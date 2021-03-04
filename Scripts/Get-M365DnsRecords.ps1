<#
.SYNOPSIS
    Get all required DNS records for custom domains in M365.
.DESCRIPTION
    Get all required DNS records for custom domains in M365.
.PARAMETER InitialDomain
    Enter the (desired) initial tenant domain such as myname.onmicrosoft.com.
.PARAMETER AcceptedDomains
    Enter all the custom domain records for which you wish to generate the required DNS records in string array format (single string is allowed).
.PARAMETER IncludeCurrentRecords
    Add this switch to determine if the script should retrieve the current custom domain DNS records (such as the SPF record) to include current configurations in the new records.
.EXAMPLE
Get-M365DnsRecords.ps1 -InitialDomain 'mytenant.onmicrosoft.com' -AcceptedDomains 'mydomain.com'
Get-M365DnsRecords.ps1 -InitialDomain 'mytenant.onmicrosoft.com' -AcceptedDomains 'mydomain.com','myotherdomain.com' -IncludeCurrentRecords
.LINK
    https://github.com/bearmannl/posh/blob/main/Scripts/Get-M365DnsRecords.ps1
.NOTES
    Authors : Mike Beerman
    Company : Rubicon
    Date : 2021-03-04
    Version : 1.0
#>

param (
    [string]$InitialDomain,
    [array]$AcceptedDomains,
    [switch]$IncludeCurrentRecords
)

$initialDomainPrefix = $InitialDomain.Replace('.onmicrosoft.com', '')

$dnsRecords = @()

function Get-O365DnsRecords {
    param (
        [string]$DefaultDomain,
        [string]$Domain,
        [string]$InjectSpf,
        [bool]$SpfSoftfail,
        [string]$DmarcDomainPolicy,
        [string]$DmarcSubdomainPolicy,
        [string]$DmarcRuaEmail,
        [string]$DmarcRufEmail
    )

    $injectSpfFormatted = if ([string]::IsNullOrWhiteSpace($InjectSpf)) { ' ' } else { ' {0} ' -f $InjectSpf }
    $spfFailtoken = if ($SpfSoftfail) { '~' } else { '-' }
    $dmarcDomainPolicy = if ([string]::IsNullOrWhiteSpace($DmarcDomainPolicy)) { 'none' } else { $DmarcDomainPolicy }
    $dmarcSubdomainPolicy = if ([string]::IsNullOrWhiteSpace($DmarcSubdomainPolicy)) { 'none' } else { $DmarcSubdomainPolicy }
    $dmarcRuaFormatted = if ([string]::IsNullOrWhiteSpace($DmarcRuaEmail)) { '' } else { ' rua=mailto:{0}' -f $DmarcRuaEmail }
    $dmarcRufFormatted = if ([string]::IsNullOrWhiteSpace($DmarcRufEmail)) { '' } else { ' ruf=mailto:{0}' -f $DmarcRufEmail }

    $o365DnsRecords = @([pscustomobject]@{
            category = 'MX'
            host     = '@'
            value    = '{0}-{1}.mail.protection.outlook.com' -f $Domain.Split('.')[0], $Domain.Split('.')[1]
            ttl      = 3600
            type     = 'MX'
        },
        [pscustomobject]@{
            category = 'MX'
            host     = 'autodiscover'
            value    = 'autodiscover.outlook.com'
            ttl      = 3600
            type     = 'CNAME'
        },
        [pscustomobject]@{
            category = 'SPF'
            host     = '@'
            value    = 'v=spf1 mx include:spf.protection.outlook.com{0}{1}all' -f $injectSpfFormatted, $spfFailtoken
            ttl      = 3600
            type     = 'TXT'
        },
        [pscustomobject]@{
            category = 'DKIM'
            host     = 'selector1._domainkey'
            value    = 'selector1-{0}-{1}._domainkey.{2}.onmicrosoft.com.' -f $Domain.Split('.')[0], $Domain.Split('.')[1], $DefaultDomain
            ttl      = 3600
            type     = 'CNAME'
        },
        [pscustomobject]@{
            category = 'DKIM'
            host     = 'selector2._domainkey'
            value    = 'selector2-{0}-{1}._domainkey.{2}.onmicrosoft.com.' -f $Domain.Split('.')[0], $Domain.Split('.')[1], $DefaultDomain
            ttl      = 3600
            type     = 'CNAME'
        },
        [pscustomobject]@{
            category = 'DMARC'
            host     = '_dmarc'
            value    = 'v=DMARC1 p={0} sp={1} pct=100{2}{3} fo=1' -f $DmarcDomainPolicy, $DmarcSubdomainPolicy, $dmarcRuaFormatted, $dmarcRufFormatted
            ttl      = 3600
            type     = 'TXT'
        },
        [pscustomobject]@{
            category = 'SkypeForBusiness'
            host     = 'sip'
            value    = 'sipdir.online.lync.com'
            ttl      = 3600
            type     = 'CNAME'
        },
        [pscustomobject]@{
            category = 'SkypeForBusiness'
            host     = 'lyncdiscover'
            value    = 'webdir.online.lync.com'
            ttl      = 3600
            type     = 'CNAME'
        },
        [pscustomobject]@{
            category = 'SkypeForBusiness'
            host     = '_sip'
            value    = 'sipdir.online.lync.com'
            ttl      = 3600
            type     = 'SRV'
            protocol = '_tls'
            port     = 443
            weight   = 1
            priority = 100
        },
        [pscustomobject]@{
            category = 'SkypeForBusiness'
            host     = '_sipfederationtls'
            value    = 'sipfed.online.lync.com'
            ttl      = 3600
            type     = 'SRV'
            protocol = '_tcp'
            port     = 5061
            weight   = 1
            priority = 100
        },
        [pscustomobject]@{
            category = 'Intune'
            host     = 'enterpriseregistration'
            value    = 'enterpriseregistration.windows.net'
            ttl      = 3600
            type     = 'CNAME'
        },
        [pscustomobject]@{
            category = 'Intune'
            host     = 'enterpriseenrollment'
            value    = 'enterpriseenrollment.manage.microsoft.com'
            ttl      = 3600
            type     = 'CNAME'
        })

    $o365DnsRecords
}

function Get-CurrentDnsRecords {
    param (
        [string]$Domain
    )

    $currentSpf = (Resolve-DnsName -Name $Domain -Type TXT -ErrorAction Ignore | Where-Object { $_.Strings -match 'v=spf1' }).Strings
    $currentDmarc = (Resolve-DnsName -Name ('_dmarc.{0}' -f $Domain) -Type TXT -ErrorAction Ignore).Strings

    $currentO365DnsRecords = [pscustomobject]@{
        currentSpf         = $currentSpf
        currentSpfSoftfail = if ($currentSpf -match '~all') { $true } else { $false }
        currentDmarcDP     = if ($currentDmarc -match ' p=') { $currentDmarc.Split(' p=')[1].Split('')[0] } else { $null }
        currentDmarcSDP    = if ($currentDmarc -match ' sp=') { $currentDmarc.Split(' sp=')[1].Split('')[0] } else { $null }
        currentDmarcRua    = if ($currentDmarc -match ' rua=') { $currentDmarc.Split(' rua=')[1].Split('')[0].Replace('mailto:', '') } else { $null }
        currentDmarcRuf    = if ($currentDmarc -match ' ruf=') { $currentDmarc.Split(' ruf=')[1].Split('')[0].Replace('mailto:', '') } else { $null }
    }

    $currentO365DnsRecords
}

foreach ($domain in $AcceptedDomains) {
    if ($IncludeCurrentRecords) {
        $currentRecords = Get-CurrentDnsRecords -Domain $domain
        $spf = $currentRecords.currentSpf
        $spfSoftFail = $currentRecords.currentSpfSoftfail
        $dmarcDP = $currentO365DnsRecords.currentDmarcDP
        $dmarcSDP = $currentO365DnsRecords.currentDmarcSDP
        $dmarcRua = $currentO365DnsRecords.currentDmarcRua
        $dmarcRuf = $currentO365DnsRecords.currentDmarcRuf
    
        $spfInjectContent = if ($spf) { $spf.Replace('v=spf1', '').Replace('mx', '').Replace('include:spf.protection.outlook.com', '').Replace('include:spf.protection.outlook.com', '').Replace('-all', '').Replace('~all', '').Replace('?all', '').Trim() }
    }
    else {
        $spfSoftFail = $false
    }

    $dnsRecords += Get-O365DnsRecords -DefaultDomain $initialDomainPrefix -Domain $domain -InjectSpf $spfInjectContent -SpfSoftfail $spfSoftFail -DmarcDomainPolicy $dmarcDP -DmarcSubdomainPolicy $dmarcSDP -DmarcRuaEmail $dmarcRua -DmarcRufEmail $dmarcRuf
}

$dnsRecords
