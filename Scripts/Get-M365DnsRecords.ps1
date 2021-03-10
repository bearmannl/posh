using namespace System
using namespace System.Collections.ObjectModel
<#
.SYNOPSIS
    Get all required DNS records for custom domains in M365.
.DESCRIPTION
    Get all required DNS records for custom domains in M365.
.PARAMETER InitialDomain
    Enter the (desired) initial tenant domain such as myname.onmicrosoft.com.
.PARAMETER AcceptedDomains
    Enter all the custom domain records for which you wish to generate the required DNS records in string array format (single string is allowed).
.PARAMETER IncludeValidationRecords
    PLEASE NOTE! Using this switch will involve logging into an M365 tenant with a valid admin account! Add this switch to retrieve the custom domain validation DNS records to include them in the output.
.PARAMETER IncludeCurrentRecords
    Add this switch to retrieve the current custom domain DNS records (such as the SPF record) to include current configurations in the new records.
.PARAMETER OutputWord
    Add switch to output the records into a Word document. Requires an installed Word application on your machine!
.PARAMETER WordFileName
    Desired filename for the Word document if that output is switched.
.PARAMETER WordFileOutputPath
    Desired output path for the Word document if that output is switched.
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
    [string]$InitialDomain = ($InitialDomain.Replace('.onmicrosoft.com', '')),
    [array]$AcceptedDomains,
    [switch]$IncludeValidationRecords,
    [switch]$IncludeCurrentRecords,
    [switch]$OutputWord,
    [string]$WordFileName = "dns.records.docx",
    [string]$WordFileOutputPath = (Split-Path $script:MyInvocation.MyCommand.Path)
)

#region datamodel
class Output {
    [Collection[Domain]] $AcceptedDomains

    Output() {
        $this.AcceptedDomains = @()
    }
}

class Domain {
    [string] $Name
    [Collection[DnsRecord]] $Records

    Domain() {
        $this.Records = @()
    }
}

class DnsRecord {
    [string] $Category
    [string] $Host
    [string] $Value
    [int] $TimeToLive
    [string] $Type
    [string] $Protocol
    [int] $Port
    [int] $Weight
    [int] $Priority
}
#endregion

#region functions
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

    $o365DnsRecords = [Domain]@{
        Name = $Domain
    }

    $injectSpfFormatted = if ([string]::IsNullOrWhiteSpace($InjectSpf)) { ' ' } else { ' {0} ' -f $InjectSpf }
    $spfFailtoken = if ($SpfSoftfail) { '~' } else { '-' }
    $dmarcDomainPolicy = if ([string]::IsNullOrWhiteSpace($DmarcDomainPolicy)) { 'none' } else { $DmarcDomainPolicy }
    $dmarcSubdomainPolicy = if ([string]::IsNullOrWhiteSpace($DmarcSubdomainPolicy)) { 'none' } else { $DmarcSubdomainPolicy }
    $dmarcRuaFormatted = if ([string]::IsNullOrWhiteSpace($DmarcRuaEmail)) { '' } else { ' rua=mailto:{0}' -f $DmarcRuaEmail }
    $dmarcRufFormatted = if ([string]::IsNullOrWhiteSpace($DmarcRufEmail)) { '' } else { ' ruf=mailto:{0}' -f $DmarcRufEmail }
    
    if ($IncludeValidationRecords) {
        if (!Get-InstalledModule MSOnline) { Install-Module MSOnline -Scope CurrentUser }
        Import-Module MSOnline
        try { Get-MsolDomain -ErrorAction Stop > $null }
        catch { Connect-MsolService }

        $o365DnsRecords.Records.Add(([DnsRecord]@{
                    Category   = 'Validation'
                    Host       = '@'
                    Value      = (Get-MsolDomainVerificationDNS -TenantId ((Get-MsolAccountSku).AccountObjectId[0]) -DomainName $domain -Mode DnsTxtRecord)
                    TimeToLive = 3600
                    Type       = 'TXT'
                }
            )
        )
    }

    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'MX'
                Host       = '@'
                Value      = '{0}-{1}.mail.protection.outlook.com' -f $Domain.Split('.')[0], $Domain.Split('.')[1]
                TimeToLive = 3600
                Type       = 'MX'
            }
        )
    )
    
    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'MX'
                Host       = 'autodiscover'
                Value      = 'autodiscover.outlook.com'
                TimeToLive = 3600
                Type       = 'CNAME' 
            }
        )
    )

    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'SPF'
                Host       = '@'
                Value      = (('v=spf1 mx include:spf.protection.outlook.com{0}{1}all' -f $injectSpfFormatted, $spfFailtoken).Replace('   ', ' ').Replace('  ', ' '))
                TimeToLive = 3600
                Type       = 'TXT' 
            }
        )
    )

    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'DKIM'
                Host       = 'selector1._domainkey'
                Value      = 'selector1-{0}-{1}._domainkey.{2}.onmicrosoft.com.' -f $Domain.Split('.')[0], $Domain.Split('.')[1], $DefaultDomain
                TimeToLive = 3600
                Type       = 'CNAME'
            }
        )
    )

    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'DKIM'
                Host       = 'selector2._domainkey'
                Value      = 'selector2-{0}-{1}._domainkey.{2}.onmicrosoft.com.' -f $Domain.Split('.')[0], $Domain.Split('.')[1], $DefaultDomain
                TimeToLive = 3600
                Type       = 'CNAME'
            }
        )
    )

    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'DMARC'
                Host       = '_dmarc'
                Value      = 'v=DMARC1 p={0} sp={1} pct=100{2}{3} fo=1' -f $DmarcDomainPolicy, $DmarcSubdomainPolicy, $dmarcRuaFormatted, $dmarcRufFormatted
                TimeToLive = 3600
                Type       = 'TXT'
            }
        )
    )

    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'Skype for Business'
                Host       = 'sip'
                Value      = 'sipdir.online.lync.com'
                TimeToLive = 3600
                Type       = 'CNAME'
            }
        )
    )
 
    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'Skype for Business'
                Host       = 'lyncdiscover'
                Value      = 'webdir.online.lync.com'
                TimeToLive = 3600
                Type       = 'CNAME'
            }
        )
    )
    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'SRV'
                Host       = '_sip'
                Value      = 'sipdir.online.lync.com'
                TimeToLive = 3600
                Type       = 'SRV'
                Protocol   = '_tls'
                Port       = 443
                Weight     = 1
                Priority   = 100
            }
        )
    )
    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'SRV'
                Host       = '_sipfederationtls'
                Value      = 'sipfed.online.lync.com'
                TimeToLive = 3600
                Type       = 'SRV'
                Protocol   = '_tcp'
                Port       = 5061
                Weight     = 1
                Priority   = 100
            }
        )
    )
    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'Intune & MDM'
                Host       = 'enterpriseregistration'
                Value      = 'enterpriseregistration.windows.net'
                TimeToLive = 3600
                Type       = 'CNAME'
            }
        )
    )
    $o365DnsRecords.Records.Add(([DnsRecord]@{
                Category   = 'Intune & MDM'
                Host       = 'enterpriseenrollment'
                Value      = 'enterpriseenrollment.manage.microsoft.com'
                TimeToLive = 3600
                Type       = 'CNAME'
            }
        )
    )

    return $o365DnsRecords
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

function CreateWordDocument {
    param(
        [Output]$RecordsObject
    )

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    $doc = $word.Documents.Add()
    $doc.Styles["Normal"].ParagraphFormat.SpaceBefore = 0
    $doc.Styles["Normal"].ParagraphFormat.SpaceAfter = 1
    $margin = 36 # 1.26 cm
    $doc.PageSetup.LeftMargin = $margin
    $doc.PageSetup.RightMargin = $margin
    $doc.PageSetup.TopMargin = $margin
    $doc.PageSetup.BottomMargin = $margin
    $selection = $word.Selection

    foreach ($domain in $RecordsObject.AcceptedDomains) {
        $selection.Style = "Heading 1"
        $selection.TypeText($domain.Name)
        $selection.TypeParagraph()
        $orderSequence = 'Validation', 'MX', 'SPF', 'DKIM', 'DMARC', 'Skype for Business', 'SRV', 'Intune & MDM'
        $groups = $domain.Records | Group-Object -Property Category | Sort-Object { $orderSequence.IndexOf($_.Name) }
        foreach ($group in $groups) {
            if ($group.Name -eq "DKIM" -or $group.Name -eq "DMARC" -or $group.Name -eq "SPF" -or $group.Name -eq "SRV") {
                $selection.Style = "Heading 3"
            }
            else {
                $selection.Style = "Heading 2"
            }
            $selection.TypeText($group.Name)
            $selection.TypeParagraph()
            foreach ($record in $group.Group) {
                if ($record.Port -eq 0) {
                    $selection.Style = "Normal"
                    $selection.TypeText("Host:`t`t$($record.Host)`vValue:`t`t$($record.Value)`vTimeToLive:`t$($record.TimeToLive)`vType:`t`t$($record.Type)`v")
                    $selection.TypeParagraph()
                }
                else {
                    $selection.Style = "Normal"
                    $selection.TypeText("Host:`t`t$($record.Host)`vValue:`t`t$($record.Value)`vTimeToLive:`t$($record.TimeToLive)`vType:`t`t$($record.Type) Protocol:`t`t$($record.Protocol)`vPort:`t`t$($record.Port)`vWeight:`t`t$($record.Weight)`vPriority:`t`t$($record.Priority)`v")
                    $selection.TypeParagraph()
                }
            }
        }
    }

    $outputPath = $WordFileOutputPath + "\" + $WordFileName
    $doc.SaveAs($outputPath)
    $doc.Close()
    $word.Quit()

}
#endregion

#region script execution
$dnsRecords = [Output]::New()

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
    
    $dnsRecords.AcceptedDomains += Get-O365DnsRecords -DefaultDomain $initialDomainPrefix -Domain $domain -InjectSpf $spfInjectContent -SpfSoftfail $spfSoftFail -DmarcDomainPolicy $dmarcDP -DmarcSubdomainPolicy $dmarcSDP -DmarcRuaEmail $dmarcRua -DmarcRufEmail $dmarcRuf
}

if ($OutputWord) {
    CreateWordDocument -RecordsObject $dnsRecords
}
else {
    $dnsRecords
}
#endregion
