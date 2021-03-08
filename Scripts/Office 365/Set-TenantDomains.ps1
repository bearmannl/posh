<#
.SYNOPSIS
    Add all required DNS domains for custom domains in M365 and return all the matching TXT records required for validation.
.DESCRIPTION
    Add all required DNS domains for custom domains in M365 and return all the matching TXT records required for validation.
.PARAMETER Domains
    Enter the (desired) custom tenant domains such as mydomain.com.
.EXAMPLE
Set-SetTenantDomains.ps1 -Domains 'mydomain.com'
Set-SetTenantDomains.ps1 -Domains 'mydomain.com','mydomain.eu'
.LINK
    https://github.com/bearmannl/posh/blob/main/Scripts/Set-SetTenantDomains.ps1
.NOTES
    Authors : Mike Beerman
    Company : Rubicon
    Date : 2021-03-08
    Version : 1.0
#>
param (
    [array]$Domains
)

if (!Get-InstalledModule MSOnline) { Install-Module MSOnline }
Import-Module MSOnline

try {
    Get-MsolDomain -ErrorAction Stop > $null
}
catch {
    Connect-MsolService
}

$tenantId = (Get-MsolAccountSku).AccountObjectId[0]
foreach ($domain in $Domains) { New-MsolDomain -TenantId $tenantId -Name $domain }
foreach ($domain in $Domains) { Get-MsolDomainVerificationDNS -TenantId $tenantId -DomainName $domain -Mode DnsTxtRecord }
