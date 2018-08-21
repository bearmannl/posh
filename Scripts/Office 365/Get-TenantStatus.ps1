<#
.SYNOPSIS
     Get O365 tenant name availability.
      
.DESCRIPTION
     Determines if the tenant name requested is available or already taken within Office 365.
     Provide a desired tenant name without the .onmicrosoft.com suffix to check if that tenant name is available.

.PARAMETER OrganizationName
    The desired tenant name for your organization.

.EXAMPLE
     Get-TenantStatus.ps1 -OrganizationName "microsoft"
      
.LINK
    https://www.linkedin.com/pulse/how-check-office-365-tenant-name-availability-aaron-dinnage
      
.NOTES
     Authors	: Aaron Dinnage
     Company	: Microsoft
     Date		: 2015-11-12
     Version	: 1.0
					1.1 | Mike Beerman | 2018-08-21 | Script cleanup & formatting.
#>
param(
    [Parameter(Mandatory = $true)][string]$OrganizationName
)

$tenantName = "$($OrganizationName.ToLower()).onmicrosoft.com"
Write-Host -NoNewLine "Checking availability for tenant $($tenantName) "

$uri = "https://portal.office.com/Signup/CheckDomainAvailability.ajax"
$body = "p0=$($OrganizationName)&assembly=BOX.Admin.UI%2C+Version%3D16.0.0.0%2C+Culture%3Dneutral%2C+PublicKeyToken%3Dnull&class=Microsoft.Online.BOX.Signup.UI.SignupServerCalls"

$invokeJob = Start-Job -ScriptBlock { param($uri, $body); return Invoke-RestMethod -Method Post -Uri $uri -Body $body } -ArgumentList $uri,$body
while ($invokeJob.State -Match 'Running') { Write-Host -NoNewline "."; Start-Sleep 1 }
Write-Host
Write-Host "$($tenantName) is " -NoNewLine
$response = Receive-Job -Job $invokeJob

if ($response.Contains("SessionValid") -eq $false) { Write-Host -ForegroundColor Red "[Incorrect]"; Write-Host -ForegroundColor Red $response; exit }
else {
    if ($response.Contains("<![CDATA[1]]>")) { Write-Host -ForegroundColor Green "[Available]" }
    else { Write-Host -ForegroundColor Yellow "[Taken]" }
}
