<#
.SYNOPSIS
	Remove references to missing features.
	
.DESCRIPTION
	Remove all references to the features that are marked as missing in the SharePoint Health Analyzer. Input file needs to be in a .csv format.
    Update: Script now includes ability to unlock site collections for deletion of features, follow by re-locking the site after feature deletion.

.PARAMETER DeleteReferences
	Switch to delete all references. Otherwise just list the references.

.EXAMPLE
	Remove-MissingFeatures.ps1
	Remove-MissingFeatures.ps1 -DeleteReferences
	
.NOTES
	Authors	: Mike Beerman
    Company : Rubicon BV
	Date	: 2016-12-29
	Version	: 2.0

	Authors	: Mike Beerman
    Company : Avanade Netherlands BV
	Date	: 2015-12-31
	Version	: 1.5
#>

Param (
	[switch]$DeleteReferences
)

##
#Initiate PowerShell session
##

# Set threading
$ver = $host | Select-Object Version
if($ver.Version.Major -gt 1) {$host.Runspace.ThreadOptions = "ReuseThread"}

# Load SharePoint PowerShell snap-in
$snapIn="Microsoft.SharePoint.PowerShell"
if(Get-PSSnapin $snapIn -EA "SilentlyContinue") {
}
else 
{
	if (Get-PSSnapin $snapIn -Registered -EA "SilentlyContinue") {
		Add-PSSnapin $snapIn
		$tStamp = Get-Date -Format T
		Write-Verbose "[$tStamp] - SharePoint PowerShell SnapIn loaded"
	}
$error.Clear()
}

##
#Set Variables
##

$path = ".\InputMissingFeature.txt"
$input = @(Get-Content $path)

##
#Declare Log File
##

function StartTracing
{
[CmdLetBinding()]
Param()
	$logPath = 'D:\Logs\Scripts\MissingDependencies'
	if (!(Test-Path -Path $logPath -PathType Container)) { New-Item -ItemType Directory -Force -Path $logPath }
    $LogTime = Get-Date -Format yyyy-MM-dd_h-mm
	Start-Transcript -Path "$logPath\MissingFeaturesOutput-$LogTime.rtf"
}

function Set-SiteLockStatus {
    Param(
    [Microsoft.SharePoint.SPSite]$Site,
    [bool]$ReadLocked,
    [bool]$WriteLocked,
    [bool]$ReadOnly
)
if ($ReadLocked) {
    Set-SPSite $Site -LockState NoAccess
        }elseif ($WriteLocked -AND !$ReadOnly) {
            Set-SPSite $Site -LockState NoAdditions 
                }elseif ($ReadOnly) {
                    Set-SPSite $Site -LockState ReadOnly
                    }
}

function Remove-SPFeatureFromContentDB
{
[CmdLetBinding()]
Param(
    [string]$ContentDb,
    [string]$FeatureId
)
    $db = Get-SPDatabase | Where-Object { $_.Name -eq $ContentDb }    
    
    $db.Sites | ForEach-Object {
        Remove-SPFeature -obj $_ -objName "site collection" -featId $FeatureId
        
        $_ | Get-SPWeb -Limit all | ForEach-Object {
            Remove-SPFeature -obj $_ -objName "site" -featId $FeatureId
        }
    }
}

function Remove-SPFeature
{
[CmdLetBinding()]
Param(
    $obj,
    [string]$objName,
	[string]$featId
)
    $siteType = $obj.GetType().Name
    $siteSC = $obj
    if ($siteType -eq "SPWeb") {
        $siteSC = $obj.Site
    }

    #$lockReason = $siteSC.LockIssue
    $readL = $siteSC.ReadLocked
    $writeL = $siteSC.WriteLocked
    $readO = $siteSC.ReadOnly

    $siteUrl = $siteSC.Url

    $feature = $obj.Features[$featId]    
    if ($feature -ne $null) {
        $verbosePreference = 'Continue'
        if ($DeleteReferences) {
            if ($readO -eq $true) {
            Write-Verbose "$siteUrl locked, unlocking..."
            Set-SPSite $siteSC -LockState Unlock
            }
            try {
                $obj.Features.Remove($feature.DefinitionId, $true)
                Write-Host "Feature successfully removed from" $objName ":" $obj.Url -foregroundcolor Red
            }
            catch {
                Write-Host "There has been an error trying to remove the feature:" $_
            }
            if ($readO -eq $true) {
                Write-Verbose "Locking site to previous state..."
                Set-SiteLockStatus -Site $siteSC -ReadLocked $readL -WriteLocked $writeL -ReadOnly $readO
            }
        }
        else
        {
			Write-Host "Feature found in" $objName ":" $obj.Url -foregroundcolor Red
        }
        $verbosePreference = 'SilentlyContinue'
    }
    else {
        #Write-Host "Feature ID specified does not exist in" $objName ":" $obj.Url
    }
}

#Start Logging
StartTracing

#Log the CVS Column Title Line
Write-Host "FeatureName;FeatureUrl" -foregroundcolor Red
 
foreach ($event in $input)
    {    
    $DBname = $event.split(";")[0]
	$fid = $event.split(";")[1]
    Remove-SPFeatureFromContentDB -ContentDb $dbname -FeatureId $fid
    }

#Stop Logging
Stop-Transcript