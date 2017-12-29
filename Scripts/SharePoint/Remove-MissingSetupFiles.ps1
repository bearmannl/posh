<#
.SYNOPSIS
	Remove references to missing setup files.
	
.DESCRIPTION
	Remove all references to the features that are marked as missing in the SharePoint Health Analyzer. Input file needs to be in a .csv format.
	Update: including site unlock feature.

.PARAMETER Log
	Switch to log all output to a file
	
.NOTES
	Authors	: Mike Beerman
	Company	: Rubicon BV
	Date	: 2016-12-29
	Version	: 2.0

	Authors	: Mike Beerman
	Company	: Avanade Netherlands BV
	Date	: 2016-02-09
	Version	: 1.6
#>

Param (
    [switch]$Log 
)

##
#Initiate PowerShell session
##

if ($Log) {
    # Set location to store migration logs
    $logPathFolder = "D:\Logs\Scripts\MissingDependencies"

    if (!(test-path $logPathFolder)) { New-Item -ItemType directory -Path $logPathFolder }

    $logPath = "$logPathFolder\Remove-MissingSetupFileOutput_{0:yyyy-MM-dd_HHmmss}.rtf" -f (Get-Date)
    Start-Transcript -Path $logPath	
}

# Set threading
$ver = $host | Select-Object Version
if ($ver.Version.Major -gt 1) {$host.Runspace.ThreadOptions = "ReuseThread"}

# Load SharePoint PowerShell snap-in
$snapIn = "Microsoft.SharePoint.PowerShell"
if (Get-PSSnapin $snapIn -EA "SilentlyContinue") {
}
else {
    if (Get-PSSnapin $snapIn -Registered -EA "SilentlyContinue") {
        Add-PSSnapin $snapIn
        $tStamp = Get-Date -Format T
        Write-Verbose "[$tStamp] - SharePoint PowerShell SnapIn loaded"
    }
    $error.Clear()
}

Start-SPAssignment -Global

$verbosePreference = 'Continue'

##
#Read input file
##

$path = ".\InputMissingSetupFile.txt"

$sl = (Get-Content $path | Measure-Object).Count

if ($sl -eq 0) { Write-Error 'Oops, empty file!' }
if ($sl -eq 1) { $dbn = (Get-Content $path).Split(";")[0] }
if ($sl -gt 1) { $dbn = (Get-Content $path)[0].Split(";")[0] }

if ($sl -gt 0) {
    $dbs = Get-SPDatabase | Where-Object { $_.Name -eq $dbn }
    $DBserver = $dbs.Server
}
 
$input = @(Get-Content $path)  

##
#Register functions
##

#Unlock or re-lock a site if needed.
function Set-SiteLockStatus {
    Param(
        [Microsoft.SharePoint.SPSite]$Site,
        [bool]$ReadLocked,
        [bool]$WriteLocked,
        [bool]$ReadOnly
    )
    if ($ReadLocked) {
        Set-SPSite $Site -LockState NoAccess
    }
    elseif ($WriteLocked -AND !$ReadOnly) {
        Set-SPSite $Site -LockState NoAdditions 
    }
    elseif ($ReadOnly) {
        Set-SPSite $Site -LockState ReadOnly
    }
}

#Declare SQL Query function  
function Run-SQLQuery {  
    [CmdLetBinding()]
    Param(
        [string]$SqlServer,
        [string]$SqlDatabase,
        [string]$SqlQuery
    )
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection  
    $SqlConnection.ConnectionString = "Server =" + $SqlServer + "; Database =" + $SqlDatabase + "; Integrated Security = True"  
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand  
    $SqlCmd.CommandText = $SqlQuery  
    $SqlCmd.Connection = $SqlConnection  
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter  
    $SqlAdapter.SelectCommand = $SqlCmd  
    $DataSet = New-Object System.Data.DataSet  
    $SqlAdapter.Fill($DataSet)  
    $SqlConnection.Close()  
    $DataSet.Tables[0]  
}

#Declare the GetFileUrl function  
function GetFileUrl ($filepath, $DBname) {  
    #Define SQL Query and set in Variable  
    $Query = "SELECT * from AllDocs where SetupPath = '" + $filepath + "'"  
    #Running SQL Query to get information about the MissingFiles and store it in a Table  
    $QueryReturn = @(Run-SQLQuery -SqlServer $DBserver -SqlDatabase $DBname -SqlQuery $Query | Select-Object Id, SiteId, DirName, LeafName, WebId, ListId) 

    foreach ($event in $QueryReturn) {  
        if ($event.Id -and $event.SiteId -and $event.WebId) {
            $site = Get-SPSite -Identity $event.SiteId
            $web = $site | Get-SPWeb -Identity $event.WebId
            $file = $web.GetFile([Guid]$event.Id)
            $webUrl = $web.Url
            $fileUrl = $file.Url
			
            Write-Verbose "File path	: $filepath"
            Write-Verbose "Web URL		: $webUrl"
            Write-Verbose "File URL		: $fileUrl"
			
            if ($web.RootFolder.WelcomePage -ne $file.Url) {
                $siteType = $site.GetType().Name
                $siteSC = $site
                if ($siteType -eq "SPWeb") {
                    $siteSC = $site.Site
                }

                #$lockReason = $siteSC.LockIssue
                $readL = $siteSC.ReadLocked
                $writeL = $siteSC.WriteLocked
                $readO = $siteSC.ReadOnly

                $siteUrl = $siteSC.Url
				
                if ($readO -eq $true) {
                    Write-Verbose "$siteUrl locked, unlocking..."
                    Set-SPSite $siteSC -LockState Unlock
                }

                try {
                    $file.delete()
                }
                catch {
                    Write-Warning "Failed to delete $file"
                    Write-Warning $_.Exception.Message

					<#
					Write-Verbose "Attempting te remove it again..."
                    try {
                        $docLib = $file.DocumentLibrary
                    }
                    catch {
                        Write-Warning "...last attempt failed, moving on."
					}
					#>
                }

                if ($readO -eq $true) {
                    Write-Verbose "Locking site to previous state..."
                    Set-SiteLockStatus -Site $siteSC -ReadLocked $readL -WriteLocked $writeL -ReadOnly $readO
                }
            }
            if (($web.RootFolder.WelcomePage -eq $file.Url) -and $web) { Write-Host "Can't be deleted, it's a welcome page." -foregroundcolor red }
        }
    }
}

##
#Execute functions
##

#Log the CVS Column Title Line  
Write-Host "MissingSetupFile;Url" -foregroundcolor Red

foreach ($event in $input) {
    $DBname = $event.split(";")[0]
    $filepath = $event.split(";")[1]
    #call Function
    GetFileUrl $filepath $dbname
}

$verbosePreference = 'SilentlyContinue'

Stop-SPAssignment -Global

if ($Log) {
    Stop-Transcript
}