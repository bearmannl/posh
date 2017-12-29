<#
.SYNOPSIS
	Remove references to missing web parts.
	
.DESCRIPTION
	Remove all references to the web parts that are marked as missing in the SharePoint Health Analyzer. Input file needs to be in a .csv format.

.NOTES
	Authors	: Mike Beerman
	Date	: 2015-12-31
	Version	: 1.5
#>

##
#Initiate PowerShell session
##

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

$path = ".\InputMissingWebPart.txt"

$sl = (Get-Content $path | Measure-Object).Count

if ($sl -eq 0) { Write-Error 'Oops, empty file!' }
if ($sl -eq 1) { $dbn = (Get-Content $path).Split(";")[0] }
if ($sl -gt 1) { $dbn = (Get-Content $path)[0].Split(";")[0] }

if ($sl -gt 0) {
    $dbs = Get-SPDatabase | Where-Object { $_.Name -eq $dbn }
    $DBserver = $dbs.Server
}

#Set Variables
$input = @(Get-Content $path)

#Declare Log File
function StartTracing {
    [CmdLetBinding()]
    Param()
    $logPath = 'D:\Logs\Scripts\MissingDependencies'
    if (!(Test-Path -Path $logPath -PathType Container)) { New-Item -ItemType Directory -Force -Path $logPath }
    $LogTime = Get-Date -Format yyyy-MM-dd_h-mm
    Start-Transcript -Path "$logPath\MissingWebPartOutput-$LogTime.rtf"
}
 
#Declare SQL Query function
function Invoke-SQLQuery {
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

function GetWebPartDetails ($wpid, $DBname) {
    #Define SQL Query and set in Variable
    $Query = "SELECT * from AllDocs inner join AllWebParts on AllDocs.Id = AllWebParts.tp_PageUrlID where AllWebParts.tp_WebPartTypeID = '" + $wpid + "'"
 
    #Running SQL Query to get information about Assembly (looking in EventReceiver Table) and store it in a Table
    $QueryReturn = @(Invoke-SQLQuery -SqlServer $DBserver -SqlDatabase $DBname -SqlQuery $Query | Select-Object Id, SiteId, DirName, LeafName, WebId, ListId, tp_ZoneID, tp_DisplayName)
 
    #Actions for each element in the table returned
    foreach ($event in $QueryReturn) {
        if ($event.id -ne $null) {
            #Get Site URL
            $site = Get-SPSite -Identity $event.SiteId
    
            #Log information to Host
            Write-Host $wpid -nonewline -foregroundcolor yellow
            write-host ";" -nonewline
            write-host $event.tp_DisplayName -foregroundcolor yellow
            write-host ";" -nonewline
            write-host $site.Url -nonewline -foregroundcolor green
            write-host "/" -nonewline -foregroundcolor green
            write-host $event.LeafName -foregroundcolor green -nonewline
            write-host ";" -nonewline
            write-host $site.Url -nonewline -foregroundcolor gray
            write-host "/" -nonewline -foregroundcolor gray
            write-host $event.DirName -foregroundcolor gray -nonewline
            write-host "/" -nonewline -foregroundcolor gray
            write-host $event.LeafName -foregroundcolor gray -nonewline
            write-host "?contents=1" -foregroundcolor gray -nonewline
            write-host ";" -nonewline
            write-host $event.tp_ZoneID -foregroundcolor cyan				
        }
    }
}
 
#Start Logging
StartTracing
 
#Log the CVS Column Title Line
write-host "WebPartID;PageUrl;MaintenanceUrl;WpZoneID" -foregroundcolor Red
 
foreach ($event in $input) {
    $DBname = $event.split(";")[0]
    $wpid = $event.split(";")[1]
    GetWebPartDetails $wpid $dbname
}
    
#Stop Logging
Stop-Transcript