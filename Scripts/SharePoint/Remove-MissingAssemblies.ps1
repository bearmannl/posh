<#
.SYNOPSIS
	Remove references to missing assemblies.
	
.DESCRIPTION
	Remove all references to the assemblies that are marked as missing in the SharePoint Health Analyzer. Input file needs to be in a .csv format.

.PARAMETER DeleteReferences
	Switch to delete all references. Otherwise just list the references.

.EXAMPLE
	Remove-MissingAssemblies.ps1
	Remove-MissingAssemblies.ps1 -DeleteReferences
	
.NOTES
	Authors	: Mike Beerman
	Date	: 2015-12-31
	Version	: 1.5
#>

Param (
    [switch]$DeleteReferences
)

##
#Initiate PowerShell session
##

$verbosePreference = 'Continue'

# Set threading
$ver = $host | Select-Object Version
if ($ver.Version.Major -gt 1) {$Host.Runspace.ThreadOptions = "ReuseThread"}

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

$path = ".\InputMissingAssembly.txt"

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
    Start-Transcript -Path "$logPath\MissingAssemblyOutput-$LogTime.rtf"
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


function GetAssemblyDetails {
    [CmdLetBinding()]
    Param(
        [string]$assembly,
        [string]$DBname
    )
    #Define SQL Query and set in Variable
    $Query = "SELECT * from EventReceivers where Assembly = '" + $assembly + "'"
    #$Query = "SELECT * from EventReceivers where Assembly = 'Microsoft.Office.InfoPath.Server, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'" 

    #Running SQL Query to get information about Assembly (looking in EventReceiver Table) and store it in a Table
    $QueryReturn = @(Invoke-SQLQuery -SqlServer $DBserver -SqlDatabase $DBname -SqlQuery $Query | Select-Object Id, Name, SiteId, WebId, HostId, HostType)

    #Actions for each element in the table returned
    foreach ($event in $QueryReturn) {   
        #HostID (check http://msdn.microsoft.com/en-us/library/ee394866(v=prot.13).aspx for HostID Type reference)
        if ($event.HostType -eq 0) {
            $site = Get-SPSite -Identity $event.SiteId			
            #Get the EventReceiver Site Object
            $er = $site.EventReceivers | Where-Object {$_.Id -eq $event.Id}
			
            if ($er) {
                $sUrl = $site.Url
                $ern = $er.Name
                $erc = $er.Class
                Write-Verbose "Assembly	: $assembly"
                Write-Verbose "URL			: $sUrl"
                Write-Verbose "Event Receiver Name: $ern"
                Write-Verbose "Event Receiver Class: $erc"
				
                if ($DeleteReferences) { $er.Delete() }
            }
        }
		 
        if ($event.HostType -eq 1) {
            $site = Get-SPSite -Identity $event.SiteId
            $web = $site | Get-SPWeb -Identity $event.WebId
            #Get the EventReceiver Site Object
            $er = $web.EventReceivers | Where-Object {$_.Id -eq $event.Id}
			
            if ($er) {
                $wUrl = $web.Url
                $ern = $er.Name
                $erc = $er.Class
                Write-Verbose "Assembly	: $assembly"
                Write-Verbose "URL			: $wUrl"
                Write-Verbose "Event Receiver Name: $ern"
                Write-Verbose "Event Receiver Class: $erc"
				
                if ($DeleteReferences) { $er.Delete() }
            }
        }
		 
        if ($event.HostType -eq 2) {
            $site = Get-SPSite -Identity $event.SiteId
            $web = $site | Get-SPWeb -Identity $event.WebId
            $list = $web.Lists | Where-Object {$_.Id -eq $event.HostId}
            #Get the EventReceiver List Object
            $er = $list.EventReceivers | Where-Object {$_.Id -eq $event.Id}
			
            if ($er) {
                $sUrl = $site.Url
                $lrf = $list.RootFolder
                $ern = $er.Name
                $erc = $er.Class
                Write-Verbose "Assembly	: $assembly"
                Write-Verbose "URL			: $sUrl"
                Write-Verbose "List		: $lrf"
                Write-Verbose "Event Receiver Name: $ern"
                Write-Verbose "Event Receiver Class: $erc"
				
                if ($DeleteReferences) { $er.Delete() }
            }
        }
    }
}

#Start Logging
StartTracing

$i = 0

foreach ($event in $input) {
    $i++
	Write-Verbose "Line		: $i"
	$DBname = $event.split(";")[0]
    $assembly = $event.split(";")[1]
    GetAssemblyDetails $assembly $dbname
}
    
$verbosePreference = 'SilentlyContinue'
	
#Stop Logging
Stop-Transcript