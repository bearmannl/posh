<#
.SYNOPSIS
     Tests status of firewall ports and outputs to various formats.
      
.DESCRIPTION
     This script checks if specified ports are open *from* local server *to* remote servers, based on the server and port mapping in the .json file. Possible output formats include .csv and .htm.

.PARAMETER Output
	Required parameter, determines the desired type of output. Currently supports CSV and HTML.

.PARAMETER BaseDirectoryPath
    Optional parameter to direct the script to a different folder for the JSON config files. Uses script invocation path by default if no input is provided.
    
.PARAMETER BaseConfigFile
    Optional parameter to point the script to a different file for the JSON config. Uses script invocation path by default if no input is provided.

.PARAMETER LogFolderPath
    Optional parameter to direct the script to a different folder for the log output. Uses script invocation path by default if no input is provided.

.PARAMETER HtmlOutputFolderPath
    Optional parameter to direct the script to a different folder for the HTML output. Uses script invocation path by default if no input is provided.

.PARAMETER CsvOutputFolderPath
    Optional parameter to direct the script to a different folder for the CSV output. Uses script invocation path by default if no input is provided.

.PARAMETER TestData
	Optional switch, portscans can take a while, in case of debugging or editing styling of the script, use test data.

.EXAMPLE
     Test-NetConnections.ps1 -Output CSV -CsvOutputFolderPath "C:\Temp\Output\"
	 Test-NetConnections.ps1 -Output HTML -BaseDirectoryPath "C:\Scripts\Tests\" -BaseConfigFile "Test-NetConnections.AADCsample.json"
     Test-NetConnections.ps1 -Output HTML -TestData
     Test-NetConnections.ps1 -Output HTML -TestData -LogFolderPath "C:\Scripts\Logs\" -HtmlOutputFolderPath "C:\inetpub\wwwroot\health80\"
      
.LINK
    http://www.rubicon.nl
      
.NOTES
     Authors	: Mike Beerman
     Company	: Rubicon
     Date		: 2017-08-08
     Version	: 1.0
					1.1 | MBE | 2017-08-10 | Added HTML output, including function, param and switch.
					1.2 | MBE | 2017-08-11 | Added comments.
					1.3 | MBE | 2017-08-14 | Added explicit states for the firewall status. Includes ports which are open but have no service listening and strict ports.
                    1.4 | MBE | 2017-08-19 | Cleaned up the TestData functionality. Added basic output folder checks.
                    1.5 | MBE | 2018-11-22 | Refactored some params in line with lessons learned over the past year. Using Write-Host since this is now fully supported with the Information stream.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('CSV', 'HTML')]
    [String]$Output,
    [String]$BaseDirectoryPath,
    [String]$BaseConfigFile = "Test-NetConnections.json",
    [String]$LogFolderPath,
    [String]$HtmlOutputFolderPath,
    [String]$CsvOutputFolderPath,
    [Switch]$TestData
)

#region basic script setup
$script:localServer = $env:COMPUTERNAME
$script:cDate = (Get-Date)
$script:logFile = "Test-NetConnectionsOutput_$($script:localServer)_{0:yyyy-MM-dd_HHmmss}.rtf" -f ($script:cDate)
$script:htmlOutPutFile = "Test-NetConnectionsOutput.htm"
$script:csvOutPutFile = "Test-NetConnectionsOutput_$($script:localServer)_{0:yyyy-MM-dd_HHmmss}.csv" -f ($script:cDate)

# If no alternate BaseDirectoryPath is given, take script invocation path
if ([string]::IsNullOrEmpty($BaseDirectoryPath)) { $BaseDirectoryPath = $PSScriptRoot }
if ([string]::IsNullOrEmpty($LogFolderPath)) { $LogFolderPath = $PSScriptRoot }
# Ensure the directory path params end in a slash
if ($false -eq ($BaseDirectoryPath -Match '.+?\\$')) { $BaseDirectoryPath = $BaseDirectoryPath + "\" }
if ($false -eq ($LogFolderPath -Match '.+?\\$')) { $LogFolderPath = $LogFolderPath + "\" }

# Check location to store system logs exists, create if not
if ($null -eq (Test-Path $LogFolderPath)) { New-Item -ItemType directory -Path $LogFolderPath > $null; Write-Host "Log output directory created at path $($LogFolderPath)" }
$script:logPath = "$($LogFolderPath)$($script:logFile)"

if ($Output -eq "HTML") {
    if ([string]::IsNullOrEmpty($HtmlOutputFolderPath)) { $HtmlOutputFolderPath = $PSScriptRoot }

    if ($null -eq (Test-Path $HtmlOutputFolderPath)) {
        $title = "HTML output path invalid"
        $message = "$($HtmlOutputFolderPath) does not exist, do you want to create a folder at this path?"
		
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Creates a new folder."
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Exits the script."
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
		
        switch ($result) {
            0 { New-Item -ItemType directory -Path $HtmlOutputFolderPath > $null; Write-Host "HTML output folder created." }
            1 { Write-Error "No output folder available!" }
        }
    }
    # Ensure the directory path params end in a slash
    if ($false -eq ($HtmlOutputFolderPath -Match '.+?\\$')) { $HtmlOutputFolderPath = $HtmlOutputFolderPath + "\" }
}

if ($Output -eq "CSV") {
    if ([string]::IsNullOrEmpty($CsvOutputFolderPath)) { $CsvOutputFolderPath = $PSScriptRoot }

    if ($null -eq (Test-Path $CsvOutputFolderPath)) {
        $title = "CSV output path invalid"
        $message = "$($CsvOutputFolderPath) does not exist, do you want to create a folder at this path?"
		
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Creates a new folder."
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Exits the script."
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.UI.PromptForChoice($title, $message, $options, 0) 
		
        switch ($result) {
            0 { New-Item -ItemType directory -Path $CsvOutputFolderPath > $null; Write-Host "CSV output folder created." }
            1 { Write-Error "No output folder available!" }
        }
    }
    # Ensure the directory path params end in a slash
    if ($false -eq ($CsvOutputFolderPath -Match '.+?\\$')) { $CsvOutputFolderPath = $CsvOutputFolderPath + "\" }
}

#endregion

#region Functions
function ConvertTo-HtmlConditionalFormat {
    [CmdLetBinding()]
    param(
        $PlainHtml,
        $ConditionalStyle
    )
    Add-Type -AssemblyName System.Xml.Linq
    # load HTML content as XML
    $xml = [System.Xml.Linq.XDocument]::Parse($PlainHtml)
    $namespace = 'http://www.w3.org/1999/xhtml'
    # select the type of html elements for processing
    $elements = $xml.Descendants("{$namespace}td")
    # loop through each conditional formatting rule
    foreach ($cs in $ConditionalStyle.Keys) {
        $scriptBlock = [scriptblock]::Create($cs)
        # find the column matching the correct content value
        $columnIndex = (($xml.Descendants("{$namespace}th") | Where-Object { $_.Value -eq $ConditionalStyle.$cs[0] }).NodesBeforeSelf() | Measure-Object).Count
        # only select the elements matching the correct column
        $elements | Where-Object { ($_.NodesBeforeSelf() | Measure-Object).Count -eq $columnIndex } | ForEach-Object {
            if (&$scriptBlock) {
                # apply the actual attribute value
                $_.SetAttributeValue( "style", $ConditionalStyle.$cs[1])
            }
        }
    }
	
    Write-Output $xml.ToString()
}

function Test-PortStatus {
    [CmdLetBinding()]
    param(
        $RemoteHost,
        $RemotePort
    )
    if ($TestData) {
        # portscans can take a while, in case of debugging or editing styling, use test data
        $pStatus = Get-Random -Input "LISTENING", "LISTENING", "LISTENING", "LISTENING", "LISTENING", "OPEN", "OPEN", "OPEN", "CLOSED", "CLOSED", "STRICT", "STRICT", "UNKNOWN"
    }
    else {
        # perform a met connection test, determine status based on PingSucceeded and TcpTestSucceeded
        $testConnection = Test-NetConnection -Computername $RemoteHost -Port $RemotePort -WarningAction:SilentlyContinue
        if ($testConnection.PingSucceeded -AND $testConnection.TcpTestSucceeded) { $pStatus = "LISTENING" } 
        elseif ($testConnection.PingSucceeded -AND !$testConnection.TcpTestSucceeded) { $pStatus = "OPEN" }
        elseif (!$testConnection.PingSucceeded -AND !$testConnection.TcpTestSucceeded) { $pStatus = "CLOSED" }
        elseif (!$testConnection.PingSucceeded -AND $testConnection.TcpTestSucceeded) { $pStatus = "STRICT" }
        else { $pStatus = "UNKNOWN" }        
    }
    Write-Host "[$($RemotePort)] [" -NoNewline

    switch ($pStatus) {
        LISTENING { 
            Write-Host "$($pStatus)" -NoNewline -ForegroundColor DarkGreen
        }
        OPEN { 
            Write-Host "$($pStatus)" -NoNewline -ForegroundColor DarkYellow
        }
        CLOSED { 
            Write-Host "$($pStatus)" -NoNewline -ForegroundColor DarkRed
        }
        STRICT { 
            Write-Host "$($pStatus)" -NoNewline -ForegroundColor DarkMagenta
        }
        UNKNOWN { 
            Write-Host "$($pStatus)" -NoNewline -ForegroundColor DarkCyan
        }
        Default {
            Write-Host "$($pStatus)" -NoNewline
        }
    }
    Write-Host "]"

    Write-Output $pStatus
}

function Add-ServerResults {
    [CmdLetBinding()]
    param(
        [System.String]$LocalServer,
        [System.String]$RemoteServer,
        [System.Int32]$RemotePort,
        [System.String]$PortStatus
    )

    # create a custom object with the desired headers
    $row = New-Object System.Object
    $row | Add-Member -MemberType NoteProperty -Name "Local Server" -Value $LocalServer
    $row | Add-Member -MemberType NoteProperty -Name "Remote Server" -Value $RemoteServer
    $row | Add-Member -MemberType NoteProperty -Name "Port" -Value $RemotePort
    $row | Add-Member -MemberType NoteProperty -Name "Status" -Value $PortStatus

    # store each port as a row in the resultset
    $script:serverResults += $row
}
#endregion

#region Main script execution
$oldErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Stop"
Start-Transcript -Path $script:logPath

# load and parse the json config file 
try {
    Write-Host "Loading config from: $($script:BaseDirectoryPath)$($BaseConfigFile)"
    $script:config = Get-Content $script:BaseDirectoryPath$BaseConfigFile -Raw | ConvertFrom-Json
}
catch {
    Write-Error -Message "Unable to open base JSON config file."
}

$script:serverResults = @()

# get outbound mappings for the local server. Allows for distributing one central json file.
$server = $script:config.Servers | Where-Object { $_.Name -eq $localServer }
if ($server) {
    Write-Host "Checking connection from source server: $($localServer)"

    # process each destination server
    foreach ($connection in $server.Outbound) {
        Write-Host "Checking outbound connections to destination server: $($connection.Name)"
		
        # ping every port in the mapping
        foreach ($port in $connection.Ports) {
            $portStatus = Test-PortStatus -RemoteHost $connection.Name -RemotePort $port
            Add-ServerResults -LocalServer $localServer -RemoteServer $connection.Name -RemotePort $port -PortStatus $portStatus
        }
    }
}
else {
    Write-Host "No configuration present for the local server."
}

# output the results according to the provided param
switch -Wildcard ($Output.ToLower()) {
    "csv" {	$serverResults | Export-CSV -Path $LogFolderPath$script:csvOutPutFile }
    "html" {
        # add basic page style configuration
        $style = "<style>
		BODY{background-color:darkslategrey;color:aliceblue;}
		TABLE{border-width: 1px;border-style: solid;border-color: darkgrey;border-collapse: collapse;}
		TH{border-width: 1px;padding: 0px;border-style: solid;border-color: darkgrey;background-color:brown}
		TD{border-width: 1px;padding: 0px;border-style: solid;border-color: darkgrey;background-color:lightslategrey}
		</style>"

        # convert the results to html with rudimentary style
        $html = $serverResults | ConvertTo-Html -Title "$($script:localServer): Firewall Port Status" -Head $style -Body "<H2>$($script:cDate)</H2>" -Post "For details, contact the IT Service Center." -Pre "<p>Automatically generated by <strong>$($script:localServer)</strong>:</p>
		<p>Status legenda:</p>
		<p>LISTENING - PING OK - SERVICE OK</p>
		<p>OPEN - PING OK - SERVICE NOTOK</p>
		<p>CLOSED - PING NOTOK - SERVICE NOTOK</p>
		<p>STRICT - PING NOTOK - SERVICE OK</p>
		<p>UNKNOWN - PING ? - SERVICE ?</p>"

        # add conditional formatting rules
        $cStyle = @{}
        $cStyle.Add('$_.Value -eq "LISTENING"', ("Status", "background-color:green"))
        $cStyle.Add('$_.Value -eq "OPEN"', ("Status", "background-color:orange"))
        $cStyle.Add('$_.Value -eq "CLOSED"', ("Status", "background-color:red"))
        $cStyle.Add('$_.Value -eq "UNKNOWN"', ("Status", "background-color:purple"))
        $cStyle.Add('$_.Value -eq "STRICT"', ("Status", "background-color:blue"))

        $styledHtml = ConvertTo-HtmlConditionalFormat -PlainHtml $html -ConditionalStyle $cStyle
        $styledHtml | Out-File $HtmlOutputFolderPath$script:htmlOutPutFile -Force
    }
    default { Write-Output $serverResults }
}

Stop-Transcript
$ErrorActionPreference = $oldErrorActionPreference
#endregion