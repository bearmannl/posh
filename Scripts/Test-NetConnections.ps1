<#
.SYNOPSIS
     Tests status of firewall ports and outputs to various formats.
      
.DESCRIPTION
     This script checks if specified ports are open *from* local server *to* remote servers, based on the server and port mapping in the .json file. Possible output formats include .csv and .htm.

.PARAMETER Verbs
	Optional switch adds Verbose output to track script progress.

.PARAMETER Output
	Required parameter, determines the desired type of output.

.PARAMETER TestData
	Optional switch, portscans can take a while, in case of debugging or editing styling of the script, use test data.

.EXAMPLE
     Test-NetConnections.ps1 -Output CSV
	 Test-NetConnections.ps1 -Output HTML -Verbs
	 Test-NetConnections.ps1 -Output HTML -Verbs -TestData
      
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
#>
Param(
	[Parameter(Mandatory=$true)]
	[ValidateSet('CSV','HTML')]
	[System.String]$Output,
	[Switch]$Verbs,
	[Switch]$TestData
)

#region basic script setup
$script:localServer = $env:COMPUTERNAME

#set location to store migration logs
$script:logPathFolder = "D:\Logs\Scripts"

if(!(test-path $script:logPathFolder)) { New-Item -ItemType directory -Path $script:logPathFolder }
$script:cDate = (Get-Date)
$script:logPath = "$script:logPathFolder\Test-NetConnectionsOutput_$($script:localServer)_{0:yyyy-MM-dd_HHmmss}.rtf" -f ($script:cDate)

$script:BaseDirectory = Get-Location
$script:BaseConfig = "\Test-NetConnections.json"

$script:htmlOutputFolder = "C:\inetpub\wwwroot\health80"
$script:htmlOutPutFile = "\default.htm"
$script:csvOutPutFile = "\Test-NetConnectionsOutput_$($script:localServer)_{0:yyyy-MM-dd_HHmmss}.csv" -f ($script:cDate)

#endregion

#region Functions
function Write-VerboseCustom {
	[CmdLetBinding()]
	Param(
		[Parameter(Mandatory=$true)]
		[String]$Message
	)
	# to provide verbose feedback to users, only on select statements
	if($Verbs) {
		$oldVerbosePreference = $VerbosePreference
		$VerbosePreference = "Continue"
		Write-Verbose $Message
		$VerbosePreference = $oldVerbosePreference
	}
}

function ConvertTo-HtmlConditionalFormat {
	[CmdLetBinding()]
	Param(
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
	foreach($cs in $ConditionalStyle.Keys) {
		$scriptBlock = [scriptblock]::Create($cs)
		# find the column matching the correct content value
		$columnIndex = (($xml.Descendants("{$namespace}th") | Where-Object { $_.Value -eq $ConditionalStyle.$cs[0] }).NodesBeforeSelf() | Measure-Object).Count
		# only select the elements matching the correct column
		$elements | Where-Object { ($_.NodesBeforeSelf() | Measure-Object).Count -eq $columnIndex } | ForEach-Object {
			if(&$scriptBlock) {
				# apply the actual attribute value
				$_.SetAttributeValue( "style", $ConditionalStyle.$cs[1])
			}
		}
	}
	
	Write-Output $xml.ToString()
}

function Test-PortStatus {
	[CmdLetBinding()]
	Param(
		$RemoteHost,
		$RemotePort
	)
	$testConnection = Test-NetConnection -Computername $RemoteHost -Port $RemotePort -WarningAction:SilentlyContinue
	if($testConnection.PingSucceeded -AND $testConnection.TcpTestSucceeded){ $pStatus = "LISTENING" } 
		elseif ($testConnection.PingSucceeded -AND !$testConnection.TcpTestSucceeded) { $pStatus = "OPEN" }
		elseif (!$testConnection.PingSucceeded -AND !$testConnection.TcpTestSucceeded) { $pStatus = "CLOSED" }
		elseif (!$testConnection.PingSucceeded -AND $testConnection.TcpTestSucceeded) { $pStatus = "STRICT" }
		else { $pStatus = "UNKNOWN" }
	Write-VerboseCustom "$($RemotePort) $($pStatus)"

	Write-Output $pStatus
}

function Add-ServerResults {
	[CmdLetBinding()]
	Param(
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
Start-Transcript -Path $script:logPath

# load and parse the json config file 
try { $script:config = Get-Content $script:BaseDirectory$script:BaseConfig -Raw -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue | ConvertFrom-Json -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue }
catch { Write-Error -Message "Unable to open base config file" }

$script:serverResults = @()

if(!$TestData) {
	# loop through all configured server mappings, execute when on the correct server. Allows for distributing one central json file.
	foreach($server in $script:config.Servers) {
		if($server.Name -eq $localServer) {
			Write-VerboseCustom "Checking connection from source server: $localServer"

			# process each destination server
			foreach($connection in $server.Outbound) {
				Write-VerboseCustom "Checking outbound connections to destination server: $($connection.Name)"
				
				# ping every port in the mapping
				foreach($port in $connection.Ports) {
					$portStatus = Test-PortStatus -RemoteHost $connection.Name -RemotePort $port
					Add-ServerResults -LocalServer $localServer -RemoteServer $connection.Name -RemotePort $port -PortStatus $portStatus
				}
			}
		}
	}
} else {
	# portscans can take a while, in case of debugging or editing styling, use test data
	Add-ServerResults -LocalServer "SERVER01" -RemoteServer "SERVER02" -RemotePort 80 -PortStatus "OPEN"
	Add-ServerResults -LocalServer "SERVER01" -RemoteServer "SERVER02" -RemotePort 443 -PortStatus "CLOSED"
	Add-ServerResults -LocalServer "SERVER01" -RemoteServer "SERVER03" -RemotePort 80 -PortStatus "LISTENING"
	Add-ServerResults -LocalServer "SERVER01" -RemoteServer "SERVER03" -RemotePort 443 -PortStatus "STRICT"
}

# output the results according to the provided param
switch -Wildcard ($Output.ToLower()) {
	"csv" {	$serverResults | Export-CSV -Path $script:logPathFolder$script:csvOutPutFile }
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
		$cStyle.Add('$_.Value -eq "LISTENING"',("Status","background-color:green"))
		$cStyle.Add('$_.Value -eq "OPEN"',("Status","background-color:orange"))
		$cStyle.Add('$_.Value -eq "CLOSED"',("Status","background-color:red"))
		$cStyle.Add('$_.Value -eq "UNKNOWN"',("Status","background-color:purple"))
		$cStyle.Add('$_.Value -eq "STRICT"',("Status","background-color:blue"))

		$styledHtml = ConvertTo-HtmlConditionalFormat -PlainHtml $html -ConditionalStyle $cStyle
		$styledHtml | Out-File $script:htmlOutputFolder$script:htmlOutPutFile -Force
	}
	default { Write-Output $serverResults }
}

Stop-Transcript
#endregion