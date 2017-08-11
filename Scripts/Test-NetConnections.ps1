<#
.SYNOPSIS
     Tests status of firewall ports and outputs to various formats.
      
.DESCRIPTION
     This script checks if specified ports are open *from* local server *to* remote servers, based on the server and port mapping in the .json file. Possible output formats include .csv and .htm.

.PARAMETER Verbs
	Optional switch adds Verbose output to track script progress.

.PARAMETER Output
	Required parameter, determines the desired type of output.

.EXAMPLE
     Test-NetConnections.ps1 -Output CSV
	 Test-NetConnections.ps1 -Verbs -Output HTML
      
.LINK
    http://www.rubicon.nl
      
.NOTES
     Authors	: Mike Beerman
     Company	: Rubicon
     Date		: 2017-08-08
     Version	: 1.0
					1.1 | MBE | 2017-08-10 | Added HTML output, including function, param and switch.
					1.2 | MBE | 2017-08-11 | Added comments.
#>
Param(
	[Parameter(Mandatory=$true)]
	[ValidateSet('CSV','HTML')]
	[System.String]$Output,
	[Switch]$Verbs
)

#region basic script setup
$script:localServer = $env:COMPUTERNAME

# Set location to store migration logs
$script:logPathFolder = "D:\Logs\Scripts"

if(!(test-path $script:logPathFolder)) { New-Item -ItemType directory -Path $script:logPathFolder }
$cDate = (Get-Date)
$script:logPath = "$script:logPathFolder\Test-NetConnectionsOutput_$($script:localServer)_{0:yyyy-MM-dd_HHmmss}.rtf" -f ($cDate)

$script:BaseDirectory = Get-Location
$script:BaseConfig = "\Test-NetConnections.json"

$script:htmlOutputFolder = "C:\inetpub\wwwroot\health80"
$script:htmlOutPutFile = "\default.htm"
$script:csvOutPutFile = "\Test-NetConnectionsOutput_$($script:localServer)_{0:yyyy-MM-dd_HHmmss}.csv" -f ($cDate)

#endregion

#region Functions
function Write-VerboseCustom {
[CmdLetBinding()]
Param(
	[Parameter(Mandatory=$true)]
	[String]$Message
)
	#to provide verbose feedback to users, only on select statements
	if($Verbs) {
		$oldVerbosePreference = $VerbosePreference
		$VerbosePreference = "Continue"
		Write-Verbose $Message
		$VerbosePreference = $oldVerbosePreference
	}
}

function ConvertTo-HtmlConditionalFormat{
[CmdLetBinding()]
Param(
	$PlainHtml,
	$ConditionalStyle
)
	Add-Type -AssemblyName System.Xml.Linq
	#load HTML content as XML
	$xml = [System.Xml.Linq.XDocument]::Parse($PlainHtml)
	$namespace = 'http://www.w3.org/1999/xhtml'
	#select the type of html elements for processing
	$elements = $xml.Descendants("{$namespace}td")
	#loop through each conditional formatting rule
	foreach($cs in $ConditionalStyle.Keys) {
		$scriptBlock = [scriptblock]::Create($cs)
		#find the column matching the correct content value
		$columnIndex = (($xml.Descendants("{$namespace}th") | Where-Object { $_.Value -eq $ConditionalStyle.$cs[0] }).NodesBeforeSelf() | Measure-Object).Count
		#only select the elements matching the correct column
		$elements | Where-Object { ($_.NodesBeforeSelf() | Measure-Object).Count -eq $columnIndex } | ForEach-Object {
			if(&$scriptBlock) {
				#apply the actual attribute value
				$_.SetAttributeValue( "style", $ConditionalStyle.$cs[1])
			}
		}
	}
	
	Write-Output $xml.ToString()
}
#endregion

#region Main script execution
Start-Transcript -Path $script:logPath

#load and parse the json config file 
try { $script:config = Get-Content $script:BaseDirectory$script:BaseConfig -Raw -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue | ConvertFrom-Json -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue }
catch { Write-Error -Message "Unable to open base config file" }

$serverResults = @()

#loop through all configured server mappings, execute when on the correct server. Allows for distributing one central json file.
foreach($server in $script:config.Servers) {
	if($server.Name -eq $localServer) {
		Write-VerboseCustom "Checking connection from source server: $localServer"

		#process each destination server
		foreach($connection in $server.Outbound) {
			Write-VerboseCustom "Checking outbound connections to destination server: $($connection.Name)"
			
			#ping every port in the mapping
			foreach($port in $connection.Ports) {
				$testConnection = Test-NetConnection -Computername $connection.Name -Port $port -WarningAction:SilentlyContinue
				if($testConnection.TcpTestSucceeded) { $portStatus = "OPEN" } else { $portStatus = "CLOSED" }
				Write-VerboseCustom "$($connection.Name):$($port) - $($portStatus)"

				#create a custom object with the desired headers
				$row = New-Object System.Object
				$row | Add-Member -MemberType NoteProperty -Name "Local Server" -Value $localServer
				$row | Add-Member -MemberType NoteProperty -Name "Remote Server" -Value $connection.Name
				$row | Add-Member -MemberType NoteProperty -Name "Port" -Value $port
				$row | Add-Member -MemberType NoteProperty -Name "Status" -Value $portStatus

				#store each port as a row in the resultset
				$serverResults += $row
			}
		}
	}
}

#output the results according to the provided param
switch -Wildcard ($Output.ToLower()) {
	"csv" {	$serverResults | Export-CSV -Path $script:logPathFolder$script:csvOutPutFile }
	"html" {
		#add basic page style configuration
		$style = "<style>"
		$style = $style + "BODY{background-color:peachpuff;}"
		$style = $style + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
		$style = $style + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
		$style = $style + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
		$style = $style + "</style>"

		#convert the results to html with rudimentary style
		$html = $serverResults | ConvertTo-Html -Title "$($script:localServer): Firewall Port Status" -Head $style -Body "<H2>$($cDate)</H2>" -Pre "<P>Automatically generated by <strong>$($script:localServer)</strong>:</P>" -Post "For details, contact the IT Service Center."

		#add conditional formatting rules
		$cStyle = @{}
		$cStyle.Add('$_.Value -eq "OPEN"',("Status","background-color:green"))
		$cStyle.Add('$_.Value -eq "CLOSED"',("Status","background-color:red"))

		$styledHtml = ConvertTo-HtmlConditionalFormat -PlainHtml $html -ConditionalStyle $cStyle
		$styledHtml | Out-File $script:htmlOutputFolder$script:htmlOutPutFile -Force
	}
	default { Write-Output $serverResults }
}

Stop-Transcript
#endregion