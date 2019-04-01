<#
.SYNOPSIS
     Downloads and starts the latest PowerShell client install file direct from GitHub.
      
.DESCRIPTION
     Downloads and starts the latest PowerShell client install file direct from GitHub.

.PARAMETER Architecture
	The desired architecture of the downloaded client, if not provided, script will attempt to auto-select.

.PARAMETER AutoRun
    If used, the install file will automatically run after download. Please note that this will not work in combination with binary file downloads.
    
.PARAMETER Binary
    Download the binaries instead of an installer.

.PARAMETER Force
    Install download and run the latest version installer, regardless of the version this script is run from.

.EXAMPLE
     Get-LatestPSCore.ps1
     Get-LatestPSCore.ps1 -Architecture WinX64 -Binary
     Get-LatestPSCore.ps1 -Architecture WinX64 -AutoRun -Force
      
.LINK
    http://bearman.nl
      
.NOTES
     Authors	: Mike Beerman
     Date		: 2019-04-01
     Version	: 1.0
#>
[CmdletBinding()]
param (
    [ValidateSet('WinX84', 'WinX64', 'WinArmX84', 'WinArmX64')]
    [String]
    $Architecture,
    [Switch]
    $AutoRun,
    [Switch]
    $Binary,
    [Switch]
    $Force
)

if ([string]::IsNullOrEmpty($Architecture)) {
    $arch = ($env:PROCESSOR_ARCHITECTURE).ToLowerInvariant()
    switch ($arch) {
        amd64 { $Architecture = "WinX64" }
        ia64 { $Architecture = "WinX64" }
        arm { $Architecture = "WinArmX64" }
        x86 { $Architecture = "WinX86" }
        Default {
            Write-Host "Unable to determine Processor Architecture, please provide the correct parameter manually."
            exit
        }
    }
    Write-Host "Auto-selecting processor architecture [" -NoNewline
    Write-Host "$Architecture" -ForegroundColor Yellow -NoNewline
    Write-Host "]"
}

Write-Host "Checking for latest stable version of PowerShell, architecture $($Architecture)..." -NoNewline
$currentVersion = "v$($PSVersionTable.PSVersion)"
$rq = Invoke-WebRequest -Uri https://api.github.com/repos/PowerShell/PowerShell/releases/latest
$rqJson = ConvertFrom-Json $rq.Content
$tagName = $rqJson.tag_name
$matchingVersion = $currentVersion -eq $tagName
if ($matchingVersion) {
    if ($Force) {
        $color = [System.ConsoleColor]"Magenta"
    }
    else {
        $color = [System.ConsoleColor]"Green"
    }
    
}
else {
    $color = [System.ConsoleColor]"Yellow"
}
Write-Host " [" -NoNewline
Write-Host "$tagName" -ForegroundColor $color -NoNewline
Write-Host "]"
if ($matchingVersion) {
    if ($Force) {
        Write-Host "Versions match, but running anyway."
    }
    else {
        Write-Host "Versions match, exiting script."
        exit
    }
}
else {
    Write-Host "Versions do not match, downloading latest stable version."
}

switch ($Architecture) {
    WinX84 {
        $arch = "win-x84"
        $executableExt = "msi"
        $archiveExt = "zip"
    }
    WinX64 {
        $arch = "win-x64"
        $executableExt = "msi"
        $archiveExt = "zip"
    }
    WinArmX84 {
        $arch = "win-arm32"
        $executableExt = "zip"
        $archiveExt = "zip"
    }
    WinArmX64 {
        $arch = "win-arm64"
        $executableExt = "zip"
        $archiveExt = "zip"
    }
    WinFxDep {
        $arch = "win-fxdependent"
        $executableExt = "zip"
        $archiveExt = "zip"
    }
    Osx {
        $arch = "osx-x64"
        $executableExt = "pkg"
        $archiveExt = "tar.gz"
    }
    Rhel {
        $arch = "1.rhel.7.x86_64"
        $executableExt = "rpm"
        $archiveExt = "rpm"
    }
    Alpine {
        $arch = "linux-alpine-x64"
        $executableExt = "tar.gz"
        $archiveExt = "tar.gz"
    }
    LinArm32 {
        $arch = "linux-arm32"
        $executableExt = "tar.gz"
        $archiveExt = "tar.gz"
    }
    LinArm64 {
        $arch = "linux-arm64"
        $executableExt = "tar.gz"
        $archiveExt = "tar.gz"
    }
    Lin64FxDep {
        $arch = "linux-x64-fxdependent"
        $executableExt = "tar.gz"
        $archiveExt = "tar.gz"
    }
    Lin64 {
        $arch = "linux-x64"
        $executableExt = "tar.gz"
        $archiveExt = "tar.gz"
    }
    Deb9 {
        $arch = "1.debian.9_amd64"
        $executableExt = "deb"
        $archiveExt = "deb"
    }
    Ubu14 {
        $arch = "1.ubuntu.14.04_amd64"
        $executableExt = "deb"
        $archiveExt = "deb"
    }
    Ubu16 {
        $arch = "1.ubuntu.16.04_amd64"
        $executableExt = "deb"
        $archiveExt = "deb"
    }
    Ubu18 {
        $arch = "1.ubuntu.18.04_amd64"
        $executableExt = "deb"
        $archiveExt = "deb"
    }
}

if ($Binary) {
    $ext = $archiveExt
}
else {
    $ext = $executableExt
}

$package = $rqJson.assets | Where-Object { $_.name -like "*-$($arch).$($ext)" }
Write-Host "File selected: [" -NoNewline
Write-Host "$($package.name)" -NoNewline -ForegroundColor $color
Write-Host "]"

# Download the actual file to disk
Write-Host "Downloading..."
$outPath = "$($env:USERPROFILE)\AppData\Local\Temp\$($package.name)"
Invoke-WebRequest -Uri $package.browser_download_url -OutFile $outPath
Write-Host "File saved: $($outPath)"

if ($AutoRun -and !$Binary) {
    Write-Host "Opted for auto run, starting the MSI"
    msiexec.exe /I $outPath
}