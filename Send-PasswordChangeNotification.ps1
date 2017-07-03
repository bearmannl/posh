<#
.SYNOPSIS
    Get all users with a soon to expire password and send them a reminder to do so.

.DESCRIPTION
    This script gets all users with a soon to expire password and sends them a reminder to do so. This script was originally created by Robert Pearman and posted on the technet gallery. Since modified by Mike Beerman to include changes specific for Silverside. Settings such as the destination SMTP server and From address can only be configured within the script.

.PARAMETER ExpireInDays
    Optional input allowing an override to the number of days at which the script will send a reminder. Defaults to 30 days.

.PARAMETER Logging
    Switch which determines if a seperate CSV logfile should be created. Default to false without the switch.

.PARAMETER LogPath
    Folderpath for the logfile, plain string format.

.PARAMETER Testing
    Switch which determines if the emails should be sent to the test recipient instead of the actual users. Default to false without the switch.

.PARAMETER TestRecipient
    Optional string for setting a different recipient in testing for testing purposes. Will not be used if Testing is not included in the command or email is disabled.

.PARAMETER TestMailDisabled
    Switch for disabling email functionality.

.EXAMPLE
Send-PasswordChangeNotification.ps1
Send-PasswordChangeNotification.ps1 -Expire 60 -Logging -Logile "C:\Logs\PasswordChangeNotifications.csv" -Testing -TestRecipient "yourname@yourdomain.com"

.LINK
    https://gallery.technet.microsoft.com/Password-Expiry-Email-177c3e27
    http://bearman.nl/2016/06/23/password-change-notification-script

.NOTES
    Authors : Mike Beerman
    Company : Silverside
    Date : 2016-06-23
    Version : 2.0

    Authors : Robert Pearman (WSSMB MVP)
    Company : Windows Server Essentials
    Date : 2016-02-03
    Version : 1.4
#>

Param (
    [ValidateRange(1,365)][int]$ExpireInDays = 30,
    [switch]$Logging,
    [string]$LogPath = "C:\Logs",
    [switch]$Testing,
    [string]$TestRecipient = "yourname@yourdomain.com",
    [switch]$TestMailDisabled
)

##
# Config values
##

$smtpServer = "mail.yourdomain.local"
$from = "Your Company IT <noreply@yourdomain.com>"
$textEncoding = [System.Text.Encoding]::UTF8
$logDate = Get-Date -Format yyyyMMdd
$LogPath = ($LogPath).TrimEnd("\")
$logFile = $LogPath + "\PasswordChangeNotifications.csv"

##
#Initiate PowerShell session
##

# Set threading
$ver = $host | Select-Object Version
if($ver.Version.Major -gt 1) {$Host.Runspace.ThreadOptions = "ReuseThread"}

# Load Active Directory PowerShell snap-in
Import-Module ActiveDirectory
Write-Verbose ("[{0}] - AD PowerShell SnapIn loaded" -f (Get-Date -Format T))

$verbosePreference = 'Continue'

# Check Logging Settings
if ($Logging) {
    # Test Log Folder Path and create new directory if false
    if (!(Test-Path $LogPath)) {
        # Create CSV File and Headers
        $newItem = New-Item $LogPath -ItemType Directory
        Write-Verbose ("[{0}] - New folder created at [{1}]" -f (Get-Date -Format T), $LogPath)
    }
    # Test Log File Path and create new file if false
    if (!(Test-Path $logFile)) {
        # Create CSV File and Headers
        $newItem = New-Item $logFile -ItemType File
        Add-Content $logFile "Date,Name,EmailAddress,DaystoExpire,ExpiresOn,Notified"
        Write-Verbose ("[{0}] - New logfile created at [{1}]" -f (Get-Date -Format T), $logFile)
    }
} # End Logging Check

$DefaultmaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
$users = Get-ADUser -Filter * -Properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress | Where-Object { $_.Enabled -eq $true -and $_.PasswordNeverExpires -eq $false -and $_.passwordexpired -eq $false -and $_.UserPrincipalName -notlike 'adm*' }

# Process Each User for Password Expiry
foreach ($user in $users) {
    $name = $user.Name
    $emailAddress = $user.EmailAddress
    $passwordSetDate = $user.PasswordLastSet
    $PasswordPol = (Get-ADUserResultantPasswordPolicy $user)
    $sent = "" # Reset Sent Flag
    # Check for Fine Grained Password
    if (($PasswordPol) -ne $null) {
        $maxPasswordAge = ($PasswordPol).MaxPasswordAge
    }
    else {
        # No FGP set to Domain Default
        $maxPasswordAge = $DefaultmaxPasswordAge
    }


    $expiresOn = $passwordSetDate + $maxPasswordAge
    $today = (Get-Date)
    $daysToExpire = (New-TimeSpan -Start $today -End $expiresOn).Days

    # Set Greeting based on Number of Days to Expiry.

    # Check Number of Days to Expiry
    $messageDays = $daysToExpire

    if (($messageDays) -gt "1") {
        $messageDays = "in " + "$daysToExpire" + " days."
    }
    else {
        $messageDays = "today."
    }

    # Email Subject Set Here
    $subject="Your password will expire $messageDays"

    # Email Body Set Here, Note You can use HTML, including Images.
    $body ="
    Dear $name,
    <p> Your Password will expire $messageDays<br>
    To change your password on a PC press CTRL ALT Delete and choose Change Password <br>
    <p>Thanks, <br> 
    Your Comany IT Administrators
    </P>"

    # If Testing Is Enabled - Email Administrator
    if ($Testing) {
        $emailAddress = $TestRecipient
    } # End Testing

    # If a user has no email address listed
    if (($emailAddress) -eq $null) {
        $emailAddress = $TestRecipient 
    }# End No Valid Email

    # Send Email Message
    if (($daysToExpire -eq '0') -or ($daysToExpire -eq '14') -or ($daysToExpire -eq $ExpireInDays)) {
        $sent = "Yes"
        # If Logging is Enabled Log Details
        if ($Logging) {
            Add-Content $logFile "$logDate,$name,$emailAddress,$daysToExpire,$expiresOn,$sent"
            Write-Verbose ("[{0}] - Entry added to logfile for user [{1}]" -f (Get-Date -Format T), $user.UserPrincipalName)
        }
        # Send Email Message if not disabled by switch
        if (!($TestMailDisabled)) {
            Send-Mailmessage -SmtpServer $smtpServer -From $from -To $emailAddress -Subject $subject -Body $body -BodyAsHTML -Priority High -Encoding $textEncoding
            Write-Verbose ("[{0}] - Email sent to [{1}]" -f (Get-Date -Format T), $emailAddress)
        }
    } # End Send Message
    else {
        # Log Non Expiring Password
        $sent = "No"
        # If Logging is Enabled Log Details
        if ($Logging) {
            Add-Content $logFile "$logDate,$name,$emailAddress,$daysToExpire,$expiresOn,$sent"
            Write-Verbose ("[{0}] - Entry added to logfile for user [{1}]" -f (Get-Date -Format T), $user.UserPrincipalName)
        }
    }
} # End User Processing

$verbosePreference = 'SilentlyContinue'
# End