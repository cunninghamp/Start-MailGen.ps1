<#
.SYNOPSIS
Start-MailGen.ps1 - Test lab email traffic generation script.

.DESCRIPTION 
NOTE: THIS IS ONLY FOR USE IN TEST LAB ENVIRONMENTS

Generates email traffic within a test lab enviroment
between randomly selected mailboxes to assist with 
simulating different real world scenarios and
administration tasks.

The script is executed from the Exchange Management Shell
and requires the Active Directory remote administration tools
and the Exchange Web Services Managed API to be installed
on the server or workstation running the script.

Please refer to the installation instructions at:
http://exchangeserverpro.com/exchange-2010-test-lab-email-script

.EXAMPLE
.\Start-MailGen.ps1
Begins executing the script.

.LINK
Installation instructions:
http://exchangeserverpro.com/exchange-2010-test-lab-email-script

.NOTES
Written By: Paul Cunningham
Website:	http://exchangeserverpro.com
Twitter:	http://twitter.com/exchservpro

Additional Credits:
Dave - for pointing out random mailbox selection error
http://exchangeserverpro.com/test-lab-email-traffic-generator-powershell-script/#comment-31162

Change Log
V1.00, 16/4/2012 - Initial version
V1.01, 22/01/2014 - Fixed error with random mailbox selection that excluded last mailbox
                  - Distribution groups are now included as a possible recipient of emails
                  - Added code to load Exchange snapin so script can run from regular PowerShell prompt
                  - Recipient list is refreshed on every loop so newly created recipients will be included
                  - Added random file attachments, will occur in about 30% of messages generated
                  - Messages can be sent to multiple recipients, between 1 and $maxrecipients
#>

#requires -version 2


#-----------------------------------------
# 				References
#-----------------------------------------
#
# The following URLs are credited as references for this script, either for direct
# code reuse, provision of dummy data, or general inspiration.
#
# Glen Scales
# URL: http://gsexdev.blogspot.com.au/
#
# Mike Pfeiffer
# URL: http://www.mikepfeiffer.net/2010/04/sending-email-with-powershell-and-the-ews-managed-api/
#
# Steve Goodman
# URL: http://www.stevieg.org/2010/07/using-powershell-to-import-contacts-into-exchange-and-outlook-live/
#
# MSDN ConnectingIdType Enumeration
# URL: http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.connectingidtype(v=exchg.80).aspx
#
# EWS Managed API
# URL: http://www.microsoft.com/downloads/details.aspx?FamilyID=C3342FB3-FBCC-4127-BECF-872C746840E1&amp;displaylang=en&displaylang=en
#
# Dictionary file by Scott Hanselman
# URL: http://www.hanselman.com/blog/DictionaryPasswordGeneratorInPowershell.aspx
#
# Long text file via Project Gutenberg
# URL: http://gutenberg.net.au/ebooks04/0400191.txt
#
# Adding attachments via EWS
# URL: http://chrisbitting.com/2013/09/08/adding-attachments-to-an-email-using-exchange-web-services/?relatedposts_exclude=135
#
# RBAC Command:
# New-ManagementRoleAssignment -Name:impersonationAssignmentName -Role:ApplicationImpersonation -User:serviceaccount
#

[CmdletBinding()]
param ()


#-----------------------------------------
# 				Variables
#-----------------------------------------

#Current directory
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

#Path to the Exchange Web Services DLL
$dllfolderpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"

# Email subject line length min/max values
$minlength = 1
$maxlength = 5

# Email message body min/max lengths
$minlines = 5
$maxlines = 50

# Max recipients for a message
$maxrecipients = 5

# Dictionary file used for random subject lines
$dictionary = "$($myDir)\dict.csv"

# Text file used for random message bodies
$textfile = "$($myDir)\longtext.txt"


#-----------------------------------------
# 				Functions
#-----------------------------------------

# This function generates the email subject lines
# using the dictionary file and a random number generator
function EmailSubject {

	$subjectLength = Get-Random -Minimum $minlength -Maximum $maxlength
	[string]$subject = ""

	$i = 0
	do {

		$rand = Get-Random -Minimum 0 -Maximum ($wordcount - 1)
		$word = ($words.GetValue($rand)).Word
			
		$subject = $word + " " + $subject
		
		$i++
	}
	while ($i -lt $subjectLength)

	# Capitalise the first character of the subject line
	$subject = $subject.substring(0,1).ToUpper()+$subject.substring(1)
	
	return $subject
}

# This function generates the email message body text
# using the long text file and a random number generator
function EmailBody {
	$lineLength = Get-Random -Minimum $minlines -Maximum $maxlines
	[string]$body = ""

	$i = 0
	do {
		$rand = Get-Random -Minimum 0 -Maximum ($textlength - 1)
		$line = ($longtext.GetValue($rand))
		
		$body = $line + " " + $body
			
		$i++
	}
	while ($i -lt $lineLength)

	return $body
}

# This function uses a random number generator to decide
# whether to attach a file, and choose from one of the
# sample files available
function PickAttachment {
    
    $rand = Get-Random -Minimum 0 -Maximum 10
    if ($rand -gt 7)
    {
        $attachfile = $true
    }
    else
    {
        $attachfile = $false
    }

    if ($attachfile)
    {
        $files = @(Get-ChildItem $myDir\Attachments | where { ! $_.PSIsContainer })
        $filecount = $files.Count
        $filepick = Get-Random -Minimum 0 -Maximum ($filecount)
        $file = $files.GetValue($filepick)
    }

    return $file
}


# This function uses a random number generator to pick
# the recipient
function PickRecipient {

	$rand = Get-Random -Minimum 0 -Maximum ($recipientcount)
	$name = $recipients.GetValue($rand)
	
	return $name
}


# This function uses a random number generator to pick
# the sender
function PickSender {

	$rand = Get-Random -Minimum 0 -Maximum ($mailboxcount)
    $name = $mailboxes.GetValue($rand)
	
	return $name
}


# This function sends the email message via EWS with impersonation
# and based on email subject, body, sender and recipient details
# that are randomly generated
function SendMail {

    Write-Host "*** New email message"

    #Generate subject, body and attachment
	$emailSubject = EmailSubject
	$emailBody = EmailBody
    $emailAttachment = PickAttachment

    #Choose sender for email
	$sender = PickSender

	$SenderSmtpAddress = (Get-Mailbox $sender).PrimarySMTPAddress
	if($SenderSmtpAddress.GetType().fullname -eq "Microsoft.Exchange.Data.SmtpAddress") {
	    $EmailSender = $SenderSmtpAddress.ToString()
	}
	else {
	    $EmailSender = $SenderSmtpAddress
	}

    Write-Host "Sender: $EmailSender"
    Write-Host "Subject: $EmailSubject"
    if ($emailAttachment)
    {
        Write-Host "Attachment: $emailAttachment"
    }

    #Impersonate sender
	$impersonate = $SenderSmtpAddress

	$ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$impersonate
	$service.ImpersonatedUserId = $ImpersonatedUserId

    #Start new mail item with sender, subject, body
	$mail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
	$mail.Subject = $EmailSubject
	$mail.Body = $EmailBody
	$mail.From = $EmailSender

    #Choose recipients, and make sure recipients and sender don't match.
    #A random number of recipients will be added to the message. 
    $tocount = Get-Random -Minimum 1 -Maximum ($maxrecipients + 1)

    $i = 0
    do {
	    $recipient = PickRecipient
	    if ($recipient -eq $sender)
	    {
		    do { $recipient = PickRecipient }
		    while ($recipient -eq $sender)
	    }

	    $RecipientSmtpAddress = (Get-Recipient $recipient).PrimarySMTPAddress
	    if($RecipientSmtpAddress.GetType().fullname -eq "Microsoft.Exchange.Data.SmtpAddress") {
	        $EmailRecipient = $RecipientSmtpAddress.ToString()
	    }
	    else {
	        $EmailRecipient = $RecipientSmtpAddress
	    }

        Write-Host "Recipient: $EmailRecipient"
    
        #Add recipient to mail item
	    [Void] $mail.ToRecipients.Add($EmailRecipient)

        $i++
    }
    while ($i -lt $tocount)

    #Add the attachment
    if ($emailAttachment)
    {
        #$attachmentPath = $myDir\Attachments\$emailAttachment"
        $mail.Attachments.AddFileAttachment("$myDir\Attachments\$emailAttachment")
    }

    #Send the message
    $mail.SendAndSaveCopy()

}


#-----------------------------------------
# 				Script
#-----------------------------------------

# Check that all of the required files are present
if (Test-Path $dllfolderpath)
{
	Add-Type -Path $dllfolderpath
}
else
{
	Write-Host -ForegroundColor Yellow "Unable to locate Exchange Web Services DLL."
	EXIT
}

if (Test-Path $dictionary)
{
	$words = @(Import-Csv $dictionary)
	$wordcount = $words.count
}
else
{
	Write-Host -ForegroundColor Yellow "Unable to locate dictionary file $dictionary."
	EXIT
}

if (Test-Path $textfile)
{
	$longtext = Get-Content $textfile
	$textlength = $longtext.count
}
else
{
	Write-Host -ForegroundColor Yellow "Unable to locate text file $textfile."
	EXIT
}


#Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	try
	{
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	}
	catch
	{
		#Snapin was not loaded
		Write-Warning $_.Exception.Message
		EXIT
	}
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}


# The Active Directory module is required to retrieve the SID of the
# service account to determine the AutoDiscover URL
Import-Module ActiveDirectory


#Web services initialization
Write-Host "Preparing EWS"
$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$user = [ADSI]"LDAP://<SID=$sid>"
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService("Exchange2007_SP1")
$service.AutodiscoverUrl($user.Properties.mail)


#Start infinite loop
Write-Host -ForegroundColor White "Starting email generation loop"
do {

	#Calculate number of emails to send this hour
	[int]$hour = Get-Date -Format HH

	# You can modify these values to vary the number of emails that the
	# script will send each hour
	Switch($hour)
	{
		06 {$sendcount = 5}
		07 {$sendcount = 8}
		08 {$sendcount = 60}
		09 {$sendcount = 60}
		10 {$sendcount = 80}
		11 {$sendcount = 10}
		12 {$sendcount = 60}
		13 {$sendcount = 60}
		14 {$sendcount = 80}
		15 {$sendcount = 80}
		16 {$sendcount = 90}
		17 {$sendcount = 4}
		18 {$sendcount = 2}
		19 {$sendcount = 8}
		default {$sendcount = 10}
	}

	# On Saturday/Sunday the sendcount is set to 10 no matter which hour of the day
	[string]$dayofweek = (Get-Date).Dayofweek
	Switch($dayofweek)
	{
		"Saturday"{$sendcount = 10}
		"Sunday" {$sendcount = 10}
	}

	# A random number is added to the message count to vary
	# the results from day to day
	$rand = Get-Random -Minimum 1 -Maximum 99
	$sendcount = $sendcount + $rand

	Write-Host -ForegroundColor White "*** Will send $sendcount emails this hour"

    #Get list of mailbox users and distribution groups
    $recipients = @()
    Write-Host -ForegroundColor White "Retrieving recipient list"
    $mailboxes = @(Get-Mailbox -RecipientTypeDetails UserMailbox -resultsize Unlimited | Where {$_.Name -ne "Administrator" -and $_.Name -notlike "extest_*"})
    $mailboxcount = $mailboxes.Count

    Write-Host "$mailboxcount mailboxes found"

    $distgroups = @(Get-DistributionGroup -resultsize Unlimited)
    $distgroupcount = $distgroups.Count
    
    Write-Host "$distgroupcount distribution groups found"
    
    $recipients += $mailboxes
    $recipients += $distgroups

    $recipientcount = $recipients.count

    Write-Host "$recipientcount total recipients found"

	# Send all the emails
	$sent = 0
	do {
		$pct = $sent/$sendcount * 100
		Write-Progress -Activity "Sending $sendcount emails" -Status "$sent of $sendcount" -PercentComplete $pct

		#SendMail $emailSubject $emailBody $emailAttachment
        SendMail
		$sent++

	}
	until ($sent -eq $sendcount)
	Write-Host -ForegroundColor White "*** Finished sending $sendcount emails for hour $hour"
	
	# Check if there is any time still left in this hour
	# and sleep if there is
	[int]$endhour = Get-Date -Format HH
	if ($hour -lt 23)
	{
		[int]$nexthour = $hour + 1
		
			do {
				Write-Progress -Activity "Waiting for next hour to start" -Status "Sleeping..." -PercentComplete 0
				Write-Host -ForegroundColor Yellow "Not next hour yet, sleeping for 5 minutes"
				Start-Sleep 300
				[int]$endhour = Get-Date -Format HH
			}
			until($endhour -ge $nexthour)
		
	}
	else
	{
		[int]$nexthour = 0
		
			do {
				Write-Progress -Activity "Waiting for next hour to start" -Status "Sleeping..." -PercentComplete 0
				Write-Host -ForegroundColor Yellow "Not next hour yet, sleeping until hour $nexthour"
				Start-Sleep 300
				[int]$endhour = Get-Date -Format HH
			}
			until($endhour -eq $nexthour)
	}

}
until ($forever)
# The script will run forever until terminated by the administrator, a log off, or server restart