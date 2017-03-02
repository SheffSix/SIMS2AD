<#

.SYNOPSIS
Create Active Directory user accounts for students and staff using reports generated from SIMS.

.NOTES
File Name    :  sims2ad.ps1
Version      :  0.5
Author       :  James Beck

CHANGE LOG
##########
v0.5
	CHANGES
	-------
	. New Azure AD Connect service which replaces DirSync uses different PowerShell modules.
		Script updated to use new powershell modules MsOnline and AdSync instead oof DirSync if they are present.
v0.4.2
	BUG FIXES
	---------
	. Year 14 student leavers are never disabled.
		The code in StudentUserFuntions.ps1 @98 did not check if the year 14 buffer period was in effect. Therefore it always assumed that it was and did not process any year 14 students.
	. The Administrator's synchronisation report is counting removed groups incorrectly.
		The removed groups counter was counting how many group memberships were removed instead of how may groups were removed.
v0.4.1
	BUG FIXES
	---------
	. When removing a group membership, the e-mail report did not show the removed member if $Name was not present
	. Removing students from teaching groups, the call to function RemoveGroupMembership did not specify "$($Member.Name)"
	
v0.4
	CHANGES
	-------
	. Directly setting targetAddress and proxyAddresses to a user account proved unreliable. Now using built-in Exchange command Enable-RemoteMailbox.
	. Commands run inside If (!($Simulate)) {} statements are now followed by Else {}, writing the command to the console/log.
	
v0.3
	CHANGES
	-------
	.New config item, $MSOLDomain. Specifies native domain name of the Microsoft Online organisation.
	
	Bug Fixes
	---------
	. On premises mailboxes were not able to e-mail new MSOL mailboxes. MSOL users are now created with the targetAddress property set. <mailprefix>@SouthHunsley.microsoftonline.com
	. MSOL Mailboxes were created with an incorrect e-mail address. Users are now created with an entry added into the proxyAddresses property. <mailprefix>@southhunsley.org.uk
	
v0.2
	CHANGES
	-------
	. The script will now also create Exchange mailboxes when required.
	. Prefix groups requiring Office 365 mailboxes must now be specified in Config.ps1.
	. Users who do not have Office 365 Mailboxes will have an Exchange mailbox created IF the -ProcessExchange parameter is used.
	. New parameter -NoLeavers. Do not process leavers. New users will be created and existing users will be updated.
	. Moved the change log into the notes section of help.
	. Moved setup instructions into the description section of help.
	. Added some usage examples to help.
	. Added parameter descriptions to help.
v0.1
	CHANGES
	-------
	. Config.ps1 split into two files, Config.ps1 and Include.ps1. Eliminating risk of user inadvertantley chaging something that shouldn't be.
	. $InDev line now commented out by default
	
	Bug Fixes
	---------
	. Office365Sync funciton was checking for $ProcessMSOL incorrectly. I removed this check altoghether as it is only launched from within the ProcessOffice365
		function, which already performs this check.
	. Reports were being e-mailed when no changes had been made. The script was looking for entries in $PrevDisabledUsers when deciding if to send an e-mail or not.
		$PrevDisabledUsers is a list of users who should be disabled, and already have been, not who shouldn't, which is where I was going wrong. I didn't implememnt
		that yet. D'oh!

.DESCRIPTION
This script runs reports from SIMS then creates and configures user accounts and groups based on the results of those reports.

Permissions
###########

The user account running this script must have the following permissions:

Active Directory
----------------
	User OUs   :  Create/delete User objects    on  This object and all decendant objects.
	              Full control                  on  Descendant User objects.
	Group OUs  :  Create/delete Groups objects  on  This object and all decendant objects.
	              Full control                  on  Descendant Group objects.

DirSync Server
--------------
	Group membership        :  Administrators
	
	Local Policy (GPEdit.msc) -
		User Rights Assignment  :  Allow log on locally
			                       Allow log on through remote desktop services
		                           Log on as a batch job

Home Folder Servers
-------------------
	Group Membership  :  Backup Operators Group
	
	Folder permissions -
		Home folder shares  :  Full control     on  This folder, subfolders and file


Msol User Account
#################

Create Account
--------------
	Create an unlicensed user with Exchange administrator and User management administrator roles.

Add credentials to the Windows Credential Manager
-------------------------------------------------
	Log on to the DirSync server as the user running the script and use the CredMan.ps1 script to save the Msol Account credentials:

		.\CredMan.ps1 -AddCred -Target 'https://<MSOL-DOMAINNAME>.microsoftonline.com' -User '<USERNAME>@<MSOL-DOMAINNAME>.onmicrosoft.com' -Pass '<PASSWORD>'

	It is important to use single quotes around the parameter values when using the CredMan.ps1 script.

	Ensure that $OfficeTargetURL in the Config.ps1 file matches what you specify as the -Target parameter.
	
.EXAMPLE
.\sims2ad.ps1 -StudentUsers -DryRun False
Synchronise student accounts

.EXAMPLE
.\sims2ad.ps1 -StudentUsers -ProcessMSOL -ProcessExchange -DryRun False
Synchronise student accounts, assign an Office 365 license to each user and create Exchange mailboxes where required.

.EXAMPLE
.\sims2ad.ps1 -StudentUsers -NoLeavers
Run in Simulation mode. Create new student accounts and update existing ones, do not disable leavers.

.EXAMPLE
.\sims2ad.ps1 -ProcessGroups -RegistrationGroups -SubjectGroups -TeachingGroups -DryRun False
Create registration groups, subject groups and teaching groups. Add memberships to the groups.

.EXAMPLE
.\sims2ad.ps1 -ProcessGroups -CleanupGroups -DryRun False
Create groups of all group types. Add memebrships to the groups. Remove memberships that are no longer valid.

.EXAMPLE
.\sims2ad.ps1 -ProcessGroups -RegistrationGroups -PurgeGroupMembers -DryRun False
Remove all memberships from existing registration groups.

.PARAMETER CleanupGroups
Removes any group memberships which should no longer exist. NOTE: This takes a LONG time to complete and should only be used over weekends/holidays.

.PARAMETER DryRun
By default, the script performs a dry run, which does not make any changes. Run the script using '-DryRun False' to allow the script to make changes.

.PARAMETER ProcessGroups
Perform group synchronisation actions. If no group type is specified, all group types will be processed.

.PARAMETER PurgeGroupMembers
Remove all memberships from groups.

.PARAMETER RegistrationGroups
Perform group sychronisation actions on registration groups.

.PARAMETER StaffGroups
Perform group sychronisation actions on staff groups.

.PARAMETER SubjectGroups
Perform group sychronisation actions on subject groups.

.PARAMETER TeachingGroups
Perform group sychronisation actions on teaching groups.

.PARAMETER StaffUsers
Perform sychronisation actions on staff user accounts.

.PARAMETER StudentUsers
Perform sychronisation actions on student user accounts.

.PARAMETER NoLeavers
Do not mark any users as leavers or disable user accounts.

.PARAMETER ProcessMSOL
Set up users' Office 365 licensing.

.PARAMETER ProcessExchange
Set up users' Exchange mailboxes. This only applies to users in groups who do not have Office 365 mailboxes enabled. This can be changed by editing the Config.ps1 file.

.PARAMETER LogFile
Specify the log file to write to. Must be the full path and file name. By default, logs are saved to $ScriptRoot\Logs and are date stamped.

.PARAMETER NoLog
Do not create a log file.

.PARAMETER LDAPReport
Create a CSV report from LDAP.

.PARAMETER NoReport
Do not generate new reports, use existing data.

.PARAMETER ReportsFolder
Specify a different reports folder. The default folder is $ScriptRoot\Reports.

.PARAMETER ReportOnly
Run reports and quit.

.PARAMETER NoEmail
Suppress notification e-mails.

.PARAMETER TestData
Specify a location where test data reports are stored. Implies -NoReport.

.PARAMETER IgnoreStartupChecks
Do not check for presence of prerequisite software and modules.

#>

Param (
	[switch]$CleanupGroups,
	[ValidateSet("True",
				 "False")][string]$DryRun = "True",
	[switch]$ProcessGroups,
	[switch]$PurgeGroupMembers,
	[switch]$RegistrationGroups,
	[switch]$StaffGroups,
	[switch]$SubjectGroups,
	[switch]$TeachingGroups,
	[switch]$StaffUsers,
	[switch]$StudentUsers,
	[switch]$NoLeavers,
	[switch]$ProcessMSOL,
	[switch]$ProcessExchange,
	[string]$LogFile,
	[switch]$NoLog,
	[switch]$LDAPReport,
	[switch]$NoReport,
	[string]$ReportsFolder,
	[switch]$ReportOnly,
	[switch]$NoEmail,
	[string]$TestData,
	[switch]$IgnoreStartupChecks
)

#Exchange Powershell is a funny bugger, so we have to attempt to connect to it before we do anything else.
If ((Test-Path -Path 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1') -and $ProcessExchange -eq $true) {
	# Connect to Exchange
	If ($DryRun -eq "False") {
		. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
		Connect-ExchangeServer -auto
		# Run an exchange Cmdlet. If ActiveDirectory cmdlets are run before Exchange cmdlets, the Exchange cmdlets will fail.
		$xyz = Get-Mailbox -ResultSize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
		$xyz = $null
	} Else {
		write-host ". 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'"
		write-host "Connect-ExchangeServer -auto"
		# Run an exchange Cmdlet. If ActiveDirectory cmdlets are run before Exchange cmdlets, the Exchange cmdlets will fail.
		write-host "`$xyz = Get-Mailbox -ResultSize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue"
		write-host "`$xyz = $null"
	}
}

# Set $InDev to $true to enable development mode.
# Development mode allows you to override parameters, as well as adjusting how some
# things are processed. For example, Users are created with the
# 'msExchHideFromAddressLists' attribute set to 'TRUE'.
#$InDev = $true

if ($InDev) {
	# Argument overrides for development mode. 
	#$NoEmail = $true
	$NoLog = $true
	
	#$NoReport = $true
	#$ReportsFolder = "${PSScriptRoot}\TestReports"
	#$ReportOnly = $true
	#$LDAPReport = $true
	
	#$UsersOnly = $true
	
	#$GroupsOnly = $true
	#	$Registrationgroups = $true
	#	$StaffGroups = $true
	#	$SubjectGroups =$true
	#	$TeachingGroups = $true

	#$StudentsOnly = $true
	#$StaffOnly = $true
	
	#$ProcessMSOL = $false
	
	$IgnoreStartupChecks = $true
}

. "${PSScriptRoot}\Include.ps1"
. "${PSScriptRoot}\Config.ps1"
. "${PSScriptRoot}\coreFunctions.ps1"
. "${PSScriptRoot}\CredMan.ps1"
. "${PSScriptRoot}\ReportingFunctions.ps1"
. "${PSScriptRoot}\CommonUserFunctions.ps1"
. "${PSScriptRoot}\CommonGroupFunctions.ps1"
. "${PSScriptRoot}\StudentUserFunctions.ps1"
. "${PSScriptRoot}\StudentGroupFunctions.ps1"
#. "${PSScriptRoot}\StaffUserFunctions.ps1"
. "${PSScriptRoot}\StaffGroupFunctions.ps1"

If (!($IgnoreStartupChecks)) {
	#Check if required powershell modules are installed.
	$CheckPassed = $true
	Try { Import-Module ActiveDirectory -ErrorAction SilentlyContinue } Catch {}
	Try { Import-Module DirSync -ErrorAction SilentlyContinue } Catch {}
	If (!(Get-Module -ListAvailable -Name ActiveDirectory)) { $CheckPassed = $false }
	If (!(Get-Module -ListAvailable -Name DirSync) -and !((Get-Module -ListAvailable -Name ADSync) -and (Get-Module -ListAvailable -Name MSOnline)) -and $ProcessMSOL -eq $true) { $CheckPassed = $false }
	If (!(Test-Path -Path 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1') -and $ProcessExchange -eq $true) { $CheckPassed = $false }
			
	If ($CheckPassed -eq $false) {
		$NoLog = $true
		speak "One or more Powershell modules are not installed. Required modules are:"
		speak ""
		speak "    ActiveDirectory"
		speak "    DirSync, or ADSync and MSOnline"
		speak "    Exchange Powershell management tools"
		speak ""
		bork 1
	}	
}

#Open log file to be written
If (!($NoLog)) {$LogStream = [System.IO.StreamWriter] "$LogFile"}

If ($DryRun -eq "True") { 

	$Simulate = $true
	speak "Performing a dry run. Nothing will be changed. Use -DryRun False to perform synchronisation."

} Else { $Simulate = $false }

If ($TestData) {
	$NoReport = $true
	$NoEmail = $true
	speak "Using test data. E-mail notifications will not be sent."
}

If ($NoEmail) { speak "E-mail notifications will not be sent." }

If ($NoReport -and $ReportOnly) {
	speak "NoReport and ReportOnly cannot be used together." -LogOff
	bork 1
}

If (($RegistrationGroups -or $SubjectGroups -or $TeachingGroups -or $StaffGroups) -and !($ProcessGroups)) {
	speak "RegistrationGroups, StaffGroups, SubjectGroups and TeachingGroups can only be used in conjunction with ProcessGroups" -LogOff
	bork 1
}

If (!($StaffUsers -or $StudentUsers -or $ProcessGroups)) {
	speak "Please specify one or more of StaffUsers, StudentUsers or ProcessGroups" -LogOff
	bork 1
}

If (!($RegistrationGroups -or $SubjectGroups -or $TeachingGroups -or $StaffGroups) -and $ProcessGroups) {
	$RegistrationGroups = $true
	$SubjectGroups = $true
	$TeachingGroups = $true
	$StaffGroups = $true
}

If ($CleanupGroups -and $PurgeGroupmembers) {
	$CleanupGroups = $Null
}

RunReports
TransferStudents
#TransferStaff
ProcessOffice365
ProcessExchange
ProcessStudentGroups
ProcessStaffGroups
SendEmail

<# To Do:
	E-mail reports
	
	Work out how we are handling Leavers who have returned.
		Leavers are now re-enabled and moved back to the correct OU.
		Still need to figure out how to handle leavers/returners home folders.
	
	Staff department
	
	More 'reliable' way of determining which staff require IT acccess. I.e. Have a user defined field in SIMS.
	
	Delete unused groups. But only when specified on command line.
	
	Office 365 accounts cannot be seen by on-prem users by default. Need to add something to the script to create O365 users as 'Mail Users' on exchange.
		Something like : Enable-MailUser -Identity $($User.userprincipalname) -Alias $($User.samaccountname) -ExternalEmailAddress address@SouthHunsley.mail.onmicrosoft.com
	
#>

#Finish
bork