#Date and time stuff
$FSTimeFormat = "yyyy\-MM\-dd\THH\-mm\-ss"
$ReadableTimeFormat = "dd\/MM\/yyyy\ HH\:mm\:ss"
$StartDate = Date
$StartTime = Get-Date -Date $StartDate -f $FSTimeFormat
$ThisMonth = Get-Date -f MMM

#Set file and folder locations
If (!($LogFile)) {$LogFile = "${PSScriptRoot}\Logs\${StartTime}.log"}
If ($TestData)	{
	$TestLength = $TestData.Length
	If ($TestData.substring($TestLength -1,1) -eq "\") { 
		$ReportsFolder = $TestData.Substring(0,$TestLength -1)
	} else {
		$ReportsFolder = $TestData
	}
} ElseIf (!($ReportsFolder)) {$ReportsFolder = "${PSScriptRoot}\Reports"}
If (Test-Path -Path "${ENV:ProgramFiles(x86)}") {
	$ProgramFiles = "${ENV:ProgramFiles(x86)}"
} else {
	$ProgramFiles = "${ENV:ProgramFiles}"
}

#Office365
Function Get-O365Creds() {
	#Only set this up if it is required, so done as a function.
	$OfficeTargetURL = "https://SouthHunsley.MicrosoftOnline.com"
	$ReadCreds = Read-Creds -Target "$OfficeTargetURL"
	$SecPassword = ConvertTo-SecureString "$($ReadCreds.CredentialBlob)" -AsPlainText -Force
	$Creds = New-Object System.Management.Automation.PSCredential("$($ReadCreds.UserName)",$SecPassword)
	$ReadCreds = $null
	$SecPassword = $null
	Return $Creds
}

#Collections
$NewStaffAccounts = New-Object System.Collections.ArrayList
$DisabledStaff = New-Object System.Collections.ArrayList
$PrevDisabledStaff = New-Object System.Collections.ArrayList
$DupStaffCodes = New-Object System.Collections.ArrayList
$DuplicateStaff = New-Object System.Collections.ArrayList
$ReEnabledStaff = New-Object System.Collections.ArrayList

$NewStudentAccounts = New-Object System.Collections.ArrayList
$DisabledStudents = New-Object System.Collections.ArrayList
$PrevDisabledStudents = New-Object System.Collections.ArrayList
$ReEnabledStudents = New-Object System.Collections.ArrayList

$NewGroups = New-Object System.Collections.ArrayList
$RemovedGroups = New-Object System.Collections.ArrayList
$NewGroupMemberships = New-Object System.Collections.ArrayList
$RemovedGroupMemberships = New-Object System.Collections.ArrayList

$NewOffice365Accounts = New-Object System.Collections.ArrayList
$NewExchangeMailboxes = New-Object System.Collections.ArrayList
$NewExchangeRemoteUsers = New-Object System.Collections.ArrayList

#Environement Checks
# MandatoryProfile
# StudentHomeRoot
# StaffHomeRoot