Function Speak() {
	Param (
		[Parameter(Position=1)][string]$Message,
		[switch]$LogOff,
		[switch]$Time,
		[switch]$NoNewLine
	)
	
	If ($NoLog) { $LogOff = $TRUE }
	
	$TheTime = Now
	If ($TIME) {
		$Message = "${TheTime} ${Message}"
	}
	If ($NoNewLine) {
		write-host -NoNewLine $Message
	} Else {
		write-host $Message
	}
	
	If (!($LogOff)) {$LogStream.WriteLine("${Message}")}
}

Function Bork ([INT]$error = 0) {
	# Close opened text streams
	If ($ReadStream) {speak "Closing ${ReadStream}"; $ReadStream.Close(); $ReadStream = $NULL}
	speak "Finished." -Time
	If ($LogStream) {$LogStream.close(); $LogStream = $NULL}
	exit $error
}

Function Now {
	$Out = Get-Date -Date $(Date) -f $ReadableTimeFormat
	Return $Out
}

Function Get-WeekDayInMonth ([int]$Month, [int]$year, [int]$WeekNumber, [int]$WeekDay) {
	#From http://blog.tyang.org/2012/09/03/powershell-function-get-weekdayinmonth/
	$FirstDayOfMonth = Get-Date -Year $year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0
	#First week day of the month (i.e. first monday of the month)
	[int]$FirstDayofMonthDay = $FirstDayOfMonth.DayOfWeek
	$Difference = $WeekDay - $FirstDayofMonthDay
	If ($Difference -lt 0) {
	 $DaysToAdd = 7 - ($FirstDayofMonthDay - $WeekDay)
	} elseif ($difference -eq 0 )	{
		$DaysToAdd = 0
	} else {
		$DaysToAdd = $Difference
	}
	$FirstWeekDayofMonth = $FirstDayOfMonth.AddDays($DaysToAdd)
	Remove-Variable DaysToAdd
	#Add Weeks
	$DaysToAdd = ($WeekNumber -1)*7
	$TheDay = $FirstWeekDayofMonth.AddDays($DaysToAdd)
	If (!($TheDay.Month -eq $Month -and $TheDay.Year -eq $Year)) {
		$TheDay = $null
	}
	$TheDay
}

Function SendEmail {
	If (!($NoEmail)) {
		Add-Type -AssemblyName System.Web		
		# Create User report if any changes have bene made to user accounts
		If ($NewStaffAccounts -or $DisabledStaff -or $PrevDisabledStaff -or $DuplicateStaff -or $NewStudentAccounts -or $DisabledStudents) {
			$CreateUserReport = $true
		}
		
		# Create a user report if there have are any disabled users that still need to be cleaned up.
		# This will only apply on the first Monday of the month.
		If ($PrevDisabledStaff -or $PrevDisabledStudents) {
			$Today = Get-Date -f "dd MMMM yyyy"
			$FirstMonday = Get-Date -Date (Get-WeekdayInMonth (Get-Date -f "MM") ($Year = Get-Date -f "yyyy") 2 5) -f "dd MMMM yyyy"
			If ($Today -eq $FirstMonday) { $CreateUserReport = $true } 
			Else { 
				# Nullify the Previous disabled arrays to stop these being reported on when not required.
				$PrevDisabledStaff = $null
				$PrevDisabledStudents = $null
			}
		}
		
		# Create a group report if any changes have been made to groups
		If ($NewGroups -or $RemovedGroups -or $NewGroupMemberships -or $RemovedGroupMemberships) {
			$CreateGroupReport = $true
		}
		
		## E-mail administrators ##
		If ($CreateUserReport -eq $true -or $CreateGroupReport -eq $true) {
			$EmailAdministrators = @{}
			If (!($Simulate)) {
				$EmailAdministrators.Subject = "Administrator's user account synchronisation report"
			} Else {
				$EmailAdministrators.Subject = "(Dry Run) Administrator's user account synchronisation report"
			}
			
			$CountNewStaff = @($NewStaffAccounts).Count
			$CountNewStudents = @($NewStudentAccounts).Count
			$CountNewUsers = $($CountNewStaff + $CountNewStudents)
			$CountDisabledStaff = @($DisabledStaff).Count
			$CountDisabledStudents = @($DisabledStudents).Count
			$CountDisabledUsers = $($CountDisabledStaff + $CountDisabledStudents)
			$CountDuplicateStaff = @($DupStaffCodes).Count
			$CountNewGroups = @($NewGroups).Count
			$CountNewGroupMembers = @($NewGroupMemberships).Count
			$CountRemovedGroups = @($RemovedGroups).Count
			$CountRemovedGroupMembers = @($RemovedGroupMemberships).Count
			
			If (!($Simulate)) {
				$MessageBody = "<head><Title>User account synchronisation report</title></head><body style=""font-family: Arial,Helvetica Neue,Helvetica,sans-serif;"">"
			} Else {
				$MessageBody = "<head><Title>User account synchronisation report (Dry Run)</title></head><body style=""font-family: Arial,Helvetica Neue,Helvetica,sans-serif;"">"`
					+ "<p>This report is generated from a dry run. None of the changes listed below have been made.</p>"
			}
			$MessageBody += "<h1 style=""text-decoration: underline;"">Report Summary</h1>"`
				+ "<table><tr><td>New Users</td><td>$CountNewUsers</td></tr>"`
				+ "<tr><td>New Staff</td><td>$CountNewStaff</td></tr>"`
				+ "<tr><td>New Students</td><td>$CountNewStudents</td></tr>"`
				+ "<tr><td>Disabled Users</td><td>$CountDisabledUsers</td></tr>"`
				+ "<tr><td>Disabled Staff</td><td>$CountDisabledStaff</td></tr>"`
				+ "<tr><td>Disabled Students</td><td>$CountDisabledStudents</td></tr>"`
				+ "<tr><td>Duplicate Staff</td><td>$CountDuplicateStaff</td></tr>"`
				+ "<tr><td>New Groups</td><td>$CountNewGroups</td></tr>"`
				+ "<tr><td>Removed Groups</td><td>$CountRemovedGroups</td></tr>"`
				+ "<tr><td>New Group Members</td><td>$CountNewGroupMembers</td></tr>"`
				+ "<tr><td>Removed Group Members</td><td>$CountRemovedGroupMembers</td></tr>"`
				+ "</table>"
			$EmailAdministrators.Body = $MessageBody
		}
		
		If ($CreateUserReport -eq $true) {
			$EmailAdministrators.Body += "<h1 style=""text-decoration: underline;"">User account synchronisation report</h1>"
		}
		
		#E-mail personnel
		If ($NewStaffAccounts -or $DisabledStaff -or $DuplicateStaff){
			$EmailPersonnel = @{}
			$EmailPersonnel.Subject = "Personnel user account synchronisation report"
			$EmailPersonnel.Body = "<head><Title>User account synchronisation report</title></head><body style=""font-family: Arial,Helvetica Neue,Helvetica,sans-serif;"">"`
				+ "<h1 style=""text-decoration: underline;"">User account synchronisation report</h1>"
		}
		
		If ($NewStaffAccounts) {
			$MessageBody = "<h2>New staff</h2>"`
				+ "<p>New staff accounts have been created. The details are listed below.<p>"`
				+ "<table><tr><th>Name</th><th>Username</th><th>Password</th><th>E-mail</th></tr>"
			ForEach ($Staff In $NewStaffAccounts) {
				$XPassword = [System.Web.HttpUtility]::HtmlEncode($Staff.Password)
				$MessageBody += "<tr><td>$($Staff.Name)</td><td>$($Staff.UserName)</td><td><span style=""font-family: Courier New,Courier,Lucida Sans Typewriter,Lucida Typewriter,monospace;"">$XPassword</span></td><td>$($Staff.EmailAddress)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
			$EmailPersonnel.Body += $MessageBody
		}

		If ($DisabledStaff) {
			$MessageBody = "<h2>Disabled Staff</h2>"`
				+ "<p>The following staff accounts have been disabled.</p>"`
				+ "<table><tr><th>Name</th><th>Username</th></tr>"
			ForEach ($Staff In $DisabledStaff) {
				$MessageBody += "<tr><td>$($Staff.cn)</td><td>$($Staff.sAMAccountName)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
			$EmailPersonnel.Body += $MessageBody
		}

		If ($PrevDisabledStaff) {
			$MessageBody = "<h2>Previously Disabled Staff</h2>"`
				+ "<p>Below is a list of previously disabled accounts that have not been processed.</p>"`
				+ "<table><tr><th>Name</th><th>Username</th></tr>"
			ForEach ($Staff In $PrevDisabledStaff) {
				$MessageBody += "<tr><td>$($Staff.Name)</td><td>$($Staff.sAMAccountName)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
		}
		
		If ($DuplicateStaff) {
			$MessageBody = "<h2>Duplicate Staff</h2>"`
				+ "<p>Possible duplicate staff codes have been found.</p>"`
				+ "<p>If two or more users appear in the SIMS database, there are multiple staff with the same staff code in SIMS. If one of these also has an entry in the LDAP database, this user must not be the one to be edited in SIMS.</p>"`
				+ "<p>If there is only one user in both SIMS and LDAP and the IDs do not match, contact IT support to help resolve the problem.</p>"`
				+ "<p>The details are listed below.</p>"
			ForEach ($Code In $DupStaffCodes) {
				$SIMSRecords = $DuplicateStaff | Where-Object {$_.StaffCode -eq "$Code" -and $_.DB -eq "SIMS"}
				$LDAPRecords = $DuplicateStaff | Where-Object {$_.StaffCode -eq "$Code" -and $_.DB -eq "LDAP"}
				$MessageBody += "<h3>$Code</h3>"
				$MessageBody += "<table><tr><th>Database</th><th>ID</th><th>Name</th></tr>"
				ForEach ($SIMSRecord In $SIMSRecords) { 
					$MessageBody += "<tr><td>SIMS</td><td>$($SIMSRecord.ID)</td><td>$($SIMSRecord.Name)</td></tr>"
				}
				ForEach ($LDAPRecord In $LDAPRecords) { 
					$MessageBody += "<tr><td>LDAP</td><td>$($LDAPRecord.ID)</td><td>$($LDAPRecord.Name)</td></tr>"
				}
				$MessageBody += "</table>"
			}
			$EmailAdministrators.Body += $MessageBody
			$EmailPersonnel.Body += $MessageBody
		}
		
		#E-mail admissions 
		If ($NewStudentAccounts -or $DisabledStudents) {
			$EmailAdmissions = @{}
			$EmailAdmissions.Subject = "Admissions user account synchronisation report"
			$EmailAdmissions.Body = "<head><Title>User account synchronisation report</title></head><body style=""font-family: Arial,Helvetica Neue,Helvetica,sans-serif;"">"`
				+ "<h1 style=""text-decoration: underline;"">User account synchronisation report</h1>"
		}
		
		If ($NewStudentAccounts) {
			$MessageBody = "<h2>New Students</h2>"`
				+ "<p>New Student accounts have been created. The details are listed below.<p>"`
				+ "<table><tr><th>Name</th><th>Username</th><th>Password</th><th>E-mail</th></tr>"
			ForEach ($Student In $NewStudentAccounts) {
				$XPassword = [System.Web.HttpUtility]::HtmlEncode($Student.Password)
				$MessageBody += "<tr><td>$($Student.Name)</td><td>$($Student.UserName)</td><td><span style=""font-family: Courier New,Courier,Lucida Sans Typewriter,Lucida Typewriter,monospace;"">$XPassword</span></td><td>$($Student.EmailAddress)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
			$EmailAdmissions.Body += $MessageBody
		}

		If ($DisabledStudents) {
			$MessageBody = "<h2>Disabled Students</h2>"`
				+ "<p>The following Student accounts have been disabled.</p>"`
				+ "<table><tr><th>Name</th><th>Username</th></tr>"
			ForEach ($Student In $DisabledStudents) {
				$MessageBody += "<tr><td>$($Student.cn)</td><td>$($Student.sAMAccountName)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
			$EmailAdmissions.Body += $MessageBody
		}
		
		If ($PrevDisabledStudents) {
			$MessageBody = "<h2>Previously Disabled Students</h2>"`
				+ "<p>Below is a list of previously disabled accounts that have not been processed.</p>"`
				+ "<table><tr><th>Name</th><th>Username</th></tr>"
			ForEach ($Staff In $PrevDisabledStaff) {
				$MessageBody += "<tr><td>$($Student.cn)</td><td>$($Student.sAMAccountName)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
		}
		
		If ($CreateGroupReport -eq $true) {
			$EmailAdministrators.Body += "<h1 style=""text-decoration: underline;"">Group synchronisation report</h1>"
		}
		
		If ($NewGroups) {
			$MessageBody = "<h2>New Groups</h2>"`
				+ "<p>New groups have been created.</p>"`
				+ "<table><tr><th>Group Name</th><th>Path</th></tr>"
			ForEach ($Group In $NewGroups) {
				$MessageBody += "<tr><td>$($Group.Name)</td><td>$($Group.Path)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
		}
		
		If ($RemovedGroups) {
			$MessageBody = "<h2>Removed Groups</h2>"`
				+ "<p>Groups have been removed.</p>"`
				+ "<table><tr><th>Group Name</th></tr>"
			ForEach ($Group In $RemovedGroups) {
				$MessageBody += "<tr><td>$($Group.Name)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
		}
		
		If ($NewGroupMemberships) {
			$MessageBody = "<h2>New Group Members</h2>"`
				+ "<p>New group memebrships have been created.</p>"`
				+ "<table><tr><th>Group Name</th><th>New Member</th></tr>"
			ForEach ($Membership In $NewGroupMemberships) {
				$MessageBody += "<tr><td>$($Membership.Group)</td><td>$($Membership.Member)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
		}
		
		If ($RemovedGroupMemberships) {
			$MessageBody = "<h2>Removed Group Members</h2>"`
				+ "<p>Group memberships have been removed.</p>"`
				+ "<table><tr><th>Group Name</th><th>Removed Member</th></tr>"
			ForEach ($Membership In $RemovedGroupMemberships) {
				$MessageBody += "<tr><td>$($Membership.Group)</td><td>$($Membership.Member)</td></tr>"
			}
			$MessageBody += "</table>"
			$EmailAdministrators.Body += $MessageBody
		}
		
		If ($EmailAdministrators) {
			$EmailAdministrators.Body += "</body>"
			Send-MailMessage -To $AdministratorContacts -From $FromAddress -Subject $EmailAdministrators.Subject -BodyAsHTML $EmailAdministrators.Body
		}
		
		# If ($EmailPersonnel -and $Simulate -eq $false ) {
			# $EmailPersonnel.Body += "</body>"
			# Send-MailMessage -To $PersonnelContacts -From $FromAddress -Subject $EmailPersonnel.Subject -BodyAsHTML $EmailPersonnel.Body
		# }
		
		If ($EmailAdmissions -and $Simulate -eq $false) {
			$EmailAdmissions.Body += "</body>"
			Send-MailMessage -To $AdmissionsContacts -From $FromAddress -Subject $EmailAdmissions.Subject -BodyAsHTML $EmailAdmissions.Body
		}
	}
}