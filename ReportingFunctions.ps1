Function RunReports {
	If ($ReportOnly){
		speak "Running reports then quitting..."
		speak ""
	}
	If (!($NoReport)) {
		#USER REPORTS
		If ($StaffUsers) {
			#STAFF USER REPORTS
			RunSIMSReport $SIMSStaffUsersReport $SIMSStaffList
			RunLDAPUserReport $LDAPStaffUsersOU $LDAPStaffList $false $TRUE
			RunLDAPUserReport $LDAPStaffUsersITSOU $LDAPStaffList 
			RunLDAPUserReport $LDAPStaffUsersITTOU $LDAPStaffList 
		}
		If ($StudentUsers){
			#STUDENT USER REPORTS
			RunSIMSReport $SIMSStudentUsersReport $SIMSStudentList
			RunLDAPUserReport $LDAPStudentUsersOU $LDAPStudentList $TRUE $TRUE
		}

		#GROUP REPORTS
		If ($StaffGroups) {
			#STAFF GROUPS REPORT
			RunSIMSReport $SIMSAssociateStaffReport $SIMSAssociateStaffList
			RunSIMSReport $SIMSTeachingStaffReport $SIMSTeachingStaffList
			
			RunLDAPGroupMemberReport $LDAPAssociateStaffGroup $LDAPAssociateStaffList
			RunLDAPGroupMemberReport $LDAPTeachingStaffGroup $LDAPTeachingStaffList
		}
		#STUDENT GROUP REPORTS
		If ($RegistrationGroups) {
			#REGISTRATION GROUPS REPORT
			RunSIMSReport $SIMSRegGroupReport $SIMSRegGroupList
			RunSIMSReport $SIMSStudentRegGroupReport $SIMSStudentRegList
		}
		If ($SubjectGroups -or $TeachingGroups) {
			#SUBJECT GROUPS REPORT
			RunSIMSReport $SIMSSubjectGroupReport $SIMSTeachingSubList
		}
		If ($TeachingGroups) {
			#TEACHING GROUPS REPORT
			RunSIMSReport $SIMSTeachingGroupReport $SIMSStudentTeachList
		}
	} else {
		speak "Reports are not being run by user request."
	}
		
	If ($ReportOnly) {bork}
}
	
Function RunSIMSReport($ReportName, $OutputFile) {
	If (Test-Path -Path ${OutputFile}){
		Remove-Item -Path ${OutputFile} -Force -ErrorAction SilentlyContinue
	}
	speak "Running SIMS report ${Reportname}..." -Time
	&  $CommandReporter /TRUSTED /SERVERNAME:"$DBServer" /DATABASENAME:"$DBName" /REPORT:"$ReportName" /OUTPUT:"$OutputFile"
	if ($? -and (Test-Path -Path ${OutputFile})) {
		speak "${OutputFile} created successfully." -Time
		speak ""
	} else {
		speak "Report failed" -Time
		speak ""
	}
}

Function RunLDAPUserReport($ReportOU, $OutputFile, $Recursive = $false, $FirstReport = $false) {
	If ($LDAPReport) {
		If (($FirstReport -eq $TRUE) -and (Test-Path -Path ${OutputFile})){
			Remove-Item -Path ${OutputFile} -Force -ErrorAction SilentlyContinue
		}
		speak "Running LDAP report on ${ReportOU}..." -Time
		If ($Recursive) {
			Get-ADUser -searchbase "${ReportOU}" -filter * -Properties ${LDAPUserPropertiesArray} | Where-Object {($_.DistinguishedName -notlike '*OU=External*') -or ($_.DistinguishedName -notlike '*OU=Guests*')} |Select ${LDAPUserPropertiesArray} |  Export-csv $OutputFile -Append
		} else {
			Get-ADUser -searchbase "${ReportOU}" -filter * -Properties ${LDAPUserPropertiesArray} -SearchScope OneLevel | select ${LDAPUserPropertiesArray} | Export-csv $OutputFile -Append
		}
		if ($? -and (Test-Path -Path ${OutputFile})) {
			speak "${OutputFile} created successfully." -Time
			speak ""
		} else {
			speak "Report failed" -Time
			speak ""
		}
	}
}

Function RunLDAPGroupMemberReport($ReportGroup, $OutputFile, $OverrideNoReport = $false) {
	If ($NoLDAPReport -eq $true -and $OverrideNoReport -eq $false) {
		If (Test-Path -Path ${OutputFile}) {
			Remove-Item -Path ${OutputFile} -Force -ErrorAction SilentlyContinue
		}
		speak "Running LDAP report on ${ReportGroup}..." -Time
		Get-ADGroupMember -Identity ${ReportGroup} | select distinguishedname, samaccountname | Export-Csv $OutputFile
		if ($? -and (Test-Path -Path ${OutputFile})) {
			speak "${OutputFile} created successfully." -Time
			speak ""
		} else {
			speak "Report failed" -Time
			speak ""
		}
	}
}