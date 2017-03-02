Function ProcessStaffGroups() {
	If ($StaffGroups) {
		StaffGroups
	}
}

Function StaffGroups() {
	speak ""
	speak "" -Time
	speak "==========================="
	speak " Processing Staff Groups... "
	speak "==========================="
	speak ""
	
	#Add staff to groups
	speak "Adding staff to Associate staff group..."
	speak ""
	$ReadFile = $SIMSAssociateStaffList
	If (Test-Path -Path $ReadFile) {
		$ReadStream = [System.IO.StreamReader] "${ReadFile}"
		$Staff = $ReadStream.ReadLine()
	} else {
		speak "ERROR: ${ReadFile} does not exist." -Time
		bork 1
	}
	While (!($ReadStream.EndOfStream)) {
		$Staff = $ReadStream.ReadLine()
		$Account = SearchUser -POBox $Staff -Quiet
		If ($Account -and $Account -ne "ERROR") {
			If (SearchGroup $LDAPAssociateStaffGroup) {
				If (!(GetGroupMembership "$LDAPAssociateStaffGroup" "$($Account.DistinguishedName)" )) {
					speak "Add $($Account.cn) to $LDAPAssociateStaffGroup"
					AddGroupMembership "$LDAPAssociateStaffGroup" "$($Account.DistinguishedName)" 
				} Else {
					speak "$($Account.cn) already a member of $LDAPAssociateStaffGroup."
				}
			} Else {
				speak "ERROR: The group $LDAPAssociateStaffGroup does not exist"
			}
		} Else {
			speak "ERROR: No account found or duplicate account for $Staff"
		}
	}
	$ReadStream.Close()
	speak ""
	speak "Adding staff to Teaching staff group..."
	speak ""
	$ReadFile = $SIMSTeachingStaffList
	If (Test-Path -Path $ReadFile) {
		$ReadStream = [System.IO.StreamReader] "${ReadFile}"
		$Staff = $ReadStream.ReadLine()
	} else {
		speak "ERROR: ${ReadFile} does not exist." -Time
		bork 1
	}
	While (!($ReadStream.EndOfStream)) {
		$Staff = $ReadStream.ReadLine()
		$Account = SearchUser -POBox $Staff -Quiet
		If ($Account -and $Account -ne "ERROR") {
			If (SearchGroup $LDAPTeachingStaffGroup) {
				If (!(GetGroupMembership "$LDAPTeachingStaffGroup" "$($Account.DistinguishedName)" )) {
					speak "Add $($Account.cn) to $LDAPTeachingStaffGroup"
					AddGroupMembership "$LDAPTeachingStaffGroup" "$($Account.DistinguishedName)" 
				} Else {
					speak "$($Account.cn) already a member of $LDAPTeachingStaffGroup."
				}
			} Else {
				speak "ERROR: The group $LDAPTeachingStaffGroup does not exist"
			}
		} Else {
			speak "ERROR: No account found or duplicate account for $Staff"
		}
	}
	$ReadStream.Close()
	speak ""
	
	#Remove staff from groups
	If ($CleanupGroups){
		speak "Removing staff from Associate Staff groups..."
		speak ""
		$Members = EnumerateGroupMembers $LDAPAssociateStaffGroup
		ForEach ($Member In $Members) {
			speak "." -nonewline -logoff
			$Member = Get-ADUser $Member -Properties sAMAccountName,Name,postOfficeBox | select sAMAccountName,Name,postOfficeBox
			$PostOfficeBox = $Member.postOfficeBox
			$Remove = $true
			$ReadFile = $SIMSAssociateStaffList
			If (Test-Path -Path $ReadFile) {
				$ReadStream = [System.IO.StreamReader] "${ReadFile}"
				$Staff = $ReadStream.ReadLine()
				While (!($ReadStream.EndOfStream)) {
					$Staff = $ReadStream.ReadLine()
					If ("$PostOfficeBox" -eq "$Staff" ) {$Remove = $false}
				}
				$ReadStream.Close()
			} else {
				speak "ERROR: ${ReadFile} does not exist." -Time
				bork 1
			}
			If ($Remove) {
				RemoveGroupMembership "$LDAPAssociateStaffGroup" "$($Member.sAMAccountName)" "$($Member.Name)"
				
			}
		}
		speak ""; speak ""
		speak "Removing staff from Teaching Staff groups..."
		speak ""
		$Members = EnumerateGroupMembers $LDAPTeachingStaffGroup
		ForEach ($Member In $Members) {
			speak "." -nonewline -logoff
			$Member = Get-ADUser $Member -Properties sAMAccountName,Name,postOfficeBox | select sAMAccountName,Name,postOfficeBox
			$PostOfficeBox = $Member.postOfficeBox
			$Remove = $true
			$ReadFile = $SIMSTeachingStaffList
			If (Test-Path -Path $ReadFile) {
				$ReadStream = [System.IO.StreamReader] "${ReadFile}"
				$Staff = $ReadStream.ReadLine()
				While (!($ReadStream.EndOfStream)) {
					$Staff = $ReadStream.ReadLine()
					If ("$PostOfficeBox" -eq "$Staff" ) {$Remove = $false}
				}
				$ReadStream.Close()
			} else {
				speak "ERROR: ${ReadFile} does not exist." -Time
				bork 1
			}
			If ($Remove) {
				RemoveGroupMembership "$LDAPTeachingStaffGroup" "$($Member.sAMAccountName)" "$($Member.Name)"
				
			}
		}
	}
}