Function ProcessStudentGroups() {
	RegistrationGroups
	SubjectGroups
	TeachingGroups
}

Function RegistrationGroups() {
	If ($RegistrationGroups) {
		speak ""
		speak "" -Time
		speak "==================================="
		speak " Processing Registration Groups... "
		speak "==================================="
		speak ""
		
		If ($PurgeGroupMembers) {
			speak "Purging Registration group members..."
			speak ""
			$Groups = EnumerateGroups $LDAPRegGroupOU
			ForEach ($Group In $Groups) {
				speak "Purging $($Group.Name)"
				If (!($Simulate)) { Get-ADGroupMember $Group | ForEach-Object {Remove-ADGroupMember $Group $_ -Confirm:$false } }
				Else { speak "Get-ADGroupMember $Group | ForEach-Object {Remove-ADGroupMember $Group $_ -Confirm:$false }" }
			}
			speak ""; speak ""
			return
		}
		
		#Create Registration groups
		speak "Creating Registration groups..."
		speak ""
		$ReadFile = $SIMSRegGroupList
		If (Test-Path -Path $ReadFile) {
			$ReadStream = [System.IO.StreamReader] "${ReadFile}"
			$Group = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${ReadFile} does not exist." -Time
			bork 1
		}
		While (!($ReadStream.EndOfStream)) {
			$Group = $ReadStream.Readline() -Replace """", ""
			$Group = $Group.Split(",")
			$GroupName = "${LDAPRegGroupPrefix} $($Group[0])"
			$Members = $Group[1]
			If (SearchGroup "$GroupName") {
				speak "The group $GroupName exists."
			} ElseIf ($Members -ge 1) {
				speak "Creating group $GroupName..."
				NewGroup "$LDAPRegGroupOU" "$GroupName"
			}
		}
		$ReadStream.Close()
		speak ""
		
		#Add students to Registration groups
		speak "Adding students to Registration groups..."
		speak ""
		$ReadFile = $SIMSStudentRegList
		If (Test-Path -Path $ReadFile) {
			$ReadStream = [System.IO.StreamReader] "${ReadFile}"
			$Student = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${ReadFile} does not exist." -Time
			bork 1
		}
		While (!($ReadStream.EndOfStream)) {
			$Student = $ReadStream.ReadLine() -Replace """", ""
			$Student = $Student.Split(",")
			$StudentName = $Student[0]
			$StudentAdNo = ConvertAdmissionNumber $Student[1]
			$StudentReg = "${LDAPRegGroupPrefix} $($Student[2])"
			$Account = SearchUser -LooseSearch $StudentAdNo -Quiet
			If ($Account -and $Account -ne "ERROR") {
				If (SearchGroup $StudentReg) {
					If (!(GetGroupMembership "$StudentReg" "$($Account.DistinguishedName)" )) {
						speak "Add $StudentName to $StudentReg"
						AddGroupMembership "$StudentReg" "$($Account.DistinguishedName)" 
					} Else {
						speak "$StudentName already a member of $StudentReg."
					}
				} Else {
					speak "ERROR: The group $StudentReg does not exist"
				}
			} Else {
				speak "ERROR: No account found or duplicate account for $StudentName"
			}
		}
		$ReadStream.Close()
		speak ""		
		
		#Remove students from Registration groups
		If ($CleanupGroups){
			speak "Removing students from Registration groups..."
			speak ""
			$Groups = EnumerateGroups $LDAPRegGroupOU
			ForEach ($Group In $Groups) {
				$Members = EnumerateGroupMembers $Group.Name
				ForEach ($Member In $Members) {
					speak "." -nonewline -logoff
					$Remove = $true
					$ReadFile = $SIMSStudentRegList
					If (Test-Path -Path $ReadFile) {
						$ReadStream = [System.IO.StreamReader] "${ReadFile}"
						$Student = $ReadStream.ReadLine()
						While (!($ReadStream.EndOfStream)) {
							$Student = $ReadStream.ReadLine() -Replace """", ""
							$Student = $Student.Split(",")
							$StudentAdNo = ConvertAdmissionNumber $Student[1]
							If ("$($Member.sAMAccountName)" -Match "$StudentAdNo$" ) {$Remove = $false}
						}
						$ReadStream.Close()
					} else {
						speak "ERROR: ${ReadFile} does not exist." -Time
						bork 1
					}
					If ($Remove) {
						RemoveGroupMembership "$($Group.Name)" "$($Member.sAMAccountName)" "$($Member.Name)"
						
					}
				}
			}
			speak ""; speak ""
		}
		
		#Delete unused Registration groups
	}
}

Function SubjectGroups() {
	If ($SubjectGroups) {
		speak ""
		speak "" -Time
		speak "==================================="
		speak " Processing Subject Groups... "
		speak "==================================="
		speak ""
		
		If ($PurgeGroupMembers) {
			speak "Purging Subject group members..."
			speak ""
			$Groups = EnumerateGroups $LDAPSubjectGroupOU
			ForEach ($Group In $Groups) {
				speak "Purging $($Group.Name)"
				If (!($Simulate)) { Get-ADGroupMember $Group | ForEach-Object {Remove-ADGroupMember $Group $_ -Confirm:$false } }
				Else { speak "Get-ADGroupMember $Group | ForEach-Object {Remove-ADGroupMember $Group $_ -Confirm:$false }" }
			}
			speak ""; speak ""
			return
		}
		
		#Create Subject groups
		speak "Creating Subject groups..."
		speak ""
		$ReadFile = $SIMSTeachingSubList
		If (Test-Path -Path $ReadFile) {
			$ReadStream = [System.IO.StreamReader] "${ReadFile}"
			$ClassLine = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${ReadFile} does not exist." -Time
			bork 1
		}
		$AllSubjects = @()
		While (!($ReadStream.EndOfStream)) {
			$ClassLine = $ReadStream.Readline() -Replace """", ""
			$ClassLine = $ClassLine.Split(",").Trim()
			$Keystage = ClassToKeyStage $ClassLine[0]
			If ($Keystage) { $Subject = "$($ClassLine[1]) KS${Keystage}" }
				Else { $Subject = $ClassLine[1] }
			$AllSubjects += $Subject
		}
		
		$ReadStream.Close()
		$AllSubJects = $AllSubjects | Sort-Object -Unique
		ForEach ($Subject In $AllSubjects) {
			$GroupName = "${LDAPSubjectGroupPrefix} ${Subject}"
			If (SearchGroup $GroupName) { speak "The group ${GroupName} exists" }
				Else {
				speak "Creating Group ${GroupName}"
				NewGroup "$LDAPSubjectGroupOU" "$GroupName"
			}
		}
		speak ""
		
		#Remove Teaching groups from subject groups
		If ($CleanupGroups){
			speak "Removing Teaching groups from Subject groups..." -Time
			speak ""
			$Groups = EnumerateGroups $LDAPSubjectGroupOU
			ForEach ($Group In $Groups) {
				$Members = EnumerateGroupMembers $Group.Name
				ForEach ($Member In $Members) {
					speak "." -nonewline -logoff
					$Remove = $true
					$ReadFile = $SIMSTeachingSubList
					If (Test-Path -Path $ReadFile) {
						$ReadStream = [System.IO.StreamReader] "${ReadFile}"
						$Class = $ReadStream.ReadLine()
						While (!($ReadStream.EndOfStream)) {
							$Class = $ReadStream.ReadLine() -Replace """", ""
							$Class = $Class.Split(",")
							$GroupName = "${LDAPTeachingGroupPrefix} $($Class[0])" -Replace("/","-")
							If ("$($Member.sAMAccountName)" -Eq "$GroupName" ) {$Remove = $false}
						}
						$ReadStream.Close()
					} else {
						speak "ERROR: ${ReadFile} does not exist." -Time
						bork 1
					}
					If ($Remove) {
						RemoveGroupMembership "$($Group.Name)" "$($Member.sAMAccountName)" "$($Member.Name)"
					}
				}
			}
			speak ""; speak ""
		}
		
		#Delete unused Subject groups
	}
}

Function TeachingGroups() {
	If ($TeachingGroups) {
		$Regex = '(Mon|Tue|Wed|Thu|Fri)(:)?([1235]|4a|4b)' #Excludes special non-timetabled classes such as "Support", "College", "Football"
		speak ""
		speak "" -Time
		speak "==================================="
		speak " Processing Teaching Groups... "
		speak "==================================="
		speak ""
		
		If ($PurgeGroupMembers) {
			speak "Purging Teaching group members..."
			speak ""
			$Groups = EnumerateGroups $LDAPTeachingGroupOU
			ForEach ($Group In $Groups) {
				speak "Purging $($Group.Name)"
				If (!($Simulate)) { Get-ADGroupMember $Group | ForEach-Object {Remove-ADGroupMember $Group $_ -Confirm:$false } }
				Else { speak "Get-ADGroupMember $Group | ForEach-Object {Remove-ADGroupMember $Group $_ -Confirm:$false }"}
			}
			speak ""; speak ""
			return
		}
		
		#Create Teaching groups
		speak "Creating Teaching groups..."
		speak ""
		$ReadFile = $SIMSTeachingSubList
		If (Test-Path -Path $ReadFile) {
			$ReadStream = [System.IO.StreamReader] "${ReadFile}"
			$Group = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${ReadFile} does not exist." -Time
			bork 1
		}
		While (!($ReadStream.EndOfStream)) {
			$Group = $ReadStream.Readline() -Replace """", ""
			$Group = $Group.Split(",").Trim()
			$GroupName = "${LDAPTeachingGroupPrefix} $($Group[0])" -Replace("/","-")
			$Members = $Group[1]
			If (SearchGroup "$GroupName") {
				speak "The group $GroupName exists."
			} ElseIf ($Members -ge 1) {
				If (!("$GroupName" -Match $Regex)){
					speak "Creating group $GroupName..."
					NewGroup "$LDAPTeachingGroupOU" "$GroupName"
				}
			}
		}
		$ReadStream.Close()
		speak ""
		
		#Add Teaching groups to subject groups
		speak "Add Teaching groups to Subject groups..." -Time
		speak ""
		$ReadFile = $SIMSTeachingSubList
		If (Test-Path -Path $ReadFile) {
			$ReadStream = [System.IO.StreamReader] "${ReadFile}"
			$Group = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${ReadFile} does not exist." -Time
			bork 1
		}
		While (!($ReadStream.EndOfStream)) {
			$ClassLine = $ReadStream.Readline() -Replace """", ""
			$ClassLine = $ClassLine.Split(",").Trim()
			$GroupName = "${LDAPTeachingGroupPrefix} $($ClassLine[0])" -Replace("/","-")
			$Keystage = ClassToKeyStage $ClassLine[0]
			If ($Keystage) { $Subject = "$($ClassLine[1]) KS${Keystage}" }
				Else { $Subject = $ClassLine[1] }
			If ($Group = SearchGroup "$GroupName") {
				#$Group.DistinguishedName
				$SubjectGroup = "$LDAPSubjectGroupPrefix $Subject"
				#$Subjectgroup
				If (SearchGroup "$SubjectGroup") {
					If (!(GetGroupMembership "$SubjectGroup" "$($Group.DistinguishedName)")){
						speak "Add $($Group.Name) to $SubjectGroup"
						AddGroupMembership "$SubjectGroup" "$($Group.DistinguishedName)" -ObjectType "group"
					} Else {
						speak "$($Group.Name) already a member of $SubjectGroup."
					}
				}
			}
		}
		$ReadStream.Close()
		speak ""
		
		#Add students to Teaching groups
		speak "Add students to teaching groups" -Time
		$ReadFile = $SIMSStudentTeachList
		If (Test-Path -Path $ReadFile) {
			$ReadStream = [System.IO.StreamReader] "${ReadFile}"
			$Group = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${ReadFile} does not exist." -Time
			bork 1
		}
		While (!($ReadStream.EndOfStream)) {
			$StudentClass = $ReadStream.ReadLine() -Replace """",""
			$StudentClass = $StudentClass.Split(",").Trim()
			$StudentName = $StudentClass[0]
			$StudentAdNo = ConvertAdmissionNumber $StudentClass[1]
			$ClassName = "${LDAPTeachingGroupPrefix} $($StudentClass[2])" -Replace("/","-")
			$Account = SearchUser -LooseSearch $StudentAdNo -Quiet
			If ($Account) {
				If (SearchGroup $ClassName) {
					If (!(GetGroupMembership "$ClassName" "$($Account.DistinguishedName)" )) {
						speak "Add $StudentName to $ClassName"
						AddGroupMembership "$Classname" "$($Account.DistinguishedName)" 
					} Else {
						speak "$StudentName already a member of $ClassName."
					}
				} Else {
					speak "ERROR: The group $Classname does not exist"
				}
			} Else {
				speak "ERROR: No account found for $StudentName"
			}
		}
		$ReadStream.Close()
		speak ""
		
		#Remove students from Teaching groups
		If ($CleanupGroups){
			speak "Removing students from Teaching groups..." -Time
			speak ""
			$Groups = EnumerateGroups $LDAPTeachingGroupOU
			ForEach ($Group In $Groups) {
				speak "." -NoNewLine -LogOff
				$Members = EnumerateGroupMembers $Group.Name
				ForEach ($Member In $Members) {
					speak "." -nonewline -logoff
					$Remove = $true
					$ReadFile = $SIMSStudentTeachList
					If (Test-Path -Path $ReadFile) {
						$ReadStream = [System.IO.StreamReader] "${ReadFile}"
						$Student = $ReadStream.ReadLine()
						While (!($ReadStream.EndOfStream)) {
							$Student = $ReadStream.ReadLine() -Replace """", ""
							$Student = $Student.Split(",")
							$StudentAdNo = ConvertAdmissionNumber $Student[1]
							If ("$($Member.sAMAccountName)" -Match "$StudentAdNo$" ) {$Remove = $false}
						}
						$ReadStream.Close()
					} else {
						speak "ERROR: ${ReadFile} does not exist." -Time
						bork 1
					}
					If ($Remove) {
						speak "" -LogOff
						RemoveGroupMembership "$($Group.Name)" "$($Member.sAMAccountName)" "$($Member.Name)"
					}
				}
			}
			speak ""; speak ""
		}
		
		#Delete unused Teaching groups
	}
}

Function ClassToKeyStage ($Class) {
	$Regex ='[^0-9]'
	$Result = $Class.Substring(0,2) -Replace $Regex, ""
	$Result = GetKeyStage $Result
	Return $Result
}