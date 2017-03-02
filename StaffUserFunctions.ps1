Function TransferStaff {
	If ($StaffUsers) {
		$Ti = 0; $FN = 1; $MN = 2; $SN = 3; $SC = 4; $ID = 5; $RT = 6
		$NewStaff = New-Object System.Collections.ArrayList
		speak ""
		speak "" -Time
		speak "============================="
		speak " Processing staff members... "
		speak "============================="
		speak ""
		If (Test-Path -Path $SIMSStaffList) {
			$ReadStream = [System.IO.StreamReader] "${SIMSStaffList}"
			$User = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${SIMSStaffList} does not exist." -Time
			bork 1
		}
		While (!($ReadStream.EndOfStream)) {
			$User = $ReadStream.ReadLine() -Replace """", ""
			$User = $User.Split(",")
# If ($User[$sn] -eq "Almond") {
			If ($User[$SC] -eq "" -or !($User[$SC])) {
				speak "$($User[$FN]) $($User[$SN]) has no staff code" -Time
			} else {
				$UserName = "staff$($User[$SC])".ToLower() -Replace " ", ""
				# If the user does not have a middle name, force it to be empty
				If ($User[$MN] -eq "" -or $User[$MN] -eq " "){$User[$MN] = $NULL}
				$UserAccount = SearchUser -Search $UserName
				 If (!($UserAccount)) {
					$NewUser = @{}
					$NewUser.Title = $User[$Ti]
					$NewUser.FirstName = $User[$FN]
					$NewUser.MiddleNames = $User[$MN]
					$NewUser.Surname = $User[$SN]
					$NewUser.UserName = $UserName
					$NewUser.Password = GetTempPassword
					$NewUser.ID = $User[$ID]
					$NewStaff.Add($NewUser) > $null
					speak "Will create new user: $($NewUser.FirstName) $($NewUser.Surname)."
					speak ""
					$NewUser = $null
				} ElseIf ($UserAccount -ne "ERROR") {
					If ($UserAccount.PostOfficeBox -eq $User[$ID] ) {
						speak "User $Username exists"
						UpdateUser -Account $UserAccount -Title $User[$Ti] -GivenName $User[$FN] -MiddleNames $User[$MN] -Surname $User[$SN] -Company "Staff" -POBox $User[$ID] -UserName "$UserName"
						speak ""
					} ElseIf ($UserAccount.PostOfficeBox -ne $User[$ID] ) {
						speak "WARNING: User $Username exists and is not this user. Possible duplicate staff IDs."
						$DupStaffCodes.Add($UserName) > $null
						speak ""
					} Else {
						speak "WARNING: Unable to properly identify existing User $UserName. Possible duplicate staff IDs."
						$DupStaffCodes.Add($UserName) > $null
						speak ""
					}	
				}
			}
# } # End if statement to capture a single user
			$User = $null
		}
		$ReadStream.Close()
		If ($NewStaff) {
			CreateStaffAccounts $NewStaff
			CreateStaffHomeFolders $NewStaff
		}
		If ($DupStaffCodes) { DuplicateStaff $DupStaffCodes }
		DisableStaffLeavers
	}
}

Function DuplicateStaff($StaffCodes) {
	$Ti = 0; $FN = 1; $MN = 2; $SN = 3; $SC = 4; $ID = 5; $RT = 6 
	speak "There are $($StaffCodes.count) possible duplicate staff codes."
	speak ""
	$StaffCodes = $StaffCodes | Sort-Object -Unique
	ForEach ($StaffCode In $StaffCodes) {
		speak "$StaffCode :"
		If (Test-Path -Path $SIMSStaffList) {
			$ReadStream = [System.IO.StreamReader] "${SIMSStaffList}"
			$User = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${SIMSStaffList} does not exist." -Time
			bork 1
		}
		speak "Accounts existing in SIMS:"
		While (!($ReadStream.EndOfStream)) {
			$User = $ReadStream.ReadLine() -Replace """", ""
			$User = $User.Split(",")
			If ($User[$SC] -eq "" -or !($User[$SC])) {
				$UserName= $null
			} else {
				$UserName = "staff$($User[$SC])".ToLower().Trim()
			}
			If ( $UserName -eq $StaffCode ) {
				speak "	$($User[$ID]): $($User[$FN]) $($User[$SN])"
				$DupStaffAccountInfo = New-Object PSObject
				$DupStaffAccountInfo | Add-Member -MemberType NoteProperty -Name ID -Value "$($User[$ID])"
				$DupStaffAccountInfo | Add-Member -MemberType NoteProperty -Name Name -Value "$($User[$FN]) $($User[$SN])"
				$DupStaffAccountInfo | Add-Member -MemberType NoteProperty -Name StaffCode -Value "$StaffCode"
				$DupStaffAccountInfo | Add-Member -MemberType NoteProperty -Name DB -Value "SIMS"
				$DuplicateStaff.Add($DupStaffAccountInfo) > $null
			}		
		}
		speak "Accounts existing in LDAP:"
		$UserAccount = SearchUser -Search $StaffCode -Quiet
		If ($UserAccount) {
			speak "	$($UserAccount.PostOfficeBox): $($UserAccount.GivenName) $($UserAccount.sn)"
			$DupStaffAccountInfo = New-Object PSObject
				$DupStaffAccountInfo | Add-Member -MemberType NoteProperty -Name ID -Value "$($UserAccount.PostOfficeBox)"
				$DupStaffAccountInfo | Add-Member -MemberType NoteProperty -Name Name -Value "$($UserAccount.GivenName) $($UserAccount.sn)"
				$DupStaffAccountInfo | Add-Member -MemberType NoteProperty -Name StaffCode -Value "$StaffCode"
				$DupStaffAccountInfo | Add-Member -MemberType NoteProperty -Name DB -Value "LDAP"
				$DuplicateStaff.Add($DupStaffAccountInfo) > $null
		}		
		speak ""
	}
}

Function GetTempPassword() {
	Param(
		[int]$Length = 10
	)
	$ascii=$NULL;
	For ($a = 48; $a -le 122; $a++) { 
		$ascii+=,[char][byte]$a
	}
	For ($Loop = 1; $Loop -le $Length; $Loop++) {
		$Character = ($Ascii | Get-Random)
		If (!($Character -eq '`')){
			$TempPassword += $Character
		} Else {
			$Loop--
		}
		
	}
	Return $TempPassword
}

Function CreateStaffHomeFolders($Accounts) {
	speak ""
	speak "" -Time
	speak "===================================="
	speak " Setting up staff home folders... "
	speak "===================================="
	speak ""
	ForEach ($Account In $Accounts) {
		$HomeRoot = "${StaffHomeRoot}"
		$NewFolderName = "$($Account.UserName)"
		$NewFolder = "${HomeRoot}\${NewFolderName}"
		speak "Creating ${NewFolder}..."
		If (Test-Path -Path "${HomeRoot}") {
			If (!(Test-Path -Path "${NewFolder}")) {
				If (!($Simulate)) { 
					New-Item -Path "${HomeRoot}" -Name "$NewFolderName" -ItemType Directory > $null
					$Acl = Get-Acl "$NewFolder"
					$Ar = New-Object System.Security.AccessControl.FileSystemAccessRule("$($Account.UserName)", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
					$Ao = New-Object System.Security.Principal.NTAccount("$NetBIOSDomain", "$($Account.UserName)")
					$Acl.SetAccessRule($Ar)
					$Acl.SetOwner($Ao)
					Set-Acl "$NewFolder" $Acl > $null
					$Ar = $null
					$Ao = $null
					$Acl = $null
				}
			} Else {
				speak "WARNING: The folder ${NewFolder} already exists."
			}	
		} Else {
			speak "ERROR: Root folder ${HomeRoot} does not exist."
		}
		speak ""
	}
	speak "Sleeping to allow AD replication to work..."
	if (!($Simulate)) { Start-Sleep -s 15 }
}

Function CreateStaffAccounts($Accounts) {
	speak ""
	speak "" -Time
	speak "=================================="
	speak " Creating new staff accounts... "
	speak "=================================="
	speak ""
	ForEach ($Account In $Accounts) {
		$OU = "${LDAPStaffUsersOU}"
		$Name = "$($Account.Surname) $($Account.FirstName.SubString(0,1)) $($Account.Title)"
		speak "${Name}..."
		$i = 2; $Created = $false; Do {
			$DN = "CN=${Name},${OU}"
			$SearchResult = Get-ADUser -LDAPFilter "(DistinguishedName=${DN})" -Properties CN | Select sAMAccountname
			If (!($SearchResult)) {
				speak "...Creating as ${DN}..."
				$Created = $true
			} Else {
				speak "...${DN} exists. Trying another account name..."
				If ($Account.MiddleNames) {
					$Name = "$($Account.FirstName) $($Account.MiddleNames.SubString(0,$i)). $($Account.Surname)"
				} else {
					$Name = "$($Account.FirstName) $($Account.Surname) ${i}"
				}
			}
			$i++
		} Until ($Created)
		
		$i = 2; $Created = $false; Do {
			$Email = "$($Account.FirstName).$($Account.Surname)${EmailSuffix}".ToLower() -Replace " ", "."
			$SearchResult = Get-ADUser -LDAPFilter "(EMailAddress=${Email})" -Properties EMailAddress | Select EMailAddress
			If (!($SearchResult)) { $Created = $true }
				Else {
				$Email = $Email = "$($Account.FirstName).$($Account.Surname)${i}${EmailSuffix}".ToLower() -Replace " ", "."
			}
			$i++			
		} Until ($Created)

		$Args = [pscustomobject]@{
			Name = "${Name}"
			GivenName = "$($Account.FirstName)"
			Surname = "$($Account.Surname)"
			DisplayName = "${Name}"
			EmailAddress = "${Email}"
			POBox = "$($Account.ID)"
			City = "$UserCity"
			UserPrincipalName = "${Email}"
			SamAccountName = "$($Account.UserName)"
			HomeDirectory = "${StaffHomeRoot}\$($Account.UserName)"		
			HomeDrive = "$StaffHomeDrive"
			Company = "Staff"
			AccountPassword = (ConvertTo-SecureString -AsPlainText $Account.Password -Force)
			Path = "$OU"
			ChangePasswordAtLogon = $true
			CannotChangePassword = $false
# ****** Need a department field from somewhere ****** #
		}
		speak "Attempting to creating user CN=$($Account.Name),${OU}..."
		If (!($Simulate)) { $Args | New-ADUser }
		
		speak "Adding user to group G Staff.."
		If (!($Simulate)) { 
			Add-ADGroupMember "G Staff" -Members "${DN}"
			Add-ADGroupMember "SW Filter Staff" -Members "${DN}"
		}
		If (!($Simulate) -and ($InDev)) {
			# Hide user account from Exchange address lists if we are running in development mode.
			$UserProp = get-aduser -Identity "$($Account.Username)" -Properties *
			$UserProp.msExchHideFromAddressLists = "True"
			Set-ADUser -Instance $UserProp
			$UserProp = $null
		}
		RegisterOffice365 -Email "${EmailAddress}" -LicenseType "Staff"
		$NewStaffAccountInfo = New-Object PSObject
		$NewStaffAccountInfo | Add-Member -MemberType NoteProperty -Name Name -Value "$($Account.FirstName) $($Account.Surname)"
		$NewStaffAccountInfo | Add-Member -MemberType NoteProperty -Name EmailAddress -Value "${Email}"
		$NewStaffAccountInfo | Add-Member -MemberType NoteProperty -Name UserName -Value "$($Account.UserName)"
		$NewStaffAccountInfo | Add-Member -MemberType NoteProperty -Name Password -Value "$($Account.Password)"
		$NewStaffAccounts.Add($NewStaffAccountInfo) > $null
		$NewStaffAccountInfo = $null
		speak ""
	}
}

Function DisableStaffLeavers() {
	speak ""
	speak "" -Time
	speak "=========================="
	speak " Disable staff leavers... "
	speak "=========================="
	speak ""
	$Ti = 0; $FN = 1; $MN = 2; $SN = 3; $SC = 4; $ID = 5; $RT = 6 
	$Accounts = Get-ADUser -SearchBase "${LDAPStaffUsersOU}" `
		-LDAPFilter "(objectclass=user)"`
		-Properties DistinguishedName,CN,sAMAccountName,Description,Department,Enabled `
		-SearchScope OneLevel
	ForEach ($Account In $Accounts) {
		$DN = $Account.DistinguishedName
		$CN = $Account.CN
		$UN = $Account.sAMAccountName
		$Ds = $Account.Description
		$Dp = $Account.Department
		If ("$($DN.Substring(0,4))" -ne "CN=!" ) {
			$ReadStream = [System.IO.StreamReader] "${SIMSStaffList}"
			$SIMS = $ReadStream.ReadLine()
			$ExistsInSIMS = $false
			While (!($ReadStream.EndOfStream)) {
				$SIMS = $ReadStream.ReadLine() -Replace """", ""
				$SIMS = $SIMS.Split(",")
				$SIMSUN = "staff$($SIMS[$SC])".ToLower().Trim()
				If ("$UN" -eq "$SIMSUN") { $ExistsInSIMS = $true }
			}
			$ReadStream.Close()
			If (!($ExistsInSIMS)) {
				speak "No staff in SIMS matching $UN, $CN"
				$Disable = @{}
				If ("$($Account.Enabled)" -eq "False") {
					speak "User account is already disabled"
					$PrevDisabledStaff.Add($Account) > $null
				} Else {
					speak "User account is enabled"
					$Now = Now
					$Ds = "${Ds} Leaver;${Now};"
					If (!($Simulate)) { Set-ADUser -Identity ${UN} -Enabled $false -Description "${Ds}" }
					$DisabledStaff.Add($Account) > $null
					speak "Updated Enabled to False"
					speak "Updated Description to $Ds"
				}
				speak ""
			}
		} Else {
			speak "$UN, $CN is excluded from processing."
			speak ""
		}
	}
}