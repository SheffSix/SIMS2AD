Function TransferStudents {
	If ($StudentUsers) {
		$FN = 0; $MN = 1; $SN = 2; $AN = 3; $YG = 4; $YT = 5; $BD = 6; $UP = 7
		$NewStudents = New-Object System.Collections.ArrayList
		speak ""
		speak "" -Time
		speak "================================"
		speak " Processing on roll students... "
		speak "================================"
		speak ""
		If (Test-Path -Path $SIMSStudentList) {
			$ReadStream = [System.IO.StreamReader] "${SIMSStudentList}"
			$User = $ReadStream.ReadLine()
		} else {
			speak "ERROR: ${SIMSStudentList} does not exist." -Time
			bork 1
		}
		While (!($ReadStream.EndOfStream)) {
			$User = $ReadStream.ReadLine() -Replace """", ""
			$User = $User.Split(",")
			$User[$YG] = $User[$YG] -Replace " ", "" -Replace "Year", ""
			$UserPassword = DateToPassword $User[$BD]
			If ($User[$AN] -eq "" -or !($User[$AN])) {
				speak "$($User[$FN]) $($User[$SN]) has no admission number" -Time
			} ElseIf ($User[$YG] -eq "" -or !($User[$YG])) {
				speak "$($User[$FN]) $($User[$SN]) has no year group specified" -Time
			} ElseIf ($User[$YT] -eq "" -or !($User[$YT])) {
				speak "$($User[$FN]) $($User[$SN]) has no year taught in specified" -Time
			} Else {
				$User[$AN] = ConvertAdmissionNumber $User[$AN]
				$UserPrefix = GetPrefix $User[$YT]
				$UserName = "${UserPrefix}$($User[$FN].SubString(0,1).ToLower())$($User[$SN].SubString(0,1).ToLower())$($User[$AN])"
				# If the user does not have a middle name, force it to be empty
				If ($User[$MN] -eq "" -or $User[$MN] -eq " "){$User[$MN] = $NULL}
				# If the user does not have a UPN, set a temporary one
				If ($User[$UP] -eq "" -or $User[$UP] -eq " " -or !($User[$UP])) { $User[$UP] = "TempUPN_$($User[$AN])"}
				$UserAccount = SearchUser -Search $UserName -LooseSearch $User[$AN]
				If (!($UserAccount)) {
					$NewStudent = @{}
					$NewStudent.FirstName = $User[$FN]
					$NewStudent.MiddleNames = $User[$MN]
					$NewStudent.Surname = $User[$SN]
					$NewStudent.UserName = $UserName
					$NewStudent.Password = $UserPassword
					$NewStudent.Prefix = $UserPrefix
					$NewStudent.YearGroup = $User[$YG]
					$NewStudent.UPN = $User[$UP]
					$NewStudents.Add($NewStudent) > $NULL
					speak "Will create new user: $($NewStudent.FirstName) $($NewStudent.Surname), Taught in Year $($User[$YT])"
					speak ""
					$NewStudent = $NULL
				} ElseIf ($UserAccount -ne "ERROR") {
					If ($UserAccount.sAMAccountName -eq $Username) {
						speak "User $Username exists"
					} Else {
						speak "User $Username exists as $($UserAccount.sAMAccountName)"
						$Username = $UserAccount.sAMAccountName
					}
					UpdateUser -Account $UserAccount -GivenName $User[$FN] -MiddleNames $User[$MN] -Surname $User[$SN] -DisplayName "$($User[$FN]) $($User[$SN])" -Company "Student" -Department "${UserPrefix}prefix" -POBox $User[$UP] -UserName "$UserName"]
					speak ""
				}
			}
			#bork 0
		}
		$ReadStream.Close()
		If ($NewStudents) {
			CreateStudentAccounts $NewStudents
			CreateStudentHomeFolders $NewStudents
		}
		If (!($NoLeavers)) { DisableStudentLeavers }
	}
}

Function DisableStudentLeavers() {
	speak ""
	speak "" -Time
	speak "============================"
	speak " Disable student leavers... "
	speak "============================"
	speak ""
	#Are we in the Year 14 leavers buffer period? This is between 1st August and 31st October.
	#During this time, Year 14 leavers will not be disabled to allow UCAS applications to be completed.
	If ($ThisMonth -ge "08" -and $ThisMonth -le "10") {
		$Y14Buffer = $true
		speak "Year 14 buffer period is in effect. Year 14's will not be processed until 1st November."
	}
	$FN = 0; $MN = 1; $SN = 2; $AN = 3; $YG = 4; $YT = 5; $BD = 6; $UP = 7
	$Accounts = Get-ADUser -SearchBase "${LDAPStudentUsersOU}" `
		-LDAPFilter "(objectclass=user)"`
		-Properties DistinguishedName,CN,sAMAccountName,Description,Department,Enabled
	ForEach ($Account In $Accounts) {
		$DN = $Account.DistinguishedName
		$CN = $Account.CN
		$UN = $Account.sAMAccountName
		$Ds = $Account.Description
		$Dp = $Account.Department
		If ("$($DN.Substring(0,4))" -ne "CN=!" ) {
			If ("$Dp" -ne "${Year14Prefix}prefix" -or $Y14Buffer -ne $true) {
				$ReadStream = [System.IO.StreamReader] "${SIMSStudentList}"
				$SIMS = $ReadStream.ReadLine()
				$ExistsInSIMS = $false
				While (!($ReadStream.EndOfStream)) {
					$SIMS = $ReadStream.ReadLine() -Replace """", ""
					$SIMS = $SIMS.Split(",")
					$SIMS[$AN] = ConvertAdmissionNumber $SIMS[$AN]
					If ("$UN" -Match "$($SIMS[$AN])$") { $ExistsInSIMS = $true }
				}
				$ReadStream.Close()
				If (!($ExistsInSIMS)) {
					speak "No student in SIMS matching $UN, $CN"
					$Disable = @{}
					If ("$($Account.Enabled)" -eq "False") {
						speak "User account is already disabled"
						$PrevDisabledStudents.Add($Account) > $null
					} Else {
						speak "User account is enabled"
						$Now = Now
						$Ds = "${Ds} Leaver;${Now};"
						If (!($Simulate)) { Set-ADUser -Identity ${UN} -Enabled $false -Description "${Ds}" }
						Else { speak "Set-ADUser -Identity ${UN} -Enabled $false -Description ""${Ds}""" }
						$DisabledStudents.Add($Account) > $null
						speak "Updated Enabled to False"
						speak "Updated Description to $Ds"
					}
					speak ""
				}
			} Else {
				speak "Not processing Year 14 $UN, $CN"
				speak ""
			}
		} Else {
			speak "$UN, $CN is excluded from processing."
			speak ""
		}
	}
}

Function CreateStudentHomeFolders($Accounts) {
	speak ""
	speak "" -Time
	speak "===================================="
	speak " Setting up student home folders... "
	speak "===================================="
	speak ""
	ForEach ($Account In $Accounts) {
		$HomeRoot = "${StudentHomeRoot}\$($Account.Prefix)prefix_home"
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
				} Else {
					speak "New-Item -Path ""${HomeRoot}"" -Name ""$NewFolderName"" -ItemType Directory > $null"
					speak "`$Acl = Get-Acl ""$NewFolder"""
					speak "`$Ar = New-Object System.Security.AccessControl.FileSystemAccessRule(""$($Account.UserName)"", ""FullControl"", ""ContainerInherit,ObjectInherit"", ""None"", ""Allow"")"
					speak "`$Ao = New-Object System.Security.Principal.NTAccount(""$NetBIOSDomain"", ""$($Account.UserName)"")"
					speak "`$Acl.SetAccessRule(`$Ar)"
					speak "`$Acl.SetOwner(`$Ao)"
					speak "Set-Acl ""$NewFolder"" `$Acl > `$null"
					speak "`$Ar = `$null"
					speak "`$Ao = `$null"
					speak "`$Acl = `$null"
				}
			} Else {
				speak "WARNING: The folder ${NewFolder} already exists."
			}	
		} Else {
			speak "ERROR: Root folder ${HomeRoot} does not exist."
		}
		speak ""
	}
}

Function CreateStudentAccounts($Accounts) {
	speak ""
	speak "" -Time
	speak "=================================="
	speak " Creating new student accounts... "
	speak "=================================="
	speak ""
	ForEach ($Account In $Accounts) {
		$OU = "OU=$($Account.prefix)prefix,${LDAPStudentUsersOU}"
		$Name = "$($Account.FirstName) $($Account.Surname)"
		speak "${Name}..."
		$i = 2; $Created = $false; Do {
			$DN = "CN=${Name},${OU}"
			$SearchResult = Get-ADUser -LDAPFilter "(DistinguishedName=${DN})" -Properties CN | Select sAMAccountname
			If (!($SearchResult)) {
				speak "Creating as ${DN}..."
				$Created = $true
			} Else {
				speak "${DN} exists. Trying another account name..."
				If ($Account.MiddleNames) {
					$Name = "$($Account.FirstName) $($Account.MiddleNames.SubString(0,$i)). $($Account.Surname)"
				} else {
					$Name = "$($Account.FirstName) $($Account.Surname) ${i}"
				}
			}
			$i++
		} Until ($Created)
		If ("$($Account.Prefix)" -eq "$Year7Prefix") {
			$ChangePasswordAtLogon = $false
			$CannotChangePassword = $true
		} Else {
			$ChangePasswordAtLogon = $true
			$CannotChangePassword = $false
		}
		$Args = [pscustomobject]@{
			Name = "${Name}"
			GivenName = "$($Account.FirstName)"
			Surname = "$($Account.Surname)"
			DisplayName = "$($Account.FirstName) $($Account.Surname)"
			EmailAddress = "$($Account.UserName)${EmailSuffix}" -Replace " ", ""
			POBox = "$($Account.UPN)"
			City = "$UserCity"
			UserPrincipalName = "$($Account.UserName)${UserNameSuffix}"
			SamAccountName = "$($Account.UserName)"
			ProfilePath = "$MandatoryProfile"
			HomeDirectory = "${StudentHomeRoot}\$($Account.Prefix)prefix_home\$($Account.UserName)"
			HomeDrive = "$StudentHomeDrive"
			Department = "$($Account.Prefix)prefix"
			Company = "Student"
			AccountPassword = (ConvertTo-SecureString -AsPlainText $Account.Password -Force)
			Path = "$OU"
			ChangePasswordAtLogon = $ChangePasswordAtLogon
			CannotChangePassword = $CannotChangePassword
			Enabled = $true
		}
		speak "Attempting to creating user ${DN}..."
		If (!($Simulate)) { $Args | New-ADUser }
		Else { "$Args | New-ADUser" }
		speak "Adding user to group $($Account.Prefix)prefix.."
		If (!($Simulate)) { Add-ADGroupMember "$($Account.Prefix)prefix" -members "CN=${Name},${OU}" }
		Else { "Add-ADGroupMember ""$($Account.Prefix)prefix"" -members ""CN=${Name},${OU}""" }
		If (!($Simulate) -and ($InDev)) {
			# Hide user account from Exchange address lists if we are running in development mode.
			speak "In-Dev Mode: Sleep for 5 seconds..."
			Start-Sleep -s 5
			speak "In-Dev Mode: Hide user from Exchange address lists..."
			$UserProp = get-aduser -LDAPFilter "(DistinguishedName=${DN})" -Properties *
			$UserProp.msExchHideFromAddressLists = "True"
			Set-ADUser -Instance $UserProp
			$UserProp = $null
		} ElseIf ($InDev) {
			speak "In-Dev Mode: Sleep for 5 seconds..."
			speak "Start-Sleep -s 5"
			speak "In-Dev Mode: Hide user from Exchange address lists..."
			speak "$UserProp = get-aduser -LDAPFilter ""(DistinguishedName=${DN})"" -Properties *"
			speak "$UserProp.msExchHideFromAddressLists = ""True"""
			speak "Set-ADUser -Instance $UserProp"
			speak "$UserProp = $null"
		}
		ForEach ($p in $Office365MailboxPrefixes) {
			If ("$P" -eq "$($Account.Prefix)") {
				$CreateOffice365Mailbox = $true
			}
		}
		If ($CreateOffice365Mailbox -eq $true) {
			RegisterOffice365 -Email "$($Account.UserName)${EmailSuffix}" -LicenseType "Student"
			RegisterExchangeRemoteMailbox -Identity "$($Account.UserName)${UserNameSuffix}" -RemoteRoutingAddress "SMTP:$($Account.UserName)@${MSOLDomain}"
		} Else {
			RegisterOffice365 -Email "$($Account.UserName)${EmailSuffix}" -LicenseType "Student" -NoMailbox
			RegisterExchangeMailbox -Identity "$($Account.UserName)" -Database "$($Account.Prefix)prefix"
		}
		
		$CreateOffice365Mailbox = $null
		
		$NewStudentAccountInfo = New-Object PSObject
		$NewStudentAccountInfo | Add-Member -MemberType NoteProperty -Name Name -Value "$($Account.FirstName) $($Account.Surname)"
		$NewStudentAccountInfo | Add-Member -MemberType NoteProperty -Name EmailAddress -Value "$($Account.UserName)${EmailSuffix}"
		$NewStudentAccountInfo | Add-Member -MemberType NoteProperty -Name UserName -Value "$($Account.UserName)"
		$NewStudentAccountInfo | Add-Member -MemberType NoteProperty -Name Password -Value "$($Account.Password)"
		$NewStudentAccounts.Add($NewStudentAccountInfo) > $null
		$NewStudentAccountInfo = $null
		
		speak ""
	}
	speak "Sleeping for 15 seconds to allow AD replication to work..."
	If (!($Simulate)) { Start-Sleep -s 15 }
	Else { Speak "Start-Sleep -s 15"}
}

Function DateToPassword($Date) {
	$Password = Get-Date -Date $Date -f "yyyyMMdd"
	Return $Password
}

Function GetPrefix($YearGroup) {
	Switch ($YearGroup) {
		"7" {$Prefix = $Year7Prefix; break}
		"8" {$Prefix = $Year8Prefix; break}
		"9" {$Prefix = $Year9Prefix; break}
		"10" {$Prefix = $Year10Prefix; break}
		"11" {$Prefix = $Year11Prefix; break}
		"12" {$Prefix = $Year12Prefix; break}
		"13" {$Prefix = $Year13Prefix; break}
		"14" {$Prefix = $Year14Prefix; break}
	}
	Return $Prefix
}

Function GetKeyStage($YearGroup) {
	Switch($YearGroup) {
		"7" {$KeyStage = "3"; break}
		"8" {$KeyStage = "3"; break}
		"9" {$KeyStage = "3"; break}
		"10" {$KeyStage = "4"; break}
		"11" {$KeyStage = "4"; break}
		"12" {$KeyStage = "5"; break}
		"13" {$KeyStage = "5"; break}
		"14" {$KeyStage = "5"; break}
		default { $KeyStage = $null }
	}
	Return $KeyStage
}

Function ConvertAdmissionNumber($AdmissionNo) {
	# Students admitted from September 2013 onwards use the full six-digit admission
	# number. Students admitted before September 2013 use the four-digit admission
	# number. The exception to this are users 8723 and 8725, who were created manually
	# with a four-digit number.
	[int]$TestNumber = [convert]::toInt32($AdmissionNo, 10)
	If ($TestNumber -lt $FirstSixDigitAdNo -or $TestNumber -eq 8723 -or $TestNumber -eq 8725) {
		$AdmissionNo = "$TestNumber"
		Return $AdmissionNo
	} else {
		Return $AdmissionNo
	}
}