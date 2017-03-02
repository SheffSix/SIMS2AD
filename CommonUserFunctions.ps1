Function SearchUser() {
	Param (
		[string]$Search,
		[string]$LooseSearch,
		[string]$POBox,
		[switch]$Quiet
	)
	If ($Search) {
		If (!($Quiet)) { speak "Searching for user account ${Search}" -Time }
		$SearchResult = Get-ADUser -SearchBase $LDAPRoot `
			-LDAPfilter "(sAMAccountName=${Search})" `
			-Properties ${LDAPUserPropertiesArray} `
			| Select ${LDAPUserPropertiesArray}
		If ($SearchResult -and !($Quiet)){ speak "User Accounts found" }
		ElseIf (!($Quiet)) { speak "No user found for the search term ${Search}" }
	}
	if (!($SearchResult) -and $LooseSearch) {
		If (!($Quiet)) { speak "Searching for loose search term ${LooseSearch}" }
		$SearchResult = Get-ADUser -SearchBase $LDAPRoot `
			-Properties ${LDAPUserPropertiesArray} `
			-LDAPFilter "(&(sAMAccountName=*${LooseSearch}*)(!(sAMAccountName=DCA58E22-8080-44AF-8))(!(sAMAccountName=SM_*))(!(sAMAccountName=SUPPORT_388945a0)))" `
			| Select ${LDAPUserPropertiesArray}
		If ($SearchResult) {
			If (!($Quiet)) { speak "User Account(s) found " }
			If ($SearchResult.Count -gt 1) {
				If (!($Quiet)) { 
					speak "ERROR: Duplicate users found for the loose search term ${LooseSearch}"
					speak ""
				}
				Return "ERROR"
			}
		} ElseIf (!($Quiet)) { speak "No user found for the loose search term ${LooseSearch}" }
	}
	if (!($SearchResult) -and $POBox) {
		If (!($Quiet)) { speak "Searching for POBox ${POBox}" }
		$SearchResult = Get-ADUser -SearchBase $LDAPRoot `
			-Properties ${LDAPUserPropertiesArray} `
			-LDAPFilter "(postOfficeBox=${POBox})" `
			| Select ${LDAPUserPropertiesArray}
		If ($SearchResult) {
			If (!($Quiet)) { speak "User Account(s) found" }
			If ($SearchResult.Count -gt 1) {
				If (!($Quiet)) { 
					speak "ERROR: Duplicate users found for the POBox ${POBox}"
					speak ""
				}
				Return "ERROR"
			}
		} ElseIf (!($Quiet)) { speak "No user found for the POBox ${POBox}" }
	}
	If (!($Search) -and !($LooseSearch) -and !($POBox)) {
		If (!($Quiet)) { 
			speak "ERROR: Niether Search, LooseSearch or POBox were specified." -Time
			speak ""
		}
	}
	
	if (!($SearchResult)) {
		Return $NULL
	}
	Return $SearchResult
}

Function RegisterOffice365() {
	Param(
		[string]$Email,
		[string]$LicenseType,
		[switch]$NoMailbox
	)
	$Account = @{}
	$Account.Email = $Email
	$Account.LicenseType = $LicenseType
	If ($NoMailbox -eq $true) { $Account.NoMailbox = $true }
		else { $account.NoMailbox = $false }
	$NewOffice365Accounts.Add($Account) > $null
}

Function ProcessOffice365() {
	If ($NewOffice365Accounts -and ($ProcessMSOL -eq $true)) {
		speak ""
		speak "" -Time
		speak "==================================="
		speak " Activating Office 365 licenses..."
		speak "==================================="
		speak ""
		speak "Connecting to MSOL Service..." -NoNewLine
		If (!($Simulate)) {
			If ( Get-Module -ListAvailable -Name MSOnline ) {
				try {
					try { Import-Module MSOnline } catch { bork 89 }
					$Connected = Get-MsolDomain -ErrorAction SilentlyContinue
				} catch {
					$Connected = $null
				}	
			} Else {
				try {
					try { Import-Module DirSync } catch { bork 98 }
					$Connected = Get-MsolDomain -ErrorAction SilentlyContinue
				} catch {
					$Connected = $null
				}
			}
			If (!($Connected)) {
				$Creds = Get-O365Creds
				Connect-MsolService -Credential $Creds
				$Creds = $null
				speak "Done!"
			} Else {
				speak "Already connected!"
			}
		}
		Office365Sync
		ForEach ($Account In $NewOffice365Accounts) {
			If (!($Simulate)) { Set-MsolUser -UserPrincipalName "$($Account.Email)" -UsageLocation GB }
			Else { speak "Set-MsolUser -UserPrincipalName ""$($Account.Email)"" -UsageLocation GB"}
			If ($Account.LicenseType -eq "Student") {
				speak "Activating Office 365 Student license $($Account.Email)..."
				If (!($Simulate)) { 
					If ($Account.NoMailbox -eq $false) { Set-MsolUserLicense -UserPrincipalName "$($Account.Email)" -AddLicenses $Office365StudentLicense -LicenseOptions $Office365StudentLicenseOptions }
					Else { Set-MsolUserLicense -UserPrincipalName "$($Account.Email)" -AddLicenses $Office365StudentLicense -LicenseOptions $Office365StudentLicenseOptionsNoMBX }
				} else {
					If ($Account.NoMailbox -eq $false) { speak "Set-MsolUserLicense -UserPrincipalName ""$($Account.Email)"" -AddLicenses $Office365StudentLicense -LicenseOptions $Office365StudentLicenseOptions # Student with 365 mailbox" }
					Else { speak "Set-MsolUserLicense -UserPrincipalName ""$($Account.Email)"" -AddLicenses $Office365StudentLicense -LicenseOptions $Office365StudentLicenseOptionsNoMBX # Student with Exchange mailbox" }
				}
			} ElseIf ($Account.LicenseType -eq "Staff") {
				speak "Activating Office 365 Staff license $($Account.Email)..."
				If (!($Simulate)) { Set-MsolUserLicense -UserPrincipalName "$($Account.Email)" -AddLicenses $Office365StaffLicense -LicenseOptions $Office365StaffLicenseOptions }
				Else { speak "Set-MsolUserLicense -UserPrincipalName ""$($Account.Email)"" -AddLicenses $Office365StaffLicense -LicenseOptions $Office365StaffLicenseOptions" }
			} Else {
				speak "Invalid license type specified. Unable to license Office 365 account $($Account.Email)."
			}
		}
		Office365Sync -s 60
	}
}

Function Office365Sync() {
	Param (
		[int]$S = 300
	)
	If (!($Simulate)) {
		If (Get-Module -ListAvailable -Name ADSync) {
			try {
				try { Import-Module MSOnline } catch { bork 143 }
				$Connected = Get-MsolDomain -ErrorAction SilentlyContinue
			} catch {
				$Connected = $null
			}
		}
	}
	If (Get-Module -ListAvailable -Name ADSync) {
		speak "Starting AD Sync Cycle (ADSync)..."
		If (!($Simulate)) { Start-ADSyncSyncCycle }
	} Else {
		speak "Starting Online Coexistence Sync (DirSync)..."
		If (!($Simulate)) { Start-OnlineCoexistenceSync }
	}
	speak "Sleeping for ${S} seconds to allow DirSync to finish..."
	If (!($Simulate)) { Start-Sleep -s $S }
	Else { Speak "Start-Sleep -s $S"}
}

Function RegisterExchangeMailbox() {
	Param (
		[string]$Identity,
		[string]$Database
	)
	$Account = @{}
	$Account.Identity = $Identity
	$Account.Database = $Database
	$NewExchangeMailboxes.Add($Account) > $null
}

Function RegisterExchangeRemoteMailbox() {
	param (
		[string]$Identity,
		[string]$RemoteRoutingAddress
	)
	$Account = @{}
	$Account.Identity = $Identity
	$Account.RemoteRoutingAddress = $RemoteRoutingAddress
	$NewExchangeRemoteUsers.Add($Account) > $null
}

Function ProcessExchange () {
	If ($NewExchangeMailboxes -and ($ProcessExchange -eq $true)) {
		speak ""
		speak "" -Time
		speak "================================"
		speak " Enabling Exchange mailboxes..."
		speak "================================"
		speak ""
		ForEach ($Account in $NewExchangeMailboxes) {
			speak "Enabling Exchange Mailbox for $($Account.Identity)..."
			If (!($Simulate)) { try { Enable-Mailbox -Identity "$($Account.Identity)" -Database "$($Account.Database)" } catch {} }
			Else { speak "Enable-Mailbox -Identity ""$($Account.Identity)"" -Database ""$($Account.Database)""" }
		}
	}
	If ($NewExchangeRemoteUsers -and ($ProcessExchange -eq $true)) {
		speak ""
		speak "" -Time
		speak "======================================="
		speak " Enabling Remote Exchange mailboxes..."
		speak "======================================="
		speak ""
		ForEach ($Account in $NewExchangeRemoteUsers) {
			speak "Enabling Remote mailbox for $($Account.Identity)..."
			If (!($Simulate)) { try { Enable-RemoteMailbox -Identity "$($Account.Identity)" -RemoteRoutingAddress "$($Account.RemoteRoutingAddress)" } catch {} }
			Else { speak "Enable-RemoteMailbox -Identity ""$($Account.Identity)"" -RemoteRoutingAddress ""$($Account.RemoteRoutingAddress)""" }
		}
	}
}

Function UpdateUser() {
	Param (
		[array]$Account,
		[string]$Title,
		[string]$UserName,
		[string]$GivenName,
		[string]$MiddleNames,
		[string]$Surname,
		[string]$DisplayName,
		[string]$Company,
		[string]$Department,
		[string]$POBox
	)
	# Update GivenName
	If ($GivenName) { UpdateAccountProperty "$($Account.DistinguishedName)" "GivenName" "$GivenName" "$($Account.GivenName)" }
	
	# Update sn
	If ($Surname) { UpdateAccountProperty "$($Account.DistinguishedName)" "Surname" "$Surname" "$($Account.sn)" }
	
	# Update DisplayName
	If ($Displayname) { UpdateAccountProperty "$($Account.DistinguishedName)" "DisplayName" "$DisplayName" "$($Account.DisplayName)" }
	
	# Update Department
	If ($Department) { UpdateAccountProperty "$($Account.DistinguishedName)" "Department" "$Department" "$($Account.Department)" }
	
	# Update Company
	If ($Company) { UpdateAccountProperty "$($Account.DistinguishedName)" "Company" "$Company" "$($Account.Company)" }
	
	# Update PostofficeBox
	If ($POBox) { UpdateAccountProperty "$($Account.DistinguishedName)" "POBox" "$POBox" "$($Account.PostOfficeBox)" }
	
	# Enabled disabled account
	EnableDisabledAccount "$($Account.DistinguishedName)" "$($Account.Enabled)" "$($Account.Description)"
	
	# Students only
	If ("$Company" -eq "Student") {
		# Update ProfilePath
		UpdateAccountProperty "$($Account.DistinguishedName)" "ProfilePath" "$MandatoryProfile" "$($Account.ProfilePath)"
		
		# Update Prefix group membership
		If (!(GetGroupMembership "$Department" $($Account.DistinguishedName))) {
			AddGroupMembership "$Department" "$($Account.DistinguishedName)"
		}
		
		# Move user account
		$AccountMoved = MoveUserAccount $Account "OU=${Department},${LDAPStudentUsersOU}" "$MiddleNames"
		If ($AccountMoved) {
			$Account = $AccountMoved
		}
	}
	
	If (!($AccountMoved)) {
		# Rename account if required
		If ("$Company" -eq "Student") {
			$AccountRenamed = UpdateAccountCommonName	"$($Account.DistinguishedName)" "$UserName" "$GivenName $Surname" "$($Account.CN)" "$MiddleNames"
		} ElseIf ("$Company" -eq "Staff") {
			$AccountRenamed = UpdateAccountCommonName	"$($Account.DistinguishedName)" "$UserName" "$Surname $($GivenName.SubString(0,1)) $Title" "$($Account.CN)"
		}
		If ($AccountRenamed) {
			speak "Account has been renamed, regathering account information."
			$Account = SearchUser -Search "$UserName"
		}
	}
	
	If (!($Displayname)) { 
		$DisplayName = $Account.CN
		UpdateAccountProperty "$($Account.DistinguishedName)" "DisplayName" "$DisplayName" "$($Account.DisplayName)" 
	}
}

Function MoveUserAccount () {
	Param(
		[Parameter(Position=1)][array]$Account,
		[Parameter(Position=2)][string]$TargetOU,
		[Parameter(Position=3)][string]$Extra
	)
	If (!($UserName)) { $UserName = "$($Account.sAMAccountName)" }
	$TargetDN = "CN=$($Account.CN),${TargetOU}"
	If ("$($Account.DistinguishedName)" -ne $TargetDN) {
		speak "User to be moved to $TargetOU..."
		$TempCN = "TemporaryCN-SIMS2AD"
		$TempDN = $Account.DistinguishedName -replace "CN=$($Account.cn)", "CN=$TempCN"
		speak "CHANGE: Renaming the account temporarily to avoid naming conflicts..."
		If (!($Simulate)) { Rename-ADObject -Identity "$($Account.DistinguishedName)" -NewName "$TempCN" }
		Else {speak "Rename-ADObject -Identity ""$($Account.DistinguishedName)"" -NewName ""$TempCN"""}
		speak "CHANGE: Moving the account..."
		if (!($Simulate)) { Move-ADObject -Identity "$TempDN" -TargetPath "${TargetOU}" }
		Else {speak "Move-ADObject -Identity ""$TempDN"" -TargetPath ""${TargetOU}"""}
		$TempDN = "CN=${TempCN},${TargetOU}"
		speak "NEW TEMP DN: $TempDN"
		if (!($Simulate)) { 
			$AccountRenamed = UpdateAccountCommonName "${TempDN}" "$UserName" "$($Account.GivenName) $($Account.sn)" "$TempCN" "$Extra"
			speak "Account has been moved, regathering account information."
			$Account = SearchUser -Search "$UserName"
		} Else {
			speak "$AccountRenamed = UpdateAccountCommonName ""${TempDN}"" ""$UserName"" ""$($Account.GivenName) $($Account.sn)"" ""$TempCN"" ""$Extra"""
			speak "Account has been moved, regathering account information."
			speak "$Account = SearchUser -Search ""$UserName"""
		}
		Return $Account
	}
	Return $false
}

Function UpdateAccountCommonName($DN, $sAMAccountName, $Required, $Current, $Extra) {
	If ("$Required" -ne "$Current") {
		speak "Attempting to rename user..."
		$Required = $required.Trim()
		$i = 2; $Renamed = $false; Do {
			$NewDN = $DN -replace $Current, $Required
			$SearchResult = Get-ADUser -LDAPFilter "(DistinguishedName=${NewDN})" -Properties CN | Select sAMAccountname, CN
			If (!($SearchResult)) {
				$Renamed = $TRUE
			} ElseIf ("$($SearchResult.sAMAccountName)" -eq "$sAMAccountName") {
				speak "... $Required is already named correctly."
				$Renamed = $TRUE
				$DoNothing = $TRUE
			} Else {
				speak "$Required already exists..."
				If ($Extra) {
					$Initial = "$($Extra.SubString(0,$i))."
					$Required = $Required -replace "(.*) (.*)", "`$1$Initial`$2"
				} else {
					$Required = "$Required $i"
				}				
			}
			$i++
		} Until ($Renamed)
		If (!($DoNothing)) {
			speak "CHANGE: Update CN from ""$Current"" to ""$Required"""
			speak "CHANGE: DistinguishedName is now $NewDN"
			If (!($Simulate)) { 
				Rename-ADObject -Identity "$DN" -NewName "$Required"
				Return $TRUE
			} Else {
				speak "Rename-ADObject -Identity ""$DN"" -NewName ""$Required"""
				Return $TRUE
			}
		}
	}
	Return $false
}

Function EnableDisabledAccount($DN, $Current, $Description) {
	# speak "Update of Enabled has been requested."
	# speak "Required Value : ${Required}"
	# speak "Current Value  : ${Current}"
	If ("$Current" -ne "True" ) {
		$regex = 'Leaver;(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[012])/(19|20)[0-9]{2} ([01][0-9]|2[0-3])(?::([0-5][0-9])){2};'
		If (!($Simulate)) {iex "Set-ADUser -Identity ""$DN"" -Enabled `$true"}
		Else { speak "iex ""Set-ADUser -Identity """"$DN"""" -Enabled `$true"""}
		speak "CHANGE: Enabled $DN"
		$NewDesc = [regex]::replace($Description, $Regex, "").Trim()
		UpdateAccountProperty "$DN" "Description" "$NewDesc" "$Description" 
	}
	# speak ""
}

Function UpdateAccountProperty($DN, $Property, $Required, $Current) {
	# speak "Update of $Property has been requested."
	# speak "Required Value : ${Required}"
	# speak "Current Value  : ${Current}"
	If ("$Required" -ne "$Current") {
		If (!($Simulate)) {iex "Set-ADUser -Identity ""$DN"" -$($Property) ""$Required"""}
		Else { speak "iex ""Set-ADUser -Identity """"$DN"""" -$($Property) """"$Required"""""}
		speak "CHANGE: Updated $Property from ""$Current"" to ""$Required"""
	}
	# speak ""
}