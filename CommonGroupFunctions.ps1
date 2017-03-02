Function NewGroup($Path, $Name) {
	If (!(SearchGroup "$Name")) {
		If (!($Simulate)) { New-ADGroup -Path "$Path" -Name "$Name" -GroupScope Universal }
		Else { speak "New-ADGroup -Path ""$Path"" -Name ""$Name"" -GroupScope Universal" }
		$NewGroup = New-Object PSObject
		$NewGroup | Add-Member -MemberType NoteProperty -Name "Name" -Value "$Name"
		$NewGroup | Add-Member -MemberType NoteProperty -Name "Path" -Value "$Path"
		$NewGroups.Add($NewGroup) > $null
	} Else { Speak "Unable to create group $Name. A group with that name already exists."}
}

Function RemoveGroup() {
	# $OldGroup = New-Object PSObject
	# $OldGroup | Add-Member -MemberType NoteProperty -Name "Name" -Value "$Name"
	# $OldGroup | Add-Member -MemberType NoteProperty -Name "Path" -Value "$Path"
	# $RemovedGroups.Add($OldGroup) > $null
}

Function SearchGroup($sAMAccountName) {
	$Result = Get-ADGroup -SearchBase $LDAPRoot -LDAPFilter "(sAMAccountName=$sAMAccountName)"
	Return $Result
}

Function EnumerateGroups($SearchBase) {
	$Result = Get-ADGroup -SearchBase $SearchBase -Filter {ObjectClass -eq "group"}
	Return $Result
}

Function EnumerateGroupMembers($Name) {
	$Result = Get-ADGroupMember -Identity "$Name"
	Return $Result
}

Function GetGroupMembership($Group,$DN) {
	$objEntry = [adsi]("LDAP://"+$DN)
	$objEntry.MemberOf | where { $_ -match $group }
}

Function AddGroupMembership() {
	Param(
		[Parameter(position=1)][string]$Group,
		[Parameter(position=2)][string]$DN,
		[string]$ObjectType = "User"
	)
	speak "CHANGE: Adding $objecttype to group ${Group}"
	If (!($Simulate)) { Add-ADGroupMember "$Group" -members "$DN" }
	Else {speak "Add-ADGroupMember ""$Group"" -members ""$DN"""}
	$NewGroupMember = New-Object PSObject
	$NewGroupMember | Add-Member -MemberType NoteProperty -Name "Group" -Value "$Group"
	$NewGroupMember | Add-Member -MemberType NoteProperty -Name "Member" -Value "$DN"
	$NewGroupMemberships.Add($NewGroupMember) > $null
}

Function RemoveGroupMembership($Group,$sAMAccountName,$Name) {
	If ($Name) { speak "Removing $Name, $sAMAccountName from $Group." } Else { speak "Removing $sAMAccountName from $Group." }
	If (!($Simulate)) { Remove-ADGroupMember -Identity "$Group" -Members "$sAMAccountName" -Confirm:$false }
	Else { speak "Remove-ADGroupMember -Identity ""$Group"" -Members ""$sAMAccountName"" -Confirm:$false" }
	$OldGroupMember = New-Object PSObject
	$OldGroupMember | Add-Member -MemberType NoteProperty -Name "Group" -Value "$Group"
	If ($Name) { $OldGroupMember | Add-Member -MemberType NoteProperty -Name "Member" -Value "$Name" }
	Else { $OldGroupMember | Add-Member -MemberType NoteProperty -Name "Member" -Value "$sAMAccountName" }
	$RemovedGroupMemberships.Add($OldGroupMember) > $null
}