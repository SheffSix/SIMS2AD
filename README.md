# SIMS2AD

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
