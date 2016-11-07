#########################################################
#				Mailbox managing tool														#
#																												#
# 	Version 1.3																					#
# 	27.10.2016		04.11.2016														#
# 	Erstellt Christoph Becker														#
#########################################################

##	SET STRINGS AND VARIABLES
$finish_info = 'To finish the array, press "RETURN" without typing another account.'
$menutext = @"

===================
MAILBOX MANAGING UI
===================

1: Press '1' setting up 'Send As' permissions.
2: Press '2' setting up 'Send on Behalf' permissions.
3: Press '3' to grant a user 'Full Access' to another mailbox.
4: Press '4' to enable an Exchange user mailbox.
5: Press '5' to create new distribution lists.
6: Type 'DISABLE' to disable a users Exchange mailbox.
H: Press 'H' for help.
Q: Press 'Q' to quit this UI.

"@

$sa_string = @"

=================
ADD SEND AS PERMS
=================

"@

$so_string = @"

==================
ADD SEND ON BEHALF
==================

"@

$fa_string = @"

================
ADD FULL ACCESS
================

"@

$enable_string = @"

================
MAILBOX CREATION
================

"@

$dist_string = @"

===========================
DISTRIBUTION GROUP CREATION
===========================

"@

$disable_string = @"

===============
DISABLE MAILBOX
===============

"@

$help_string = @"
====
HELP
====

All scripts are using variables. The 'AFFECTED USER' variable is for the affected mailbox you want to give permissions on.
Example:
AFFECTED USER = TMUSTERMA
GRANT TO USER = SMUSTERFR
This would result that SMUSTERFR will get permissions on TMUSTERMA's mailbox

"@

##	SCRIPT
function Show-Menu
{
	param (
		[string]$Title = 'MAILBOX MANAGING UI'
	)
	cls
	$menutext
}

do
{
	Show-Menu
 	echo ""
 	$input = Read-Host "> MAKE A SELECTION"
  switch ($input)
	{

## GRANT SEND AS
		'1' {
			cls
			$menutext
			$finish_info
			$sa_string
			$mailbox = read-host "> AFFECTED USER"
			$users = @()
			do {
					$input = (Read-Host "> GRANT TO USER")
					if ($input -ne '') {$users += $input}
				}
			until ($input -eq '')
			foreach ($user in $users) {
				Get-Mailbox $mailbox -resultsize unlimited | Add-ADPermission -User $user -Extendedrights "Send As"
			}
			Write-Host "Press any key to return to main menu ..."
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}

## GRANT SEND ON BEHALF
		'2' {
			cls
			$menutext
			$finish_info
			$sa_string
			$mailbox = read-host "> AFFECTED USER"
			$users = @()
			do {
					$input = (Read-Host "> GRANT TO USER")
					if ($input -ne '') {$users += $input}
			}
			until ($input -eq '')
			foreach ($user in $users) {
				Get-Mailbox $mailbox -resultsize unlimited | set-mailbox -GrantSendOnBehalfTo @{Add=$user}
			}
			Write-Host "Press any key to return to main menu ..."
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}

## FULL ACCESS TO MAILBOX
		'3' {
			cls
			$menutext
			$finish_info
			$fa_string
			$mailbox = read-host "> AFFECTED USER"
			$users = @()
			do {
				$input = (Read-Host "> GRANT TO USER")
				if ($input -ne '') {$users += $input}
			}
			until ($input -eq '')
			foreach ($user in $users) {
				Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess
			}
			Write-Host "Press any key to return to main menu ..."
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}

## ENABLE MAILBOX
		'4' {
			cls
			$menutext
			$finish_info
			$enable_string
			$users = @()
			do {
				$input = (Read-Host "> AFFECTED USER")
				if ($input -ne '') {$users += $input}
			}
			until ($input -eq '')
			foreach ($user in $users) {
				Enable-Mailbox -Identity $user -RetentionPolicy "Default HSDG Retention Policy"
				$primary_smtp = Get-Mailbox -Identity $user | select primarySmtpAddress
			}
			Write-Host $primary_smtp
			Write-Host "Press any key to return to main menu ..."
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}

## DISTRIBUTION LIST CREATION
		'5' {
			cls
			$menutext
			$finish_info
			$dist_string
			$region = read-host "> REGION"
			$location = read-host "> LOCATION"
			$ou_full = "OU=DST,OU=Groups,OU=" + $location + ",OU=" + $region + ",OU=HSDG,DC=hsdg-ad,DC=int"
			$lists = @()
			do {
				$input = (Read-Host "> FULL NAME")
				if ($input -ne '') {$lists += $input}
			}
			until ($input -eq '')
			foreach ($list in $lists) {
				Write-Host ""
				Write-Host "CREATION OF THE LIST"$list
				Write-Host ""
				$alias_full = read-host "> ALIAS"
				$owner = read-host "> OWNER"
				$psmtp = $alias_full +"@hamburgsud.com"
				New-DistributionGroup -name $list -Alias $alias_full -DisplayName $alias_full -OrganizationalUnit $ou_full -ManagedBy $owner -MemberDepartRestriction Closed -MemberJoinRestriction Closed -PrimarySmtpAddress $psmtp | Out-Null
				Set-DistributionGroup -identity $list -RequireSenderAuthenticationEnabled $False
				Write-Host ""
				Write-Host "######## COPY FOR RESOLUTION ########"
				Write-Host "Created following list:"
				Write-Host `t$alias_full
				Write-Host "Email address"
				Write-Host `t$psmtp
				Write-Host "Please be aware that it can take up to 24 hours until the changes are visible in the Outlook address book."
				Write-Host "The distribution list is already able to receive emails."
				Write-Host "#####################################"
				Write-Host ""
				}
			Write-Host "Press any key to return to main menu ..."
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}

## DISABLE A USERS MAILBOX
		'DISABLE' {
			cls
			$menutext
			'You chose to disable an Exchange user mailbox.'
			$disable_string
			$user = read-host "> AFFECTED USER"
			Disable-Mailbox -Identity $user
			Write-Host "Press any key to return to main menu ..."
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}

## SHORT HELP
		'H' {
			cls
			$menutext
			'You chose "Help".'
			$help_string
			Write-Host "Press any key to return to main menu ..."
			$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}

		'q' {
			return
		}
	}

}
until ($input -eq 'q')
