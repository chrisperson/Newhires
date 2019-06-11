#Connect to Exchange and Skype Remote PS
$UserCredential = Get-Credential
$ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://emsmbxv01.seattle.main.gatesfoundation.org/PowerShell/ -Authentication Kerberos -Credential $UserCredential
$SfBSession = New-PSSession -ConnectionUri https://seas4bfe01.seattle.main.gatesfoundation.org/OCSPowerShell -Credential $UserCredential

#Check if connection to Exchange On-Premises is working
if (!($ExSession)) {
	Write-Host "No connection to exchange server, exiting script" -foregroundcolor RED
	start-sleep 10
	exit
	}
#Check if connection to SfB is working
if (!($SfBSession)) {
	Write-Host "No connection to exchange server, exiting script" -foregroundcolor RED
	start-sleep 10
	exit
}

import-pssession -session $ExSession -allowclobber
import-pssession -session $SfBSession -allowclobber

#Import Active Directory Module
Import-Module ActiveDirectory

#Create a quick check for each service

#Import Data
$UserImport = import-csv .\Newhires.csv

#Define Arrays
$UsersToProvision = @()
$UsersToSkip = @()
$ADUserList = @()
$AssignPhoneNumbers = @()
$NoEnterpriseVoice = @()

#DefineGlobalVariables
$EVListFile = (get-date -format yyyyMMddhhmmss) + "_AssignNumbers.csv"
$NoEVListFile = (get-date -format yyyyMMddhhmmss) + "_NoEnterpriseVoice.csv"
$SkipFile = (get-date -format yyyyMMddhhmmss) + "_SkippedUsers.csv"
$NewHireCopy = (get-date -format yyyyMMddhhmmss) + "_NewHires.csv"
$RoutingDomain = "@bmgf.mail.onmicrosoft.com"

#Create AD User list for each object in csv and appent Office Location from CSV.
#Insert UserImport variable saftey checks...

foreach ($item in $UserImport) {
	$User = get-aduser $item.SamAccountName -properties *
	$User | Add-Member -Force -MemberType NoteProperty -Name "UserOffice" -value $item.UserOffice
	$ADUserList += $User
}

#Check each user in the list for unique email address and sip address
foreach ($User in $ADUserList) {
	If ((!(get-recipient $User.UserPrincipalName -erroraction SilentlyContinue)) -and (!(get-csuser $User.UserPrincipalName -erroraction SilentlyContinue))) {
		$UsersToProvision += $User
	}
	Else {
		$UsersToSkip += $User
	}
}

#Consider making this a function - consider manual run unless specified in call

If ($UsersToProvision) {
	Write-Host "Woud you like to provision the following Mailboxes and Skype for Business Accounts" -foregroundcolor Yellow

	Foreach ($User in $UsersToProvision) {
	Write-Host $User.UserOffice "," $User.DisplayName "," $User.SamAccountName "," $User.UserPrincipalName -foregroundcolor Yellow
	}
	$Response = Read-Host "( Y / N )"
		Switch ($Response) {
			Y {Write-Host "Provisioning of Services Approved" -foregroundcolor Green ; $ReviewConfirmed=$true}
			N {Write-Host "Provisioning of Services Denied" -foregroundcolor RED ; $ReviewConfirmed=$false}
			Default {Write-Host "Invalid Response" -foregroundcolor RED ; $ReviewConfirmed=$false}
		}
	if ($ReviewConfirmed -ne $true) {
		write-host "Confirmation Denied, closing process" -foregroundcolor RED
		start-sleep 10
		exit
	}
}

# Enable Skype Accounts in Pool appropriate to office location
Function Enable-BMGFAccount {
	Foreach ($User in $UsersToProvision) {
		Switch ($User.UserOffice) {
			"LON" {
				$RegistrarPool = "LONS4BFE01.seattle.main.gatesfoundation.org"
				$RemoteRoutingAddress = $User.SamAccountName + $RoutingDomain
				enable-Remotemailbox $User.UserPrincipalName -remoteroutingaddress $RemoteRoutingAddress
				enable-csuser -identity  $User.UserPrincipalName -RegistrarPool $RegistrarPool -SipAddressType UserPrincipalName -SipDomain gatesfoundation.org
			}
			"BER" {
				$RegistrarPool = "LONS4BFE01.seattle.main.gatesfoundation.org"
				$RemoteRoutingAddress = $User.SamAccountName + $RoutingDomain
				enable-Remotemailbox $User.UserPrincipalName -remoteroutingaddress $RemoteRoutingAddress
				enable-csuser -identity  $User.UserPrincipalName -RegistrarPool $RegistrarPool -SipAddressType UserPrincipalName -SipDomain gatesfoundation.org
			}
			"BEJ" {
				$RegistrarPool = "BEJS4BFE01.seattle.main.gatesfoundation.org"
				$RemoteRoutingAddress = $User.SamAccountName + $RoutingDomain
				enable-Remotemailbox $User.UserPrincipalName -remoteroutingaddress $RemoteRoutingAddress
				enable-csuser -identity  $User.UserPrincipalName -RegistrarPool $RegistrarPool -SipAddressType UserPrincipalName -SipDomain gatesfoundation.org
			}
			"DEL" {
				$RegistrarPool = "DELS4BFE01.seattle.main.gatesfoundation.org"
				$RemoteRoutingAddress = $User.SamAccountName + $RoutingDomain
				enable-Remotemailbox $User.UserPrincipalName -remoteroutingaddress $RemoteRoutingAddress
				enable-csuser -identity  $User.UserPrincipalName -RegistrarPool $RegistrarPool -SipAddressType UserPrincipalName -SipDomain gatesfoundation.org
			}
			"PAT" {
				$RegistrarPool = "DELS4BFE01.seattle.main.gatesfoundation.org"
				$RemoteRoutingAddress = $User.SamAccountName + $RoutingDomain
				enable-Remotemailbox $User.UserPrincipalName -remoteroutingaddress $RemoteRoutingAddress
				enable-csuser -identity  $User.UserPrincipalName -RegistrarPool $RegistrarPool -SipAddressType UserPrincipalName -SipDomain gatesfoundation.org
			}
			"JNB" {
				$RegistrarPool = "JNBS4BFE01.seattle.main.gatesfoundation.org"
				$RemoteRoutingAddress = $User.SamAccountName + $RoutingDomain
				enable-Remotemailbox $User.UserPrincipalName -remoteroutingaddress $RemoteRoutingAddress
				enable-csuser -identity  $User.UserPrincipalName -RegistrarPool $RegistrarPool -SipAddressType UserPrincipalName -SipDomain gatesfoundation.org
			}	
			Default {
				$RegistrarPool = "SEAS4BPOOL01.gatesfoundation.org"
				$RemoteRoutingAddress = $User.SamAccountName + $RoutingDomain
				enable-Remotemailbox $User.UserPrincipalName -remoteroutingaddress $RemoteRoutingAddress
				enable-csuser -identity  $User.UserPrincipalName -RegistrarPool $RegistrarPool -SipAddressType UserPrincipalName -SipDomain gatesfoundation.org
			}
		}
	}
}


# Configure Pool/Office Specific Settings for Skype
Function Set-BMGFAccountPolicies {
	Foreach ($User in $UsersToProvision) {
		Switch ($User.UserOffice) {
			"SEA" {
				set-csuser $User.UserPrincipalName -EnterpriseVoiceEnabled $True
				grant-csdialplan $User.UserPrincipalName -policy "SeattleLync.seattle.main.gatesfoundation.org"
				grant-csvoicepolicy $User.UserPrincipalName -policy "LyncSeattleUsers - Unrestricted"
				grant-cshostedvoicemailpolicy $User.UserPrincipalName -policy "Office365UM"
			}
			"WDC" {
				set-csuser $User.UserPrincipalName -EnterpriseVoiceEnabled $True
				grant-csdialplan $User.UserPrincipalName -policy "ECOLync.seattle.main.gatesfoundation.org"
				grant-csvoicepolicy $User.UserPrincipalName -policy "LyncECOUsers - Unrestricted"
				grant-cshostedvoicemailpolicy $User.UserPrincipalName -policy "Office365UM"
			}
			"LON" {
				set-csuser $User.UserPrincipalName -EnterpriseVoiceEnabled $True
				Grant-CsDialPlan $User.UserPrincipalName -PolicyName "LondonLync.seattle.main.gatesfoundation.org"
				Grant-CsVoicePolicy $User.UserPrincipalName -PolicyName "Lync UK Users - Unrestricted"
				grant-cshostedvoicemailpolicy $User.UserPrincipalName -policy "Office365UM"
			}			
			"BEJ" {
				set-csuser $User.UserPrincipalName -EnterpriseVoiceEnabled $True
				Grant-CsDialPlan $User.UserPrincipalName -PolicyName "BeijingLync.seattle.main.gatesfoundation.org"
				Grant-CsVoicePolicy $User.UserPrincipalName -PolicyName "Lync China Users - Unrestricted"
				grant-cshostedvoicemailpolicy $User.UserPrincipalName -policy "Office365UM"
			}
			"DEL" {
				set-csuser $User.UserPrincipalName -EnterpriseVoiceEnabled $True
				grant-csdialplan $User.UserPrincipalName "NewDelhiLync.seattle.main.gatesfoundation.org"
				grant-csvoicepolicy $User.UserPrincipalName "IN-NewDelhi-International"
				grant-cshostedvoicemailpolicy $User.UserPrincipalName -policy "Office365UM"
			}
			"JNB" {
				set-csuser $User.UserPrincipalName -EnterpriseVoiceEnabled $True
			}	
			Default {
			}
		}
	}
}

Enable-BMGFAccount
### Sleep for 15 seconds for Skype/AD Replication
start-sleep -s 15
#Configure Policies for New Accounts
Set-BMGFAccountPolicies

If ($UsersToSkip) {
	Write-Host "The following users have been skipped due to a non unique SMTP Address or SIP Address" -foregroundcolor Yellow
	Foreach ($User in $UsersToSkip) {
	Write-Host $User.UserOffice "," $User.DisplayName "," $User.SamAccountName "," $User.UserPrincipalName -foregroundcolor Yellow
	}
}
start-sleep 15
foreach ($User in $UsersToProvision) {
$UserData = get-csuser $User.UserPrincipalName
if ($UserData.EnterpriseVoiceEnabled -ne $True) {
	$NoEnterpriseVoice += $UserData
	}
else {
	$AssignPhoneNumbers += $UserData
	}
}

$UsersToSkip | Select-Object UserOffice, DisplayName, SamAccountName,UserPrincipalName | export-csv .\logs\$SkipFile
Write-Host "Please enable the following users for Exchange UM, Intercall, and assign a Phone Number." -foregroundcolor Green
$AssignPhoneNumbers | Select-Object DisplayName, SipAddress, RegistrarPool
$AssignPhoneNumbers | Select-Object DisplayName, SipAddress, RegistrarPool | export-csv .\logs\$EVListFile
Write-Host "Please enable the following users for Intercall." -foregroundcolor Yellow
$NoEnterpriseVoice | Select-Object DisplayName, SipAddress, RegistrarPool 
$NoEnterpriseVoice | Select-Object DisplayName, SipAddress, RegistrarPool | export-csv .\logs\$NoEVListFile
$UserImport | export-csv .\logs\$NewHireCopy

#Clear NewHires.csv
(Get-Content .\NewHires.csv |  Select-Object -First 1) | Out-File .\NewHires.CSV
#Submit Zendesk Ticket for Newhire/Phone Number Assignments.