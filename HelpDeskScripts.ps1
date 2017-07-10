#------------------------------------------------------------------------
# Title: Service Desk Console Functions
# Author: Malcolm Sims (VVRB)
# Notes: Production Module
# Notes: 
#Sensitive Company information removed from module
#------------------------------------------------------------------------##

####==========================================================================================
#======================================Email Confirmations
####==========================================================================================
	<#
.Synopsis
   Email Confirmation for Service Desk Console
.DESCRIPTION
   This command is used to send a email confimation including command performed and closing script to Service Desk Technician..
.EXAMPLE
   Send-ConfirmationEmail -DistributionEmail
#>
function Send-ConfirmationEmail
{
    [CmdletBinding()]
    Param
    (
        #SMTP Server
        [string]
        $SMTPServer = "Enter Mail Server",
        #Specifies Distribution List Email
        [switch]$DistributionList,
        #Specifies Mail Contact Email
        [switch]$MailContact,
        #ResourceMailbox Email
        [switch]$ResourceMailbox,
        #Termination Email
        [switch]$UserTermination,
        #Leave of Absence Email
        [switch]$LeaveofAbsence,
        #Return from Leave of Absence
        [switch]$ReturnFromLeaveofAbsence,
        #New User Creation
        [switch]$NewUser
    )

    Begin
    {
		#Creation of variables that are common throughout this command.

		#Email of technician
        $TechEmail = (Get-ADUser $env:USERNAME -Properties emailaddress).emailaddress

		#Subject of Notification including Case number
        $EmailSubject = "Console Notification - Case#: $CaseID"

		#Subject without Case number
		$EmailSubject2 = "Console Notification"

		#Email address used for Console monitoring
        $ConsoleEmail = 'MonitoringBox@Company.com'

		#Group Mailbox for Service Desk Team
		$TechMailbox = 'GroupITMailbox@Company.com'

		#Tech's First name
		$TFirstName = (Get-ADUser $env:username).GivenName

		#Tech's Last name
        $TLastName = (Get-ADUser $env:username).Surname
    }
    Process
    {

    if ($DistributionList)
    {
        $Name = "$DisplayName" 
		$OwnerFirstName = (Get-ADUser $Owner).GivenName
		$OwnerLastName = (Get-ADUser $Owner).Surname

        $EmailBody = "This message is to confirm that a distribution group has been successfully created.
		
Closing Script:

Hi customername, we have created the $Name distribution list as requested. We have listed $OwnerFirstName $OwnerLastName as the manager of the list. Thank you, $TFirstName $TLastName, IT Service Desk."

    Send-MailMessage -To $TechEmail -Bcc $ConsoleEmail -Subject $EmailSubject -Body $EmailBody -From $ConsoleEmail -SmtpServer $SMTPServer
       
    }  #if 
    if ($MailContact)
    {
         $Name = $DisplayName

        $EmailBody = "This message is to confirm that a Mail Contact has been successfully created
		
Name: $Alias
EmailAddress: $ExternalEmail
Case#: $CaseID
		
Closing Script: 
		
Hello customername, a Contact has been created for $Alias.  The e-mail address is $ExternalEmail. Please note that any e-mail sent to the e-mail address before forwarding was set up, will not forward to the new contact address.  It can be accessed by signing into the Webmail service.  Thank you, $TFirstName $TLastName, IT Service Desk."

    Send-MailMessage -To $TechEmail -Bcc $ConsoleEmail -Subject $EmailSubject -Body $EmailBody -From $ConsoleEmail -SmtpServer $SMTPServer

    }  #if
    if ($ResourceMailbox)
    {
        $Name = (Get-Mailbox $LogonName).DisplayName

		$MailboxEmail = (Get-Mailbox $LogonName).PrimarySMTPaddress

        $EmailBody = "The mailbox [$Name] has been successfully created. The owners are Owner: $NoteOwner$Note2Owner & Co-Owner: $NoteOwner2$Note2Owner2.
		
Name: $Name
Email Address: $MailboxEmail
Owner: $NoteOwner $Note2Owner
Co-Owner: $NoteOwner2 $Note2Owner2
Alt Owner: $NoteOwner3 $Note3Owner3
		
Closing Script:
		
Hi customername, we have completed your request and have created the $Name mailbox, with an email address of $MailboxEmail. Thank you, $TFirstName $TLastName, IT Service Desk"

    Send-MailMessage -To $TechEmail -Bcc $ConsoleEmail -Subject "$EmailSubject - Mailbox Created" -Body $EmailBody -From $ConsoleEmail -SmtpServer $SMTPServer

    } #if
    if ($UserTermination)
    {
        $Name = (Get-ADUser $Username).Name

		$EmailDate = Get-Date -Format d

        $EmailBody = "This message is to inform you that a user termination has started. The Pst for the mailbox is currently being exported. Once the export is finished the mailbox will be disabled automatically.
		
Name: $Name
Username: $Username
Date Started: $EmailDate
		
Closing Script:
		
Hello customername, we were delighted to assist you.  The network account for $Name has been disabled.  Their email address has been removed from the GAL, and they have been removed from all memberships and access. Please contact the IT Service Desk at 999-999-9999 opt 1 if you have any questions or we can help further.  I will close the case.  Thank you, $TFirstName $TLastName, IT Service Desk."

    Send-MailMessage -To $TechEmail -Bcc $ConsoleEmail -Subject $EmailSubject2 -Body $EmailBody -From $ConsoleEmail -SmtpServer $SMTPServer

    } #if
    if ($LeaveofAbsence)
    {
       if ((Get-ADUser $Username -Properties manager).manager -ne $null) {
		   
	    $FirstName = (Get-ADUser $Username).GivenName
		$LastName = (Get-ADUser $Username).SurName
		$ManagerFirstName = ((Get-ADUser $Username -Properties manager).manager | get-aduser).GivenName
		$ManagerLastName = ((Get-ADUser $Username -Properties manager).manager | get-aduser).SurName
		$ManagerEmail = ((get-aduser $Username -Properties manager).manager | get-aduser -Properties emailaddress).emailaddress 

		$EmailBody = "This message is to inform you that [$FirstName $LastName] has been placed on leave and his/her network access has been disabled. Be advised a copy of this message has been sent to [$ManagerFirstName $ManagerLastName] the manager of [$FirstName $LastName] as well."
		$EmailBody2 = "This message is to inform you that [$FirstName $LastName] has been placed on leave and his/her network access has been disabled. Be advised a copy of this message has been sent to [$ManagerFirstName $ManagerLastName] the manager of [$FirstName $LastName] as well.
		
Hi customername, we were delighted to assist you. We have placed $FirstName $LastName on LOA per your request. If you have any questions, please contact me at (999) 999-9999.  Thank you, $TFirstName $TLastName, IT Service Desk.
		"
    Send-MailMessage -To $ManagerEmail -Bcc $ConsoleEmail -Subject $EmailSubject -Body $EmailBody -From $TechMailbox -SmtpServer $SMTPServer
	Send-MailMessage -To $TechEmail -Subject $EmailSubject -Body $EmailBody2 -From $ConsoleEmail -SmtpServer $SMTPServer
  			 } else {
		$FirstName = (Get-ADUser $Username).GivenName
		$LastName = (Get-ADUser $Username).SurName

		$EmailBody = "This message is to inform you that [$FirstName $LastName] has been placed on leave and his/her network access has been disabled. Be advised a copy of this message was not sent to a manager as there was not one listed on the account.
		
Hi customername, we were delighted to assist you. We have placed $FirstName $LastName on LOA per your request. If you have any questions, please contact me at (866) 984-3571.  Thank you, $TFirstName $TLastName, IT Service Desk.
		"
    Send-MailMessage -To $TechEmail -Bcc $ConsoleEmail -Subject $EmailSubject -Body $EmailBody -From $TechMailbox -SmtpServer $SMTPServer
	  } #else
        
    }
    if ($ReturnFromLeaveofAbsence)
    {
		if ((Get-ADUser $Username -Properties manager).manager -ne $null) {

		$Name = (Get-ADUser $Username).Name
        $FirstName = (Get-ADUser $Username).GivenName
		$LastName = (Get-ADUser $Username).SurName
		$ManagerFirstName = ((Get-ADUser $Username -Properties manager).manager | get-aduser).GivenName
		$ManagerLastName = ((Get-ADUser $Username -Properties manager).manager | get-aduser).SurName
		$ManagerEmail = ((get-aduser $Username -Properties manager).manager | get-aduser -Properties emailaddress).emailaddress
		$EmailSubject2 = "Return From Leave of Absence $Name"

        $EmailBody = "This message is to inform you that [$FirstName $LastName] has returned from LOA. The password has been set to $NewPassword. Be advised that a copy of this message has been sent to [$ManagerFirstName $ManagerLastName] the manager listed on the account."
		$EmailBody2 = "This message is to inform you that [$FirstName $LastName] has returned from LOA. The password has been set to $NewPassword. Be advised that a copy of this message has been sent to [$ManagerFirstName $ManagerLastName] the manager listed on the account.
		
Hi customername, we enabled the account for $FirstName $LastName and notified their manager, $ManagerFirstName $ManagerLastName accordingly.  Thanks, $TFirstName $TLastName, IT Service Desk.
		"
    Send-MailMessage -To $ManagerEmail -Subject $EmailSubject2 -Cc $ManagerEmail -Body $EmailBody -From $ConsoleEmail -SmtpServer $SMTPServer
	Send-MailMessage -To $TechEmail -Bcc $ConsoleEmail -Subject $EmailSubject2 -Body $EmailBody2 -From $ConsoleEmail -SmtpServer $SMTPServer	
	} else {
			$Name = (Get-ADUser $Username).Name
        $FirstName = (Get-ADUser $Username).GivenName
		$LastName = (Get-ADUser $Username).SurName
		$NewUEmail = (Get-ADUser $Username).EmailAddress
		$EmailSubject2 = "Return From Leave of Absence $Name"

        $EmailBody = "This message is to inform you that [$FirstName $LastName] has returned from LOA. The password has been set to $NewPassword. Be advised that a copy of this email has not been sent to a manager as there was not one listed.
		
Hi customername, we enabled the account for $FirstName $LastName and notified their manager, <managername> accordingly.  Thanks, $TFirstName $TLastName, IT Service Desk.
		"
    Send-MailMessage -To $TechEmail -Bcc $ConsoleEmail -Subject $EmailSubject -Body $EmailBody -From $ConsoleEmail -SmtpServer $SMTPServer
		} #else
    }
    if ($NewUser)
    {
        $Name = "$FirstName $LastName"

        $EmailBody = "This message is to inform you that a network account has been created.

Name: $Name
Start Date: $StartDate
Case#: $CaseID

Closing Script:

Hello customername, a  network account has been created for $Name.  The Network Login is $Username and the email address is $NewUEmail.  Thank you, $TFirstName $TLastName, IT Service Desk"


    Send-MailMessage -To $TechEmail -Bcc $ConsoleEmail -Subject $EmailSubject -Body $EmailBody -From $ConsoleEmail -SmtpServer $SMTPServer
} #if
    }
    End
    {}
}#function

	#================================================================================================================================
	#==============================================Unlock User (Menu Modify)===========================================================================
	#================================================================================================================================
	
	
	function Unlock-User
	{
		[CmdletBinding()]
		[OutputType([string])]
		Param
		(
			#Specifies an Active Directory user object to be unlocked.
			[Parameter(Mandatory = $true,
					   ValueFromPipelineByPropertyName = $true)]
			[string]$SamAccountName
		)
		
		Begin
		{
		}
		Process
		{
			foreach ($Username in $SamAccountName) {
					Unlock-ADAccount $Username
			} #foreach
		}
		End
		{}
	}

    #================================================================================================##
    #===================================Lockout Status===============================================##
    #================================================================================================##

    function Get-ADLockoutStatus {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
		#SamAccountName of the Active Directory user.
       [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
       [string]
       $SamAccountName

    )

    Begin
    {
    }

    Process
    {

    foreach ($Username in $SamAccountName)
        {

                Write-Verbose 'Querying Domain Controller'
                Get-ADUser $Username -Properties DisplayName,PasswordLastSet,LockedOut,LockoutTime,LastBadPasswordAttempt,badPwdCount |
                Select-Object DisplayName,PasswordLastSet,LockedOut,LockoutTime,LastBadPasswordAttempt,badPwdCount

        }
    }
    
    End
    {}
}
	
	#================================================================================================================================
	#==============================================US Distribution Functions===========================================================================
	#================================================================================================================================
	
<#
.Synopsis
  Creates a New Distribution List
.DESCRIPTION
  Creates a new Distribution list. This cmdlet will create the Distribution list and also fill out the AD properties per Company requirements. 
  policy. 
.EXAMPLE
   New-DistributionList -DisplayName US Bob List -CaseID 1231231 -Owner A01testuser -Alias USBobList -Purpose 'Bob Account Communication'
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Creates Distribution list the can be seen in EAC and AD.
.FUNCTIONALITY
   Requires that a connection to EAC be made.
#>
	
	function New-DistributionList
	{
		[CmdletBinding()]
		param (
			#Display Name of Distribution List
			[parameter(Mandatory = $True)]
			[string]$DisplayName,
			#Alias.
			[parameter(Mandatory = $True)]
			[string]$Alias,
			[Parameter(Mandatory=$true)]
			#Case Number per case
			[int]$CaseID,
			#Owner of the Distribution List - must be SamAccountName
			[parameter(Mandatory = $true)]
			[string]$Owner,
			#Purpose for Distribution List
			[parameter (Mandatory = $true)]
			[string]$Purpose,
			#Switch for Company ME Group
			[switch]$CompanyME,
			#Invokes Pester Integration tests
			[switch]$Test
		)
		
		BEGIN
		{}
		PROCESS
		{
			#Variable for technicians initials
			$Tech = $env:username
			
			$Date = Get-Date -UFormat "%m/%d/%y"
			
			$Desc = "Created per Case# $CaseID - $Date - $Tech"
			
			#Organizational Unit of Distribution List
			$OUnit = 'Company.local/IT/Shared/DistributionLists'

			if ($CompanyME)	{
				#Hash Table for M&E Distro Variables
			$UProp = @{
				Name = "Company ME: $DisplayName";
				Alias = "Company_ME_$Alias";
				OrganizationalUnit = "$OUnit";
				DisplayName = "Company ME: $DisplayName";
				MemberJoinRestriction = "Closed";
				PrimarySmtpAddress = "Company_ME_$Alias@Companyme.com";
				Notes = "$Desc";
				Managedby = "$Owner"
			} #HashCompanyME
			} else {
				#Hash Table for Distro Variables
			$UProp = @{
				Name = "DL: $DisplayName";
				Alias = "DL_$Alias";
				OrganizationalUnit = "$OUnit";
				DisplayName = "DL: $DisplayName";
				MemberJoinRestriction = "Closed";
				PrimarySmtpAddress = "DL_$Alias@Company.com";
				Notes = "$Desc";
				Managedby = "$Owner"
			} #Hash
			} #else 
			
			New-DistributionGroup @UProp

			if ($Test) {

				if ($CompanyME) {
					Start-Sleep -Seconds 2
				Describe "Integration Tests: $DisplayName" {
					BeforeAll {
						$CDistro = Get-DistributionGroup "Company_ME_$Alias"
					}#BeforeAll
					Context "Exchange Online Tests" {
						It "M&E Name is Correct" {
							$CDistro.Name | Should Be $UProp.Name
						}#it
						It "M&E Alias is correct" {
							$CDistro.Alias | Should Be $UProp.Alias
						}#it
						It "M&E OU is correct" {
							$CDistro.OrganizationalUnit | Should Be $UProp.OrganizationalUnit
						}#it
						It "M&E Display Name is correct" {
							$CDistro.DisplayName | Should Be $UProp.DisplayName
						}#it
						It "M&E Member Join Restriction Closed" {
							$CDistro.MemberJoinRestriction | Should Be $UProp.MemberJoinRestriction
						}#it
						It "M&E Primary SMTP is correct" {
							$CDistro.PrimarySMTPaddress | Should Be $UProp.PrimarySMTPaddress
						}#it
					}#Context
				}#Describe
				} else {
				Start-Sleep -Seconds 2
				Describe "Integration Tests: $DisplayName" {
					BeforeAll {
						$CDistro = Get-DistributionGroup "DL_$Alias"
					}#BeforeAll
					Context "Exchange Online Tests" {
						It "Name is Correct" {
							$CDistro.Name | Should Be $UProp.Name
						}#it
						It "Alias is correct" {
							$CDistro.Alias | Should Be $UProp.Alias
						}#it
						It "OU is correct" {
							$CDistro.OrganizationalUnit | Should Be $UProp.OrganizationalUnit
						}#it
						It "Display Name is correct" {
							$CDistro.DisplayName | Should Be $UProp.DisplayName
						}#it
						It "Member Join Restriction Closed" {
							$CDistro.MemberJoinRestriction | Should Be $UProp.MemberJoinRestriction
						}#it
						It "Primary SMTP is correct" {
							$CDistro.PrimarySMTPaddress | Should Be $UProp.PrimarySMTPaddress
						}#it
					}#Context
				}#Describe
			}#else
		}#if
		}
		END
		{
			Send-ConfirmationEmail -DistributionList
		}#end
	}
	
				
	#================================================================================================================================
	#============================================== User Termination (Menu Modify)===========================================================================
	#================================================================================================================================
	
<#
.Synopsis
  Termination of Company User AD Account.
.DESCRIPTION
  Populates AD properties per the Service Desk checklist, disables the AD account, removes manager field and resets users password.
  This cmdlet will also move the AD account to the 'Disabled Accounts' OU and remove all memberships to AD groups and Distribution lists.
.EXAMPLE
Shows the user being removed including the Active Directory notes being added and removal of all associated groups. The test switch is also used to run integration tests via Pester. 
  PS C:\windows> Remove-CompanyUser -Identity bbuilder -CaseID 1234567 -TermPassword 'Te1231231' -Test

Describing Integration Tests

  Context Active Directory Tests
    [+] Confirms Description Entered 3.09s
    [+] Confirms Note section edited 30ms
    [+] Account Disabled 101ms
    [+] Unable to Change Password 32ms
    [+] Confirms Manager removed 51ms
    [+] Confirms Memberships & Groups Removed 60ms

  Context Exchange Online Tests
    [+] Placed in Disabled Users OU 101ms
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   1st part of the Service Desk termination policy.
.FUNCTIONALITY
   1st part of the Company termination procedure. 2nd part is found in the Remove-CompanyUserMailbox cmdlet.
#>
	
	function Remove-CompanyUser
	{
		[CmdletBinding()]
		param (
			#Username of Terminated User
			[parameter(Mandatory = $True)]
			[string]$Username,
			#Case number.
			[parameter(Mandatory = $True)]
			[int]$CaseID,
			#Term User Pass
			[parameter(Mandatory = $true)]
		[string]$TermPassword,
		[switch]$Test
		)
		
		Begin
		{
		}
		
		Process
		{
			#Main Nasa Domain Controller
			
			$TermDate = Get-Date -UFormat "%m/%d/%y"
			
			$TermProperties = @{
                                identity = $Username;
                                Enabled = $false;
                                CannotChangePassword = $True;
                                Manager = $null;
                                                }
			
			#Variable Hash Table
			$Entry = @"

Disabled $TermDate Case# $CaseID - $env:username
"@
			
			$Entry2 = "Disabled $TermDate per Case# $CaseID - $env:username"
			
			$S = Get-ADUser $Username -Properties info
			
			#Setting of users Description and also the info property.Also applies properties in $TermProperties Hash table.
			Set-ADUser @TermProperties -Description $Entry2 -Replace @{ info = $S.info + $Entry }
			
			#Resets user's password
			Set-ADAccountPassword $Username -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $TermPassword -Force)
			
			#Moves ad acocunt to the disabled users OU.
			(Get-ADUser $Username).objectguid | Move-ADObject -TargetPath "OU=Users,OU=Terminated,DC=Company,DC=local"
			
			#Removes user from all AD groups except for the Domain Local group.
		    (Get-ADUser $Username -Properties memberof).memberof | Remove-ADGroupMember -Members "$Username" -Confirm:$false
		
		If ($Test)
		{
            Start-Sleep -Seconds 2

			Describe "Integration Tests" {
				BeforeAll {
					$AdServer = 'DoaminController.Company.local'
					$CUser = Get-ADUser $Username -Server $AdServer -Properties *
					$CMailbox = Get-Mailbox $Username
				} #BeforeAll
				Context "Active Directory Tests" {
					It "Confirms Description Entered" {
						$CUser.Description | Should Be $Entry2
					} #it
					It "Confirms Note section edited" {
						$CUser.info | Should Match $Entry
					} #it
					It "Account Disabled" {
						$CUser.Enabled | Should Be $true
					} #it
					It "Unable to Change Password" {
						$CUser.CannotChangePassword | Should Be $true
					} #it
					It "Confirms Manager removed" {
						$CUser.Manager | Should BeNullOrEmpty
					} #it
					It "Confirms Memberships & Groups Removed" {
					$CUser.Memberof | Should BeNullOrEmpty
                    } #it
				} #Context
				Context "Exchange Online Tests" {
					It "Placed in Disabled Users OU" {
						$CMailbox.OrganizationalUnit | Should BeExactly 'Company.local/Terminated/Users'
					} #it
				} #Context
			}
		} #Switch
		} #Process
		End
		{
		}
	}
	
<#
.Synopsis
  Removes user mailbox per the Service Desk termination procedures.
.DESCRIPTION
   This cmdlet will hide a users mailbox from the GAL, remove any forwarding setup and que the mailbox for export to pst in the pst
   shared folder. Running this cmdlet in the Service Desk Console gives the option to have the mailbox added to a list for mailbox 
   removal from the Exchange database. This option in the console will remove the Exchange properties in AD. This has not been added 
   to this cmdlet, but can be ran by using the below command. 
   Add to Disable List: (Get-Mailbox 'Enter Username').Alias | Out-File '\\usmokscr300d\c$\MailboxestoRemove.csv' -Append
.EXAMPLE
   Remove-CompanyUserMailbox -Username A01testuser
.EXAMPLE
   Remove-CompanyUserMailbox -Username A01testuser

   (Get-Mailbox 'A01testuser).Alias | Out-File '\\usmokscr300d\c$\MailboxestoRemove.csv' -Append
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   2nd part of the Company termination procedure.
.COMPONENT
   The component this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
	
	function Remove-CompanyUserMailbox
	{
		[CmdletBinding()]
		param (
			#Username of Terminated User
			[parameter(Mandatory = $True)]
		[string]$Username,
		#Invokes Pester Integration Tests
		[switch]$Test
		)
		
		Begin
		{
		}
		
		Process
		{
			$ExpDate = Get-Date -UFormat '%m%d%y'

            $RemoveProperties = @{
                                    identity = $Username;
                                    ForwardingAddress = $null;
                                    HiddenFromAddressListsEnabled = $true;
                                    }#Hash
			
			#Hides the email address from the GAL and removes all forwarding.
			Set-Mailbox @RemoveProperties
			
			#Adds a new request to export the users mailbox.
			New-MailboxExportRequest -Mailbox $Username -FilePath "\\FileServer\PSTEXPORT\$ExpDate-$Username.pst"

			#Adds Username to text list for the automated Termination 
			(Get-Mailbox $Username).alias | Out-File '\\SourceServer\C$\Console\Production\Text Lists\MailboxesToRemove.txt' -Append
		
			Send-ConfirmationEmail -UserTermination
		
		if ($Test)
		{
			Describe "Integration Tests" {
				BeforeAll {
					$ExportStatus = (Get-MailboxExportRequest -Mailbox $id).Status
				} #BeforeAll
				Context "Exchange Online Tests" {
					It "Hidden From Gal" {
						$CUserMailbox.HiddenFromAddressListsEnabled | Should Be $RemoveProperties.HiddenFromAddressListsEnabled
					} #it
					It "Email Forwarding Removed" {
						$CUserMailbox.ForwardingAddress | Should BeNullOrEmpty
					} #it
					It "Mailbox Export Started" {
						$ExportStatus | Should Be 'Queued'
					} #it
					It "Added to Removal List" {
						'\\SourceServer\C$\Console\Production\Text Lists\MailboxesToRemove.txt' | Should Contain $Username
					} #it
				} #Context
			} #Describe
		} #if
	} #Process
	End
		{
		}
	}

###===============================================================================================================####
#=================================================Leave of Absence Functions=========================================#
#####===========================================================================================================######

<#
.Synopsis
  Puts user on Leave of Absence status per Helpdesk Guidelines.
.DESCRIPTION
   Sets AD properties and disables account for the Leave of Absence status. Will move account to the Leave of Absence Users OU.
.EXAMPLE
   Invoke-UserLOA -Username bbuilder -CaseID '1234567' -NewPassword 'Bb1234567'
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Puts a user on LOA status
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
	
	function Invoke-UserLOA
	{
		[CmdletBinding()]
		param
		(
			#Username Parameter
			[parameter(Mandatory = $true)]
			[string]$Username,
			#CaseID Parameter
			[Parameter(Mandatory = $true)]
			[int]$CaseID,
			#Password Parameter
			[Parameter(Mandatory = $true)]
			[string]$NewPassword,
			#Invokes Pester integration tests
			[switch]$Test
		)
		
		BEGIN
		{}
		PROCESS
		{
			#Main Nasa Domain Controller
			$NasaDCServer = "Company.DomainController.local"

			$Date = Get-Date -UFormat "%m/%d/%y"
			
			#Technician Initials
			$Tech = $env:username
			
			#Hash Table
			$UProp = @{
				identity = "$Username"
			}
			
			#Property for notes section
			$Entry = @"

LOA $Date Case# $CaseID - $Tech
"@
			#Variable for Notes section
			$CurrInfo = Get-ADUser @UProp -Properties info
			
			$Description = "LOA - Do not reset password $Date - $Tech"
			
			#Setting of user properties including Description and info fields
			Set-ADUser @UProp -Description $Description -Replace @{ info = $CurrInfo.info + $Entry }
			
			#Change of password
			Set-ADAccountPassword $Username -Reset -NewPassword (Convertto-SecureString -AsPlainText "$NewPassword" -Force)
			
			#Movement of the Active Directory account to the Leave of Absence OU
			(Get-ADUser $Username).objectguid | Move-ADObject -TargetPath "OU=LOA,OU=Disabled,DC=Company,DC=local"

			if ($Test) {

				Describe "Integration Test: LOA $Username" {
					BeforeAll {
						$CLOA = Get-ADUser $Username -Server $NasaDCServer -Properties *
					}#BeforeAll
					Context "Active Directory Tests" {
						It "Description is set" {
							$CLOA.Descripton | Should Be $Description
						}#it
						It "Moved to LOA Ou" {
							$CLOA.OrganizationalUnit | Should Match "OU=LOA,OU=Disabled,DC=Company,DC=local"
						}#it
					}#Context
				}#Describe
			}#if "Test"
		}
		END
		{
			Send-ConfirmationEmail -LeaveofAbsence
		}
	} #function
	
	<#
.Synopsis
   Returns user from Leave of Absence
.DESCRIPTION
   Sets AD properties and enables account for return from LOA.
.EXAMPLE
   Invoke-UserLOAReturn -Username bbuilder -CaseID '1234567' -NewPassword 'Bb1234567'
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Returns a user from Loa, but you will still need to move them to the correct OU
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
	
	function Invoke-UserLOAReturn
	{
		[CmdletBinding()]
		param
		(
			#SamAccountName of user
			[parameter(Mandatory = $true)]
			[string]$Username,
			#Case Number
			[Parameter(Mandatory = $true)]
			[int]$CaseID,
			#Password 
			[Parameter(Mandatory = $true)]
			[string]$NewPassword
		)
		BEGIN
		{}
		PROCESS
		{
			$Date = Get-Date -UFormat "%m/%d/%y"
			
			#Technician Initials
			$Tech = $env:username
			
			#Variable Hash Table
			$UProp = @{
				identity = "$Username"
			}
			
			#Hash for notes section
			$Entry = @"

Return from LOA $Date Case# $CaseID - $Tech 
"@
			#Variable used for placing in the correct Organizational Unit
			$Office = (Get-ADUser $Username -Properties office).Office

			#Variable to get existing notes section
			$CurrInfo = Get-ADUser @UProp -Properties info

			#Setting of user properties
			Set-ADUser @UProp -Description $null -Replace @{ info = $CurrInfo.info + $Entry }

			#Enabling of user account
			Enable-ADAccount @UProp
			
			#Changing of network password
			Set-ADAccountPassword @UProp -Reset -NewPassword (Convertto-SecureString -AsPlainText "$NewPassword" -Force)

						#If statement to place Active Directory account in the correct Organizational Unit
                        if ($Office)
                        {
                        $SearchPath = (Get-ADOrganizationalUnit -Filter {Name -eq $Office}).DistinguishedName

                        $UserOuPath = (Get-ADOrganizationalUnit -Filter {Name -eq 'Users'} -SearchBase $SearchPath).DistinguishedName

                        (Get-ADUser $Username).objectguid | Move-ADObject -TargetPath $UserOuPath
                        } #if
		}
		
		END
		{
			Send-ConfirmationEmail -ReturnFromLeaveofAbsence
		}
	}


#====================================================================================================================================================================
####===============================================================Resource Mailbox Functions=========================================================
#====================================================================================================================================================================


<#
.Synopsis
   Creates Resource Mailbox in Nasa domain
.DESCRIPTION
   Resource Mailbox is created in the Nasa domain that can be seen in EAC and also Active Directory
.EXAMPLE
   New-ResourceMailbox -DisplayName 'Malc Test MB' -LogonName MalcTestMB -NewPassword Mt1231231 -Owner msims -Owner2 bbuilder
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Will create a Resource mailbox.
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
	function New-ResourceMailbox
	{
		[CmdletBinding()]
		param (
			#Logon name of mailbox
			[parameter(Mandatory = $True)]
			[string]$LogonName,
			#First name of mailbox.
			[parameter(Mandatory = $True)]
			[string]$DisplayName,
			#First Owner of mailbox
			[parameter(Mandatory = $true)]
			[string]$Owner,
			#Password for Resource Mailbox
			[string]$NewPassword,
			#Invokes Pester integration tests
            [switch]$Test
		)
		
		BEGIN
		{}
		PROCESS
		{
			$UsedPassword = ConvertTo-SecureString -AsPlainText $NewPassword -Force
			
			#OU that the user will be added to
			$OUnit = 'Company.Companytravel.local/IT/Shared/Mailboxes'
			
			#Hash Table for the Resource Mailbox
			$UProp = @{
				Name = "_grp_$DisplayName"
				Alias = "$LogonName";
				OrganizationalUnit = "$OUnit";
				UserPrincipalName = "$LogonName@Company.com";
				SamAccountName = "$LogonName";
				DisplayName = "$DisplayName";
				ResetPasswordOnNextLogon = $False;
				Password = $UsedPassword
			}

			$MailboxProperties = @{
				identity = $LogonName;
				EmailAddressPolicyEnabled = $False;
				EmailAddresses = "$LogonName@Company.com"
			}
			
			#Command to create the Resource mailbox
			New-Mailbox @UProp -Shared
            
            #Applies the parameters defined in $MailboxProperties Hashtable
            Set-Mailbox @MailboxProperties
			
			#Variable for adding owner permission
			$GrpName = (Get-Mailbox $LogonName).Name
				
				#creates array to fo owners to be added
				$collection = "$Owner"
				
				foreach ($GrpUser in $collection)
				{
					Add-MailboxPermission $LogonName -User $GrpUser -AccessRights FullAccess -InheritanceType All -ErrorAction SilentlyContinue
					
					Add-ADPermission "$GrpName" -User $GrpUser -AccessRights ExtendedRight -ExtendedRights "Send As" -ErrorAction SilentlyContinue
				} #foreach

            if ($Test) {
                Describe "Integration Tests" {
                    BeforeAll {
                        	$CMailbox = Get-Mailbox $LogonName
                    }#BeforeAll
                    Context "Exchange Online" {
                        It "Name is Correct" {
                        	$CMailbox.Name | Should Be $UProp.Name
                        }#it
                        It "Alias is Correct" {
                        	$CMailbox.Alias | Should Be $UProp.Alias
                        }#it
                        It "Correct OU" {
                        	$CMailbox.OrganizationalUnit | Should Be $UProp.OrganizationalUnit
                        }#it
                        It "UPN is Correct" { 
                        	$CMailbox.UserPrincipalName | Should Be $UProp.UserPrincipalName
						}#it
						It "SamAccountName is Correct" {
							$CMailbox.SamAccountName | Should Be $UProp.SamAccountName
						}#it
						It "Display Name is Correct"  {
							$CMailbox.DisplayName | Should Be $UProp.DisplayName
						}#it
						It "Disabled Password Reset On Logon" {
							$CMailbox.ResetPasswordOnNextLogon | Should Be $UProp.ResetPasswordOnNextLogon
						}#it
						It "Email Address Policy Disabled" {
							$CMailbox.EmailAddressPolicyEnabled | Should Be $MailboxProperties.EmailAddressPolicyEnabled
						}#it
						It "Correct Email Address Present" {
							$CMailbox.EmailAddresses | Should Match $UProp.EmailAddresses
						}#it
						It "Custom Attribute Present" {
							$CMailbox.CustomAttribute11 | Should Be $MailboxProperties.CustomAttribute11
						}#it
                    }#Context
                }#Describe
            }
			}
		END
		{}
		
	}
	
	<#
.Synopsis
   Sets AD properties per Service Desk checklists
.DESCRIPTION
   Long description
.EXAMPLE
   Set-ResourceMailbox -LogonName MalcTestMB -CaseID 1231231 -Purpose 'A New Testbx' -Owner 'msims' -Owner2 'bbuilder'
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   After creating a Resource mailbox this cmdlet can be ran to give it the default properties
#>
	
	function Set-ResourceMailbox
	{
		[CmdletBinding()]
		param (
			#Username of the new hire
			[parameter(Mandatory = $True)]
			[Alias('User')]
			[string]$LogonName,
			#Case number provided by new hire case.
			[parameter(Mandatory = $True)]
			[Alias('Case')]
			[int]$CaseID,
			#Initials of technician creating the new hire
			[string]$Tech = $env:username,
			#1st Owner of mailbox - Uses SamAccountName
			[parameter(Mandatory = $true)]
			[string]$Owner,
			#Purpose for mailbox
			[parameter(Mandatory = $true)]
			[string]$Purpose,
			#Invoke Pester Integration Tests
			[switch]$Test
			
		)
		
		BEGIN
		{
		}
		PROCESS
		{	
			$Date = Get-Date -UFormat "%m/%d/%y"
			 
			#User Hash table
			$UProp = @{
				identity = "$LogonName";
                Server = "Company.DomainController.local";
				Enabled = $True;
			}

			#Getting Owner name for Notes section in AD
			$NoteOwner = (Get-ADUser $Owner).GivenName
			$Note2Owner = (Get-ADUser $Owner).Surname
			
			#Setting of Notes section within Active Directory
			Set-ADUser  @UProp -Add @{
				info = "Created $Date Case $CaseID - $Tech
Owner: $NoteOwner $Note2Owner
Purpose: $Purpose"	
			} #Splat

		Write-Verbose 'Notes Added to Active directory Object'
				
			if ($Test) {

				Describe "Integration Tests" {
					BeforeAll {		
						$AdServer = "Company.DomainController.local"		
						$CRMailbox = Get-ADUser $LogonName -Properties * -Server $AdServer
					}#Beforeal
						Context "Active Directory Tests" {
							It "Account is Enabled" {
								$CRMailbox.Enabled | Should Be $UProp.Enabled
							}#it
						}#Context
				}#describe
			}
        }
		END
		{
			Send-ConfirmationEmail -ResourceMailbox
		}
	}
	#================================================================================================================================
	#==============================================Mail Contact Functions============================================================
	#================================================================================================================================
	
<#
.Synopsis
  Creates a New Company Mail Contact
.DESCRIPTION
  Creates a new Company Mail Contact. This cmdlet will create the mail contact and also fill out the AD properties per Company 
  policy. 
.EXAMPLE
   New-CompanyMailContact -FirstName Bob -LastName Builder -CaseID 1231231 -ExternalEmail Bob.Builder@aol.com -Location CR
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Creates Mail Contact that can be seen in EAC and AD.
.FUNCTIONALITY
   Requires that a connection to EAC be made.
#>
	
	function New-CompanyMailContact
	{
		[CmdletBinding()]
		param (
			#First name of contact
			[Parameter(Mandatory=$True)]
			[string]$FirstName,
			#Last name of contact
			[Parameter(Mandatory=$True)]
			[string]$LastName,
			#Case Number
			[Parameter(Mandatory=$true)]
			[int]$CaseID,
			#External Email address for user
			[Parameter(Mandatory=$true)]
			[string]$ExternalEmail,
			#Location of the contact
			[Parameter(Mandatory=$false)]
			[string]$Location,
			#Invokes Pester integration tests
			[switch]$Test
		)
		
		BEGIN
		{}
		PROCESS
		{
			#Organizational unit
			$OUnit = 'Company.DC.local/IT/Shared/Contacts'
				
				$Date = Get-Date -Format d
				
				#Variable for Contact Alias
				$Alias = "$FirstName$LastName"
				
				#Hashtable for Contact creation
				$UProp = @{
					FirstName = "$FirstName";
					LastName = "$LastName";
					Name = "$LastName, $FirstName"
					Alias = "$Alias";
					DisplayName = "$LastName, $FirstName ($Location)";
					ExternalEmailAddress = $ExternalEmail;
					OrganizationalUnit = "$OUnit";
				}

				#Hashtable for setting of contact
                $MProp = @{
                    identity = "$Alias";
					Notes = "Contact Created - Case# $CaseID $Date $env:USERNAME";
                        }
			
			#Create Mail Contact
			New-MailContact @UProp 

			#Hides contact from the Global Address Book
			Set-MailContact @MProp -HiddenFromAddressListsEnabled $True
			
			#Adds notes to Active Directory notes section
			Set-Contact @MProp

			#if statement for adding location to contact
            if ($Location)
            {
			Set-Contact @MProp -CountryOrRegion "$Location"
            } #if

			if ($Test) {
				Describe "Integration Tests" {
					BeforeAll {
						$CContact = Get-MailContact $Alias
					}#BeforeAll
					Context "Exchange Online & Active Directory Tests" {
						It "First Name Correct" {
							$CContact.FirstName | Should Be $UProp.FirstName
						}#it
						It "Correct Last Name" {
							$CContact.LastName | Should Be $UProp.LastName
						}#it
						It "Name Correct" {
							$CContact.Name | Should Be $UProp.Name
						}#it
						It "Correct Alias" {
							$CContact.Alias | Should Be $Alias
						}#it
						It "Display Name correct" {
							$CContact.DisplayName | Should Be $UProp.DisplayName
						}#it
						It "External Email Address Correct" {
							$CContact.ExternalEmailAddress | Should Be $ExternalEmail
						}#it
						It "Correct OU" {
							$CCContact.OrgranizationalUnit | Should Be $OUnit
						}#it
					}#Context
				}#Describe
			}#if

			Send-ConfirmationEmail -MailContact
		}
		END
		{}	
	}

<#
.Synopsis
  Removes a Company Mail Contact
.DESCRIPTION
  Removes a Company mail contact. This command will ask for th aliasand remove it from the enviroment. 
.EXAMPLE
   Remove-CompanyMailContact BobBuilder
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
  Removes Mail contact from EAC and Active Directory
.FUNCTIONALITY
   Requires that a connection to EAC be made.
#>
	
	function Remove-CompanyMailContact
	{
		[CmdletBinding()]
		param (
			#Alias of contact
			[Parameter(Mandatory=$True)]
			[string]$Alias
		)
		
		BEGIN
		{}
		PROCESS
		{
			Remove-MailContact $Alias
		}
		END
		{}	
	}

	#================================================================================================================================
	#==============================================Company New User Functions===========================================================================
	#================================================================================================================================

<#
.Synopsis
  Creates a New Company User Account including mailbox
.DESCRIPTION
  Creates a new Company Active Directory account including an Exchange mailbox. 
.EXAMPLE
   New-CompanyUser -Username bbuilder -FirstName 'Bob' -LastName 'Builder' -NewPassword ******** -Location 'US'
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Creates Mail Contact that can be seen in EAC and AD.
.FUNCTIONALITY
   Requires that a connection to EAC be made.
#>
function New-CompanyUser
{
	
	[CmdletBinding()]
	param (
		#Username of the new hire
		[parameter(Mandatory = $True)]
		[string]$Username,
		#First name of user.
		[parameter(Mandatory = $True)]
		[string]$FirstName,
		#Last Name of user.
		[string]$LastName,
		#Password for new user
		[string]$NewPassword,
		#Location to switch process of creation
		[Parameter(Mandatory = $false)]
		[string]$Location,
		#Invoke Pester integration tests
        [switch]$Test,
		#Switch for Contractor1 employee
		[switch]$Contractor1,
		#Switch for Contractor2 employee
		[switch]$Contractor2
		
	)
	BEGIN
	{ }
	PROCESS
	{
		#Takes Password from input
		$UsedPassword = ConvertTo-SecureString -AsPlainText $NewPassword -Force
		Write-Verbose 'Password Set'
		
		#Organizational unit that the user will be added to
		$OUnit = 'Company.DC.local/IT/Shared/Pending New Hires'
		
		if ($Contractor1)
		{
			#Hash Table for the new user
			$UProp = @{
				Name = "$LastName, $FirstName (Contractor1)";
				Alias = "$Username";
				OrganizationalUnit = "$OUnit";
				UserPrincipalName = "$FirstName.$LastName@Contractor1.com";
				SamAccountName = "$Username";
				FirstName = "$FirstName";
				LastName = "$LastName";
                Password = $UsedPassword;
                ResetPasswordOnNextLogon = $false;
			}#Hash

            $UMailboxProperties = @{
                Identity = $Username;
                CustomAttribute11 = 'DomainMail';
                EmailAddressPolicyEnabled = $false;
                EmailAddresses = "$FirstName.$LastName@Contractor1.com"
            }#Hash

			Write-Verbose 'Variables Set'
			
			#Command to create the new users mailbox
			New-Mailbox @UProp
			Write-Verbose 'Mailbox and user created in EAC'
			
			#Command to add the alias email addresses for Lync and the normal Companytravel.com email address
			Set-Mailbox @UMailboxProperties
			Write-Verbose 'Email Addresses Set'

			#Variable for Users EmailAddress
			$OriginalUserEmail = ((Get-Mailbox $Username).EmailAddresses)

			#Adding Sip Email addresses - Disabled until Email Policy has been revised
			Set-Mailbox $Username -EmailAddresses ($OriginalUserEmail += "smtp:$FirstName.$LastName@lync.Company.com", "smtp:$Firstname.$LastName@Company.com", "sip:$FirstName.$LastName@lync.Company.com") -Verbose
			Write-Verbose 'Email Addresses Set for Contractor1 User'
		} #if
		
		#switch used to know if user should be M&E account
		elseif ($Contractor2)
		{
			#Hash Table for the new user
			$UProp = @{
				Name = "$LastName, $FirstName ($Location)";
				Alias = "$Username";
				OrganizationalUnit = "$OUnit";
				UserPrincipalName = "$FirstName.$LastName@Contractor2.com";
				SamAccountName = "$Username";
				FirstName = "$FirstName";
				LastName = "$LastName";
                ResetPasswordOnNextLogon = $false;
                Password = $UsedPassword;
			}#Hash

            $UMailboxProperties = @{
                Identity = $Username;
                CustomAttribute11 = 'Domain-Contractor2Mail';
                EmailAddressPolicyEnabled = $false;
                EmailAddresses = "$FirstName.$LastName@Contractor2.com"
            }#Hash

			Write-Verbose 'Variables Set'
			#Command to create the new users mailbox
			New-Mailbox @UProp
			
			#Disabling Email Policy and setting address manually.#####Waiting for confirmation#####
			Set-Mailbox @UMailboxProperties
			Write-Verbose 'Mailbox and user created in EAC'
			
			#Adding Sip Email addresses - Disabled until Email Policy has been revised
			Set-Mailbox $Username -EmailAddresses (((Get-Mailbox $Username).EmailAddresses) += "$FirstName.$LastName@Company.com")
			Write-Verbose 'Email Addresses Set for Contractor2 User'
		}
		else
		{
			#Hash Table for the new user
			$UProp = @{
				Name = "$LastName, $FirstName ($Location)"
				Alias = "$Username";
				OrganizationalUnit = "$OUnit";
				UserPrincipalName = "$FirstName.$LastName@Company.com";
				SamAccountName = "$Username";
				FirstName = "$FirstName";
				LastName = "$LastName";
                ResetPasswordOnNextLogon = $false;
                Password = $UsedPassword;
			}#Hash 

            $UMailboxProperties = @{
                Identity = $Username;
                CustomAttribute11 = 'DomainMail';
                EmailAddressPolicyEnabled = $false;
                EmailAddresses = "$FirstName.$LastName@Company.com"
            }#Hash

			Write-Verbose 'Variables Set'
			
			#Command to create the new users mailbox
			New-Mailbox @UProp -Verbose
			Write-Verbose 'Mailbox and user created in EAC'
			
			#Disabling of Email Address Policy, sets email addresses manually. 
			Set-Mailbox @UMailboxProperties -Verbose

			#Original Email Address Variable 
			$OriginalUserEmail = ((Get-Mailbox $Username).EmailAddresses)
			
			#Command to add the alias email addresses for Lync and the normal Companytravel.com email address
			Set-Mailbox $Username -EmailAddresses ($OriginalUserEmail += "smtp:$FirstName.$LastName@lync.Company.com", "sip:$FirstName.$LastName@lync.Copmany.com")
			Write-Verbose 'Email Addresses Set'
		} #else

        if ($Test) {
        Describe "Intergration Tests" {
            BeforeAll {
				$CUser = Get-Mailbox $Username
            }#BeforeAll
            Context "Exchange Online Tests" {
				It "Full Name is Correct" {
					$CUser.Name | Should Be $UProp.Name
				}#it
				It "Alias is Correct" {
					$CUser.Alias | Should Be $UProp.Alias
				}#it
				It "Account in the Correct OU" {
                    $CUser.OrganizationalUnit | Should Be $UProp.OrganizationalUnit
                }#it
                It "SamAccountName is Correct" {
                    $CUser.SamAccountName | Should Be $UProp.SamaccountName
                }#it
				It "Pwd Reset on Logon Disabled" {
					$CUser.ResetPasswordOnNextLogon | Should Be $UProp.ResetPasswordOnNextLogon
				}#it
				It "Custom Atribute Set" {
					$CUser.CustomAttribute11 | Should Be $UMailboxProperties.CustomAttribute11
				}#it
				It "Email Policy Disabled" {
					$CUser.EmailAddressPolicyEnabled | Should Be $UMailboxProperties.EmailAddressPolicyEnabled
				}#it
            }#Context 
        }#Describe
	}#if "Test"
	}#Process
	
	END
	{ }
}

function Set-CompanyUser
{
	
	[CmdletBinding()]
	param (
		#Username of the new hire-Foir
		[parameter(Mandatory = $True)]
		[Alias('User')]
		[string]$Username,
		#Case number provided by new hire case.
		[parameter(Mandatory = $True)]
		[Alias('Case')]
		[int]$CaseID,
		#Initials of technician creating the new hire - Required
		[string]$Tech = $env:username,
		#Start date of new hire
		[parameter(Mandatory = $True)]
		[DateTime]$StartDate,
		#First name of new user
		[parameter()]
		[string]$FirstName,
		#Last name of new user
		[parameter()]
		[string]$LastName,
		#Location that decides which users account to create
		[parameter(Mandatory = $true)]
		[string]$Location,
		#Invoke Pester integration tests
		[switch]$Test
	)
	
	BEGIN
	{
	}
	PROCESS
	{
		$Date = Get-Date -UFormat "%m/%d/%y"
		
		#Main DC in Nasa Domain
		$CompanyDCServer = "Company.DC.local"
		
		if ($Location -eq 'BR')
		{
			#US Wikipedia Group
			$Group1 = 'Wikipedia Brazil'
			
			#User Hash table
			$UProp = @{
				identity = "$Username";
				Server = "$CompanyDCServer";
			}
			
			#Group 1 Hash table
			$MProp = @{
				identity = "$Group1";
				Members = "$Username";
				Server = "$CompanyDCServer";
			}

			Write-Verbose 'All Variables Set'
			#Setting of new user. Including the description, displayname and the info properties.
			Set-ADUser $Username -Description "Start Date - $StartDate" -Country 'BR' -DisplayName "$LastName, $FirstName (US)" -Add @{ info = "HRStuff Case $CaseID - $Date - $Tech" } -ErrorAction SilentlyContinue -Server "Company.DC.local"
			Write-Verbose 'AD Properties Added'
			
			#Adding of user to the group 1 Hash table
			Add-ADGroupMember @MProp
			Write-Verbose 'Added to Wiki group'
		} #if
		
		if ($Location -like '*US*')
		{
			#US Wikipedia Group
			$Group1 = 'Wiki US Group'
			
			#User Hash table
			$UProp = @{
				identity = "$Username";
				Server = "Company.DC.local";
			}
			
			#Group 1 Hash table
			$MProp = @{
				identity = "$Group1";
				Members = "$Username";
				Server = "Company.DC.local";
			}

			Write-Verbose 'All Variables Set'
			
			#Setting of new user. Including the description, displayname and the info properties.
			Set-ADUser $Username -Description "Start Date - $StartDate" -Country 'US' -DisplayName "$LastName, $FirstName (US)" -Add @{ info = "HRStuff Case $CaseID - $Date - $Tech" } -ErrorAction SilentlyContinue -Server "Company.DC.local"
			Write-Verbose 'AD Properties Added'
			
			#Adding of user to the group 1 Hash table
			Add-ADGroupMember @MProp
			Write-Verbose 'Added to Wiki group'
		} #if
		
		if ($Location -like '*CA*')
		{
			#US Wikipedia Group
			$Group1 = 'Wiki CA Group'
			
			#User Hash table
			$UProp = @{
				identity = "$Username";
				Server = "Company.DC.local";
			}
			
			#Group 1 Hash table
			$MProp = @{
				identity = "$Group1";
				Members = "$Username";
				Server = "Company.DC.local";
			}#Hash
			
			Write-Verbose 'All Variables Set'
			
			#Setting of new user. Including the description, displayname and the info properties.
			Set-ADUser $Username -Description "Start Date - $StartDate" -Country 'CA' -DisplayName "$LastName, $FirstName (CA)" -Add @{ info = "HRSC Case $CaseID - $Date - $Tech" } -ErrorAction SilentlyContinue -Server "Company.DC.local"
			Write-Verbose 'AD Properties Added'
			
			#Adding of user to the group 1 Hash table
			Add-ADGroupMember @MProp
			Write-Verbose 'Added to Wiki group'
		} #if
		
		if ($Location -eq 'CR')
		{
			#Main DC in Nasa Domain
			$NasaDCServer = "Company.DC.local"
			
			$Date = Get-Date -UFormat "%m/%d/%y"
			
			#US Wikipedia Group
			$Group1 = 'Wiki CR Group'
			
			#User Hash table
			$UProp = @{
				identity = "$Username";
				Server = "Company.DC.local";
			}#Hash
			
			#Group 1 Hash table
			$MProp = @{
				identity = "$Group1";
				Members = "$Username";
				Server = "Company.DC.local";
			}#Hash
			Write-Verbose 'All Variables Set'
			
			#Setting of new user. Including the description, displayname and the info properties.
			Set-ADUser $Username -Description "Start Date - $StartDate" -Country 'CR' -DisplayName "$LastName, $FirstName (CR)" -Add @{ info = "New Hire Case $CaseID - $Date - $Tech" } -ErrorAction SilentlyContinue -Server "Company.DC.local"
			Write-Verbose 'AD Properties Added'
			
			#Adding of user to the group 1 Hash table
			Add-ADGroupMember @MProp
			Write-Verbose 'Added to Wiki group'
		} #if
		
		if ($Location -eq 'IN')
		{
			#Main DC in Nasa Domain
			$NasaDCServer = "Company.DC.local"
			
			$Date = Get-Date -UFormat "%m/%d/%y"
			
			#US Wikipedia Group
			$Group1 = 'Wiki IN Group'
			
			#Group 1 Hash table
			$MProp = @{
				identity = "$Group1";
				Members = "$Username";
				Server = "Company.DC.local";
			}#Hash
			Write-Verbose 'All Variables Set'
			
			#Setting of new user. Including the description, displayname and the info properties.
			Set-ADUser $Username -Description "Start Date - $StartDate" -Country 'IN' -DisplayName "$LastName, $FirstName (IN)" -Add @{ info = "New Hire Case $CaseID - $Date - $Tech" } -ErrorAction SilentlyContinue -Server "Company.DC.local"
			Write-Verbose 'AD Properties Added'0
			
			#Adding of user to the group 1 Hash table
			Add-ADGroupMember @MProp
			Write-Verbose 'Added to Wiki group'
		} #if
		
		if ($Location -eq 'ME')
		{
			#Main DC in Nasa Domain
			$NasaDCServer = "Company.DC.local"
			
			#Pulls date of certain format
			$Date = Get-Date -UFormat "%m/%d/%y"
			
			#US Wikipedia Group
			$Group1 = 'Wiki US Group'
			
			  #Group 1 Hash table
			$MProp = @{
				identity = "$Group1";
				Members = "$Username";
				Server = "Company.DC.local";
			}#Hash
			
			Write-Verbose 'All Variables Set'
			
			#Setting of new user. Including the description, displayname and the info properties.
			Set-ADUser $Username -Description "Start Date - $StartDate" -Country 'US' -DisplayName "$LastName, $FirstName (US)" -Add @{ info = "M&E HR Case $CaseID - $Date - $Tech" } -Verbose -Server "Company.DC.local"
			Write-Verbose 'AD Properties Added'
			
			#Adding of user to the group 1 Hash table
			Add-ADGroupMember @MProp
			Write-Verbose 'Added to Wiki group'
			} #if
	
		Send-ConfirmationEmail -NewUser 
	}
	
	end
	{ }
}

<#
.Synopsis
   Grants Full & Send As permissions
.DESCRIPTION
   Grants Full & Send As permissions for Active Directory security group to mailbox in Exchange Online
.EXAMPLE
   Add-CompanySecurityGroup -SecurityGroup 'TWC Employees' -Destinationmailbox servicedeskmailbox
.EXAMPLE
   Another example of how to use this cmdlet
#>

function Add-CompanySecurityGroup {
    [CmdletBinding()]
    param (

    #Securtiy Group to be added
    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $SecurityGroup,

        #Mailbox Security Group will be added to
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Destinationmailbox
    )
    
    begin {
    }
    
    process {
        #Variable for mailbox
		$Mailbox = Get-Mailbox $DestinationMailbox

		#Variable for Security Group
		$Group = Get-ADGroup $SecurityGroup

		#Notes for AD Account
		$Notes = "
Access to $Mailbox - $env:USERNAME"

		Add-MailboxPermission $Mailbox.Alias -User $Group.DistinguishedName -AccessRights FullAccess -InheritanceType All 
							
		Add-ADPermission $Mailbox.Name -User $Group.DistinguishedName -AccessRights ExtendedRight -ExtendedRights "Send As" 

		#Setting of Notes section
		Set-ADGroup $Group.SamAccountName -Replace @{info = $Group.info + $Notes}
    }#process
    
    end {
    }#end
}#function

#############################################
<#
.SYNOPSIS
Starts Remote Connection to computer

.DESCRIPTION
Uses Landesk HTML5 protocol to start a remote desktop session using the Opera or Chrome browser.
This command starts a new process on those browers and enters the computer name after confirming that
it is online.

.PARAMETER ComputerName
Hostname or IP address of Target Machine

.EXAMPLE
Connect-ToComputer -ComputerName 'USmokscr300d

Creates a connection to the target computer USmokscr300d using either opera or chrome

.NOTES
General notes
#>
function Connect-ToComputer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,Position=0)]
        [String[]]$ComputerName
    )
    
    begin {
    }
    
    process {
        $ConnectionTest = Test-Connection $ComputerName -Count 1 -Quiet

        if ($ConnectionTest) {
            $MainVariable = (Test-Connection $ComputerName -Count 1).Address

            try {
                Start-Process chrome https:\\"$MainVariable":4343
            } catch {
                Start-Process opera https:\\"$MainVariable":4343
            } 
        }#if 
    }#process
    
    end {
    }
}
