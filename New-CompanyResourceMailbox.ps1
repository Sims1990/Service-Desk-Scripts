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
			Send-ConfirmationEmail -ResourceMailboxEmail
		}
	}