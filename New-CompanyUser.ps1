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
	}
	
	end
	{ }
}