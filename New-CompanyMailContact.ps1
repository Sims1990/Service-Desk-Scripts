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

			Send-ConfirmationEmail -MailContactEmail
		}
		END
		{}	
	}#Function

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
	}#Function