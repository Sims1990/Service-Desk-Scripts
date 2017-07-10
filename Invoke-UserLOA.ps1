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
		}
	} #function