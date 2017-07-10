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
		{		}
	}