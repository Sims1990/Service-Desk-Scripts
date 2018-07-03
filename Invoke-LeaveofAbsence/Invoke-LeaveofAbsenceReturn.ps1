<#
.Synopsis
   Returns user from Leave of Absence
.DESCRIPTION
   Sets AD properties and enables account for return from LOA.
.EXAMPLE
   Disable-LOAUser -Username bbuilder -CaseID '1234567' -NewPassword 'Bb1234567'
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

function Invoke-LeaveofAbsenceReturn
	{
		[CmdletBinding()]
		param
		(
			#SamAccountName of user
			[parameter(Mandatory = $true)]
			[string]$Username
		)
		BEGIN
		{}
		PROCESS
		{

			#Creation of Password
            Write-Verbose "[ReturnFromLeaveofAbsence] Preparing to create temp password"
            
            function CreatePassword {
				[System.Web.Security.Membership]::GeneratePassword(16,5)
            }
            
			$NewPassword = CreatePassword
			Write-Verbose "[ReturnFromLeaveofAbsence] Password created"
			
			#Variable Hash Table
			$UProp = @{
				identity = "$Username"
			}
			
			#Hash for notes section
			$Entry = @"

Enter account notes here 
"@

			#Variable to get existing notes section
			$CurrentAccountNotes = Get-ADUser @UProp -Properties info
			Write-Verbose "[ReturnFromLeaveofAbsence] Confirming Notes variable: $CurrentAccountNotes"

			#Enabling of user account
			Write-Verbose "[ReturnFromLeaveofAbsence] Preparing to enable Account"
			Enable-ADAccount @UProp -Verbose
			Write-Verbose "[ReturnFromLeaveofAbsence] Account enabled"

			#Changing of network password
			Write-Verbose "[ReturnFromLeaveofAbsence] Preparing to change users password"
			Set-ADAccountPassword @UProp -Reset -NewPassword (Convertto-SecureString -AsPlainText "$NewPassword" -Force) -Verbose
			Write-Verbose "[ReturnFromLeaveofAbsence] User passwored reset"

			#Setting of user properties
			Write-Verbose "[ReturnFromLeaveofAbsence] Preparing to update Active Directory Account"
			Set-ADUser @UProp -Description $null -Replace @{ info = $CurrentAccountNotes.info + $Entry } -ChangePasswordAtLogon $true
			Write-Verbose "[ReturnFromLeaveofAbsence] Account updated"

		}
		
		END
		{
			$Properties = @{'Name'="$Username";
			'Technician'="$env:username";
			'Return'="$true"}
 $obj = New-Object -TypeName PSObject -Property $Properties

 Write-Output $obj
}
	}