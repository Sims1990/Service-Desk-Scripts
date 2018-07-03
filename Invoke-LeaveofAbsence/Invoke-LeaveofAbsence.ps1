<#
.Synopsis
  Puts user on Leave of Absence status per Company Guidelines.
.DESCRIPTION
   Sets Active Directory properties, disables account and moves account to different organizational ou.
.EXAMPLE
   Invoke-LeaveofAbsence -Username bbuilder
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

function Invoke-LeaveofAbsence
	{
		[CmdletBinding()]
		param
		(
			#Username Parameter
			[parameter(Mandatory = $true)]
			[string]$Username
		)
		
		BEGIN
		{}
		PROCESS
		{
			#Adding of .net assembly
			Add-Type -AssemblyName System.Web

			Write-Verbose "[LeaveofAbsence]Preparing to create new password"
			function CreatePassword {
				[System.Web.Security.Membership]::GeneratePassword(16,5)
			}
			Write-Verbose "[LeaveofAbsence]Password function loaded"

			#Creation of Password
			$NewPassword = CreatePassword
			Write-Verbose "[LeaveofAbsence]Password created"
			
			#Hash Table
			$UProp = @{
				identity = "$Username"
			}
			
			#Property for notes section [Feel free to edit this section]
			$Notes = @"

Enter Notes here
"@

			Write-Verbose "[LeaveofAbsence]Preparing to add notes to account"
			#Variable for Notes section
			$CurrentInformation = Get-ADUser @UProp -Properties info
			Write-Verbose "[LeaveofAbsence]Notes set"

			Write-Verbose "[LeaveofAbsence]Adding Notes to Description in Acitve Directory"
			$Description = "Custom Message Here - $env:username"
			Write-Verbose "[LeaveofAbsence]Description set"

			#Setting of user properties including Description and info fields
			Write-Verbose "[LeaveofAbsence]Writing remaining Active Directory Notes"
			Set-ADUser @UProp -Description $Description -Replace @{ info = $CurrentInformation.info + $Notes }
			Write-Verbose "[LeaveofAbsence]Notes updated on Active Directory Account"
			
			#Change of password
			Write-Verbose "[LeaveofAbsence]Resetting password"
			Set-ADAccountPassword $Username -Reset -NewPassword (Convertto-SecureString -AsPlainText "$NewPassword" -Force)
			Write-Verbose "[LeaveofAbsence]Password reset"

		}#process

		END
		{
			$Properties = @{
			'Technician'="$env:username";
			'Return'="$false"}
 $obj = New-Object -TypeName PSObject -Property $Properties

Write-Output $obj
		}
	} #function