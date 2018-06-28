#====================================================================================================================================================================
####===============================================================Resource Mailbox Functions=========================================================
#====================================================================================================================================================================


<#
.Synopsis
   Creates Resource Mailbox in Exchange
.DESCRIPTION
   Resource Mailbox is created 
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Will create a generic mailbox with a @<example>com email address.
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
			#Alias name of mailbox
			[parameter(Mandatory = $True)]
			[string]$Alias,
			#Display Name of mailbox.
			[parameter(Mandatory = $True)]
			[string]$DisplayName,
			#Owner of mailbox
			[parameter(Mandatory = $True)]
			[string]$Owner,
			#Purpose of Mailbox - Uses SamAccountName
			[parameter(Mandatory = $true)]
			[string]$Notes
		)
		
		BEGIN
		{}
		PROCESS
		{
			
				#Organizational Unit where mailbox will be created.
				$OUnit = 'Enter Organizational Unit Here'
                Write-verbose '[ResourceMailboxCreate] Creating Organizational Unit Variable'
                
				$UProp = @{
					Name = "$DisplayName"
					Alias = "$Alias";
					OrganizationalUnit = "$OUnit";
					UserPrincipalName = "$Alias@Domain.com";
					SamAccountName = "$Alias";
					DisplayName = "$DisplayName"
                }#UProp	
			
			#Command to create the Resource mailbox
			Write-Verbose '[ResourceMailboxCreate] Beginning to create mailbox'
			New-Mailbox @UProp -Shared -Verbose
			Write-Verbose '[ResourceMailboxCreate] Mailbox Created'
			
			#Date Variable
			$Date = Get-Date -UFormat "%m/%d/%y"
			Write-Verbose '[ResourceMailboxCreate] Date Variable Set'

			#User Hash table
			$UProp = @{
				identity = "$Alias";
			}#User Hash Table
			
			#Variable for adding owner permission
			$GrpName = (Get-Mailbox $Alias).Name
			
			#Getting Owner name for Notes section
			Write-Verbose '[ResourceMailboxCreate] Gathering Mailbox Owner Information'
			$NoteOwner = (Get-User $Owner).FirstName
			$Note2Owner = (Get-User $Owner).LastName
			Write-Verbose '[ResourceMailboxCreate] Owner Information Gathered'
			
			#Setting of Notes section
			Write-Verbose '[ResourceMailboxCreate]Beginning to Write Notes to Active Directory'
			Set-User  @UProp -Notes "Created $Date - $env:username
Owner: $NoteOwner $Note2Owner
Notes: $Purpose" -Verbose
			Write-Verbose '[ResourceMailboxCreate] Notes written to Active Directory'
			#Notes

            #Production code adds multiple owners here
			Write-Verbose '[ResourceMailboxCreate] Creating Array of Owners.'
			$collection = "$Owner"
			
			foreach ($GrpUser in $collection)
			{
				Write-Verbose "[ResourceMailboxCreate][User=$GrpUser] Beginning to Grant Permissons to $Alias"
				Add-MailboxPermission $Alias -User $GrpUser -AccessRights FullAccess -InheritanceType All -Verbose
				
				Add-ADPermission "$GrpName" -User $GrpUser -AccessRights ExtendedRight -ExtendedRights "Send As" -Verbose
				Write-Verbose "[ResourceMailboxCreate][User=$GrpUser] Finished Granting Permissions to $Alias"
			} #foreach
			
		}#Process
		END
		{
			
			
			$Properties = @{'Name'="$DisplayName";
			'Date'="$Date";
			'Technician'="$env:username";
			'Owner'="$Owner"}
 $obj = New-Object -TypeName PSObject -Property $Properties

 Write-Output $obj
 		}
		
    }