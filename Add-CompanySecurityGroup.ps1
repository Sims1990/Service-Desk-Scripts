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