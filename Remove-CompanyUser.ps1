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