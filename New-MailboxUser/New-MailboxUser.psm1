function New-MailboxUser
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
		[string]$LastName
		
	)
	BEGIN
	{ 
		
	 }
	PROCESS
	{

		

				#Adding of .net assembly
				Add-Type -AssemblyName System.Web -Verbose

				function CreatePassword {
					[System.Web.Security.Membership]::GeneratePassword(16,5)
				}
		
				#Create New Password
				$NewPassword = CreatePassword
				Write-Verbose "[New User Creation] Random Password Created"
		

		#Takes Password from input
		Write-Verbose "[New User Creation] Applying Password"
		$UsedPassword = ConvertTo-SecureString -AsPlainText $NewPassword -Force -Verbose
		Write-Verbose "[New User Creation] Password Set For Account"
		
		#Organizational unit that the user will be added to
		$OUnit = 'Enter organizational unit'
		Write-Verbose "[New User Creation] Organizational Unit Variable Created"
		
		<#Put in for special case
		Write-Verbose "[New User Creation] Removing and extra spaces in Lastname"
		$LastNameNoSpace = $LastName -replace '\s',''
		Write-Verbose "[New User Creation] Removing extra spaces in Firstname"
		$FirstNameNoSpace = $FirstName -replace '\s',''
		Write-Verbose "[New User Creation] Spaces Removed"#>
		
		
			#Hash Table for the new user
			$UProp = @{
				Name = "$LastName, $FirstName"
				Alias = "$Username";
				OrganizationalUnit = "$OUnit";
				UserPrincipalName = "$FirstName.$LastNameNoSpace@Domain.com";
				SamAccountName = "$Username";
				FirstName = "$FirstName";
				LastName = "$LastName";
                ResetPasswordOnNextLogon = $false;
                Password = $UsedPassword;
			}#Hash 

            $UMailboxProperties = @{
                Identity = $Username;
				EmailAddresses = "$FirstName.$LastNameNoSpace@Domain.com";
				
            }#Hash

			Write-Verbose "[New User Creation] User Variables Set"
			
			#Command to create the new users mailbox
			Write-Verbose "[New User Creation] Preparing to Create Mailbox"
			New-Mailbox @UProp -Verbose
			Write-Verbose "[New User Creation] Mailbox Created"
			
			#Disabling of Email Address Policy, sets email addresses manually.
			Write-Verbose "[New User Creation] Preparing to Configure Properties"
			Set-Mailbox @UMailboxProperties -Verbose
			Write-Verbose "[New User Creation] Properties Configured"
		
			#Setting of DisplayName
			Set-Mailbox $Username -DisplayName "$LastName, $FirstName"
			Write-Verbose "[New User Creation] Display Name set"

			#Setting of new user. Including the description, displayname and the info properties.
			Set-User $Username -Notes "Enter Notes here"
			Write-Verbose "[New User Creation] Notes Applied"
			Write-Verbose "[New User Creation] Password Set to Reset"
			
	}#Process
	
	END
	{ 
		
		$Properties = @{'FirstName'="$FirstName";
		'LastName'="$LastName";
		'Username'="$Username";
		'Technician'="$env:USERNAME";}
$obj = New-Object -TypeName PSObject -Property $Properties

Write-Output $obj}
}