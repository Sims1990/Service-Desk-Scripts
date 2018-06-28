#================================================================================================================================
	#==============================================US Distribution Functions===========================================================================
	#================================================================================================================================
	
<#
.Synopsis
  Creates a New Distribution List
.DESCRIPTION
  Creates a new Distribution list. This cmdlet will create the Distribution list 
  policy. 
.EXAMPLE
   Create-DistributionList -DisplayName US Bob List -Owner A01testuser -Alias USBobList -Purpose 'Bob Account Communication'
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Creates Distribution list the can be seen in EAC and AD.
.FUNCTIONALITY
   Requires that a connection to EAC be made.
#>

function Create-DistributionList
	{
		[CmdletBinding()]
		param (
			#Display Name of Distribution List
			[parameter(Mandatory = $True)]
			[string]$DisplayName,
			#Alias.
			[parameter(Mandatory = $True)]
			[string]$Alias,
			#Owner of the Distribution List
			[parameter(Mandatory = $true)]
			[string]$Owner,
			#Purpose for Distribution List
			[parameter (Mandatory = $true)]
			[string]$Purpose
		)
		
		BEGIN
		{
			Write-Verbose "[DistributionList]Starting Alias and DisplayName checks"
			#If Statements to check formats are correct [Custom check requested]
			if ($DisplayName -like "*DL:*"){
			Throw "
Incorrect Displayname format. Please exclude the 'DL:'. 

Example: 

Display Name: Bell Account

------------------------------------

The Service Desk Console will add the DL: to the name entered in the display name field."
		}
			if ($Alias -like "*DL_*") {
				Throw "
Incorrect Alias format error. Please exclude the 'DL_'.

Example: 

Alias: BellAccount

----------------------------------------

The Service Desk console will add the DL_ to the name entered in the alias field."
			}
			Write-Verbose "[DistributionList]Pre-Run checks done"
	}
		PROCESS
		{
			Write-Verbose "[DistributionList]Creating variables"
			
			#OwnerUsername
			$OwnerName = (Get-User $Owner).FirstName + (Get-User $Owner).LastName
			Write-Verbose "[DistributionList]Owner variable created"
			
			
Write-Verbose "[DistributionList]Description variable finished"

			#Places in desired Organizational Unit
			$OUnit = 'Enter Organizational Unit Here'
			Write-Verbose "[DistributionList]Organizational Unit selected"
			
			
				#Hash Table for Variables
			$UProp = @{
				Name = "DL: $DisplayName";
				Alias = "DL_$Alias";
				OrganizationalUnit = "$OUnit";
				DisplayName = "DL: $DisplayName";
				PrimarySmtpAddress = "DL_$Alias@Domain.com";
				Notes = "$Notes";
				Managedby = "$Owner"
			} #Hash
			Write-Verbose "[DistributionList]Hash tables created"
			
			New-DistributionGroup @UProp -Verbose
			Write-Verbose "[DistributionList]Distribution List created"
		}
		END
		{
			$Properties = @{'Name'="$DisplayName";
			'Owner'="$Owner";
			'Notes'="$Notes";
			'Technician'="$env:username"}
 $obj = New-Object -TypeName PSObject -Property $Properties
Write-Output $obj
		}#end
	}