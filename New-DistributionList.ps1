<#
.Synopsis
  Creates a New Distribution List
.DESCRIPTION
  Creates a new Distribution list. This cmdlet will create the Distribution list and also fill out the AD properties per Company requirements. This cmdlet will require a connection to Exchange online. 
.EXAMPLE
   New-DistributionList -DisplayName US Bob List -CaseID 1231231 -Owner A01testuser -Alias USBobList -Purpose 'Bob Account Communication'
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Creates Distribution list the can be seen in Exchange Online & Active Directory.
.FUNCTIONALITY
   Requires that a connection to EAC be made.
#>
	
	function New-DistributionList
	{
		[CmdletBinding()]
		param (
			#Display Name of Distribution List
			[parameter(Mandatory = $True)]
			[string]$DisplayName,
			#Alias.
			[parameter(Mandatory = $True)]
			[string]$Alias,
			[Parameter(Mandatory=$true)]
			#Case Number per case
			[int]$CaseID,
			#Owner of the Distribution List - must be SamAccountName
			[parameter(Mandatory = $true)]
			[string]$Owner,
			#Purpose for Distribution List
			[parameter (Mandatory = $true)]
			[string]$Purpose,
			#Switch for Company ME Group
			[switch]$CompanyME,
			#Invokes Pester Integration tests
			[switch]$Test
		)
		
		BEGIN
		{}
		PROCESS
		{
			#Variable for technicians initials
			$Tech = $env:username
			
			$Date = Get-Date -UFormat "%m/%d/%y"
			
			$Desc = "Created per Case# $CaseID - $Date - $Tech"
			
			#Organizational Unit of Distribution List
			$OUnit = 'Company.local/IT/Shared/DistributionLists'

			if ($CompanyME)	{
				#Hash Table for M&E Distro Variables
			$UProp = @{
				Name = "Company ME: $DisplayName";
				Alias = "Company_ME_$Alias";
				OrganizationalUnit = "$OUnit";
				DisplayName = "Company ME: $DisplayName";
				MemberJoinRestriction = "Closed";
				PrimarySmtpAddress = "Company_ME_$Alias@Companyme.com";
				Notes = "$Desc";
				Managedby = "$Owner"
			} #HashCompanyME
			} else {
				#Hash Table for Distro Variables
			$UProp = @{
				Name = "DL: $DisplayName";
				Alias = "DL_$Alias";
				OrganizationalUnit = "$OUnit";
				DisplayName = "DL: $DisplayName";
				MemberJoinRestriction = "Closed";
				PrimarySmtpAddress = "DL_$Alias@Company.com";
				Notes = "$Desc";
				Managedby = "$Owner"
			} #Hash
			} #else 
			
			New-DistributionGroup @UProp

			if ($Test) {

				if ($CompanyME) {
					Start-Sleep -Seconds 2
				Describe "Integration Tests: $DisplayName" {
					BeforeAll {
						$CDistro = Get-DistributionGroup "Company_ME_$Alias"
					}#BeforeAll
					Context "Exchange Online Tests" {
						It "M&E Name is Correct" {
							$CDistro.Name | Should Be $UProp.Name
						}#it
						It "M&E Alias is correct" {
							$CDistro.Alias | Should Be $UProp.Alias
						}#it
						It "M&E OU is correct" {
							$CDistro.OrganizationalUnit | Should Be $UProp.OrganizationalUnit
						}#it
						It "M&E Display Name is correct" {
							$CDistro.DisplayName | Should Be $UProp.DisplayName
						}#it
						It "M&E Member Join Restriction Closed" {
							$CDistro.MemberJoinRestriction | Should Be $UProp.MemberJoinRestriction
						}#it
						It "M&E Primary SMTP is correct" {
							$CDistro.PrimarySMTPaddress | Should Be $UProp.PrimarySMTPaddress
						}#it
					}#Context
				}#Describe
				} else {
				Start-Sleep -Seconds 2
				Describe "Integration Tests: $DisplayName" {
					BeforeAll {
						$CDistro = Get-DistributionGroup "DL_$Alias"
					}#BeforeAll
					Context "Exchange Online Tests" {
						It "Name is Correct" {
							$CDistro.Name | Should Be $UProp.Name
						}#it
						It "Alias is correct" {
							$CDistro.Alias | Should Be $UProp.Alias
						}#it
						It "OU is correct" {
							$CDistro.OrganizationalUnit | Should Be $UProp.OrganizationalUnit
						}#it
						It "Display Name is correct" {
							$CDistro.DisplayName | Should Be $UProp.DisplayName
						}#it
						It "Member Join Restriction Closed" {
							$CDistro.MemberJoinRestriction | Should Be $UProp.MemberJoinRestriction
						}#it
						It "Primary SMTP is correct" {
							$CDistro.PrimarySMTPaddress | Should Be $UProp.PrimarySMTPaddress
						}#it
					}#Context
				}#Describe
			}#else
		}#if
		}
		END
		{
		}#end
	}