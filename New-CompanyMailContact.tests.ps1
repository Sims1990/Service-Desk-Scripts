$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Exchange Mail Contact" -Tag MailContact {
    BeforeEach {
        $ExModule = New-Module -Name Exchange2 -Function "Set-Mailbox","New-MailboxExportRequest","Get-Mailbox","New-Mailbox","Add-MailboxPermission","Add-ADPermission","Send-ConfirmationEmail", "New-MailContact", "Set-MailContact", "Remove-MailContact" -ScriptBlock {
                                                                                                    Function Set-Mailbox {'Set-Mailbox'};
                                                                                                    Function New-MailboxExportRequest {'New-MailboxExportRequest'};
                                                                                                    Function Get-Mailbox {'Get-Mailbox'};
                                                                                                    Function New-Mailbox{'New-Mailbox'};
                                                                                                    Function Add-MailboxPermission{'Add-MailboxPermission'};
                                                                                                    Function Add-ADPermission{'Add-ADPermission'};
                                                                                                    Function Send-ConfirmationEmail{'Send-ConfirmationEmail'};
                                                                                                    Function New-MailContact {'New-MailContact'};
                                                                                                    Function Set-MailContact {'Set-MailContact'};
                                                                                                    Function Remove-MailContact {'Remove-MailContact'}}
        $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount", "Get-ADOrganizationalUnit", "Set-Contact" -ScriptBlock {
                                                                                                    Function New-ADUser {"New-ADUser"};
                                                                                                    Function Get-ADUser {"Get-ADUser"};
                                                                                                    Function Set-ADUser {"Set-ADUser"};
                                                                                                    Function Unlock-ADAccount {"Unlock-ADAccount"};
                                                                                                    Function Set-ADAccountPassword {"Set-ADAccountPassword"};
                                                                                                    Function Remove-ADGroupMember {"Remove-ADGroupMember"};
                                                                                                    Function Move-ADObject {"Move-ADObject"};
                                                                                                    Function Enable-ADAccount {"Enable-ADAccount"};
                                                                                                    Function Disable-ADAccount {"Disable-ADAccount"};
                                                                                                    Function Get-ADOrganizationalUnit {"Get-ADOrganizationalUnit"};
                                                                                                    Function Set-Contact {'Set-Contact'}
 }#Scriptblock
        $ADModule | Import-Module     
        $ExModule | Import-Module

        mock New-Mailbox 
        mock Add-ADPermission
        mock Add-MailboxPermission
        mock Set-Mailbox
        mock Start-Sleep
        mock Write-Host
        mock Send-ConfirmationEmail
        mock Set-ADUser
        mock Set-Contact
        mock Set-MailContact
        mock New-MailContact
        mock Remove-MailContact
    }#BeforeEach
    AfterEach {
        Get-Module ActiveDirectory | Remove-Module
        Get-Module Exchange2 | Remove-Module
    }#AfterEach
    Context "Exchange MailContact" {
        It "CaseID Only Accepts Integers" {
        {New-CompanyMailContact -FirstName 'Test' -LastName 'Tester' -CaseID 'gdgdgd' -ExternalEmail 'test1@hv.com' -Location 'US'} | Should throw
        }#it
        It "Creates New MailContact" {
            New-CompanyMailContact -FirstName 'Test' -LastName 'Tester' -CaseID '123123' -ExternalEmail 'test1@hv.com' -Location 'US'
            Assert-MockCalled New-MailContact -Times 1 -Scope It -Exactly
        }#
        It "Hides Contact from GAL" {
            Assert-MockCalled Set-MailContact -Parameter {@{HiddenFromAddressListsEnabled="$true"}} -Scope Context -Exactly
        }#it
        It "Adds Notes to Active Directory" {
            New-CompanyMailContact -FirstName 'Test' -LastName 'Tester' -CaseID '1231231' -ExternalEmail 'test1@hv.com' -Location 'US'
            Assert-MockCalled Set-Contact -Times 2 -Scope It -Exactly
        }#it
        It "Confirmation Email Sent" {
            New-CompanyMailContact -FirstName 'Test' -LastName 'Tester' -CaseID '1231231' -ExternalEmail 'test1@hv.com' -Location 'US'
            Assert-MockCalled Send-ConfirmationEmail -ParameterFilter{@{MailContactEmail='yes'}}
            Assert-MockCalled Send-ConfirmationEmail -Times 1 -Scope It -Exactly
        }#it
    }#Context
    Context "Removal of MailContact" {
        It "Remove MailContact" {
            Remove-CompanyMailContact 'test1'
            Assert-MockCalled Remove-MailContact -Times 1 -Scope Context -Exactly
        }#it
    }#Context
}#Describe