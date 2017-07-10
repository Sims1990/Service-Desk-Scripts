$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Adding Company Security Group Function" {
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
        $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount", "Get-ADOrganizationalUnit", "Set-Contact", "Get-ADGroup", "Set-ADGroup" -ScriptBlock {
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
                                                                                                    Function Set-Contact {'Set-Contact'};
                                                                                                    Function Get-ADGroup {'Get-ADGroup'};
                                                                                                    Function Set-ADGroup {'Set-ADGroup'}
 }#Scriptblock
        $ADModule | Import-Module     
        $ExModule | Import-Module

        mock Get-Mailbox -ParameterFilter {@{identity='test1';DistingishedName='fjfjf'}}
        mock Get-ADGroup -ParameterFilter {@{identity='test1';DistingishedName='fjfjf'}}
        mock Set-ADGroup
        mock Add-MailboxPermission
        mock Add-ADPermission
    }#BeforeEach
    AfterEach {

    }#AfterEach
    Context "Add-CompanySecurityGroup" {
        It "Cmd Runs Correctly" {
            Add-CompanySecurityGroup -SecurityGroup 'test1' -Destinationmailbox 'msims'
        }#it
        It "Getting Destination Mailbox Variable" {
            Assert-MockCalled Get-Mailbox -Times 1 -Scope Context -Exactly
        }#it
        It "Permissions Added" {
            Assert-MockCalled Add-MailboxPermission -Times 1 -Scope Context -Exactly
            Assert-MockCalled Add-ADPermission -Times 1 -Scope Context -Exactly
        }#it
        It "Notes Section Updated" {
            Assert-MockCalled Set-ADGroup -Times 1 -Scope Context -Exactly
        }
    }
}