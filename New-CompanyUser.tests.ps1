$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "New User Creation" -Tags NewUserCreate {
    BeforeEach {
        $ExModule = New-Module -Name Exchange2 -Function "Set-Mailbox","New-MailboxExportRequest","Get-Mailbox","New-Mailbox","Add-MailboxPermission","Add-ADPermission","Send-ConfirmationEmail", "New-MailContact", "Set-MailContact", "Remove-MailContact","Add-ADGroupMember" -ScriptBlock {
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
        $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Add-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount", "Get-ADOrganizationalUnit", "Set-Contact" -ScriptBlock {
                                                                                                    Function New-ADUser {"New-ADUser"};
                                                                                                    Function Get-ADUser {"Get-ADUser"};
                                                                                                    Function Set-ADUser {"Set-ADUser"};
                                                                                                    Function Unlock-ADAccount {"Unlock-ADAccount"};
                                                                                                    Function Set-ADAccountPassword {"Set-ADAccountPassword"};
                                                                                                    Function Remove-ADGroupMember {"Remove-ADGroupMember"};
                                                                                                    Function Add-ADGroupMember {"Add-ADGroupMember"}
                                                                                                    Function Move-ADObject {"Move-ADObject"};
                                                                                                    Function Enable-ADAccount {"Enable-ADAccount"};
                                                                                                    Function Disable-ADAccount {"Disable-ADAccount"};
                                                                                                    Function Get-ADOrganizationalUnit {"Get-ADOrganizationalUnit"};
                                                                                                    Function Set-Contact {'Set-Contact'}
 }#Scriptblock
        $ADModule | Import-Module     
        $ExModule | Import-Module

        mock New-Mailbox
        mock Set-Mailbox
        mock Add-ADGroupMember
        mock Send-ConfirmationEmail
        mock Start-Sleep
        mock Set-Aduser
        mock Add-ADGroupMember

    $SUser = Set-CompanyUser -Username 'test1' -CaseID '1231231' -StartDate '03/03/2018' -FirstName 'test' -LastName '1' -Location 'US'
    }#BeforeEach
    AfterEach {
        Get-Module ActiveDirectory | Remove-Module
        Get-Module Exchange2 | Remove-Module
    }#AfterEach
    Context "Create New User" -Tag CreateUser {
        It "New User is Created" {
            New-CompanyUser -Username 'test1' -FirstName 'test' -LastName '1' -NewPassword 'P@ssw0rd' -Location 'US'
            Assert-MockCalled New-Mailbox -Times 1 -Scope It -Exactly
    }#it
        It "Custom Attribute Added" {
            Assert-MockCalled Set-Mailbox -ParameterFilter {@{CustomAttribute11='NASAMail'}} -Scope Context
        }#it
        It "Email Policy Disabled" {
            Assert-MockCalled Set-Mailbox -Parameterfilter {@{EmailAddressPolicyEnabled="$false"}} -Scope Context
        }#It
        It "Pwd Reset on Logon Disabled" {
            Assert-MockCalled Set-Mailbox -ParameterFilter {@{ResetPasswordONNextLogon="$false"}}
        }#it
    }#Context
    Context "Setting of Attributes" -Tag SetUser {
        It "CaseID only accepts Integers" {
            {Set-CompanyUser -Username 'test1' -CaseID 'tererr' -StartDate '03/03/2018' -FirstName 'test' -LastName '1' -Location 'US'} | Should throw
        }#it
        It "Start Date has to be a DateTime Property" {
            {Set-CompanyUser -Username 'test1' -CaseID '1231231' -StartDate '332018' -FirstName 'test' -LastName '1' -Location 'US'} | Should throw

        }#it
        It "No Sleeping Within Module" {
            Assert-MockCalled Start-Sleep -Times 0 -Scope Context -Exactly
        }#it
        It "Sets Properties for AD User" {
            $SUser
            Assert-MockCalled Set-ADUser -Times 1 -Scope It -Exactly
        }#it
        It "Adding Account to Wiki Group" {
            $SUser
            Assert-MockCalled Add-ADGroupMember -Times 1 -Scope It -Exactly
        }#it
    }#Context
}#Describe