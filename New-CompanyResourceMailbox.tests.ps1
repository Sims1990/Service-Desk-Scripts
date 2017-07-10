$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Resource Mailbox Functions" -Tags ExchangeMailbox {
    BeforeEach {
        $ExModule = New-Module -Name Exchange2 -Function "Set-Mailbox","New-MailboxExportRequest","Get-Mailbox","New-Mailbox","Add-MailboxPermission","Add-ADPermission","Send-ConfirmationEmail" -ScriptBlock {
                                                                                                    Function Set-Mailbox {'Set-Mailbox'};
                                                                                                    Function New-MailboxExportRequest {'New-MailboxExportRequest'};
                                                                                                    Function Get-Mailbox {'Get-Mailbox'};
                                                                                                    Function New-Mailbox{'New-Mailbox'};
                                                                                                    Function Add-MailboxPermission{'Add-MailboxPermission'};
                                                                                                    Function Add-ADPermission{'Add-ADPermission'};
                                                                                                    Function Send-ConfirmationEmail{'Send-ConfirmationEmail'}}
        $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount", "Get-ADOrganizationalUnit" -ScriptBlock {
                                                                                                    Function New-ADUser {"New-ADUser"};
                                                                                                    Function Get-ADUser {"Get-ADUser"};
                                                                                                    Function Set-ADUser {"Set-ADUser"};
                                                                                                    Function Unlock-ADAccount {"Unlock-ADAccount"};
                                                                                                    Function Set-ADAccountPassword {"Set-ADAccountPassword"};
                                                                                                    Function Remove-ADGroupMember {"Remove-ADGroupMember"};
                                                                                                    Function Move-ADObject {"Move-ADObject"};
                                                                                                    Function Enable-ADAccount {"Enable-ADAccount"};
                                                                                                    Function Disable-ADAccount {"Disable-ADAccount"};
                                                                                                    Function Get-ADOrganizationalUnit {"Get-ADOrganizationalUnit"}
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
    }#BeforeEach
    AfterEach {
        Get-Module Exchange2 | Remove-Module
    }
    Context "Creating Resource Mailbox Exchange Side" -Tag CreateResourceMailbox {
        It "Creation of Group Mailbox One Owner" {
            New-ResourceMailbox -DisplayName 'Test Bx' -LogonName 'testbx' -NewPassword 'Mt1231231' -Owner 'msims'
        }#it
        It "Adds One Owner" {
            New-ResourceMailbox -DisplayName 'Test Bx' -LogonName 'testbx' -NewPassword 'Mt1231231' -Owner 'msims'
            Assert-MockCalled Add-MailboxPermission -Times 1 -Scope It -Exactly
            Assert-MockCalled Add-ADPermission -Times 1 -Scope It -Exactly
        }
        It "Creates Mailbox in EAC" {
            New-ResourceMailbox -DisplayName 'Test Bx' -LogonName 'testbx' -NewPassword 'Mt1231231' -Owner 'msims'
            Assert-MockCalled New-Mailbox -Times 1 -Scope It -Exactly
        }#it
        It "Sets defined properties" {
            New-ResourceMailbox -DisplayName 'Test Bx' -LogonName 'testbx' -NewPassword 'Mt1231231' -Owner 'msims'
            Assert-MockCalled Set-Mailbox -Times 1 -Scope It -Exactly 
        }#it
    }#Context
    Context "Setting of Resource Mailbox Active Directory" -Tag SetResourceMailbox {
        It "Does not write to console" {
            Set-ResourceMailbox -LogonName 'test23' -CaseID '1231231' -Owner 'msims' -Purpose 'Testing'
            Assert-MockCalled Write-Host -Times 0 -Scope Context -Exactly
        }#It
        It "Setting of Mailbox Attributes" {
            Set-ResourceMailbox -LogonName 'test23' -CaseID '1231231' -Owner 'msims' -Purpose 'Testing'
            Assert-MockCalled Set-ADUser -Times 1 -Scope It -Exactly
        }#It
        It "CaseID Only Accepts Integers" {
            {Set-ResourceMailbox -LogonName 'test23' -CaseID 'werwerw' -Owner 'msims' -Purpose 'Testing'} | Should throw
        }#it
        It "No Sleep Commands" {
            Assert-MockCalled Start-Sleep -Times 0 -Scope Context -Exactly
        }#It
        It "Confirmation Email Sent" {
            Set-ResourceMailbox -LogonName 'test23' -CaseID '1231231' -Owner 'msims' -Purpose 'Testing'
            Assert-MockCalled Send-ConfirmationEmail -ParameterFilter {@{ResourceMailboxEmail='yes';ErrorAction='SilentlyContinue'}} -Times 1 -Scope It -Exactly
        }#it
    }#Context
}#Describe