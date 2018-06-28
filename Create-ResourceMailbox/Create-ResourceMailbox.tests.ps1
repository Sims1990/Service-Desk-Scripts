$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

####==============================Resource Mailbox
#####
Describe "Resource Mailbox Functions" -Tags ExchangeMailbox {
    BeforeEach {
        $ExModule = New-Module -Name Exchange2 -Function "Set-Mailbox","New-MailboxExportRequest","Get-Mailbox","New-Mailbox","Add-MailboxPermission","Add-ADPermission","Get-User","Set-User" -ScriptBlock {
                                                                                                    Function Set-Mailbox {'Set-Mailbox'};
                                                                                                    Function New-MailboxExportRequest {'New-MailboxExportRequest'};
                                                                                                    Function Get-Mailbox {'Get-Mailbox'};
                                                                                                    Function New-Mailbox{'New-Mailbox'};
                                                                                                    Function Add-MailboxPermission{'Add-MailboxPermission'};
                                                                                                    Function Add-ADPermission{'Add-ADPermission'};
                                                                                                    Function Get-User{'Get-User'};
                                                                                                    Function Set-User{'Set-User'}}
        $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount", "Get-ADOrganizationalUnit", "New-ADGroup", "Get-ADGroup", "Get-ADDomain" -ScriptBlock {
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
                                                                                                    Function New-ADGroup {"New-ADGroup"};
                                                                                                    Function Get-ADGroup {"Get-ADGroup"};
                                                                                                    Function Get-ADDomain {"Get-ADDomain"}
 }#Scriptblock
        $ADModule | Import-Module     
        $ExModule | Import-Module
        
        mock New-Mailbox 
        mock Add-ADPermission
        mock Add-MailboxPermission
        mock Get-Mailbox
        mock Set-Mailbox
        mock Start-Sleep
        mock Write-Host
        mock Set-User
        mock New-ADGroup
        mock Get-ADGroup
        mock Get-ADDomain
    }#BeforeEach
    AfterEach {
        Get-Module Exchange2 | Remove-Module
    }
    Context "Creating Resource Mailbox Exchange Side"  {
        It "Creation of Group Mailbox One Owner" {
            New-ResourceMailbox -DisplayName 'Test Bx' -Alias 'testbx' -Owner 'msims' -Notes 'Testing'
        }#it
        It "Adds One Owner" {
            New-ResourceMailbox -DisplayName 'Test Bx' -Alias 'testbx' -Owner 'msims' -Notes 'Testing'
            Assert-MockCalled Add-MailboxPermission -Times 1 -Scope It -Exactly
            Assert-MockCalled Add-ADPermission -Times 1 -Scope It -Exactly
        }
        It "Creates Mailbox in EAC" {
            New-ResourceMailbox -DisplayName 'Test Bx' -Alias 'testbx' -Owner 'msims' -Notes 'Testing'
            Assert-MockCalled New-Mailbox -Times 1 -Scope It -Exactly
        }#it
    }#Context
    Context "Setting of Resource Mailbox Active Directory"  {
        It "Does not write to console" {
            New-ResourceMailbox -DisplayName 'Test Bx' -Alias 'testbx' -Owner 'msims' -Notes 'Testing'
            Assert-MockCalled Write-Host -Times 0 -Scope Context -Exactly
        }#It
        It "Setting of Mailbox Attributes" {
            New-ResourceMailbox -DisplayName 'Test Bx' -Alias 'testbx' -Owner 'msims' -Notes 'Testing'
            Assert-MockCalled Set-User -Times 1 -Scope It -Exactly
        }#It
        It "No Sleep Commands" {
            Assert-MockCalled Start-Sleep -Times 0 -Scope Context -Exactly
        }#It
    }#Context
}#Describe
