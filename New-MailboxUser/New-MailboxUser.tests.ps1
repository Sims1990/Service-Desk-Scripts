$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "New User Creation" -Tags NewUserCreate {
    BeforeEach {
        $ExModule = New-Module -Name Exchange2 -Function "CreatePassword","Set-Mailbox","New-MailboxExportRequest","Get-Mailbox","New-Mailbox","Add-MailboxPermission","Add-ADPermission","Send-BCDConfirmation", "New-MailContact", "Set-MailContact", "Remove-MailContact","Add-ADGroupMember","Get-User", "Set-User" -ScriptBlock {
            Function Set-Mailbox {'Set-Mailbox'};
            Function Get-Mailbox {'Get-Mailbox'};
            Function New-Mailbox{'New-Mailbox'};
            Function Set-User {'Set-User'};
            Function Get-User {'Get-User'};
            Function CreatePassword {"CreatePassword"}}
$ADModule = New-Module -Name ActiveDirectory -Function "Get-ADDomain", "Get-ADUser", "Set-ADAccountExpiration", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount", "Get-ADOrganizationalUnit", "Set-Contact", "Add-ADGroupMember" -ScriptBlock {
            Function New-ADUser {"New-ADUser"};
            Function Get-ADUser {"Get-ADUser"};
            Function Set-ADUser {"Set-ADUser"};
            Function Get-ADDomain {"Get-ADDomain"};
            Function Unlock-ADAccount {"Unlock-ADAccount"};
            Function Set-ADAccountPassword {"Set-ADAccountPassword"};
            Function Remove-ADGroupMember {"Remove-ADGroupMember"};
            Function Add-ADGroupMember {"Add-ADGroupMember"}
            Function Move-ADObject {"Move-ADObject"};
            Function Enable-ADAccount {"Enable-ADAccount"};
            Function Disable-ADAccount {"Disable-ADAccount"};
            Function Get-ADOrganizationalUnit {"Get-ADOrganizationalUnit"};
            Function Set-Contact {'Set-Contact'};
            Function Set-ADAccountExpiration {'Set-ADAccountExpiration'}
}#Scriptblock    
$ExModule | Import-Module
mock New-Mailbox
mock Set-Mailbox
mock Get-User
mock Set-User
mock Get-ADDomain
mock Set-ADAccountExpiration
mock CreatePassword {return $true}
mock Add-Type
mock Start-Sleep
mock Set-ADUser             

    }
    AfterEach {
        Get-Module ActiveDirectory | Remove-Module
        Get-Module Exchange2 | Remove-Module
    }
    Context "Create New User" {  
        It 'Cmdlet runs correctly' {
            New-MailboxUser -Username 'test1' -FirstName 'test' -LastName '1'
        }#it
        It "New User is Created" {
            Assert-MockCalled New-Mailbox -Times 1 -Scope Context -Exactly   
        }#it
        It "Email Policy Disabled" {
            Assert-MockCalled Set-Mailbox -Scope Context
        }#It
        It "Pwd Reset on Logon Disabled" {
            Assert-MockCalled Set-Mailbox -ParameterFilter {@{ResetPasswordONNextLogon="$false"}}
        }#it
        It "Add .net assembly to create password" {
            Assert-MockCalled Add-Type -Times 1 -Exactly -Scope Context
        }
    }#Context
    Context "Setting of Attributes" {
        It "No sleep commands" {
            New-MailboxUser -Username 'test1' -FirstName 'test' -LastName '1'            
            Assert-MockCalled Start-Sleep -Times 0 -Scope Context -ParameterFilter { $Seconds -eq '10'} -Exactly
        }#it
        It "No Active Directory commands" {
            Assert-MockCalled Set-ADUser -Times 0 -Scope Context -Exactly -ParameterFilter { $identity -ne "$null" }
        }#it
        It "writes notes in EAC" {            
            Assert-MockCalled Set-User -Times 1 -Scope Context -ParameterFilter { $identity -ne "$null" }
        }
    }#Context
    
}#Describe