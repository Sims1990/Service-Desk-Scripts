$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Termination Functions" -Tag TerminationFunctions {
    Context "Remove-User" {
        BeforeEach {
            $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject" -ScriptBlock {
                                                                                                    Function New-ADUser {"New-ADUser"};
                                                                                                    Function Get-ADUser {"Get-ADUser"};
                                                                                                    Function Set-ADUser {"Set-ADUser"};
                                                                                                    Function Unlock-ADAccount {"Unlock-ADAccount"};
                                                                                                    Function Set-ADAccountPassword {"Set-ADAccountPassword"};
                                                                                                    Function Remove-ADGroupMember {"Remove-ADGroupMember"};
                                                                                                    Function Move-ADObject {"Move-ADObject"}
 }#Scriptblock

 $ADModule | Import-Module  

            mock -CommandName Get-ADUser -MockWith {@{Username='test1';
                                                        CaseID='1231231';
                                                    NewPassword='P@ssw0rd';
                                                    ObjectGuid='c3ccdb13-a23f-4264-a295-1887c72676aa';
                                                    MemberOf='none'}
            }#mock
            mock -CommandName Set-Aduser -MockWith {@{Username='test1'}
            }#mock
            Mock -CommandName Set-ADAccountPassword #-MockWith {@{identity='test1'}
       # }#mock
            mock -CommandName Remove-ADGroupMember -MockWith {@{Members='test1'}

            }#Mock
         
    }#BeforeEach
    It "Remove-User CMD called" {
        Remove-User -Username test1 -CaseID 1231231 -TermPassword 'P@ssw0rd'
    }#it
    It "Get-ADUser Called" {
        Assert-MockCalled -CommandName Get-AdUser -Times 3 -Scope Context
    }#it
    It "Set-ADUser Called" {
        Assert-MockCalled -CommandName Set-AdUser -Times 1 -Scope Context
    }#it
    It "Set-ADAccountPassword Called" {
        Assert-MockCalled -CommandName Set-ADAccountPassword -Times 1 -Scope Context
    }#it
    It "Remove-ADGroupMember Called" {
        Assert-MockCalled -CommandName Remove-ADGroupMember -Times 1 -Scope Context
    }#it 
    AfterEach {
            Get-Module Exchange | Remove-Module
            Get-Module ActiveDirectory | Remove-Module
        }#AfterEach -
    }#Context - Termination Functions
    Context "Remove-UserMailbox" {
        BeforeEach {
            $ExModule = New-Module -Name Exchange2 -Function "Set-Mailbox","New-MailboxExportRequest","Get-Mailbox" -ScriptBlock {
                                                                Function Set-Mailbox {'Set-Mailbox'};
                                                                Function New-MailboxExportRequest {'New-MailboxExportRequest'};
                                                                Function Get-Mailbox {'Get-Mailbox'}}
                            
            $ExModule | Import-Module                                                    

            mock -CommandName Set-Mailbox -MockWith {@{identity='test1'}}
            mock -CommandName Out-File -MockWith{@{FilePath='\\FilerServer\c$'}}
            mock -CommandName New-MailboxExportRequest 
            mock -CommandName Get-Mailbox
            mock -CommandName Send-ConfirmationEmail -MockWith{@{TerminationEmail='yes'}}
        }
    It "Mailbox Removal Ran" {
        Remove-UserMailbox -Username 'test1'
    }#it
    It "Forwarding address removed and hidden from GAL" {
        Assert-MockCalled -CommandName Set-Mailbox -Times 1 -Scope Context
    }#it
    It "Invoke Request for Mailbox Export" {
        Assert-MockCalled -CommandName New-MailboxExportRequest -Times 1 -Scope Context
    }#it
    It "Pull alias of the user" {
        Assert-MockCalled -CommandName Get-Mailbox
    }#it
    It "Write to Text List for removal" {
        Assert-MockCalled -CommandName Out-File -Times 1 -Scope Context -ExclusiveFilter {$Append}
    }#it
    It "Confirmation Email Sent" {
        Assert-MockCalled -CommandName Send-ConfirmationEmail -ExclusiveFilter {$TerminationEmail -eq 'yes'}
    }
    }#Context - Removal of mailbox
}#Describe -Termination Functions