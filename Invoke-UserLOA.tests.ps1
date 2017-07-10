$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"


Describe "Leave of Absence Functions" -Tag LOAFunctions{
    Context "Placing User on Leave of Absence" {
        BeforeEach {
            $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount" -ScriptBlock {
                                                                                                    Function New-ADUser {"New-ADUser"};
                                                                                                    Function Get-ADUser {"Get-ADUser"};
                                                                                                    Function Set-ADUser {"Set-ADUser"};
                                                                                                    Function Unlock-ADAccount {"Unlock-ADAccount"};
                                                                                                    Function Set-ADAccountPassword {"Set-ADAccountPassword"};
                                                                                                    Function Remove-ADGroupMember {"Remove-ADGroupMember"};
                                                                                                    Function Move-ADObject {"Move-ADObject"};
                                                                                                    Function Enable-ADAccount {"Enable-ADAccount"};
                                                                                                    Function Disable-ADAccount {"Disable-ADAccount"}
 }#Scriptblock
 $ADModule | Import-Module  

            mock -CommandName Unlock-ADAccount -MockWith{@{identity='test1'}} -Verifiable
            mock -CommandName Get-ADUser -MockWith {@{Username='test1';
                                                    CaseID='1231231';
                                                    NewPassword='P@ssw0rd';
                                                    ObjectGuid='c3ccdb13-a23f-4264-a295-1887c72676aa';
                                                    MemberOf='none';
                                                    info='Who knows'}
            }#mock
            mock -CommandName Set-Aduser -MockWith {@{Username='test1'}
            }#mock
            Mock -CommandName Set-ADAccountPassword

            mock -CommandName Move-ADObject         
    }#BeforeEach
        AfterEach {
                Get-Module ActiveDirectory | Remove-Module
            }#AfterEach - LOA
        It "Enable-LOA Successfully Ran" {
            Invoke-UserLOA -Username 'test1' -CaseID '1231231' -NewPassword 'P@ssw0rd'
        }#it
        It "Grab Current Info Properties" {
            Assert-MockCalled Get-AdUser -Times 1 -Scope Context
        }#it
        It "Adding Description to user" {
            Assert-MockCalled -CommandName Set-ADUser -Times 1 -Scope Context
        }#it
        It "Resetting of Users Password" {
            Assert-MockCalled -CommandName Set-ADAccountPassword
        }#it
        It "Move AD Account to Disabled OU" {
            Assert-MockCalled -CommandName Move-ADObject  -Times 1 -Scope Context
        }#it
    }#Context
}#Describe