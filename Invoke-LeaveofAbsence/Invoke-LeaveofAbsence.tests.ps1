$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"


Describe "Leave of Absence Functions" -Tag LOAFunctions{
    BeforeEach {
        $TestModule = New-Module -Name ActiveDirectory -Function "CreatePassword", "Get-ADOrganizationalUnit", "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount" -ScriptBlock {
                                                                                                Function New-ADUser {"New-ADUser"};
                                                                                                Function Get-ADUser {"Get-ADUser"};
                                                                                                Function Set-ADUser {"Set-ADUser"};
                                                                                                Function Unlock-ADAccount {"Unlock-ADAccount"};
                                                                                                Function Set-ADAccountPassword {"Set-ADAccountPassword"};
                                                                                                Function Remove-ADGroupMember {"Remove-ADGroupMember"};
                                                                                                Function Move-ADObject {"Move-ADObject"};
                                                                                                Function Enable-ADAccount {"Enable-ADAccount"};
                                                                                                Function Disable-ADAccount {"Disable-ADAccount"};
                                                                                                Function CreatePassword {"CreatePassword"}
                                                                                                Function Get-ADOrganizationalUnit{"Get-ADOrganizationalUnit"}
}#Scriptblock
$TestModule | Import-Module

mock -CommandName Unlock-ADAccount -MockWith{@{identity='test1'}} -Verifiable
        mock -CommandName Get-ADUser -MockWith {@{Username='test1';
                                                ObjectGuid='c3ccdb13-a23f-4264-a295-1887c72676aa';
                                                MemberOf='none';
                                                info='Who knows'}
        }#mock
mock -CommandName Set-Aduser -MockWith {@{Username='test1'}
        }#mock
mock -CommandName Set-ADAccountPassword -MockWith{@{Reset=$true}}
mock ConvertTo-SecureString   
mock CreatePassword {return $true} 
        

             
}#BeforeEach
    AfterEach {
            Get-Module ActiveDirectory | Remove-Module
            Get-Module $TestModule | Remove-Module
        }#AfterEach - LOA
    Context "Placing User on Leave of Absence" {          
        
        It "Enable-LOA Successfully Ran" {
            Invoke-LeaveofAbsence 'test1'
        }#it
        It "Grab Current Info Properties" {
            Assert-MockCalled Get-ADUser -Times 1 -Scope Context
        }#it
        It "Adding Description to user" {
            Assert-MockCalled -CommandName Set-ADUser -Times 1 -Scope Context
        }#it
        It "Resetting of Users Password" {
            Assert-MockCalled -CommandName Set-ADAccountPassword
        }#it
    }#Context
}