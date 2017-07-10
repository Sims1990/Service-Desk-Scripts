$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Leave of Absence Functions" -Tag LOAFunctions{
    Context "Returning User from Leave of Absence" {
        BeforeEach {
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

            mock -CommandName Unlock-ADAccount -MockWith{@{identity='test1'}} -Verifiable
            mock -CommandName Get-ADUser -MockWith {@{Username='test1';
                                                    CaseID='1231231';
                                                    NewPassword='P@ssw0rd';
                                                    ObjectGuid='c3ccdb13-a23f-4264-a295-1887c72676aa';
                                                    Office='MO999'}
            }#mock
            mock -CommandName Set-Aduser -MockWith {@{Username='test1'}
            }#mock
            Mock -CommandName Set-ADAccountPassword #mock

            mock -CommandName Enable-ADAccount
        mock -CommandName Move-ADObject #-MockWith {@{TargetPath='Disable'}}        
    }#BeforeEach
        AfterEach {
                Get-Module ConsoleFunctionsFull2 | Remove-Module
                Get-Module ActiveDirectory | Remove-Module
            }#AfterEach - LOA
        It "Running of Return from LOA Function" {
            Invoke-UserLOAReturn -Username test1 -CaseID 1231231 -NewPassword 'P@ssw0rd'
        }#it
        It "CaseID only accepts integers" {
            {Invoke-UserLOAReturn -Username 'test1' -CaseID 'fjfjfjfj' -NewPassword 'P@ssw0rd'} | Should throw
        }
        It "Office Variable created, CurrInfo Variable created and GUID pulled" {
            Assert-MockCalled -CommandName Get-ADUser -Times 3 -Scope Context
        }#it
    }#Context
}#Describe