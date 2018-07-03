$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe 'Return From Leave of Absence'{
    BeforeEach {
        $TestModule = New-Module -Name ActiveDirectory -Function "CreatePassword", "Get-ADOrganizationalUnit", "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Enable-ADAccount", "Disable-ADAccount" -ScriptBlock {
            Function New-ADUser {"New-ADUser"};
            Function Get-ADUser {"Get-ADUser"};
            Function Set-ADUser {"Set-ADUser"};
            Function Unlock-ADAccount {"Unlock-ADAccount"};
            Function Set-ADAccountPassword {"Set-ADAccountPassword"};
            Function Remove-ADGroupMember {"Remove-ADGroupMember"};
            Function Enable-ADAccount {"Enable-ADAccount"};
            Function Disable-ADAccount {"Disable-ADAccount"};
            Function CreatePassword {"CreatePassword"}
            Function Get-ADOrganizationalUnit{"Get-ADOrganizationalUnit"}
}#Scriptblock
$TestModule | Import-Module

        mock Unlock-ADAccount -MockWith {@{identity='test1'}} -Verifiable
        mock Get-ADUser -MockWith {@{Username='test1';
                                    ObjectGuid='c3ccdb13-a23f-4264-a295-1887c72676aa';
                                    Office='MO999'}
                            }#mock
        Mock Set-ADUser -MockWith {@{Username='test1'}}
        Mock Set-ADAccountPassword -MockWith{@{Reset=$true}}
        Mock Enable-ADAccount
        Mock CreatePassword {return $true}
        Mock Write-Output
        Mock New-Object
        Mock Convertto-SecureString
    }#BeforeEach
    AfterEach {
        Get-Module $TestModule | Remove-Module
    }
    Context "Returning User from Leave of Absence" {
        
        It "Running of Return from LOA Function" {
            Invoke-LeaveofAbsenceReturn -Username test1 
        }#it
        It "CurrInfo Variable created and GUID pulled" {
            Assert-MockCalled -CommandName Get-ADUser -Times 1 -Scope Context
        }#it
        It 'Creates temp password' {
            Assert-MockCalled CreatePassword
            Assert-MockCalled Set-ADAccountPassword 
        }#it
        It 'Enables Account in Active Directory' {
            Assert-MockCalled Enable-ADAccount -Times 1 -Exactly
        }#it
        It 'Sets Active Directory notes' {
            Assert-MockCalled Set-ADUser
        }#it
        It 'PSObject created and written as output' {
            Assert-MockCalled New-Object -Times 1 -Exactly
            Assert-MockCalled Write-Output
        }
    }#Context
}#Describe
