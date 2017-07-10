$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Lockout Status" -Tag 'LockoutStatus' {
    Context "Unlock-ADUser" {
        BeforeEach {
            $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount" -ScriptBlock {
                                                                                                    Function New-ADUser {"New-ADUser"};
                                                                                                    Function Get-ADUser {"Get-ADUser"};
                                                                                                    Function Set-ADUser {"Set-ADUser"};
                                                                                                    Function Unlock-ADAccount {"Unlock-ADAccount"}
 }#Scriptblock
 $ADModule | Import-Module  

         mock -CommandName Get-ADUser -MockWith {Return@{Username='test1'}}
        mock -CommandName Unlock-ADAccount -MockWith {@{identity='test1'}}
        
    }#BeforeEach
        It "Calls Unlock-ADAccount function" {
            Unlock-User 'test1'
            Assert-MockCalled -CommandName Unlock-ADAccount -Times 1 -Scope Context
        }#it
        AfterEach {
            Get-Module ActiveDirectory | Remove-Module
        }
    }#Context - Unlock-ADUser
    ###LockOut Stats###
    Context "Get-Lockoutstatus" {
        BeforeEach {
         $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount" -ScriptBlock {
                                                                                                    Function New-ADUser {"New-ADUser"};
                                                                                                    Function Get-ADUser {"Get-ADUser"};
                                                                                                    Function Set-ADUser {"Set-ADUser"};
                                                                                                    Function Unlock-ADAccount {"Unlock-ADAccount"}
 }#Scriptblock
 $ADModule | Import-Module  
             mock -CommandName Get-ADUser -MockWith {@{Username='test1';CaseID='1231231';NewPassword='P@ssw0rd';ObjectGuid='c3ccdb13-a23f-4264-a295-1887c72676aa'}} -Verifiable

        }#BeforeEach
        It "Should Call Get-AdUser" {
            Get-ADLockoutStatus 'test1'
           Assert-MockCalled -CommandName Get-ADUser -Times 1 -Scope It
        }#it
        AfterEach {
            Get-Module ConsoleFunctionsFull2 | Remove-Module
            Get-Module ActiveDirectory | Remove-Module
        }
    }#Context    
}#Describe
Describe "Distribution List Functions" {
    Context "Create DL" {
     BeforeEach {
         $DummyModule = New-Module -Name Exchange -Function "New-DistributionGroup" -ScriptBlock {
                                                                Function New-DistributionGroup {'New-DistributionGroup'}}
         $DummyModule | Import-Module
        #Import-Module 'C:\Users\msims\Desktop\PoSh Demos\Service Desk\TestBed - ServiceDesk\ConsoleFunctions\Modules\ConsoleFunctionsFull2.psm1'
         mock -CommandName New-DistributionGroup -MockWith {@{Name="DL:"}} -Verifiable
         mock -CommandName Send-ConfirmationEmail -Verifiable

     }#BeforeEach
     It "Should Call New-DistributionGroup" {
         New-DistributionList -DisplayName 'test 1' -Alias 'test1' -CaseID '1231231' -Owner 'bbuilder' -Purpose 'mocktest'  
         Assert-MockCalled -CommandName New-DistributionGroup -Times 1 -Scope it
     }#it
     It "Suppress, but call Confirmation Email" {
         Assert-MockCalled -CommandName Send-ConfirmationEmail -Times 1 -Scope Context
     }
    }#Context 
}#Describe -Distribution List Functions
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
        Remove-CompanyUser -Username test1 -CaseID 1231231 -TermPassword 'P@ssw0rd'
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
            mock -CommandName Send-ConfirmationEmail
        }
    It "Mailbox Removal Ran" {
        Remove-CompanyUserMailbox -Username 'test1'
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
        Assert-MockCalled -CommandName Send-ConfirmationEmail -ExclusiveFilter {$UserTermination}
    }
    }#Context - Removal of mailbox
}#Describe -Termination Functions
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
            mock -CommandName Send-ConfirmationEmail -MockWith{@{LOAEmail='no'}
        }#mock
        mock -CommandName Move-ADObject #-MockWith {@{TargetPath='Disable'}}        
    }#BeforeEach
        AfterEach {
                Get-Module ConsoleFunctionsFull2 | Remove-Module
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
        It "Confirmation Email Sent Out" {
            Assert-MockCalled -CommandName Send-ConfirmationEmail -Times 1 -Scope Context
        }
    }#Context
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
            mock -CommandName Send-ConfirmationEmail -MockWith{@{LOAEmail='no'}
        }#mock
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
    Context "Creating Resource Mailbox Exchange Side" {
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
    Context "Setting of Resource Mailbox Active Directory" {
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
Describe "Exchange Mail Contact" -Tag MailContact {
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
        $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount", "Get-ADOrganizationalUnit", "Set-Contact" -ScriptBlock {
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
                                                                                                    Function Set-Contact {'Set-Contact'}
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
        mock Set-Contact
        mock Set-MailContact
        mock New-MailContact
        mock Remove-MailContact
    }#BeforeEach
    AfterEach {
        Get-Module ActiveDirectory | Remove-Module
        Get-Module Exchange2 | Remove-Module
    }#AfterEach
    Context "Exchange MailContact" {
        It "CaseID Only Accepts Integers" {
        {New-CompanyMailContact -FirstName 'Test' -LastName 'Tester' -CaseID 'gdgdgd' -ExternalEmail 'test1@hv.com' -Location 'US'} | Should throw
        }#it
        It "Creates New MailContact" {
            New-CompanyMailContact -FirstName 'Test' -LastName 'Tester' -CaseID '123123' -ExternalEmail 'test1@hv.com' -Location 'US'
            Assert-MockCalled New-MailContact -Times 1 -Scope It -Exactly
        }#
        It "Hides Contact from GAL" {
            Assert-MockCalled Set-MailContact -Parameter {@{HiddenFromAddressListsEnabled="$true"}} -Scope Context -Exactly
        }#it
        It "Adds Notes to Active Directory" {
            New-CompanyMailContact -FirstName 'Test' -LastName 'Tester' -CaseID '1231231' -ExternalEmail 'test1@hv.com' -Location 'US'
            Assert-MockCalled Set-Contact -Times 2 -Scope It -Exactly
        }#it
        It "Confirmation Email Sent" {
            New-CompanyMailContact -FirstName 'Test' -LastName 'Tester' -CaseID '1231231' -ExternalEmail 'test1@hv.com' -Location 'US'
            Assert-MockCalled Send-ConfirmationEmail -ParameterFilter{@{MailContactEmail='yes'}}
            Assert-MockCalled Send-ConfirmationEmail -Times 1 -Scope It -Exactly
        }#it
    }#Context
    Context "Removal of MailContact" {
        It "Remove MailContact" {
            Remove-CompanyMailContact 'test1'
            Assert-MockCalled Remove-MailContact -Times 1 -Scope Context -Exactly
        }#it
    }#Context
}#Describe
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
        $ADModule = New-Module -Name ActiveDirectory -Function "Get-ADUser", "New-ADUser", "Set-ADUser", "Unlock-ADAccount", "Set-ADAccountPassword", "Add-ADGroupMember", "Set-ADGroupMember", "Remove-ADGroupMember", "Move-ADObject", "Enable-ADAccount", "Disable-ADAccount", "Get-ADOrganizationalUnit", "Set-Contact" -ScriptBlock {
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
                                                                                                    Function Set-Contact {'Set-Contact'};
                                                                                                    Function Add-ADGroupMember {'Add-ADGroupMember'};
                                                                                                    Function Set-ADGroupMember {'Set-ADGroupMember'}
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
    Context "Create New User" {
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
    Context "Setting of Attributes" {
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
        It "Adding Account to Wiki Groups" {
            $SUser
            Assert-MockCalled Add-ADGroupMember -Times 1 -Scope It -Exactly
        }#it
        It "Confrimation Email Sent" {
            $SUser
            Assert-MockCalled Send-ConfirmationEmail -ParameterFilter {@{NewUser='yes'}}
            Assert-MockCalled Send-ConfirmationEmail -Scope it
        }#it
    }#Context
}#Describe

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
Describe "HTMl5 Remote Connection" -Tag HTML5Connect {
    BeforeEach {
mock Start-Process
    }#BeforeEach
    AfterEach {

    }#AfterEach
    Context "Connect-ToComputer" {
        It "Command runs successfully" {
            Connect-ToComputer -ComputerName 'usmokscr300d'
        }#it
        It "Starts Remote Session" {
            Assert-MockCalled Start-Process -Scope Context -Times 1 -Exactly
        }#it
        It "Fails if empty string is given" {
            mock Test-Connection -MockWith {
                return $false
            }
            {Connect-ToComputer "$Test"} | Should throw
        }
    }
}
