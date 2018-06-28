$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Distribution List Functions" -Tag DistroList {
    BeforeEach {
        
                                                               $ExModule = New-Module -Name Exchange2 -Function "New-DistributionGroup","Set-Mailbox","New-MailboxExportRequest","Get-Mailbox","New-Mailbox","Add-MailboxPermission","Add-ADPermission","Get-User","Set-User" -ScriptBlock {
                                                                   
                                                                   Function Set-Mailbox {'Set-Mailbox'};
                                                                   Function New-DistributionGroup {'New-DistributionGroup'};
                                                                   Function New-MailboxExportRequest {'New-MailboxExportRequest'};
                                                                   Function Get-Mailbox {'Get-Mailbox'};
                                                                   Function New-Mailbox{'New-Mailbox'};
                                                                   Function Add-MailboxPermission{'Add-MailboxPermission'};
                                                                   Function Add-ADPermission{'Add-ADPermission'};
                                                                   Function Get-User{'Get-User'};
                                                                   Function Set-User{'Set-User'}}
        
        $ExModule | Import-Module
        mock -CommandName New-DistributionGroup 
        mock Get-Date
        mock Get-User
        mock Export-Csv {return $null}

    }#BeforeEach
    AfterEach {
        $ExModule | Remove-Module
    }
    Context "Create DL" {
     
     It "Should Call New-DistributionGroup" {
         Create-DistributionList -DisplayName 'test 1' -Alias 'test1' -Owner 'bbuilder' -Purpose 'mocktest'  
         Assert-MockCalled -CommandName New-DistributionGroup -Times 1 -Scope Context
     }#it
     It "Gets the owners first and last name" {
         Assert-MockCalled Get-User -Times 2 -Scope Context -Exactly
     }
    }#Context 

    Context "Addressing issues with Distribution entries" {
        It "Confirms display name is in the correct format" {
            Create-DistributionList -DisplayName 'test 1' -Alias 'test1' -Owner 'bbuilder' -Purpose 'mocktest'  
            {Create-DistributionList -DisplayName 'DL: test 1' -Alias 'test1' -Owner 'bbuilder' -Purpose 'mocktest'} | Should throw 

        }#it
        It "Confirms alias is in the correct format" {
            Create-DistributionList -DisplayName 'test 1' -Alias 'test1' -Owner 'bbuilder' -Purpose 'mocktest'  
            {Create-DistributionList -DisplayName 'test 1' -Alias 'DL_test1' -Owner 'bbuilder' -Purpose 'mocktest'} | Should throw  

        }#it
    }#Context
}#Describe