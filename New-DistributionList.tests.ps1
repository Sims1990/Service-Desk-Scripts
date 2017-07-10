$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Distribution List Functions" {
    Context "Create DL" {
     BeforeEach {
         $DummyModule = New-Module -Name Exchange -Function "New-DistributionGroup" -ScriptBlock {
                                                                Function New-DistributionGroup {'New-DistributionGroup'}}
         $DummyModule | Import-Module
         mock -CommandName New-DistributionList -MockWith {@{Name="DL:"}} -Verifiable

     }#BeforeEach
     It "Should Call New-DistributionGroup" {
         New-DistributionList -DisplayName 'test 1' -Alias 'test1' -CaseID '1231231' -Owner 'bbuilder' -Purpose 'mocktest'  
         Assert-MockCalled -CommandName New-DistributionList -Times 1 -Scope it
     }#it
    }#Context 
}#Describe -Distribution List Functions