@{
    RootModule = 'Get-OneFileFromZip.psm1'
    ModuleVersion = '1.1.1.0'
    GUID = '4883c45e-d997-4c7d-b835-7587bbfc8e3f'
    Author = 'Roy Ashbrook'
    CompanyName = 'ashbrook.io'
    Copyright = '(c) 2021 royashbrook. All rights reserved.'
    Description = 'Returns one file from a zip. Uses Streamreader ReadToEnd method to return results.'
    FunctionsToExport = @('Get-OneFileFromZip','Get-SheetByName')
    AliasesToExport = @()
    CmdletsToExport = @()
    VariablesToExport = @()
}