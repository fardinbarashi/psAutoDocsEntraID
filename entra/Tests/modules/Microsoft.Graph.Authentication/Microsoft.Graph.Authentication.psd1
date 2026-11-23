@{
    RootModule        = 'Microsoft.Graph.Authentication.psm1'
    ModuleVersion     = '99.0.0'
    GUID              = '5a6a1f1e-1111-4222-8333-444455556666'
    FunctionsToExport = @('Connect-MgGraph', 'Get-MgContext', 'Disconnect-MgGraph', 'Invoke-MgGraphRequest')
}
