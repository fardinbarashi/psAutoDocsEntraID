# Autodoc v1.0
function Connect-EntraGraph {
    <#
        Connects to Microsoft Graph with the given scopes.
        Throws if the connection fails so the caller can stop cleanly.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Scopes
    )

    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $Scopes | Out-Null

    $context = Get-MgContext
    if ($null -eq $context) {
        throw "The connection to Microsoft Graph failed."
    }

    Write-Host "Connected to Microsoft Graph. Tenant: $($context.TenantId)" -ForegroundColor DarkGreen
    return $context
}
