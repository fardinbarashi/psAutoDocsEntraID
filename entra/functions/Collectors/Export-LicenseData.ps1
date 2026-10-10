# Autodoc v1.0
function Export-LicenseData {
    <#
        Section 1.2.2 - Subscribed SKUs / license inventory.
        Requires Organization.Read.All.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths
    )

    $Section = 'Section 1.2.2 : License data'
    try {
        Write-Host "Start $Section"

        Write-Host "Get license data..."
        $skus = Get-MgSubscribedSku -All
        $licenseInventory = $skus | ForEach-Object {
            [pscustomobject]@{
                skuId            = $_.SkuId
                skuPartNumber    = $_.SkuPartNumber
                enabledUnits     = $_.PrepaidUnits.Enabled
                consumedUnits    = $_.ConsumedUnits
                freeUnits        = $_.PrepaidUnits.Enabled - $_.ConsumedUnits
                capabilityStatus = $_.CapabilityStatus
            }
        }

        Write-Host "Export license to files..."
        Export-InventoryData -Data $licenseInventory -BaseName 'LicensesInformation' -Paths $Paths

        Write-Host "Done. License Information collected: $($licenseInventory.Count)" -ForegroundColor DarkGreen
        Write-Host "End $Section" -ForegroundColor Green

        return @{ LicenseInventory = $licenseInventory }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
