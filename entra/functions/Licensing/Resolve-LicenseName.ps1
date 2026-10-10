# Autodoc v1.0
function Resolve-LicenseName {
    <# Resolves a SKU GUID to a friendly name, falling back to the GUID. (Original.) #>
    param(
        [Parameter(Mandatory)]$AssignedLicense,
        [Parameter(Mandatory)][hashtable]$LicenseRef,
        [Parameter(Mandatory)][hashtable]$GraphRef
    )

    $skuGuid = if ($AssignedLicense.SkuId) { [string]$AssignedLicense.SkuId } else { $null }
    if ($skuGuid -and $LicenseRef.SkuByGuid.ContainsKey($skuGuid))    { return $LicenseRef.SkuByGuid[$skuGuid] }
    if ($skuGuid -and $GraphRef.SkuPartByGuid.ContainsKey($skuGuid))  { return $GraphRef.SkuPartByGuid[$skuGuid] }
    if ($skuGuid) { return $skuGuid }
    return $null
}
