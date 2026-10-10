# Autodoc v1.0
function Get-LicenseReferenceData {
    <#
        Loads the Microsoft licensing reference CSV and builds lookup tables
        for resolving SKU and service plan names. (Original, unchanged.)
    #>
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -Path $Path)) { throw "License reference file could not be found: $Path" }
    Write-Host "Loading license reference file: $Path"
    try { $rows = Import-Csv -Path $Path -ErrorAction Stop }
    catch { throw "Could not read the license reference file with Import-Csv. Please check the file. Error: $($_.Exception.Message)" }
    if (-not $rows) { throw "License reference file is empty or could not be parsed." }

    $skuByGuid          = @{}
    $skuByStringId      = @{}
    $planById           = @{}
    $planByName         = @{}
    $skuPlansBySkuGuid  = @{}
    $skuPlansByStringId = @{}

    foreach ($row in $rows) {
        $productDisplayName = $row.Product_Display_Name
        $stringId           = $row.String_Id
        $guid               = $row.GUID
        $servicePlanName    = $row.Service_Plan_Name
        $servicePlanId      = $row.Service_Plan_Id
        $friendlyPlanName   = $row.Service_Plans_Included_Friendly_Names

        if ($guid -and -not $skuByGuid.ContainsKey($guid)) { $skuByGuid[$guid] = $productDisplayName }
        if ($stringId -and -not $skuByStringId.ContainsKey($stringId)) { $skuByStringId[$stringId] = $productDisplayName }
        if ($servicePlanId -and -not $planById.ContainsKey($servicePlanId)) { $planById[$servicePlanId] = if ([string]::IsNullOrWhiteSpace($friendlyPlanName)) { $servicePlanName } else { $friendlyPlanName }}
        if ($servicePlanName -and -not $planByName.ContainsKey($servicePlanName)) { $planByName[$servicePlanName] = if ([string]::IsNullOrWhiteSpace($friendlyPlanName)) { $servicePlanName } else { $friendlyPlanName } }
        if ($guid) {
            if (-not $skuPlansBySkuGuid.ContainsKey($guid)) { $skuPlansBySkuGuid[$guid] = @() }
            $skuPlansBySkuGuid[$guid] += $row }
        if ($stringId) {
            if (-not $skuPlansByStringId.ContainsKey($stringId)) { $skuPlansByStringId[$stringId] = @() }
            $skuPlansByStringId[$stringId] += $row }
    }

    @{
        Rows               = $rows
        SkuByGuid          = $skuByGuid
        SkuByStringId      = $skuByStringId
        PlanById           = $planById
        PlanByName         = $planByName
        SkuPlansBySkuGuid  = $skuPlansBySkuGuid
        SkuPlansByStringId = $skuPlansByStringId
    }
}
