# Autodoc v1.0
function Get-GraphSkuData {
    <#
        Loads subscribed SKUs from Microsoft Graph and builds lookup tables
        for SKU and service-plan names. Used together with the CSV reference
        data by Resolve-LicenseName / Resolve-PlanName.
    #>
    [CmdletBinding()]
    param()

    Write-Host "Loading subscribed SKUs from Microsoft Graph..."
    $graphSkus = Get-MgSubscribedSku -All -ErrorAction Stop

    $skuPartByGuid = @{}
    $skuPartByName = @{}
    $planById      = @{}
    $planByName    = @{}

    foreach ($sku in $graphSkus) {
        $guid = if ($sku.SkuId) { [string]$sku.SkuId } else { $null }
        if ($guid -and -not $skuPartByGuid.ContainsKey($guid)) { $skuPartByGuid[$guid] = $sku.SkuPartNumber }
        if ($sku.SkuPartNumber -and -not $skuPartByName.ContainsKey($sku.SkuPartNumber)) { $skuPartByName[$sku.SkuPartNumber] = $sku.SkuPartNumber }
        foreach ($plan in $sku.ServicePlans) {
            $planId = if ($plan.ServicePlanId) { [string]$plan.ServicePlanId } else { $null }
            if ($planId -and -not $planById.ContainsKey($planId)) { $planById[$planId] = $plan.ServicePlanName }
            if ($plan.ServicePlanName -and -not $planByName.ContainsKey($plan.ServicePlanName)) { $planByName[$plan.ServicePlanName] = $plan.ServicePlanName }
        }
    }

    @{
        GraphSkus     = $graphSkus
        SkuPartByGuid = $skuPartByGuid
        SkuPartByName = $skuPartByName
        PlanById      = $planById
        PlanByName    = $planByName
    }
}
