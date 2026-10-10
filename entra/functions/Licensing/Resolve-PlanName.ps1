# Autodoc v1.0
function Resolve-PlanName {
    <# Resolves a service plan ID to a friendly name, optionally dropping RETIRED plans. (Original.) #>
    param(
        [Parameter(Mandatory)]$AssignedPlan,
        [Parameter(Mandatory)][hashtable]$LicenseRef,
        [Parameter(Mandatory)][hashtable]$GraphRef,
        [switch]$ExcludeRetired
    )

    $planId = $null
    if ($AssignedPlan.ServicePlanId) { $planId = [string]$AssignedPlan.ServicePlanId }
    $resolvedName = $null
    if ($planId -and $LicenseRef.PlanById.ContainsKey($planId))     { $resolvedName = $LicenseRef.PlanById[$planId] }
    elseif ($planId -and $GraphRef.PlanById.ContainsKey($planId))   { $resolvedName = $GraphRef.PlanById[$planId] }
    elseif ($planId) { $resolvedName = $planId }
    if ($ExcludeRetired -and $resolvedName -match '^RETIRED') { return $null }
    return $resolvedName
}
