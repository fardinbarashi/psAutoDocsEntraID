# Autodoc v1.0
function Get-GroupApplications {
    <# App role assignments granted to the group, as resource display names. #>
    param([string]$GroupId)

    $apps = Get-MgGroupAppRoleAssignment -GroupId $GroupId -All -ErrorAction SilentlyContinue
    if (-not $apps) { return "" }

    $appNames = foreach ($app in $apps) {
        if     ($app.ResourceDisplayName) { $app.ResourceDisplayName }
        elseif ($app.ResourceId)          { [string]$app.ResourceId }
    }
    return ($appNames | Sort-Object -Unique) -join "; "
}
