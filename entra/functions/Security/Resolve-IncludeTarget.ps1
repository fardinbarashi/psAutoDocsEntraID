# Autodoc v1.0
function Resolve-IncludeTarget {
    <#
        Interprets an authentication-method includeTargets collection into a
        scope (All / Selected / None / Unknown) and an optional target group id.
    #>
    param([array]$Targets)

    if (-not $Targets -or $Targets.Count -eq 0) { return [pscustomobject]@{ Scope = "None"; GroupId = $null } }

    $allUsersTarget = $Targets | Where-Object { $_.targetType -eq "allUsers" }
    if ($allUsersTarget) { return [pscustomobject]@{ Scope = "All"; GroupId = $null } }

    $groupTarget = $Targets | Where-Object { $_.targetType -eq "group" } | Select-Object -First 1
    if ($groupTarget) { return [pscustomobject]@{ Scope = "Selected"; GroupId = $groupTarget.id } }

    return [pscustomobject]@{ Scope = "Unknown"; GroupId = $null }
}
