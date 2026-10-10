# Autodoc v1.0
function Get-GroupRoleAssignments {
    <# Directory role assignments where the group is the principal, resolved to role names. #>
    param([string]$GroupId)

    $assignments = Get-MgRoleManagementDirectoryRoleAssignment -All -Filter "principalId eq '$GroupId'" -ErrorAction SilentlyContinue
    if (-not $assignments) { return "" }

    $roleNames = foreach ($assignment in $assignments) {
        try {
            $roleDef = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $assignment.RoleDefinitionId -ErrorAction Stop
            if ($roleDef.DisplayName) { $roleDef.DisplayName } else { $assignment.RoleDefinitionId }
        }
        catch { $assignment.RoleDefinitionId }
    }
    return ($roleNames | Sort-Object -Unique) -join "; "
}
