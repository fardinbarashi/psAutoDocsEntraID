# Autodoc v1.0
function Get-GroupMemberOfInfo {
    <# Splits a group's memberOf into group / AU / directory-role name lists. #>
    param([string]$GroupId)

    $memberOf = Get-MgGroupMemberOf -GroupId $GroupId -All -ErrorAction SilentlyContinue
    $groupMemberships    = @()
    $administrativeUnits = @()
    $directoryRoles      = @()

    foreach ($item in $memberOf) {
        $odataType   = $item.AdditionalProperties['@odata.type']
        $displayName = $item.AdditionalProperties['displayName']

        if     ($odataType -eq '#microsoft.graph.group')              { $groupMemberships    += $displayName }
        elseif ($odataType -eq '#microsoft.graph.administrativeUnit') { $administrativeUnits += $displayName }
        elseif ($odataType -eq '#microsoft.graph.directoryRole')      { $directoryRoles      += $displayName }
    }

    @{
        GroupMemberships    = ($groupMemberships    | Sort-Object -Unique) -join "; "
        AdministrativeUnits = ($administrativeUnits | Sort-Object -Unique) -join "; "
        MemberOfRoles       = ($directoryRoles      | Sort-Object -Unique) -join "; "
    }
}
