# Autodoc v1.0
function Get-GroupMembersWithDetails {
    <#
        Returns user members of a group enriched with UPN, department, office
        and manager. Uses the shared user lookup first and falls back to a direct
        Graph query (caching the result back into the lookup). Skips non-user members.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$GroupDisplayName,
        [Parameter(Mandatory)][hashtable]$UserDetailsLookup
    )

    $memberResults = @()
    try { $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop }
    catch {
        Write-Host "  Could not get members for $GroupDisplayName : $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $memberResults
    }

    foreach ($member in $members) {
        $odataType = $member.AdditionalProperties['@odata.type']
        if ($odataType -ne '#microsoft.graph.user') { continue }   # users only

        $userId     = $member.Id
        $userDetail = $null

        if ($UserDetailsLookup.ContainsKey($userId)) {
            $userDetail = $UserDetailsLookup[$userId]
        }
        else {
            # Fallback - fetch directly from Graph if not in the lookup
            try {
                $u = Get-MgUser -UserId $userId -Property "Id,DisplayName,UserPrincipalName,Department,OfficeLocation" -ExpandProperty "manager(`$select=id,userPrincipalName,displayName)" -ErrorAction Stop
                $managerUpn         = $null
                $managerDisplayName = $null
                if ($u.Manager) {
                    if ($u.Manager.AdditionalProperties.ContainsKey('userPrincipalName')) { $managerUpn = $u.Manager.AdditionalProperties['userPrincipalName'] }
                    if ($u.Manager.AdditionalProperties.ContainsKey('displayName'))       { $managerDisplayName = $u.Manager.AdditionalProperties['displayName'] }
                }
                $userDetail = [pscustomobject]@{
                    Id                 = $u.Id
                    DisplayName        = $u.DisplayName
                    UserPrincipalName  = $u.UserPrincipalName
                    Department         = $u.Department
                    OfficeLocation     = $u.OfficeLocation
                    Manager            = $managerUpn
                    ManagerDisplayName = $managerDisplayName
                }
                $UserDetailsLookup[$userId] = $userDetail
            }
            catch { continue }
        }

        $memberResults += [pscustomobject]@{
            GroupId            = $GroupId
            GroupDisplayName   = $GroupDisplayName
            UserId             = $userDetail.Id
            DisplayName        = $userDetail.DisplayName
            UserPrincipalName  = $userDetail.UserPrincipalName
            Department         = $userDetail.Department
            OfficeLocation     = $userDetail.OfficeLocation
            Manager            = $userDetail.Manager
            ManagerDisplayName = $userDetail.ManagerDisplayName
        }
    }
    return $memberResults
}
