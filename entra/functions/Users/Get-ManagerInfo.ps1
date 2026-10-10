# Autodoc v1.0
function Get-ManagerInfo {
    <#
        Returns a user's manager (id / UPN / display name), using the expanded
        manager property when available and falling back to a Graph call.
    #>
    param(
        [Parameter(Mandatory)]$User
    )

    $managerId          = $null
    $managerUpn         = $null
    $managerDisplayName = $null

    if ($User.Manager) {
        $managerId = $User.Manager.Id
        if ($User.Manager.AdditionalProperties.ContainsKey('userPrincipalName')) { $managerUpn = $User.Manager.AdditionalProperties['userPrincipalName'] }
        if ($User.Manager.AdditionalProperties.ContainsKey('displayName'))       { $managerDisplayName = $User.Manager.AdditionalProperties['displayName'] }
    }

    if (-not $managerId -and -not $managerUpn) {
        try {
            $mgr = Get-MgUserManager -UserId $User.Id -ErrorAction Stop
            $managerId = $mgr.Id
            if ($mgr.AdditionalProperties.ContainsKey('userPrincipalName')) { $managerUpn = $mgr.AdditionalProperties['userPrincipalName'] }
            if ($mgr.AdditionalProperties.ContainsKey('displayName'))       { $managerDisplayName = $mgr.AdditionalProperties['displayName'] }
        }
        catch {
            $managerId          = $null
            $managerUpn         = $null
            $managerDisplayName = $null
        }
    }

    @{
        ManagerId                = $managerId
        ManagerUserPrincipalName = $managerUpn
        ManagerDisplayName       = $managerDisplayName
    }
}
