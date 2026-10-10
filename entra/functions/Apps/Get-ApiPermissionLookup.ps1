# Autodoc v1.0
function Get-ApiPermissionLookup {
    <#
        Builds the table needed to turn API permission GUIDs into names.

        An app registration stores only identifiers: resourceAppId plus a list
        of permission GUIDs. The human-readable names ("User.Read.All",
        "Directory.Read.All") live on the RESOURCE's service principal - for
        Graph permissions that is the Microsoft Graph service principal - in two
        collections:
          AppRoles                 application permissions  (ResourceAccess type 'Role')
          Oauth2PermissionScopes   delegated permissions    (ResourceAccess type 'Scope')

        Returns:
          @{ ResourceName = @{ appId -> display name }
             Permission   = @{ "appId/permissionGuid" -> permission name } }
    #>
    [CmdletBinding()]
    param(
        $ServicePrincipals
    )

    $resourceName = @{}
    $permission   = @{}

    foreach ($sp in @($ServicePrincipals)) {
        $appId = [string]$sp.AppId
        if (-not $appId) { continue }
        if (-not $resourceName.ContainsKey($appId)) { $resourceName[$appId] = $sp.DisplayName }

        foreach ($role in @($sp.AppRoles)) {
            if ($role.Id) { $permission["$appId/$([string]$role.Id)"] = $role.Value }
        }
        foreach ($scope in @($sp.Oauth2PermissionScopes)) {
            if ($scope.Id) { $permission["$appId/$([string]$scope.Id)"] = $scope.Value }
        }
    }

    @{ ResourceName = $resourceName; Permission = $permission }
}
