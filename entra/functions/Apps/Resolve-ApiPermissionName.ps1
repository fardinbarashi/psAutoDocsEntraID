# Autodoc v1.0
function Resolve-ApiPermissionName {
    <#
        Turns an application's RequiredResourceAccess into readable text, e.g.
          "Microsoft Graph: Directory.Read.All (Application), User.Read (Delegated)"

        Falls back to the raw GUID for any permission the lookup doesn't cover -
        that happens when the resource has no service principal in this tenant.
        Returns @{ Text = '...'; Count = <n> }.
    #>
    [CmdletBinding()]
    param(
        $RequiredResourceAccess,
        [Parameter(Mandatory)][hashtable]$Lookup
    )

    $parts = @()
    $count = 0

    foreach ($resource in @($RequiredResourceAccess)) {
        $appId = [string]$resource.ResourceAppId
        if (-not $appId) { continue }

        $resName = if ($Lookup.ResourceName.ContainsKey($appId)) { $Lookup.ResourceName[$appId] } else { $appId }

        $perms = foreach ($access in @($resource.ResourceAccess)) {
            $count++
            $key  = "$appId/$([string]$access.Id)"
            $name = if ($Lookup.Permission.ContainsKey($key)) { $Lookup.Permission[$key] } else { [string]$access.Id }
            $kind = switch ("$($access.Type)") {
                'Role'  { 'Application' }
                'Scope' { 'Delegated' }
                default { "$($access.Type)" }
            }
            "$name ($kind)"
        }

        if ($perms) { $parts += "${resName}: $($perms -join ', ')" }
    }

    @{ Text = ($parts -join ' | '); Count = $count }
}
