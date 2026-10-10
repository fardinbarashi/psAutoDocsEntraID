# Autodoc v1.0
function Resolve-PrincipalType {
    <#
        Classifies a principal ID as User / Group / ServicePrincipal using the
        supplied lookups (@{ User; Group; Sp }), or "Unknown" if not found.
    #>
    param(
        [string]$PrincipalId,
        [Parameter(Mandatory)][hashtable]$Lookups
    )

    if ([string]::IsNullOrWhiteSpace($PrincipalId)) { return $null }
    if ($Lookups.User.ContainsKey($PrincipalId))  { return "User" }
    if ($Lookups.Group.ContainsKey($PrincipalId)) { return "Group" }
    if ($Lookups.Sp.ContainsKey($PrincipalId))    { return "ServicePrincipal" }
    return "Unknown"
}
