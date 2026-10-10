# Autodoc v1.0
function Resolve-PrincipalName {
    <#
        Resolves a principal ID to a display name using the supplied lookups.
        -Lookups is a hashtable: @{ User = <ht>; Group = <ht>; Sp = <ht> }.
        Falls back to the raw ID if not found.
    #>
    param(
        [string]$PrincipalId,
        [Parameter(Mandatory)][hashtable]$Lookups
    )

    if ([string]::IsNullOrWhiteSpace($PrincipalId)) { return $null }
    if ($Lookups.User.ContainsKey($PrincipalId))  { return $Lookups.User[$PrincipalId] }
    if ($Lookups.Group.ContainsKey($PrincipalId)) { return $Lookups.Group[$PrincipalId] }
    if ($Lookups.Sp.ContainsKey($PrincipalId))    { return $Lookups.Sp[$PrincipalId] }
    return $PrincipalId
}
