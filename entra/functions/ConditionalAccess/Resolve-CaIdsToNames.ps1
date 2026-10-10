# Autodoc v1.0
function Resolve-CaIdsToNames {
    <#
        Resolves a set of CA condition IDs (users/groups/roles/apps/locations)
        to display names, using an optional special-values map first
        (e.g. "All" -> "All Users") then a lookup table, falling back to the raw ID.
        Returns a ";"-joined string.
    #>
    param(
        [object[]]$Ids,
        [hashtable]$Lookup,
        [hashtable]$SpecialMap = @{}
    )

    if (-not $Ids) { return $null }

    $resolved = foreach ($id in $Ids) {
        if ($null -eq $id) { continue }
        $key = [string]$id
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if     ($SpecialMap -and $SpecialMap.ContainsKey($key)) { $SpecialMap[$key] }
        elseif ($Lookup     -and $Lookup.ContainsKey($key))     { $Lookup[$key] }
        else   { $key }
    }
    return ($resolved -join ";")
}
