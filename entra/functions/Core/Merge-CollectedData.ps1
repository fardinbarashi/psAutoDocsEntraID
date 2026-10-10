# Autodoc v1.0
function Merge-CollectedData {
    <#
        Merges a collector's returned hashtable of named datasets into the
        shared $Target dictionary. Ignores $null results (a failed section),
        so the summary can still run with whatever was collected.
        $Target is a reference type, so edits are visible to the caller.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Target,
        $Result
    )

    if ($Result -is [System.Collections.IDictionary]) {
        foreach ($key in $Result.Keys) { $Target[$key] = $Result[$key] }
    }
}
