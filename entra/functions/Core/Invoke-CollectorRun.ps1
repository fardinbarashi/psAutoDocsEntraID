# Autodoc v1.0
function Invoke-CollectorRun {
    <#
        Runs the selected collectors (by Key) in registry order, passing the
        shared run context and merging each result into $Context.Collected.
        Unknown keys are ignored; a failing collector is logged by itself and
        does not stop the run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Registry,
        [Parameter(Mandatory)][string[]]$Selected,
        [Parameter(Mandatory)][hashtable]$Context
    )

    # Tell the exporters which datasets to actually write. A collector may
    # produce several datasets; Export-InventoryData writes only the selected
    # ones (see that function). Unset would mean "write everything".
    $Context.Paths.SelectedOutputs = @($Selected)

    $ranOwners = @{}
    foreach ($def in $Registry) {
        if ($Selected -notcontains $def.Key) { continue }

        # Entries can share a collector (Owner); run that collector only once,
        # even when several of its datasets are selected.
        $owner = if ($def.Owner) { $def.Owner } else { $def.Key }
        if ($ranOwners.ContainsKey($owner)) { continue }
        $ranOwners[$owner] = $true

        Write-Host ""
        Write-Host "==================== $($def.Name) ====================" -ForegroundColor Cyan
        $result = & $def.Invoke $Context
        Merge-CollectedData -Target $Context.Collected -Result $result
    }

    return $Context.Collected
}
