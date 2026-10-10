# Autodoc v1.0
function Resolve-AutodocSource {
    <#
        Resolves which export folder a report should build from: the explicit
        -SourceFolder if given and valid, otherwise the newest export under
        ExportsRoot. Returns $null if nothing usable is found.
    #>
    [CmdletBinding()]
    param(
        [string]$SourceFolder,
        [Parameter(Mandatory)][string]$ExportsRoot
    )

    if ($SourceFolder -and (Test-Path $SourceFolder)) { return $SourceFolder }

    if (Test-Path $ExportsRoot) {
        $newest = Get-ChildItem -Path $ExportsRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1
        if ($newest) { return $newest.FullName }
    }
    return $null
}
