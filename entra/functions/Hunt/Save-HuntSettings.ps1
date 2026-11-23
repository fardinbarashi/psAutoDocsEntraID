# Autodoc v1.0
function Get-HuntSettings {
    <#
        Reads Entra\Config\WorkloadIdentityHunt.json as an ordered hashtable
        (empty when the file is missing or unreadable).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return [ordered]@{} }
    try { $h = Read-AutodocJson -Path $Path -AsHashtable; if ($h) { $h } else { [ordered]@{} } }
    catch { Write-Host "  Could not read $Path : $($_.Exception.Message)" -ForegroundColor Yellow; [ordered]@{} }
}

function Save-HuntSettings {
    <#
        Merges -Values into Entra\Config\WorkloadIdentityHunt.json and writes
        it back; keys not in -Values (About, thresholds) are kept as they are.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Values
    )
    $cfg = Get-HuntSettings -Path $Path
    foreach ($k in $Values.Keys) { $cfg[$k] = $Values[$k] }
    $cfg | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
    Write-Host "  Hunt settings saved: $Path" -ForegroundColor DarkGray
}
