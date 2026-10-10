# Autodoc v1.0
function Get-AutodocExportData {
    <#
        Loads one exported dataset from an export folder. Prefers JSON, falls
        back to CSV if the JSON is missing or unreadable (e.g. truncated).
        Always returns an array (possibly empty), never $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExportFolder,
        [Parameter(Mandatory)][string]$BaseName
    )

    $jsonFile = Get-ChildItem -Path (Join-Path $ExportFolder 'rawDataJson') -Filter "$BaseName*.json" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($jsonFile) {
        try   { return @(Get-Content -Path $jsonFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json) }
        catch { Write-Host "  JSON unreadable for $BaseName ($($_.Exception.Message)) - trying CSV." -ForegroundColor Yellow }
    }

    $csvFile = Get-ChildItem -Path (Join-Path $ExportFolder 'rawDataCsv') -Filter "$BaseName*.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($csvFile) {
        try   { return @(Import-Csv -Path $csvFile.FullName -Encoding UTF8) }
        catch { Write-Host "  CSV unreadable for $BaseName ($($_.Exception.Message))." -ForegroundColor Yellow }
    }

    Write-Host "  No data found for $BaseName." -ForegroundColor DarkYellow
    return @()
}
