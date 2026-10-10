# Autodoc v1.0
function Export-InventoryData {
    <#
        Writes a dataset to CSV, JSON and Excel in one call.
        Replaces the repeated Export-Csv / ConvertTo-Json / Export-Excel block.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$BaseName,
        [Parameter(Mandatory)][hashtable]$Paths,
        [string]$WorksheetName,
        $JsonData,          # optional: write a different shape to JSON than to CSV/Excel
        [int]$JsonDepth = 15
    )

    if (-not $WorksheetName) { $WorksheetName = $BaseName }
    if ($null -eq $JsonData) { $JsonData = $Data }
    $fileDate = $Paths.FileDate

    # Per-dataset selection: if the run limited which outputs to write (set by
    # Invoke-CollectorRun from the picker), skip any dataset whose BaseName was
    # not selected. When SelectedOutputs is unset, write everything as before.
    if ($Paths.SelectedOutputs -and @($Paths.SelectedOutputs).Count -gt 0 -and (@($Paths.SelectedOutputs) -notcontains $BaseName)) {
        Write-Host "  Skipping '$BaseName' (not selected)." -ForegroundColor DarkGray
        return
    }

    $csvFile   = Join-Path $Paths.Csv   "$BaseName $fileDate.csv"
    $jsonFile  = Join-Path $Paths.Json  "$BaseName $fileDate.json"
    $excelFile = Join-Path $Paths.Excel "$BaseName $fileDate.xlsx"

    $formats = $Paths.Formats
    if (-not $formats) { $formats = @('Csv','Json','Excel') }   # default: all three

    if ($formats -contains 'Csv')   { $Data     | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8 -Force }
    if ($formats -contains 'Json')  { $JsonData | ConvertTo-Json -Depth $JsonDepth | Out-File -Path $jsonFile -Force }
    if ($formats -contains 'Excel') { $Data     | Export-Excel -Path $excelFile -WorksheetName $WorksheetName -AutoSize -TableName $WorksheetName }
}
