# Autodoc v1.0
function New-AutodocPaths {
    <#
        Builds every path the script uses for a single run and creates the
        folder structure. Returns a hashtable so nothing relies on global state.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root
    )

    $runStamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'   # per-run export folder
    $fileDate = Get-Date -Format 'yyyy-MM-dd'            # used in file names

    $filesRoot  = Join-Path $Root      'files'
    $exportRoot = Join-Path $filesRoot "exports\$runStamp"
    $summary    = Join-Path $exportRoot 'rawDataTenantSummary'

    $paths = @{
        Root     = $Root
        RunStamp = $runStamp
        FileDate = $fileDate

        Files = $filesRoot
        Logs  = Join-Path $filesRoot 'logs'
        Cache = Join-Path $filesRoot 'cache'

        Export  = $exportRoot
        Csv     = Join-Path $exportRoot 'rawDataCsv'
        Json    = Join-Path $exportRoot 'rawDataJson'
        Excel   = Join-Path $exportRoot 'rawDataExcel'
        Summary = $summary

        # Reference / settings files
        LicenseReferenceFile = Join-Path $filesRoot 'cache\ms-licensing-reference.csv'
        MsLearnConfig        = Join-Path $Root 'Config\AzureDocs.json'
    }

    # Consolidated summary output files
    $paths.SummaryExcel = Join-Path $summary "TenantSummary $fileDate.xlsx"
    $paths.SummaryJson  = Join-Path $summary "TenantSummary $fileDate.json"
    $paths.SummaryCsv   = Join-Path $summary "TenantSummary $fileDate.csv"

    # Create the folder structure
    foreach ($folder in @($paths.Logs, $paths.Cache, $paths.Export,
                          $paths.Csv, $paths.Json, $paths.Excel, $paths.Summary)) {
        if (-not (Test-Path -Path $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }
    }

    # Transcript file (depends on Logs existing)
    $paths.TranscriptFile = Join-Path $paths.Logs "AutodocEntra - $runStamp.txt"

    return $paths
}
