# Autodoc v1.0
<#
    Autodoc
    ------------------------------------------------------------------
    Launcher for the documentation modules. Product logic lives under a
    per-product folder next to this script

    Run modes:
      .\Autodoc.ps1                        # shows the GUI ( collect / Report)
      .\Autodoc.ps1 -All                   # collect every section, no GUI
      .\Autodoc.ps1 -Collectors UserInformation,ConditionalAccess    # collect only these datasets
      .\Autodoc.ps1 -All -Formats Excel    # collect everything, Excel output only
      .\Autodoc.ps1 -Mode Report           # step 2: build Excel charts from latest export
      .\Autodoc.ps1 -Mode Hunt             # workload identity hunt into the latest export
      .\Autodoc.ps1 -Mode Hunt -ExportFolder <path>   # ...into this export (new folder if none exists)

    Author : Fardin Barashi
    GitHub : https://github.com/fardinbarashi/psAutoDocs
    Requires -Version 7.4
#>

param(
    [ValidateSet('Collect','Report','Hunt')][string]$Mode,   # bypass the GUI
    [string[]]$Collectors,   # collect only these registry keys
    [string[]]$Formats,      # any of: Csv, Json, Excel (default: all three)
    [switch]$All,            # collect everything without the GUI
    [switch]$NoGui,          # headless full collect (same as -All)
    [string]$ExportFolder    # -Mode Hunt: export folder to write into (default: latest)
)

$ErrorActionPreference = 'Continue'
$ScriptRoot = $PSScriptRoot
$EntraRoot  = Join-Path $ScriptRoot 'Entra'
$ExportsRoot = Join-Path $EntraRoot 'files\exports'

# --- Unblock files downloaded from the internet (GitHub zip) --------------
# Windows marks every file in a downloaded zip as "from the internet", and the
# RemoteSigned execution policy then asks before each one is loaded. Removing
# that mark once here stops the prompts for the rest of the files.
if ($env:OS -eq 'Windows_NT') {   # $IsWindows does not exist in Windows PowerShell 5.1
    Get-ChildItem -Path $ScriptRoot -Recurse -File -Include '*.ps1','*.psm1','*.psd1' |
        Unblock-File -ErrorAction SilentlyContinue
}

# --- Load configuration + functions from the Entra module -----------------
$Settings = Import-PowerShellDataFile (Join-Path $EntraRoot 'Config\Settings.psd1')
. (Join-Path $EntraRoot 'Config\Paths.ps1')
Get-ChildItem (Join-Path $EntraRoot 'Functions') -Recurse -Filter '*.ps1' |
    ForEach-Object { . $_.FullName }

$registry        = Get-CollectorRegistry
$selectedFormats = if ($Formats) { @($Formats) } else { @('Csv','Json','Excel') }

# --- Dispatcher: turns a GUI request (or a parameter set) into work --------
$runAction = {
    param($request)

    switch ($request.Action) {
        'ReportDocs' {
            $fmts = @($request.Formats)
            if (-not $fmts -or $fmts.Count -eq 0) {
                Write-Host "No formats selected - tick at least one (Excel, SVG, HTML, Word, Security)." -ForegroundColor Yellow
            } else {
                foreach ($fmt in $fmts) {
                    Invoke-AutodocReport -EntraRoot $EntraRoot -Kind $fmt -SourceFolder $request.SourceFolder `
                        -WordLanguage $request.WordLanguage -WordCompanyLogo $request.WordCompanyLogo -WordClientLogo $request.WordClientLogo
                }
            }
        }
        'Hunt' {
            Invoke-AutodocHunt -EntraRoot $EntraRoot -Settings $Settings -TargetFolder $request.TargetFolder `
                -NewFolder:([bool]$request.NewFolder) -HuntSettings $request.HuntSettings -BuildReport:([bool]$request.BuildReport) `
                -Language $(if ($request.Language) { $request.Language } else { 'sv' }) `
                -ConsultantLogo $request.WordCompanyLogo -CustomerLogo $request.WordClientLogo | Out-Null
        }
        default {
            Invoke-AutodocCollect -EntraRoot $EntraRoot -Settings $Settings -Registry $registry `
                                  -Selected $request.Collectors -Formats $request.Formats
        }
    }
}

# --- Headless modes (no GUI) ----------------------------------------------
if     ($Mode -eq 'Hunt') {
    $target = Resolve-AutodocSource -SourceFolder $ExportFolder -ExportsRoot $ExportsRoot
    & $runAction ([pscustomobject]@{ Action = 'Hunt'; TargetFolder = $target; NewFolder = (-not $target) }); return
}
elseif ($Mode -eq 'Report') { & $runAction ([pscustomobject]@{ Action = 'ReportDocs'; Formats = @('Excel') }); return }
elseif ($Collectors)        { & $runAction ([pscustomobject]@{ Action = 'Collect'; Collectors = @($Collectors);   Formats = $selectedFormats }); return }
elseif ($All -or $NoGui -or $Mode -eq 'Collect') {
                              & $runAction ([pscustomobject]@{ Action = 'Collect'; Collectors = @($registry.Key); Formats = $selectedFormats }); return }

# --- GUI mode: window stays open, Run calls back into $runAction ----------
try {
    Show-CollectorPicker -Registry $registry `
                         -XamlPath (Join-Path $ScriptRoot 'Config\CollectorPicker.xaml') `
                         -ExportsRoot $ExportsRoot `
                         -OnRun $runAction
}
catch {
    Write-Host "GUI unavailable ($($_.Exception.Message)) - collecting all sections." -ForegroundColor Yellow
    & $runAction ([pscustomobject]@{ Action = 'Collect'; Collectors = @($registry.Key); Formats = $selectedFormats })
}
