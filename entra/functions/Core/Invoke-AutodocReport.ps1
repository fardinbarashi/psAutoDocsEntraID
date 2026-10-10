# Autodoc v1.0
function Invoke-AutodocReport {
    <#
        STEP 2 wrapper - runs a report builder with a transcript, so report runs
        are logged the same way collections are. The log lands next to the
        collection logs in Entra\files\logs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EntraRoot,
        [Parameter(Mandatory)][ValidateSet('Excel','Visio','Svg','Word','Html')][string]$Kind,
        [string]$SourceFolder,
        [string]$WordLanguage,
        [string]$WordCompanyLogo,
        [string]$WordClientLogo
    )

    $logFolder = Join-Path $EntraRoot 'files\logs'
    if (-not (Test-Path $logFolder)) { New-Item -Path $logFolder -ItemType Directory -Force | Out-Null }

    $stamp   = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $logFile = Join-Path $logFolder "Report-$Kind - $stamp.txt"

    Start-Transcript -Path $logFile -Force | Out-Null
    try {
        Write-Host "Report run: $Kind    $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
        $exportsRoot = Join-Path $EntraRoot 'files\exports'
        $src = Resolve-AutodocSource -SourceFolder $SourceFolder -ExportsRoot $exportsRoot
        if ($src) {
            Write-ExportReadme -ExportFolder $src | Out-Null   # refresh the export README
            # Keep the service-limits config current (in Entra\Config), live from
            # Microsoft when reachable. A refresh failure must never break a report.
            try {
                $limitsConfig = Join-Path (Split-Path (Split-Path $PSScriptRoot)) 'Config/EntraServiceLimits.json'
                Update-EntraServiceLimits -OutputPath $limitsConfig | Out-Null
            }
            catch { Write-Host "  Service limits: refresh skipped ($($_.Exception.Message))." -ForegroundColor DarkYellow }
        }
        $extra = @{}
        if ($src) { $extra['SourceFolder'] = $src }
        $outFolder = $null   # folder to open in Explorer once the build succeeds

        switch ($Kind) {
            'Excel' {
                Build-EntraExcelReport -ExportsRoot $exportsRoot @extra | Out-Null
                if ($src) { $outFolder = Join-Path $src 'Report\excelkpi' }
            }
            'Visio' {
                # Visio files only (the two .vsdx). SVG has its own tab now.
                if (-not $src) {
                    Write-Host "No export folder to build a map from - run a collection first." -ForegroundColor Yellow
                } else {
                    Write-Host "Building Visio maps from: $src" -ForegroundColor Cyan
                    Build-LicenceMapVisio -SourceFolder $src | Out-Null
                    Build-OrgMapVisio     -SourceFolder $src | Out-Null
                    Write-Host ""
                    Write-Host "Done. Visio files: $(Join-Path $src 'Report\visio')" -ForegroundColor Green
                    $outFolder = Join-Path $src 'Report\visio'
                }
            }
            'Svg' {
                # SVG files only (all maps), into the svg folder.
                if (-not $src) {
                    Write-Host "No export folder to build a map from - run a collection first." -ForegroundColor Yellow
                } else {
                    Write-Host "Building SVG maps from: $src" -ForegroundColor Cyan
                    Build-EntraSvgMaps -SourceFolder $src | Out-Null
                    Write-SvgReadme -SvgFolder (Join-Path $src 'Report\svg')                 # README describing the SVG files
                    # PDF copies of every map, into a pdf subfolder.
                    $svgFolder = Join-Path $src 'Report\svg'
                    $pdfFolder = Join-Path $src 'Report\pdf'
                    $pdfMade = 0
                    foreach ($svgFile in Get-ChildItem -Path $svgFolder -Filter '*.svg' -ErrorAction SilentlyContinue) {
                        $pdfPath = Join-Path $pdfFolder ($svgFile.BaseName + '.pdf')
                        if (Convert-SvgToPdf -SvgPath $svgFile.FullName -PdfPath $pdfPath) { $pdfMade++ }
                    }
                    if ($pdfMade -gt 0) { Write-Host "  PDF copies: $pdfMade maps -> $pdfFolder" -ForegroundColor DarkGreen }
                    Write-Host ""
                    Write-Host "Done. SVG files: $(Join-Path $src 'Report\svg')" -ForegroundColor Green
                    $outFolder = Join-Path $src 'Report\svg'
                }
            }
            'Word'  {
                $wx = @{}
                if ($WordLanguage)    { $wx['Language'] = $WordLanguage }
                if ($WordCompanyLogo) { $wx['ConsultantLogo'] = $WordCompanyLogo }
                if ($WordClientLogo)  { $wx['CustomerLogo'] = $WordClientLogo }
                Build-EntraWordReport -ExportsRoot $exportsRoot @extra @wx
                if ($src) { $outFolder = Join-Path $src 'Report\word' }
            }
            'Html'  {
                if (-not $src) {
                    Write-Host "No export folder to build an HTML report from - run a collection first." -ForegroundColor Yellow
                } else {
                    Write-Host "Building HTML report from: $src" -ForegroundColor Cyan
                    Build-EntraSvgMaps    -SourceFolder $src | Out-Null   # ensure the maps exist
                    Build-EntraHtmlReport -SourceFolder $src | Out-Null   # inline them into one HTML
                    $outFolder = Join-Path $src 'Report\html'
                }
            }
        }

        # open the folder holding the freshly built files
        if ($outFolder -and (Test-Path $outFolder)) {
            try { Start-Process -FilePath explorer.exe -ArgumentList "`"$outFolder`"" -ErrorAction SilentlyContinue }
            catch { Write-Host "  (couldn't open $outFolder : $($_.Exception.Message))" -ForegroundColor DarkGray }
        }
    }
    catch {
        Write-Host "Report failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    }
    finally {
        Write-Host "Log: $logFile" -ForegroundColor DarkGray
        Stop-Transcript | Out-Null
    }
}
