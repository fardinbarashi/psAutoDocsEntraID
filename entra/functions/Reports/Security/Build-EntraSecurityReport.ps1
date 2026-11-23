# Autodoc v1.0
function Build-EntraSecurityReport {
    <#
        Builds the "Security and improvements" set into <export>\Report\security,
        all with the same name ("Säkerhet och förbättringar" / "Security and
        improvements") and timestamp:
          .json         the proposals and hunt results - the data the other
                        files are built from (rebuild later with -DataJson)
          .svg / .pdf   map: one card per area, coloured by highest severity
          .xlsx         proposals, affected objects and hunt results
          .docx         the document (with the map)
        Step 1 computes the proposals from the export's JSON (no Graph calls) and
        saves them as JSON; step 2 builds every file from that JSON file.
    #>
    [CmdletBinding()]
    param(
        [string]$SourceFolder,
        # Rebuild Word, Excel and map from an earlier .json instead of the export
        [string]$DataJson,
        [string]$Language = 'sv',
        [string]$ConsultantLogo,
        [string]$CustomerLogo
    )

    if (-not $SourceFolder -and -not $DataJson) { throw 'Give -SourceFolder (an export) or -DataJson (an earlier result).' }
    $cfgPath = Join-Path (Split-Path (Split-Path (Split-Path $PSScriptRoot))) 'Config\SecurityImprovements.json'
    $names = (Read-AutodocJson -Path $cfgPath -AsHashtable).Text.FileName
    $lang = if ("$Language" -eq 'en') { 'en' } else { 'sv' }
    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'

    if ($DataJson) {
        # rebuild from an earlier result: the output lands next to that file, in
        # the file's language unless -Language is given
        Write-Host "Building security and improvements from: $DataJson" -ForegroundColor Cyan
        $data = if ($PSBoundParameters.ContainsKey('Language')) { Import-SecurityImprovementData -Path $DataJson -Language $lang } else { Import-SecurityImprovementData -Path $DataJson }
        $lang = $data.Language
        $name = "$($names[$lang])"; if (-not $name) { $name = 'Security and improvements' }
        if (-not $SourceFolder -and $data.SourceFolder -and (Test-Path $data.SourceFolder)) { $SourceFolder = $data.SourceFolder }
        $outDir = Split-Path $DataJson -Parent
        $base = Join-Path $outDir "$name $stamp"
        $jsonPath = $DataJson
    }
    else {
        Write-Host "Building security and improvements from: $SourceFolder" -ForegroundColor Cyan
        $name = "$($names[$lang])"; if (-not $name) { $name = 'Security and improvements' }
        $outDir = Join-Path $SourceFolder 'Report\security'
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        $base = Join-Path $outDir "$name $stamp"
        # step 1: proposals from the export JSON -> one JSON file
        $jsonPath = Save-SecurityImprovementData -Data (Get-SecurityImprovementData -SourceFolder $SourceFolder -Language $lang) -Path "$base.json" -SourceFolder $SourceFolder
        Write-Host "  Data: $jsonPath" -ForegroundColor DarkGreen
        # step 2: everything below is built from that file
        $data = Import-SecurityImprovementData -Path $jsonPath
    }

    $result = [ordered]@{ Json = $jsonPath; Svg = $null; Pdf = $null; Excel = $null; Word = $null }

    # map first, so the Word file can embed it
    try {
        $result.Svg = Build-EntraSecurityMapSvg -Data $data -Path "$base.svg"
        if (Convert-SvgToPdf -SvgPath $result.Svg -PdfPath "$base.pdf") { $result.Pdf = "$base.pdf" }
        Write-Host "  Map: $($result.Svg)" -ForegroundColor DarkGreen
    } catch { Write-Host "  Map failed: $($_.Exception.Message)" -ForegroundColor Red }

    try {
        $result.Excel = Build-EntraSecurityExcelReport -Data $data -Path "$base.xlsx" -SourceFolder $SourceFolder
        Write-Host "  Excel: $($result.Excel)" -ForegroundColor DarkGreen
    } catch { Write-Host "  Excel failed: $($_.Exception.Message)" -ForegroundColor Red }

    try {
        $png = $null
        if ($result.Svg) {
            $png = Join-Path ([System.IO.Path]::GetTempPath()) ("secmap_" + [guid]::NewGuid().ToString('N') + '.png')
            if (-not (Convert-SvgToPng -SvgPath $result.Svg -PngPath $png)) { $png = $null }
        }
        $result.Word = Build-EntraSecurityWordReport -Data $data -Path "$base.docx" -SourceFolder $SourceFolder -MapPng $png `
            -ConsultantLogo $ConsultantLogo -CustomerLogo $CustomerLogo
        if ($png -and (Test-Path $png)) { Remove-Item -LiteralPath $png -Force -ErrorAction SilentlyContinue }
        Write-Host "  Word: $($result.Word)" -ForegroundColor DarkGreen
    } catch { Write-Host "  Word failed: $($_.Exception.Message)" -ForegroundColor Red }

    $sev = @($data.SevOrder | ForEach-Object { $s = $_; "$s $(@($data.Improvements | Where-Object Severity -eq $s).Count)" }) -join ', '
    Write-Host ""
    Write-Host "  $($data.Config.Text.Console_Done[$data.Language])  ($(@($data.Improvements).Count): $sev)" -ForegroundColor Green
    Write-Host "  $(if ($data.Language -eq 'sv') { 'Ändra förslagen i JSON-filen och bygg om med' } else { 'Edit the proposals in the JSON file and rebuild with' }): Build-EntraSecurityReport -DataJson `"$jsonPath`"" -ForegroundColor DarkGray
    return [pscustomobject]$result
}
