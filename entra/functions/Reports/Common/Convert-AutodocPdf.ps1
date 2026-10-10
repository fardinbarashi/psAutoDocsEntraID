# Autodoc v1.0
function Get-PdfConverterExe {
    <#
        Locates an external converter executable by probing the usual install
        paths (Windows) and PATH (any OS). Returns the full path or $null.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Edge', 'Chrome', 'Soffice')][string]$Which)

    $candidates = switch ($Which) {
        'Edge'    { @(
                "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
                'msedge'
            ) }
        'Chrome'  { @(
                "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
                'chrome', 'google-chrome'
            ) }
        'Soffice' { @(
                "$env:ProgramFiles\LibreOffice\program\soffice.exe",
                "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
                'soffice', 'libreoffice'
            ) }
    }
    foreach ($c in $candidates) {
        if ($c -match '[\\/]') { if (Test-Path $c) { return $c } }
        else { $cmd = Get-Command $c -ErrorAction SilentlyContinue; if ($cmd) { return $cmd.Source } }
    }
    return $null
}

function Convert-WithSoffice {
    # Shared LibreOffice fallback: converts any supported file to PDF. soffice
    # writes <basename>.pdf into the out dir, so we convert to a temp dir and
    # move the result to the requested path.
    [CmdletBinding()]
    param([string]$SrcPath, [string]$PdfPath, [string]$Soffice)

    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("autodocpdf_" + [guid]::NewGuid().ToString('N'))
    New-Item -Path $tmp -ItemType Directory -Force | Out-Null
    try {
        & $Soffice --headless --convert-to pdf --outdir $tmp $SrcPath 2>&1 | Out-Null
        $made = Join-Path $tmp ([System.IO.Path]::GetFileNameWithoutExtension($SrcPath) + '.pdf')
        if (Test-Path $made) { Move-Item -Path $made -Destination $PdfPath -Force; return (Test-Path $PdfPath) }
        return $false
    }
    finally { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

function Convert-SvgToPdf {
    <#
        SVG map -> PDF. Prefers Edge/Chrome headless (built into Windows), falls
        back to LibreOffice. Returns $true on success.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SvgPath, [Parameter(Mandatory)][string]$PdfPath)

    if (-not (Test-Path $SvgPath)) { return $false }
    $dir = Split-Path $PdfPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }

    # 1) Edge or Chrome headless
    $browser = Get-PdfConverterExe -Which Edge
    if (-not $browser) { $browser = Get-PdfConverterExe -Which Chrome }
    if ($browser) {
        try {
            $uri = 'file:///' + ((Resolve-Path $SvgPath).Path -replace '\\', '/')
            $args = @('--headless=new', '--disable-gpu', '--no-pdf-header-footer', "--print-to-pdf=$PdfPath", $uri)
            Start-Process -FilePath $browser -ArgumentList $args -Wait -WindowStyle Hidden -ErrorAction Stop
            if (Test-Path $PdfPath) { return $true }
        }
        catch { }
    }

    # 2) LibreOffice fallback
    $soffice = Get-PdfConverterExe -Which Soffice
    if ($soffice) { return (Convert-WithSoffice -SrcPath $SvgPath -PdfPath $PdfPath -Soffice $soffice) }

    Write-Host "  PDF: no Edge/Chrome or LibreOffice available to convert $(Split-Path $SvgPath -Leaf)." -ForegroundColor DarkYellow
    return $false
}

function Convert-SvgToPng {
    <#
        SVG map -> PNG (for embedding in the Word report). Prefers Edge/Chrome
        headless screenshot (built into Windows), falls back to LibreOffice.
        Returns $true on success.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SvgPath, [Parameter(Mandatory)][string]$PngPath, [int]$MaxPx = 2600)

    if (-not (Test-Path $SvgPath)) { return $false }
    $dir = Split-Path $PngPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }

    # read the SVG's pixel size so the browser window matches the whole canvas
    $w = 1600; $h = 1000
    try {
        $head = Get-Content -LiteralPath $SvgPath -TotalCount 6 -Raw
        if ($head -match "<svg[^>]*\bwidth='([\d.]+)'[^>]*\bheight='([\d.]+)'") { $w = [int][double]$Matches[1]; $h = [int][double]$Matches[2] }
        elseif ($head -match '<svg[^>]*\bwidth="([\d.]+)"[^>]*\bheight="([\d.]+)"') { $w = [int][double]$Matches[1]; $h = [int][double]$Matches[2] }
    }
    catch { }
    if ($w -gt $MaxPx) { $h = [int]($h * $MaxPx / $w); $w = $MaxPx }
    if ($w -lt 100) { $w = 1600 }
    if ($h -lt 100) { $h = 1000 }

    # 1) Edge or Chrome headless screenshot
    $browser = Get-PdfConverterExe -Which Edge
    if (-not $browser) { $browser = Get-PdfConverterExe -Which Chrome }
    if ($browser) {
        try {
            $uri = 'file:///' + ((Resolve-Path $SvgPath).Path -replace '\\', '/')
            $args = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--default-background-color=FFFFFFFF', "--screenshot=$PngPath", "--window-size=$w,$h", $uri)
            Start-Process -FilePath $browser -ArgumentList $args -Wait -WindowStyle Hidden -ErrorAction Stop
            if (Test-Path $PngPath) { return $true }
        }
        catch { }
    }

    # 2) LibreOffice fallback
    $soffice = Get-PdfConverterExe -Which Soffice
    if ($soffice) {
        try {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $tmp -Force | Out-Null
            & $soffice --headless --convert-to png --outdir $tmp $SvgPath 2>&1 | Out-Null
            $made = Get-ChildItem -Path $tmp -Filter '*.png' | Select-Object -First 1
            if ($made) { Copy-Item -LiteralPath $made.FullName -Destination $PngPath -Force; Remove-Item $tmp -Recurse -Force; return (Test-Path $PngPath) }
            Remove-Item $tmp -Recurse -Force
        }
        catch { }
    }

    Write-Host "  PNG: no Edge/Chrome or LibreOffice available to convert $(Split-Path $SvgPath -Leaf)." -ForegroundColor DarkYellow
    return $false
}
