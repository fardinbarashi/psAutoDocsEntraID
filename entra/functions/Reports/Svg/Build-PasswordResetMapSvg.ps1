# Autodoc v1.0
function Build-PasswordResetMapSvg {
    <#
        SVG card of the tenant's Self-Service Password Reset (SSPR) configuration:
        one configuration card listing every setting (user scope, methods, admin
        settings). SSPR is tenant-wide, so this is a single card, not a grid.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SourceFolder)

    $data = Get-PasswordResetData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No password reset (SSPR) data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $cfg = $data.Config

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg

    $lines = @(Get-PasswordResetLines -Config $cfg)

    $marginTop = 0.5; $marginBottom = 0.6; $tenantH = 1.35; $gapY = 0.6
    $legendH = (Get-LegendHeight -LineCount 2)
    $lineH = 0.26; $cardPad = 0.2
    $cardH = 0.30 + ($lines.Count + 1) * 0.19   # match the rendered LinesTop spacing (was too tall)
    $pageW = 9.0

    # shared "Data source" block (JSON path only) at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'PasswordReset' -WidthIn $dsW -FontSize 10 -PathOnly
    $dsH = $dsBlock.Height
    $pageH = [math]::Max(8.5, $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY + $cardH + $marginBottom)

    $shapes = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)"; Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Self-Service Password Reset (SSPR)"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ Text = 'How to read this map'; Bold = $true; Align = 'start' }
        @{ BoldPrefix = 'Card = '; Text = 'the tenant-wide SSPR configuration'; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=5.6; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    # SSPR config card
    $cardY = $legendY - $legendH / 2 - $gapY - $cardH / 2
    $cardLines = @(@{ Text = 'SSPR configuration'; Bold = $true; Align = 'start' })
    foreach ($l in $lines) {
        # split "Label: value" so the label can be bold
        $idx = $l.IndexOf(':')
        if ($idx -gt 0) {
            $label = $l.Substring(0, $idx + 1) + ' '
            $rest  = $l.Substring($idx + 1).Trim()
            $cardLines += @{ BoldPrefix = $label; Text = $rest; Align = 'start' }
        } else {
            $cardLines += @{ Text = $l; Bold = $false; Align = 'start' }
        }
    }
    $shapes += @{ Id='sspr'; Kind='Rectangle'; Lines=$cardLines; LinesTop=$true; TopInset=0.14
                  X=$pageW/2; Y=$cardY; W=7.0; H=$cardH; Fill='#EEF2F7'; Line='#888888'; FontSize=10 }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Password reset (SSPR) $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: SSPR configuration ($(@($lines).Count) settings)" -ForegroundColor DarkGray
    return $svg
}
