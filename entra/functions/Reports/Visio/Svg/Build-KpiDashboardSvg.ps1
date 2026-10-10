# Autodoc v1.0
function Build-KpiDashboardSvg {
    <#
        A one-page dashboard of the tenant, drawn as doughnut charts in
        self-contained SVG (no external libraries). Because it is a normal map it
        flows into the SVG, PDF and HTML reports automatically. Values come from
        the shared resolver.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder
    )

    $v = Get-WordTemplateValues -SourceFolder $SourceFolder
    if (-not $v) { Write-Host "No data for the KPI dashboard." -ForegroundColor Yellow; return }
    function N([string]$k) { $x = $v[$k]; if ($null -eq $x -or $x -eq '') { 0 } else { [int]$x } }

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $tName  = if ($tenant) { "$($tenant.TenantName)" } else { 'Microsoft Entra ID' }
    $tDom   = if ($tenant) { "$($tenant.PrimaryDomain)" } else { '' }
    $tEnv   = ("$(Get-TenantHybridLabel -Tenant $tenant)" -replace '^Environment:\s*', '')
    $today  = Get-Date -Format 'yyyy-MM-dd'

    $charts = @(
        @{ T = 'Users';               S = @(@{N='Members';V=(N 'Members');C='#0F6CBD'}, @{N='Guests';V=(N 'Guests');C='#4CA6E0'}) }
        @{ T = 'User status';         S = @(@{N='Enabled';V=(N 'EnabledUsers');C='#2E7D32'}, @{N='Disabled';V=(N 'DisabledUsers');C='#B4B4B4'}) }
        @{ T = 'Groups by type';      S = @(@{N='Security';V=(N 'SecurityGroups');C='#0F6CBD'}, @{N='Microsoft 365';V=(N 'M365Groups');C='#4CA6E0'}, @{N='Dynamic';V=(N 'DynamicGroups');C='#F9A825'}) }
        @{ T = 'Group source';        S = @(@{N='Cloud';V=(N 'CloudGroups');C='#0F6CBD'}, @{N='On-premises';V=(N 'OnPremisesGroups');C='#8C6D3F'}) }
        @{ T = 'Conditional Access';  S = @(@{N='Enabled';V=(N 'EnabledPolicies');C='#2E7D32'}, @{N='Report-only';V=(N 'ReportOnlyPolicies');C='#F9A825'}, @{N='Disabled';V=(N 'DisabledPolicies');C='#B4B4B4'}) }
        @{ T = 'App registrations';   S = @(@{N='Single-tenant';V=(N 'SingleTenantApps');C='#0F6CBD'}, @{N='Multi-tenant';V=(N 'MultiTenantApps');C='#4CA6E0'}) }
        @{ T = 'Enterprise apps SSO'; S = @(@{N='Configured';V=(N 'SsoConfigured');C='#2E7D32'}, @{N='Not configured';V=(N 'WithoutSso');C='#B4B4B4'}) }
        @{ T = 'RBAC assignments';    S = @(@{N='Active';V=(N 'ActiveAssignments');C='#0F6CBD'}, @{N='Eligible';V=(N 'EligibleAssignments');C='#F9A825'}) }
    )

    $W = 1440; $margin = 40; $cols = 4
    $headTop = 40; $headH = 150; $gridTop = $headTop + $headH + 34
    $cellW = [math]::Round(($W - 2 * $margin) / $cols); $rowH = 360
    $rows = [math]::Ceiling($charts.Count / $cols)
    $noteH = 70
    $H = $gridTop + $rows * $rowH + 24 + $margin

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("<svg xmlns='http://www.w3.org/2000/svg' width='$W' height='$H' viewBox='0 0 $W $H' font-family='Segoe UI, sans-serif'>")
    [void]$sb.AppendLine("<rect x='0' y='0' width='$W' height='$H' fill='#ffffff'/>")

    # header
    $hw = $W - 2 * $margin
    [void]$sb.AppendLine("<rect x='$margin' y='$headTop' width='$hw' height='$headH' rx='6' fill='#0F6CBD' stroke='#0B4C87'/>")
    $hx = $margin + 28; $hy = $headTop + 40
    [void]$sb.AppendLine("<text x='$hx' y='$hy' font-size='20' font-weight='bold' fill='#ffffff'>$([System.Net.WebUtility]::HtmlEncode($tName))</text>")
    foreach ($ln in @($tDom, "Environment: $tEnv", 'Microsoft Entra ID', "Generated $today")) {
        $hy += 24
        $wt = if ($ln -eq 'Microsoft Entra ID') { "font-weight='bold'" } else { '' }
        [void]$sb.AppendLine("<text x='$hx' y='$hy' font-size='13' fill='#ffffff' $wt>$([System.Net.WebUtility]::HtmlEncode($ln))</text>")
    }

    # doughnut
    function Pt($cx, $cy, $r, $deg) {
        $a = ($deg - 90) * [math]::PI / 180
        "$([math]::Round($cx + $r * [math]::Cos($a), 1)) $([math]::Round($cy + $r * [math]::Sin($a), 1))"
    }

    for ($i = 0; $i -lt $charts.Count; $i++) {
        $c = $charts[$i]
        $col = $i % $cols; $row = [math]::Floor($i / $cols)
        $cellX = $margin + $col * $cellW; $cellY = $gridTop + $row * $rowH
        $cx = $cellX + $cellW / 2; $cyd = $cellY + 118; $rO = 70; $rI = 44
        $total = ($c.S | Measure-Object -Property V -Sum).Sum

        # title
        [void]$sb.AppendLine("<text x='$([math]::Round($cx,1))' y='$($cellY + 22)' font-size='14' font-weight='bold' fill='#1a1a1a' text-anchor='middle'>$([System.Net.WebUtility]::HtmlEncode($c.T))</text>")

        # doughnut segments
        if ($total -le 0) {
            [void]$sb.AppendLine("<circle cx='$([math]::Round($cx,1))' cy='$cyd' r='$rO' fill='none' stroke='#ECECEC' stroke-width='$($rO-$rI)'/>")
        }
        else {
            $ang = 0.0
            foreach ($s in $c.S) {
                if ($s.V -le 0) { continue }
                $sweep = [math]::Min(359.9, 360 * $s.V / $total)
                $a0 = $ang; $a1 = $ang + $sweep; $ang = $a1
                $large = if ($sweep -gt 180) { 1 } else { 0 }
                $oS = Pt $cx $cyd $rO $a0; $oE = Pt $cx $cyd $rO $a1
                $iE = Pt $cx $cyd $rI $a1; $iS = Pt $cx $cyd $rI $a0
                [void]$sb.AppendLine("<path d='M $oS A $rO $rO 0 $large 1 $oE L $iE A $rI $rI 0 $large 0 $iS Z' fill='$($s.C)'/>")
            }
        }
        # centre total
        [void]$sb.AppendLine("<text x='$([math]::Round($cx,1))' y='$($cyd+2)' font-size='20' font-weight='bold' fill='#1a1a1a' text-anchor='middle'>$('{0:N0}' -f $total)</text>")
        [void]$sb.AppendLine("<text x='$([math]::Round($cx,1))' y='$($cyd+20)' font-size='10' fill='#5b6470' text-anchor='middle'>total</text>")

        # legend
        $ly = $cellY + 210; $lx = $cellX + 34
        foreach ($s in $c.S) {
            $pct = if ($total -gt 0) { [math]::Round(100 * $s.V / $total) } else { 0 }
            [void]$sb.AppendLine("<rect x='$lx' y='$($ly-10)' width='12' height='12' rx='2' fill='$($s.C)'/>")
            [void]$sb.AppendLine("<text x='$($lx+18)' y='$ly' font-size='12' fill='#1a1a1a'>$([System.Net.WebUtility]::HtmlEncode($s.N)): $('{0:N0}' -f [int]$s.V) ($pct%)</text>")
            $ly += 22
        }
    }

    [void]$sb.AppendLine('</svg>')

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    # remove any earlier dashboard file so no stale copy lingers
    Get-ChildItem -Path $svgDir -Filter '*ashboard*.svg' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    $svg = Join-Path $svgDir "Entra ID - Dashboard $stamp.svg"
    Set-Content -LiteralPath $svg -Value $sb.ToString() -Encoding UTF8
    Write-Host "  SVG: KPI dashboard ($($charts.Count) charts)" -ForegroundColor DarkGray
    return $svg
}
