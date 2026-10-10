# Autodoc v1.0
function Build-ConditionalAccessMapSvg {
    <#
        SVG map of Conditional Access policies. One card per policy, coloured by
        state (green = enabled, amber = report-only, grey = disabled), with four
        summary lines: WHO (users/groups/roles, include & exclude), APPS (target
        cloud apps), WHEN (locations / platforms / device filter), and GRANT
        (required controls / MFA / authentication strength).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$PerRow = 3
    )

    $data = Get-ConditionalAccessMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No Conditional Access data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $policies = $data.Policies

    # full (raw) policy records, keyed by name, so the card can surface any
    # populated attribute that the summary rows don't already cover
    $fullCa = Get-CaFullData -SourceFolder $SourceFolder
    $rawByName = @{}
    if ($fullCa) { foreach ($rp in $fullCa.Policies) { $rawByName["$($rp.DisplayName)"] = $rp } }

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg
    # Conditional-access cards use the dedicated service icon file
    # (files\cache\Azure icons\Svg\10233-icon-service-Conditional-Access.svg),
    # resolved by its file name via the SVG-preferred icon index.
    $policyIcon = Get-MapIcon -IconSet $iconSet -Sku '10233-icon-service-Conditional-Access' -IconMap $iconMap -PreferSvg

    # state -> fill colour
    function Get-StateFill([string]$state) {
        switch ($state) {
            'enabled'                              { return '#C8E6C9' }  # green
            'enabledForReportingButNotEnforced'    { return '#FFE0B2' }  # amber (report-only)
            'disabled'                             { return '#E0E0E0' }  # grey
            default                                { return '#F0F0F0' }
        }
    }
    function Get-StateLabel([string]$state) {
        switch ($state) {
            'enabled'                              { return 'ON' }
            'enabledForReportingButNotEnforced'    { return 'REPORT-ONLY' }
            'disabled'                             { return 'OFF' }
            default                                { return $state }
        }
    }
    # collapse an include/exclude pair into a readable phrase
    function Join-IncExc([string]$inc, [string]$exc) {
        $parts = @()
        if ($inc) { $parts += $inc }
        if ($exc) { $parts += "excl. $exc" }
        if (-not $parts) { return '-' }
        return ($parts -join ', ')
    }

    $gapX = 0.3; $gapY = 0.4
    $tenantH = 1.25; $marginTop = 0.5; $marginBottom = 0.6
    $legendH = (Get-LegendHeight -LineCount 4)
    $cols = [math]::Max(1, [math]::Min($PerRow, 3))
    $lineH = 0.2; $cardPad = 0.14
    $pageW = 13.0
    $nodeW = ($pageW - 2 - ($cols - 1) * $gapX) / $cols
    $innerW = $nodeW - 2 * $cardPad

    # shared "Data source" block (JSON path + field glossary) shown at the very top
    $dsW = [math]::Max(6.0, $pageW - 2)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'ConditionalAccess' -WidthIn $dsW -FontSize 10 -GridFields
    $dsH = $dsBlock.Height

    # Usable text width inside a card (inches) for wrapping. Text is drawn to the
    # RIGHT of the icon (see Export-MapAsSvg): the icon occupies W*0.34 at W*0.05
    # from the left, with 6px padding on each side. Wrap to 97% of that as a
    # cushion for bold text and glyph-measurement rounding.
    $cardAvailIn = ($nodeW - $nodeW * 0.05 - $nodeW * 0.34 - 2 * (6 / 96)) * 0.97

    # Wrap a value to lines that fit a given width (inches) at a font size,
    # breaking on natural separators and hard-splitting any single token wider
    # than a line. The first line can reserve room for a bold key prefix
    # (e.g. "WHO: ") via -FirstReserveIn. Always returns at least one line.
    function Split-TextToWidth {
        param([string]$Text, [double]$AvailIn, [double]$FontSize, [double]$FirstReserveIn = 0)
        $Text = "$Text"
        $out = @()
        $cur = ''
        $firstBudget = $AvailIn - $FirstReserveIn
        $tokens = [regex]::Split($Text, '(?<=[\s;,/_@-])')
        foreach ($tok in $tokens) {
            $budget = if ($out.Count -eq 0) { $firstBudget } else { $AvailIn }
            while ((Measure-MapTextWidth $tok $FontSize) -gt $budget -and $tok.Length -gt 1) {
                $budget = if ($out.Count -eq 0 -and -not $cur) { $firstBudget } else { $AvailIn }
                $n = $tok.Length
                while ($n -gt 1 -and (Measure-MapTextWidth ($cur + $tok.Substring(0, $n)) $FontSize) -gt $budget) { $n-- }
                $out += ($cur + $tok.Substring(0, $n)).TrimEnd()
                $cur = ''
                $tok = $tok.Substring($n)
            }
            $budget = if ($out.Count -eq 0) { $firstBudget } else { $AvailIn }
            if ($cur -and (Measure-MapTextWidth ($cur + $tok) $FontSize) -gt $budget) {
                $out += $cur.TrimEnd(); $cur = $tok
            }
            else {
                $cur += $tok
            }
        }
        if ($cur) { $out += $cur.TrimEnd() }
        if ($out.Count -eq 0) { $out = @('') }
        return $out
    }

    # build card content + measured height. The four summary rows (WHO / APPS /
    # WHEN / GRANT) are kept; extra export parameters are folded into them and
    # two optional rows (RISK / SESSION) appear only when the policy sets them.
    # Long values wrap to the card's usable width so nothing spills outside.
    $cards = @()
    foreach ($p in $policies) {
        # WHO: users, roles, groups (include / exclude) - rendered as a readable
        # block below (each user = bold name + UPN line) via Get-CaWhoLines
        $roles = Join-IncExc $p.IncludeRoles  $p.ExcludeRoles
        $grp   = Join-IncExc $p.IncludeGroups $p.ExcludeGroups

        # APPS: target apps (include / exclude) or user actions
        $apps = Join-IncExc $p.IncludeApps $p.ExcludeApps
        if ($p.UserActions) { $apps = (@($apps, "user actions: $($p.UserActions)") -join '; ') }

        # WHEN: locations, platforms (include / exclude), device filter, client apps
        $whenParts = @()
        $locs = Join-IncExc $p.Locations  $p.ExcludeLocations
        if ($locs  -ne '-') { $whenParts += "loc: $locs" }
        $plats = Join-IncExc $p.Platforms $p.ExcludePlatforms
        if ($plats -ne '-') { $whenParts += "platform: $plats" }
        if ($p.DeviceFilter) { $whenParts += "device: $($p.DeviceFilter)" }
        if ($p.ClientAppTypes) { $whenParts += "client: $($p.ClientAppTypes)" }
        $when = if ($whenParts.Count) { $whenParts -join '; ' } else { 'any location / platform' }

        # GRANT: controls (with operator when several), auth strength, terms of use
        $grantParts = @()
        if ($p.Grant) {
            $ctrls = $p.Grant
            $nCtrl = @(($ctrls -split '[;,]').Where({ $_.Trim() })).Count
            if ($p.GrantOperator) { $ctrls = "$ctrls ($($p.GrantOperator))" }
            $grantParts += $ctrls
        }
        if ($p.AuthStrength) { $grantParts += "auth: $($p.AuthStrength)" }
        if ($p.TermsOfUse)   { $grantParts += "terms: $($p.TermsOfUse)" }
        $grant = if ($grantParts.Count) { $grantParts -join ', ' } else { '-' }

        # RISK: sign-in / user / service-principal (only when the policy sets them)
        $riskParts = @()
        if ($p.SignInRisk) { $riskParts += "sign-in: $($p.SignInRisk)" }
        if ($p.UserRisk)   { $riskParts += "user: $($p.UserRisk)" }
        if ($p.SpRisk)     { $riskParts += "SP: $($p.SpRisk)" }

        # SESSION: frequency / persistent browser / app-enforced / cloud app security
        $sessParts = @()
        if ($p.SignInFreq)        { $sessParts += "freq: $($p.SignInFreq)" }
        if ($p.PersistentBrowser) { $sessParts += "persistent browser: $($p.PersistentBrowser)" }
        if ($p.AppEnforced)       { $sessParts += "app-enforced: $($p.AppEnforced)" }
        if ($p.CloudAppSecurity)  { $sessParts += "cloud app security: $($p.CloudAppSecurity)" }

        $rows = @(
            @{ Key = 'APPS: ';  Text = $apps }
            @{ Key = 'WHEN: ';  Text = $when }
            @{ Key = 'GRANT: '; Text = $grant }
        )
        if ($riskParts.Count) { $rows += @{ Key = 'RISK: ';    Text = ($riskParts -join '; ') } }
        if ($sessParts.Count) { $rows += @{ Key = 'SESSION: '; Text = ($sessParts -join '; ') } }

        # any other populated attribute not already summarised above, so nothing
        # with a value is hidden. Skips GUID '*Raw' duplicates of the resolved
        # fields and the internal RecordType/PolicyId markers.
        $raw = $rawByName["$($p.Name)"]
        if ($raw) {
            $covered = @(
                'IncludeUsersResolved', 'ExcludeUsersResolved', 'IncludeRolesResolved', 'ExcludeRolesResolved',
                'IncludeGroupsResolved', 'ExcludeGroupsResolved', 'IncludeApplicationsResolved', 'ExcludeApplicationsResolved',
                'IncludeUserActionsRaw', 'IncludeLocationsResolved', 'ExcludeLocationsResolved', 'IncludePlatforms', 'ExcludePlatforms',
                'DeviceFilterRule', 'DeviceFilterMode', 'ClientAppTypes', 'BuiltInControls', 'AuthenticationStrengthResolved',
                'AuthenticationStrengthId', 'TermsOfUse', 'GrantOperator', 'SignInRiskLevels', 'UserRiskLevels', 'ServicePrincipalRiskLevels',
                'SignInFrequencyValue', 'SignInFrequencyType', 'PersistentBrowserMode', 'AppEnforcedRestrictionsEnabled',
                'CloudAppSecurityType', 'State', 'DisplayName', 'RecordType', 'PolicyId'
            )
            foreach ($prop in $raw.PSObject.Properties) {
                $n = $prop.Name
                $v = "$($prop.Value)".Trim()
                if (-not $v -or $v -eq 'null') { continue }
                if ($covered -contains $n) { continue }
                if ($n -like '*Raw') { continue }
                $rows += @{ Key = ($n + ': '); Text = $v }
            }
        }

        # pre-wrap into explicit lines: bold title, then bold key + wrapped value
        $lines = @()
        $titleText = "$($p.Name)  [$(Get-StateLabel $p.State)]"
        foreach ($t in @(Split-TextToWidth -Text $titleText -AvailIn $cardAvailIn -FontSize 10)) {
            $lines += @{ Text = $t; Bold = $true; Align = 'start' }
        }
        # readable WHO block: each targeted user as a bold name line + "UPN: ..." line
        $lines += @(Get-CaWhoLines -IncludeUsers $p.IncludeUsers -ExcludeUsers $p.ExcludeUsers -Roles $roles -Groups $grp -AvailIn $cardAvailIn -FontSize 10)
        foreach ($r in $rows) {
            $reserve = Measure-MapTextWidth $r.Key 9
            $parts   = @(Split-TextToWidth -Text $r.Text -AvailIn $cardAvailIn -FontSize 10 -FirstReserveIn $reserve)
            $lines += @{ BoldPrefix = $r.Key; Text = $parts[0]; Align = 'start' }
            for ($k = 1; $k -lt $parts.Count; $k++) { $lines += @{ Text = $parts[$k]; Bold = $false; Align = 'start' } }
        }

        $cardH = $cardPad + $lines.Count * $lineH + $cardPad
        $cards += [pscustomobject]@{ Name = $p.Name; State = $p.State; Lines = $lines; Height = $cardH; Fill = (Get-StateFill $p.State) }
    }

    # masonry placement
    $colHeights = @(0) * $cols; $placement = @()
    foreach ($card in $cards) {
        $c = 0; for ($k = 1; $k -lt $cols; $k++) { if ($colHeights[$k] -lt $colHeights[$c]) { $c = $k } }
        $placement += [pscustomobject]@{ Col = $c; TopOffset = $colHeights[$c] }
        $colHeights[$c] += $card.Height + $gapY
    }
    $tallestCol = ($colHeights | Measure-Object -Maximum).Maximum
    $vSpan = $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY + 0.6 + $tallestCol + $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    $enabledN  = @($policies | Where-Object State -eq 'enabled').Count
    $reportN   = @($policies | Where-Object State -eq 'enabledForReportingButNotEnforced').Count
    $disabledN = @($policies | Where-Object State -eq 'disabled').Count

    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)"; Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "CA policies: $($policies.Count)   (on $enabledN / report $reportN / off $disabledN)"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10; Columns=$dsBlock.Columns; GridFrom=$dsBlock.GridFrom
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ Text = 'How to read this map'; Bold = $true; Align = 'start' }
        @{ BoldPrefix = 'Green = '; Text = 'enabled   Amber = report-only   Grey = disabled'; Align = 'start' }
        @{ BoldPrefix = 'WHO / APPS / WHEN = '; Text = 'who it targets, which apps, and the conditions'; Align = 'start' }
        @{ BoldPrefix = 'GRANT = '; Text = 'required controls (MFA, compliant device, auth strength)'; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=7.4; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
    $gridTop = $legendY - $legendH / 2 - $gapY - 0.6
    for ($i = 0; $i -lt $cards.Count; $i++) {
        $card = $cards[$i]; $pl = $placement[$i]; $cardH = $card.Height
        $cx = $left + $pl.Col * ($nodeW + $gapX) + $nodeW / 2
        $cy = $gridTop - $pl.TopOffset - $cardH / 2
        $shapes += @{ Id="ca$i"; Kind='Rectangle'; Lines=$card.Lines; LinesTop=$true; Icon=$policyIcon
                      X=$cx; Y=$cy; W=$nodeW; H=$cardH; Fill=$card.Fill; Line='#888888'; FontSize=9 }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Conditional Access policies $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: $($policies.Count) CA policies (on $enabledN / report $reportN / off $disabledN)" -ForegroundColor DarkGray
    return $svg
}
