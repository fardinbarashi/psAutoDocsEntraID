# Autodoc v1.0
function Build-RbacMapSvg {
    <#
        RBAC as six category sections. Each section has a coloured header and the
        roles in that category as cards; each card lists the people who hold the
        role (filled dot = active, hollow dot = eligible), capped for readability.
        Privileged roles are marked "P" (and a red border), read-only roles "R" -
        symbols are always paired with text so nothing depends on colour alone.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$Columns = 2,
        [int]$MaxUsers = 0      # 0 = show all users per role
    )

    $data = Get-RbacMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No RBAC data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant

    $categories = [ordered]@{
        'Identity & Access'         = @('Global Administrator', 'Global Reader', 'User Administrator', 'Authentication Administrator', 'Privileged Authentication Administrator', 'Directory Readers', 'Directory Writers', 'Directory Synchronization Accounts', 'Attribute Definition Reader')
        'Collaboration & Messaging' = @('SharePoint Administrator', 'Teams Administrator', 'Exchange Administrator', 'Exchange Recipient Administrator', 'Groups Administrator')
        'Device Management'         = @('Intune Administrator', 'Device Managers', 'Azure AD Joined Device Local Administrator')
        'Security & Compliance'     = @('Security Administrator', 'Security Reader', 'Compliance Administrator', 'Purview Workload Content Writer')
        'Applications & Platform'   = @('Application Administrator', 'Cloud Application Administrator', 'Power Platform Administrator')
        'Operations & Support'      = @('Helpdesk Administrator', 'Service Support Administrator', 'Reports Reader')
    }
    $catColour = @{
        'Identity & Access'         = @{ Head = '#E4D7F3'; Card = '#F4EFFB'; Line = '#7E57C2' }  # purple
        'Collaboration & Messaging' = @{ Head = '#D8EAD1'; Card = '#EEF6EA'; Line = '#5B9A4E' }  # green
        'Device Management'         = @{ Head = '#D4E3F5'; Card = '#EBF2FB'; Line = '#3F72AF' }  # blue
        'Security & Compliance'     = @{ Head = '#F6E0C6'; Card = '#FBF1E4'; Line = '#D98A34' }  # orange
        'Applications & Platform'   = @{ Head = '#F5EEC2'; Card = '#FBF8E2'; Line = '#C4AB2E' }  # yellow
        'Operations & Support'      = @{ Head = '#CFEBE8'; Card = '#EAF6F4'; Line = '#3E9E97' }  # teal
    }
    $privileged = @('Global Administrator', 'Privileged Authentication Administrator', 'Security Administrator', 'Compliance Administrator')
    $readonly   = @('Global Reader', 'Directory Readers', 'Attribute Definition Reader', 'Security Reader', 'Reports Reader')

    $rolesByName = @{}
    foreach ($r in $data.Roles) { $rolesByName[$r.Name] = $r }

    # geometry (SVG user units are px at 96/inch; fonts in px)
    $marginTop = 0.5; $marginBottom = 0.6; $marginSide = 0.6
    $tenantH = 1.7; $legendH = 1.35; $gapY = 0.55
    $cardW = 5.0; $cardGapX = 0.35; $cardGapY = 0.3
    $headerH = 0.62; $catGap = 0.75
    $lineH = 0.205; $cardPadTop = 0.30; $cardPadBot = 0.16
    $fsRole = 15; $fsUser = 14; $fsCat = 22

    $cols = [math]::Max(1, $Columns)
    $gridW = $cols * $cardW + ($cols - 1) * $cardGapX
    $pageW = [math]::Max(11, $gridW + 2 * $marginSide)

    # shared "Data source" block (JSON path + field glossary) shown at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'RBAC' -WidthIn $dsW -FontSize 10 -GridFields -GridMaxColumns 3
    $dsH = $dsBlock.Height

    $cleanName = { param($p) $s = "$p"; if ($s -match '^(.*?)\s*<') { $matches[1].Trim() } else { $s.Trim() } }

    # build the sections (only categories/roles present in the data)
    $sections = @()
    foreach ($cat in $categories.Keys) {
        $present = @($categories[$cat] | Where-Object { $rolesByName.ContainsKey($_) })
        if (-not $present) { continue }
        $cards = @()
        $catAssignments = 0
        foreach ($role in $present) {
            $members = @($rolesByName[$role].Group)
            $n = $members.Count
            $catAssignments += $n
            $isPriv = $privileged -contains $role
            $isRead = $readonly -contains $role
            $mark = if ($isPriv) { 'P  ' } elseif ($isRead) { 'R  ' } else { '' }
            $userWord = if ($n -eq 1) { 'user' } else { 'users' }

            $lines = @()
            $lines += @{ Text = "$mark$role  ($n $userWord)"; Bold = $true; Align = 'start' }
            $shown = if ($MaxUsers -gt 0) { $members | Select-Object -First $MaxUsers } else { $members }
            foreach ($m in $shown) {
                $sym = if ("$($m.Kind)" -eq 'Active') { [char]0x25CF } else { [char]0x25CB }
                $suffix = if ("$($m.Kind)" -ne 'Active') { ' (eligible)' } else { '' }
                $nm = & $cleanName $m.Principal
                if ($nm.Length -gt 46) { $nm = $nm.Substring(0, 45) + [char]0x2026 }
                $lines += @{ Text = "$sym  $nm$suffix"; Bold = $false; Align = 'start' }
            }
            if ($MaxUsers -gt 0 -and $n -gt $MaxUsers) { $lines += @{ Text = "+ $($n - $MaxUsers) more"; Bold = $false; Align = 'start' } }

            $h = $cardPadTop + $lines.Count * $lineH + $cardPadBot
            $cards += [pscustomobject]@{ Lines = $lines; H = $h; Priv = $isPriv }
        }
        $sections += [pscustomobject]@{ Name = $cat; Roles = $present.Count; Assignments = $catAssignments; Cards = $cards }
    }

    # masonry height for a set of cards over N columns
    $sectionHeight = {
        param($cards)
        $colH = @(0) * $cols
        foreach ($c in $cards) {
            $mi = 0; for ($i = 1; $i -lt $cols; $i++) { if ($colH[$i] -lt $colH[$mi]) { $mi = $i } }
            $colH[$mi] += $c.H + $cardGapY
        }
        ($colH | Measure-Object -Maximum).Maximum
    }

    # privileged summary
    $privHolders = @{}
    foreach ($r in $data.Roles) { if ($privileged -contains $r.Name) { foreach ($m in $r.Group) { $privHolders["$($m.Principal)"] = $true } } }

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg

    # total page height
    $contentH = 0.0
    foreach ($sec in $sections) { $contentH += $headerH + $gapY + (& $sectionHeight $sec.Cards) + $catGap }
    $pageH = [math]::Max(8.5, $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY + $contentH + $marginBottom)

    $shapes = @()
    $today = Get-Date -Format 'yyyy-MM-dd'
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10; Columns=$dsBlock.Columns; GridFrom=$dsBlock.GridFrom
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = 'RBAC by category'; Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ BoldPrefix = 'Privileged-role holders: '; Text = "$($privHolders.Count)  (of $($data.Total) assignments; active $($data.ActiveN) / eligible $($data.EligibleN))"; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=14; Icon=$tenantIcon }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ Text = 'Legend'; Bold = $true; Align = 'start' }
        @{ BoldPrefix = "$([char]0x25CF)  "; Text = 'Active Role'; Align = 'start' }
        @{ BoldPrefix = "$([char]0x25CB)  "; Text = 'Eligible Role'; Align = 'start' }
        @{ BoldPrefix = 'P  '; Text = 'Privileged Role (also red border)'; Align = 'start' }
        @{ BoldPrefix = 'R  '; Text = 'Read-Only Access'; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$legendY; W=($pageW - 2 * $marginSide); H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=14 }

    $gridLeft = ($pageW - $gridW) / 2
    $curTop = $legendY - $legendH / 2 - $gapY
    $si = 0
    foreach ($sec in $sections) {
        $pal = $catColour[$sec.Name]

        # section header (full width, large, category colour)
        $headCy = $curTop - $headerH / 2
        $shapes += @{ Id="sec$si"; Kind='Rectangle'
                      Lines=@(@{ Text = "$($sec.Name)   -   $($sec.Roles) roles, $($sec.Assignments) assignments"; Bold = $true; Align = 'start' })
                      LinesTop=$true; TopInset=0.12
                      X=$pageW/2; Y=$headCy; W=$gridW; H=$headerH; Fill=$pal.Head; Line=$pal.Line; FontSize=$fsCat }
        $cardsTop = $curTop - $headerH - $gapY

        # masonry the role cards
        $colBottom = @($cardsTop) * $cols
        foreach ($c in $sec.Cards) {
            $mi = 0; for ($i = 1; $i -lt $cols; $i++) { if ($colBottom[$i] -gt $colBottom[$mi]) { $mi = $i } }
            $cx = $gridLeft + $mi * ($cardW + $cardGapX) + $cardW / 2
            $cy = $colBottom[$mi] - $c.H / 2
            $edge = if ($c.Priv) { '#C81E1E' } else { $pal.Line }
            $shapes += @{ Id="card$si-$mi-$([math]::Round($colBottom[$mi],2))"; Kind='Rectangle'; Lines=$c.Lines; LinesTop=$true; TopInset=0.10
                          X=$cx; Y=$cy; W=$cardW; H=$c.H; Fill=$pal.Card; Line=$edge; FontSize=$fsUser }
            $colBottom[$mi] = $colBottom[$mi] - $c.H - $cardGapY
        }
        $lowest = ($colBottom | Measure-Object -Minimum).Minimum
        $curTop = $lowest - $catGap + $cardGapY
        $si++
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - RBAC roles by category $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: $($sections.Count) category sections" -ForegroundColor DarkGray
    return $svg
}
