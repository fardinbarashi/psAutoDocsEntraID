# Autodoc v1.0
function Build-RbacMatrixSvg {
    <#
        RBAC broken into six category matrices (by function / risk). Each matrix
        has the people who hold a role in that category as ROWS and the category's
        roles as (rotated) COLUMNS. Cell symbols: a filled dot for an active
        assignment, a hollow dot for a PIM-eligible one. Sensitive/privileged
        roles get a red column header so high-risk access stands out. A "Data
        source" block at the top shows which JSON file the data comes from.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder
    )

    $data = Get-RbacMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No RBAC data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant

    # role -> category (in the order the categories should appear)
    $categories = [ordered]@{
        'Identity & Access Management' = @('Global Administrator', 'Global Reader', 'User Administrator', 'Authentication Administrator', 'Privileged Authentication Administrator', 'Directory Readers', 'Directory Writers', 'Directory Synchronization Accounts', 'Attribute Definition Reader')
        'Collaboration & Messaging'    = @('Exchange Administrator', 'Exchange Recipient Administrator', 'SharePoint Administrator', 'Teams Administrator', 'Groups Administrator')
        'Device Management'            = @('Intune Administrator', 'Device Managers', 'Azure AD Joined Device Local Administrator')
        'Security & Compliance'        = @('Security Administrator', 'Security Reader', 'Compliance Administrator', 'Purview Workload Content Writer')
        'Applications & Platform'      = @('Application Administrator', 'Cloud Application Administrator', 'Power Platform Administrator')
        'Operations & Support'         = @('Helpdesk Administrator', 'Service Support Administrator', 'Reports Reader')
    }
    $sensitive = @('Global Administrator', 'Privileged Authentication Administrator', 'Security Administrator', 'Compliance Administrator')

    # cell lookup: "role|principal" -> kind (Active wins over Eligible)
    $cell = @{}
    $presentRoles = @{}
    foreach ($r in $data.Roles) {
        $presentRoles[$r.Name] = $true
        foreach ($n in $r.Group) {
            $key = "$($r.Name)|$($n.Principal)"
            if (-not $cell.ContainsKey($key) -or $n.Kind -eq 'Active') { $cell[$key] = $n.Kind }
        }
    }

    # build the per-category matrices (only roles present in the data, only
    # people who hold at least one role in that category)
    $matrices = @()
    foreach ($cat in $categories.Keys) {
        $catRoles = @($categories[$cat] | Where-Object { $presentRoles[$_] })
        if (-not $catRoles) { continue }
        $personSet = @{}
        foreach ($rn in $catRoles) {
            foreach ($k in $cell.Keys) {
                if ($k -like "$rn|*") { $personSet[$k.Substring($rn.Length + 1)] = $true }
            }
        }
        if (-not $personSet.Count) { continue }
        $persons = @($personSet.Keys | Sort-Object @{ e = { -(@($catRoles | Where-Object { $cell.ContainsKey("$_|$($_)") }).Count) } }, { $_ })
        # count roles held per person for sorting (busiest first)
        $held = @{}
        foreach ($p in $personSet.Keys) { $held[$p] = @($catRoles | Where-Object { $cell.ContainsKey("$_|$p") }).Count }
        $persons = @($personSet.Keys | Sort-Object @{ e = { -$held[$_] } }, { $_ })
        $matrices += [pscustomobject]@{ Cat = $cat; Roles = $catRoles; Persons = $persons }
    }

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg

    # geometry
    $marginTop = 0.5; $marginBottom = 0.6; $tenantH = 1.55; $gapY = 0.5
    $labelW = 3.2; $colW = 0.5; $rowH = 0.4; $headH = 2.6; $titleH = 0.4; $catGap = 0.7

    $maxCols = ($matrices | ForEach-Object { $_.Roles.Count } | Measure-Object -Maximum).Maximum
    if (-not $maxCols) { $maxCols = 1 }
    $pageW = [math]::Max(11, $labelW + $maxCols * $colW + 2.0)

    $dsW = [math]::Max(6.0, [math]::Min($pageW - 1, 12.0))
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'RBAC' -WidthIn $dsW -FontSize 10 -PathOnly
    $dsH = $dsBlock.Height

    $matricesH = 0.0
    foreach ($m in $matrices) { $matricesH += $titleH + $headH + $m.Persons.Count * $rowH + $catGap }
    $legendH = 0.95

    $pageH = [math]::Max(8.5, $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY + $matricesH + $marginBottom)

    $shapes = @()
    $today = Get-Date -Format 'yyyy-MM-dd'
    $topY  = $pageH - $marginTop - $tenantH / 2
    $dsY = $topY - $tenantH / 2 - $gapY - $dsH / 2

    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    # summary counts for the tenant block: how many principals hold a privileged
    # (sensitive) role, out of everyone holding any role
    $allHolders = @{}; $privHolders = @{}
    foreach ($k in $cell.Keys) {
        $parts = $k -split '\|', 2
        $allHolders[$parts[1]] = $true
        if ($sensitive -contains $parts[0]) { $privHolders[$parts[1]] = $true }
    }
    $privCount = $privHolders.Count
    $allCount = $allHolders.Count

    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = "RBAC by category  ($($matrices.Count) matrices)"; Bold = $false; Align = 'start' }
            @{ BoldPrefix = 'Privileged-role holders: '; Text = "$privCount  (of $allCount with any role)"; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ Text = 'How to read this matrix'; Bold = $true; Align = 'start' }
        @{ Text = "$([char]0x25CF) active assignment"; Bold = $false; Align = 'start' }
        @{ Text = "$([char]0x25CB) eligible (PIM)"; Bold = $false; Align = 'start' }
        @{ Text = 'Red column header = sensitive / privileged role'; Bold = $false; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$legendY; W=($pageW - 2); H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    # lay out each category matrix top-down
    $curTop = $legendY - $legendH / 2 - $gapY
    $mi = 0
    foreach ($m in $matrices) {
        $gridW = $labelW + $m.Roles.Count * $colW
        $gridLeft = ($pageW - $gridW) / 2

        # category title band
        $titleCy = $curTop - $titleH / 2
        $shapes += @{ Id="mtitle$mi"; Kind='Rectangle'; Lines=@(@{ Text="$($m.Cat)  ($($m.Persons.Count) people)"; Bold=$true; Align='start' }); LinesTop=$true; TopInset=0.06
                      X=$pageW/2; Y=$titleCy; W=$gridW; H=$titleH; Fill='#E8EDF3'; Line='#B8BFC7'; FontSize=11 }
        $headTop = $curTop - $titleH

        # rotated role column headers (sensitive roles get a red header)
        for ($c = 0; $c -lt $m.Roles.Count; $c++) {
            $role = $m.Roles[$c]
            $colCx = $gridLeft + $labelW + $c * $colW + $colW / 2
            $isSens = $sensitive -contains $role
            $hfill = if ($isSens) { '#F8D2D5' } else { '#EEF2F7' }
            $hline = if ($isSens) { '#C81E1E' } else { '#CCCCCC' }
            $hname = "$role"; if ($hname.Length -gt 48) { $hname = $hname.Substring(0, 47) + [char]0x2026 }
            $shapes += @{ Id="mh$mi-$c"; Kind='Rectangle'; Text=$hname; Rotate=-90
                          X=$colCx; Y=($headTop - $headH/2); W=($colW - 0.03); H=($headH - 0.06)
                          Fill=$hfill; Line=$hline; FontSize=9; Bold=$isSens }
        }

        # person rows
        for ($rIdx = 0; $rIdx -lt $m.Persons.Count; $rIdx++) {
            $person = $m.Persons[$rIdx]
            $rowCy = $headTop - $headH - $rIdx * $rowH - $rowH / 2
            $pname = "$person"; if ($pname.Length -gt 52) { $pname = $pname.Substring(0, 51) + [char]0x2026 }
            $shapes += @{ Id="mr$mi-$rIdx"; Kind='Rectangle'; Text=$pname
                          X=($gridLeft + $labelW/2); Y=$rowCy; W=($labelW - 0.04); H=($rowH - 0.04)
                          Fill='#F4F6F9'; Line='#DDDDDD'; FontSize=10 }
            for ($c = 0; $c -lt $m.Roles.Count; $c++) {
                $role = $m.Roles[$c]
                $k = "$role|$person"
                $cellCx = $gridLeft + $labelW + $c * $colW + $colW / 2
                $sym = '-'; $fill = '#FFFFFF'
                if ($cell.ContainsKey($k)) {
                    if ($cell[$k] -eq 'Active') { $sym = [char]0x25CF; $fill = '#C8E6C9' }
                    else { $sym = [char]0x25CB; $fill = '#FFE0B2' }
                }
                $shapes += @{ Id="mc$mi-$rIdx-$c"; Kind='Rectangle'; Text="$sym"; CenterText=$true
                              X=$cellCx; Y=$rowCy; W=($colW - 0.03); H=($rowH - 0.04)
                              Fill=$fill; Line='#DDDDDD'; FontSize=12 }
            }
        }

        $curTop = $headTop - $headH - $m.Persons.Count * $rowH - $catGap
        $mi++
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - RBAC matrix (roles x people) $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG (matrix): $($matrices.Count) category matrices" -ForegroundColor DarkGray
    return $svg
}
