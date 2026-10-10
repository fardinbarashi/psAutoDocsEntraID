# Autodoc v1.0
function Build-GroupDeptMatrixSvg {
    <#
        Group -> department matrices, one SVG FILE PER FIRST LETTER. Each file
        contains every group whose name starts with that letter (rows) against
        ALL departments those groups have members in (rotated columns). Because
        each letter is its own file, every department is included without making
        a single unopenable file. Cell numbers = members of the group in that
        department; each group row gets its own colour.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SourceFolder)

    $data = Get-GroupDeptMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No group-member data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $tree = $data.Tree

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg

    # group the groups by first letter (non-letters -> '#')
    $byLetter = [ordered]@{}
    foreach ($g in ($data.Groups | Sort-Object)) {
        $ch = ("$g".Substring(0, 1)).ToUpper()
        if ($ch -notmatch '[A-Z]') { $ch = '#' }
        if (-not $byLetter.Contains($ch)) { $byLetter[$ch] = @() }
        $byLetter[$ch] += $g
    }

    # geometry
    $marginTop = 0.5; $marginBottom = 0.6; $marginSide = 0.5
    $tenantH = 1.6; $gapY = 0.5
    $rowLabelW = 3.0; $colW = 0.42; $rowH = 0.34; $headH = 2.6; $titleH = 0.5
    $fsTitle = 20; $fsLabel = 11; $fsCell = 10; $fsHead = 10
    # per-row colours: label = medium tint, value-cell fill = light tint,
    # borders/name = strong. Each group row cycles through these so rows are
    # easy to tell apart across the wide matrix.
    $rowMed    = @('#81C784', '#E57373', '#BA68C8', '#64B5F6', '#FFB74D', '#4DB6AC', '#F06292', '#A1887F', '#7986CB', '#FF8A65', '#4DD0E1', '#DCE775')
    $rowLight  = @('#C8E6C9', '#FFCDD2', '#E1BEE7', '#BBDEFB', '#FFE0B2', '#B2DFDB', '#F8BBD0', '#D7CCC8', '#C5CAE9', '#FFCCBC', '#B2EBF2', '#F0F4C3')
    $rowStrong = @('#2E7D32', '#C62828', '#6A1B9A', '#1565C0', '#EF6C00', '#00695C', '#AD1457', '#5D4037', '#283593', '#D84315', '#00838F', '#9E9D24')
    # per-department header colours: interleaved so neighbouring columns never
    # share a similar hue (avoids a rainbow gradient of look-alike colours)
    $deptPalette = @('#EF9A9A', '#90CAF9', '#A5D6A7', '#FFCC80', '#CE93D8', '#80DEEA', '#FFF59D', '#B39DDB', '#C5E1A5', '#F48FB1', '#80CBC4', '#FFAB91', '#9FA8DA', '#E6EE9C', '#BCAAA4', '#81D4FA', '#FFE082', '#B0BEC5', '#D1C4E9', '#CFD8DC')

    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $today = Get-Date -Format 'yyyy-MM-dd'
    $written = @()

    foreach ($letter in $byLetter.Keys) {
        $grps = @($byLetter[$letter])
        # all departments (with members) for this letter's groups
        $deptSet = [ordered]@{}
        foreach ($g in $grps) { if ($tree[$g]) { foreach ($d in $tree[$g].Keys) { if ([int]$tree[$g][$d] -gt 0) { $deptSet[$d] = $true } } } }
        $depts = @($deptSet.Keys | Sort-Object)
        if (-not $depts) { continue }

        $gridW = $rowLabelW + $depts.Count * $colW
        $pageW = [math]::Max(11, $gridW + 2 * $marginSide)

        # data-source box: leave a margin from the page edges (path only -> cheap)
        $dsW = [math]::Min($pageW - 2 * $marginSide, 12)
        $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'GroupMembers' -WidthIn $dsW -FontSize 10 -PathOnly
        $dsH = $dsBlock.Height

        $matH  = $titleH + $headH + $grps.Count * $rowH
        $pageH = [math]::Max(8.5, $marginTop + $dsH + $gapY + $tenantH + $gapY + $matH + $marginBottom)

        $shapes = [System.Collections.Generic.List[object]]::new()
        $topY  = $pageH - $marginTop - $tenantH / 2
        $dsY = $topY - $tenantH / 2 - $gapY - $dsH / 2

        [void]$shapes.Add(@{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                             X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 })

        $tenantLines = if ($tenant) {
            @(
                @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
                @{ Text = "Groups starting with `"$letter`"  -  group x department"; Bold = $false; Align = 'start' }
                @{ BoldPrefix = 'Cell number = '; Text = 'group members who are in that department (empty = 0)'; Align = 'start' }
                @{ Text = "$($grps.Count) groups, $($depts.Count) departments"; Bold = $false; Align = 'start' }
                @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
            )
        } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
        [void]$shapes.Add(@{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                             X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=12; Icon=$tenantIcon })

        $gridLeft = ($pageW - $gridW) / 2
        $titleCy  = ($dsY - $dsH / 2 - $gapY) - $titleH / 2
        $titleText = "Groups starting with `"$letter`"   -   $($grps.Count) groups, $($depts.Count) departments"
        # widen the title band when the grid is narrow so the title never spills
        $titleBandW = [math]::Min($pageW - 2 * $marginSide, [math]::Max($gridW, (Measure-MapTextWidth $titleText $fsTitle) + 0.5))
        [void]$shapes.Add(@{ Id='ltitle'; Kind='Rectangle'
                             Lines=@(@{ Text = $titleText; Bold = $true; Align = 'start' })
                             LinesTop=$true; TopInset=0.12
                             X=$pageW/2; Y=$titleCy; W=$titleBandW; H=$titleH; Fill='#E6EAF0'; Line='#B8BFC7'; FontSize=$fsTitle })
        $headTop = $dsY - $dsH / 2 - $gapY - $titleH

        for ($c = 0; $c -lt $depts.Count; $c++) {
            $colCx = $gridLeft + $rowLabelW + $c * $colW + $colW / 2
            $dn = "$($depts[$c])"; if ($dn.Length -gt 44) { $dn = $dn.Substring(0, 43) + [char]0x2026 }
            [void]$shapes.Add(@{ Id="h$c"; Kind='Rectangle'; Text=$dn; Rotate=-90
                                 X=$colCx; Y=($headTop - $headH/2); W=($colW - 0.03); H=($headH - 0.06)
                                 Fill=$deptPalette[$c % $deptPalette.Count]; Line='#78909C'; FontSize=$fsHead })
        }

        for ($r = 0; $r -lt $grps.Count; $r++) {
            $grp = $grps[$r]
            $ci  = $r % $rowMed.Count
            $rowCy = $headTop - $headH - $r * $rowH - $rowH / 2
            $gname = "$grp"; if ($gname.Length -gt 42) { $gname = $gname.Substring(0, 41) + [char]0x2026 }
            [void]$shapes.Add(@{ Id="rl$r"; Kind='Rectangle'; Text=$gname
                                 X=($gridLeft + $rowLabelW/2); Y=$rowCy; W=($rowLabelW - 0.04); H=($rowH - 0.04)
                                 Fill=$rowMed[$ci]; Line=$rowStrong[$ci]; StrokeWidth='1.6'; FontSize=$fsLabel })
            for ($c = 0; $c -lt $depts.Count; $c++) {
                $v = if ($tree[$grp]) { [int]$tree[$grp][$depts[$c]] } else { 0 }
                $cellCx = $gridLeft + $rowLabelW + $c * $colW + $colW / 2
                if ($v -gt 0) {
                    # value cell: the department's colour, border = the group's colour
                    [void]$shapes.Add(@{ Id="c$r-$c"; Kind='Rectangle'; Text="$v"; CenterText=$true
                                         X=$cellCx; Y=$rowCy; W=($colW - 0.03); H=($rowH - 0.03)
                                         Fill=$deptPalette[$c % $deptPalette.Count]; Line=$rowStrong[$ci]; StrokeWidth='2'; FontSize=$fsCell })
                }
                else {
                    # empty cell: the group's own colour, so each row reads as a band
                    [void]$shapes.Add(@{ Id="c$r-$c"; Kind='Rectangle'; Text=''
                                         X=$cellCx; Y=$rowCy; W=($colW - 0.03); H=($rowH - 0.03)
                                         Fill=$rowLight[$ci]; Line=$rowLight[$ci]; FontSize=$fsCell })
                }
            }
        }

        $safe = if ($letter -eq '#') { '0-9 and symbols' } else { $letter }
        $svg = Join-Path $svgDir "Entra ID - Groups x departments - $safe $stamp.svg"
        Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null
        $written += $svg
    }

    Write-Host "  SVG (matrix): $($written.Count) files, one per letter, $($data.Groups.Count) groups" -ForegroundColor DarkGray
    return $svgDir
}
