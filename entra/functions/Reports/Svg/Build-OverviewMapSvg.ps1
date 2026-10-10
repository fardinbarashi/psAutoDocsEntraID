# Autodoc v1.0
function Build-OverviewMapSvg {
    <#
        A single combined overview. Two styles via -Style:
          'hub'      : tenant in the centre, one node per area (Licences, Groups,
                       Owners, App registrations, Enterprise apps, Conditional
                       Access), each with its headline count; arrows from the
                       tenant out to every area. Clean, board-level summary.
          'relations': the same area nodes plus relationship arrows between them
                       (Groups -> CA, Groups -> Enterprise apps, Owners -> Groups,
                       CA -> Enterprise apps), showing how the pieces connect.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [ValidateSet('hub','relations','tree')][string]$Style = 'hub'
    )

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $ic = { param($n) Get-MapIcon -IconSet $iconSet -Sku $n -IconMap $iconMap -PreferSvg }

    # ---- gather the headline numbers from each area (all optional) ----
    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $lic  = Get-LicenceMapData         -SourceFolder $SourceFolder
    $grp  = Get-GroupsMapData          -SourceFolder $SourceFolder
    $own  = Get-GroupOwnerMapData      -SourceFolder $SourceFolder
    $areg = Get-AppRegMapData          -SourceFolder $SourceFolder
    $eapp = Get-EnterpriseAppMapData   -SourceFolder $SourceFolder
    $ca   = Get-ConditionalAccessMapData -SourceFolder $SourceFolder

    # enterprise-app count consistent with the enterprise map, where Microsoft
    # first-party apps are hidden by default (falls back to all when the flag is
    # absent, i.e. on exports collected before IsMicrosoftApp was added)
    $eappFull  = Get-EnterpriseAppFullData -SourceFolder $SourceFolder
    $eappShown = if ($eappFull) { @($eappFull.Apps | Where-Object { "$($_.IsMicrosoftApp)".Trim() -ne 'Yes' }).Count } elseif ($eapp) { $eapp.Total } else { 0 }

    if ($Style -eq 'tree') {
        return Build-OverviewTree -SourceFolder $SourceFolder -Tenant $tenant `
                 -Lic $lic -Grp $grp -Own $own -Areg $areg -Eapp $eapp -EappShown $eappShown -Ca $ca `
                 -IconSet $iconSet -IconMap $iconMap
    }

    # extra areas the hub was missing (users with member/guest split, RBAC,
    # departments and office locations - same data the tree shows)
    $usr  = Get-UsersMapData      -SourceFolder $SourceFolder
    $rbac = Get-RbacMapData       -SourceFolder $SourceFolder
    $dept = Get-DepartmentMapData -SourceFolder $SourceFolder
    $dom  = Get-DomainMapData     -SourceFolder $SourceFolder
    $sspr = Get-PasswordResetData -SourceFolder $SourceFolder
    $members = 0; $guests = 0
    if ($usr) { foreach ($t in $usr.Types) { if ("$($t.Name)" -match 'Guest') { $guests = $t.Count } elseif ("$($t.Name)" -match 'Member') { $members = $t.Count } } }

    $areas = @()
    if ($lic)  { $areas += @{ Key='lic';  Label='Licences';          Sub="$($lic.Purchased.Count) purchased"; Icon=(& $ic 'SKU');                     Fill='#DCE7F3' } }
    if ($grp)  { $areas += @{ Key='grp';  Label='Groups';            Sub="$($grp.Total) groups";              Icon=(& $ic '10223-icon-service-Groups');                  Fill='#E7F0DC' } }
    if ($own)  { $areas += @{ Key='own';  Label='Group owners';      Sub="$($own.Owners.Count) owners";       Icon=(& $ic '10223-icon-service-GroupOwner');                   Fill='#F3E7DC' } }
    if ($areg) { $areas += @{ Key='areg'; Label='App registrations'; Sub="$($areg.Total) apps";               Icon=(& $ic '10232-icon-service-App-Registrations');       Fill='#EDDCF3' } }
    if ($eapp) { $areas += @{ Key='eapp'; Label='Enterprise apps';   Sub="$eappShown apps";               Icon=(& $ic '10225-icon-service-Enterprise-Applications'); Fill='#DCF3EF' } }
    if ($ca)   { $areas += @{ Key='ca';   Label='Conditional Access';Sub="$($ca.Total) policies";             Icon=(& $ic '10233-icon-service-Conditional-Access');      Fill='#F3DCE4' } }
    if ($usr)  { $areas += @{ Key='usr';  Label='Users'; Icon=(& $ic '10230-icon-service-Users'); Fill='#E7F3F0'
                              Lines=@( @{ Text='Users'; Bold=$true; Align='middle' }, @{ Text="$($usr.Total) users"; Bold=$false; Align='middle' }, @{ Text="$members members  -  $guests guests"; Bold=$false; Align='middle' } ) } }
    if ($rbac -and @($rbac.Roles).Count)       { $areas += @{ Key='rbac';   Label='RBAC roles';       Sub="$(@($rbac.Roles).Count) roles";              Icon=(& $ic '10340-icon-service-Entra-Identity-Roles-and-Administrators'); Fill='#F3E7EE' } }
    if ($dept -and @($dept.Departments).Count) { $areas += @{ Key='dept';   Label='Departments';      Sub="$(@($dept.Departments).Count) departments";  Icon=(& $ic 'department');      Fill='#E7EEF3' } }
    if ($dept -and @($dept.Offices).Count)     { $areas += @{ Key='office'; Label='Office locations'; Sub="$(@($dept.Offices).Count) offices";           Icon=(& $ic 'office-location'); Fill='#F3F0E7' } }
    if ($dom -and @($dom.Total))               { $areas += @{ Key='dom';    Label='Domains';          Sub="$($dom.Total) domains";                       Icon=(& $ic '10222-icon-service-Entra-Domain-Services'); Fill='#EDE7F3' } }
    if ($sspr)                                 { $areas += @{ Key='sspr';   Label='Password reset';   Sub='SSPR';                                        Icon=(& $ic 'sspr-user-lock'); Fill='#F3EEE7' } }

    if (-not $areas) { Write-Host "No data to build an overview from." -ForegroundColor Yellow; return }

    # ---- layout: tenant hub centred, areas on a ring around it ----
    $pageW = 15.5; $pageH = 11.0
    $cxHub = $pageW / 2; $cyHub = $pageH / 2
    $tenantW = 3.2; $tenantH = 1.2
    $nodeW = 2.6; $nodeH = 1.2
    $rx = 5.6; $ry = 4.0    # ring radii

    $shapes = @(); $links = @()
    $today = Get-Date -Format 'yyyy-MM-dd'

    # tenant hub (centre)
    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true;  Align = 'middle' }
            @{ Text = "$($tenant.PrimaryDomain)"; Bold = $false; Align = 'middle' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'middle' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'middle' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'middle' }) }
    $shapes += @{ Id='hub'; Kind='Rectangle'; Lines=$tenantLines
                  X=$cxHub; Y=$cyHub; W=$tenantW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=(& $ic 'Azure-Active-Directory') }

    # area nodes on the ring
    $n = $areas.Count
    for ($i = 0; $i -lt $n; $i++) {
        $ang = (-90 + ($i * 360.0 / $n)) * [math]::PI / 180.0   # start at top, clockwise
        $x = $cxHub + $rx * [math]::Cos($ang)
        $y = $cyHub + $ry * [math]::Sin($ang)
        $a = $areas[$i]; $a['X'] = $x; $a['Y'] = $y
        $aLines = if ($a.Lines) { $a.Lines } else { @(@{ Text=$a.Label; Bold=$true; Align='middle' }, @{ Text=$a.Sub; Bold=$false; Align='middle' }) }
        $shapes += @{ Id=$a.Key; Kind='Rectangle'; Lines=$aLines
                      X=$x; Y=$y; W=$nodeW; H=$nodeH; Fill=$a.Fill; Line='#888888'; FontSize=10; Icon=$a.Icon }
        # hub -> area arrow (always)
        $links += @{ From='hub'; To=$a.Key }
    }

    # relationship arrows (only in 'relations' style)
    if ($Style -eq 'relations') {
        $have = @{}; foreach ($a in $areas) { $have[$a.Key] = $true }
        $rel = @(
            @{ From='own';  To='grp'  }   # owners own groups
            @{ From='grp';  To='ca'   }   # groups targeted by CA
            @{ From='grp';  To='eapp' }   # groups assigned to enterprise apps
            @{ From='ca';   To='eapp' }   # CA protects apps
            @{ From='areg'; To='eapp' }   # app reg backs an enterprise app
        )
        foreach ($r in $rel) { if ($have[$r.From] -and $have[$r.To]) { $links += @{ From=$r.From; To=$r.To; Rel=$true } } }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $mapName = if ($Style -eq 'hub') { 'EntraID - Hub' } else { "Entra ID - Tenant overview - $Style" }
    $svg = Join-Path $svgDir "$mapName $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -Connector $links -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG overview ($Style): $($areas.Count) areas" -ForegroundColor DarkGray
    return $svg
}

function Build-OverviewTree {
    <#
        Expanded overview: tenant hub at the top, the area nodes in a row beneath
        it, and every item of each area listed in a column under its area node.
        Arrows: hub -> each area, and area -> each of its items. Shows the whole
        tenant as one tree. Grows downward, so any number of items fits.
    #>
    param(
        [string]$SourceFolder, $Tenant, $Lic, $Grp, $Own, $Areg, $Eapp, $EappShown, $Ca,
        $IconSet, [hashtable]$IconMap
    )
    $ic = { param($n) Get-MapIcon -IconSet $IconSet -Sku $n -IconMap $IconMap -PreferSvg }

    # ---- build the per-area item lists (name + key figure) ----
    $columns = @()

    if ($Lic) {
        $items = foreach ($s in (@($Lic.Purchased) + @($Lic.Viral))) {
            @{ Name = $s.Sku; Sub = "$($s.Consumed)/$($s.Enabled) used"; Icon = (& $ic 'SKU') }
        }
        $columns += @{ Key='lic'; Label='Licences'; Sub="$(@($Lic.Purchased).Count + @($Lic.Viral).Count) SKUs"; Icon=(& $ic 'SKU'); Fill='#DCE7F3'; Items=@($items) }
    }
    if ($Grp) {
        $items = foreach ($cat in $Grp.Categories) { foreach ($g in $cat.Group) {
            $src = if ($g.Synced) { 'synced' } else { 'cloud' }
            $mail = if ($g.MailOn) { 'mail' } else { 'no mail' }
            @{ Name = $g.Name; Sub = "$($g.Members) members - $src - $mail"; Icon = (& $ic 'groups') }
        } }
        $columns += @{ Key='grp'; Label='Groups'; Sub="$($Grp.Total) groups"; Icon=(& $ic '10223-icon-service-Groups'); Fill='#E7F0DC'; Items=@($items) }
    }
    if ($Own) {
        $items = foreach ($o in $Own.Owners) {
            @{ Name = $o.Owner; Sub = "owns $(@($o.Groups).Count) groups"; Icon = (& $ic 'users') }
        }
        $columns += @{ Key='own'; Label='Group owners'; Sub="$($Own.Owners.Count) owners"; Icon=(& $ic '10223-icon-service-GroupOwner'); Fill='#F3E7DC'; Items=@($items) }
    }
    $appFull = Get-AppFullData -SourceFolder $SourceFolder
    if ($Areg -and $appFull) {
        $items = foreach ($a in $appFull.Apps) {
            $det = Get-AppDetailLines -App $a -View 'appreg'
            @{ Name = "$($a.AppName)"; Sub = ''; Detail = @($det); Icon = (& $ic 'app-registrations') }
        }
        $columns += @{ Key='areg'; Label='App registrations'; Sub="$($appFull.Total) apps"; Icon=(& $ic '10232-icon-service-App-Registrations'); Fill='#EDDCF3'; Items=@($items) }
    }
    if ($Eapp) {
        $eFull = Get-EnterpriseAppFullData -SourceFolder $SourceFolder
        $eApps = if ($eFull) { @($eFull.Apps | Where-Object { "$($_.IsMicrosoftApp)".Trim() -ne 'Yes' }) } else { @() }
        $items = foreach ($a in $eApps) {
            $det = Get-AppDetailLines -App $a -View 'enterprise'
            @{ Name = "$($a.AppName)"; Sub = ''; Detail = @($det); Icon = (& $ic 'enterprise-applications') }
        }
        $columns += @{ Key='eapp'; Label='Enterprise apps'; Sub="$(@($eApps).Count) apps"; Icon=(& $ic '10225-icon-service-Enterprise-Applications'); Fill='#DCF3EF'; Items=@($items) }
    }
    if ($Ca) {
        $caFull = Get-CaFullData -SourceFolder $SourceFolder
        $items = if ($caFull) {
            foreach ($p in $caFull.Policies) {
                # everything except the who-parts (users/roles/groups) - those are
                # rendered as a readable WHO block (bold name + UPN) in the item loop
                $rest = @(Get-CaDetailLines -Policy $p | Where-Object { $_ -notmatch '^(Include|Exclude) (users|roles|groups)\b' })
                $state = "$($p.State)".Trim()
                @{ Name = "$($p.DisplayName)  [$state]"; Sub = ''; Detail = $rest; Icon = (& $ic 'conditional-access')
                   CaWho = @{ IncludeUsers = "$($p.IncludeUsersResolved)"; ExcludeUsers = "$($p.ExcludeUsersResolved)"; Roles = "$($p.IncludeRolesResolved)"; Groups = "$($p.IncludeGroupsResolved)" } }
            }
        } else { @() }
        $columns += @{ Key='ca'; Label='Conditional Access'; Sub="$($Ca.Total) policies"; Icon=(& $ic '10233-icon-service-Conditional-Access'); Fill='#F3DCE4'; Items=@($items) }
    }

    # Departments and Office locations (from GroupMembers), each with member counts
    $dept = Get-DepartmentMapData -SourceFolder $SourceFolder
    if ($dept -and @($dept.Departments).Count) {
        $items = foreach ($d in $dept.Departments) {
            $w = if ($d.Members -eq 1) { 'member' } else { 'members' }
            @{ Name = $d.Name; Sub = "$($d.Members) $w"; Icon = (& $ic 'users') }
        }
        $columns += @{ Key='dept'; Label='Departments'; Sub="$(@($dept.Departments).Count) departments"; Icon=(& $ic 'department'); Fill='#E7EEF3'; Items=@($items) }
    }
    if ($dept -and @($dept.Offices).Count) {
        $items = foreach ($o in $dept.Offices) {
            $w = if ($o.Members -eq 1) { 'member' } else { 'members' }
            @{ Name = $o.Name; Sub = "$($o.Members) $w"; Icon = (& $ic 'users') }
        }
        $columns += @{ Key='office'; Label='Office locations'; Sub="$(@($dept.Offices).Count) offices"; Icon=(& $ic 'office-location'); Fill='#F3F0E7'; Items=@($items) }
    }

    # Domains (from TenantInformationDomain): type + default flag
    $dom = Get-DomainMapData -SourceFolder $SourceFolder
    if ($dom -and @($dom.Total)) {
        $items = foreach ($ty in $dom.Types) { foreach ($d in $ty.Group) {
            $tag = if ($d.IsDefault) { ' [default]' } else { '' }
            @{ Name = "$($d.Name)$tag"; Sub = $d.Type; Icon = (& $ic 'Azure-Active-Directory') }
        } }
        $columns += @{ Key='dom'; Label='Domains'; Sub="$($dom.Total) domains"; Icon=(& $ic '10222-icon-service-Entra-Domain-Services'); Fill='#EDE7F3'; Items=@($items) }
    }

    # RBAC roles: each role with the principals that hold it
    $rbac = Get-RbacMapData -SourceFolder $SourceFolder
    if ($rbac -and @($rbac.Roles).Count) {
        $items = foreach ($r in $rbac.Roles) {
            $det = foreach ($n in ($r.Group | Sort-Object @{ e = { $_.Kind } }, Principal)) {
                $p = "$($n.Principal)"
                if ($p -match '^(.*?)\s*<(.+)>\s*$') { $nm = $matches[1].Trim(); $upn = $matches[2].Trim() } else { $nm = $p; $upn = '' }
                $elig = if ($n.Kind -eq 'Eligible') { ' (eligible)' } else { '' }
                @{ Text = "$nm$elig"; Bold = $true }
                if ($upn) { @{ Text = "UPN: $upn"; Bold = $false } }
            }
            @{ Name = "$($r.Name)  ($(@($r.Group).Count))"; Sub = ''; Detail = @($det); Icon = (& $ic 'users') }
        }
        $columns += @{ Key='rbac'; Label='RBAC roles'; Sub="$(@($rbac.Roles).Count) roles"; Icon=(& $ic '10340-icon-service-Entra-Identity-Roles-and-Administrators'); Fill='#F3E7EE'; Items=@($items) }
    }

    # Users: summary + the key breakdowns as detail lines
    $usr = Get-UsersMapData -SourceFolder $SourceFolder
    if ($usr) {
        $items = @()
        $items += @{ Name = "Accounts ($($usr.Total))"; Sub = ''; Detail = @(
            "Enabled: $($usr.Enabled)  Disabled: $($usr.Disabled)"
            "Without mailbox: $($usr.NoMail)"
            "Multi-method auth: $($usr.MultiAuth)"
        ); Icon = (& $ic 'users') }
        foreach ($t in $usr.Types)   { $items += @{ Name = "Type: $($t.Name)";   Sub = "$($t.Count) users"; Icon = (& $ic 'users') } }
        foreach ($d in $usr.Domains) { $items += @{ Name = $d.Name;               Sub = "$($d.Count) users"; Icon = (& $ic 'users') } }
        foreach ($a in $usr.Auth)    { $items += @{ Name = "Auth: $($a.Name)";    Sub = "$($a.Count) users"; Icon = (& $ic 'users') } }
        $columns += @{ Key='usr'; Label='Users'; Sub="$($usr.Total) users"; Icon=(& $ic '10230-icon-service-Users'); Fill='#E7F3F0'; Items=@($items) }
    }

    # Password reset (SSPR): one config item with all settings as detail lines
    $sspr = Get-PasswordResetData -SourceFolder $SourceFolder
    if ($sspr) {
        $det = @(Get-PasswordResetLines -Config $sspr.Config)
        $items = @(@{ Name = 'SSPR configuration'; Sub = ''; Detail = $det; Icon = (& $ic 'sspr-user-lock') })
        $columns += @{ Key='sspr'; Label='Password reset'; Sub='SSPR'; Icon=(& $ic 'sspr-user-lock'); Fill='#F3EEE7'; Items=@($items) }
    }

    if (-not $columns) { Write-Host "No data to build an overview from." -ForegroundColor Yellow; return }

    # ---- layout ----
    $colW = 2.7; $colGap = 0.5
    $areaH = 0.9; $itemH = 0.62; $itemGap = 0.16
    $tenantW = 3.4; $tenantH = 1.2
    $marginTop = 0.5; $marginSide = 0.5; $marginBottom = 0.6
    $hubToAreaGap = 0.9; $areaToItemGap = 0.6

    $nCols = $columns.Count
    $pageW = [math]::Max(11, $marginSide * 2 + $nCols * $colW + ($nCols - 1) * $colGap)

    # each item's height depends on how many detail lines it has (name + sub + details)
    $lineH = 0.19
    $fsItem = 9
    # items have no icon (only the top category boxes do), so text uses the full
    # card width; wrap to that so nothing spills outside
    $itemAvail = ($colW - 0.18) * 0.97
    foreach ($col in $columns) {
        foreach ($it in $col.Items) {
            $ls = @()
            foreach ($p in @(Split-MapTextToWidth -Text "$($it.Name)" -AvailIn $itemAvail -FontSize $fsItem)) { $ls += @{ Text = $p; Bold = $true; Align = 'start' } }
            if ($it.CaWho) {
                # Conditional Access: each targeted user as a bold name + "UPN: ..." line
                $ls += @(Get-CaWhoLines -IncludeUsers $it.CaWho.IncludeUsers -ExcludeUsers $it.CaWho.ExcludeUsers -Roles $it.CaWho.Roles -Groups $it.CaWho.Groups -AvailIn $itemAvail -FontSize $fsItem)
            }
            if ($it.Sub) { foreach ($p in @(Split-MapTextToWidth -Text "$($it.Sub)" -AvailIn $itemAvail -FontSize $fsItem)) { $ls += @{ Text = $p; Bold = $false; Align = 'start' } } }
            if ($it.Detail) {
                foreach ($d in @($it.Detail)) {
                    $dtext = if ($d -is [hashtable]) { "$($d.Text)" } else { "$d" }
                    $dbold = if ($d -is [hashtable]) { [bool]$d.Bold } else { $false }
                    foreach ($p in @(Split-MapTextToWidth -Text $dtext -AvailIn $itemAvail -FontSize $fsItem)) { $ls += @{ Text = $p; Bold = $dbold; Align = 'start' } }
                }
            }
            $it['Lines'] = $ls
            $it['H'] = [math]::Max($itemH, 0.16 + $ls.Count * $lineH + 0.12)
        }
    }

    # tallest column (by summed item heights) decides page height
    $colBlockHeights = foreach ($col in $columns) {
        $sum = 0.0; foreach ($it in $col.Items) { $sum += $it.H + $itemGap }
        [math]::Max($itemH, $sum - $itemGap)
    }
    $colBlockH = ($colBlockHeights | Measure-Object -Maximum).Maximum
    if (-not $colBlockH) { $colBlockH = $itemH }
    $pageH = $marginTop + $tenantH + $hubToAreaGap + $areaH + $areaToItemGap + $colBlockH + $marginBottom
    $pageH = [math]::Max(8.5, $pageH)

    $shapes = @(); $links = @()
    $today = Get-Date -Format 'yyyy-MM-dd'

    # tenant hub, top centre
    $cxHub = $pageW / 2
    $cyHub = $pageH - $marginTop - $tenantH / 2
    $tenantLines = if ($Tenant) {
        @(
            @{ Text = "$($Tenant.TenantName)"; Bold = $true;  Align = 'middle' }
            @{ Text = "$($Tenant.PrimaryDomain)"; Bold = $false; Align = 'middle' }
            @{ Text = (Get-TenantHybridLabel -Tenant $Tenant); Bold = $false; Align = 'middle' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'middle' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'middle' }) }
    $shapes += @{ Id='hub'; Kind='Rectangle'; Lines=$tenantLines
                  X=$cxHub; Y=$cyHub; W=$tenantW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=(& $ic 'Azure-Active-Directory') }

    # area row + item columns
    $gridW = $nCols * $colW + ($nCols - 1) * $colGap
    $left  = ($pageW - $gridW) / 2
    $areaCy = $cyHub - $tenantH / 2 - $hubToAreaGap - $areaH / 2
    $itemsTopY = $areaCy - $areaH / 2 - $areaToItemGap

    for ($c = 0; $c -lt $nCols; $c++) {
        $col = $columns[$c]
        $cx = $left + $c * ($colW + $colGap) + $colW / 2

        # area node
        $shapes += @{ Id=$col.Key; Kind='Rectangle'
                      Lines=@(@{ Text=$col.Label; Bold=$true; Align='middle' }, @{ Text=$col.Sub; Bold=$false; Align='middle' })
                      X=$cx; Y=$areaCy; W=$colW; H=$areaH; Fill=$col.Fill; Line='#888888'; FontSize=11; Icon=$col.Icon }
        $links += @{ From='hub'; To=$col.Key }

        # items under the area node
        $iy = $itemsTopY
        $idx = 0
        foreach ($it in $col.Items) {
            $id = "$($col.Key)_$idx"
            $ih = $it.H
            $cyItem = $iy - $ih / 2
            $shapes += @{ Id=$id; Kind='Rectangle'; Lines=$it.Lines; LinesTop=$true; TopInset=0.10
                          X=$cx; Y=$cyItem; W=$colW; H=$ih; Fill=$col.Fill; Line='#B8BFC7'; FontSize=$fsItem }
            if ($idx -eq 0) { $links += @{ From=$col.Key; To=$id } }
            $iy -= ($ih + $itemGap)
            $idx++
        }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "EntraID - Overview tree $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -Connector $links -PageWidth $pageW -PageHeight $pageH | Out-Null

    $totalItems = ($columns | ForEach-Object { $_.Items.Count } | Measure-Object -Sum).Sum
    Write-Host "  SVG overview (tree): $($columns.Count) areas, $totalItems items" -ForegroundColor DarkGray
    return $svg
}
