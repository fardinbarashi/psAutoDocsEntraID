# Autodoc v1.0
function Build-EntraWordReport {
    <#
        Builds a Word system-documentation report (.docx) from an export, in pure
        PowerShell - it writes the Office Open XML parts and zips them, so it needs
        neither Word installed nor any extra module.

        This first milestone lays down the whole document skeleton (cover page,
        table of contents, and every section heading so the TOC is complete) and
        fills section 1 (Inledning / Versionshantering). The remaining sections are
        placeholders that we fill in area by area.

        Metadata that isn't in Microsoft Graph (author, customer, contacts, logos,
        version) comes from Entra\Config\WordReportProfile.psd1; anything blank is
        shown as a [placeholder] so it's easy to complete in Word.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExportsRoot,
        [string]$SourceFolder,
        [string]$Language = 'sv',
        [string]$ConsultantLogo,
        [string]$CustomerLogo
    )

    $src = if ($SourceFolder) { $SourceFolder } else { Resolve-AutodocSource -ExportsRoot $ExportsRoot }
    if (-not $src) { Write-Host "No export to build a Word report from." -ForegroundColor Yellow; return }
    $lang = if ("$Language" -eq 'en') { 'en' } else { 'sv' }

    # ---- profile (metadata that isn't in Graph) ----
    $entraRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot))   # Functions\Reports\Word -> Entra
    # document text - editable per language under files\cache\wordTemplates\<lang> (also matches eng/swe)
    $txtAll = @{}
    $txtFile = Join-Path $entraRoot 'Config\wordText.json'
    if (Test-Path $txtFile) { try { $txtAll = Get-Content -LiteralPath $txtFile -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable } catch { $txtAll = @{} } }
    function T($k) { $e = $txtAll[$k]; if ($e -and "$($e[$lang])") { "$($e[$lang])" } elseif ($e -and "$($e['en'])") { "$($e['en'])" } else { "[$k]" } }
    $prof = @{}
    $profilePath = Join-Path $entraRoot 'Config\WordReportProfile.json'
    if (Test-Path $profilePath) { try { $prof = Get-Content -LiteralPath $profilePath -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable } catch { $prof = @{} } }
    # real values live centrally in Config\WordReportProfile.json; a blank field falls back to the per-language default/placeholder in Config\wordText.json
    $consultant = if ("$($prof.ConsultantCompany)") { "$($prof.ConsultantCompany)" } else { T 'ConsultantPlaceholder' }
    $author     = if ("$($prof.Author)") { "$($prof.Author)" } else { T 'AuthorPlaceholder' }
    $customer   = if ("$($prof.CustomerName)") { "$($prof.CustomerName)" } else { T 'CustomerPlaceholder' }
    $docId      = if ("$($prof.DocumentId)") { "$($prof.DocumentId)" } else { T 'DocIdDefault' }
    $version    = if ("$($prof.Version)")    { "$($prof.Version)" }    else { '1.0' }
    $firstChg   = if ("$($prof.FirstChange)"){ "$($prof.FirstChange)" }else { T 'FirstChangeDefault' }
    $today      = Get-Date -Format 'yyyy-MM-dd'

    # logos: GUI parameter -> cached logo (files\cache\wordTemplates\template\img) -> profile -> text placeholder
    $logoImgDir = Join-Path $entraRoot 'files\cache\wordTemplates\template\img'
    function CachedLogo($base) { if (Test-Path $logoImgDir) { $f = Get-ChildItem -Path $logoImgDir -Filter "$base.*" -ErrorAction SilentlyContinue | Select-Object -First 1; if ($f) { return $f.FullName } } ; return '' }
    $consultantLogo = if ($ConsultantLogo -and (Test-Path $ConsultantLogo)) { $ConsultantLogo } elseif (CachedLogo 'companylogo') { CachedLogo 'companylogo' } elseif ("$($prof.ConsultantLogo)" -and (Test-Path "$($prof.ConsultantLogo)")) { "$($prof.ConsultantLogo)" } else { '' }
    $customerLogo   = if ($CustomerLogo   -and (Test-Path $CustomerLogo))   { $CustomerLogo }   elseif (CachedLogo 'clientlogo')  { CachedLogo 'clientlogo' }  elseif ("$($prof.CustomerLogo)"   -and (Test-Path "$($prof.CustomerLogo)"))   { "$($prof.CustomerLogo)" }   else { '' }

    $tenant = @(Get-AutodocExportData -ExportFolder $src -BaseName 'TenantInformation') | Select-Object -First 1
    $users   = @(Get-AutodocExportData -ExportFolder $src -BaseName 'UserInformation')
    $groups  = @(Get-AutodocExportData -ExportFolder $src -BaseName 'EntraGroups')
    $apps    = @(Get-AutodocExportData -ExportFolder $src -BaseName 'AppRegEnterpriseApps')
    $domains = @(Get-AutodocExportData -ExportFolder $src -BaseName 'TenantInformationDomain')
    $licenses = @(Get-AutodocExportData -ExportFolder $src -BaseName 'LicensesInformation')
    $eapps    = @(Get-AutodocExportData -ExportFolder $src -BaseName 'EnterpriseApps')
    $rbac     = @(Get-AutodocExportData -ExportFolder $src -BaseName 'RBAC')
    $sspr     = @(Get-AutodocExportData -ExportFolder $src -BaseName 'PasswordReset') | Select-Object -First 1
    $caRoot   = @(Get-AutodocExportData -ExportFolder $src -BaseName 'ConditionalAccess') | Select-Object -First 1
    $caPol    = @($caRoot.ConditionalAccessPolicies)
    $namedLoc = @($caRoot.ConditionalAccessNamedLocations)
    $vals     = Get-WordTemplateValues -SourceFolder $src
    # counts computed directly from the raw data (independent of the resolver)
    $uGuests   = @($users | Where-Object { "$($_.UserType)" -eq 'Guest' }).Count
    $uMembers  = @($users | Where-Object { "$($_.UserType)" -ne 'Guest' }).Count
    $uEnabled  = @($users | Where-Object { $_.AccountEnabled }).Count
    $uDisabled = @($users | Where-Object { -not $_.AccountEnabled }).Count
    $rbacActive   = @($rbac | Where-Object { $_.RecordType -eq 'RoleAssignmentActive' }).Count     # same rule as the Excel RBAC sheet
    $rbacEligible = @($rbac | Where-Object { $_.RecordType -eq 'RoleAssignmentEligible' }).Count
    $caEnabled  = @($caPol | Where-Object { "$($_.State)" -eq 'enabled' }).Count
    $caReport   = @($caPol | Where-Object { "$($_.State)" -match 'report' }).Count
    $caDisabled = @($caPol | Where-Object { "$($_.State)" -eq 'disabled' }).Count
    $appMulti   = @($apps | Where-Object { "$($_.SignInAudience)" -match 'MultipleOrgs|PersonalMicrosoftAccount' }).Count
    $appSingle  = (@($apps).Count - $appMulti)
    $ssoYes = @($eapps | Where-Object { $_.SSOConfigured -eq $true -or "$($_.SSOConfigured)" -eq 'True' }).Count
    $ssoNo  = (@($eapps).Count - $ssoYes)
    # ---- tenant-structure figures (all from the exported JSON) ----
    function Num($n) { ([int]$n).ToString('N0', [Globalization.CultureInfo]::InvariantCulture) -replace ',', ' ' }
    function Pct($v) { $s = '{0:0.#}' -f [double]$v; if ($lang -eq 'sv') { $s = $s -replace '\.', ',' }; "$s %" }
    $rbacAssign = @($rbac | Where-Object { "$($_.RecordType)" -like 'RoleAssignment*' })
    if (-not $rbacAssign.Count) { $rbacAssign = @($rbac | Where-Object { "$($_.RoleDefinitionName)" }) }

    # ---- OOXML building blocks ----
    function XmlEsc([string]$s) { if ($null -eq $s) { return '' }; $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' }
    function Par {
        param([string]$Text = '', [string]$Style = '', [string]$Align = '', [switch]$Bold, [int]$Size = 0, [string]$Color = '')
        $ppr = ''; if ($Style) { $ppr += "<w:pStyle w:val='$Style'/>" }; if ($Align) { $ppr += "<w:jc w:val='$Align'/>" }
        $rpr = ''; if ($Bold) { $rpr += '<w:b/>' }; if ($Size) { $rpr += "<w:sz w:val='$Size'/><w:szCs w:val='$Size'/>" }; if ($Color) { $rpr += "<w:color w:val='$Color'/>" }
        $ppr = if ($ppr) { "<w:pPr>$ppr</w:pPr>" } else { '' }; $rpr = if ($rpr) { "<w:rPr>$rpr</w:rPr>" } else { '' }
        "<w:p>$ppr<w:r>$rpr<w:t xml:space='preserve'>$(XmlEsc $Text)</w:t></w:r></w:p>"
    }
    # bold red heading + red line, used where the export SHOULD have data but none was found
    function RedMissing($heading) {
        (Par -Text "$($heading):" -Bold -Color 'C00000') + (Par -Text (T 'NoData') -Color 'C00000')
    }
    # neutral grey note (text comes from the language file)
    function GreyPar($k) { Par -Text (T $k) -Color '808080' }
    function Manual { GreyPar 'Manual' }
    function H1($t) { Par -Text $t -Style 'Heading1' }
    function Note($t) { Par -Text $t -Color '808080' }
    function BoldRuns($text, $phrases) {
        $rest = "$text"; $runs = ''
        while ($rest.Length) {
            $bestIdx = -1; $bestLen = 0
            foreach ($ph in $phrases) { if (-not $ph) { continue }; $i = $rest.IndexOf("$ph"); if ($i -ge 0 -and ($bestIdx -lt 0 -or $i -lt $bestIdx -or ($i -eq $bestIdx -and "$ph".Length -gt $bestLen))) { $bestIdx = $i; $bestLen = "$ph".Length } }
            if ($bestIdx -lt 0) { $runs += "<w:r><w:t xml:space='preserve'>$(XmlEsc $rest)</w:t></w:r>"; break }
            if ($bestIdx -gt 0) { $runs += "<w:r><w:t xml:space='preserve'>$(XmlEsc $rest.Substring(0, $bestIdx))</w:t></w:r>" }
            $runs += "<w:r><w:rPr><w:b/></w:rPr><w:t xml:space='preserve'>$(XmlEsc $rest.Substring($bestIdx, $bestLen))</w:t></w:r>"
            $rest = $rest.Substring($bestIdx + $bestLen)
        }
        $runs
    }
    function ParBold($text, $phrases) { "<w:p>$(BoldRuns $text $phrases)</w:p>" }
    function BulletBold($text, $phrases) { "<w:p><w:pPr><w:ind w:left='420' w:hanging='220'/></w:pPr><w:r><w:t xml:space='preserve'>&#8226;  </w:t></w:r>$(BoldRuns $text $phrases)</w:p>" }
    function H2($t) { Par -Text $t -Style 'Heading2' }
    function PageBreak { "<w:p><w:r><w:br w:type='page'/></w:r></w:p>" }
    function Tbl {
        param([object[]]$Rows, [int[]]$ColW, [switch]$Header, [switch]$BoldFirstCol)
        $total = ($ColW | Measure-Object -Sum).Sum
        $grid = ($ColW | ForEach-Object { "<w:gridCol w:w='$_'/>" }) -join ''
        $trs = ''
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            $tcs = ''
            for ($c = 0; $c -lt $Rows[$r].Count; $c++) {
                $hdr = ($Header -and $r -eq 0)
                $dataIdx = if ($Header) { $r - 1 } else { $r }
                $shd = if ($hdr) { "<w:shd w:val='clear' w:color='auto' w:fill='4472C4'/>" } elseif ($dataIdx % 2 -eq 1) { "<w:shd w:val='clear' w:color='auto' w:fill='F2F2F2'/>" } else { '' }
                $b = ($hdr -or ($BoldFirstCol -and $c -eq 0))
                $col = if ($hdr) { "<w:color w:val='FFFFFF'/>" } else { '' }
                $rpr = if ($b -or $col) { "<w:rPr>$(if($b){'<w:b/>'})$col</w:rPr>" } else { '' }
                $tcs += "<w:tc><w:tcPr><w:tcW w:w='$($ColW[$c])' w:type='dxa'/>$shd</w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r>$rpr<w:t xml:space='preserve'>$(XmlEsc "$($Rows[$r][$c])")</w:t></w:r></w:p></w:tc>"
            }
            $trs += "<w:tr>$tcs</w:tr>"
        }
        "<w:tbl><w:tblPr><w:tblW w:w='$total' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='BFBFBF'/><w:left w:val='single' w:sz='4' w:color='BFBFBF'/><w:bottom w:val='single' w:sz='4' w:color='BFBFBF'/><w:right w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideH w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideV w:val='single' w:sz='4' w:color='BFBFBF'/></w:tblBorders></w:tblPr><w:tblGrid>$grid</w:tblGrid>$trs</w:tbl>"
    }

    # ---- image building blocks (logos) ----
    $images = [System.Collections.Generic.List[object]]::new()
    $imgSeq = 0
    function Get-ImageSize($path) {
        try {
            $b = [System.IO.File]::ReadAllBytes($path)
            if ($b.Length -gt 24 -and $b[0] -eq 0x89 -and $b[1] -eq 0x50) {                     # PNG
                return @{ W = (([int]$b[16] -shl 24) -bor ([int]$b[17] -shl 16) -bor ([int]$b[18] -shl 8) -bor [int]$b[19]); H = (([int]$b[20] -shl 24) -bor ([int]$b[21] -shl 16) -bor ([int]$b[22] -shl 8) -bor [int]$b[23]) }
            }
            if ($b.Length -gt 4 -and $b[0] -eq 0xFF -and $b[1] -eq 0xD8) {                        # JPEG
                $i = 2
                while ($i -lt $b.Length - 8) {
                    if ($b[$i] -ne 0xFF) { $i++; continue }
                    $m = $b[$i + 1]
                    if ($m -ge 0xC0 -and $m -le 0xC3) { return @{ W = (([int]$b[$i + 7] -shl 8) -bor [int]$b[$i + 8]); H = (([int]$b[$i + 5] -shl 8) -bor [int]$b[$i + 6]) } }
                    $i += 2 + (([int]$b[$i + 2] -shl 8) -bor [int]$b[$i + 3])
                }
            }
        } catch {}
        return @{ W = 0; H = 0 }
    }
    function LogoDrawing($path, $maxWidthInch) {
        if (-not $path -or -not (Test-Path $path)) { return $null }
        $script:imgSeq++
        $ext = ([System.IO.Path]::GetExtension($path)).TrimStart('.').ToLower(); if ($ext -eq 'jpeg') { $ext = 'jpg' }
        $mediaName = "image$($script:imgSeq).$ext"; $relId = "rIdImg$($script:imgSeq)"
        $d = Get-ImageSize $path
        $wIn = $maxWidthInch; $hIn = if ($d.W -gt 0) { [math]::Round($maxWidthInch * $d.H / $d.W, 3) } else { [math]::Round($maxWidthInch * 0.4, 3) }
        $wEmu = [int]($wIn * 914400); $hEmu = [int]($hIn * 914400)
        $images.Add(@{ Path = $path; MediaName = $mediaName; RelId = $relId })
        $id = $script:imgSeq
        "<w:drawing><wp:inline distT='0' distB='0' distL='0' distR='0' xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'><wp:extent cx='$wEmu' cy='$hEmu'/><wp:docPr id='$id' name='Logo$id'/><a:graphic xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/picture'><pic:pic xmlns:pic='http://schemas.openxmlformats.org/drawingml/2006/picture'><pic:nvPicPr><pic:cNvPr id='$id' name='Logo$id'/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed='$relId'/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='$wEmu' cy='$hEmu'/></a:xfrm><a:prstGeom prst='rect'><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"
    }

    function Get-MapPng($namePattern, $builder, $maxWidthInch) {
        $svgDir = Join-Path $src 'Report\svg'
        $find = { Get-ChildItem -Path $svgDir -Filter "*$namePattern*.svg" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -Last 1 }
        $svg = & $find
        if (-not $svg -and $builder) { try { & $builder | Out-Null } catch { } ; $svg = & $find }
        if (-not $svg) { return $null }
        $png = Join-Path ([System.IO.Path]::GetTempPath()) ("wrmap_" + [guid]::NewGuid().ToString('N') + '.png')
        if (Convert-SvgToPng -SvgPath $svg.FullName -PngPath $png) { return (LogoDrawing $png $maxWidthInch) }
        return $null
    }

    # ---- simple horizontal bar chart (SVG -> PNG) for KPI distributions ----
    # ---- KPI subsection: H2 + 2-col table (+ optional bar) from label/value pairs ----
    # ---- body ----
    $sb = [System.Text.StringBuilder]::new()

    # KPI cards: the same figures as the cards at the top of each Excel KPI sheet
    $kpiCards = try { Get-KpiCardData -SourceFolder $src } catch { [ordered]@{} }
    function CardVal($area, $label) {
        $c = @($kpiCards[$area] | Where-Object { $_.Label -eq $label }) | Select-Object -First 1
        if ($c -and "$($c.Value)" -match '^\d+$') { Num $c.Value } elseif ($c) { "$($c.Value)" } else { '' }
    }
    function KpiCards($area) {
        $cards = @($kpiCards[$area])
        if (-not $cards.Count) { return '' }
        $per = if ($cards.Count -le 6) { $cards.Count } else { [math]::Ceiling($cards.Count / 2) }
        $xml = ''
        for ($i = 0; $i -lt $cards.Count; $i += $per) {
            $chunk = @($cards[$i..([math]::Min($i + $per, $cards.Count) - 1)])
            $w = [int](9000 / $chunk.Count)
            $xml += Tbl -Header -Rows @(@($chunk | ForEach-Object { $_.Label }), @($chunk | ForEach-Object { CardVal $area $_.Label })) -ColW @(1..$chunk.Count | ForEach-Object { $w })
            $xml += Par ''
        }
        $xml
    }

    # cover
    # cover table: label | value | consultant logo (logo cell spans both rows)
    $cLogo = LogoDrawing $consultantLogo 1.9
    $logoRun = if ($cLogo) { "<w:r>$cLogo</w:r>" } else { "<w:r><w:t xml:space='preserve'>$(XmlEsc "[$consultant logo]")</w:t></w:r>" }
    $cw = @(1900, 2600, 3000)
    $cvTc = { param($w, $text, $bold, $fill) "<w:tc><w:tcPr><w:tcW w:w='$w' w:type='dxa'/>$(if($fill){"<w:shd w:val='clear' w:color='auto' w:fill='$fill'/>"})<w:vAlign w:val='center'/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r>$(if($bold){'<w:rPr><w:b/></w:rPr>'})<w:t xml:space='preserve'>$(XmlEsc $text)</w:t></w:r></w:p></w:tc>" }
    $logoTop = "<w:tc><w:tcPr><w:tcW w:w='$($cw[2])' w:type='dxa'/><w:vMerge w:val='restart'/><w:vAlign w:val='center'/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/><w:jc w:val='center'/></w:pPr>$logoRun</w:p></w:tc>"
    $logoCont = "<w:tc><w:tcPr><w:tcW w:w='$($cw[2])' w:type='dxa'/><w:vMerge/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr></w:p></w:tc>"
    $coverRows = "<w:tr>$(& $cvTc $cw[0] 'Document Type' $true '')$(& $cvTc $cw[1] 'System Documentation' $false '')$logoTop</w:tr>" +
                 "<w:tr>$(& $cvTc $cw[0] 'Area' $true 'F2F2F2')$(& $cvTc $cw[1] 'Entra ID' $false 'F2F2F2')$logoCont</w:tr>"
    $coverGrid = ($cw | ForEach-Object { "<w:gridCol w:w='$_'/>" }) -join ''
    [void]$sb.Append("<w:tbl><w:tblPr><w:tblW w:w='$(($cw|Measure-Object -Sum).Sum)' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='BFBFBF'/><w:left w:val='single' w:sz='4' w:color='BFBFBF'/><w:bottom w:val='single' w:sz='4' w:color='BFBFBF'/><w:right w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideH w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideV w:val='single' w:sz='4' w:color='BFBFBF'/></w:tblBorders></w:tblPr><w:tblGrid>$coverGrid</w:tblGrid>$coverRows</w:tbl>")
    1..5 | ForEach-Object { [void]$sb.Append((Par '')) }
    [void]$sb.Append((Par -Text $docId -Style 'Title' -Align 'center'))
    [void]$sb.Append((Par -Text $customer -Align 'center' -Size 36))
    1..7 | ForEach-Object { [void]$sb.Append((Par '')) }
    $kLogo = LogoDrawing $customerLogo 2.6
    if ($kLogo) { [void]$sb.Append("<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r>$kLogo</w:r></w:p>") } else { [void]$sb.Append((Par -Text (T 'Cover_CustomerLogo') -Align 'center')) }
    [void]$sb.Append((PageBreak))

    # table of contents (populated when the field is updated in Word)
    [void]$sb.Append((Par -Text (T 'TOC_Title') -Style 'TOCHeading'))
    [void]$sb.Append("<w:p><w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText xml:space='preserve'> TOC \o &quot;1-3&quot; \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType='separate'/></w:r><w:r><w:t xml:space='preserve'>$(XmlEsc (T 'TOC_Placeholder'))</w:t></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r></w:p>")
    [void]$sb.Append((PageBreak))

    # ============ 1 Inledning ============
    $boldWords = @('Kundens', 'Kunden', 'Kund', 'Leverantören', 'Leverantör', 'Konsultbolaget', 'Konsultbolag', "$customer", "$consultant")
    [void]$sb.Append((H1 (T 'S1')))
    [void]$sb.Append((Par (T 'CH1_p1')))
    [void]$sb.Append((Par (T 'CH1_p2')))
    [void]$sb.Append((Par (T 'CH1_p3')))
    [void]$sb.Append((Par (T 'CH1_p4')))
    [void]$sb.Append((Par (T 'CH1_p5')))
    [void]$sb.Append((H2 (T 'CH1_1')))
    [void]$sb.Append((ParBold ((T 'CH1_1_body') -f $customer) $boldWords))
    [void]$sb.Append((BulletBold (T 'CH1_1_b1') $boldWords))
    [void]$sb.Append((BulletBold (T 'CH1_1_b2') $boldWords))
    [void]$sb.Append((BulletBold ((T 'CH1_1_b3') -f $consultant) $boldWords))
    [void]$sb.Append((H2 (T 'S1_1')))
    [void]$sb.Append((Par (T 'S1_1_body')))
    [void]$sb.Append((H2 (T 'CH1_3')))
    [void]$sb.Append((Tbl -Header -Rows @(@((T 'L_Document'), (T 'L_Reference')), @((T 'Ref_docLabel'), 'N/A')) -ColW @(5000, 4000)))
    [void]$sb.Append((H2 (T 'S1_2')))
    [void]$sb.Append((Tbl -Header -Rows @(@((T 'F_Version'), (T 'F_Date'), (T 'L_ChangedBy'), (T 'L_Change')), @($version, $today, $author, $firstChg)) -ColW @(1200, 1800, 3000, 3000)))

    # ============ 2 Förvaltning och support ============
    [void]$sb.Append((H1 (T 'S2')))
    [void]$sb.Append((H2 (T 'CH2_1')))
    [void]$sb.Append((Par (T 'CH2_1_body')))
    [void]$sb.Append((H2 (T 'CH2_2')))
    [void]$sb.Append((Tbl -Header -Rows @(
                @('Role', 'Name', 'Phone', 'Email'),
                @('Operations', 'Operationsdesk', 'XXX-XXX-XXX', 'autodoc.@contoso.se'),
                @('Support', 'Supportdesk', 'XXX-XXX-XXX', 'autodoc.@contoso.se')
            ) -ColW @(2000, 2500, 2200, 2800)))
    [void]$sb.Append((H2 (T 'S2_customer')))
    [void]$sb.Append((Tbl -Header -Rows @(
                @('Role', 'Name', 'Phone', 'Email'),
                @('', '', '', '')
            ) -ColW @(2000, 2500, 2200, 2800)))

    # ============ 3 Allmän systembeskrivning ============
    [void]$sb.Append((H1 (T 'CH3_general')))
    [void]$sb.Append((Par (T 'CH3_gen_p1')))
    [void]$sb.Append((Par (T 'CH3_gen_p2')))

    # ============ 4 Arkitektur och konfiguration ============
    [void]$sb.Append((H1 (T 'S3')))
    [void]$sb.Append((Par (T 'S3_intro')))

    [void]$sb.Append((H2 (T 'S3_1')))
    if ($tenant) {
        # Table 1 - tenant identity
        $trows = @(
            @('Category', 'Value'),
            @('Tenant', "$($tenant.TenantName)"),
            @('Tenant ID', "$($tenant.TenantId)"),
            @('Primary domain', "$($tenant.PrimaryDomain)"),
            @('Country', "$($tenant.CountryCode)")
        )
        [void]$sb.Append((Tbl -Header -BoldFirstCol -Rows $trows -ColW @(3000, 6000)))
        [void]$sb.Append((Par ''))
        # narrative summary (figures filled from the JSON)
        [void]$sb.Append((Par ((T 'CH4_text_p1') -f (CardVal 'Dashboard' 'Users'), (CardVal 'Dashboard' 'Devices'), (CardVal 'Dashboard' 'Assigned licences'), (CardVal 'Dashboard' 'Purchased SKUs'), (CardVal 'Dashboard' 'Unused units'))))
        [void]$sb.Append((Par ((T 'CH4_text_p2') -f (CardVal 'Dashboard' 'Total groups'), (CardVal 'Dashboard' 'Security groups'), (CardVal 'Dashboard' 'Microsoft 365 groups'), (CardVal 'Dashboard' 'Total app registrations'), (CardVal 'Dashboard' 'Apps with service principal'), (CardVal 'Dashboard' 'Distinct roles'), (CardVal 'Dashboard' 'Role assignments'), (CardVal 'Dashboard' 'Privileged assignments'))))
        [void]$sb.Append((Par ''))
        # Table 2 - key figures (two label/value pairs per row)
        $krows = @(, @('Category', 'Value', 'Category', 'Value'))
        foreach ($pair in @(@('Users', 'Total groups'), @('Devices', 'Security groups'), @('Assigned licences', 'Microsoft 365 groups'),
                @('Purchased SKUs', 'Total app registrations'), @('Unused units', 'Apps with service principal'),
                @('Distinct roles', 'Role assignments'), @('Privileged assignments', ''))) {
            $krows += , @($pair[0], (CardVal 'Dashboard' $pair[0]), $pair[1], $(if ($pair[1]) { CardVal 'Dashboard' $pair[1] } else { '' }))
        }
        [void]$sb.Append((Tbl -Header -Rows $krows -ColW @(2600, 1900, 2600, 1900)))
    }
    else { [void]$sb.Append((RedMissing (T 'S3_1'))) }

    [void]$sb.Append((H2 (T 'S3_2')))
    if ($tenant -and "$($tenant.HybridStatus)") {
        $envLabel = "$($tenant.HybridStatus)"
        $syncU = @($users | Where-Object { $_.OnPremisesSyncEnabled }).Count
        $syncG = @($groups | Where-Object { $_.OnPremisesSyncEnabled }).Count
        if ($envLabel -match 'Onprem') {
            [void]$sb.Append((ParBold ((T 'S3_2_hybrid') -f $envLabel, $syncU, $syncG) @('Vilket synkverktyg (Entra Connect eller Cloud Sync) som används framgår inte av katalogdatan.', 'Which sync tool (Entra Connect or Cloud Sync) is used is not exposed by the directory data.')))
        }
        else { [void]$sb.Append((Par ((T 'S3_2_cloud') -f $envLabel))) }
    }
    else { [void]$sb.Append((RedMissing (T 'S3_2'))) }

    [void]$sb.Append((H2 (T 'S3_3')))
    if ($domains.Count) {
        $domUsers = @{}
        $umdDom = try { Get-UsersMapData -SourceFolder $src } catch { $null }
        if ($umdDom) { foreach ($x in @($umdDom.Domains)) { $domUsers["$($x.Name)".ToLower()] = [int]$x.Count } }
        $drows = @(, @('Domain', 'Type', 'Default', 'Users'))
        foreach ($d in ($domains | Sort-Object @{ E = { - [int][bool]$_.IsDefault } }, Domain)) {
            $du = $domUsers["$($d.Domain)".ToLower()]
            $drows += , @("$($d.Domain)", "$($d.Type)", $(if ($d.IsDefault) { T 'L_Yes' } else { '' }), $(if ($du) { Num $du } else { '0' }))
        }
        [void]$sb.Append((Tbl -Header -Rows $drows -ColW @(4000, 1900, 1500, 1600)))
    }
    else { [void]$sb.Append((RedMissing (T 'S3_3'))) }

    # ============ 4 Systemkarta ============
    [void]$sb.Append((H1 (T 'S4')))
    [void]$sb.Append((Par (T 'S4_intro')))
    $hubDraw = Get-MapPng 'Hub' { Build-OverviewMapSvg -SourceFolder $src -Style hub } 6.3
    if ($hubDraw) { [void]$sb.Append("<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r>$hubDraw</w:r></w:p>") }
    else { [void]$sb.Append((RedMissing (T 'S4'))) }

    # ============ 5 Microsoft Entra ID (KPI dashboard + pointer to the Excel workbook) ============
    [void]$sb.Append((H1 (T 'S_KPI')))
    $kpiSvg = Build-KpiDashboardSvg -SourceFolder $src   # always rebuild so the image reflects current data
    if ($kpiSvg -and (Test-Path $kpiSvg)) {
        $kpiPng = Join-Path ([System.IO.Path]::GetTempPath()) ("wrdash_" + [guid]::NewGuid().ToString('N') + '.png')
        if (Convert-SvgToPng -SvgPath $kpiSvg -PngPath $kpiPng) { [void]$sb.Append("<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r>$(LogoDrawing $kpiPng 6.3)</w:r></w:p>"); [void]$sb.Append((Par '')) }
    }
    [void]$sb.Append((Par (T 'KPI_seeExcel')))

    # ============ 6 Systemkonfiguration och inställningar ============
    [void]$sb.Append((H1 (T 'CH6')))

    [void]$sb.Append((H2 (T 'S5_1')))
    [void]$sb.Append((Par (T 'Exp_Users')))
    if ($users.Count) {
        $urows = @(, @((T 'L_Category'), (T 'L_Count')))
        $urows += , @((T 'L_Members'), ('{0:N0}' -f $uMembers))
        $urows += , @((T 'L_GuestsB2B'), ('{0:N0}' -f $uGuests))
        $urows += , @((T 'L_Enabled'), ('{0:N0}' -f $uEnabled))
        $urows += , @((T 'L_Disabled'), ('{0:N0}' -f $uDisabled))
        [void]$sb.Append((Tbl -Header -Rows $urows -ColW @(5000, 4000)))
        $umd = try { Get-UsersMapData -SourceFolder $src } catch { $null }
        if ($umd -and @($umd.Domains).Count) {
            [void]$sb.Append((Par -Text 'Users per domain' -Bold))
            $ddrows = @(, @('Domain', 'Users'))
            foreach ($d in $umd.Domains) { $ddrows += , @("$($d.Name)", (Num $d.Count)) }
            [void]$sb.Append((Tbl -Header -Rows $ddrows -ColW @(6000, 3000)))
        }
        if ($umd -and @($umd.Auth).Count) {
            [void]$sb.Append((Par -Text 'Authentication methods' -Bold))
            $amrows = @(, @('Method', 'Users'))
            foreach ($a in $umd.Auth) { $amrows += , @("$($a.Name)", (Num $a.Count)) }
            [void]$sb.Append((Tbl -Header -Rows $amrows -ColW @(6000, 3000)))
        }
    }
    else { [void]$sb.Append((RedMissing (T 'S5_1'))) }

    [void]$sb.Append((H2 (T 'S5_2')))
    [void]$sb.Append((Par (T 'Exp_Groups')))
    if ($kpiCards['Groups']) { [void]$sb.Append((KpiCards 'Groups')) }
    else { [void]$sb.Append((RedMissing (T 'S5_2'))) }
    [void]$sb.Append((Par (T 'Grp_types_p')))
    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi 08 Groups overview, check also svg Groups'))
    [void]$sb.Append((H2 (T 'Grp_size_h')))
    [void]$sb.Append((Par (T 'Grp_size_p')))
    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi 08 Groups overview, check also svg Groups'))
    [void]$sb.Append((H2 (T 'Grp_dyn_h')))
    [void]$sb.Append((Par (T 'Grp_dyn_p')))
    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi 08 Groups overview, check also svg Groups'))
    [void]$sb.Append((H2 (T 'Grp_owner_h')))
    [void]$sb.Append((Par (T 'Grp_owner_p')))
    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi 08 Groups overview, check also svg Groups'))
    [void]$sb.Append((H2 (T 'Grp_deptown_h')))
    [void]$sb.Append((Par (T 'Grp_deptown_p')))
    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi 08 Groups overview, check also svg Groups'))

    [void]$sb.Append((H2 (T 'S5_3')))
    [void]$sb.Append((Par (T 'Exp_RBAC')))
    [void]$sb.Append((KpiCards 'Rbac'))
    if ($rbac.Count) { [void]$sb.Append((Par ((T 'S5_3_body') -f $rbacActive, $rbacEligible))) }
    else { [void]$sb.Append((RedMissing (T 'S5_3'))) }
    [void]$sb.Append((Par (T 'Role_p1')))
    if ($rbacAssign.Count) {
        $rrows = @(, @('Role', 'Assignments', 'Share'))
        foreach ($g in (@($rbacAssign | Group-Object RoleDefinitionName | Sort-Object Count -Descending))) { $rrows += , @("$($g.Name)", (Num $g.Count), (Pct ([double]$g.Count / $rbacAssign.Count * 100))) }
        [void]$sb.Append((Tbl -Header -Rows $rrows -ColW @(5000, 2000, 2000)))
    }

    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi 12 RBAC overview, check also svg for RBAC'))
    [void]$sb.Append((H2 (T 'S5_4')))
    [void]$sb.Append((Par (T 'Exp_PIM')))
    if ($rbacEligible -gt 0) { [void]$sb.Append((Par ((T 'S5_4_inuse') -f $rbacEligible))) }
    else { [void]$sb.Append((GreyPar 'S5_4_none')) }

    [void]$sb.Append((H2 (T 'S5_5')))
    [void]$sb.Append((Par (T 'Exp_Guests')))
    $guestList = @($users | Where-Object { "$($_.UserType)" -eq 'Guest' })
    if ($guestList.Count) { [void]$sb.Append((Par ((T 'S5_5_some') -f $guestList.Count))) }
    else { [void]$sb.Append((Par (T 'S5_5_none'))) }

    [void]$sb.Append((H2 (T 'S6_1')))
    [void]$sb.Append((Par (T 'Exp_CA')))
    [void]$sb.Append((KpiCards 'ConditionalAccess'))
    if ($caPol.Count) {
        [void]$sb.Append((Par ((T 'S6_1_body') -f $caPol.Count, $caEnabled, $caReport, $caDisabled)))
        $crows = @(, @((T 'L_Policy'), (T 'L_State')))
        foreach ($p in ($caPol | Sort-Object DisplayName)) { $crows += , @("$($p.DisplayName)", "$($p.State)") }
        [void]$sb.Append((Tbl -Header -Rows $crows -ColW @(6500, 2500)))
    }
    else { [void]$sb.Append((RedMissing (T 'S6_1'))) }

    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi 09 Conditional access, check also svg for Conditional'))
    [void]$sb.Append((H2 (T 'S6_3')))
    [void]$sb.Append((Par (T 'Exp_MFA')))
    [void]$sb.Append((GreyPar 'S6_3_body'))
    [void]$sb.Append((Note 'Kopiera SVG Entra ID - Users (types, domains, auth, titles), Authentication methods'))
    [void]$sb.Append((Note 'Add data from EntraID Authentication methods | Policies'))
    [void]$sb.Append((Tbl -Header -Rows @(@('Name', 'Target', 'Enabled', 'Method'), @('', '', '', ''), @('', '', '', ''), @('', '', '', '')) -ColW @(2600, 2200, 2200, 2000)))

    [void]$sb.Append((H2 (T 'S6_4')))
    [void]$sb.Append((Par (T 'Exp_SSPR')))
    if ($sspr) {
        $enabledState = "$($sspr.EnabledState)$($sspr.Enablement)$($sspr.State)"
        if ($enabledState) { [void]$sb.Append((Par ((T 'S6_4_status') -f $enabledState))) }
        else { [void]$sb.Append((GreyPar 'S6_4_present')) }
    }
    else { [void]$sb.Append((RedMissing (T 'S6_4'))) }

    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi fliken for SSPR, check also svg for SSPR'))
    [void]$sb.Append((H2 (T 'L_AppRegs')))
    [void]$sb.Append((Par (T 'Exp_AppReg')))
    [void]$sb.Append((KpiCards 'AppRegistrations'))
    if ($apps.Count) {
        $expired = @($apps | Where-Object { $_.CredentialExpired -eq $true -or "$($_.CredentialExpired)" -eq 'True' }).Count
        [void]$sb.Append((Par ((T 'S7_1_body') -f $apps.Count, $appSingle, $appMulti, $expired)))
    }
    else { [void]$sb.Append((RedMissing (T 'L_AppRegs'))) }
    [void]$sb.Append((Par (T 'App_kpiRisk')))

    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi fliken for this category, check also svg for the category'))
    [void]$sb.Append((H2 (T 'S7_2')))
    [void]$sb.Append((Par (T 'Exp_Enterprise')))
    [void]$sb.Append((Par (T 'Ent_p1')))
    # same basis as the Excel KPI sheet "11 Enterprise apps": registrations that have a service principal
    function IsYes($v) { $v -eq $true -or "$v" -match '^(Yes|True)$' }
    $spApps = @($apps | Where-Object { IsYes $_.ServicePrincipalExists })
    if ($spApps.Count) {
        $spTotal  = $spApps.Count
        $spGraph  = @($spApps | Where-Object { IsYes $_.UsesGraphPermissions }).Count
        $spPublic = @($spApps | Where-Object { IsYes $_.PublicClient }).Count
        $grpApps = @{}
        foreach ($g in $groups) { foreach ($an in @("$($g.Applications)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) { $grpApps[$an] = $true } }
        $spViaGrp = @($spApps | Where-Object { $grpApps.ContainsKey("$($_.AppName)") }).Count
        [void]$sb.Append((KpiCards 'EnterpriseApps'))
        $kpiTbl = {
            param($title, $hdr, $pairs)
            [void]$sb.Append((Par ''))
            [void]$sb.Append((Par -Text $title -Bold))
            $rows = @(, @($hdr, 'Apps', 'Share'))
            foreach ($p in $pairs) { $rows += , @("$($p[0])", (Num $p[1]), (Pct ([double]$p[1] / $spTotal * 100))) }
            [void]$sb.Append((Tbl -Header -Rows $rows -ColW @(4600, 2200, 2200)))
        }
        $ssoPairs = @($spApps | Group-Object { if ("$($_.SSOType)") { "$($_.SSOType)" } else { 'None' } } | Sort-Object Count -Descending | ForEach-Object { , @($_.Name, $_.Count) })
        & $kpiTbl 'Authentication protocol (SSO type)' 'SSO type' $ssoPairs
        $bands = @($spApps | ForEach-Object { try { Get-SignInRecencyBand -LastSignIn $_.LastSignInDateTime } catch { 'Never / no data' } })
        $recPairs = @(foreach ($b in 'Last 7 days', '8-30 days', '31-90 days', '90+ days', 'Never / no data') { $c = @($bands | Where-Object { $_ -eq $b }).Count; if ($c) { , @($b, $c) } })
        & $kpiTbl 'Last sign-in' 'Last sign-in' $recPairs
        & $kpiTbl 'Microsoft Graph permissions' 'Graph' @(@('Uses Graph API', $spGraph), @('No Graph API', ($spTotal - $spGraph)))
        & $kpiTbl 'Client type' 'Client' @(@('Public client', $spPublic), @('Confidential client', ($spTotal - $spPublic)))
        & $kpiTbl 'App assignment through groups' 'Assignment' @(@('Assigned via group', $spViaGrp), @('No group assignment', ($spTotal - $spViaGrp)))
    }
    [void]$sb.Append((Note 'Add dashboard and data from ExcelKpi fliken for this category, check also svg for the category'))

    # ============ 9 Beroenden till andra system ============
    [void]$sb.Append((H1 (T 'CH9')))
    [void]$sb.Append((H2 (T 'CH9_1'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH9_2'))); [void]$sb.Append((Manual))

    # ============ 10 Licenser och avtal ============
    [void]$sb.Append((H1 (T 'CH10')))
    [void]$sb.Append((H2 (T 'S8_1')))
    if ($licenses.Count) {
        $lrows = @(, @((T 'L_SKU'), (T 'L_LicEnabled'), (T 'L_LicConsumed'), (T 'L_LicFree')))
        foreach ($l in ($licenses | Sort-Object @{ E = { [int]$_.consumedUnits } } -Descending)) {
            $lrows += , @("$($l.skuPartNumber)", ('{0:N0}' -f [int]$l.enabledUnits), ('{0:N0}' -f [int]$l.consumedUnits), ('{0:N0}' -f [int]$l.freeUnits))
        }
        [void]$sb.Append((Tbl -Header -Rows $lrows -ColW @(3600, 1800, 1800, 1800)))
    }
    else { [void]$sb.Append((RedMissing (T 'S8_1'))) }

    [void]$sb.Append((H2 (T 'CH_LicShare')))
    if ($licenses.Count) {
        $licSorted = @($licenses | Sort-Object @{ E = { [int]$_.consumedUnits } } -Descending)
        $totCons = [int]((@($licenses | ForEach-Object { [int]$_.consumedUnits }) | Measure-Object -Sum).Sum); if ($totCons -le 0) { $totCons = 1 }
        if ($licSorted.Count -ge 3) {
            $a = $licSorted[0]; $b = $licSorted[1]; $c = $licSorted[2]
            [void]$sb.Append((Par ((T 'CH_LicShare_intro') -f `
                        "$($a.skuPartNumber)", (Num $a.consumedUnits), (Pct ([double][int]$a.consumedUnits / $totCons * 100)), `
                        "$($b.skuPartNumber)", (Num $b.consumedUnits), (Pct ([double][int]$b.consumedUnits / $totCons * 100)), `
                        "$($c.skuPartNumber)", (Num $c.consumedUnits), (Pct ([double][int]$c.consumedUnits / $totCons * 100)))))
        }
        $srows = @(, @('Licence', 'Users', 'Share'))
        foreach ($l in $licSorted) {
            $srows += , @("$($l.skuPartNumber)", (Num $l.consumedUnits), (Pct ([double][int]$l.consumedUnits / $totCons * 100)))
        }
        [void]$sb.Append((Tbl -Header -Rows $srows -ColW @(4600, 2200, 2200)))
    }
    else { [void]$sb.Append((RedMissing (T 'CH_LicShare'))) }

    [void]$sb.Append((H2 (T 'Lic_util_h')))
    if ($licenses.Count) {
        [void]$sb.Append((Par (T 'Lic_util_p1')))
        $urows = @(, @('Sku', 'Category', 'Enabled', 'Consumed', 'Free', 'Util %', 'Status'))
        foreach ($l in ($licenses | Sort-Object @{ E = { [int]$_.consumedUnits } } -Descending)) {
            $en = [int]$l.enabledUnits; $co = [int]$l.consumedUnits
            $util = if ($en -gt 0) { Pct ([double]$co / $en * 100) } else { Pct 0 }
            $cat = try { Get-LicenseSkuCategory -SkuPartNumber "$($l.skuPartNumber)" } catch { '' }
            $urows += , @("$($l.skuPartNumber)", "$cat", (Num $l.enabledUnits), (Num $l.consumedUnits), (Num $l.freeUnits), $util, "$($l.capabilityStatus)")
        }
        [void]$sb.Append((Tbl -Header -Rows $urows -ColW @(2500, 1300, 1150, 1300, 1150, 1050, 1300)))
        [void]$sb.Append((Par (T 'Lic_util_p2')))
    }
    else { [void]$sb.Append((RedMissing (T 'Lic_util_h'))) }

    [void]$sb.Append((H2 (T 'Lic_assigned_h')))
    if ($licenses.Count) {
        [void]$sb.Append((Par (T 'Lic_assigned_p')))
        $arows = @(, @('Sku', 'Assigned', 'Unused'))
        foreach ($l in ($licenses | Sort-Object @{ E = { [int]$_.freeUnits } } -Descending)) { $arows += , @("$($l.skuPartNumber)", (Num $l.consumedUnits), (Num $l.freeUnits)) }
        [void]$sb.Append((Tbl -Header -Rows $arows -ColW @(4600, 2200, 2200)))
    }
    else { [void]$sb.Append((RedMissing (T 'Lic_assigned_h'))) }

    [void]$sb.Append((H2 (T 'Lic_top10_h')))
    if ($licenses.Count) {
        [void]$sb.Append((Par (T 'Lic_top10_p')))
        $trows2 = @(, @('Sku', 'Assigned units'))
        foreach ($l in (@($licenses | Sort-Object @{ E = { [int]$_.consumedUnits } } -Descending | Select-Object -First 10))) { $trows2 += , @("$($l.skuPartNumber)", (Num $l.consumedUnits)) }
        [void]$sb.Append((Tbl -Header -Rows $trows2 -ColW @(6000, 3000)))
    }
    else { [void]$sb.Append((RedMissing (T 'Lic_top10_h'))) }

    [void]$sb.Append((H2 (T 'Lic_dept_h')))
    [void]$sb.Append((Par (T 'Lic_dept_p')))
    [void]$sb.Append((H2 (T 'Lic_office_h')))
    [void]$sb.Append((Par (T 'Lic_office_p')))
    [void]$sb.Append((H2 (T 'Lic_peruser_h')))
    [void]$sb.Append((Par (T 'Lic_peruser_p')))

    # ============ 11 Certifikat ============
    [void]$sb.Append((H1 (T 'CH11')))
    $fedDomains = @($domains | Where-Object { "$($_.Type)" -match 'Federat' })
    if ($fedDomains.Count) { $fedList = (($fedDomains | ForEach-Object { $_.Domain }) -join ', '); [void]$sb.Append((Par ((T 'S11_fed') -f $fedList))) }
    else { [void]$sb.Append((GreyPar 'S11_none')) }
    [void]$sb.Append((H2 (T 'CH11_1'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH11_2'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH11_3'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH11_4'))); [void]$sb.Append((Manual))

    # ============ 12 Användarkonton, systemkonton och lösenord ============
    [void]$sb.Append((H1 (T 'CH12')))
    [void]$sb.Append((H2 (T 'CH12_1'))); [void]$sb.Append((Manual))

    # ============ 13 Behörighetsgrupper (RBAC) ============
    [void]$sb.Append((H1 (T 'CH13')))
    if ($rbac.Count) { [void]$sb.Append((Par ((T 'S5_3_body') -f $rbacActive, $rbacEligible))) }
    else { [void]$sb.Append((RedMissing (T 'CH13'))) }

    # ============ 14 Säkerhet ============
    [void]$sb.Append((H1 (T 'CH14')))
    [void]$sb.Append((H2 (T 'CH14_1'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH14_2'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH14_3'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH14_4'))); [void]$sb.Append((Manual))

    # ============ 15 Övervakning ============
    [void]$sb.Append((H1 (T 'CH15')))
    [void]$sb.Append((GreyPar 'S9_body'))

    # ============ 16 Logghantering ============
    [void]$sb.Append((H1 (T 'CH16')))
    [void]$sb.Append((H2 (T 'CH16_1'))); [void]$sb.Append((Manual))

    # ============ 17 Backup och återläsning ============
    [void]$sb.Append((H1 (T 'CH17')))
    [void]$sb.Append((H2 (T 'CH17_1'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH17_2'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH17_3'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH17_4'))); [void]$sb.Append((Manual))
    [void]$sb.Append((H2 (T 'CH17_5'))); [void]$sb.Append((Manual))

    # ============ 18 Överlämning och godkännande ============
    [void]$sb.Append((H1 (T 'CH18')))
    [void]$sb.Append((Manual))

    $sectPr = "<w:sectPr><w:footerReference w:type='default' r:id='rId3'/><w:pgSz w:w='11906' w:h='16838'/><w:pgMar w:top='1417' w:right='1417' w:bottom='1417' w:left='1417' w:header='708' w:footer='708' w:gutter='0'/></w:sectPr>"
    $bodyXml = $sb.ToString() + $sectPr

    # ---- fixed parts ----
    $documentXml = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n" +
    "<w:document xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><w:body>$bodyXml</w:body></w:document>"

    # footer with the metadata table + page number
    $pageField = "<w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText xml:space='preserve'> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r><w:r><w:t xml:space='preserve'> (</w:t></w:r><w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText xml:space='preserve'> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r><w:r><w:t>)</w:t></w:r></w:p>"
    function FtrCell($w, $text, $bold) {
        $rpr = if ($bold) { '<w:rPr><w:b/><w:sz w:val=''16''/></w:rPr>' } else { '<w:rPr><w:sz w:val=''16''/></w:rPr>' }
        "<w:tc><w:tcPr><w:tcW w:w='$w' w:type='dxa'/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r>$rpr<w:t xml:space='preserve'>$(XmlEsc "$text")</w:t></w:r></w:p></w:tc>"
    }
    $fw = @(2400, 2400, 1500, 1000, 1700)
    $ftrHead = '<w:tr>' + (FtrCell $fw[0] 'Prepared by' $true) + (FtrCell $fw[1] 'Document ID' $true) + (FtrCell $fw[2] 'Date' $true) + (FtrCell $fw[3] 'Version' $true) + (FtrCell $fw[4] 'Page' $true) + '</w:tr>'
    $sidaCell = "<w:tc><w:tcPr><w:tcW w:w='$($fw[4])' w:type='dxa'/></w:tcPr>$pageField</w:tc>"
    $ftrVal = '<w:tr>' + (FtrCell $fw[0] $author $false) + (FtrCell $fw[1] $docId $false) + (FtrCell $fw[2] $today $false) + (FtrCell $fw[3] $version $false) + $sidaCell + '</w:tr>'
    $ftrGrid = ($fw | ForEach-Object { "<w:gridCol w:w='$_'/>" }) -join ''
    $restW = $fw[1] + $fw[2] + $fw[3] + $fw[4]
    $ftrClass = '<w:tr>' + (FtrCell $fw[0] 'Informationsklass' $true) + "<w:tc><w:tcPr><w:tcW w:w='$restW' w:type='dxa'/><w:gridSpan w:val='4'/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr></w:p></w:tc>" + '</w:tr>'
    $footerTbl = "<w:tbl><w:tblPr><w:tblW w:w='$(($fw|Measure-Object -Sum).Sum)' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='BFBFBF'/><w:bottom w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideH w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideV w:val='single' w:sz='4' w:color='BFBFBF'/></w:tblBorders></w:tblPr><w:tblGrid>$ftrGrid</w:tblGrid>$ftrHead$ftrVal$ftrClass</w:tbl>"
    $footerXml = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<w:ftr xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`">$footerTbl</w:ftr>"

    $contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>
'@
    $relsRoot = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
'@
    $imgRels = ''
    foreach ($im in $images) { $imgRels += "<Relationship Id=`"$($im.RelId)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/$($im.MediaName)`"/>" }
    $relsDoc = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/><Relationship Id=`"rId2`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings`" Target=`"settings.xml`"/><Relationship Id=`"rId3`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer`" Target=`"footer1.xml`"/><Relationship Id=`"rId4`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering`" Target=`"numbering.xml`"/>$imgRels</Relationships>"
    $settingsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:updateFields w:val="true"/></w:settings>
'@
    $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="sv-SE"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr><w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="6AA84F"/></w:pBdr></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="2"/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="9"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style></w:styles>
'@

    $numberingXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:suff w:val="space"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:suff w:val="space"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/><w:suff w:val="space"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>
'@

    # ---- assemble + zip ----
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("wr_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp '_rels') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp 'word') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp 'word\_rels') -Force | Out-Null

    Set-Content -LiteralPath (Join-Path $tmp '[Content_Types].xml') -Value $contentTypes -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp '_rels\.rels') -Value $relsRoot -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\document.xml') -Value $documentXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\styles.xml') -Value $stylesXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\settings.xml') -Value $settingsXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\numbering.xml') -Value $numberingXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\footer1.xml') -Value $footerXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\_rels\document.xml.rels') -Value $relsDoc -Encoding UTF8

    if ($images.Count) {
        New-Item -ItemType Directory -Path (Join-Path $tmp 'word\media') -Force | Out-Null
        foreach ($im in $images) { Copy-Item -LiteralPath $im.Path -Destination (Join-Path $tmp "word\media\$($im.MediaName)") -Force }
    }

    $wordDir = Join-Path $src 'Report\word'
    if (-not (Test-Path $wordDir)) { New-Item -ItemType Directory -Path $wordDir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $baseName = (($docId -replace '[\\/:*?"<>|]', ' ') -replace '\s+', ' ').Trim()
    if (-not $baseName) { $baseName = T 'DocIdDefault' }
    $out = Join-Path $wordDir "$baseName $stamp.docx"
    if (Test-Path $out) { Remove-Item -LiteralPath $out -Force }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $out)
    Remove-Item -LiteralPath $tmp -Recurse -Force

    Write-Host "  Word report: $out" -ForegroundColor DarkGreen
    Write-Host ""
    if ($lang -eq 'sv') {
        Write-Host "  Nu är Word-filen klar, men allt kopieras inte över från Excel och SVG rakt in i Word." -ForegroundColor Yellow
        Write-Host "  Ta en titt på Word-filen och de andra filerna, och kopiera och klistra in diagram och bilder i systemdokumentationen." -ForegroundColor Yellow
    }
    else {
        Write-Host "  The Word file is ready, but not everything is copied over from Excel and SVG straight into Word." -ForegroundColor Yellow
        Write-Host "  Review the Word file and the other files, and copy and paste charts and images into the system documentation." -ForegroundColor Yellow
    }
    return $out
}
