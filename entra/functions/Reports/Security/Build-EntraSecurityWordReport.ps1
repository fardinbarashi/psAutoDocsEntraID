# Autodoc v1.0
function Build-EntraSecurityWordReport {
    <#
        "Security and improvements" Word document (.docx), written as raw Office
        Open XML like the system documentation, so it needs neither Word nor an
        extra module. Same cover, footer and styles as the system document.

        Chapters: introduction, summary (with the map), priority actions, all
        improvement proposals per area, the workload identity hunt, next steps.
        Input is the object from Get-SecurityImprovementData; texts come from
        Entra\Config\SecurityImprovements.json and the cover/footer metadata
        from Entra\Config\WordReportProfile.json.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path,
        [string]$SourceFolder,
        [string]$MapPng,
        [string]$ConsultantLogo,
        [string]$CustomerLogo
    )

    $cfg  = $Data.Config
    $lang = $Data.Language
    function TX($k) { $e = $cfg.Text[$k]; if ($e -and "$($e[$lang])") { "$($e[$lang])" } elseif ($e) { "$($e['en'])" } else { "[$k]" } }
    function SevName($s) { TX "Sev_$s" }
    $sevFill = @{ CRITICAL = 'F8D7DA'; HIGH = 'F3E5F5'; MEDIUM = 'FFF3CD'; INFO = 'E8F0FE' }

    # ---- profile (metadata that isn't in Graph), shared with the system document ----
    $entraRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot))   # Functions\Reports\Security -> Entra
    $prof = @{}
    $profilePath = Join-Path $entraRoot 'Config\WordReportProfile.json'
    if (Test-Path $profilePath) { try { $prof = Read-AutodocJson -Path $profilePath -AsHashtable } catch { $prof = @{} } }
    $tenant     = $Data.Tenant
    $consultant = if ("$($prof.ConsultantCompany)") { "$($prof.ConsultantCompany)" } else { '[Consultant]' }
    $author     = if ("$($prof.Author)") { "$($prof.Author)" } else { '[Author]' }
    $customer   = if ("$($prof.CustomerName)") { "$($prof.CustomerName)" } elseif ($tenant) { "$($tenant.TenantName)" } else { '[Customer]' }
    $docId      = TX 'FileName'
    $version    = if ("$($prof.Version)") { "$($prof.Version)" } else { '1.0' }
    $today      = Get-Date -Format 'yyyy-MM-dd'
    $logoImgDir = Join-Path $entraRoot 'files\cache\wordTemplates\template\img'
    function CachedLogo($base) { if (Test-Path $logoImgDir) { $f = Get-ChildItem -Path $logoImgDir -Filter "$base.*" -ErrorAction SilentlyContinue | Select-Object -First 1; if ($f) { return $f.FullName } } ; return '' }
    $consultantLogo = if ($ConsultantLogo -and (Test-Path $ConsultantLogo)) { $ConsultantLogo } elseif (CachedLogo 'companylogo') { CachedLogo 'companylogo' } elseif ("$($prof.ConsultantLogo)" -and (Test-Path "$($prof.ConsultantLogo)")) { "$($prof.ConsultantLogo)" } else { '' }
    $customerLogo   = if ($CustomerLogo -and (Test-Path $CustomerLogo)) { $CustomerLogo } elseif (CachedLogo 'clientlogo') { CachedLogo 'clientlogo' } elseif ("$($prof.CustomerLogo)" -and (Test-Path "$($prof.CustomerLogo)")) { "$($prof.CustomerLogo)" } else { '' }
    function Num($n) { ([int]$n).ToString('N0', [Globalization.CultureInfo]::InvariantCulture) -replace ',', ' ' }

    # ---- OOXML building blocks ----
    function XmlEsc([string]$s) { if ($null -eq $s) { return '' }; $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' }
    function Par {
        param([string]$Text = '', [string]$Style = '', [string]$Align = '', [switch]$Bold, [int]$Size = 0, [string]$Color = '')
        $ppr = ''; if ($Style) { $ppr += "<w:pStyle w:val='$Style'/>" }; if ($Align) { $ppr += "<w:jc w:val='$Align'/>" }
        $rpr = ''; if ($Bold) { $rpr += '<w:b/>' }; if ($Size) { $rpr += "<w:sz w:val='$Size'/><w:szCs w:val='$Size'/>" }; if ($Color) { $rpr += "<w:color w:val='$Color'/>" }
        $ppr = if ($ppr) { "<w:pPr>$ppr</w:pPr>" } else { '' }; $rpr = if ($rpr) { "<w:rPr>$rpr</w:rPr>" } else { '' }
        "<w:p>$ppr<w:r>$rpr<w:t xml:space='preserve'>$(XmlEsc $Text)</w:t></w:r></w:p>"
    }
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


    function H3($t) { Par -Text $t -Style 'Heading3' }
    # table whose first column carries the severity colour (SevCol = column index, -1 = none)
    function SevTbl {
        param([object[]]$Rows, [int[]]$ColW, [string[]]$SevKeys, [int]$SevCol = 0, [switch]$NoHeader)
        $total = ($ColW | Measure-Object -Sum).Sum
        $grid = ($ColW | ForEach-Object { "<w:gridCol w:w='$_'/>" }) -join ''
        $trs = ''
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            $tcs = ''
            for ($c = 0; $c -lt $Rows[$r].Count; $c++) {
                $hdr = ($r -eq 0 -and -not $NoHeader)
                $shd = if ($hdr) { "<w:shd w:val='clear' w:color='auto' w:fill='4472C4'/>" }
                       elseif ($c -eq $SevCol -and $SevKeys[$r] -and $sevFill[$SevKeys[$r]]) { "<w:shd w:val='clear' w:color='auto' w:fill='$($sevFill[$SevKeys[$r]])'/>" }
                       elseif ($r % 2 -eq 0) { "<w:shd w:val='clear' w:color='auto' w:fill='F2F2F2'/>" } else { '' }
                $rpr = if ($hdr) { "<w:rPr><w:b/><w:color w:val='FFFFFF'/></w:rPr>" } elseif (($c -eq $SevCol -and $SevKeys[$r]) -or ($NoHeader -and $c -eq 0)) { '<w:rPr><w:b/></w:rPr>' } else { '' }
                $tcs += "<w:tc><w:tcPr><w:tcW w:w='$($ColW[$c])' w:type='dxa'/>$shd</w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r>$rpr<w:t xml:space='preserve'>$(XmlEsc "$($Rows[$r][$c])")</w:t></w:r></w:p></w:tc>"
            }
            $trs += "<w:tr>$tcs</w:tr>"
        }
        "<w:tbl><w:tblPr><w:tblW w:w='$total' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='BFBFBF'/><w:left w:val='single' w:sz='4' w:color='BFBFBF'/><w:bottom w:val='single' w:sz='4' w:color='BFBFBF'/><w:right w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideH w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideV w:val='single' w:sz='4' w:color='BFBFBF'/></w:tblBorders></w:tblPr><w:tblGrid>$grid</w:tblGrid>$trs</w:tbl>"
    }
    function Short($s, [int]$max) { $s = "$s"; if ($s.Length -gt $max) { $s.Substring(0, $max - 1) + '…' } else { $s } }

    $impr = @($Data.Improvements)
    $hunt = $Data.Hunt
    $sb = [System.Text.StringBuilder]::new()

    # ---- cover: label | value | consultant logo, then title, customer, customer logo ----
    $cLogo = LogoDrawing $consultantLogo 1.9
    $logoRun = if ($cLogo) { "<w:r>$cLogo</w:r>" } else { "<w:r><w:t xml:space='preserve'>$(XmlEsc "[$consultant logo]")</w:t></w:r>" }
    $cw = @(1900, 2600, 3000)
    $cvTc = { param($w, $text, $bold, $fill) "<w:tc><w:tcPr><w:tcW w:w='$w' w:type='dxa'/>$(if($fill){"<w:shd w:val='clear' w:color='auto' w:fill='$fill'/>"})<w:vAlign w:val='center'/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r>$(if($bold){'<w:rPr><w:b/></w:rPr>'})<w:t xml:space='preserve'>$(XmlEsc $text)</w:t></w:r></w:p></w:tc>" }
    $logoTop  = "<w:tc><w:tcPr><w:tcW w:w='$($cw[2])' w:type='dxa'/><w:vMerge w:val='restart'/><w:vAlign w:val='center'/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/><w:jc w:val='center'/></w:pPr>$logoRun</w:p></w:tc>"
    $logoCont = "<w:tc><w:tcPr><w:tcW w:w='$($cw[2])' w:type='dxa'/><w:vMerge/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr></w:p></w:tc>"
    $coverRows = "<w:tr>$(& $cvTc $cw[0] 'Document Type' $true '')$(& $cvTc $cw[1] (TX 'DocType') $false '')$logoTop</w:tr>" +
                 "<w:tr>$(& $cvTc $cw[0] 'Area' $true 'F2F2F2')$(& $cvTc $cw[1] 'Entra ID' $false 'F2F2F2')$logoCont</w:tr>"
    $coverGrid = ($cw | ForEach-Object { "<w:gridCol w:w='$_'/>" }) -join ''
    [void]$sb.Append("<w:tbl><w:tblPr><w:tblW w:w='$(($cw|Measure-Object -Sum).Sum)' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='BFBFBF'/><w:left w:val='single' w:sz='4' w:color='BFBFBF'/><w:bottom w:val='single' w:sz='4' w:color='BFBFBF'/><w:right w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideH w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideV w:val='single' w:sz='4' w:color='BFBFBF'/></w:tblBorders></w:tblPr><w:tblGrid>$coverGrid</w:tblGrid>$coverRows</w:tbl>")
    1..5 | ForEach-Object { [void]$sb.Append((Par '')) }
    [void]$sb.Append((Par -Text $docId -Style 'Title' -Align 'center'))
    [void]$sb.Append((Par -Text $customer -Align 'center' -Size 36))
    1..7 | ForEach-Object { [void]$sb.Append((Par '')) }
    $kLogo = LogoDrawing $customerLogo 2.6
    if ($kLogo) { [void]$sb.Append("<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r>$kLogo</w:r></w:p>") }
    [void]$sb.Append((PageBreak))

    # ---- table of contents ----
    [void]$sb.Append((Par -Text $(if ($lang -eq 'sv') { 'Innehåll' } else { 'Contents' }) -Style 'TOCHeading'))
    [void]$sb.Append("<w:p><w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText xml:space='preserve'> TOC \o &quot;1-3&quot; \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType='separate'/></w:r><w:r><w:t xml:space='preserve'>$(XmlEsc $(if ($lang -eq 'sv') { 'Högerklicka och välj Uppdatera fält.' } else { 'Right-click and choose Update Field.' }))</w:t></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r></w:p>")
    [void]$sb.Append((PageBreak))

    # ---- 1 introduction ----
    [void]$sb.Append((H1 (TX 'Intro_H')))
    [void]$sb.Append((Par ((TX 'Intro_p1') -f $customer)))
    [void]$sb.Append((Par (TX 'Intro_p2')))
    [void]$sb.Append((Note ((TX 'Intro_p3') -f $(if ($SourceFolder) { Split-Path $SourceFolder -Leaf } else { '' }), $today)))

    # ---- 2 summary ----
    [void]$sb.Append((H1 (TX 'Summary_H')))
    [void]$sb.Append((Par (TX 'Summary_p')))
    $sumRows = @(, @((TX 'Severity'), (TX 'Proposal'), (TX 'HuntFindings')))
    $sumKeys = @('')
    foreach ($s in $Data.SevOrder) {
        $sumRows += , @((SevName $s), (Num @($impr | Where-Object Severity -eq $s).Count), $(if ($hunt.Ran) { Num @($hunt.Findings | Where-Object Severity -eq $s).Count } else { '-' }))
        $sumKeys += $s
    }
    $sumRows += , @((TX 'Total'), (Num $impr.Count), $(if ($hunt.Ran) { Num @($hunt.Findings).Count } else { '-' }))
    $sumKeys += ''
    [void]$sb.Append((SevTbl -Rows $sumRows -ColW @(3000, 3000, 3000) -SevKeys $sumKeys))
    [void]$sb.Append((Par ''))
    [void]$sb.Append((Par (TX 'Summary_Area')))
    $areaRows = @(, (@((TX 'Area')) + @($Data.SevOrder | ForEach-Object { SevName $_ }) + @((TX 'Total'))))
    foreach ($area in $Data.AreaOrder) {
        $items = @($impr | Where-Object Area -eq $area)
        $name = @($cfg.Areas[$area][$lang], $area) | Where-Object { $_ } | Select-Object -First 1
        $areaRows += , (@($name) + @($Data.SevOrder | ForEach-Object { $s = $_; "$(@($items | Where-Object Severity -eq $s).Count)" }) + @("$($items.Count)"))
    }
    [void]$sb.Append((SevTbl -Rows $areaRows -ColW @(2900, 1200, 1200, 1200, 1200, 1300) -SevKeys @() -SevCol -1))
    if ($MapPng -and (Test-Path $MapPng)) {
        $mapRun = LogoDrawing $MapPng 6.3
        if ($mapRun) { [void]$sb.Append((Par '')); [void]$sb.Append("<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r>$mapRun</w:r></w:p>") }
    }

    # ---- 3 priority actions ----
    [void]$sb.Append((H1 (TX 'Prio_H')))
    $prio = @($impr | Where-Object { $_.Severity -in 'CRITICAL', 'HIGH' })
    if ($prio.Count) {
        [void]$sb.Append((Par (TX 'Prio_p')))
        $rows = @(, @((TX 'Severity'), (TX 'Area'), (TX 'Proposal'), (TX 'Count'))); $keys = @('')
        foreach ($p in $prio) { $rows += , @((SevName $p.Severity), $p.AreaName, $p.Title, $(if ($p.Count) { Num $p.Count } else { '-' })); $keys += $p.Severity }
        [void]$sb.Append((SevTbl -Rows $rows -ColW @(1300, 2500, 4200, 1000) -SevKeys $keys))
    } else { [void]$sb.Append((Par (TX 'Prio_None'))) }

    # ---- 4 all proposals per area ----
    [void]$sb.Append((H1 (TX 'Impr_H')))
    [void]$sb.Append((Par (TX 'Impr_p')))
    foreach ($area in $Data.AreaOrder) {
        $items = @($impr | Where-Object Area -eq $area)
        if (-not $items.Count) { continue }
        [void]$sb.Append((H2 $items[0].AreaName))
        foreach ($p in $items) {
            [void]$sb.Append((H3 $p.Title))
            $rows = @(@((TX 'Severity'), (SevName $p.Severity)),
                      @((TX 'Finding'), $p.Finding),
                      @((TX 'Recommendation'), $p.Recommendation),
                      @((TX 'Reference'), $p.Reference))
            # two-column card: label | text; the severity row is coloured
            [void]$sb.Append((SevTbl -Rows $rows -ColW @(2200, 6800) -SevKeys @($p.Severity, '', '', '') -SevCol 1 -NoHeader))
            [void]$sb.Append((Par ''))
        }
    }

    # ---- 5 workload identity hunt ----
    [void]$sb.Append((H1 (TX 'Hunt_H')))
    [void]$sb.Append((Par (TX 'Hunt_p')))
    if (-not $hunt.Ran) {
        [void]$sb.Append((Note (TX 'Hunt_NotRun')))
    } elseif (-not @($hunt.Findings).Count) {
        [void]$sb.Append((Par (TX 'Hunt_NoFindings')))
    } else {
        $sevRank = @{ CRITICAL = 0; HIGH = 1; MEDIUM = 2; INFO = 3 }
        [void]$sb.Append((H2 (TX 'Hunt_Sum_H')))
        $rows = @(, @((TX 'Severity'), 'Category', (TX 'Count'))); $keys = @('')
        foreach ($g in @($hunt.Findings | Group-Object Severity, Category | Sort-Object @{ Expression = { $sevRank["$($_.Group[0].Severity)"] } }, @{ Expression = 'Count'; Descending = $true })) {
            $rows += , @((SevName $g.Group[0].Severity), $g.Group[0].Category, (Num $g.Count)); $keys += "$($g.Group[0].Severity)"
        }
        [void]$sb.Append((SevTbl -Rows $rows -ColW @(1500, 6300, 1200) -SevKeys $keys))

        $top = @($hunt.Findings | Where-Object { $_.Severity -in 'CRITICAL', 'HIGH' })
        if ($top.Count) {
            [void]$sb.Append((H2 (TX 'Hunt_Top_H')))
            $show = @($top | Select-Object -First 40)
            $rows = @(, @((TX 'Severity'), 'Category', 'App / user', 'IP', 'Time')); $keys = @('')
            foreach ($f in $show) { $rows += , @((SevName $f.Severity), (Short $f.Category 48), (Short $f.AppName 40), (Short $f.IP 30), "$($f.Time)"); $keys += "$($f.Severity)" }
            [void]$sb.Append((SevTbl -Rows $rows -ColW @(1100, 2900, 2200, 1300, 1500) -SevKeys $keys))
            if ($top.Count -gt $show.Count) { [void]$sb.Append((Note ((TX 'Hunt_Top_More') -f $show.Count, $top.Count))) }
        }

        if (@($hunt.Triage).Count) {
            [void]$sb.Append((H2 (TX 'Hunt_Triage_H')))
            [void]$sb.Append((Par (TX 'Hunt_Triage_p')))
            $rows = @(, @('User', 'Verdict', 'IoC sign-ins', 'IoC IPs')); $keys = @('')
            foreach ($t in @($hunt.Triage | Sort-Object @{ Expression = { $sevRank["$($_.Severity)"] } })) {
                $rows += , @("$($t.User)", "$($t.Verdict)", "$($t.IocSignIns)", (Short $t.IocIPs 40)); $keys += "$($t.Severity)"
            }
            [void]$sb.Append((SevTbl -Rows $rows -ColW @(2900, 2600, 1300, 2200) -SevKeys $keys -SevCol 1))
        }

        if (@($hunt.Credentials).Count) {
            [void]$sb.Append((H2 (TX 'Hunt_Cred_H')))
            [void]$sb.Append((Par (TX 'Hunt_Cred_p')))
            $cr = @($hunt.Credentials)
            $isTrue = { param($v) "$v" -match '^(True|Yes)$' }
            $activeSecrets = @($cr | Where-Object { $_.CredentialType -eq 'ClientSecret' -and -not (& $isTrue $_.Expired) })
            $rows = @(@('Credential', (TX 'Count')),
                @('Client secrets (active)', (Num $activeSecrets.Count)),
                @('Client secrets without owner', (Num @($activeSecrets | Where-Object { $_.Object -eq 'Application' -and -not "$($_.Owners)" }).Count)),
                @('Client secrets on service principal objects', (Num @($activeSecrets | Where-Object { $_.Object -eq 'ServicePrincipal' }).Count)),
                @('Certificates (active)', (Num @($cr | Where-Object { $_.CredentialType -eq 'Certificate' -and -not (& $isTrue $_.Expired) }).Count)),
                @('Federated credentials', (Num @($cr | Where-Object { $_.CredentialType -eq 'FederatedCredential' }).Count)),
                @('Expired', (Num @($cr | Where-Object { & $isTrue $_.Expired }).Count)),
                @('Added recently', (Num @($cr | Where-Object { & $isTrue $_.AddedRecently }).Count)))
            [void]$sb.Append((SevTbl -Rows $rows -ColW @(6000, 3000) -SevKeys @() -SevCol -1))
        }
        [void]$sb.Append((Par ''))
        [void]$sb.Append((Note (TX 'Hunt_Html')))
    }

    # ---- 6 next steps ----
    [void]$sb.Append((H1 (TX 'Next_H')))
    [void]$sb.Append((Par (TX 'Next_p1')))
    [void]$sb.Append((Par (TX 'Next_p2')))
    [void]$sb.Append((Note (TX 'Next_p3')))

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


    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    if (Test-Path $Path) { Remove-Item -LiteralPath $Path -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $Path)
    Remove-Item -LiteralPath $tmp -Recurse -Force
    return $Path
}
