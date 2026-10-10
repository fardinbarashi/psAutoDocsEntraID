# Autodoc v1.0
function New-VisioDocument {
    <#
        Writes a .vsdx from a simple shape/connector model.

        A .vsdx is an OPC package - a ZIP with the same layout as .xlsx - so it
        can be produced without Visio installed. The parts written here are the
        minimum Visio needs: content types, package and document relationships,
        the document, one page, and the page contents holding the shapes.

        Shapes carry inline geometry rather than referencing masters. That keeps
        the package self-contained: no stencil has to resolve when the file is
        opened, at the cost of the shapes being plain geometry instead of smart
        Visio shapes.

        Coordinates are in inches with the origin bottom-left, which is Visio's
        own system; the callers work in the same units.

        Icons ride along as separate Foreign shapes laid over their node, each
        pointing at a PNG in visio/media. Visio's ForeignData records carry
        raster data only, so SVG icons are skipped here - they still appear in
        the SVG map, which is drawn from the same model.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Shape,          # @{ Id; Text; X; Y; W; H; Kind; Fill; Line; FontSize }
        $Connector = @(),                      # @{ From; To }
        [double]$PageWidth  = 11,
        [double]$PageHeight = 8.5,
        [string]$PageName   = 'Page-1',
        [string]$Title      = 'Autodoc'
    )

    function Esc([string]$t) { [System.Security.SecurityElement]::Escape($t) }

    # ---------- geometry for one shape ----------
    function ShapeXml($s, [int]$id) {
        $kind = if ($s.Kind) { $s.Kind } else { 'Rectangle' }
        $fill = if ($s.Fill) { $s.Fill } else { '#FFFFFF' }
        $line = if ($s.Line) { $s.Line } else { '#444444' }
        $font = if ($s.FontSize) { $s.FontSize } else { 9 }
        $noFill = if ($fill -eq 'none') { 1 } else { 0 }
        $noLine = if ($line -eq 'none') { 1 } else { 0 }
        if ($fill -eq 'none') { $fill = '#FFFFFF' }
        if ($line -eq 'none') { $line = '#FFFFFF' }

        # Simplest reliable Visio text model: plain text with literal newlines
        # PLAIN-TEXT model - the only one that renders correctly in Visio.
        # cp/pp runs collapse multi-line text onto one row (confirmed twice), so
        # formatting is per-SHAPE, not per-line: a block is bold as a whole, and
        # aligned as a whole. Lines that need a different weight are drawn as
        # separate shapes by the caller (see the split tenant/legend blocks).
        if ($s.Lines) {
            $textLines = @($s.Lines | ForEach-Object {
                $t = if ($_.BoldPrefix) { "$($_.BoldPrefix)$($_.Text)" } else { "$($_.Text)" }
                @{ Text = $t; Bold = [bool]$_.Bold; Middle = ($_.Align -eq 'middle') }
            })
        } else {
            $middle = ($s.Kind -eq 'Ellipse') -or $s.CenterText
            $textLines = @(("$($s.Text)" -split "`n") | ForEach-Object { @{ Text = $_; Bold = $false; Middle = $middle } })
        }
        $vertAlign   = 0                                   # top (middle collapses lines)
        $linePattern = if ($s.Dashed) { 2 } else { 1 }
        $topMargin   = if ($s.TopInset) { $s.TopInset } else { 0.08 }

        # Shape-level bold: bold when every line is bold; and a whole shape can be
        # forced bold via -Bold on the shape.
        $allBold   = $s.Bold -or ((@($textLines | Where-Object { -not $_.Bold }).Count -eq 0) -and ($textLines.Count -gt 0))
        $allMiddle = (@($textLines | Where-Object { -not $_.Middle }).Count -eq 0) -and ($textLines.Count -gt 0)
        $horzAlign = if ($allMiddle) { 1 } else { 0 }
        # Bold = bit 1 in Visio's Style bitmask (17 was wrong and was ignored).
        # Set both the shape-level Char.Style cell and a Character section row.
        if ($allBold) {
            $charStyle   = "<Cell N='Char.Style' V='1'/>"
            $charSection = "<Section N='Character'><Row IX='0'><Cell N='Style' V='1'/><Cell N='Color' V='#000000'/><Cell N='Size' V='$($font/72)'/></Row></Section>"
        } else {
            $charStyle   = ''
            $charSection = ''
        }
        $paraSection = "<Section N='Paragraph'><Row IX='0'><Cell N='HorzAlign' V='$horzAlign'/></Row></Section>"

        # Plain body: lines joined by literal newlines, no runs.
        $richText = (@($textLines | ForEach-Object { Esc $_.Text }) -join "`n")

        if ($kind -eq 'Ellipse') {
            # An ellipse is a single EllipticalArcTo pair around the bounding box
            $geom = @"
        <Section N='Geometry' IX='0'>
          <Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/>
          <Row T='MoveTo' IX='1'><Cell N='X' V='0'/><Cell N='Y' V='$($s.H/2)'/></Row>
          <Row T='EllipticalArcTo' IX='2'><Cell N='X' V='$($s.W/2)'/><Cell N='Y' V='$($s.H)'/><Cell N='A' V='$($s.W)'/><Cell N='B' V='$($s.H/2)'/><Cell N='C' V='0'/><Cell N='D' V='1'/></Row>
          <Row T='EllipticalArcTo' IX='3'><Cell N='X' V='$($s.W/2)'/><Cell N='Y' V='0'/><Cell N='A' V='0'/><Cell N='B' V='$($s.H/2)'/><Cell N='C' V='0'/><Cell N='D' V='1'/></Row>
        </Section>
"@
        }
        else {
            $geom = @"
        <Section N='Geometry' IX='0'>
          <Cell N='NoFill' V='$noFill'/><Cell N='NoLine' V='$noLine'/>
          <Row T='RelMoveTo' IX='1'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row>
          <Row T='RelLineTo' IX='2'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row>
          <Row T='RelLineTo' IX='3'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row>
          <Row T='RelLineTo' IX='4'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row>
          <Row T='RelLineTo' IX='5'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row>
        </Section>
"@
        }

        @"
      <Shape ID='$id' NameU='Shape$id' Type='Shape' LineStyle='0' FillStyle='0' TextStyle='0'>
        <Cell N='PinX' V='$($s.X)'/><Cell N='PinY' V='$($s.Y)'/>
        <Cell N='Width' V='$($s.W)'/><Cell N='Height' V='$($s.H)'/>
        <Cell N='LocPinX' V='$($s.W/2)' F='Width*0.5'/><Cell N='LocPinY' V='$($s.H/2)' F='Height*0.5'/>
        <Cell N='FillForegnd' V='$fill'/><Cell N='LineColor' V='$line'/><Cell N='LineWeight' V='0.01'/><Cell N='LinePattern' V='$linePattern'/>
        <Cell N='Char.Size' V='$($font/72)'/><Cell N='Char.Color' V='#000000'/>$charStyle<Cell N='VerticalAlign' V='$vertAlign'/>
        <Cell N='LeftMargin' V='0.06'/><Cell N='RightMargin' V='0.06'/><Cell N='TopMargin' V='$topMargin'/><Cell N='BottomMargin' V='0.03'/>
        $charSection
        $paraSection
$geom
        <Text>$richText</Text>
      </Shape>
"@
    }

    # ---------- a connector is a two-point line shape ----------
    function ConnectorXml($from, $to, [int]$id) {
        $x1 = $from.X; $y1 = $from.Y - $from.H / 2
        $x2 = $to.X;   $y2 = $to.Y + $to.H / 2
        $w = [math]::Max([math]::Abs($x2 - $x1), 0.001)
        $h = [math]::Max([math]::Abs($y2 - $y1), 0.001)
        $px = [math]::Min($x1, $x2); $py = [math]::Min($y1, $y2)
        $rx1 = ($x1 - $px) / $w; $ry1 = ($y1 - $py) / $h
        $rx2 = ($x2 - $px) / $w; $ry2 = ($y2 - $py) / $h
        @"
      <Shape ID='$id' NameU='Conn$id' Type='Shape' LineStyle='0' FillStyle='0' TextStyle='0'>
        <Cell N='PinX' V='$($px + $w/2)'/><Cell N='PinY' V='$($py + $h/2)'/>
        <Cell N='Width' V='$w'/><Cell N='Height' V='$h'/>
        <Cell N='LocPinX' V='$($w/2)' F='Width*0.5'/><Cell N='LocPinY' V='$($h/2)' F='Height*0.5'/>
        <Cell N='LineColor' V='#888888'/><Cell N='LineWeight' V='0.01'/><Cell N='EndArrow' V='4'/>
        <Section N='Geometry' IX='0'>
          <Cell N='NoFill' V='1'/><Cell N='NoLine' V='0'/>
          <Row T='RelMoveTo' IX='1'><Cell N='X' V='$rx1'/><Cell N='Y' V='$ry1'/></Row>
          <Row T='RelLineTo' IX='2'><Cell N='X' V='$rx2'/><Cell N='Y' V='$ry2'/></Row>
        </Section>
      </Shape>
"@
    }

    # ---------- an icon becomes a Foreign shape over its node ----------
    # ---------- build the page ----------
    $byId = @{}; foreach ($s in $Shape) { $byId[$s.Id] = $s }
    $sb = [System.Text.StringBuilder]::new()
    $media = @()      # @{ Name; Bytes; RelId }
    $n = 0
    foreach ($s in $Shape) {
        $n++; [void]$sb.AppendLine((ShapeXml $s $n))
        # Icons are intentionally NOT embedded in the .vsdx. Visio renders a
        # ForeignData image as a crossed-out placeholder here, so the drawing is
        # kept clean with colour and text only. The icons still appear in the
        # .svg map, which is generated from the same model by Export-MapAsSvg.
    }
    foreach ($c in $Connector) {
        if (-not $byId.ContainsKey($c.From) -or -not $byId.ContainsKey($c.To)) { continue }
        $n++; [void]$sb.AppendLine((ConnectorXml $byId[$c.From] $byId[$c.To] $n))
    }

    $pageRels = ($media | ForEach-Object {
        "  <Relationship Id='$($_.RelId)' Type='http://schemas.microsoft.com/visio/2010/relationships/image' Target='../media/$($_.Name)'/>"
    }) -join "`n"
    $pngDefault = if ($media) { "  <Default Extension='png' ContentType='image/png'/>" } else { '' }

    $V = "http://schemas.microsoft.com/office/visio/2012/main"
    $R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    $parts = @{
        '[Content_Types].xml' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
  <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
  <Default Extension='xml' ContentType='application/xml'/>
$pngDefault
  <Override PartName='/docProps/app.xml' ContentType='application/vnd.openxmlformats-officedocument.extended-properties+xml'/>
  <Override PartName='/docProps/core.xml' ContentType='application/vnd.openxmlformats-package.core-properties+xml'/>
  <Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/>
  <Override PartName='/visio/pages/pages.xml' ContentType='application/vnd.ms-visio.pages+xml'/>
  <Override PartName='/visio/pages/page1.xml' ContentType='application/vnd.ms-visio.page+xml'/>
  <Override PartName='/visio/windows.xml' ContentType='application/vnd.ms-visio.windows+xml'/>
</Types>
"@
        '_rels/.rels' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/>
  <Relationship Id='rId2' Type='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties' Target='docProps/core.xml'/>
  <Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties' Target='docProps/app.xml'/>
</Relationships>
"@
        'docProps/core.xml' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
  xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dcterms='http://purl.org/dc/terms/'
  xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
  <dc:title>$(Esc $Title)</dc:title><dc:creator>Autodoc</dc:creator>
  <cp:lastModifiedBy>Autodoc</cp:lastModifiedBy>
  <dcterms:created xsi:type='dcterms:W3CDTF'>$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')</dcterms:created>
</cp:coreProperties>
"@
        'docProps/app.xml' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
  xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'>
  <Application>Microsoft Visio</Application><Company>Autodoc</Company>
</Properties>
"@
        'visio/document.xml' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<VisioDocument xmlns='$V' xmlns:r='$R' xml:space='preserve'>
  <DocumentSettings TopPage='0' DefaultTextStyle='0' DefaultLineStyle='0' DefaultFillStyle='0' DefaultGuideStyle='0'>
    <GlueSettings>9</GlueSettings><SnapSettings>65847</SnapSettings>
  </DocumentSettings>
  <FaceNames>
    <FaceName NameU='Segoe UI' UnicodeRanges='0 0 0 0' CharSets='0 0' Panos='2 11 5 2 4 2 4 2 2 3' Flags='325'/>
  </FaceNames>
</VisioDocument>
"@
        'visio/_rels/document.xml.rels' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/>
  <Relationship Id='rId2' Type='http://schemas.microsoft.com/visio/2010/relationships/windows' Target='windows.xml'/>
</Relationships>
"@
        'visio/windows.xml' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Windows xmlns='$V' xmlns:r='$R' ClientWidth='1000' ClientHeight='700'>
  <Window ID='0' WindowType='Drawing' WindowState='1073741824' Document='visio/document.xml' Page='0'
          ViewScale='1' ViewCenterX='$($PageWidth/2)' ViewCenterY='$($PageHeight/2)'/>
</Windows>
"@
        'visio/pages/pages.xml' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Pages xmlns='$V' xmlns:r='$R' xml:space='preserve'>
  <Page ID='0' NameU='$(Esc $PageName)' Name='$(Esc $PageName)' ViewScale='-1' ViewCenterX='$($PageWidth/2)' ViewCenterY='$($PageHeight/2)'>
    <PageSheet LineStyle='0' FillStyle='0' TextStyle='0'>
      <Cell N='PageWidth' V='$PageWidth'/><Cell N='PageHeight' V='$PageHeight'/>
      <Cell N='PageScale' V='1'/><Cell N='DrawingScale' V='1'/>
    </PageSheet>
    <Rel r:id='rId1'/>
  </Page>
</Pages>
"@
        'visio/pages/_rels/pages.xml.rels' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page1.xml'/>
</Relationships>
"@
        'visio/pages/_rels/page1.xml.rels' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
$pageRels
</Relationships>
"@
        'visio/pages/page1.xml' = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<PageContents xmlns='$V' xmlns:r='$R' xml:space='preserve'>
  <Shapes>
$($sb.ToString())  </Shapes>
</PageContents>
"@
    }

    # ---------- write the package ----------
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("vsdx_" + [guid]::NewGuid().ToString('N'))
    New-Item -Path $tmp -ItemType Directory -Force | Out-Null
    foreach ($part in $parts.Keys) {
        $target = Join-Path $tmp $part
        New-Item -Path (Split-Path $target -Parent) -ItemType Directory -Force | Out-Null
        [IO.File]::WriteAllText($target, $parts[$part], (New-Object Text.UTF8Encoding $false))
    }
    if ($media) {
        $mediaDir = Join-Path $tmp 'visio/media'
        New-Item -Path $mediaDir -ItemType Directory -Force | Out-Null
        foreach ($m in $media) { [IO.File]::WriteAllBytes((Join-Path $mediaDir $m.Name), $m.Bytes) }
    }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::CreateFromDirectory($tmp, $Path)
    Remove-Item $tmp -Recurse -Force
    $Path
}
