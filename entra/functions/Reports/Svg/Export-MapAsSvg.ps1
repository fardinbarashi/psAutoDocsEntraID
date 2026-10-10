# Autodoc v1.0
function Export-MapAsSvg {
    <#
        Renders a shape/connector model as a self-contained SVG picture that
        can be viewed in any browser.

        Shape coordinates use a bottom-left origin while SVG's origin is
        top-left, so Y is flipped here. Units are converted from inches at
        96 px per inch.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Shape,
        $Connector = @(),
        [double]$PageWidth = 11,
        [double]$PageHeight = 8.5,
        [int]$Dpi = 96
    )

    function Esc([string]$t) { [System.Security.SecurityElement]::Escape($t) }
    function PX([double]$inch) { [math]::Round($inch * $Dpi, 1) }
    function FlipY([double]$y) { [math]::Round(($PageHeight - $y) * $Dpi, 1) }

    $w = PX $PageWidth; $h = PX $PageHeight
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("<svg xmlns='http://www.w3.org/2000/svg' width='$w' height='$h' viewBox='0 0 $w $h' font-family='Segoe UI, sans-serif'>")
    [void]$sb.AppendLine("  <rect width='100%' height='100%' fill='#FFFFFF'/>")
    [void]$sb.AppendLine("  <defs><marker id='arrow' markerWidth='8' markerHeight='8' refX='6' refY='3' orient='auto'><path d='M0,0 L6,3 L0,6 z' fill='#888888'/></marker></defs>")

    $byId = @{}; foreach ($s in $Shape) { $byId[$s.Id] = $s }
    foreach ($c in $Connector) {
        if (-not $byId.ContainsKey($c.From) -or -not $byId.ContainsKey($c.To)) { continue }
        $f = $byId[$c.From]; $t = $byId[$c.To]
        # Route to the facing edges so the arrowhead always points into the box.
        # Boxes roughly level with the hub (sides of the ring) are reached
        # horizontally - a vertical route there leaves a zero-length final leg
        # and the arrow never orients; boxes above/below are reached vertically.
        $dy = $t.Y - $f.Y
        if ([math]::Abs($dy) -lt (($f.H + $t.H) / 2)) {
            $y1 = FlipY $f.Y; $y2 = FlipY $t.Y
            if ($t.X -ge $f.X) { $x1 = PX ($f.X + $f.W / 2); $x2 = PX ($t.X - $t.W / 2) }
            else               { $x1 = PX ($f.X - $f.W / 2); $x2 = PX ($t.X + $t.W / 2) }
            $midX = [math]::Round(($x1 + $x2) / 2, 1)
            [void]$sb.AppendLine("  <path d='M $x1 $y1 H $midX V $y2 H $x2' fill='none' stroke='#888888' stroke-width='1.2' marker-end='url(#arrow)'/>")
        }
        else {
            $x1 = PX $f.X; $x2 = PX $t.X
            if ($t.Y -lt $f.Y) {
                $y1 = FlipY ($f.Y - $f.H / 2); $y2 = FlipY ($t.Y + $t.H / 2)
            }
            else {
                $y1 = FlipY ($f.Y + $f.H / 2); $y2 = FlipY ($t.Y - $t.H / 2)
            }
            $mid = [math]::Round(($y1 + $y2) / 2, 1)
            [void]$sb.AppendLine("  <path d='M $x1 $y1 V $mid H $x2 V $y2' fill='none' stroke='#888888' stroke-width='1.2' marker-end='url(#arrow)'/>")
        }
    }

    foreach ($s in $Shape) {
        $sw = PX $s.W; $sh = PX $s.H
        $cx = PX $s.X; $cy = FlipY $s.Y
        $x  = [math]::Round($cx - $sw / 2, 1); $y = [math]::Round($cy - $sh / 2, 1)
        $fill = if ($s.Fill) { $s.Fill } else { '#FFFFFF' }
        $line = if ($s.Line) { $s.Line } else { '#444444' }
        $drawShape = ($fill -ne 'none' -or $line -ne 'none')
        $fs   = if ($s.FontSize) { $s.FontSize } else { 9 }
        $dash = if ($s.Dashed) { " stroke-dasharray='6,4'" } else { '' }
        $strokeW = if ($s.StrokeWidth) { $s.StrokeWidth } else { '1.2' }
        if ($drawShape) {
            if ($s.Kind -eq 'Ellipse') {
                [void]$sb.AppendLine("  <ellipse cx='$cx' cy='$cy' rx='$([math]::Round($sw/2,1))' ry='$([math]::Round($sh/2,1))' fill='$fill' stroke='$line' stroke-width='$strokeW'$dash/>")
            }
            else {
                [void]$sb.AppendLine("  <rect x='$x' y='$y' width='$sw' height='$sh' rx='4' fill='$fill' stroke='$line' stroke-width='$strokeW'$dash/>")
            }
        }
        # The icon sits on the left; the text is centred in the space that
        # remains to the right of it, with padding kept on both sides.
        $pad = 6
        $textLeft = $x + $pad
        if ($s.Icon) {
            $iconSide = [math]::Round([math]::Min($sh * 0.6, $sw * 0.34), 1)
            $ix = [math]::Round($x + $sw * 0.05, 1)
            $iy = [math]::Round($cy - $iconSide / 2, 1)
            [void]$sb.AppendLine("  <image x='$ix' y='$iy' width='$iconSide' height='$iconSide' href='data:$($s.Icon.Mime);base64,$($s.Icon.Base64)'/>")
            $textLeft = $ix + $iconSide + $pad
        }
        $textRight = $x + $sw - $pad
        $colour = '#000000'

        if ($s.Lines) {
            # Rich model: each line carries its own bold flag and alignment.
            # Used by the tenant box, bands, and (top-anchored) org cards.
            $lines = @($s.Lines)
            $lineStep = if ($s.LinesTop) { 0.185 * 96 } else { $fs + 3 }
            if ($s.LinesTop) {
                $inset = if ($s.TopInset) { PX $s.TopInset } else { 0 }
                $startY = ($y + $pad + $fs + $inset)
            } else {
                $startY = $cy - (($lines.Count - 1) * $lineStep) / 2 + $fs / 3
            }
            # Optional grid: lines from GridFrom onward are laid out in Columns
            # columns (column-major) instead of one per row. Header lines above
            # GridFrom stay full width. Only shapes that set Columns use this.
            $gcols = if ($s.Columns -and [int]$s.Columns -gt 1) { [int]$s.Columns } else { 1 }
            $gfrom = if ($s.GridFrom) { [int]$s.GridFrom } else { 0 }
            $gN    = $lines.Count - $gfrom
            $gRows = if ($gcols -gt 1 -and $gN -gt 0) { [math]::Ceiling($gN / $gcols) } else { 0 }
            $gColW = if ($gcols -gt 1) { ($textRight - $textLeft) / $gcols } else { 0 }
            for ($i = 0; $i -lt $lines.Count; $i++) {
                $ln = $lines[$i]
                if ($gcols -gt 1 -and $i -ge $gfrom) {
                    $g = $i - $gfrom
                    $col = [math]::Floor($g / $gRows)
                    $row = $g % $gRows
                    $tx = [math]::Round($textLeft + $col * $gColW, 1)
                    $ly = [math]::Round($startY + ($gfrom + $row) * $lineStep, 1)
                    $anchor = 'start'
                }
                else {
                    $ly = [math]::Round($startY + $i * $lineStep, 1)
                    $anchor = if ($ln.Align -eq 'middle') { 'middle' } else { 'start' }
                    $tx = if ($ln.Align -eq 'middle') { [math]::Round(($textLeft + $textRight) / 2, 1) } else { [math]::Round($textLeft, 1) }
                }
                if ($ln.BoldPrefix) {
                    # Legend line: the prefix (e.g. "Red:") is bold, the rest normal.
                    # Rendered as two tspans in one left-aligned text element.
                    $pfx = Esc $ln.BoldPrefix
                    $rest = Esc $ln.Text
                    [void]$sb.AppendLine("  <text x='$tx' y='$ly' font-size='$fs' fill='$colour' text-anchor='start'><tspan font-weight='bold'>$pfx</tspan>$rest</text>")
                } else {
                    $weight = if ($ln.Bold) { " font-weight='bold'" } else { '' }
                    [void]$sb.AppendLine("  <text x='$tx' y='$ly' font-size='$fs' fill='$colour' text-anchor='$anchor'$weight>$(Esc $ln.Text)</text>")
                }
            }
        }
        else {
            if ($s.Rotate) {
                # Rotated single-line label (e.g. matrix column headers): drawn
                # centred in the cell and rotated around its centre so long names
                # read vertically. Only shapes that set Rotate use this path.
                $ly = [math]::Round($cy + $fs / 3, 1)
                $bw = if ($s.Bold) { " font-weight='bold'" } else { '' }
                [void]$sb.AppendLine("  <g transform='rotate($($s.Rotate) $cx $cy)'><text x='$cx' y='$ly' font-size='$fs' fill='$colour' text-anchor='middle'$bw>$(Esc "$($s.Text)")</text></g>")
            }
            else {
            # Plain model: left-aligned next to the icon, wrapped to keep one
            # font size. About 0.6 * font size per character.
            # Centre when the shape is an ellipse or explicitly asks for it
            # (free/viral blocks set CenterText); otherwise left-align by the icon.
            $centre = (($s.Kind -eq 'Ellipse') -or $s.CenterText) -and (-not $s.Icon)
            if ($centre) {
                $avail = ($sw - $pad * 2) * 0.92; $textX = [math]::Round($cx, 1); $anchor = 'middle'
            } else {
                $avail = $textRight - $textLeft; $textX = [math]::Round($textLeft, 1); $anchor = 'start'
            }
            $maxChars = if ($avail -gt 0) { [math]::Max(6, [math]::Floor($avail / ($fs * 0.6))) } else { 999 }
            $wrapped = @()
            foreach ($line in @("$($s.Text)" -split "`n")) {
                if ($line.Length -le $maxChars) { $wrapped += $line; continue }
                $words = $line -split '(?<=[_\s/-])'
                $cur = ''
                foreach ($w in $words) {
                    while ($w.Length -gt $maxChars) {
                        if ($cur) { $wrapped += $cur.TrimEnd(); $cur = '' }
                        $wrapped += $w.Substring(0, $maxChars)
                        $w = $w.Substring($maxChars)
                    }
                    if (($cur + $w).Length -gt $maxChars -and $cur) { $wrapped += $cur.TrimEnd(); $cur = $w }
                    else { $cur += $w }
                }
                if ($cur) { $wrapped += $cur.TrimEnd() }
            }
            $lineStep = $fs + 2
            if ($s.LinesTop) {
                $startY = $y + $pad + $fs
            } else {
                $startY = $cy - (($wrapped.Count - 1) * $lineStep) / 2 + $fs / 3
            }
            $bw = if ($s.Bold) { " font-weight='bold'" } else { '' }
            for ($i = 0; $i -lt $wrapped.Count; $i++) {
                $ly = [math]::Round($startY + $i * $lineStep, 1)
                [void]$sb.AppendLine("  <text x='$textX' y='$ly' font-size='$fs' fill='$colour' text-anchor='$anchor'$bw>$(Esc $wrapped[$i])</text>")
            }
            }
        }
    }
    [void]$sb.AppendLine('</svg>')
    [IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object Text.UTF8Encoding $false))
    $Path
}
