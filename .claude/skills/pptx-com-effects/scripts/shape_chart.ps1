# Charts drawn with plain shapes instead of the chart object: quicker to style, free to exaggerate, and every
# piece stays an ordinary shape. Column/Bar = rectangles, Pie = "partial circle" shapes, Donut = pie + a
# background-coloured circle on top, Ring = one percentage as a donut gauge.
# Example: .\shape_chart.ps1 -Type Column -Values 120,450,950,1030,700 -Labels '1st|2nd|3rd|4th|5th' -Title 'SALES' -Out C:\work\chart.pptx -Preview C:\work\chart.png
#          .\shape_chart.ps1 -Type Donut -Values 20,30,50 -Labels 'A|B|C' -Center 'SHARE' -Out C:\work\donut.pptx
#          .\shape_chart.ps1 -Type Ring -Values 72 -Center 'DONE' -Out C:\work\ring.pptx
param(
    [ValidateSet('Column', 'Bar', 'Pie', 'Donut', 'Ring')][string]$Type = 'Column',
    [Parameter(Mandatory)][double[]]$Values,      # Ring: one number 0..100
    [string]$Labels = '',                         # 'a|b|c', one per value
    [string]$Title = '',                          # '|' = line break
    [int]$Highlight = 0,                          # 1-based item drawn in the accent colour; 0 = the largest
    [string]$Unit = '',                           # appended to value labels, e.g. '%'
    [string]$Center = '',                         # Donut/Ring: text in the hole
    [string]$Color = '#3A3A3A',                   # ordinary items (pies step from this towards light grey)
    [string]$Accent = '#E8604C',
    [string]$TextColor = '#1B1B1B',
    [string]$Background = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Type -eq 'Ring' -and $Values.Count -ne 1) { throw '-Type Ring takes exactly one value (0..100)' }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Fill($shape, [int]$bgr) { $shape.Fill.ForeColor.RGB = $bgr; $shape.Line.Visible = 0 }
function Add-Label([string]$text, $size, [int]$bgr, [string]$font, $cx, $cy) {
    $t = New-Text $text $size $bgr $font
    $t.Left = [single]($cx - [double]$t.Width / 2); $t.Top = [single]($cy - [double]$t.Height / 2)
    $t
}
# Pie angles: degrees clockwise, 0 = 3 o'clock, kept in -180..180. $from/$sweep are measured from 12 o'clock.
function Add-Slice($cx, $cy, $r, [double]$from, [double]$sweep, [int]$bgr) {
    if ($sweep -ge 359.9) { $p = $script:sl.Shapes.AddShape(9, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r)) }
    else {
        $p = $script:sl.Shapes.AddShape(142, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r))   # msoShapePie
        $norm = { param($a) $x = ($a + 180) % 360; if ($x -lt 0) { $x += 360 }; $x - 180 }
        $p.Adjustments.Item(1) = [single](& $norm ($from - 90)); $p.Adjustments.Item(2) = [single](& $norm ($from + $sweep - 90))
    }
    Set-Fill $p $bgr
    $p
}
function Get-Grey([int]$i, [int]$n) {             # $Color for the first slice, stepping to light grey
    $h = $Color.TrimStart('#'); $t = if ($n -le 1) { 0 } else { $i / ($n - 1) * 0.75 }
    $rgb = 0, 2, 4 | ForEach-Object { $v = [Convert]::ToInt32($h.Substring($_, 2), 16); [int]($v + (225 - $v) * $t) }
    $rgb[0] + $rgb[1] * 256 + $rgb[2] * 65536
}

try {
    New-Deck $Background
    $names = if ($Labels) { $Labels -split '\|' } else { @() }
    $ink = ConvertTo-Bgr $TextColor; $acc = ConvertTo-Bgr $Accent; $base = ConvertTo-Bgr $Color; $bg = ConvertTo-Bgr $Background
    $n = $Values.Count
    $max = ($Values | Measure-Object -Maximum).Maximum
    $hi = if ($Highlight -gt 0) { $Highlight - 1 } else { [array]::IndexOf($Values, $max) }
    if ($Title) { $tt = New-FitText $Title 34 $ink $Font ($W * 0.5); $tt.Left = 56; $tt.Top = 44; $tt.Name = 'Title' }

    if ($Type -eq 'Column') {
        $baseY = 440; $maxH = 280; $slot = [math]::Min(110, 760 / $n); $bw = $slot * 0.55
        $x0 = ($W - $slot * $n) / 2
        for ($i = 0; $i -lt $n; $i++) {
            $h = [math]::Max(4, $maxH * $Values[$i] / $max); $cx = $x0 + $slot * ($i + 0.5)
            $b = New-Rect ([single]($cx - $bw / 2)) ([single]($baseY - $h)) ([single]$bw) ([single]$h); $b.Name = "Bar$($i + 1)"
            $c = if ($i -eq $hi) { $acc } else { $base }; Set-Fill $b $c
            $size = if ($i -eq $hi) { 22 } else { 14 }
            [void](Add-Label "$($Values[$i])$Unit" $size $c $Font $cx ($baseY - $h - $size * 0.9))
            if ($i -lt $names.Count) { [void](Add-Label $names[$i] 12 $ink 'Arial' $cx ($baseY + 16)) }
        }
    } elseif ($Type -eq 'Bar') {
        $x0 = 250; $maxW = 520; $slot = [math]::Min(70, 330 / $n); $bh = $slot * 0.55
        $y0 = 150 + (330 - $slot * $n) / 2
        for ($i = 0; $i -lt $n; $i++) {
            $w = [math]::Max(4, $maxW * $Values[$i] / $max); $cy = $y0 + $slot * ($i + 0.5)
            $b = New-Rect ([single]$x0) ([single]($cy - $bh / 2)) ([single]$w) ([single]$bh); $b.Name = "Bar$($i + 1)"
            $c = if ($i -eq $hi) { $acc } else { $base }; Set-Fill $b $c
            $v = New-Text "$($Values[$i])$Unit" 14 $c $Font; $v.Left = [single]($x0 + $w + 10); $v.Top = [single]($cy - [double]$v.Height / 2)
            if ($i -lt $names.Count) { $l = New-Text $names[$i] 13 $ink 'Arial'; $l.Left = [single]($x0 - 14 - [double]$l.Width); $l.Top = [single]($cy - [double]$l.Height / 2) }
        }
    } else {
        $cx = $W / 2; $cy = 290; $r = 170
        if ($Type -eq 'Ring') {
            $track = Add-Slice $cx $cy $r 0 360 (ConvertTo-Bgr '#E3E3E3'); $track.Name = 'Track'
            $p = Add-Slice $cx $cy $r 0 (3.6 * [math]::Min(100, $Values[0])) $acc; $p.Name = 'Value'
        } else {
            $sum = ($Values | Measure-Object -Sum).Sum; $from = 0.0
            for ($i = 0; $i -lt $n; $i++) {
                $sweep = 360 * $Values[$i] / $sum
                $c = if ($Highlight -gt 0 -and $i -eq $hi) { $acc } else { Get-Grey $i $n }
                $p = Add-Slice $cx $cy $r $from $sweep $c; $p.Name = "Slice$($i + 1)"
                $mid = ($from + $sweep / 2) * [math]::PI / 180
                $k = if ($Type -eq 'Donut') { 0.8 } else { 0.62 }
                $txt = "$([math]::Round(100 * $Values[$i] / $sum))%"
                if ($i -lt $names.Count -and $Type -eq 'Pie') { $txt = "$($names[$i])`r$txt" }
                $dark = ($i / [math]::Max(1, $n - 1)) -lt 0.5 -or ($Highlight -gt 0 -and $i -eq $hi)
                $lab = Add-Label $txt 15 $(if ($dark) { 0xFFFFFF } else { $ink }) $Font ($cx + $r * $k * [math]::Sin($mid)) ($cy - $r * $k * [math]::Cos($mid))
                $lab.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                $from += $sweep
            }
        }
        if ($Type -ne 'Pie') {
            $hr = $r * $(if ($Type -eq 'Ring') { 0.78 } else { 0.62 })
            $hole = $sl.Shapes.AddShape(9, [single]($cx - $hr), [single]($cy - $hr), [single](2 * $hr), [single](2 * $hr)); Set-Fill $hole $bg; $hole.Name = 'Hole'
            if ($Type -eq 'Ring') { [void](Add-Label "$($Values[0])%" 60 $acc $Font $cx ($cy - $(if ($Center) { 14 } else { 0 }))) }
            if ($Center) { [void](Add-Label $Center $(if ($Type -eq 'Ring') { 16 } else { 26 }) $ink $Font $cx ($cy + $(if ($Type -eq 'Ring') { 38 } else { 0 }))) }
        }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
