# Empty-corner check ("fill it with a rectangle"): draw a translucent rectangle from the corner where the design
# starts to the corner where it ends, and see which parts of that rectangle hold nothing. A slide that feels bare
# usually has one such hole; put a small element there (an icon, a 3-D object, a number) instead of rearranging
# everything. This script does the same on a copy of a slide: it finds the bounds of the content, splits them into a
# grid, marks the empty cells, and with -Fill places a picture or .glb model in the largest hole.
# The original deck is not changed; the result is saved to -Out (one slide: the check, or with -Fill the repaired slide
# followed by the check).
# Example: .\balance_check.ps1 -Deck C:\work\report.pptx -Slide 3 -Out C:\work\check.pptx -Preview C:\work\check.png
param(
    [Parameter(Mandatory)][string]$Deck,
    [int]$Slide = 1,
    [int[]]$Grid = @(3, 3),                       # columns, rows inside the content bounds
    [double]$Threshold = 0.06,                    # a cell counts as empty when less than this share of it is covered
    [string]$Fill = '',                           # picture or .glb to put into the largest empty area
    [string]$Color = '#2F6BFF',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in @($Deck) + $Fill) { if ($p -and -not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    $script:app = New-Object -ComObject PowerPoint.Application
    $script:pres = $app.Presentations.Open((Resolve-Path $Deck).Path, -1, -1, 0)      # read-only, untitled copy, no window
    $script:W = [double]$pres.PageSetup.SlideWidth; $script:H = [double]$pres.PageSetup.SlideHeight
    for ($i = $pres.Slides.Count; $i -ge 1; $i--) { if ($i -ne $Slide) { $pres.Slides($i).Delete() } }
    $script:sl = $pres.Slides(1)

    # content = every shape except backgrounds (anything covering most of the slide) and things off the slide
    $boxes = @()
    foreach ($s in $sl.Shapes) {
        $l = [math]::Max(0, [double]$s.Left); $t = [math]::Max(0, [double]$s.Top)
        $r = [math]::Min($W, [double]$s.Left + [double]$s.Width); $b = [math]::Min($H, [double]$s.Top + [double]$s.Height)
        if ($r -le $l -or $b -le $t) { continue }
        if (($r - $l) * ($b - $t) -gt 0.6 * $W * $H) { continue }
        $boxes += , @($l, $t, $r, $b)
    }
    if (-not $boxes.Count) { throw 'no content shapes on that slide' }
    $x0 = ($boxes | ForEach-Object { $_[0] } | Measure-Object -Minimum).Minimum; $y0 = ($boxes | ForEach-Object { $_[1] } | Measure-Object -Minimum).Minimum
    $x1 = ($boxes | ForEach-Object { $_[2] } | Measure-Object -Maximum).Maximum; $y1 = ($boxes | ForEach-Object { $_[3] } | Measure-Object -Maximum).Maximum
    $cols = $Grid[0]; $rows = $Grid[1]; $cw = ($x1 - $x0) / $cols; $ch = ($y1 - $y0) / $rows; $empty = @()
    for ($r = 0; $r -lt $rows; $r++) {
        for ($c = 0; $c -lt $cols; $c++) {
            $cl = $x0 + $cw * $c; $ct = $y0 + $ch * $r; $cover = 0
            # ponytail: overlapping shapes are counted twice, which only matters near the threshold; exact union area if a slide ever misreads
            foreach ($bx in $boxes) { $ow = [math]::Min($cl + $cw, $bx[2]) - [math]::Max($cl, $bx[0]); $oh = [math]::Min($ct + $ch, $bx[3]) - [math]::Max($ct, $bx[1]); if ($ow -gt 0 -and $oh -gt 0) { $cover += $ow * $oh } }
            if ($cover / ($cw * $ch) -lt $Threshold) { $empty += , @($c, $r, $cl, $ct) }
        }
    }
    "content bounds: {0:N0},{1:N0} - {2:N0},{3:N0} of {4:N0} x {5:N0}; empty cells: {6} of {7}" -f $x0, $y0, $x1, $y1, $W, $H, $empty.Count, ($cols * $rows)
    foreach ($e in $empty) { "  empty: column $($e[0] + 1), row $($e[1] + 1)" }

    if ($Fill -and $empty.Count) {                                    # repaired copy first, the check after it
        [void]$sl.Duplicate(); $fix = $pres.Slides(1); $script:sl = $pres.Slides(2)
        $e = $empty[0]; $size = [math]::Min($cw, $ch) * 0.8; $fl = $e[2] + ($cw - $size) / 2; $ft = $e[3] + ($ch - $size) / 2
        if ($Fill -like '*.glb') { $obj = $fix.Shapes.Add3DModel((Resolve-Path $Fill).Path, 0, -1, [single]$fl, [single]$ft, [single]$size, [single]$size) }
        else { $obj = $fix.Shapes.AddPicture((Resolve-Path $Fill).Path, 0, -1, 0, 0, -1, -1); $obj.LockAspectRatio = -1; $obj.Height = [single]$size; $obj.Left = [single]($e[2] + ($cw - [double]$obj.Width) / 2); $obj.Top = [single]$ft }
        $obj.Name = 'Filler'
    }
    $veil = New-Rect ([single]$x0) ([single]$y0) ([single]($x1 - $x0)) ([single]($y1 - $y0)); $veil.Fill.ForeColor.RGB = (ConvertTo-Bgr $Color); $veil.Fill.Transparency = 0.7; $veil.Line.Visible = 0; $veil.Name = 'CheckBounds'
    foreach ($e in $empty) {
        $mark = New-Rect ([single]$e[2]) ([single]$e[3]) ([single]$cw) ([single]$ch); $mark.Fill.Visible = 0; $mark.Line.ForeColor.RGB = 0x0000FF; $mark.Line.Weight = 2.5; $mark.Line.DashStyle = 4
        $tag = New-Text 'EMPTY' 14 0x0000FF 'Arial Black'; $tag.Left = [single]($e[2] + ($cw - [double]$tag.Width) / 2); $tag.Top = [single]($e[3] + ($ch - [double]$tag.Height) / 2)
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
