# Photos in shapes and arrangements other than a row of equal rectangles - the kind of layouts PowerPoint's
# Design Ideas pane proposes, built on purpose: each picture is cut to a shape (picture first, then the shape ->
# Merge Shapes > Intersect), so the aspect ratio is never distorted.
#   Columns  full-height strips side by side          Mosaic   two stacked + one tall
#   Circles  big circle in the middle, smaller ones overlapping behind it
#   Slant    leaning parallelograms                   Scatter  tilted prints with a white border and shadow
#   Brush    one photo with a torn, brush-stroke edge and room for text beside it
#   Blob     one photo in a soft organic shape with two faint larger echoes of itself around it
#   Diagonal three photos divided by slanted cuts, edge to edge
#   Frames   photos inset in white cards (one tall, two small) on a grey page
#   Arc      colour field on the left, photo with a round edge on the right, dashed arc along the curve
#   Grid     same-size cells at an even gap under a centred title (-Grid cols,rows; the images repeat to fill it)
#   Callouts one photo in the middle, four translucent caption bars reaching in from both sides (-Captions)
# Example: .\photo_layout.ps1 -Images a.jpg,b.jpg,c.jpg -Layout Circles -Title 'OUR TEAM' -Out C:\work\photos.pptx -Preview C:\work\photos.png
param(
    [Parameter(Mandatory)][string[]]$Images,
    [ValidateSet('Columns', 'Mosaic', 'Circles', 'Slant', 'Scatter', 'Brush', 'Blob', 'Diagonal', 'Frames', 'Arc', 'Grid', 'Callouts')][string]$Layout = 'Mosaic',
    [string]$Title = '',                          # '|' = line break
    [string]$Body = '',
    [string]$Color = '#1B1B1B',
    [string]$Background = '#F4F3EF',
    [string]$Accent = '#FDB515',                  # Arc: colour field
    [string]$Font = 'Arial Black',
    [int[]]$Grid = @(5, 3),                       # Grid: columns, rows
    [string]$Captions = '',                       # Callouts: 'head;text|head;text|...' (four; left-top, right-top, left-bottom, right-bottom)
    [int]$Seed = 7,                               # Brush / Scatter: change it for a different edge or tilt
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent
$script:made = 0

# Cut a picture to $frame (any shape): cover the frame's box without stretching, then intersect.
function Add-Cropped([string]$path, $frame, $fit = $null) {   # $fit: place the picture by this box instead (keeps copies aligned)
    $script:made++; $fb = Get-Box $frame; $frame.Name = "Frame$script:made"; if ($fit) { $fb = $fit }
    $pic = $sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = "Src$script:made"
    $pic.Width = [single]$fb.W; if ([double]$pic.Height -lt $fb.H) { $pic.Height = [single]$fb.H }
    $pic.Left = [single]($fb.L + ($fb.W - [double]$pic.Width) / 2); $pic.Top = [single]($fb.T + ($fb.H - [double]$pic.Height) / 2)
    $shot = Merge-Shapes 3 @("Src$script:made", "Frame$script:made"); $shot.Name = "Photo$script:made"
    $shot
}
function Add-Frame([int]$type, $left, $top, $width, $height) { $sl.Shapes.AddShape($type, [single]$left, [single]$top, [single]$width, [single]$height) }

try {
    New-Deck $Background
    $n = $Images.Count; $ink = ConvertTo-Bgr $Color; $rng = New-Object System.Random $Seed
    $textLeft = 60; $textTop = 60; $textW = $W * 0.5; $centred = $false
    switch ($Layout) {
        'Columns' {
            $gap = 4; $cw = ($W - $gap * ($n - 1)) / $n
            for ($i = 0; $i -lt $n; $i++) { [void](Add-Cropped $Images[$i] (Add-Frame 1 (($cw + $gap) * $i) 0 $cw $H)) }
            $ink = 0xFFFFFF; $textTop = $H - 150
        }
        'Mosaic' {
            if ($n -lt 3) { throw 'Mosaic needs three images' }
            $m = 30; $gap = 8; $lw = ($W - 2 * $m - $gap) * 0.42; $rw = ($W - 2 * $m - $gap) - $lw; $bh = ($H - 2 * $m - $gap) / 2
            [void](Add-Cropped $Images[0] (Add-Frame 1 $m $m $lw $bh))
            [void](Add-Cropped $Images[1] (Add-Frame 1 $m ($m + $bh + $gap) $lw $bh))
            [void](Add-Cropped $Images[2] (Add-Frame 1 ($m + $lw + $gap) $m $rw ($H - 2 * $m)))
            $ink = 0xFFFFFF; $textLeft = $m + $lw + $gap + 30; $textTop = $H - $m - 130
        }
        'Circles' {
            $big = 330; $small = 230; $cy = $H / 2 + 30; $mid = [math]::Floor($n / 2)
            $spread = ($big + $small) / 2 - 40                        # neighbours tuck 40pt behind the middle one
            for ($i = 0; $i -lt $n; $i++) {
                if ($i -eq $mid) { continue }
                $cx = $W / 2 + ($i - $mid) * $spread + [math]::Sign($i - $mid) * ([math]::Abs($i - $mid) - 1) * ($small - 40 - $spread)
                [void](Add-Cropped $Images[$i] (Add-Frame 9 ($cx - $small / 2) ($cy - $small / 2) $small $small))
            }
            $c = Add-Cropped $Images[$mid] (Add-Frame 9 ($W / 2 - $big / 2) ($cy - $big / 2) $big $big)
            $c.Line.Visible = -1; $c.Line.ForeColor.RGB = (ConvertTo-Bgr $Background); $c.Line.Weight = 6
            $textTop = 34; $textW = $W - 120
        }
        'Slant' {
            $lean = 70; $gap = 10; $top = 150; $bh = $H - $top - 40
            $cw = ($W - 120 + $lean - $gap * ($n - 1)) / $n
            for ($i = 0; $i -lt $n; $i++) {
                $f = Add-Frame 2 (60 + (($W - 120 - $cw) / [math]::Max(1, $n - 1)) * $i) $top $cw $bh      # msoShapeParallelogram
                $f.Adjustments.Item(1) = [single]($lean / $bh)
                [void](Add-Cropped $Images[$i] $f)
            }
            $textTop = 44
        }
        'Scatter' {
            $pw = [math]::Min(300, ($W - 120) / $n + 60); $ph = $pw * 0.78; $cy = $H / 2 + 40
            for ($i = 0; $i -lt $n; $i++) {
                $cx = 60 + $pw / 2 + (($W - 120 - $pw) / [math]::Max(1, $n - 1)) * $i
                $shot = Add-Cropped $Images[$i] (Add-Frame 1 ($cx - $pw / 2) ($cy - $ph / 2 + $rng.Next(-30, 31)) $pw $ph)
                $shot.Line.Visible = -1; $shot.Line.ForeColor.RGB = 0xFFFFFF; $shot.Line.Weight = 9
                $sh = $shot.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Transparency = 0.65; $sh.Blur = 14; $sh.OffsetX = 3; $sh.OffsetY = 6
                $shot.Rotation = [single]((($rng.Next(-9, 10)) + 360) % 360)
            }
            $textTop = 40
        }
        'Blob' {
            $cx = 320; $cy = $H / 2 - 20; $r = 180
            $fit = @{ L = $cx - 1.6 * $r; T = $cy - 1.6 * $r; W = 3.2 * $r; H = 3.2 * $r }
            # organic outline = a circle united with a tilted ellipse on its lower right; $k scales it about the centre
            function New-Blob([double]$k, [string]$tag) {
                $a = Add-Frame 9 ($cx - $r * $k) ($cy - $r * $k) (2 * $r * $k) (2 * $r * $k); $a.Name = "BA$tag"
                $b = Add-Frame 9 ($cx + (0.55 - 0.62) * $r * $k) ($cy + (0.52 - 0.48) * $r * $k) (1.24 * $r * $k) (0.96 * $r * $k); $b.Rotation = 42; $b.Name = "BB$tag"
                Merge-Shapes 1 @("BA$tag", "BB$tag")
            }
            foreach ($ring in @(1.2, 0.62), @(1.1, 0.4)) {             # faint echoes: the photo again, larger, under a veil of background colour
                [void](Add-Cropped $Images[0] (New-Blob $ring[0] "p$($ring[0])") $fit)
                $veil = New-Blob $ring[0] "v$($ring[0])"; $veil.Fill.ForeColor.RGB = (ConvertTo-Bgr $Background); $veil.Fill.Transparency = [single](1 - $ring[1]); $veil.Line.Visible = 0
            }
            [void](Add-Cropped $Images[0] (New-Blob 1 'main') $fit)
            $textLeft = $W * 0.62; $textTop = $H / 2 - 60; $textW = $W * 0.33
        }
        'Diagonal' {
            if ($n -lt 3) { throw 'Diagonal needs three images' }
            $g = 6; $x0 = $W * 0.52; $dx = $W * 0.16; $xm = $x0 - $dx / 2           # the cut runs from (x0,0) to (x0-dx,H)
            $polys = @(
                @(0, 0, $x0, 0, ($x0 - $dx), $H, 0, $H),
                @(($x0 + $g), 0, $W, 0, $W, ($H / 2 - $g / 2), ($xm + $g), ($H / 2 - $g / 2)),
                @(($xm + $g), ($H / 2 + $g / 2), $W, ($H / 2 + $g / 2), $W, $H, ($x0 - $dx + $g), $H))
            for ($i = 0; $i -lt 3; $i++) {
                $q = $polys[$i]; $fb = $sl.Shapes.BuildFreeform(0, [single]$q[0], [single]$q[1])
                for ($j = 2; $j -lt 8; $j += 2) { $fb.AddNodes(0, 0, [single]$q[$j], [single]$q[$j + 1]) }
                $fb.AddNodes(0, 0, [single]$q[0], [single]$q[1])
                [void](Add-Cropped $Images[$i] ($fb.ConvertToShape()))
            }
            $ink = 0xFFFFFF; $textLeft = 50; $textTop = $H - 130; $textW = $W * 0.3
        }
        'Frames' {
            if ($n -lt 3) { throw 'Frames needs three images' }
            $page = New-Rect 0 0 $W $H; $page.Fill.ForeColor.RGB = (ConvertTo-Bgr '#C9C9C9'); $page.Line.Visible = 0
            $cards = @(@(50, 40, 480, 460), @(560, 40, 350, 222), @(560, 278, 350, 222))
            for ($i = 0; $i -lt 3; $i++) {
                $c = $cards[$i]; $card = New-Rect $c[0] $c[1] $c[2] $c[3]; $card.Fill.ForeColor.RGB = 0xFFFFFF; $card.Line.Visible = 0
                $ih = $c[3] * 0.8; $iw = [math]::Min($c[2] * 0.62, $ih * 0.78)
                [void](Add-Cropped $Images[$i] (Add-Frame 1 ($c[0] + ($c[2] - $iw) / 2) ($c[1] + ($c[3] - $ih) / 2) $iw $ih))
            }
            $textTop = $H + 40                                         # no room for a title in this layout
        }
        'Arc' {
            $field = New-Rect 0 0 $W $H; $field.Fill.ForeColor.RGB = (ConvertTo-Bgr $Accent); $field.Line.Visible = 0
            $R = $W * 0.62; $ccx = $W * 0.8; $ccy = $H * 0.55
            [void](Add-Cropped $Images[0] (Add-Frame 9 ($ccx - $R) ($ccy - $R) (2 * $R) (2 * $R)))
            $R2 = $R + 26; $arc = Add-Frame 9 ($ccx - $R2) ($ccy - $R2) (2 * $R2) (2 * $R2); $arc.Fill.Visible = 0
            $arc.Line.ForeColor.RGB = (ConvertTo-Bgr '#E8453C'); $arc.Line.Weight = 3; $arc.Line.DashStyle = 4
            $dot = Add-Frame 9 90 ($H - 110) 34 34; $dot.Fill.ForeColor.RGB = (ConvertTo-Bgr '#E8453C'); $dot.Line.Visible = 0
            $textLeft = 50; $textTop = 70; $textW = $W * 0.22
        }
        'Grid' {
            $cols = $Grid[0]; $rows = $Grid[1]; $m = 60; $top = 124; $gap = 8
            $cw = ($W - 2 * $m - $gap * ($cols - 1)) / $cols; $ch = ($H - $top - 36 - $gap * ($rows - 1)) / $rows
            for ($i = 0; $i -lt $cols * $rows; $i++) {
                [void](Add-Cropped $Images[($i + [math]::Floor($i / $cols)) % $n] (Add-Frame 1 ($m + ($cw + $gap) * ($i % $cols)) ($top + ($ch + $gap) * [math]::Floor($i / $cols)) $cw $ch))
            }
            $rule = $sl.Shapes.AddLine([single]($W * 0.32), 104, [single]($W * 0.68), 104); $rule.Line.ForeColor.RGB = (ConvertTo-Bgr $Accent); $rule.Line.Weight = 1.5
            $textTop = 44; $textW = $W * 0.6; $centred = $true
        }
        'Callouts' {
            $pw = $W * 0.36; $ph = $H * 0.6; $px = ($W - $pw) / 2; $py = $H - $ph - 36; $m = 60
            [void](Add-Cropped $Images[0] (Add-Frame 1 $px $py $pw $ph))
            $caps = @($Captions -split '\|' | Where-Object { $_ }); $bh = $ph * 0.26; $bw = $px + $pw / 3 - $m
            for ($i = 0; $i -lt [math]::Min(4, $caps.Count); $i++) {
                $isLeft = ($i % 2 -eq 0); $row = [math]::Floor($i / 2); $part = $caps[$i] -split ';', 2
                $x = $m; $y = $py + $ph * 0.06 + $row * $ph * 0.38
                if (-not $isLeft) { $x = $px + $pw * 2 / 3; $y += $ph * 0.13 }      # right pair sits half a step lower
                $bar = New-Rect ([single]$x) ([single]$y) ([single]$bw) ([single]$bh); $bar.Line.Visible = 0
                $bar.Fill.ForeColor.RGB = (ConvertTo-Bgr '#6E6E6E'); $bar.Fill.Transparency = 0.3; $bar.Name = "Bar$($i + 1)"
                $tw = $px - $m - 28; $ty = $y + 12
                foreach ($line in @(@($part[0], 15, $Font), @($(if ($part.Count -gt 1) { $part[1] } else { '' }), 11, 'Arial'))) {
                    if (-not $line[0]) { continue }
                    $c = New-FitText $line[0] $line[1] 0xFFFFFF $line[2] $tw; $c.Top = [single]$ty; $ty += [double]$c.Height + 4
                    if ($isLeft) { $c.Left = [single]($x + 14) } else { $c.TextFrame.TextRange.ParagraphFormat.Alignment = 3; $c.Left = [single]($x + $bw - 14 - [double]$c.Width) }
                }
            }
            $textTop = 36; $textW = $W * 0.7; $centred = $true
        }
        'Brush' {
            # torn right edge: a polygon whose right side zigzags by random amounts
            $edge = $W * 0.6; $fb = $sl.Shapes.BuildFreeform(0, 0, 0)
            $fb.AddNodes(0, 0, [single]$edge, 0)
            for ($y = 0; $y -lt $H; $y += 9) {
                $fb.AddNodes(0, 0, [single]($edge + $rng.Next(-70, 40)), [single]($y + 4))
                $fb.AddNodes(0, 0, [single]($edge + $rng.Next(-110, -20)), [single]($y + 9))
            }
            $fb.AddNodes(0, 0, 0, $H); $fb.AddNodes(0, 0, 0, 0)
            [void](Add-Cropped $Images[0] ($fb.ConvertToShape()))
            $textLeft = $W * 0.64; $textTop = $H / 2 - 70; $textW = $W * 0.32
        }
    }
    $y = $textTop
    if ($Title) { $t = New-FitText $Title 40 $ink $Font $textW; $t.Left = [single]$textLeft; $t.Top = [single]$y; $t.Name = 'Title'; $y += [double]$t.Height + 8; if ($centred) { $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2; $t.Left = [single](($W - [double]$t.Width) / 2) } }
    if ($Body) { $b = New-FitText $Body 16 $ink 'Arial' $textW; $b.Left = [single]$textLeft; $b.Top = [single]$y; $b.Name = 'Body'; if ($centred) { $b.Left = [single](($W - [double]$b.Width) / 2) } }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
