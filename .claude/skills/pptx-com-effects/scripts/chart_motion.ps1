# Charts and headline figures drawn from plain shapes (no chart objects), each with its own motion. Same colour / font parameters as scene_page.ps1.
#   Stream      stream graph: the layers bloom open from one thin line (Morph), then a light beam sweeps across and the legend follows
#               -Title T  -Sub 'big figure;caption'  -Cards 'x label|x label|...'  -Items 'layer;v1;v2;...|...' (2-6 layers)  [-Images night photo]
#   Race        bar chart race: one slide per year, the bars change length and overtake each other (Morph); the leader takes the accent colour
#               -Title T  [-Sub unit note]  -Cards 'year|year|...'  -Items 'name;v1;v2;...|...' (3-7 names)
#   Swarm       100 dots = 100 people in sunflower-packed clusters; the dots that change group light up and fly across (Morph)
#               -Title T  -Cards 'before|after'  -Items 'group;before;after|...' (2-4 groups, both columns must add up to the same total)
#   GifNumber   one enormous figure whose digits show an animated GIF (or a photo), echo outlines behind it, facts on the right
#               -Figure '87%'  -Images gif-or-photo  -Title T  [-Sub line under the figure]  -Items 'value;line|...' (2-3)
#   Odometer    a counter: every digit sits in its own tile and spins more than a full turn to the final value (Morph); then a small trend line draws itself
#               -Figure '1,284,650'  [-Sub unit]  -Title T  -Cards 'caption line|caption line'  -Items 'label;value|...' (trend, 3-8 points)
#   GlassStats  three frosted-glass cards (blurred pieces of the photo itself) rise at different speeds over the photo (Morph); a level bar fills in each
#               -Images photo  -Title T  -Items 'value;unit;label;share 0..1|...' (2-4)
# -Accent2 is the second colour of the two-tone gradients (Stream layers); default: half way from -Accent to -Light.
# All six play by themselves.
# Example: .\chart_motion.ps1 -Layout Race -Title 'Rides by city' -Cards '2023|2024|2025' -Items 'A;10;14;22|B;12;13;15|C;6;15;19' -Out C:\work\race.pptx -Preview C:\work\race.mp4
param(
    [Parameter(Mandatory)][ValidateSet('Stream', 'Race', 'Swarm', 'GifNumber', 'Odometer', 'GlassStats')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Figure = '',
    [string]$Note = '',
    [string]$Accent = '#7C6BFF',
    [string]$Accent2 = '',
    [string]$AccentDark = '',
    [string]$Light = '#FCEFE6',
    [string]$Dark = '#241C45',
    [string]$Ink = '#1F1B2E',
    [string]$Font = 'Malgun Gothic',
    [string]$HeavyFont = '',
    [double]$Duration = 0.9,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent
. (Join-Path $PSScriptRoot 'scene_helpers.ps1')
if (-not $Accent2) { $Accent2 = Mix $Accent $Light 0.5 }
function Steps([int]$count) { for ($i = 1; $i -lt $count; $i++) { Add-MorphSlide $Duration } }
function Auto($slide, $sec) { $slide.SlideShowTransition.AdvanceOnTime = -1; $slide.SlideShowTransition.AdvanceTime = [single]$sec }
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Place($shape, $x, $y) { $shape.Left = [single]$x; $shape.Top = [single]$y }
function Tail($r, [int]$from = 1) { if ($r.Count -gt $from) { ($r[$from..($r.Count - 1)]) -join "`r" } else { '' } }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Path($pts, [string]$hex, $weight) {                           # open line through the points
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $s = Tag ($fb.ConvertToShape()); $s.Fill.Visible = 0; $s.Line.ForeColor.RGB = (Rgb $hex); $s.Line.Weight = [single]$weight; $s
}
function Grad($shape, [string]$c1, [string]$c2, $angle = 0) {          # two-colour fill
    $shape.Fill.TwoColorGradient(1, 1); $shape.Fill.GradientAngle = [single]$angle
    $gs = $shape.Fill.GradientStops; $gs.Item(1).Color.RGB = (Rgb $c1); $gs.Item(2).Color.RGB = (Rgb $c2)
}
function Need1 { if (-not $Images.Count) { throw "$Layout needs -Images <picture>" } }
function NeedFigure { if (-not $Figure) { throw "$Layout needs -Figure" } }
function Grow($shape, $factor) {                                       # scale about the centre
    $cxs = [double]$shape.Left + [double]$shape.Width / 2; $cys = [double]$shape.Top + [double]$shape.Height / 2
    $shape.LockAspectRatio = -1; $shape.Width = [single]([double]$shape.Width * $factor); Place $shape ($cxs - [double]$shape.Width / 2) ($cys - [double]$shape.Height / 2)
}

try {
    New-Deck $(if ($Layout -in 'Race', 'Odometer') { $Light } else { $Dark })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count; $labels = @($Cards -split '\|'); $hold = 3

    switch ($Layout) {
        'Stream' {
            if ($Images.Count) { (Add-Cover $Images[0]).Name = 'Photo'; $veil = Box 0 0 $W $H $Dark; $veil.Fill.Transparency = 0.58 }
            Head $true; Foot $true
            $m = $rows[0].Count - 1; $seg = 10; $np = ($m - 1) * $seg + 1; $x0 = 65; $x1 = 620; $cy = 332; $maxH = 262
            $val = New-Object 'double[,]' $n, $np; $tot = New-Object 'double[]' $np
            for ($q = 0; $q -lt $np; $q++) {                                                                                    # eased between the data points: soft shoulders instead of corners
                $j = [int][math]::Min($m - 2, [math]::Floor($q / $seg)); $u = ($q - $j * $seg) / $seg; $e = (1 - [math]::Cos([math]::PI * $u)) / 2
                for ($i = 0; $i -lt $n; $i++) { $a = [double]$rows[$i][$j + 1]; $b = [double]$rows[$i][$j + 2]; $val[$i, $q] = $a + ($b - $a) * $e; $tot[$q] += $val[$i, $q] }
            }
            $k = $maxH / ($tot | Measure-Object -Maximum).Maximum; $cum = @($tot | ForEach-Object { - $_ * $k / 2 })           # centred on the middle line
            for ($j = 0; $j -lt $m; $j++) {
                $gx = $x0 + ($x1 - $x0) * $j / ($m - 1); $ln = Box ($gx - 0.5) 192 1 280 '#FFFFFF'; $ln.Fill.Transparency = 0.86
                if ($j -lt $labels.Count) { [void](Tx $labels[$j] 15 $dText ($gx - 40) 482 80 0 $Font 2) }
            }
            for ($i = 0; $i -lt $n; $i++) {
                $col = Mix $Accent $Accent2 ($i / [math]::Max(1, $n - 1)); $up = @(); $dn = @()
                for ($q = 0; $q -lt $np; $q++) { $px = $x0 + ($x1 - $x0) * $q / ($np - 1); $up += , @($px, ($cy + $cum[$q])); $cum[$q] += $val[$i, $q] * $k; $dn += , @($px, ($cy + $cum[$q])) }
                [array]::Reverse($dn); $lay = Poly ($up + $dn) $col; Grad $lay (Mix $col $Dark 0.55) $col 0
                $lay.Line.Visible = -1; $lay.Line.ForeColor.RGB = (Rgb $Dark); $lay.Line.Weight = 1.5; $lay.Name = "L$i"
            }
            $rx = 665; $rw = 235; $hero = @($Sub -split ';')
            if ($Sub) { [void](Tx $hero[0] 54 '#FFFFFF' $rx 150 $rw $heavyBold $heavy); if ($hero.Count -gt 1) { [void](Tx $hero[1] 17 $dText $rx 220 $rw 0 $Font 1 0 1.3) } }
            $Duration = 1.6; Steps 2
            for ($i = 0; $i -lt $n; $i++) { $p = $pres.Slides(1).Shapes("L$i"); $p.Height = 2; $p.Top = [single]($cy - 1) }   # slide 1: every layer pressed into one line
            $beam = Box ($x0 - 1) 188 2 288 '#FFFFFF'; Glow $beam $Accent2 0.15 10 0
            $fx = $sl.TimeLine.MainSequence.AddEffect($beam, 0, 0, $AFTER); $fx.Behaviors.Add(1).MotionEffect.Path = ('M 0 0 L {0:F4} 0 E' -f (($x1 - $x0) / $W)); $fx.Timing.Duration = 1.8
            $ly = 282; $lp = [math]::Min(42, 190 / $n)
            for ($i = 0; $i -lt $n; $i++) {
                $col = Mix $Accent $Accent2 ($i / [math]::Max(1, $n - 1))
                $set = @((Disc ($rx + 7) ($ly + 12) 6 $col), (Tx $rows[$i][0] 17 '#FFFFFF' ($rx + 24) $ly 140), (Tx $rows[$i][$m] 17 $col $rx $ly $rw -1 $Font 3))
                Anim (Grp $set) $FADE $WITH 0.5 (0.2 + 1.5 * ($i + 0.5) / $n); $ly += $lp
            }
            Auto $pres.Slides(1) 0.5; $hold = 5
        }
        'Race' {
            Head; Foot; $ny = $rows[0].Count - 1; $vmax = 0.0
            foreach ($r in $rows) { for ($k = 1; $k -le $ny; $k++) { $vmax = [math]::Max($vmax, [double]$r[$k]) } }
            $top = 150; $pitch = [math]::Min(58, 342 / $n); $bh = $pitch * 0.62; $bx = 205; $bmax = 560
            (Tx $labels[0] 190 (Mix $Light $Ink 0.1) ($W - 55 - 580) 282 580 $heavyBold $heavy 3).Name = 'Year'                   # the year sits behind the bars as a watermark
            if ($Sub) { [void](Tx $Sub 15 $gray ($W - 65 - 300) 122 300 0 $Font 3) }
            for ($r = 0; $r -lt $n; $r++) { [void](Mid (Tx ('{0:00}' -f ($r + 1)) 16 $gray 65 0 30 -1) ($top + $r * $pitch) $pitch) }   # the rank numbers stay; the names move past them
            for ($i = 0; $i -lt $n; $i++) {
                (Box $bx $top 20 $bh $Ink 5).Name = "Bar$i"
                (Mid (Tx $rows[$i][0] 19 $Ink 100 0 100 -1) $top $pitch).Name = "Name$i"
                (Mid (Tx '0' 22 $Ink $bx 0 140 $heavyBold $heavy) $top $pitch).Name = "Val$i"
            }
            $Duration = 1.1; Steps $ny
            for ($k = 0; $k -lt $ny; $k++) {
                $slide = $pres.Slides($k + 1); if ($k -lt $labels.Count) { $slide.Shapes('Year').TextFrame.TextRange.Text = $labels[$k] }
                $order = @(0..($n - 1) | Sort-Object { - [double]$rows[$_][$k + 1] })
                for ($r = 0; $r -lt $n; $r++) {
                    $i = $order[$r]; $y = $top + $r * $pitch; $bwid = [math]::Max(10, $bmax * [double]$rows[$i][$k + 1] / $vmax)
                    $b = $slide.Shapes("Bar$i"); $b.Top = [single]($y + ($pitch - $bh) / 2); $b.Width = [single]$bwid; $b.Adjustments.Item(1) = [single][math]::Min(0.5, 5 / [math]::Min($bwid, $bh))
                    $b.Fill.ForeColor.RGB = (Rgb $(if ($r -eq 0) { $Accent } else { Mix $Ink $Light (0.12 + 0.6 * $r / $n) }))
                    $pen = Rgb $(if ($r -eq 0) { $AccentDark } else { $Ink })
                    $nm = $slide.Shapes("Name$i"); $nm.Top = [single]$y; $nm.TextFrame.TextRange.Font.Color.RGB = $pen
                    $v = $slide.Shapes("Val$i"); $v.Top = [single]$y; $v.Left = [single]($bx + $bwid + 12); $v.TextFrame.TextRange.Text = $rows[$i][$k + 1]; $v.TextFrame.TextRange.Font.Color.RGB = $pen
                }
                if ($k -lt $ny - 1) { Auto $slide 0.8 }
            }
            $hold = 2
        }
        'Swarm' {
            Head $true; Foot $true; $rd = 7; $sp0 = 10.6; $cy = 290; $dim = Mix $Dark '#FFFFFF' 0.4
            $c1 = @(); $k1 = @(); $c2 = @(); $k2 = @(); $pool = New-Object System.Collections.Queue; $sumA = 0
            for ($c = 0; $c -lt $n; $c++) {                                                                                    # the outer dots of a shrinking group are the ones that leave
                $bf = [int]$rows[$c][1]; $af = [int]$rows[$c][2]; $stay = [math]::Min($bf, $af); $sumA += $af
                for ($k = 0; $k -lt $bf; $k++) { $c1 += $c; $k1 += $k; if ($k -lt $stay) { $c2 += $c; $k2 += $k } else { $c2 += -1; $k2 += 0; $pool.Enqueue($c1.Count - 1) } }
            }
            if ($sumA -ne $c1.Count) { throw "Swarm: the before and after columns must add up to the same total ($($c1.Count) vs $sumA)" }
            for ($c = 0; $c -lt $n; $c++) { for ($k = [int]$rows[$c][1]; $k -lt [int]$rows[$c][2]; $k++) { $d = $pool.Dequeue(); $c2[$d] = $c; $k2[$d] = $k } }
            function At($c, $k) { $ang = $k * 2.39996; $rr = $sp0 * [math]::Sqrt($k + 0.5); @((65 + ($W - 130) * ($c + 0.5) / $n + $rr * [math]::Cos($ang)), ($cy + $rr * [math]::Sin($ang))) }   # sunflower packing
            for ($d = 0; $d -lt $c1.Count; $d++) { $p = At ($c1[$d]) ($k1[$d]); (Disc $p[0] $p[1] $rd $dim).Name = "D$d" }
            for ($c = 0; $c -lt $n; $c++) {
                $mx = 65 + ($W - 130) * ($c + 0.5) / $n
                (Tx $rows[$c][1] 50 '#FFFFFF' ($mx - 110) 386 220 $heavyBold $heavy 2).Name = "Cnt$c"
                [void](Tx $rows[$c][0] 18 $dText ($mx - 110) 452 220 -1 $Font 2)
            }
            (Tx $labels[0] 64 '#FFFFFF' ($W - 65 - 300) 32 300 $heavyBold $heavy 3).Name = 'Year'
            $Duration = 1.9; Steps 2; $s2 = $sl
            for ($d = 0; $d -lt $c1.Count; $d++) {
                if ($c2[$d] -eq $c1[$d]) { continue }
                $p = At ($c2[$d]) ($k2[$d]); $o = $s2.Shapes("D$d"); Place $o ($p[0] - $rd) ($p[1] - $rd); $o.Fill.ForeColor.RGB = (Rgb $Accent); Glow $o $Accent 0.4 8 0
            }
            for ($c = 0; $c -lt $n; $c++) {
                $diff = [int]$rows[$c][2] - [int]$rows[$c][1]; $mx = 65 + ($W - 130) * ($c + 0.5) / $n; $t = $s2.Shapes("Cnt$c"); $t.TextFrame.TextRange.Text = $rows[$c][2]
                if ($diff -gt 0) { $t.TextFrame.TextRange.Font.Color.RGB = (Rgb $Accent) }
                if ($diff) { [void](Tx ('{0:+#;-#}' -f $diff) 19 $(if ($diff -gt 0) { $Accent } else { $dText }) ($mx + 46) 398 90 -1) }
            }
            if ($labels.Count -gt 1) { $s2.Shapes('Year').TextFrame.TextRange.Text = $labels[1] }
            Auto $pres.Slides(1) 1.2; $hold = 3
        }
        'GifNumber' {
            Need1; NeedFigure; Head $true; Foot $true
            $t = New-Text $Figure 250 0xFFFFFF $heavy; if (-not $HeavyFont) { $t.TextFrame.TextRange.Font.Bold = -1 }
            if ([double]$t.Width -gt 530) { $t.TextFrame.TextRange.Font.Size = [single][math]::Floor(250 * 530 / [double]$t.Width) }
            Place $t 58 (300 - [double]$t.Height / 2)
            $glyph = New-GlyphShape $t; $t.Delete(); $gb = Get-Box $glyph
            foreach ($e in 3, 2, 1) {                                                                                           # echo outlines stepping away behind the figure
                $o = $glyph.Duplicate(); $o.Name = "Echo$e"; Place $o ($gb.L + 9 * $e) ($gb.T + 9 * $e); $o.Fill.Visible = 0
                $o.Line.Visible = -1; $o.Line.ForeColor.RGB = (Rgb $Accent); $o.Line.Weight = 1.5; $o.Line.Transparency = [single](0.1 + 0.25 * $e)
            }
            $g = $sl.Shapes.AddPicture($Images[0], 0, -1, 0, 0, -1, -1); $g.LockAspectRatio = -1; $g.Height = [single]($gb.H + 30); if ([double]$g.Width -lt $gb.W + 30) { $g.Width = [single]($gb.W + 30) }
            Place $g ($gb.L + ($gb.W - [double]$g.Width) / 2) ($gb.T + ($gb.H - [double]$g.Height) / 2)
            $num = Merge-Shapes 3 @($g.Name, $glyph.Name); $num.Name = 'Number'                                                 # picture first: the result stays a (playing) picture
            if ($Sub) { [void](Tx $Sub 20 '#FFFFFF' 62 ($gb.T + $gb.H + 44) 540 -1 $Font 1 0 1.3) }
            foreach ($e in 1, 2, 3) { Anim $sl.Shapes("Echo$e") $FADE $AFTER 0.3 }
            $rx = 640; $rw = 255; $ry = 150; $rp = [math]::Min(120, 342 / $n)
            for ($i = 0; $i -lt $n; $i++) {
                $set = @((Box $rx $ry $rw 1.2 (Mix $Dark '#FFFFFF' 0.4)), (Tx $rows[$i][0] 40 $Accent $rx ($ry + 10) $rw $heavyBold $heavy), (Tx (Tail $rows[$i]) 17 (Mix $Dark '#FFFFFF' 0.9) $rx ($ry + 59) $rw 0 $Font 1 0 1.3))
                Anim (Grp $set) $RISE $AFTER 0.5; $ry += $rp
            }
            $hold = 5
        }
        'Odometer' {
            NeedFigure; $chars = $Figure.ToCharArray(); $nd = @($chars | Where-Object { [char]::IsDigit($_) }).Count; $ns = $chars.Count - $nd
            $size = 112.0; $unitW = $(if ($Sub) { Wd $Sub 34 } else { 0 }); $room = $W - 130 - $unitW
            $need = $nd * $size * 0.76 + ($chars.Count - 1) * 7 + $ns * $size * 0.2; if ($need -gt $room) { $size = $size * $room / $need }
            $tw = $size * 0.76; $sw = $size * 0.2; $winTop = 150; $x = 65; $tiles = @(); $seps = @(); $digits = @(); $tops = $null; $lineH = 0
            $seq = ((0..9) + (0..9)) -join '|'                                                                                  # two rounds of digits: every strip turns at least once
            foreach ($ch in $chars) {
                if (-not [char]::IsDigit($ch)) { $seps += , @($x, "$ch"); $x += $sw + 7; continue }
                $tiles += (Box $x $winTop $tw 20 $Dark 10)
                $t = Tx $seq $size $Light $x $winTop $tw $heavyBold $heavy 2 0 0.92; $t.Name = "Reel$($digits.Count)"; $digits += [int]"$ch"
                if (-not $tops) { $tr = $t.TextFrame.TextRange; $tops = @(1..20 | ForEach-Object { [double]$tr.Lines($_).BoundTop - [double]$t.Top }); $lineH = ($tops[19] - $tops[0]) / 19 }
                $t.Top = [single]($winTop - $tops[0]); $x += $tw + 7
            }
            foreach ($tl in $tiles) { $tl.Height = [single]$lineH; $tl.Adjustments.Item(1) = [single](10 / [math]::Min($tw, $lineH)) }
            [void](Box 0 0 $W $winTop $Light); [void](Box 0 ($winTop + $lineH) $W ($H - $winTop - $lineH) $Light)               # only one row of digits shows between the two covers
            foreach ($tl in $tiles) { $seam = Box $tl.Left ($winTop + $lineH / 2 - 0.6) $tw 1.2 $Dark; $seam.Fill.Transparency = 0.5 }
            foreach ($s in $seps) { [void](Tx $s[1] $size $Ink $s[0] ($winTop - $tops[0]) $sw $heavyBold $heavy 2 0 0.92) }
            if ($Sub) { [void](Tx $Sub 34 $AccentDark ($x + 4) ($winTop + $lineH - 50) ($unitW + 20) -1) }
            Head; Foot
            $y0 = $winTop + $lineH + 44; [void](Box 65 $y0 46 4 $Accent 2)
            if ($Cards) { [void](Tx $Cards 19 (Mix $Ink '#FFFFFF' 0.22) 65 ($y0 + 22) 370 0 $Font 1 0 1.5) }
            $Duration = 2.4; Steps 2
            for ($d = 0; $d -lt $digits.Count; $d++) { $sl.Shapes("Reel$d").Top = [single]($winTop - $tops[10 + $digits[$d]]) }
            if ($n -gt 1) {                                                                                                     # the trend line exists on slide 2 only and draws itself after the roll
                $gx0 = 490; $gx1 = 880; $gy0 = $y0 + 34; $gy1 = 462; $vals = @($rows | ForEach-Object { [double]$_[1] })
                $hi = ($vals | Measure-Object -Maximum).Maximum; $lo = ($vals | Measure-Object -Minimum).Minimum; $lo = $lo - ($hi - $lo) * 0.3
                $pts = @(); for ($i = 0; $i -lt $n; $i++) { $pts += , @(($gx0 + ($gx1 - $gx0) * $i / ($n - 1)), ($gy1 - ($gy1 - $gy0) * ($vals[$i] - $lo) / ($hi - $lo))) }
                $area = Poly ($pts + @(, @($gx1, $gy1)) + @(, @($gx0, $gy1))) $Accent; Grad $area (Mix $Accent $Light 0.35) $Light 90
                $line = Path $pts $Accent 3.5
                Anim $area $WIPE $AFTER 1.1 0.1 4; Anim $line $WIPE $WITH 1.1 0.1 4
                for ($i = 0; $i -lt $n; $i++) { [void](Tx $rows[$i][0] 15 $gray ($pts[$i][0] - 40) ($gy1 + 8) 80 0 $Font 2) }
                $end = $pts[$n - 1]; $dot = Disc $end[0] $end[1] 8 $AccentDark; Glow $dot $Accent 0.3 10 0
                $tag = Tx $(if ($rows[$n - 1].Count -gt 2) { $rows[$n - 1][2] } else { $rows[$n - 1][1] }) 21 $AccentDark ($end[0] - 150) ($end[1] - 40) 150 $heavyBold $heavy 3
                Anim (Grp @($dot, $tag)) $FADE $AFTER 0.4
            }
            Auto $pres.Slides(1) 0.4; $hold = 5
        }
        'GlassStats' {
            Need1; $zoom = 1.1; $photo = Add-Cover $Images[0]; Grow $photo $zoom; $photo.PictureFormat.Brightness = 0.42; $photo.Name = 'Photo'
            $gap = 26; $cw = [math]::Min(262, ($W - 130 - $gap * ($n - 1)) / $n); $ch = 252; $xa = ($W - ($n * $cw + ($n - 1) * $gap)) / 2; $stag = 196, 164, 214, 180
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $x = $xa + $i * ($cw + $gap); $y = $stag[$i % 4]
                $g = Add-Cover $Images[0]; Grow $g $zoom; $g.AutoShapeType = 5; $g.Adjustments.Item(1) = 0.08
                $c = $g.PictureFormat.Crop; $c.ShapeLeft = [single]$x; $c.ShapeWidth = [single]$cw; $c.ShapeTop = [single]$y; $c.ShapeHeight = [single]$ch    # same picture, same place: only the window is smaller
                $e = $g.Fill.PictureEffects.Insert(2); $e.EffectParameters(1).Value = 70; $g.PictureFormat.Brightness = 0.56
                $g.Line.Visible = -1; $g.Line.ForeColor.RGB = 0xFFFFFF; $g.Line.Transparency = 0.5; $g.Line.Weight = 1.25; $g.Name = "Glass$i"
                $vt = Tx $r[0] 56 '#FFFFFF' ($x + 26) ($y + 28) ($cw - 40) $heavyBold $heavy; $set = @($vt)
                if ($r.Count -gt 1 -and $r[1]) { $set += (Tx $r[1] 19 $Accent ($x + 26 + [double]$vt.TextFrame.TextRange.BoundWidth + 4) ($y + 28 + 36) 90 -1) }
                if ($r.Count -gt 2) { $set += (Tx $r[2] 19 '#FFFFFF' ($x + 26) ($y + 122) ($cw - 52) -1 $Font 1 0 1.35) }
                $track = Box ($x + 26) ($y + $ch - 38) ($cw - 52) 6 '#FFFFFF' 3; $track.Fill.Transparency = 0.72; $set += $track
                (Grp $set).Name = "Copy$i"
                if ($r.Count -gt 3) { (Box ($x + 26) ($y + $ch - 38) (($cw - 52) * [double]$r[3]) 6 $Accent 3).Name = "Fill$i" }
            }
            $head = @(); if ($Kicker) { $head += (Tx $Kicker 12.5 $Accent 65 45 600 -1 $Font 1 3) }; $head += (Tx $Title 32 '#FFFFFF' 65 68 830 -1); (Grp $head).Name = 'Head'
            Foot $true
            $Duration = 1.5; Steps 2; $s1 = $pres.Slides(1)
            Grow $s1.Shapes('Photo') (1 / $zoom); $s1.Shapes('Head').Top = 400
            for ($i = 0; $i -lt $n; $i++) {
                $c = $s1.Shapes("Glass$i").PictureFormat.Crop; $c.ShapeHeight = 4; $c.ShapeTop = [single]($H - 5)
                $s1.Shapes("Copy$i").Top = [single]($H + 40 + 170 * $i)                                                          # further down toward the right: the cards land one after another
                if ($rows[$i].Count -gt 3) { $s1.Shapes("Fill$i").Delete(); Anim $sl.Shapes("Fill$i") $WIPE $(if ($i) { $WITH } else { $AFTER }) 0.9 (0.15 * $i) 4 }
            }
            Auto $s1 0.9; $hold = 4
        }
    }
    Save-Deck $Out $Preview $hold
    "saved $Out"
} catch { Close-Deck; throw }
