# Covers, contents and pitch scenes with their own movement (2026-10 batch, part 4). Same colour / font parameters as scene_page.ps1.
#   Policy     ministry policy cover: title slides in, a slanted slab with a giant numeral closes from the right, a seal lands, a ticker runs
#              -Kicker k  -Title 'line//line'  -Sub s  -Cards 'big;label|seal line//seal line'  -Items 'ticker word|word|...'  -Note 'organisation  date'
#   Numeral    typographic cover: giant numerals rise from behind a horizon line, the title stands vertically, a corner of the page folds
#              -Title 'column//column'  -Cards '27'  -Kicker k  -Sub 'line//line'  -Items 'right line|right line'  -Note n
#   Shelf      contents: book spines with vertical titles fall onto a shelf, the last one leans   -Title T  -Sub s  -Items 'head;small|...' (3-6)
#   Export     world map: arcs leave one country one after another, each lands with a pulse and a label; a running total counts up
#              -Title T  -Cards 'South Korea;label'  -Sub 'unit;caption;right line'  -Items 'Country;label;number|...' (2-6)
#   KoreaPins  map of Korea (assets\korea_sido.json): pins drop one per slide, a route draws, the side card shows the stop
#              -Title T  -Sub 'intro line'  -Items 'name;lon;lat;text|...' (2-7)
#   AppIntro   product intro: screenshots as phones that rise and tilt in perspective (3-D rotation), callouts pop in
#              -Images shot1,shot2,shot3  -Title T  -Sub s  -Items 'callout;feature line|...' (one per image)
#   WhyUs      pitch "why us": the problem gets struck through, a stage spotlight falls on our column, proof figures rise
#              -Title T  -Sub 'label;problem text'  -Cards 'label;answer head;answer line|name;score|name;score|...(ours last)'  -Items 'value;unit;caption|...' (3)
#   Strategy   pitch summary: strategies on a slanted panel, one budget bar cut into pieces, half-dial gauges whose needles swing (Morph)
#              -Title T  -Items 'head;line|...' (3)  -Cards 'label;amount|...'  -Sub 'budget caption;unit;gauge caption'  -Kpi 'caption;shown;0..1|...' (3)
# Export draws assets\world_110m.json (Natural Earth, public domain; country names as in that file). All layouts play by themselves.
# Example: .\cover_motion.ps1 -Layout Shelf -Title 'CONTENTS' -Items 'Why;now|What;plan|How;budget' -Out C:\work\shelf.pptx -Preview C:\work\shelf.mp4
param(
    [Parameter(Mandatory)][ValidateSet('Policy', 'Numeral', 'Shelf', 'Export', 'KoreaPins', 'AppIntro', 'WhyUs', 'Strategy')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string]$Kpi = '',                             # Strategy only: gauges
    [string[]]$Images = @(),
    [string]$Map = '',                             # outline file (default: world_110m.json, korea_sido.json for KoreaPins)
    [string]$Note = '',
    [string]$Accent = '#7C6BFF',
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
if (-not $Map) { $Map = Join-Path (Split-Path $PSScriptRoot) $(if ($Layout -eq 'KoreaPins') { 'assets\korea_sido.json' } else { 'assets\world_110m.json' }) }
if ($Layout -in 'Export', 'KoreaPins' -and -not (Test-Path $Map)) { throw "map outline not found: $Map" }
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent
. (Join-Path $PSScriptRoot 'scene_helpers.ps1')
$ZOOM = 53                                                             # faded zoom entrance
function Steps([int]$count) { for ($i = 1; $i -lt $count; $i++) { Add-MorphSlide $Duration } }
function Auto($slide, $sec) { $slide.SlideShowTransition.AdvanceOnTime = -1; $slide.SlideShowTransition.AdvanceTime = [single]$sec }
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Place($shape, $x, $y) { $shape.Left = [single]$x; $shape.Top = [single]$y }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Stroke($x1, $y1, $x2, $y2, [string]$hex, $weight) { $s = Tag ($sl.Shapes.AddLine([single]$x1, [single]$y1, [single]$x2, [single]$y2)); $s.Line.ForeColor.RGB = (Rgb $hex); $s.Line.Weight = [single]$weight; $s }
function Grow($shape, $factor) {                                       # scale about the centre
    $cxs = [double]$shape.Left + [double]$shape.Width / 2; $cys = [double]$shape.Top + [double]$shape.Height / 2
    $shape.LockAspectRatio = -1; $shape.Width = [single]([double]$shape.Width * $factor); Place $shape ($cxs - [double]$shape.Width / 2) ($cys - [double]$shape.Height / 2)
}
function VTx([string]$text, $size, [string]$hex, $x, $y, $bh, [int]$bold = 0, [string]$face = $Font) {   # vertical (top-to-bottom) text; "//" starts the next column to the left
    $t = Tag ($sl.Shapes.AddTextbox(6, [single]$x, [single]$y, 40, [single]$bh))
    $f = $t.TextFrame; $f.WordWrap = 0; $f.AutoSize = 1; $f.MarginLeft = 0; $f.MarginRight = 0; $f.MarginTop = 0; $f.MarginBottom = 0
    $r = $f.TextRange; $r.Text = ($text -replace '\s*//\s*', "`r"); $r.Font.Name = $face; $r.Font.NameFarEast = $face; $r.Font.Size = [single]$size
    $r.Font.Color.RGB = (Rgb $hex); $r.Font.Bold = $bold; Place $t $x $y
    $t
}
function FadeOut($shape, $dur, $delay) { $e = $sl.TimeLine.MainSequence.AddEffect($shape, 10, 0, 2); $e.Exit = -1; $e.Timing.Duration = [single]$dur; $e.Timing.TriggerDelayTime = [single]$delay }
function WipeDir($dx, $dy) { if ([math]::Abs($dx) -lt 0.5 * [math]::Abs($dy)) { if ($dy -gt 0) { 1 } else { 3 } } elseif ($dx -gt 0) { 4 } else { 2 } }
function Need1 { if (-not $Images.Count) { throw "$Layout needs -Images <picture>" } }
function DrawMap($x, $y, $bw, $bh, [string]$fill, [string]$edge) {    # one freeform per outline, named after the region, grouped as 'Map'
    $script:mapInfo = Get-Content $Map -Raw -Encoding UTF8 | ConvertFrom-Json
    $sc = [math]::Min($bw / [double]$mapInfo.width, $bh / [double]$mapInfo.height); $mw = [double]$mapInfo.width * $sc; $mh = [double]$mapInfo.height * $sc
    $ml = $x + ($bw - $mw) / 2; $mt = $y + ($bh - $mh) / 2; $names = @()
    foreach ($c in $mapInfo.shapes) {
        $k = 0
        foreach ($ring in $c.rings) {
            $k++; $fb = $sl.Shapes.BuildFreeform(0, [single]($ml + $mw * $ring[0][0]), [single]($mt + $mh * $ring[0][1]))
            for ($q = 1; $q -lt $ring.Count; $q++) { $fb.AddNodes(0, 0, [single]($ml + $mw * $ring[$q][0]), [single]($mt + $mh * $ring[$q][1])) }
            $f = $fb.ConvertToShape(); $f.Fill.ForeColor.RGB = (Rgb $fill); $f.Line.ForeColor.RGB = (Rgb $edge); $f.Line.Weight = 0.4
            $f.Name = $(if ($k -eq 1) { $c.name } else { "$($c.name) $k" }); $names += $f.Name
        }
    }
    $g = $sl.Shapes.Range([string[]]$names).Group(); $g.Name = 'Map'
    @{ L = [double]$g.Left; T = [double]$g.Top; W = [double]$g.Width; H = [double]$g.Height; ML = $ml; MT = $mt; MW = $mw; MH = $mh }
}
function Spot([string]$country) {                                      # centre and size of a country's largest outline, as shares of the map
    $c = $mapInfo.shapes | Where-Object { $_.name -eq $country } | Select-Object -First 1; if (-not $c) { throw "no such country in the map: $country" }
    $best = $null
    foreach ($ring in $c.rings) {
        $xs = $ring | ForEach-Object { [double]$_[0] }; $ys = $ring | ForEach-Object { [double]$_[1] }
        $b = @{ x0 = ($xs | Measure-Object -Minimum).Minimum; x1 = ($xs | Measure-Object -Maximum).Maximum; y0 = ($ys | Measure-Object -Minimum).Minimum; y1 = ($ys | Measure-Object -Maximum).Maximum }
        $b.area = ($b.x1 - $b.x0) * ($b.y1 - $b.y0); if (-not $best -or $b.area -gt $best.area) { $best = $b }
    }
    @{ cx = ($best.x0 + $best.x1) / 2; cy = ($best.y0 + $best.y1) / 2; w = $best.x1 - $best.x0; h = $best.y1 - $best.y0 }
}

try {
    New-Deck $(if ($Layout -in 'Policy', 'Export', 'WhyUs') { $Dark } else { $Light })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count; $crd = Rows $Cards; $secs = 3

    switch ($Layout) {
        'Policy' {
            $slab = Poly @(@(640, -10), @(($W + 10), -10), @(($W + 10), ($H + 10)), @(520, ($H + 10))) $Accent
            $slab.Fill.TwoColorGradient(1, 1); $slab.Fill.ForeColor.RGB = (Rgb (Mix $Accent '#FFFFFF' 0.28)); $slab.Fill.BackColor.RGB = (Rgb $AccentDark); $slab.Fill.GradientAngle = 300
            $big = Tx $crd[0][0] 230 $Dark 602 34 350 $heavyBold $heavy 2 0 0.9
            $lab = Tx $crd[0][1] 30 $Dark 602 ([double]$big.Top + [double]$big.Height - 22) 350 -1 $Font 2
            (Grp @($slab, $big, $lab)).Name = 'Slab'
            $set = @(); if ($Kicker) { $set += (Tx $Kicker 15 $Accent 66 56 500 -1 $Font 1 3) }
            $tt = Tx $Title 88 '#FFFFFF' 62 84 540 $heavyBold $heavy 1 0 1.02; $set += $tt; $yy = [double]$tt.Top + [double]$tt.Height + 10
            $set += (Box 66 $yy 400 4 $Accent); if ($Sub) { $set += (Tx $Sub 32 (Mix $Accent '#FFFFFF' 0.55) 66 ($yy + 20) 480 -1) }
            if ($Note) { $set += (Tx $Note 19 '#FFFFFF' 66 428 480 0 $Font 1 1) }
            (Grp $set).Name = 'Copy'
            if ($crd.Count -gt 1) {                                                                                              # the seal straddles the edge of the slab
                $scx = 590; $scy = 398; $d1 = Disc $scx $scy 60 $Dark; $d1.Line.Visible = -1; $d1.Line.ForeColor.RGB = (Rgb $Accent); $d1.Line.Weight = 3
                $d2 = Disc $scx $scy 51 $Dark; $d2.Line.Visible = -1; $d2.Line.ForeColor.RGB = (Rgb $Accent); $d2.Line.Weight = 1
                $st = Mid (Tx $crd[1][0] 15 '#FFFFFF' ($scx - 50) ($scy - 50) 100 -1 $Font 2 0 1.3) ($scy - 50) 100
                $seal = Grp @($d1, $d2, $st); $seal.Name = 'Seal'; $seal.Rotation = 346
            }
            [void](Box 0 480 $W 44 (Mix $Dark '#FFFFFF' 0.1))
            $words = @($rows | ForEach-Object { $_[0] }); $one = ($words -join '     /     ') + '     /     '
            $strip = New-Text ($one * 4) 17 (Rgb (Mix $Dark '#FFFFFF' 0.85)) $Font; Place $strip 40 (480 + (44 - [double]$strip.Height) / 2); $strip.Name = 'Tick'
            $Duration = 1.2; Steps 2
            $fx = $sl.TimeLine.MainSequence.AddEffect($sl.Shapes('Tick'), 0, 0, $AFTER); $fx.Behaviors.Add(1).MotionEffect.Path = 'M 0 0 L -0.7 0 E'      # the ticker keeps running on the settled cover
            $fx.Timing.Duration = 16; $fx.Timing.SmoothStart = 0; $fx.Timing.SmoothEnd = 0
            $s1 = $pres.Slides(1); $s1.Shapes('Copy').Left = [single]([double]$s1.Shapes('Copy').Left - 640); $s1.Shapes('Slab').Left = [single]([double]$s1.Shapes('Slab').Left + 480)
            if ($crd.Count -gt 1) {
                $g = $s1.Shapes('Seal'); Grow $g 3.6; $g.Rotation = 20
                foreach ($it in $g.GroupItems) { if ($it.Fill.Visible) { $it.Fill.Transparency = 1 }; if ($it.Line.Visible) { $it.Line.Transparency = 1 }; if ($it.HasTextFrame -and $it.TextFrame.HasText) { $it.TextFrame2.TextRange.Font.Fill.Transparency = 1 } }
            }
            Auto $s1 0.4; $secs = 9
        }
        'Numeral' {
            $y0 = 372
            $num = New-Text $Cards 400 (Rgb $Accent) $heavy; if (-not $HeavyFont) { $num.TextFrame.TextRange.Font.Bold = -1 }
            $num.TextFrame.TextRange.Font.Size = [single][math]::Floor(400 * 620 / [double]$num.Width); $num.Name = 'Num'
            Place $num ($W - [double]$num.Width * 0.9) ($y0 - [double]$num.Height * 0.62)
            (Box -5 $y0 ($W + 10) ($H - $y0 + 5) (Mix $Light $Ink 0.06)).Name = 'Ground'; (Box -5 ($y0 - 1.5) ($W + 10) 3 $Ink).Name = 'Hz'
            $vt = VTx $Title 52 $Ink 70 50 300 $heavyBold $heavy; $vt.Name = 'VT'
            if ([double]$vt.Height -gt $y0 - 84) { $vt.TextFrame.TextRange.Font.Size = [single][math]::Floor(52 * ($y0 - 84) / [double]$vt.Height); Place $vt 70 50 }   # the columns must end above the horizon
            if ($Kicker) { (Tx $Kicker 13 $Accent ([double]$vt.Left + [double]$vt.Width + 26) 56 300 -1 $Font 1 3).Name = 'Kick' }
            $low = @(); if ($Sub) { $low += (Tx $Sub 23 $Ink 70 ($y0 + 26) 440 -1 $Font 1 0 1.35) }
            if ($Note) { $low += (Tx $Note 15 $gray 70 488 440 0 $Font 1 1) }
            $yy = $y0 + 30; foreach ($r in $rows) { $low += (Tx $r[0] 17 $body ($W - 65 - 360) $yy 360 0 $Font 3); $yy += 32 }
            if ($low.Count -gt 1) { (Grp $low).Name = 'Low' } elseif ($low.Count) { $low[0].Name = 'Low' }
            $fs = 118; (Poly @(@(($W - $fs), 0), @($W, 0), @($W, $fs)) $Dark).Name = 'FoldA'
            $flap = Poly @(@(($W - $fs), 0), @(($W - $fs), $fs), @($W, $fs)) (Mix $Light $Ink 0.16); $flap.Name = 'FoldB'; Glow $flap $Ink 0.6 14 6
            $Duration = 1.5; Steps 2
            $s1 = $pres.Slides(1); $drop = $H - 34 - $y0
            $s1.Shapes('Num').Top = [single]($H + 10); $g = $s1.Shapes('Ground'); $g.Top = [single]($H - 34); $s1.Shapes('Hz').Top = [single]($H - 35.5)
            if ($low.Count) { $s1.Shapes('Low').Top = [single]([double]$s1.Shapes('Low').Top + $drop) }
            $s1.Shapes('VT').Top = [single](-340)
            foreach ($nm in 'FoldA', 'FoldB') { $f = $s1.Shapes($nm); $f.Width = 3; $f.Height = 3; Place $f ($W - 3) 0 }
            Auto $s1 0.4
        }
        'Shelf' {
            $shY = 436; $yy = 150
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 $yy 280 -1 $Font 1 3); $yy += 26 }
            $tt = Tx $Title 46 $Ink 62 $yy 290 $heavyBold $heavy 1 0 1.08; $tt.TextFrame.WordWrap = 0
            if ([double]$tt.Width -gt 270) { $tt.TextFrame.TextRange.Font.Size = [single][math]::Floor(46 * 270 / [double]$tt.Width) }
            $yy += [double]$tt.Height + 14
            [void](Box 65 $yy 46 4 $Accent); if ($Sub) { [void](Tx $Sub 19 $body 65 ($yy + 22) 270 0 $Font 1 0 1.45) }
            $wds = 100, 84, 106, 90, 100, 86; $hts = 340, 296, 362, 310, 330, 292
            $fills = @($Accent, $Dark, (Mix $Accent $Light 0.6), $AccentDark, (Mix $Dark $Light 0.28), (Mix $Accent $Dark 0.5))
            $x = 372; $prevH = 0; $info = @()
            for ($i = 0; $i -lt $n; $i++) {
                $bw = $wds[$i % 6]; $bh = $hts[$i % 6]; $lean = ($i -eq $n - 1 -and $n -gt 2); if ($lean) { $bh = [math]::Round($prevH * 0.88) }
                $fill = $fills[$i % 6]; $fg = $(if ($i % 6 -eq 2) { $Dark } else { '#FFFFFF' }); $by = $shY - $bh
                $parts = @((Box $x $by $bw $bh $fill 3), (Tx ('{0:00}' -f ($i + 1)) 24 $fg $x ($by + 16) $bw $heavyBold $heavy 2), (Box ($x + $bw * 0.22) ($by + 56) ($bw * 0.56) 2 $fg))
                $v = VTx (($rows[$i][0..([math]::Min(1, $rows[$i].Count - 1))]) -join '//') 23 $fg $x ($by + 76) ($bh - 130) -1
                if ($rows[$i].Count -gt 1) { $p2 = $v.TextFrame.TextRange.Paragraphs(2); $p2.Font.Size = 15; $p2.Font.Bold = 0 }
                $fit = ($bh - 76 - 50) / [double]$v.Height
                if ($fit -lt 1) { $v.TextFrame.TextRange.Paragraphs(1).Font.Size = [single][math]::Floor(23 * $fit); if ($rows[$i].Count -gt 1) { $v.TextFrame.TextRange.Paragraphs(2).Font.Size = [single][math]::Floor(15 * $fit) } }
                Place $v ($x + ($bw - [double]$v.Width) / 2) ($by + 76); $parts += $v
                $parts += (Box $x ($by + $bh - 36) $bw 9 (Mix $fill $(if ($i % 6 -eq 2) { $Dark } else { '#FFFFFF' }) 0.25))
                $g = Grp $parts; $g.Name = "Sp$i"
                if ($lean) {                                                                                                     # the last book leans on the others
                    $th = 9 * [math]::PI / 180; $ccx = $x - 5 + $bw / 2 * [math]::Cos($th) + $bh / 2 * [math]::Sin($th); $ccy = $shY - ($bw / 2 * [math]::Sin($th) + $bh / 2 * [math]::Cos($th))
                    Place $g ($ccx - $bw / 2) ($ccy - $bh / 2); $g.Rotation = 351; $x = $ccx + $bw / 2 * [math]::Cos($th) + $bh / 2 * [math]::Sin($th) - $bw
                }
                $info += , @($bh, $lean); $x += $bw + 5; $prevH = $bh
            }
            [void](Box 350 $shY ($x + 17 - 350) 10 $Dark); [void](Box 350 ($shY + 10) ($x + 17 - 350) 7 (Mix $Light $Ink 0.14))
            if ($Note) { [void](Tx $Note 13 $gray 350 ($shY + 30) 566 0 $Font 1) }
            $Duration = 1.5; Steps 2
            $s1 = $pres.Slides(1)
            for ($i = 0; $i -lt $n; $i++) { $g = $s1.Shapes("Sp$i"); $g.Top = [single](-($info[$i][0] + 70) - 170 * $i); $g.Rotation = $(if ($i % 2) { 9 } else { 351 }) }   # higher toward the right: they land one after another
            Auto $s1 0.4
        }
        'Export' {
            $m = DrawMap 64 92 832 330 (Mix $Dark '#FFFFFF' 0.17) $Dark; $origin = $crd[0][0]
            foreach ($it in $sl.Shapes('Map').GroupItems) { if ($it.Name -eq $origin -or $it.Name -like "$origin [0-9]*") { $it.Fill.ForeColor.RGB = (Rgb $Accent) } }
            Head $true
            $o = Spot $origin; $ax = $m.ML + $m.MW * $o.cx; $ay = $m.MT + $m.MH * $o.cy; $sb = @($Sub -split ';'); $unit = $sb[0]
            $r1 = Disc $ax $ay 20 $Accent; $r1.Fill.Transparency = 0.78; $r2 = Disc $ax $ay 11 $Accent; $r2.Fill.Transparency = 0.5
            if ($sb.Count -gt 1) { [void](Tx $sb[1] 15 $dText 65 440 400 -1) }
            if ($sb.Count -gt 2) { [void](Tx $sb[2] 19 '#FFFFFF' ($W - 65 - 400) 468 400 -1 $Font 3) }
            if ($Note) { [void](Tx $Note 11 (Mix $Dark '#FFFFFF' 0.5) ($W - 65 - 400) 500 400 0 $Font 3) }
            $sum = 0; $prev = $null; $late = @()
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $b = Spot $r[0]; $bx = $m.ML + $m.MW * $b.cx; $by = $m.MT + $m.MH * $b.cy
                $dist = [math]::Sqrt(($bx - $ax) * ($bx - $ax) + ($by - $ay) * ($by - $ay)); $kx = ($ax + $bx) / 2; $ky = [math]::Max(96, ($ay + $by) / 2 - [math]::Max(36, $dist * 0.24))
                $fb = $sl.Shapes.BuildFreeform(0, [single]$ax, [single]$ay)
                for ($q = 1; $q -le 28; $q++) { $u = $q / 28; $v = 1 - $u; $fb.AddNodes(0, 0, [single]($v * $v * $ax + 2 * $v * $u * $kx + $u * $u * $bx), [single]($v * $v * $ay + 2 * $v * $u * $ky + $u * $u * $by)) }
                $arc = Tag ($fb.ConvertToShape()); $arc.Fill.Visible = 0; $arc.Line.ForeColor.RGB = (Rgb $Accent); $arc.Line.Weight = 2.25
                Anim $arc $WIPE $AFTER 0.7 $(if ($i) { 0.15 } else { 0.5 }) (WipeDir ($bx - $ax) ($by - $ay))
                $ring = Disc $bx $by 16 $Accent; $ring.Fill.Transparency = 0.62; Anim $ring $ZOOM $WITH 0.4 0.6
                $dot = Disc $bx $by 5.5 '#FFFFFF'; Anim $dot $FADE $WITH 0.2 0.6
                $text = "$($r[1])  $($r[2])$unit"; $pw = (Wd $text 15) - 12
                if ([math]::Abs($bx - $ax) -lt 90) { $lx = $bx - $pw / 2; $ly = $(if ($by + 54 -gt 428) { $by - 50 } else { $by + 16 }) } else { $lx = $(if ($bx -lt $ax) { $bx - 16 - $pw } else { $bx + 16 }); $ly = $by - 16 }   # beside the dot, away from the origin
                $lx = [math]::Max(16, [math]::Min($W - 16 - $pw, $lx))
                $pill = Grp (Pill $lx $ly $pw 32 $Light $text 15 $Ink); Anim $pill $RISE $WITH 0.4 0.65
                $sum += [double]$r[2]; $tot = Tx (('{0:N0}' -f $sum) + $unit) 46 $Accent 63 460 520 $heavyBold $heavy
                Anim $tot $FADE $WITH 0.25 0.7; if ($prev) { FadeOut $prev 0.2 0.6 }; $prev = $tot
            }
            [void](Disc $ax $ay 5 '#FFFFFF'); if ($crd[0].Count -gt 1) { [void](Tx $crd[0][1] 15 '#FFFFFF' ($ax - 80) ($ay - 44) 160 -1 $Font 2) }
            $secs = [int](4 + 1.0 * $n)
        }
        'KoreaPins' {
            $m = DrawMap 40 22 430 496 (Mix $Accent $Light 0.7) $Light; $bd = $mapInfo.bounds; foreach ($it in $sl.Shapes('Map').GroupItems) { $it.Line.Weight = 0.9 }                               # bounds = lon0, lat0, lon1, lat1 (linear in both)
            $yy = 52; if ($Kicker) { (Tx $Kicker 12.5 $Accent 520 $yy 380 -1 $Font 1 3).Name = 'Kick'; $yy += 26 }
            (Tx $Title 32 $Ink 520 $yy 380 -1 $Font 1 0 1.15).Name = 'Head'
            $shots = ($Images.Count -gt 0); $cy0 = $(if ($shots) { 178 } else { 196 }); $chh = $(if ($shots) { 312 } else { 222 }); $po = $(if ($shots) { 114 } else { 0 })
            $card = Box 520 $cy0 375 $chh '#FFFFFF' 18; Glow $card $Accent 0.8 22 8; $card.Name = 'Card'
            (Tx $Sub 24 $Ink 548 ($cy0 + 66 + $po / 2) 320 -1 $Font 1 0 1.45).Name = 'Info0'
            $pt = @()
            for ($k = 1; $k -le $n; $k++) {
                $r = $rows[$k - 1]; $pt += , @(($m.ML + $m.MW * ([double]$r[1] - $bd[0]) / ($bd[2] - $bd[0])), ($m.MT + $m.MH * ($bd[3] - [double]$r[2]) / ($bd[3] - $bd[1])))
                $set = @((Tx ('STOP {0:00}' -f $k) 13 $Accent 548 ($cy0 + 26 + $po) 320 -1 $Font 1 3), (Tx $r[0] 42 $Ink 546 ($cy0 + 48 + $po) 330 $heavyBold $heavy))
                if ($r.Count -gt 3) { $set += (Tx $r[3] 21 $body 548 ($cy0 + 116 + $po) 320 0 $Font 1 0 1.45) }
                if ($shots -and $k -le $Images.Count) {
                    # the photo is cut to a band through its middle and sits inside the top of the card
                    $shot = Tag ($sl.Shapes.AddPicture($Images[$k - 1], 0, -1, 532, [single]($cy0 + 12), -1, -1)); $shot.LockAspectRatio = -1; $shot.Width = 351
                    $cr = $shot.PictureFormat.Crop; $cr.ShapeTop = [single]([double]$shot.Top + ([double]$shot.Height - 110) / 2); $cr.ShapeHeight = 110; $shot.Top = [single]($cy0 + 12)
                    $shot.AutoShapeType = 5; $shot.Adjustments.Item(1) = 0.1; $set += $shot
                }
                (Grp $set).Name = "Info$k"; (Disc (530 + 28 * ($k - 1)) ($cy0 + $chh + 22) 7 $Accent).Name = "Dot$k"
            }
            for ($k = 2; $k -le $n; $k++) { $s = Stroke $pt[$k - 2][0] $pt[$k - 2][1] $pt[$k - 1][0] $pt[$k - 1][1] $AccentDark 2.25; $s.Line.DashStyle = 4; $s.Name = "Seg$k" }
            for ($k = 1; $k -le $n; $k++) {
                $px = $pt[$k - 1][0]; $py = $pt[$k - 1][1]
                $g = Grp @((Poly @(@(($px - 10), ($py - 20)), @(($px + 10), ($py - 20)), @($px, $py)) $AccentDark), (Disc $px ($py - 27) 14 $AccentDark), (Mid (Tx "$k" 14 '#FFFFFF' ($px - 14) ($py - 41) 28 -1 $Font 2) ($py - 41) 28))
                $g.Name = "Pin$k"; Glow $g $Ink 0.6 8 4
            }
            $Duration = 0.8; Steps ($n + 1)
            for ($s = 0; $s -le $n; $s++) {
                $slide = $pres.Slides($s + 1); $script:sl = $slide
                for ($k = 1; $k -le $n; $k++) {
                    if ($k -gt $s) { $slide.Shapes("Pin$k").Top = [single](-80 - 30 * ($k - $s)) }                                # waiting above the slide
                    if ($k -ge 2) { if ($k -gt $s) { $slide.Shapes("Seg$k").Delete() } elseif ($k -eq $s) { Anim $slide.Shapes("Seg$k") $WIPE $AFTER 0.5 0.1 (WipeDir ($pt[$k - 1][0] - $pt[$k - 2][0]) ($pt[$k - 1][1] - $pt[$k - 2][1])) } }
                    if ($k -gt $s) { $slide.Shapes("Dot$k").Fill.ForeColor.RGB = (Rgb (Mix $Light $Ink 0.16)) }
                }
                for ($k = 0; $k -le $n; $k++) { if ($k -ne $s) { $slide.Shapes("Info$k").Delete() } }
                if ($s -lt $n) { Auto $slide $(if ($s) { 1.9 } else { 0.8 }) }
            }
        }
        'AppIntro' {
            Need1; [void](Disc 690 286 236 (Mix $Accent $Light 0.72)); $rg = Disc 690 286 262 $Light; $rg.Fill.Visible = 0; $rg.Line.Visible = -1; $rg.Line.ForeColor.RGB = (Rgb (Mix $Accent $Light 0.45)); $rg.Line.Weight = 1.25
            $yy = 66; if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 $yy 320 -1 $Font 1 3); $yy += 28 }
            $tt = Tx $Title 38 $Ink 62 $yy 340 $heavyBold $heavy 1 0 1.12; $yy += [double]$tt.Height + 14
            if ($Sub) { $st = Tx $Sub 18 $body 65 $yy 310 0 $Font 1 0 1.45; $yy += [double]$st.Height + 24 }
            $cnt = [math]::Min(3, $Images.Count)
            $spec = $(switch ($cnt) { 1 { , @(585, 62, 410, -16) } 2 { @(470, 96, 360, 24), @(700, 62, 410, -22) } default { @(432, 124, 322, 28), @(588, 58, 414, -12), @(790, 124, 322, -30) } })
            for ($j = 0; $j -lt $cnt; $j++) {
                $q = $spec[$j]; $ph = $sl.Shapes.AddPicture($Images[$j], 0, -1, [single]$q[0], [single]$q[1], -1, -1); $ph.LockAspectRatio = -1; $ph.Height = [single]$q[2]; Place $ph $q[0] $q[1]
                $ph.AutoShapeType = 5; $ph.Adjustments.Item(1) = 0.13; $ph.Line.Visible = -1; $ph.Line.ForeColor.RGB = (Rgb $Ink); $ph.Line.Weight = 6
                $t3 = $ph.ThreeD; $t3.Visible = -1; $t3.SetPresetCamera(15); $t3.FieldOfView = 45; $t3.RotationY = [single]$q[3]; $t3.RotationX = 6; $t3.Depth = 9; $t3.ExtrusionColor.RGB = (Rgb $Ink); $t3.PresetLighting = 24   # a picture with a bezel and a body: one shape, so it tilts as one
                $ph.Name = "Ph$j"
            }
            if ($Note) { [void](Tx $Note 11 $gray 65 508 400) }
            $Duration = 1.3; Steps 2
            for ($j = 0; $j -lt [math]::Min($n, $cnt); $j++) {                                                                   # callouts and feature lines exist on the second slide only
                $q = $spec[$j]; $pwid = [double]$q[2] * 0.46; $text = $rows[$j][0]; $pw = (Wd $text 16) + 18
                $pos = $(switch ($j) { 0 { @(($q[0] - $pw * 0.42), ($q[1] + 46)) } 1 { @(($q[0] + $pwid / 2 - $pw / 2), ($q[1] + $q[2] - 34)) } default { @(($W - 18 - $pw), ($q[1] + 178)) } })
                $bb = Box $pos[0] $pos[1] $pw 40 '#FFFFFF' 20; Glow $bb $Ink 0.72 14 6
                $dd = Disc ($pos[0] + 20) ($pos[1] + 20) 6 $Accent; $tx = Mid (Tx $text 16 $Ink ($pos[0] + 34) $pos[1] ($pw - 44) -1) $pos[1] 40
                Anim (Grp @($bb, $dd, $tx)) $ZOOM $AFTER 0.4 $(if ($j) { 0.25 } else { 0.5 })
                if ($rows[$j].Count -gt 1) {
                    $ln = Grp @((Tx ('{0:00}' -f ($j + 1)) 17 $Accent 65 ($yy + 1) 36 $heavyBold $heavy), (Tx $rows[$j][1] 18 $Ink 104 $yy 270 -1 $Font 1 0 1.3))
                    Anim $ln $FADE $WITH 0.4 $(if ($j) { 0.25 } else { 0.5 }); $yy += [double]$ln.Height + 14
                }
            }
            $s1 = $pres.Slides(1)
            for ($j = 0; $j -lt $cnt; $j++) { $ph = $s1.Shapes("Ph$j"); $ph.ThreeD.RotationY = 0; $ph.ThreeD.RotationX = 0; $ph.Top = [single]($H + 40 + 120 * $j) }
            Auto $s1 0.4; $secs = 6
        }
        'WhyUs' {
            Head $true; $sb = @($Sub -split ';'); $ans = $crd[0]; $m = $crd.Count - 1
            [void](Grp (Pill 65 126 (Wd $sb[0] 14) 30 $dCard $sb[0] 14 $dText))
            $pt = Tx $sb[1] 19 $dText 65 166 330 0 $Font 1 0 1.35; $pb = [double]$pt.Top + [double]$pt.Height
            $cut = Box 65 ($pb + 3) 300 4 $Accent 2; Anim $cut $WIPE $AFTER 0.6 0.7 4                                             # the problem is underlined with a marker
            $set = @(Pill 65 ($pb + 16) (Wd $ans[0] 14) 30 $Accent $ans[0] 14 $Dark); $ah = Tx $ans[1] 24 '#FFFFFF' 65 ($pb + 54) 340 -1 $Font 1 0 1.2; $set += $ah
            if ($ans.Count -gt 2) { $set += (Tx $ans[2] 15 $dText 65 ([double]$ah.Top + [double]$ah.Height + 5) 340 0 $Font 1 0 1.3) }
            Anim (Grp $set) $RISE $AFTER 0.5 0.1
            $sx0 = 450; $sw = 445; $fl = 352; $maxH = 186; $pitch = $sw / $m; $smax = (@($crd[1..$m] | ForEach-Object { [double]$_[1] }) | Measure-Object -Maximum).Maximum
            [void](Box $sx0 $fl $sw 1.5 $dLine)
            for ($q = 0; $q -lt $m; $q++) {
                $r = $crd[$q + 1]; $px = $sx0 + $pitch * ($q + 0.5); $ph = $maxH * [double]$r[1] / $smax; $ours = ($q -eq $m - 1)
                if ($ours) {
                    $cone = Poly @(@(($px - 13), 84), @(($px + 13), 84), @(($px + 70), $fl), @(($px - 70), $fl)) $Accent; $cone.Fill.Transparency = 0.8
                    $pool = Tag ($sl.Shapes.AddShape(9, [single]($px - 76), [single]($fl - 9), 152, 18)); $pool.Line.Visible = 0; $pool.Fill.ForeColor.RGB = (Rgb $Accent); $pool.Fill.Transparency = 0.6
                    $lamp = Poly @(@(($px - 12), 70), @(($px + 12), 70), @(($px + 20), 86), @(($px - 20), 86)) $dLine
                    Anim (Grp @($cone, $pool, $lamp)) $FADE $AFTER 0.6 0.2                                                       # the spotlight comes on
                }
                $bar = Box ($px - 28) ($fl - $ph) 56 $ph $(if ($ours) { $Accent } else { $dLine }); Anim $bar $WIPE $(if ($ours) { $WITH } else { $AFTER }) $(if ($ours) { 0.7 } else { 0.3 }) $(if ($ours) { 0.3 } else { 0 }) 3
                $sc = Tx $r[1] $(if ($ours) { 34 } else { 19 }) $(if ($ours) { '#FFFFFF' } else { $dText }) ($px - 60) ($fl - $ph - $(if ($ours) { 46 } else { 28 })) 120 $heavyBold $heavy 2; Anim $sc $FADE $WITH 0.3 $(if ($ours) { 0.8 } else { 0.1 })
                [void](Tx $r[0] 15 $(if ($ours) { $Accent } else { $dText }) ($px - $pitch / 2) ($fl + 9) $pitch $(if ($ours) { -1 } else { 0 }) $Font 2)
            }
            $cw = 830 / [math]::Max(1, $n); [void](Box 65 408 830 1 $dLine)
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $x = 65 + $cw * $i
                $v = Tx "$($r[0])$($r[1])" 46 $Accent $x 416 ($cw - 16) $heavyBold $heavy; if ($r[1]) { $u = $v.TextFrame.TextRange.Characters($r[0].Length + 1, $r[1].Length); $u.Font.Size = 21; $u.Font.Name = $Font; $u.Font.NameFarEast = $Font; $u.Font.Bold = -1 }
                $set = @($v, (Tx $r[2] 15 $dText $x 476 ($cw - 16) 0 $Font 1 0 1.3)); Anim (Grp $set) $RISE $AFTER 0.45 $(if ($i) { 0 } else { 0.3 })
            }
            Foot $true; $secs = 10
        }
        'Strategy' {
            [void](Poly @(@(-10, -10), @(332, -10), @(264, ($H + 10)), @(-10, ($H + 10))) $Dark)
            if ($Kicker) { $kk = Tx $Kicker 12.5 $soft 48 45 270 -1 $Font 1 3; $kk.TextFrame.WordWrap = 0 }
            $tt = Tx $Title 29 '#FFFFFF' 48 70 250 -1 $Font 1 0 1.18; $yy = [double]$tt.Top + [double]$tt.Height + 34
            for ($i = 0; $i -lt $n; $i++) {
                $set = @((Tx ('{0:00}' -f ($i + 1)) 17 $soft 48 ($yy + 3) 36 $heavyBold $heavy), (Tx $rows[$i][0] 20 '#FFFFFF' 86 $yy 168 -1 $Font 1 0 1.2))
                if ($rows[$i].Count -gt 1) { $set += (Tx $rows[$i][1] 15 $dText 86 ([double]$set[1].Top + [double]$set[1].Height + 5) 160 0 $Font 1 0 1.35) }
                $g = Grp $set; $g.Name = "St$i"; $yy += [double]$g.Height + 22
            }
            $sb = @($Sub -split ';'); $unit = $sb[1]; $m = $crd.Count; $total = (@($crd | ForEach-Object { [double]$_[1] }) | Measure-Object -Sum).Sum
            $bx0 = 372; $bwid = 523; $by = 142; $bht = 58
            [void](Tx $sb[0] 15 $gray $bx0 46 300 -1); [void](Tx (('{0:0.#}' -f $total) + $unit) 46 $Ink ($bx0 - 2) 64 400 $heavyBold $heavy)
            $x = $bx0; $segs = @()
            for ($j = 0; $j -lt $m; $j++) {
                $amt = [double]$crd[$j][1]; $sw = $bwid * $amt / $total; $fill = Mix $Accent $Dark (0.8 * $j / [math]::Max(1, $m - 1)); $top = $by
                (Box ($x + 4) $top ($sw - 8) $bht $fill).Name = "Seg$j"
                (Mid (Tx ('{0:0}%' -f (100 * $amt / $total)) 18 '#FFFFFF' ($x + 4) $top ($sw - 8) -1 $Font 2) $top $bht).Name = "Pct$j"
                $lw = [math]::Max($sw, 96)
                (Grp @((Tx $crd[$j][0] 16 $Ink ($x + $sw / 2 - $lw / 2) ($by + $bht + 22) $lw -1 $Font 2), (Tx "$($crd[$j][1])$unit" 15 $body ($x + $sw / 2 - $lw / 2) ($by + $bht + 46) $lw 0 $Font 2))).Name = "Lab$j"
                $segs += , @($x, $sw); $x += $sw
            }
            [void](Box $bx0 272 $bwid 1 (Mix $Light $Ink 0.18)); if ($sb.Count -gt 2) { [void](Tx $sb[2] 15 $gray $bx0 284 400 -1) }
            $kp = Rows $Kpi; $cell = $bwid / [math]::Max(1, $kp.Count); $cy = 396; $rad = 64
            for ($q = 0; $q -lt $kp.Count; $q++) {
                $cx = $bx0 + $cell * ($q + 0.5); $val = [math]::Min(0.95, [double]$kp[$q][2])
                foreach ($arc in @("Trk$q", (Mix $Light $Ink 0.13), 0), @("Val$q", $Accent, ((180 + 180 * $val) % 360))) {
                    $a = $sl.Shapes.AddShape(20, [single]($cx - $rad), [single]($cy - $rad), [single](2 * $rad), [single](2 * $rad)); $a.Line.Visible = 0; $a.Fill.ForeColor.RGB = (Rgb $arc[1])
                    $a.Adjustments.Item(1) = 180; $a.Adjustments.Item(2) = [single]$arc[2]; $a.Adjustments.Item(3) = 0.2; $a.Name = $arc[0]
                }
                $ghost = Box ($cx - 4) $cy 8 ($rad - 8) $Ink; $ghost.Fill.Transparency = 1                                         # mirror half: keeps the pivot in the middle of the group
                $nd = Grp @((Poly @(@(($cx - 4), $cy), @(($cx + 4), $cy), @($cx, ($cy - $rad + 8))) $Ink), $ghost); $nd.Name = "Ndl$q"; $nd.Rotation = [single]((270 + 180 * $val) % 360)
                [void](Disc $cx $cy 7 $Ink)
                [void](Tx $kp[$q][1] 30 $Ink ($cx - $cell / 2) ($cy + 14) $cell $heavyBold $heavy 2); [void](Tx $kp[$q][0] 15 $body ($cx - $cell / 2 + 6) ($cy + 56) ($cell - 12) 0 $Font 2 0 1.3)
            }
            if ($Note) { [void](Tx $Note 10 $gray ($W - 65 - 300) 516 300 0 $Font 3) }
            $Duration = 1.4; Steps 2
            $s1 = $pres.Slides(1)
            for ($j = 0; $j -lt $m; $j++) {                                                                                      # before: one uncut bar
                $g = $s1.Shapes("Seg$j"); $g.Left = [single]$segs[$j][0]; $g.Width = [single]($segs[$j][1] + 0.5); $g.Top = [single]($by + 18); $g.Height = 22; $g.Fill.ForeColor.RGB = (Rgb $AccentDark)
                $s1.Shapes("Pct$j").Delete(); $s1.Shapes("Lab$j").Delete()
            }
            for ($q = 0; $q -lt $kp.Count; $q++) { $s1.Shapes("Val$q").Adjustments.Item(2) = 183; $s1.Shapes("Ndl$q").Rotation = 270 }
            for ($i = 0; $i -lt $n; $i++) { $g = $s1.Shapes("St$i"); $g.Left = [single]([double]$g.Left - 330 - 60 * $i) }
            Auto $s1 0.5
        }
    }
    Save-Deck $Out $Preview $secs
    "saved $Out"
} catch { Close-Deck; throw }
