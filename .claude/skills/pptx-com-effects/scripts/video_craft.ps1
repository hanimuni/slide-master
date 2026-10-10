# More scenes rebuilt from tutorial videos (2026-10 batch, part 3). Same colour / font parameters as scene_page.ps1.
#   SplitTitle  cover: the title is solid on the photo and punched out where it runs onto a slanted panel; the photo drifts closer
#               -Images photo  -Title 'YOUR STORY'  [-Kicker k] [-Sub s]
#   GlassRise   a frosted-glass panel with the text rises from the bottom edge of the photo   -Images photo  -Title T  -Items 'line|line'
#   Loupe       round pieces of the photo grow out of their own spot, one by one, with a label   -Images photo  -Title T  -Items 'fx;fy;label|...' (2-4, fx/fy = 0..1)
#   Lotus       contents: petals unfold from one petal and the numbered items come in around them   [-Images photo]  -Title T  -Items 'head|head|...' (3-7)
#   Bars3D      columns with two shaded sides and a roof, the largest one in the accent colour   -Title T  -Items 'label;value;unit|...' (2-6)
#   FlightPath  a plane flies an arc across the world map, then the title card grows at the destination   -Title T [-Sub s]  -Items 'From country;To country'
#   RegionZoom  the world map zooms into one country per click; only that country keeps its colour    -Title T  -Items 'Country;head;line;line|...'
#   Interlock   slanted photo bands and the text slide in from opposite sides and lock together   -Images photo  -Title T [-Sub s]  -Items 'line|line|...'
#   Ladder      full-height colour panels climb over the photo cover like a ladder   -Images photo  -Title T [-Sub s]  -Items 'head;line;line|...' (3-4)
# FlightPath and RegionZoom draw the map from assets\world_110m.json (Natural Earth, public domain; country names as in that file).
# SplitTitle, GlassRise, Loupe, Lotus, FlightPath and Interlock play by themselves; RegionZoom and Ladder wait for a click.
# Example: .\video_craft.ps1 -Layout Bars3D -Title 'Share by brand' -Items 'K;47;%|A;38;%|M;15;%' -Out C:\work\bars.pptx -Preview C:\work\bars.mp4
param(
    [Parameter(Mandatory)][ValidateSet('SplitTitle', 'GlassRise', 'Loupe', 'Lotus', 'Bars3D', 'FlightPath', 'RegionZoom', 'Interlock', 'Ladder')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Map = '',                             # outline file for FlightPath / RegionZoom (default: assets\world_110m.json)
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
if (-not $Map) { $Map = Join-Path (Split-Path $PSScriptRoot) 'assets\world_110m.json' }
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent
. (Join-Path $PSScriptRoot 'scene_helpers.ps1')
function Steps([int]$count) { for ($i = 1; $i -lt $count; $i++) { Add-MorphSlide $Duration } }
function Auto($slide, $sec) { $slide.SlideShowTransition.AdvanceOnTime = -1; $slide.SlideShowTransition.AdvanceTime = [single]$sec }
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Place($shape, $x, $y) { $shape.Left = [single]$x; $shape.Top = [single]$y }
function Tail($r, [int]$from = 1) { if ($r.Count -gt $from) { ($r[$from..($r.Count - 1)]) -join "`r" } else { '' } }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Need1 { if (-not $Images.Count) { throw "$Layout needs -Images <photo>" } }
function Grow($shape, $factor) {                                       # scale about the centre
    $cxs = [double]$shape.Left + [double]$shape.Width / 2; $cys = [double]$shape.Top + [double]$shape.Height / 2
    $shape.LockAspectRatio = -1; $shape.Width = [single]([double]$shape.Width * $factor); Place $shape ($cxs - [double]$shape.Width / 2) ($cys - [double]$shape.Height / 2)
}
function DrawMap($x, $y, $bw, $bh, [string]$fill, [string]$edge) {    # one freeform per outline, named after the country, grouped as 'Map'
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
    $sky = Mix $Accent '#FFFFFF' 0.88
    New-Deck $(switch ($Layout) { 'FlightPath' { $sky } { $_ -in 'Bars3D', 'RegionZoom', 'Interlock' } { $Light } default { $Dark } })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count

    switch ($Layout) {
        'SplitTitle' {
            Need1; $photo = Add-Cover $Images[0]; $photo.Name = 'Photo'
            $t = New-Text $Title 150 0 $heavy; if (-not $HeavyFont) { $t.TextFrame.TextRange.Font.Bold = -1 }
            if ([double]$t.Width -gt $W - 110) { $t.TextFrame.TextRange.Font.Size = [single][math]::Floor(150 * ($W - 110) / [double]$t.Width) }
            Set-Center $t; $t.Top = [single]([double]$t.Top - 14); $tr = $t.TextFrame.TextRange
            $cut = $Title.IndexOf(' ')                                                                                          # the panel starts in the gap between the words (or in the middle)
            $xs = $(if ($cut -gt 0) { $c = $tr.Characters($cut + 1, 1); [double]$c.BoundLeft + [double]$c.BoundWidth / 2 } else { $W / 2 })
            $ym = [double]$tr.BoundTop + [double]$tr.BoundHeight / 2; $slant = 0.2; $tl = [double]$tr.BoundLeft; $tt = [double]$tr.BoundTop; $tb = $tt + [double]$tr.BoundHeight
            $panel = Poly @(@(($xs + ($ym + 10) * $slant), -10), @(($W + 10), -10), @(($W + 10), ($H + 10)), @(($xs - ($H + 10 - $ym) * $slant), ($H + 10))) $Light
            $mask = Merge-Shapes 2 @($panel.Name, $t.Name); $mask.Fill.ForeColor.RGB = (Rgb $Light); $mask.Line.Visible = 0; $mask.Name = 'Mask'   # combine: solid outside the panel, a hole inside it
            if ($Kicker) { [void](Tx $Kicker 15 '#FFFFFF' ($tl + 4) ($tt - 18) 400 -1 $Font 1 4) }
            if ($Sub) { [void](Tx $Sub 21 $Ink ($W - 65 - 330) ($tb + 22) 330 -1 $Font 3 0 1.3) }
            if ($Note) { [void](Tx $Note 13 $AccentDark ($W - 65 - 330) 52 330 -1 $Font 3 2) }
            $Duration = 6; Steps 2; Grow $pres.Slides(2).Shapes('Photo') 1.14; Auto $pres.Slides(1) 0.3
        }
        'GlassRise' {
            Need1; $photo = Add-Cover $Images[0]; $photo.PictureFormat.Brightness = 0.44; $photo.Name = 'Photo'
            $gx = 66; $gy = 62; $gw = 430; $gh = $H - 124
            $g = Add-Cover $Images[0]; $g.AutoShapeType = 5; $g.Adjustments.Item(1) = 0.06
            $c = $g.PictureFormat.Crop; $c.ShapeLeft = [single]$gx; $c.ShapeWidth = [single]$gw; $c.ShapeTop = [single]$gy; $c.ShapeHeight = [single]$gh    # same picture, same place: only the window is smaller
            $e = $g.Fill.PictureEffects.Insert(2); $e.EffectParameters(1).Value = 70; $g.PictureFormat.Brightness = 0.34
            $g.Line.Visible = -1; $g.Line.ForeColor.RGB = 0xFFFFFF; $g.Line.Transparency = 0.45; $g.Line.Weight = 1.25; $g.Name = 'Glass'
            $set = @(); $y = $gy + 44
            if ($Kicker) { $set += (Tx $Kicker 12.5 '#FFFFFF' ($gx + 40) $y ($gw - 80) -1 $Font 1 3); $y += 30 }
            $tt = Tx $Title 42 '#FFFFFF' ($gx + 40) $y ($gw - 80) -1 $Font 1 0 1.15; $set += $tt; $y += [double]$tt.Height + 14
            $set += (Box ($gx + 40) $y 46 4 '#FFFFFF' 2); $y += 26
            foreach ($r in $rows) { $b = Tx $r[0] 19 '#FFFFFF' ($gx + 40) $y ($gw - 80) 0 $Font 1 0 1.45; $set += $b; $y += [double]$b.Height + 14 }
            (Grp $set).Name = 'Copy'
            $Duration = 1.1; Steps 2
            $s1 = $pres.Slides(1); $c = $s1.Shapes('Glass').PictureFormat.Crop; $c.ShapeHeight = 4; $c.ShapeTop = [single]($H - 5); $s1.Shapes('Copy').Top = [single]($H + 30)
            Grow $pres.Slides(2).Shapes('Photo') 1.06; Auto $s1 0.7
        }
        'Loupe' {
            Need1; (Add-Cover $Images[0]).Name = 'Base'
            $scrim = Box 0 0 $W $H $Dark; $scrim.Fill.Transparency = 1; $scrim.Name = 'Scrim'
            $r0 = 56; $big = 1.95
            for ($i = 0; $i -lt $n; $i++) {
                $cx = [double]$rows[$i][0] * $W; $cy = [double]$rows[$i][1] * $H
                $copy = Add-Cover $Images[0]; $ring = $sl.Shapes.AddShape(9, [single]($cx - $r0), [single]($cy - $r0), [single](2 * $r0), [single](2 * $r0))
                $lens = Merge-Shapes 3 @($copy.Name, $ring.Name); $lens.Name = "Lens$i"; $lens.Line.ForeColor.RGB = 0xFFFFFF; $lens.Line.Weight = 4; $lens.Line.Visible = 0     # picture first: the result stays a picture
                if ($rows[$i].Count -gt 2) { $lab = $rows[$i][2]; $pw = Wd $lab 19; $ly = $(if ($cy + $r0 * $big + 60 -lt $H) { $cy + $r0 * $big + 12 } else { $cy - $r0 * $big - 52 }); (Grp (Pill ($cx - $pw / 2) $ly $pw 40 '#FFFFFF' $lab 19 $Ink)).Name = "Tag$i" }
            }
            $head = @(); if ($Kicker) { $head += (Tx $Kicker 12.5 '#FFFFFF' 65 45 600 -1 $Font 1 3) }; if ($Title) { $head += (Tx $Title 34 '#FFFFFF' 65 68 700 -1) }
            if ($head.Count) { (Grp $head).Name = 'Head' }
            $Duration = 0.8; Steps ($n + 1)
            for ($s = 0; $s -le $n; $s++) {
                $slide = $pres.Slides($s + 1); if ($s) { $slide.Shapes('Scrim').Fill.Transparency = 0.5 }
                for ($i = 0; $i -lt $n; $i++) {
                    $lens = $slide.Shapes("Lens$i"); $has = ($rows[$i].Count -gt 2)
                    if ($i -lt $s) { Grow $lens $big; $lens.Line.Visible = -1; $lens.Line.ForeColor.RGB = 0xFFFFFF; $lens.Line.Weight = 4; Glow $lens '#000000' 0.5 22 8 }
                    else { if ($s) { Grow $lens 0.12 }; if ($has) { $slide.Shapes("Tag$i").Delete() } }                           # waiting: a bright dot on the dimmed photo
                }
                if ($s -lt $n) { Auto $slide $(if ($s) { 0.9 } else { 0.7 }) }
            }
        }
        'Lotus' {
            if ($Images.Count) { $ph = Add-Cover $Images[0]; $ph.PictureFormat.Brightness = 0.33; $ph.Name = 'Photo' }
            $px = $W / 2; $py = 418; $rr = 150; $half = 96; $lh = [math]::Sqrt($rr * $rr - $half * $half); $step = [math]::Min(30, 132 / [math]::Max(1, $n - 1)); $fan = [math]::Min(38, 156 / [math]::Max(1, $n - 1))
            for ($i = 0; $i -lt $n; $i++) {
                $a = $sl.Shapes.AddShape(9, [single]($px - $half - $rr), [single]($py - $lh - $rr), [single](2 * $rr), [single](2 * $rr)); $b = $sl.Shapes.AddShape(9, [single]($px + $half - $rr), [single]($py - $lh - $rr), [single](2 * $rr), [single](2 * $rr))
                $pet = Merge-Shapes 3 @($a.Name, $b.Name); $pet.Line.Visible = 0; $pet.Fill.TwoColorGradient(1, 1); $pet.Fill.GradientAngle = 90   # two circles, intersected: one petal
                $gs = $pet.Fill.GradientStops; $gs.Item(1).Color.RGB = 0xFFFFFF; $gs.Item(1).Transparency = 0.72; $gs.Item(2).Color.RGB = (Rgb (Mix $Accent '#FFFFFF' 0.55)); $gs.Item(2).Transparency = 0.18
                $pet.Name = "Pet$i"
            }
            for ($i = 0; $i -lt $n; $i++) {
                $deg = ($i - ($n - 1) / 2) * $fan; $b = $deg * [math]::PI / 180; $ix = $px + 300 * [math]::Sin($b); $iy = $py - 40 - 262 * [math]::Cos($b)
                $al = $(if ($deg -lt -5) { 3 } elseif ($deg -gt 5) { 1 } else { 2 }); $bx = $(if ($al -eq 3) { $ix - 210 } elseif ($al -eq 1) { $ix } else { $ix - 105 })
                (Grp @((Tx ('{0:00}' -f ($i + 1)) 40 '#FFFFFF' $bx ($iy - 44) 210 $heavyBold $heavy $al), (Tx $rows[$i][0] 18 '#FFFFFF' $bx ($iy + 6) 210 -1 $Font $al 0 1.25))).Name = "It$i"
            }
            $cap = @((Tx $Title 26 '#FFFFFF' 0 ($py + 26) $W -1 $Font 2)); if ($Kicker) { $cap += (Tx $Kicker 12 $soft 0 ($py + 66) $W -1 $Font 2 3) }
            (Grp $cap).Name = 'Cap'
            $Duration = 1.3; Steps 2
            $s2 = $pres.Slides(2); $s1 = $pres.Slides(1)
            for ($i = 0; $i -lt $n; $i++) {
                $deg = ($i - ($n - 1) / 2) * $step; $a = $deg * [math]::PI / 180; $p = $s2.Shapes("Pet$i"); $p.Rotation = [single](($deg + 360) % 360)     # turned about the lower tip
                Place $p ($px + $lh * [math]::Sin($a) - [double]$p.Width / 2) ($py - $lh * [math]::Cos($a) - [double]$p.Height / 2)
                $it = $s1.Shapes("It$i"); $side = ($i - ($n - 1) / 2)
                if ([math]::Abs($side) -lt 0.3) { $it.Top = [single]([double]$it.Top - 320) } else { $it.Left = [single]([double]$it.Left + [math]::Sign($side) * (520 + 90 * [math]::Abs($side))) }
            }
            if ($Images.Count) { Grow $s1.Shapes('Photo') 1.1 }
            Auto $s1 0.6
        }
        'Bars3D' {
            [void](Box 0 ($H * 0.73) $W ($H * 0.27) (Mix $Light $Ink 0.07)); Head
            $vals = @($rows | ForEach-Object { [double]$_[1] }); $vmax = ($vals | Measure-Object -Maximum).Maximum; $hi = [array]::IndexOf($vals, $vmax)
            $base = 446; $maxH = 250; $fw = 46; $sk = 21; $pitch = [math]::Min(200, 780 / $n)
            for ($i = 0; $i -lt $n; $i++) {
                $x = $W / 2 + ($i - ($n - 1) / 2) * $pitch; $yt = $base - $maxH * $vals[$i] / $vmax; $on = ($i -eq $hi)
                $cl = $(if ($on) { $AccentDark } else { Mix $Ink '#FFFFFF' 0.56 }); $cr = $(if ($on) { $Accent } else { Mix $Ink '#FFFFFF' 0.7 }); $ct = $(if ($on) { Mix $Accent '#FFFFFF' 0.55 } else { Mix $Ink '#FFFFFF' 0.86 })
                $sh = Tag ($sl.Shapes.AddShape(9, [single]($x - $fw - 6), [single]($base - $sk - 6), [single](2 * $fw + 44), [single](2 * $sk))); $sh.Line.Visible = 0; $sh.Fill.ForeColor.RGB = (Rgb $Ink); $sh.Fill.Transparency = 0.86; $sh.SoftEdge.Type = 3
                $parts = @((Poly @(@(($x - $fw), ($base - $sk)), @($x, $base), @($x, $yt), @(($x - $fw), ($yt - $sk))) $cl),                # left side, right side, roof
                    (Poly @(@($x, $base), @(($x + $fw), ($base - $sk)), @(($x + $fw), ($yt - $sk)), @($x, $yt)) $cr),
                    (Poly @(@($x, $yt), @(($x + $fw), ($yt - $sk)), @($x, ($yt - 2 * $sk)), @(($x - $fw), ($yt - $sk))) $ct))
                $bar = Grp $parts; Anim $bar $WIPE $AFTER 0.55 0 3
                $unit = $(if ($rows[$i].Count -gt 2) { $rows[$i][2] } else { '' })
                $v = Tx "$($rows[$i][1])$unit" $(if ($on) { 44 } else { 32 }) $(if ($on) { $AccentDark } else { $Ink }) ($x - 90) ($yt - 2 * $sk - $(if ($on) { 62 } else { 48 })) 180 $heavyBold $heavy 2
                Anim $v $FADE $WITH 0.4 0.25
                [void](Tx $rows[$i][0] 18 $(if ($on) { $AccentDark } else { $body }) ($x - 90) ($base + 16) 180 -1 $Font 2)
            }
            Foot
        }
        'FlightPath' {
            $m = DrawMap 50 50 ($W - 100) 400 (Mix $Accent '#FFFFFF' 0.62) $sky
            $a = Spot $rows[0][0]; $b = Spot $rows[0][1]; $ax = $m.ML + $m.MW * $a.cx; $ay = $m.MT + $m.MH * $a.cy; $bx = $m.ML + $m.MW * $b.cx; $by = $m.MT + $m.MH * $b.cy
            foreach ($pin in @($ax, $ay, 'DEPART'), @($bx, $by, 'ARRIVE')) {
                [void](Disc $pin[0] $pin[1] 7 $AccentDark); $d = Disc $pin[0] $pin[1] 13 $AccentDark; $d.Fill.Transparency = 0.7
                [void](Tx $pin[2] 12 $AccentDark ($pin[0] - 60) ($pin[1] + 16) 120 -1 $Font 2 2)
            }
            $lift = [math]::Max(90, [math]::Abs($bx - $ax) * 0.28)
            $c1x = $ax + ($bx - $ax) / 3; $c1y = $ay - $lift; $c2x = $ax + 2 * ($bx - $ax) / 3; $c2y = $by - $lift
            $fb = $sl.Shapes.BuildFreeform(0, [single]$ax, [single]$ay)
            for ($q = 1; $q -le 32; $q++) {                                                                                     # the same cubic curve the plane follows, as short straight pieces
                $u = $q / 32; $v = 1 - $u
                $fb.AddNodes(0, 0, [single]($v * $v * $v * $ax + 3 * $v * $v * $u * $c1x + 3 * $v * $u * $u * $c2x + $u * $u * $u * $bx), [single]($v * $v * $v * $ay + 3 * $v * $v * $u * $c1y + 3 * $v * $u * $u * $c2y + $u * $u * $u * $by))
            }
            $trail = Tag ($fb.ConvertToShape()); $trail.Fill.Visible = 0; $trail.Line.ForeColor.RGB = (Rgb $AccentDark); $trail.Line.Weight = 2.25; $trail.Line.DashStyle = 3; $trail.Name = 'Trail'
            $dir = $(if ($bx -lt $ax) { -1 } else { 1 })                                                                         # the plane is drawn already facing its way: a rotated shape would turn its path too
            $plane = Poly @(@(($ax + 20 * $dir), $ay), @(($ax - 4 * $dir), ($ay - 5)), @(($ax - 12 * $dir), ($ay - 20)), @(($ax - 18 * $dir), ($ay - 20)), @(($ax - 13 * $dir), ($ay - 4)), @(($ax - 22 * $dir), ($ay - 2)), @(($ax - 27 * $dir), ($ay - 9)), @(($ax - 31 * $dir), ($ay - 9)), @(($ax - 28 * $dir), $ay),
                @(($ax - 31 * $dir), ($ay + 9)), @(($ax - 27 * $dir), ($ay + 9)), @(($ax - 22 * $dir), ($ay + 2)), @(($ax - 13 * $dir), ($ay + 4)), @(($ax - 18 * $dir), ($ay + 20)), @(($ax - 12 * $dir), ($ay + 20)), @(($ax - 4 * $dir), ($ay + 5))) $AccentDark
            $plane.Name = 'Plane'; Glow $plane $AccentDark 0.6 8 4
            $secs = 2.6; Anim $trail $WIPE $AFTER $secs 0.4 $(if ($dir -lt 0) { 2 } else { 4 })
            $fx = $sl.TimeLine.MainSequence.AddEffect($plane, 0, 0, $WITH)                                                       # custom path, in shares of the slide
            $path = 'M 0 0 C {0:F4} {1:F4} {2:F4} {3:F4} {4:F4} {5:F4} E' -f (($c1x - $ax) / $W), (($c1y - $ay) / $H), (($c2x - $ax) / $W), (($c2y - $ay) / $H), (($bx - $ax) / $W), (($by - $ay) / $H)
            $fx.Behaviors.Add(1).MotionEffect.Path = $path; $fx.Timing.Duration = [single]$secs; $fx.Timing.TriggerDelayTime = 0.4
            $card = Box ($bx - 20) ($by - 8) 40 16 '#FFFFFF' 8; $card.Fill.Transparency = 1; $card.Name = 'Card'
            if ($Kicker) { [void](Tx $Kicker 12.5 $AccentDark 65 40 600 -1 $Font 1 3) }
            $Duration = 1.0; Steps 2
            $s2 = $pres.Slides(2); $script:sl = $s2
            while ($s2.TimeLine.MainSequence.Count) { $s2.TimeLine.MainSequence.Item(1).Delete() }
            $p2 = $s2.Shapes('Plane'); Place $p2 ([double]$p2.Left + $bx - $ax) ([double]$p2.Top + $by - $ay); $p2.Name = 'Landed'                                      # a different name: it must not morph back from the start
            $cw = 720; $ch = 170; $c2 = $s2.Shapes('Card'); $c2.Width = [single]$cw; $c2.Height = [single]$ch; Place $c2 (($W - $cw) / 2) ($H - $ch - 44); $c2.Adjustments.Item(1) = 0.14; $c2.Fill.Transparency = 0.08; Glow $c2 $AccentDark 0.75 26 10
            $c2.ZOrder(0)
            [void](Tx $Title 44 $AccentDark (($W - $cw) / 2) ($H - $ch - 44 + 30) $cw -1 $Font 2)
            if ($Sub) { [void](Tx $Sub 20 $body (($W - $cw) / 2) ($H - $ch - 44 + 104) $cw 0 $Font 2) }
            Auto $pres.Slides(1) ($secs + 1.0)
        }
        'RegionZoom' {
            $pale = Mix $Light $Ink 0.08
            $m = DrawMap 60 96 ($W - 120) 400 (Mix $Accent '#FFFFFF' 0.3) $Light
            (Tx $Title 30 $Ink 0 44 $W -1 $Font 2).Name = 'Head'; if ($Kicker) { (Tx $Kicker 12.5 $Accent 0 22 $W -1 $Font 2 3).Name = 'Kick' }
            for ($k = 0; $k -lt $n; $k++) {
                $r = $rows[$k]; $x = $(if ($k % 2) { $W - 65 - 400 } else { 65 })
                $set = @((Tx ('REGION {0:00}' -f ($k + 1)) 12.5 $Accent $x 150 400 -1 $Font 1 3), (Tx $(if ($r.Count -gt 1) { $r[1] } else { $r[0] }) 52 $AccentDark $x 176 400 -1 $Font 1 0 1.1))
                if ($r.Count -gt 2) { $set += (Tx (Tail $r 2) 21 $body $x ([double]$set[1].Top + [double]$set[1].Height + 18) 400 0 $Font 1 0 1.5) }
                (Grp $set).Name = "Info$k"
            }
            $Duration = 1.3; Steps ($n + 1)
            for ($s = 1; $s -le $n; $s++) {
                $slide = $pres.Slides($s + 1); $k = $s - 1; $country = $rows[$k][0]; $sp = Spot $country
                $g = $slide.Shapes('Map'); $accentBgr = Rgb $Accent; $paleBgr = Rgb $pale
                foreach ($it in $g.GroupItems) { $it.Fill.ForeColor.RGB = $(if ($it.Name -eq $country -or $it.Name -like "$country [0-9]*") { $accentBgr } else { $paleBgr }) }
                $zoom = [math]::Min(310 / ($sp.h * $m.MH), 360 / ($sp.w * $m.MW)); $tx = $(if ($k % 2) { $W * 0.28 } else { $W * 0.72 })
                $g.LockAspectRatio = -1; $g.Width = [single]($m.W * $zoom)
                Place $g ($tx - (($m.ML - $m.L) + $m.MW * $sp.cx) * $zoom) ($H / 2 + 10 - (($m.MT - $m.T) + $m.MH * $sp.cy) * $zoom)
                $slide.Shapes('Head').Delete(); if ($Kicker) { $slide.Shapes('Kick').Delete() }
            }
            for ($s = 0; $s -le $n; $s++) { for ($k = 0; $k -lt $n; $k++) { if ($k -ne $s - 1) { $pres.Slides($s + 1).Shapes("Info$k").Delete() } } }
        }
        'Interlock' {
            Need1; $deg = 60; $bwid = 138; $pit = 150; $len = 1100; $ccx = 772; $ccy = $H / 2
            $ux = [math]::Cos($deg * [math]::PI / 180); $uy = - [math]::Sin($deg * [math]::PI / 180); $vx = - $uy; $vy = $ux; $stag = 36, -44, 28
            for ($j = 0; $j -lt 3; $j++) {
                $bx = $ccx + $vx * ($j - 1) * $pit + $ux * $stag[$j]; $by = $ccy + $vy * ($j - 1) * $pit + $uy * $stag[$j]
                $pic = Add-Cover $Images[0]; Grow $pic 1.25; $pic.Left = [single]([double]$pic.Left + 150)
                $bar = $sl.Shapes.AddShape(1, [single]($bx - $len / 2), [single]($by - $bwid / 2), [single]$len, [single]$bwid); $bar.Rotation = [single](360 - $deg)
                (Merge-Shapes 3 @($pic.Name, $bar.Name)).Name = "Band$j"                                                         # one photo, cut into three leaning bands
            }
            $set = @(); $y = 70
            if ($Kicker) { $set += (Tx $Kicker 12.5 $Accent 65 $y 400 -1 $Font 1 3); $y += 30 }
            $tt = Tx $Title 56 $Ink 62 $y 430 $heavyBold $heavy 1 0 1.05; $set += $tt; $y += [double]$tt.Height + 8
            if ($Sub) { $st = Tx $Sub 24 $AccentDark 65 $y 400 -1; $set += $st; $y += [double]$st.Height + 22 }
            for ($i = 0; $i -lt $n; $i++) {
                $set += (Tx ('{0:00}' -f ($i + 1)) 17 $Accent 65 ($y + 2) 40 $heavyBold $heavy); $b = Tx $rows[$i][0] 19 $body 104 $y 250 0 $Font 1 0 1.3; $set += $b; $y += [double]$b.Height + 12
            }
            (Grp $set).Name = 'Copy'
            $Duration = 1.1; Steps 2
            $s1 = $pres.Slides(1); $s1.Shapes('Copy').Left = [single](-520)
            for ($j = 0; $j -lt 3; $j++) { $p = $s1.Shapes("Band$j"); $far = (760 + 160 * $j) * $(if ($j % 2) { -1 } else { 1 }); Place $p ([double]$p.Left + $ux * $far) ([double]$p.Top + $uy * $far) }   # each band leaves along its own slope
            Auto $s1 0.5
        }
        'Ladder' {
            Need1; $ph = Add-Cover $Images[0]; $ph.PictureFormat.Brightness = 0.4
            $hero = @((Tx $Title 58 '#FFFFFF' ($W - 65 - 600) 70 600 -1 $Font 3 0 1.1)); if ($Sub) { $hero += (Tx $Sub 24 $soft ($W - 65 - 600) ([double]$hero[0].Top + [double]$hero[0].Height + 12) 600 -1 $Font 3) }
            if ($Kicker) { $hero += (Tx $Kicker 12.5 '#FFFFFF' ($W - 65 - 600) 42 600 -1 $Font 3 3) }
            [void](Grp $hero)
            $cw = $W / $n
            for ($i = 0; $i -lt $n; $i++) {
                $col = Mix (Mix $Accent '#FFFFFF' 0.1) $Dark ($i / [math]::Max(1, $n - 1)); $x = $i * $cw; $mx = $x + $cw / 2
                $set = @((Box $x 0 ($cw + 1) $H $col), (Disc $mx 168 46 '#FFFFFF'))
                $set += (Mid (Tx ('{0:00}' -f ($i + 1)) 30 $col ($mx - 46) 122 92 $heavyBold $heavy 2) 122 92)
                $hd = Tx $rows[$i][0] 27 '#FFFFFF' ($x + 30) 244 ($cw - 60) -1 $Font 2 0 1.2; $set += $hd
                if ($rows[$i].Count -gt 1) { $set += (Tx (Tail $rows[$i]) 18 '#FFFFFF' ($x + 34) ([double]$hd.Top + [double]$hd.Height + 18) ($cw - 68) 0 $Font 2 0 1.5) }
                (Grp $set).Name = "Col$i"
            }
            $Duration = 1.3; Steps 2
            for ($i = 0; $i -lt $n; $i++) { $pres.Slides(1).Shapes("Col$i").Top = [single]($H + 30 + 190 * $i) }                 # further down toward the right: they arrive one after another
        }
    }
    Save-Deck $Out $Preview $(if ($Layout -eq 'Bars3D') { 2.2 + 0.6 * $n } else { 2 })
    "saved $Out"
} catch { Close-Deck; throw }
