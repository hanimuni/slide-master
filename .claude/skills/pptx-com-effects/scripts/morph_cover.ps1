# Covers, contents pages and closing slides that play as one Morph sequence after a single click (2026-10 batch).
# Same colour / font parameters as scene_page.ps1, plus -Accent2 for a two-tone gradient.
#   AuroraType   cover: every letter of the title is a window onto one moving GIF; the letters fly in scattered and assemble
#                -Images flow.gif  -Title 'AI NEXT'  [-Kicker k] [-Sub s] [-Items 'chip|chip|chip'] [-Note n]
#   LiquidTitle  cover: liquid rises inside the outlined title, then a wave floods the slide and becomes the new ground
#                -Title 'POP!'  [-Images cutout.png] [-Kicker k] [-Sub s] [-Note n]
#   Tilt3D       cover: the extruded title swings round from edge-on to facing the room, then settles beside the copy
#                -Title 'FLOW//DESK'  [-Images dark photo] [-Kicker k] [-Sub s] [-Items 'chip|chip'] [-Note n]
#   MosaicFlip   cover: colour tiles turn away in three diagonal waves and leave a photo mosaic around the title plate
#                -Images photo  -Title T  [-Kicker k] [-Sub s] [-Cards 'word|word|word|word' (one big word per tile)]
#   Ribbon       contents: a folded ribbon unrolls row by row over its pale track; each row is one agenda item
#                -Title T  -Items 'head;detail|head;detail|...' (3-6)  [-Kicker k] [-Note n]
#   IsoStack     contents: a flat card turns into an isometric stack, the slabs lift apart and each one lights its chapter
#                -Title T  -Items 'head;line|...' (3-5)  [-Kicker k] [-Images square picture for the top slab]
#   CountFill    closing: the closing word fills up step by step while a counter runs to its target, then the message comes in
#                -Title 'THANK YOU'  -Sub message  -Items 'target;unit;label|contact|contact'  [-Kicker k] [-Note n]
#   FoldCard     closing: a closed leaflet opens to the left and to the right: message, QR code, contacts
#                -Images qr.png  -Title T  -Sub 'cover text'  -Cards 'line|line' (left panel)  -Items 'QR caption;url|label;value|label;value|...'
#   TravelType   closing: the camera travels down a GIF-filled message line by line, pulls back, and a glass card rises
#                -Images flow.gif[,photo]  -Title 'SEE YOU//NEXT//YEAR'  [-Kicker k] [-Sub s] [-Items 'label;value|...']
# All nine advance by themselves (every slide but the last has an "after" time). GIFs: any seamless loop, e.g. make_pattern_gif.py.
# Example: .\morph_cover.ps1 -Layout Ribbon -Title 'Programme' -Items 'Opening;10:00|Session 1;10:40|Panel;13:30' -Out C:\work\r.pptx -Preview C:\work\r.mp4
param(
    [Parameter(Mandatory)][ValidateSet('AuroraType', 'LiquidTitle', 'Tilt3D', 'MosaicFlip', 'Ribbon', 'IsoStack', 'CountFill', 'FoldCard', 'TravelType')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Note = '',
    [string]$Accent = '#7C6BFF',
    [string]$Accent2 = '',                         # second colour of a two-tone gradient (liquid, ribbon); '' = one colour
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
function Steps([int]$count) { for ($i = 1; $i -lt $count; $i++) { Add-MorphSlide $Duration } }
function Auto($slide, $sec) { $slide.SlideShowTransition.AdvanceOnTime = -1; $slide.SlideShowTransition.AdvanceTime = [single]$sec }
function Dur([int]$k, $sec) { $pres.Slides($k).SlideShowTransition.Duration = [single]$sec }
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Place($shape, $x, $y) { $shape.Left = [single]$x; $shape.Top = [single]$y }
function Centre($shape, $x, $y) { Place $shape ($x - [double]$shape.Width / 2) ($y - [double]$shape.Height / 2) }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Need1 { if (-not $Images.Count) { throw "$Layout needs -Images" } }
function Grow($shape, $factor) {                                       # scale about the centre
    $cxs = [double]$shape.Left + [double]$shape.Width / 2; $cys = [double]$shape.Top + [double]$shape.Height / 2
    $shape.LockAspectRatio = -1; $shape.Width = [single]([double]$shape.Width * $factor); Place $shape ($cxs - [double]$shape.Width / 2) ($cys - [double]$shape.Height / 2)
}
function Brk([string]$s) { $s -replace '\s*(\||//)\s*', "`r" }
function Dup($s) { $d = $s.Duplicate(); $d.Left = $s.Left; $d.Top = $s.Top; $d.Name = "Dup$($d.Id)"; $d }
function Fit([string]$text, $size, $maxW, $maxH, [int]$align = 1) {   # heavy text, as large as the box allows ("//" = new line)
    $t = New-Text (Brk $text) $size 0xFFFFFF $heavy; $r = $t.TextFrame.TextRange; if (-not $HeavyFont) { $r.Font.Bold = -1 }
    $r.ParagraphFormat.Alignment = $align; $r.ParagraphFormat.SpaceWithin = 0.9
    for ($pass = 0; $pass -lt 4; $pass++) {                            # measure the laid-out text itself; repeat because the first reading can be stale
        $k = [math]::Min([double]$maxW / [math]::Max([double]$t.Width, [double]$r.BoundWidth), [double]$maxH / [double]$r.BoundHeight); if ($k -ge 1) { break }
        $r.Font.Size = [single][math]::Floor([double]$r.Font.Size * $k * 0.99)
    }
    $t
}
function Paint($s, [string]$a, [string]$b) {                          # one colour, or a top-to-bottom gradient a -> b
    if ($b) { $s.Fill.TwoColorGradient(1, 1); $gs = $s.Fill.GradientStops; $gs.Item(1).Color.RGB = (Rgb $a); $gs.Item(2).Color.RGB = (Rgb $b) } else { $s.Fill.Solid(); $s.Fill.ForeColor.RGB = (Rgb $a) }
    $s.Line.Visible = 0
}
function Wave($x0, $x1, $top, $bottom, $amp, $len) {                  # a block whose top edge is a sine wave
    $x0 = [double]$x0; $x1 = [double]$x1; $top = [double]$top; $bottom = [double]$bottom                              # a bare "-40" argument arrives as text
    $pts = @(); for ($x = $x0; $x -le $x1; $x += 20) { $pts += , @($x, ($top + $amp * [math]::Sin(6.2832 * $x / $len))) }
    $pts += , @($x1, $bottom); $pts += , @($x0, $bottom); Poly $pts $Accent
}
function Lum([string]$hex) { $v = [Convert]::ToInt32($hex.TrimStart('#'), 16); (0.299 * (($v -shr 16) -band 255) + 0.587 * (($v -shr 8) -band 255) + 0.114 * ($v -band 255)) / 255 }
function ChipsC([string]$spec, $y, [string]$fill, [string]$hex) {     # a centred row of pills
    $list = @($spec -split '\|' | Where-Object { $_ }); $total = -10; foreach ($c in $list) { $total += (Wd $c 15) + 10 }
    Chips $spec (($W - $total) / 2) $y $fill $hex
}
function Tone($grp, [string]$hex) { foreach ($it in $grp.GroupItems) { if ($it.HasTextFrame) { $it.TextFrame.TextRange.Font.Color.RGB = (Rgb $hex) } } }
function GifIn([string]$gif, $glyph, $box) {                          # the GIF, cut to the glyph shape; $box = the area the GIF must cover
    $g = $sl.Shapes.AddPicture($gif, 0, -1, 0, 0, -1, -1); $g.LockAspectRatio = -1; $g.Height = [single]($box.H + 40); if ([double]$g.Width -lt $box.W + 40) { $g.Width = [single]($box.W + 40) }
    Place $g ($box.L + ($box.W - [double]$g.Width) / 2) ($box.T + ($box.H - [double]$g.Height) / 2)
    $r = Merge-Shapes 3 @($g.Name, $glyph.Name); if ($r.Type -ne 13) { throw 'intersect did not keep the picture' }; $r
}

try {
    New-Deck $(switch ($Layout) { { $_ -in 'Ribbon', 'IsoStack', 'CountFill' } { $Light } default { $Dark } })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count; $hold = 3
    $two = $(if ($Accent2) { $Accent2 } else { $AccentDark })

    switch ($Layout) {
        'AuroraType' {
            Need1
            $o1 = Disc 250 150 240 $Accent; $o1.Fill.Transparency = 0.74; $o1.SoftEdge.Type = 6; $o1.Name = 'Orb1'
            $o2 = Disc 740 410 260 $two; $o2.Fill.Transparency = 0.78; $o2.SoftEdge.Type = 6; $o2.Name = 'Orb2'
            $t = Fit $Title 170 ($W - 130) 250 2; Set-Center $t; $t.Top = [single](232 - [double]$t.Height / 2)
            $tr = $t.TextFrame.TextRange; $fs = [double]$tr.Font.Size; $tb = Get-Box $t; $cnt = 0
            for ($i = 1; $i -le $tr.Text.Length; $i++) {                                                                        # one glyph shape per letter, each cut out of the same GIF in the same place
                $ch = $tr.Characters($i, 1); if (-not $ch.Text.Trim()) { continue }
                $one = New-Text $ch.Text $fs 0xFFFFFF $heavy; if (-not $HeavyFont) { $one.TextFrame.TextRange.Font.Bold = -1 }; Place $one $ch.BoundLeft $ch.BoundTop
                $gl = New-GlyphShape $one; $one.Delete(); (GifIn $Images[0] $gl $tb).Name = "L$cnt"; $cnt++
            }
            $t.Delete(); $y = $tb.T + $tb.H + 2
            if ($Kicker) { (Tx $Kicker 15 (Mix $two '#FFFFFF' 0.35) 0 78 $W -1 $Font 2 5).Name = 'Kick' }
            $set = @(); if ($Sub) { $st = Tx $Sub 25 '#FFFFFF' 0 $y $W -1 $Font 2 0 1.3; $set += $st; $y += [double]$st.Height + 22 }
            if ($Items) { $set += ChipsC $Items $y $dCard '#FFFFFF' }
            if ($set.Count) { (Grp $set).Name = 'Info' }
            if ($Note) { [void](Tx $Note 11 (Mix $Dark '#FFFFFF' 0.5) 0 506 $W 0 $Font 2) }
            $Duration = 1.7; Steps 3; Dur 3 0.9
            $rnd = New-Object System.Random 7
            foreach ($k in 1, 2) {
                $slide = $pres.Slides($k); if ($set.Count) { $slide.Shapes('Info').Top = [single]($H + 30) }; if ($Kicker) { $slide.Shapes('Kick').Delete() }
                for ($i = 0; $i -lt $cnt; $i++) {
                    $p = $slide.Shapes("L$i"); $p.Top = [single]([double]$p.Top + 62)
                    if ($k -eq 1) {                                                                                             # thrown around the edges, large and turned
                        $ang = $i * 2.39996 + 0.6; $rad = 300 + 130 * $rnd.NextDouble(); Grow $p (2.0 + 1.5 * $rnd.NextDouble())
                        Centre $p ($W / 2 + 1.3 * $rad * [math]::Cos($ang)) ($H / 2 + 0.72 * $rad * [math]::Sin($ang)); $p.Rotation = [single]((330 + 60 * $rnd.NextDouble()) % 360)
                    }
                }
            }
            $s1 = $pres.Slides(1); Place $s1.Shapes('Orb1') 520 220; Place $s1.Shapes('Orb2') -60 -80
            Auto $s1 0.5; Auto $pres.Slides(2) 0.5
        }
        'LiquidTitle' {
            $t = Fit $Title 210 ($W - 150) 300 2; Set-Center $t; $t.Top = [single](248 - [double]$t.Height / 2)
            $edge = New-GlyphShape $t; $nb = Get-Box $edge; $amp = 9; $rnd = New-Object System.Random 5
            $wv = Wave -20 ($W + 500) 0 ($H + 160) $amp 240; Paint $wv $Accent $Accent2; $set = @($wv)
            for ($i = 0; $i -lt 12; $i++) { $set += (Disc ($nb.L + 240 + $nb.W * $rnd.NextDouble()) (34 + $nb.H * 0.85 * $rnd.NextDouble()) (4 + 7 * $rnd.NextDouble()) (Mix $Accent '#FFFFFF' 0.6)) }
            (Grp $set).Name = 'Liquid'
            $mrect = New-Rect -20 -20 ($W + 40) ($H + 40); $d1 = Dup $t; $mask = Merge-Shapes 4 @($mrect.Name, $d1.Name); $mask.Fill.Solid(); $mask.Fill.ForeColor.RGB = (Rgb $Dark); $mask.Line.Visible = 0; $mask.Name = 'Mask'
            $edge.Fill.Visible = 0; $edge.Line.Visible = -1; $edge.Line.ForeColor.RGB = (Rgb $Accent); $edge.Line.Weight = 1.75; $edge.Name = 'Edge'; $edge.ZOrder(0)
            $frect = New-Rect -40 -12 ($W + 80) ($H + 80); $d2 = Dup $t; $fbody = Merge-Shapes 4 @($frect.Name, $d2.Name); Paint $fbody $Accent $Accent2   # the flood with the title punched out: the ground under it shows as letters
            $crest = Wave -40 ($W + 40) -40 -10 $amp 240; (Grp @($fbody, $crest)).Name = 'Flood'                                  # the crest is a separate strip so the holes are cut from a plain rectangle
            $t.Delete()
            if ($Kicker) { (Tx $Kicker 15 $Light 65 46 600 -1 $Font 1 5).Name = 'Kick' }
            if ($Sub) { $st = Tx $Sub 25 $Light 65 0 560 -1 $Font 1 0 1.3; $st.Top = [single]($H - 52 - [double]$st.Height); $st.Name = 'Sub' }
            if ($Note) { (Tx $Note 13 $Light ($W - 65 - 400) 48 400 -1 $Font 3 3).Name = 'Note' }
            if ($Images.Count) { $pr = $sl.Shapes.AddPicture($Images[0], 0, -1, 0, 0, -1, -1); $pr.LockAspectRatio = -1; $pr.Height = 330; Place $pr ($W - 56 - [double]$pr.Width) ($H - 318); $pr.Rotation = 8; $pr.Name = 'Prod'; Glow $pr $Dark 0.55 24 12 }
            $Duration = 2.4; Steps 3; Dur 3 1.4
            $s3 = $pres.Slides(3); $s3.Shapes('Liquid').Top = [single]($H + 40); $s3.Shapes('Edge').Delete()
            foreach ($nm in 'Kick', 'Sub', 'Note') { foreach ($sh in $s3.Shapes) { if ($sh.Name -eq $nm) { $sh.TextFrame.TextRange.Font.Color.RGB = (Rgb $Dark) } } }
            foreach ($k in 1, 2) {
                $slide = $pres.Slides($k); $slide.Shapes('Flood').Top = [single]($H + 30); $lq = $slide.Shapes('Liquid')
                if ($k -eq 1) { Place $lq -20 ($nb.T + $nb.H + 12) } else { Place $lq -260 ($nb.T - 30) }                         # empty, then full and pushed sideways so the crests travel
                if ($Images.Count) { $q = $slide.Shapes('Prod'); $q.Top = [single]($H + 60); $q.Rotation = 350 }
            }
            Auto $pres.Slides(1) 0.5; Auto $pres.Slides(2) 0.5
        }
        'Tilt3D' {
            if ($Images.Count) { (Add-Cover $Images[0]).Name = 'Photo' }
            $t = Fit $Title 150 480 280; $gs =New-GlyphShape $t; $t.Delete(); $gs.Name = 'Tit'; $gs.Line.Visible = 0; $gs.Fill.Solid(); $gs.Fill.ForeColor.RGB = (Rgb $Accent); Place $gs 74 96
            $d3 = $gs.ThreeD; $d3.Visible = -1; $d3.SetPresetCamera(15); $d3.FieldOfView = 45; $d3.RotationY = 14; $d3.RotationX = 6; $d3.Depth = 26; $d3.ExtrusionColor.RGB = (Rgb (Mix $Accent '#000000' 0.5)); $d3.PresetLighting = 24
            $y = 96 + [double]$gs.Height + 26; $set = @((Box 76 $y 210 6 $Accent)); $y += 20
            if ($Sub) { $st = Tx $Sub 23 $Light 76 $y 500 -1 $Font 1 0 1.25; $set += $st; $y += [double]$st.Height + 14 }
            if ($Items) { $set += Chips $Items 76 $y $dCard $Light }
            (Grp $set).Name = 'Copy'
            if ($Kicker) { (Tx $Kicker 14 $Accent 76 52 600 -1 $Font 1 4).Name = 'Kick' }
            if ($Note) { [void](Tx $Note 11 (Mix $Dark '#FFFFFF' 0.5) 76 516 500) }
            $Duration = 1.9; Steps 3; Dur 3 1.1
            foreach ($k in 1, 2) {
                $slide = $pres.Slides($k); $slide.Shapes('Copy').Left = -640; $g2 = $slide.Shapes('Tit'); Grow $g2 1.4; Centre $g2 ($W / 2 - 40) ($H / 2 - 6)
                $q = $g2.ThreeD; if ($k -eq 1) { $q.RotationY = 82; $q.RotationX = 0; $q.Depth = 150 } else { $q.RotationY = 24; $q.RotationX = 10; $q.Depth = 60 }      # edge-on slab, then facing the room
                if ($Images.Count) { Grow $slide.Shapes('Photo') $(if ($k -eq 1) { 1.2 } else { 1.1 }) }
            }
            if ($Kicker) { $pres.Slides(1).Shapes('Kick').Delete() }
            Auto $pres.Slides(1) 0.4; Auto $pres.Slides(2) 0.6
        }
        'MosaicFlip' {
            Need1; (Add-Cover $Images[0]).Name = 'Photo'
            $cols = 6; $rws = 4; $cw = $W / $cols; $chh = $H / $rws; $gap = 5; $set = @()
            for ($c = 0; $c -le $cols; $c++) { $set += (Box ($c * $cw - $gap / 2) 0 $gap $H $Dark) }; for ($r = 0; $r -le $rws; $r++) { $set += (Box 0 ($r * $chh - $gap / 2) $W $gap $Dark) }
            (Grp $set).Name = 'Grid'
            (Box 0 (2 * $chh) (4 * $cw) (2 * $chh) $Dark).Name = 'Plate'
            $pal = @($Accent, (Mix $Accent $Light 0.6), $Light, (Mix $Dark $Accent 0.32), (Mix $Accent '#000000' 0.28), (Mix $Dark '#FFFFFF' 0.16))
            $words = @($Cards -split '\|' | Where-Object { $_ }); $spots = '01', '13', '04', '25', '10', '15'; $tiles = @()
            for ($r = 0; $r -lt $rws; $r++) {
                for ($c = 0; $c -lt $cols; $c++) {
                    if ($r -ge 2 -and $c -lt 4) { continue }
                    $col = $pal[($c * 2 + $r * 3 + $c * $r) % 6]; (Box ($c * $cw + $gap / 2) ($r * $chh + $gap / 2) ($cw - $gap) ($chh - $gap) $col).Name = "T$r$c"
                    $wi = [array]::IndexOf($spots, "$r$c"); $has = ($wi -ge 0 -and $wi -lt $words.Count)
                    if ($has) { (Mid (Tx $words[$wi] 64 $(if ((Lum $col) -gt 0.55) { $Dark } else { $Light }) ($c * $cw) ($r * $chh) $cw $heavyBold $heavy 2) ($r * $chh) $chh).Name = "W$r$c" }
                    $tiles += , @($r, $c, $(if ($r + $c -le 2) { 1 } elseif ($r + $c -le 5) { 2 } else { 3 }), $has)
                }
            }
            $px = 46; $y = 2 * $chh + 38; $set = @()
            if ($Kicker) { $set += (Tx $Kicker 14 $Accent $px $y 540 -1 $Font 1 4); $y += 30 }
            $tt = Tx $Title 64 $Light $px $y 550 $heavyBold $heavy 1 0 1.08; $set += $tt; $y += [double]$tt.Height + 14
            if ($Sub) { $set += (Tx $Sub 20 (Mix $Dark '#FFFFFF' 0.82) $px $y 540 0 $Font 1 0 1.4) }
            (Grp $set).Name = 'Copy'
            $Duration = 0.8; Steps 4
            for ($s = 1; $s -le 4; $s++) {
                $slide = $pres.Slides($s); if ($s -lt 4) { Grow $slide.Shapes('Photo') (1.16 - 0.04 * $s) }
                foreach ($q in $tiles) {
                    if ($q[2] -gt $s - 1) { continue }
                    $tl = $slide.Shapes("T$($q[0])$($q[1])"); $mx = [double]$tl.Left + [double]$tl.Width / 2; $my = [double]$tl.Top + [double]$tl.Height / 2                      # turned edge-on: alternately about the upright and the level axis
                    if (($q[0] + $q[1]) % 2) { $tl.Width = 0.6; $tl.Left = [single]($mx - 0.3) } else { $tl.Height = 0.6; $tl.Top = [single]($my - 0.3) }
                    $tl.Fill.Transparency = 1; if ($q[3]) { $slide.Shapes("W$($q[0])$($q[1])").Delete() }
                }
                if ($s -lt 4) { Auto $slide $(if ($s -eq 1) { 1.0 } else { 0.12 }) }
            }
        }
        'Ribbon' {
            if ($Kicker) { [void](Tx $Kicker 13 $Accent 65 42 700 -1 $Font 1 3) }
            [void](Tx ($Title -replace '\|', ' ') 32 $Ink 65 66 830 $heavyBold $heavy)
            if ($Note) { [void](Tx $Note 11 $gray 65 512 700) }
            $top = 142; $pitch = [math]::Min(86.0, ($H - $top - 40) / $n); $bh = [math]::Round($pitch * 0.72); $xl = 64 + $bh; $xr = $W - 64 - $bh; $geo = @()
            for ($k = 0; $k -lt $n; $k++) {
                $y = $top + $k * $pitch; $ltr = ($k % 2 -eq 0); $more = ($k -lt $n - 1); $prev = ($k -gt 0)
                $rt = $xr + $bh; $rb = $xr + $bh; $lt = $xl - $bh; $lb = $xl - $bh                                               # ends: square, folded down to the next row, or coming out of the fold above
                if ($ltr -and $more) { $rb = $xr } elseif (-not $ltr -and $prev) { $rt = $xr }
                if (-not $ltr -and $more) { $lb = $xl } elseif ($ltr -and $prev) { $lt = $xl }
                $back = $null
                if ($more) { $y2 = $y + $pitch; $back = $(if ($ltr) { @(@(($xr + $bh), $y), @(($xr + $bh), ($y2 + $bh)), @($xr, $y2), @($xr, ($y + $bh))) } else { @(@(($xl - $bh), $y), @($xl, ($y + $bh)), @($xl, $y2), @(($xl - $bh), ($y2 + $bh))) }) }
                $geo += , @{ band = @(@($lt, $y), @($rt, $y), @($rb, ($y + $bh)), @($lb, ($y + $bh))); back = $back; y = $y; ltr = $ltr; col = (Mix $Accent $two ($k / [math]::Max(1, $n - 1))) }
            }
            $pale = Mix $Light $Ink 0.07; foreach ($q in $geo) { [void](Poly $q.band $pale); if ($q.back) { [void](Poly $q.back (Mix $Light $Ink 0.12)) } }             # the track the ribbon will cover
            for ($k = 0; $k -lt $n; $k++) {
                $q = $geo[$k]; (Poly $q.band $q.col).Name = "B$k"; if ($q.back) { (Poly $q.back (Mix $q.col '#000000' 0.34)).Name = "K$k" }
                $set = @((Mid (Tx ('{0:00}' -f ($k + 1)) 23 $Light ($xl + 18) $q.y 50 $heavyBold $heavy) $q.y $bh), (Mid (Tx $rows[$k][0] 21 '#FFFFFF' ($xl + 68) $q.y 440 -1) $q.y $bh))
                if ($rows[$k].Count -gt 1) { $set += (Mid (Tx $rows[$k][1] 16 (Mix $q.col '#FFFFFF' 0.86) ($xr - 18 - 300) $q.y 300 0 $Font 3) $q.y $bh) }
                (Grp $set).Name = "T$k"
            }
            $rr = $bh * 0.62; $dk = Mix $two '#000000' 0.4
            (Grp @((Disc 0 0 $rr $dk), (Disc 0 0 ($rr * 0.66) (Mix $two '#FFFFFF' 0.18)), (Disc 0 0 ($rr * 0.36) $dk), (Disc 0 0 ($rr * 0.12) $Light))).Name = 'Roll'
            $Duration = 0.62; Steps ($n + 1)
            for ($s = 1; $s -le $n + 1; $s++) {
                $slide = $pres.Slides($s); $done = $s - 1                                                                         # rows 0 .. done-1 are unrolled on this slide
                for ($k = 0; $k -lt $n; $k++) {
                    $q = $geo[$k]
                    if ($k -ge $done) { $b = $slide.Shapes("B$k"); $right = [double]$b.Left + [double]$b.Width; $b.Width = 1.5; if (-not $q.ltr) { $b.Left = [single]($right - 1.5) }; $slide.Shapes("T$k").Delete() }
                    if ($q.back -and $k + 1 -ge $done) { $slide.Shapes("K$k").Height = 1.5 }
                }
                $roll = $slide.Shapes('Roll'); Grow $roll (1 - 0.45 * $done / $n)
                if ($done) { $q = $geo[$done - 1]; Centre $roll $(if ($q.ltr) { $xr + $bh } else { $xl - $bh }) ($q.y + $bh / 2) } else { Centre $roll ($xl - $bh) ($top + $bh / 2) }
                if ($s -le $n) { Auto $slide $(if ($s -eq 1) { 0.5 } else { 0.12 }) }
            }
        }
        'IsoStack' {
            if ($Kicker) { [void](Tx $Kicker 13 (Mix $Accent $Ink 0.35) 65 42 700 -1 $Font 1 3) }
            [void](Tx ($Title -replace '\|', ' ') 32 $Ink 65 66 830 $heavyBold $heavy)
            [void](Box 60 132 380 370 '#FFFFFF' 28); [void](Box 460 132 440 370 '#FFFFFF' 28)
            $side = 132; $mx = 250; $my = 314; $spread = [math]::Min(56.0, 170 / [math]::Max(1, $n - 1)); $rowH = 330 / $n; $vx = $mx + 0.66 * $side + 10
            for ($k = $n - 1; $k -ge 0; $k--) {                                                                                 # bottom slab first, so the top one is drawn last
                $col = $(if ($k -eq 0) { $Accent } else { Mix $Light $Ink (0.06 + 0.07 * $k) }); $cy = $my + ($k - ($n - 1) / 2) * $spread
                $sb = Box ($mx - $side / 2) ($cy - $side / 2) $side $side $col 20; $sb.Name = "S$k"
                if ($k -eq 0 -and $Images.Count) { $sb.Fill.UserPicture($Images[0]) }
                $d3 = $sb.ThreeD; $d3.Visible = -1; $d3.SetPresetCamera(1); $d3.RotationX = 314.7; $d3.RotationY = 324.6; $d3.RotationZ = 60; $d3.Depth = 16; $d3.ExtrusionColor.RGB = (Rgb (Mix $col '#000000' 0.22))
            }
            for ($k = 0; $k -lt $n; $k++) {
                $ly = 150 + $k * $rowH; $cy = $my + ($k - ($n - 1) / 2) * $spread
                $set = @((Tx ('{0:00}' -f ($k + 1)) 27 $(if ($k -eq 0) { Mix $Accent $Ink 0.25 } else { $Ink }) 500 $ly 56 $heavyBold $heavy), (Tx $rows[$k][0] 22 $Ink 560 ($ly + 1) 320 -1))
                if ($rows[$k].Count -gt 1) { $set += (Tx $rows[$k][1] 16 $body 560 ($ly + 34) 320 0 $Font 1 0 1.3) }
                (Grp $set).Name = "Lb$k"
                $ln = $sl.Shapes.AddLine([single]$vx, [single]$cy, [single]490, [single]($ly + 18)); $ln.Line.ForeColor.RGB = (Rgb (Mix $Ink '#FFFFFF' 0.55)); $ln.Line.Weight = 1.25; $ln.Name = "Ln$k"
                (Disc $vx $cy 4 $(if ($k -eq 0) { $Accent } else { $Ink })).Name = "Dt$k"
            }
            $Duration = 1.3; Steps 3; Dur 3 1.1
            foreach ($j in 1, 2) {
                $slide = $pres.Slides($j)
                for ($k = 0; $k -lt $n; $k++) {
                    Tone $slide.Shapes("Lb$k") (Mix $Light $Ink 0.2); $slide.Shapes("Ln$k").Delete(); $slide.Shapes("Dt$k").Delete()
                    $sb = $slide.Shapes("S$k")
                    if ($j -eq 1) { $sb.Top = [single]($my - $side / 2); $q = $sb.ThreeD; $q.RotationX = 359.9; $q.RotationY = 359.9; $q.RotationZ = 0.1 }                # one flat card, seen from the front
                    else { $sb.Top = [single]($my - $side / 2 + ($k - ($n - 1) / 2) * 13.5) }                                    # turned, the slabs still touching
                }
            }
            Auto $pres.Slides(1) 0.5; Auto $pres.Slides(2) 0.3
        }
        'CountFill' {
            $r0 = $rows[0]; $target = [double]$r0[0]; $unit = $(if ($r0.Count -gt 1) { $r0[1] } else { '' }); $steps = 6
            $t = Fit $Title 170 ($W - 150) 210 2; Set-Center $t; $t.Top = [single](172 - [double]$t.Height / 2)
            $edge = New-GlyphShape $t; $nb = Get-Box $edge
            $lq = Wave -260 ($W + 260) 0 ($H + 80) 5 180; Paint $lq $Accent $Accent2; $lq.Name = 'Liquid'
            $mrect = New-Rect -20 -20 ($W + 40) ($H + 40); $d1 = Dup $t; $mask = Merge-Shapes 4 @($mrect.Name, $d1.Name); $mask.Fill.Solid(); $mask.Fill.ForeColor.RGB = (Rgb $Light); $mask.Line.Visible = 0; $mask.Name = 'Mask'
            $edge.Fill.Visible = 0; $edge.Line.Visible = -1; $edge.Line.ForeColor.RGB = (Rgb $Ink); $edge.Line.Weight = 1.75; $edge.Name = 'Edge'; $edge.ZOrder(0); $t.Delete()
            if ($Kicker) { [void](Tx $Kicker 13 $Accent 0 40 $W -1 $Font 2 4) }
            $ry = $nb.T + $nb.H + 50; $bw = $W - 180
            if ($r0.Count -gt 2) { (Mid (Tx $r0[2] 21 $body 90 $ry 430 -1) $ry 72).Name = 'Lab' }
            (Mid (Tx "0$unit" 60 $Ink ($W - 90 - 430) $ry 430 $heavyBold $heavy 3) $ry 72).Name = 'Num'
            (Box 90 ($ry + 88) $bw 5 (Mix $Light $Ink 0.13) 2).Name = 'Track'; (Box 90 ($ry + 88) $bw 5 $Accent 2).Name = 'Bar'
            $set = @(); $y = $ry - 4
            if ($Sub) { $st = Tx $Sub 32 $Ink 0 $y $W $heavyBold $heavy 2 0 1.25; $set += $st; $y += [double]$st.Height + 18 }
            $set += (Box ($W / 2 - 30) $y 60 4 $Accent 2); $y += 24
            $extra = @(); for ($i = 1; $i -lt $n; $i++) { $extra += $rows[$i][0] }
            if ($extra.Count) { $set += (Tx ($extra -join '      ') 17 $body 0 $y $W 0 $Font 2) }
            (Grp $set).Name = 'Msg'; if ($Note) { [void](Tx $Note 11 $gray 0 510 $W 0 $Font 2) }
            $Duration = 0.42; Steps ($steps + 2); Dur ($steps + 2) 0.9
            for ($s = 1; $s -le $steps + 2; $s++) {
                $slide = $pres.Slides($s); $f = [math]::Min(1.0, ($s - 1) / [double]$steps); $lq2 = $slide.Shapes('Liquid')
                Place $lq2 (-260 + 36 * ($s % 2) * 2 - 30 * $s * 0) ($nb.T + $nb.H + 10 - $f * ($nb.H + 26))
                if ($s -le $steps + 1) {
                    $slide.Shapes('Num').TextFrame.TextRange.Text = ('{0:N0}' -f [math]::Round($target * (1 - (1 - $f) * (1 - $f)))) + $unit                         # ease-out: fast at first, settling on the target
                    $slide.Shapes('Bar').Width = [single][math]::Max(1.0, $bw * $f); $slide.Shapes('Msg').Top = [single]($H + 20)
                    Auto $slide $(if ($s -eq 1) { 0.6 } elseif ($s -eq $steps + 1) { 0.9 } else { 0.05 })
                } else { foreach ($nm in 'Lab', 'Num', 'Track', 'Bar') { foreach ($sh in @($slide.Shapes)) { if ($sh.Name -eq $nm) { $sh.Delete() } } } }
            }
        }
        'FoldCard' {
            Need1; $pw = 250; $phh = 300; $px = ($W - 3 * $pw) / 2; $py = 164; $paper = $Light; $shade = Mix $Light $Dark 0.55
            $halo = Disc ($W / 2) ($py + $phh / 2) 330 $Accent; $halo.Fill.Transparency = 0.84; $halo.SoftEdge.Type = 6; $halo.Name = 'Halo'
            if ($Kicker) { [void](Tx $Kicker 13 $Accent 0 40 $W -1 $Font 2 4) }
            [void](Tx ($Title -replace '\|', ' ') 38 $Light 0 66 $W $heavyBold $heavy 2)
            if ($Note) { [void](Tx $Note 11 (Mix $Dark '#FFFFFF' 0.5) 0 508 $W 0 $Font 2) }
            foreach ($j in 0, 1, 2) { $pn = Box ($px + $j * $pw) $py $pw $phh $paper; $pn.Name = "P$j"; Glow $pn '#000000' 0.7 20 10 }
            foreach ($j in 1, 2) { $cr = Box ($px + $j * $pw - 7) $py 14 $phh $Ink; $cr.Fill.TwoColorGradient(2, 1); $gs = $cr.Fill.GradientStops; $gs.Item(1).Color.RGB = (Rgb $Ink); $gs.Item(1).Transparency = 1; $gs.Item(2).Color.RGB = (Rgb $Ink); $gs.Item(2).Transparency = 0.78; $cr.Name = "Cr$j" }
            # left panel: a short message (or a second picture)
            if ($Images.Count -gt 1) { $lp = $sl.Shapes.AddPicture($Images[1], 0, -1, 0, 0, -1, -1); $lp.LockAspectRatio = -1; $lp.Height = [single]($phh - 28); if ([double]$lp.Width -lt $pw - 28) { $lp.Width = [single]($pw - 28) }
                Centre $lp ($px + $pw / 2) ($py + $phh / 2); $c = $lp.PictureFormat.Crop; $c.ShapeLeft = [single]($px + 14); $c.ShapeWidth = [single]($pw - 28); $c.ShapeTop = [single]($py + 14); $c.ShapeHeight = [single]($phh - 28); $lp.Name = 'C0' }
            else { $set = @((Box ($px + 30) ($py + 36) 40 5 $Accent 2)); if ($Cards) { $set += (Tx $Cards 29 $Ink ($px + 30) ($py + 62) ($pw - 50) $heavyBold $heavy 1 0 1.3) }; (Grp $set).Name = 'C0' }
            # centre panel: QR code and its caption
            $set = @((Pic $Images[0] ($px + $pw + ($pw - 160) / 2) ($py + 30) 160))
            if ($n) { $set += (Tx $rows[0][0] 17 $Ink ($px + $pw) ($py + 208) $pw -1 $Font 2); if ($rows[0].Count -gt 1) { $set += (Tx $rows[0][1] 15 (Mix $Accent $Ink 0.45) ($px + $pw) ($py + 240) $pw 0 $Font 2) } }
            (Grp $set).Name = 'C1'
            # right panel: contacts
            $set = @(); $y = $py + 34; $step = [math]::Min(84.0, ($phh - 50) / [math]::Max(1, $n - 1))
            for ($i = 1; $i -lt $n; $i++) { $set += (Tx $rows[$i][0] 12.5 (Mix $Accent $Ink 0.3) ($px + 2 * $pw + 28) $y ($pw - 44) -1 $Font 1 2); if ($rows[$i].Count -gt 1) { $set += (Tx $rows[$i][1] 17.5 $Ink ($px + 2 * $pw + 28) ($y + 21) ($pw - 44) -1 $Font 1 0 1.25) }; $y += $step }
            if ($set.Count) { (Grp $set).Name = 'C2' }
            # the cover: the back of the left flap, lying on the centre panel while the leaflet is closed
            (Box ($px + $pw) $py $pw $phh $Accent).Name = 'Cover'; $set = @((Box ($px + $pw + 22) ($py + 22) ($pw - 44) ($phh - 44) $Accent))
            $set[0].Line.Visible = -1; $set[0].Line.ForeColor.RGB = (Rgb $Dark); $set[0].Line.Weight = 1.25; $set[0].Fill.Visible = 0
            if ($Sub) { $set += (Mid (Tx $Sub 30 $Dark ($px + $pw) $py $pw $heavyBold $heavy 2 0 1.25) $py $phh) }
            (Grp $set).Name = 'CovT'
            $Duration = 1.0; Steps 3
            $s1 = $pres.Slides(1); $s2 = $pres.Slides(2); $s3 = $pres.Slides(3)
            foreach ($slide in $s2, $s3) { $cv = $slide.Shapes('Cover'); $cv.Width = 1.5; $cv.Fill.ForeColor.RGB = (Rgb (Mix $Accent '#000000' 0.4)); $slide.Shapes('CovT').Delete() }   # the flap has turned over: only its edge is left
            foreach ($slide in $s1, $s2) { $p2 = $slide.Shapes('P2'); $p2.Width = 1.5; $p2.Fill.ForeColor.RGB = (Rgb $shade); $slide.Shapes('Cr2').Delete(); if ($n -gt 1) { $slide.Shapes('C2').Delete() } }
            $p0 = $s1.Shapes('P0'); $p0.Width = 1.5; $p0.Left = [single]($px + $pw - 1.5); $p0.Fill.ForeColor.RGB = (Rgb $shade); $s1.Shapes('Cr1').Delete(); $s1.Shapes('C0').Delete(); $s1.Shapes('C1').Delete()
            Grow $s1.Shapes('Halo') 0.55
            Auto $s1 0.9; Auto $s2 0.25
        }
        'TravelType' {
            Need1; $hasPhoto = ($Images.Count -gt 1)
            if ($hasPhoto) { $ph = Add-Cover $Images[1]; $ph.PictureFormat.Brightness = 0.36; $ph.Name = 'Photo' }
            $t = Fit $Title 130 470 410; Place $t 66 (($H - [double]$t.Height) / 2 - 4); $tr = $t.TextFrame.TextRange; $lines = @()
            for ($i = 1; $i -le $tr.Paragraphs().Count; $i++) { $p = $tr.Paragraphs($i); $lines += , @(([double]$p.BoundLeft + [double]$p.BoundWidth / 2), ([double]$p.BoundTop + [double]$p.BoundHeight / 2), [double]$p.BoundWidth, [double]$p.BoundHeight) }
            $gl = New-GlyphShape $t; $mb = Get-Box $gl; $t.Delete(); $msg = GifIn $Images[0] $gl $mb; $msg.Name = 'Msg'; $mb = Get-Box $msg
            $gx = 566; $gy = 78; $gw = 330; $gh = $H - 156
            if ($hasPhoto) {                                                                                                   # frosted glass: the same photo, blurred, seen through a smaller window
                $g = Add-Cover $Images[1]; $g.AutoShapeType = 5; $g.Adjustments.Item(1) = 0.08
                $c = $g.PictureFormat.Crop; $c.ShapeLeft = [single]$gx; $c.ShapeWidth = [single]$gw; $c.ShapeTop = [single]$gy; $c.ShapeHeight = [single]$gh
                $e = $g.Fill.PictureEffects.Insert(2); $e.EffectParameters(1).Value = 70; $g.PictureFormat.Brightness = 0.3
                $g.Line.Visible = -1; $g.Line.ForeColor.RGB = 0xFFFFFF; $g.Line.Transparency = 0.5; $g.Line.Weight = 1.25; $g.Name = 'Glass'
            } else { (Box $gx $gy $gw $gh $dCard 22).Name = 'Glass' }
            $set = @(); $y = $gy + 38; $ix = $gx + 34; $iw = $gw - 60
            if ($Kicker) { $set += (Tx $Kicker 12.5 (Mix $Accent '#FFFFFF' 0.5) $ix $y $iw -1 $Font 1 3); $y += 30 }
            if ($Sub) { $st = Tx $Sub 28 '#FFFFFF' $ix $y $iw $heavyBold $heavy 1 0 1.25; $set += $st; $y += [double]$st.Height + 16 }
            $set += (Box $ix $y 44 4 $Accent 2); $y += 26
            foreach ($r in $rows) { $set += (Tx $r[0] 12.5 (Mix $Accent '#FFFFFF' 0.55) $ix $y $iw -1 $Font 1 2); if ($r.Count -gt 1) { $set += (Tx $r[1] 18 '#FFFFFF' $ix ($y + 20) $iw -1) }; $y += 62 }
            (Grp $set).Name = 'Info'
            $nl = $lines.Count; $Duration = 0.85; Steps ($nl + 1); Dur ($nl + 1) 1.3
            for ($s = 1; $s -le $nl; $s++) {
                $slide = $pres.Slides($s); $q = $lines[$s - 1]; $z = [math]::Min(4.2, [math]::Min(($W - 130) / $q[2], ($H - 150) / $q[3]))
                $m2 = $slide.Shapes('Msg'); $m2.LockAspectRatio = -1; $m2.Width = [single]($mb.W * $z); Place $m2 ($W / 2 - ($q[0] - $mb.L) * $z) ($H / 2 - ($q[1] - $mb.T) * $z)   # this line fills the slide
                $slide.Shapes('Info').Top = [single]($H + 40)
                if ($hasPhoto) { $c = $slide.Shapes('Glass').PictureFormat.Crop; $c.ShapeHeight = 2; $c.ShapeTop = [single]($H - 2); if ($s -lt $nl) { Grow $slide.Shapes('Photo') (1 + 0.12 * ($nl - $s)) } }
                else { $slide.Shapes('Glass').Top = [single]($H + 10) }
                Auto $slide $(if ($s -eq 1) { 0.6 } else { 0.3 })
            }
        }
    }
    Save-Deck $Out $Preview $hold
    "saved $Out"
} catch { Close-Deck; throw }
