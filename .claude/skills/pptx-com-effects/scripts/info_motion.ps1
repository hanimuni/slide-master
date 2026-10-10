# One-slide infographics that move (2026-10 batch). Same colour / font parameters as scene_page.ps1, plus -Accent2 for two-tone layouts.
#   IsoTower       isometric building drawn from freeform faces; floors light up from the ground floor, each with its figure
#                  -Title T [-Sub closing line]  -Items 'floor name;value;line|...' (3-6, top floor first)
#   Liquid         vessels drawn from shapes fill like liquid (two wave surfaces) up to their percentage, read off a side scale
#                  -Title T [-Sub s]  -Items 'bottle|cup|drop|battery|house;percent;label;line|...' (3-5)
#   Trail          a zigzag trail draws itself while a marker travels along it; a round photo and a caption open at every stop
#                  -Title T [-Sub s]  -Images photo,photo,...  -Items 'head;line|...' (3-6, one photo per stop; no photo = numbered disc)
#   Constellation  stars on a night sky connect in sequence, then the figure gets its name
#                  -Title T [-Sub name of the figure]  [-Images sky photo]  -Items 'head;line;value|...' (4-6)  [-Accent2 colour of the last star]
#   Exploded       a tight stack of isometric plates moves aside and pulls apart (Morph); every layer gets its label
#                  -Title T [-Sub s]  [-Images square photo for the bottom plate]  -Items 'head;line|...' (3-5, top layer first)
# All five play by themselves. Example:
#   .\info_motion.ps1 -Layout Liquid -Title 'Recycling' -Items 'bottle;78;PET;collected|cup;46;Cups;reused|drop;31;Rain;reused' -Out C:\work\l.pptx -Preview C:\work\l.mp4
param(
    [Parameter(Mandatory)][ValidateSet('IsoTower', 'Liquid', 'Trail', 'Constellation', 'Exploded')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
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
if (-not $Accent2) { $Accent2 = Mix $Accent '#FFFFFF' 0.5 }
function Steps([int]$count) { for ($i = 1; $i -lt $count; $i++) { Add-MorphSlide $Duration } }
function Auto($slide, $sec) { $slide.SlideShowTransition.AdvanceOnTime = -1; $slide.SlideShowTransition.AdvanceTime = [single]$sec }
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Place($shape, $x, $y) { $shape.Left = [single]$x; $shape.Top = [single]$y }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Trace($pts, [string]$hex, $weight) {                          # open line through @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $s = Tag ($fb.ConvertToShape()); $s.Fill.Visible = 0; $s.Line.ForeColor.RGB = (Rgb $hex); $s.Line.Weight = [single]$weight; $s
}
function Seg($xa, $ya, $xb, $yb, [string]$hex, $weight) { $s = Tag ($sl.Shapes.AddLine([single]$xa, [single]$ya, [single]$xb, [single]$yb)); $s.Line.ForeColor.RGB = (Rgb $hex); $s.Line.Weight = [single]$weight; $s }
function Haze($cx, $cy, $rx, $ry, [string]$hex, $transp) { $s = Tag ($sl.Shapes.AddShape(9, [single]($cx - $rx), [single]($cy - $ry), [single](2 * $rx), [single](2 * $ry))); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s.Fill.Transparency = [single]$transp; $s.SoftEdge.Type = 6; $s }
function Fld($r, [int]$k) { if ($r.Count -gt $k) { $r[$k] } else { '' } }

function IsoFloor($cx, $y, $hw, $fh, [bool]$on) {                      # one storey: left wall, right wall, roof, four windows a side
    $sk = $hw / 2
    if ($on) { $cl = Mix $Accent $Dark 0.62; $cr = Mix $Accent $Dark 0.38; $ct = $Accent; $wl = Mix $Accent $Dark 0.12; $wr = Mix $Accent '#FFFFFF' 0.55 }
    else { $cl = Mix $Dark '#FFFFFF' 0.07; $cr = Mix $Dark '#FFFFFF' 0.12; $ct = Mix $Dark '#FFFFFF' 0.19; $wl = Mix $Dark '#FFFFFF' 0.13; $wr = Mix $Dark '#FFFFFF' 0.2 }
    $parts = @((Poly @(@(($cx - $hw), ($y - $sk)), @($cx, $y), @($cx, ($y - $fh)), @(($cx - $hw), ($y - $fh - $sk))) $cl),
        (Poly @(@($cx, $y), @(($cx + $hw), ($y - $sk)), @(($cx + $hw), ($y - $fh - $sk)), @($cx, ($y - $fh))) $cr),
        (Poly @(@($cx, ($y - $fh)), @(($cx + $hw), ($y - $fh - $sk)), @($cx, ($y - $fh - 2 * $sk)), @(($cx - $hw), ($y - $fh - $sk))) $ct))
    for ($k = 0; $k -lt 4; $k++) {
        $ua = 0.1 + $k * 0.215; $ub = $ua + 0.15; $va = 0.28 * $fh; $vb = 0.74 * $fh
        $parts += (Poly @(@(($cx - $hw + $hw * $ua), ($y - $sk + $sk * $ua - $va)), @(($cx - $hw + $hw * $ub), ($y - $sk + $sk * $ub - $va)), @(($cx - $hw + $hw * $ub), ($y - $sk + $sk * $ub - $vb)), @(($cx - $hw + $hw * $ua), ($y - $sk + $sk * $ua - $vb))) $wl)
        $parts += (Poly @(@(($cx + $hw * $ua), ($y - $sk * $ua - $va)), @(($cx + $hw * $ub), ($y - $sk * $ub - $va)), @(($cx + $hw * $ub), ($y - $sk * $ub - $vb)), @(($cx + $hw * $ua), ($y - $sk * $ua - $vb))) $wr)
    }
    Grp $parts
}
function Vessel([string]$kind, $x, $base, $iw, $ih) {                  # one closed outline standing on ($x, $base)
    $parts = @()
    switch ($kind) {
        'bottle' {
            $parts += (Tag ($sl.Shapes.AddShape(5, [single]($x - $iw * 0.42), [single]($base - $ih * 0.66), [single]($iw * 0.84), [single]($ih * 0.66)))); $parts[-1].Adjustments.Item(1) = 0.16
            $parts += (Poly @(@(($x - $iw * 0.42), ($base - $ih * 0.56)), @(($x - $iw * 0.17), ($base - $ih * 0.84)), @(($x + $iw * 0.17), ($base - $ih * 0.84)), @(($x + $iw * 0.42), ($base - $ih * 0.56))) '#000000')
            $parts += (Tag ($sl.Shapes.AddShape(5, [single]($x - $iw * 0.17), [single]($base - $ih), [single]($iw * 0.34), [single]($ih * 0.3)))); $parts[-1].Adjustments.Item(1) = 0.18
        }
        'cup' {
            $parts += (Poly @(@(($x - $iw * 0.44), ($base - $ih * 0.9)), @(($x + $iw * 0.44), ($base - $ih * 0.9)), @(($x + $iw * 0.3), $base), @(($x - $iw * 0.3), $base)) '#000000')
            $parts += (Tag ($sl.Shapes.AddShape(5, [single]($x - $iw * 0.5), [single]($base - $ih), [single]$iw, [single]($ih * 0.13)))); $parts[-1].Adjustments.Item(1) = 0.4
        }
        'battery' {
            $parts += (Tag ($sl.Shapes.AddShape(5, [single]($x - $iw * 0.38), [single]($base - $ih * 0.91), [single]($iw * 0.76), [single]($ih * 0.91)))); $parts[-1].Adjustments.Item(1) = 0.16
            $parts += (Tag ($sl.Shapes.AddShape(1, [single]($x - $iw * 0.15), [single]($base - $ih), [single]($iw * 0.3), [single]($ih * 0.12))))
        }
        'house' { $parts += (Poly @(@(($x - $iw * 0.46), $base), @(($x - $iw * 0.46), ($base - $ih * 0.56)), @($x, ($base - $ih)), @(($x + $iw * 0.46), ($base - $ih * 0.56)), @(($x + $iw * 0.46), $base)) '#000000') }
        default {                                                       # drop
            $pts = @(); for ($q = 0; $q -lt 40; $q++) { $t = 2 * [math]::PI * $q / 40; $pts += , @(($x + $iw * 0.66 * [math]::Sin($t) * [math]::Pow([math]::Sin($t / 2), 1.6)), ($base - $ih + $ih * (1 - [math]::Cos($t)) / 2)) }
            $parts += (Poly $pts '#000000')
        }
    }
    if ($parts.Count -gt 1) { Tag (Merge-Shapes 1 ([string[]]@($parts | ForEach-Object { $_.Name }))) } else { $parts[0] }
}
function Surface([string]$kind, $x, $base, $iw, $ih, $level, $phase, [string]$hex) {   # the part of the vessel below a wavy surface
    $pts = @(, @(($x - $iw), ($base + 6)))
    for ($q = 0; $q -le 28; $q++) { $pts += , @(($x - $iw + 2 * $iw * $q / 28), ($level + 3.5 * [math]::Sin($q / 28 * 5 * [math]::PI + $phase))) }
    $pts += , @(($x + $iw), ($base + 6))
    $v = Vessel $kind $x $base $iw $ih; $wv = Poly $pts $hex
    $s = Tag (Merge-Shapes 3 @($v.Name, $wv.Name)); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function RoundPic([string]$path, $cx, $cy, $r) {                       # a photo cut to a circle (picture first: the result stays a picture)
    $p = $sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1); $p.LockAspectRatio = -1
    if ([double]$p.Width -ge [double]$p.Height) { $p.Height = [single](2 * $r) } else { $p.Width = [single](2 * $r) }
    Place $p ($cx - [double]$p.Width / 2) ($cy - [double]$p.Height / 2)
    $c = $sl.Shapes.AddShape(9, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r))
    Tag (Merge-Shapes 3 @($p.Name, $c.Name))
}
function Plate($cx, $cy, $side, [string]$hex, $transp, [string]$name) {  # a rounded square seen from above at an angle
    $s = $sl.Shapes.AddShape(5, [single]($cx - $side / 2), [single]($cy - $side / 2), [single]$side, [single]$side); $s.Adjustments.Item(1) = 0.1
    $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s.Fill.Transparency = [single]$transp; $s.ThreeD.SetPresetCamera(23); $s.Name = $name; $s
}

try {
    New-Deck $(if ($Layout -in 'Liquid', 'Trail') { $Light } else { $Dark })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count; $secs = 4

    switch ($Layout) {
        'IsoTower' {
            Head $true
            $hw = 128; $sk = $hw / 2; $gap = 16; $cx = 215; $base = 500
            $pitch = [math]::Min(56, (260 + $gap) / $n); $fh = $pitch - $gap
            [void](Haze $cx ($base - $sk + 6) ($hw + 60) ($sk + 22) $Accent 0.8)
            for ($f = 0; $f -lt $n; $f++) {
                $r = $rows[$n - 1 - $f]; $yf = $base - $f * $pitch; $ly = $yf - $sk - $fh / 2
                [void](IsoFloor $cx $yf $hw $fh $false)
                $lit = IsoFloor $cx $yf $hw $fh $true                                                                          # drawn right above its dark twin, so the next storey still covers its roof
                $lead = Grp @((Seg ($cx + $hw + 8) $ly 436 $ly $Accent 1.25), (Disc 438 $ly 3.5 $Accent))
                $set = @((Mid (Tx (Fld $r 1) 30 $Accent 452 ($ly - 22) 160 $heavyBold $heavy) ($ly - 22) 44), (Tx $r[0] 18 '#FFFFFF' 616 ($ly - 23) 300 -1))
                if ($r.Count -gt 2) { $set += (Tx $r[2] 15 $dText 616 ($ly + 2) 300) }
                $row = Grp $set
                Anim $lit $FADE $AFTER 0.35 $(if ($f) { 0.15 } else { 0.5 }); Anim $lead $WIPE $WITH 0.3 0.2 4; Anim $row $FADE $WITH 0.4 0.3
            }
            if ($Sub) { $pw = Wd $Sub 16; $pl = Grp (Pill 452 464 $pw 34 $Accent $Sub 16 $Dark); Anim $pl $FADE $AFTER 0.5 0.2 }
            Foot $true; $secs = 4 + 0.7 * $n
        }
        'Liquid' {
            Head; if ($Sub) { [void](Tx $Sub 17 $body 65 112 830) }
            $iw = 112; $ih = 196; $base = 372; $pitch = [math]::Min(225, 850 / $n); $pale = Mix $Light $Ink 0.07; $back = Mix $Accent $Light 0.5; $scale = Mix $Light $Ink 0.35
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $kind = $r[0]; $pct = [math]::Max(0, [math]::Min(100, [double]$r[1])); $xc = $W / 2 + ($i - ($n - 1) / 2) * $pitch; $x = $xc - 36; $level = $base - $ih * $pct / 100
                $e = Vessel $kind $x $base $iw $ih; $e.Line.Visible = 0; $e.Fill.ForeColor.RGB = (Rgb $pale)
                $wb = Surface $kind $x $base $iw $ih ($level - 5) 1.6 $back; $wf = Surface $kind $x $base $iw $ih $level 0 $Accent
                $o = Vessel $kind $x $base $iw $ih; $o.Fill.Visible = 0; $o.Line.Visible = -1; $o.Line.ForeColor.RGB = (Rgb $Ink); $o.Line.Weight = 3.5
                foreach ($tk in 0, 25, 50, 75, 100) { $ty = $base - $ih * $tk / 100; [void](Seg ($x - $iw / 2 - $(if ($tk % 50) { 12 } else { 18 })) $ty ($x - $iw / 2 - 6) $ty $scale 1.25) }
                $num = Tx "$($r[1])%" 40 $AccentDark ($x + $iw / 2 + 12) ($level - 26) 110 $heavyBold $heavy; $tr = $num.TextFrame.TextRange; $tr.Characters($tr.Length, 1).Font.Size = 20
                $mark = Grp @((Disc ($x + $iw / 2 + 5) $level 3.5 $AccentDark), $num); $num.Left = [single]($x + $iw / 2 + 16)       # a dot, not a dash: a dash reads as a minus sign
                [void](Tx (Fld $r 2) 19 $Ink ($xc - $pitch / 2 + 8) ($base + 22) ($pitch - 16) -1 $Font 2)
                if ($r.Count -gt 3) { [void](Tx $r[3] 15 $body ($xc - $pitch / 2 + 8) ($base + 52) ($pitch - 16) 0 $Font 2 0 1.3) }
                Anim $wb $WIPE $AFTER 1.0 $(if ($i) { 0.1 } else { 0.5 }) 3; Anim $wf $WIPE $WITH 1.15 $(if ($i) { 0.2 } else { 0.6 }) 3; Anim $mark $FADE $AFTER 0.3
            }
            Foot; $secs = 4 + 1.6 * $n
        }
        'Trail' {
            Head; if ($Sub) { [void](Tx $Sub 17 $AccentDark ($W - 65 - 300) 76 300 -1 $Font 3) }
            $xa = 125; $xb = 835; $yHi = 258; $yLo = 376; $rad = 54; $legs = $n - 1; $dx = ($xb - $xa) / $legs; $pts = @()
            for ($i = 0; $i -lt $n; $i++) { $pts += , @(($xa + $i * $dx), $(if ($i % 2) { $yLo } else { $yHi })) }
            [void](Trace $pts (Mix $Light $Dark 0.13) 22)
            $trail = Trace $pts $Dark 5
            $total = 1.15 * $legs; $start = 0.7
            Anim $trail $WIPE $AFTER $total $start 4
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $px = $pts[$i][0]; $py = $pts[$i][1]; $up = -not ($i % 2); $cy = $(if ($up) { $py - 20 - $rad } else { $py + 20 + $rad })
                [void](Disc $px $py 9 $Light); $dot = Disc $px $py 6.5 $Dark
                if ($i -lt $Images.Count) { $ph = RoundPic $Images[$i] $px $cy $rad; $ph.Line.Visible = -1; $ph.Line.ForeColor.RGB = 0xFFFFFF; $ph.Line.Weight = 4; Glow $ph $Dark 0.7 14 5 }
                else { $ph = Disc $px $cy $rad (Mix $Light $Dark 0.16) }
                $bx = $px + $rad * 0.72; $by = $cy + $(if ($up) { 1 } else { -1 }) * $rad * 0.72
                $stop = Grp @($ph, (Disc $bx $by 15 $Accent), (Mid (Tx "$($i + 1)" 15 '#FFFFFF' ($bx - 15) ($by - 15) 30 -1 $Font 2) ($by - 15) 30))
                $cap = @((Tx $r[0] 18 $Ink ($px - 88) 0 176 -1 $Font 2)); if ($r.Count -gt 1) { $cap += (Tx $r[1] 15 $body ($px - 88) 25 176 0 $Font 2) }
                $cg = Grp $cap; $cg.Top = [single]$(if ($up) { $py + 62 } else { $py - 62 - [double]$cg.Height })
                $at = $start + $total * $i / $legs
                Anim $dot $FADE $WITH 0.2 $at; Anim $stop $FADE $WITH 0.45 ([math]::Max(0.2, $at - 0.1)); Anim $cg $FADE $WITH 0.45 ($at + 0.15)
            }
            $mk = Disc $xa $yHi 11 $Accent; $mk.Line.Visible = -1; $mk.Line.ForeColor.RGB = 0xFFFFFF; $mk.Line.Weight = 3; Glow $mk $AccentDark 0.5 10 3
            $path = 'M 0 0'; for ($i = 1; $i -lt $n; $i++) { $path += ' L {0:F4} {1:F4}' -f (($pts[$i][0] - $xa) / $W), (($pts[$i][1] - $yHi) / $H) }
            $fx = $sl.TimeLine.MainSequence.AddEffect($mk, 0, 0, $WITH); $fx.Behaviors.Add(1).MotionEffect.Path = "$path E"
            $fx.Timing.Duration = [single]$total; $fx.Timing.TriggerDelayTime = [single]$start; $fx.Timing.SmoothStart = 0; $fx.Timing.SmoothEnd = 0
            Foot; $secs = 4 + $total
        }
        'Constellation' {
            if ($Images.Count) { [void](Add-Cover $Images[0]); $sc = Box 0 0 $W $H $Dark; $sc.Fill.Transparency = 0.38 }
            [void](Haze 220 470 330 150 $Accent 0.72); [void](Haze 760 500 330 140 $Accent2 0.78)
            $rnd = New-Object Random 11; $dust = @()
            for ($q = 0; $q -lt 70; $q++) { $d = Disc ($rnd.Next(10, 950)) ($rnd.Next(10, 530)) (0.6 + $rnd.NextDouble() * 1.1) '#FFFFFF'; $d.Fill.Transparency = [single](0.25 + $rnd.NextDouble() * 0.5); $dust += $d }
            [void](Grp $dust); Head $true
            $hiY = 0.44, 0.42, 0.46; $loY = 0.61, 0.64, 0.6; $pts = @()
            for ($i = 0; $i -lt $n; $i++) { $pts += , @((125 + $i * 710 / ($n - 1)), ($H * $(if ($i % 2) { $hiY[[math]::Floor($i / 2) % 3] } else { $loY[($i / 2) % 3] }))) }
            $first = $true
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $px = $pts[$i][0]; $py = $pts[$i][1]; $col = Mix $Accent $Accent2 ($i / ($n - 1)); $bright = Mix $col '#FFFFFF' 0.45; $up = [bool]($i % 2)
                if ($i) {
                    $qx = $pts[$i - 1][0]; $qy = $pts[$i - 1][1]; $len = [math]::Sqrt(($px - $qx) * ($px - $qx) + ($py - $qy) * ($py - $qy)); $ox = ($px - $qx) / $len * 20; $oy = ($py - $qy) / $len * 20
                    $ln = Seg ($qx + $ox) ($qy + $oy) ($px - $ox) ($py - $oy) (Mix $bright '#FFFFFF' 0.3) 1.5; Anim $ln $WIPE $AFTER 0.45 0.05 4
                }
                $st = Tag ($sl.Shapes.AddShape(91, [single]($px - 17), [single]($py - 17), 34, 34)); $st.Line.Visible = 0; $st.Fill.ForeColor.RGB = 0xFFFFFF; $st.Adjustments.Item(1) = 0.16
                $node = Grp @((Haze $px $py 30 30 $col 0.45), $st, (Disc $px $py 5 '#FFFFFF'))
                $set = @((Tx $r[0] 18 '#FFFFFF' ($px - 95) 0 190 -1 $Font 2)); $ty = 26
                if ($r.Count -gt 1 -and $r[1]) { $set += (Tx $r[1] 15 $dText ($px - 95) $ty 190 0 $Font 2) }
                if ($r.Count -gt 2) { $set[0].Top = 40; if ($set.Count -gt 1) { $set[1].Top = 66 }; $set += (Tx $r[2] 32 $bright ($px - 95) 0 190 $heavyBold $heavy 2) }
                $cg = Grp $set; $cg.Top = [single]$(if ($up) { $py - 28 - [double]$cg.Height } else { $py + 28 })
                Anim $node $FADE $(if ($first) { $AFTER } else { $WITH }) 0.4 $(if ($first) { 0.5 } else { 0.35 }); Anim $cg $FADE $WITH 0.45 $(if ($first) { 0.75 } else { 0.55 }); $first = $false
            }
            if ($Sub) { $nm = Tx $Sub 22 '#FFFFFF' 0 476 $W -1 $Font 2 6; Anim $nm $FADE $AFTER 0.8 0.5 }
            Foot $true; $secs = 5 + 1.0 * $n
        }
        'Exploded' {
            $side = 196; $cx = 290; $gapY = [math]::Min(64, 240 / [math]::Max(1, $n - 1)); $top = 308 - $gapY * ($n - 1) / 2; $thick = 9   # a wide gap: every plate must show enough of itself
            for ($i = $n - 1; $i -ge 0; $i--) {                                                                                 # bottom plate first: the upper ones lie over it
                $cy = $top + $i * $gapY; $f = $(if ($n -gt 1) { $i / ($n - 1) } else { 0 }); $col = $(if ($i -eq 0) { $Accent } else { Mix '#FFFFFF' $Dark (0.1 + 0.5 * $f) })
                [void](Plate $cx ($cy + $thick) $side (Mix $col $Dark 0.45) 0 "Under$i")
                if ($i -eq $n - 1 -and $Images.Count) {
                    $p = $sl.Shapes.AddPicture($Images[0], 0, -1, 0, 0, -1, -1); $p.LockAspectRatio = -1
                    if ([double]$p.Width -ge [double]$p.Height) { $p.Height = [single]$side } else { $p.Width = [single]$side }
                    Place $p ($cx - [double]$p.Width / 2) ($cy - [double]$p.Height / 2); $p.AutoShapeType = 5; $p.Adjustments.Item(1) = 0.1
                    $c = $p.PictureFormat.Crop; $c.ShapeLeft = [single]($cx - $side / 2); $c.ShapeWidth = [single]$side; $c.ShapeTop = [single]($cy - $side / 2); $c.ShapeHeight = [single]$side
                    $p.ThreeD.SetPresetCamera(23); $p.Name = "Top$i"
                } else { [void](Plate $cx $cy $side $col $(if ($i -eq 0) { 0.06 } else { 0.24 }) "Top$i") }
            }
            $reach = $side * 0.64
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $cy = $top + $i * $gapY
                (Grp @((Seg ($cx + $reach + 6) $cy 520 $cy $Accent 1.25), (Disc 522 $cy 3.5 $Accent))).Name = "Lead$i"
                $set = @((Mid (Tx ('{0:00}' -f ($i + 1)) 26 $Accent 538 ($cy - 20) 50 $heavyBold $heavy) ($cy - 20) 40), (Tx $r[0] 19 '#FFFFFF' 590 ($cy - 23) 320 -1))
                if ($r.Count -gt 1) { $set += (Tx $r[1] 15 $dText 590 ($cy + 3) 320) }
                (Grp $set).Name = "Info$i"
            }
            if ($Kicker) { (Tx $Kicker 12.5 $Accent 65 45 600 -1 $Font 1 3).Name = 'Kick' }
            (Tx ($Title -replace '\s*(\||//)\s*', ' ') 30 '#FFFFFF' 65 68 830 -1).Name = 'Title'
            Foot $true
            $Duration = 1.4; Steps 2
            $s1 = $pres.Slides(1); $mid = 290; $bigX = 668; $zoom = 1.3
            for ($i = 0; $i -lt $n; $i++) {
                foreach ($nm in "Under$i", "Top$i") {
                    $q = $s1.Shapes($nm); $q.LockAspectRatio = -1; $q.Width = [single]($side * $zoom); $q.Height = [single]($side * $zoom)
                    Place $q ($bigX - $side * $zoom / 2) ($mid + ($i - ($n - 1) / 2) * 15 + $(if ($nm -like 'Under*') { $thick * $zoom } else { 0 }) - $side * $zoom / 2)
                }
                $s1.Shapes("Lead$i").Delete(); $s1.Shapes("Info$i").Left = [single]($W + 40 + 60 * $i)
            }
            $t1 = $s1.Shapes('Title'); $t1.TextFrame.TextRange.Text = ($Title -replace '\s*(\||//)\s*', "`r"); $t1.TextFrame.TextRange.Font.Size = 46; $t1.Width = 380; Place $t1 65 180
            if ($Kicker) { Place $s1.Shapes('Kick') 65 150 }
            if ($Sub) { $script:sl = $s1; (Tx $Sub 20 $dText 65 ([double]$t1.Top + [double]$t1.Height + 18) 380 0 $Font 1 0 1.4).Name = 'Sub1' }
            Auto $s1 1.6; $secs = 4
        }
    }
    Save-Deck $Out $Preview ([int][math]::Ceiling($secs))
    "saved $Out"
} catch { Close-Deck; throw }
