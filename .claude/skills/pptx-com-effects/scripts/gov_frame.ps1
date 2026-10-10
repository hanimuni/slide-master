# Frames for government / public-sector planning documents (vision system, project overview, organisation, roadmap).
# Same colour / font parameters as scene_page.ps1. Every layout plays by itself (entrance animations, no clicks).
#   VisionStairs  vision system on a dark ground: terraces climb to the right, a spark runs up the steps, a beacon lights the vision
#                 [-Cards '2027' = large faint figure in the top right corner]
#                 -Title T  -Sub 'vision//two lines'  -Items 'label;item;item|label;item;item;item|label;item;...' (top step first, 3 steps)
#   VisionTemple  vision system as a temple front that builds itself: base, columns (strategies), beam (goal), roof (vision), sunrise
#                 -Title T  -Sub 'vision'  -Cards 'goal sentence|base sentence'  -Items 'head;line;line|...' (3-5 columns)
#   RibbonFlow    background -> tasks -> effects on one ribbon that makes two U-turns and unrolls along its course
#                 -Title T [-Sub 'closing line']  -Items 'label;head;sub;head;sub;head;sub|...' (3 rows, the middle one runs right to left)
#   Dossier       one-page project overview: slanted dark panel with a huge figure and facts, staggered content rows, a stamp at the end
#                 -Title 'name//two lines'  -Sub 'stamp//text'  -Cards 'label;figure;unit|label;value|label;value|label;value'  -Items 'head;line;line|...' (4)
#   HubSpoke      steering system: committee halo and task-force hub bleed off the left edge, spokes run to the divisions, signals travel out
#                 -Title T  -Sub 'hub name;line;tag'  -Cards 'committee;role|line//line'  -Items 'division;role|...' (4-6)
#   MetroOrg      organisation chart drawn as a subway map: portraits are the interchange stations, a train runs down every line
#                 -Title T  -Sub 'head title;name'  -Cards '3;unit|9;unit|86;unit'  -Items 'bureau;name;team;role;team;role;team;role|...' (2-3)  [-Images root,line1,line2,line3]
#   Gantt         quarterly roadmap: hatched plan bars, milestones, a "today" line that sweeps in and fills the finished part
#                 -Title T  -Sub 'position;label' (position in columns, e.g. 1.6)  -Cards 'year;quarter|year;quarter|...'  -Items 'task;start;end;milestone|...' (3-6)
#   Runway        roadmap as a night runway in perspective: edge lights come on from near to far, markers follow, a light runs to the goal
#                 -Title T  -Sub 'goal'  -Cards 'D-448;label'  -Items 'date;head;line|...' (3-5, nearest first)
# Example: .\gov_frame.ps1 -Layout Gantt -Title 'Roadmap' -Sub '1.5;Today' -Cards '2027;Q1|2027;Q2|2027;Q3|2027;Q4' -Items 'Design;0;1.5;Spec|Build;1;3.5;Beta|Launch;3;4;Open' -Out C:\work\g.pptx -Preview C:\work\g.mp4
param(
    [Parameter(Mandatory)][ValidateSet('VisionStairs', 'VisionTemple', 'RibbonFlow', 'Dossier', 'HubSpoke', 'MetroOrg', 'Gantt', 'Runway')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Note = '',
    [string]$Accent = '#E0B04A',
    [string]$AccentDark = '',
    [string]$Light = '#F6F1E4',
    [string]$Dark = '#0E1E3F',
    [string]$Ink = '#1B2233',
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
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Place($shape, $x, $y) { $shape.Left = [single]$x; $shape.Top = [single]$y }
function Tail($r, [int]$from = 1) { if ($r.Count -gt $from) { ($r[$from..($r.Count - 1)]) -join "`r" } else { '' } }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Ln($pts, [string]$hex, $wt) {                                 # open line through the points
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $s = Tag ($fb.ConvertToShape()); $s.Fill.Visible = 0; $s.Line.Visible = -1; $s.Line.ForeColor.RGB = (Rgb $hex); $s.Line.Weight = [single]$wt; $s
}
function Seg($x1, $y1, $x2, $y2, [string]$hex, $wt) {                  # straight stroke; a level one is a thin box so that a wipe has something to wipe
    if ([math]::Abs($y2 - $y1) -lt 1) { Box ([math]::Min($x1, $x2)) ($y1 - $wt / 2) ([math]::Abs($x2 - $x1)) $wt $hex } else { Ln @(@($x1, $y1), @($x2, $y2)) $hex $wt }
}
function Ring($cx, $cy, $rad, [string]$hex, $wt) { $s = Disc $cx $cy $rad $hex; $s.Fill.Visible = 0; $s.Line.Visible = -1; $s.Line.ForeColor.RGB = (Rgb $hex); $s.Line.Weight = [single]$wt; $s }
function Glide($shape, [string]$path, $dur, [int]$trig = 2, $delay = 0) {   # custom motion path, in shares of the slide
    $fx = $sl.TimeLine.MainSequence.AddEffect($shape, 0, 0, $trig); $fx.Behaviors.Add(1).MotionEffect.Path = $path
    $fx.Timing.Duration = [single]$dur; $fx.Timing.TriggerDelayTime = [single]$delay; $fx
}
function Face([string]$path, $cx, $cy, $rad, [string]$ringHex) {       # round portrait cut from the top of a cut-out, on a tinted disc
    $names = @((Disc $cx $cy $rad (Mix $Light $Accent 0.16)).Name)
    if ($path) {
        $p = $sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1); $p.LockAspectRatio = -1; $p.Width = [single](2 * $rad / 0.62)
        Place $p ($cx - [double]$p.Width / 2) ($cy - $rad + $rad * 0.1)
        $c = $sl.Shapes.AddShape(9, [single]($cx - $rad), [single]($cy - $rad), [single](2 * $rad), [single](2 * $rad))
        $names += (Tag (Merge-Shapes 3 @($p.Name, $c.Name))).Name                                                             # picture first: the result stays a picture
    }
    $names += (Ring $cx $cy ($rad + 2) $ringHex 4).Name
    $names
}

try {
    New-Deck $(if ($Layout -in 'VisionStairs', 'HubSpoke', 'Runway') { $Dark } else { $Light })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count; $secs = 7

    switch ($Layout) {
        'VisionStairs' {
            $sth = 92; $sgap = 6; $bot = 496; $stepX = 205
            $lx = 65 + ($n - 1) * $stepX + 12; $topY = $bot - $n * ($sth + $sgap) + $sgap; $ly = $topY - 58
            $beam = Poly @(@($lx, $ly), @(30, ($ly - 44)), @(30, ($ly + 150))) $Accent                                              # drawn first: it lies under the text
            $beam.Fill.TwoColorGradient(2, 1); $beam.Fill.GradientAngle = 0
            $gs = $beam.Fill.GradientStops; $gs.Item(1).Color.RGB = (Rgb $Accent); $gs.Item(1).Transparency = 1; $gs.Item(2).Color.RGB = (Rgb $Accent); $gs.Item(2).Transparency = 0.45
            if ($Cards) { [void](Tx $Cards 132 (Mix $Dark "#FFFFFF" 0.08) 455 34 440 $heavyBold $heavy 3) }
            Head $true
            $vis = Grp @((Tx 'VISION' 12.5 $Accent 65 ($ly - 14) 300 -1 $Font 1 4), (Tx $Sub 34 '#FFFFFF' 65 ($ly + 8) 410 $heavyBold $heavy 1 0 1.2))
            $slabs = @(); $path = 'M 0 0'; $px = 75; $py = 0
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $sx = 65 + ($n - 1 - $i) * $stepX; $sy = $bot - ($n - $i) * ($sth + $sgap) + $sgap
                $set = @((Box $sx $sy ($W - $sx) $sth (Mix $Dark '#FFFFFF' (0.07 + 0.05 * ($n - 1 - $i)))), (Box $sx $sy ($W - $sx) 4 $Accent))
                $set += (Mid (Tx $r[0] 20 $Accent ($sx + 22) $sy 100 -1 $Font 1 0 1.15) ($sy + 4) ($sth - 4))
                $m = $r.Count - 1; $ix = $sx + 132; $cwid = (895 - $ix) / $m
                for ($j = 1; $j -le $m; $j++) {
                    $set += (Box ($ix + ($j - 1) * $cwid - 13) ($sy + 28) 1 ($sth - 52) $dLine)
                    $set += (Mid (Tx $r[$j] 17 '#FFFFFF' ($ix + ($j - 1) * $cwid) $sy ($cwid - 16) -1 $Font 1 0 1.3) ($sy + 4) ($sth - 4))
                }
                $slabs += (Grp $set)
            }
            $pole = Box ($lx - 1.5) $ly 3 ($topY - $ly) (Mix $Dark '#FFFFFF' 0.45)
            $lamp = Disc $lx $ly 10 $Accent; Glow $lamp $Accent 0.25 22 0
            $sy = $bot - $sth; $spark = Disc 75 ($sy - 8) 6 $Light; Glow $spark $Accent 0.2 12 0                                    # the spark starts on the lowest step
            $px = 75; $py = $sy - 8
            for ($i = $n - 2; $i -ge 0; $i--) {
                $nx = 65 + ($n - 1 - $i) * $stepX - 9; $ny = $bot - ($n - $i) * ($sth + $sgap) + $sgap - 8
                $path += ' L {0:F4} {1:F4} L {0:F4} {2:F4}' -f (($nx - 75) / $W), (($py - ($sy - 8)) / $H), (($ny - ($sy - 8)) / $H); $px = $nx; $py = $ny
            }
            $path += ' L {0:F4} {1:F4} L {0:F4} {2:F4} E' -f (($lx - 75) / $W), (($py - ($sy - 8)) / $H), (($ly - ($sy - 8)) / $H)
            for ($i = $n - 1; $i -ge 0; $i--) { Anim $slabs[$i] $WIPE $AFTER 0.5 0 4 }
            Anim $pole $FADE $WITH 0.3 0.2
            Anim $spark $FADE $AFTER 0.2; [void](Glide $spark $path 1.8 $AFTER)
            Anim $lamp $FADE $AFTER 0.25; Anim $beam $WIPE $AFTER 0.9 0 2; Anim $vis $FADE $WITH 0.7 0.35
            Foot $true; $secs = 9
        }
        'VisionTemple' {
            $cxm = $W / 2; $stone = Mix $Light $Dark 0.08; $flute = Mix $Light $Dark 0.2; $card = Rows $Cards
            $sun = Tag ($sl.Shapes.AddShape(142, [single]($cxm - 215), [single](208 - 215), 430, 430)); $sun.Adjustments.Item(1) = 180; $sun.Adjustments.Item(2) = 0
            $sun.Line.Visible = 0; $sun.Fill.ForeColor.RGB = (Rgb (Mix $Light $Accent 0.14))
            Head
            $roof = Grp @((Poly @(@($cxm, 96), @(($cxm + 404), 208), @(($cxm - 404), 208)) $Dark), (Box ($cxm - 404) 208 808 7 $Accent),
                (Tx 'VISION' 12.5 $soft 0 131 $W -1 $Font 2 4), (Tx $Sub 25 '#FFFFFF' 0 153 $W $heavyBold $heavy 2))
            $beamSet = @((Box ($cxm - 376) 222 752 44 $Accent)); if ($card.Count) { $beamSet += (Mid (Tx $card[0][0] 18 '#FFFFFF' ($cxm - 366) 222 732 -1 $Font 2) 222 44) }
            $beam = Grp $beamSet
            $pitch = [math]::Min(190, 600 / [math]::Max(1, $n - 1)); $cols = @(); $roman = 'I', 'II', 'III', 'IV', 'V', 'VI'
            for ($i = 0; $i -lt $n; $i++) {
                $cx = $cxm + ($i - ($n - 1) / 2) * $pitch; $r = $rows[$i]
                $set = @((Box ($cx - 78) 285 156 148 $stone), (Box ($cx - 86) 274 172 11 $AccentDark), (Box ($cx - 86) 433 172 11 $AccentDark))
                foreach ($o in 7, 13) { $set += (Box ($cx - 78 + $o) 285 1.2 148 $flute); $set += (Box ($cx + 78 - $o - 1.2) 285 1.2 148 $flute) }
                $set += (Tx $roman[$i] 15 $Accent ($cx - 70) 293 140 $heavyBold $heavy 2)
                $set += (Tx $r[0] 20 $Ink ($cx - 76) 315 152 -1 $Font 2)
                if ($r.Count -gt 1) { $set += (Tx (Tail $r) 15 $body ($cx - 76) 349 152 0 $Font 2 0 1.4) }
                $cols += (Grp $set)
            }
            $baseSet = @((Box ($cxm - 388) 448 776 30 $Dark), (Box ($cxm - 408) 478 816 14 (Mix $Dark $Light 0.35)))
            if ($card.Count -gt 1) { $baseSet += (Mid (Tx $card[1][0] 15.5 '#FFFFFF' ($cxm - 380) 448 760 -1 $Font 2) 448 30) }
            $base = Grp $baseSet
            Anim $base $WIPE $AFTER 0.5 0 4
            for ($i = 0; $i -lt $n; $i++) { Anim $cols[$i] $WIPE $(if ($i) { $WITH } else { $AFTER }) 0.55 (0.16 * $i) 3 }
            Anim $beam $WIPE $AFTER 0.45 0 4; Anim $roof $FADE $AFTER 0.3; [void](Glide $roof 'M 0 -0.3 L 0 0 E' 0.55 $WITH); Anim $sun $WIPE $AFTER 1.1 0 3
            Foot; $secs = 7
        }
        'RibbonFlow' {
            Head
            $rh = 84; $gp = 30; $y0 = 138; $ro = $rh + $gp / 2; $xa = 65 + $ro; $xb = 895 - $ro; $tip = 861
            $rc = @($Dark, (Mix $Dark $Accent 0.5), $Accent); $tc = @((Mix (Mix $Dark $Accent 0.25) '#000000' 0.18), (Mix (Mix $Dark $Accent 0.75) '#000000' 0.18))
            $bands = @(); $turns = @()
            for ($k = 0; $k -lt 3; $k++) {
                $r = $rows[$k]; $ry = $y0 + $k * ($rh + $gp); $back = ($k -eq 1)
                $rx0 = $(if ($k) { $xa } else { 65 }); $rx1 = $(if ($k -eq 2) { $tip } else { $xb })
                $set = @((Box $rx0 $ry ($rx1 - $rx0) $rh $rc[$k]))
                if ($k -eq 2) { $set += (Poly @(@(($tip - 1), ($ry - 14)), @(($tip + 46), ($ry + $rh / 2)), @(($tip - 1), ($ry + $rh + 14))) $rc[$k]) }
                $lbx = $(if ($back) { $rx1 - 104 } else { $rx0 + 22 }); $ix0 = $(if ($back) { $rx0 + 8 } else { $rx0 + 124 }); $ix1 = $(if ($back) { $rx1 - 124 } else { $rx1 - 8 })
                $set += (Mid (Tx $r[0] 20 $Light $lbx $ry 90 $heavyBold $heavy 1 0 1.15) $ry $rh)
                $set += (Box $(if ($back) { $ix1 + 6 } else { $ix0 - 10 }) ($ry + 20) 1 ($rh - 40) (Mix $rc[$k] '#FFFFFF' 0.35))
                $m = [math]::Floor(($r.Count - 1) / 2); $cwid = ($ix1 - $ix0) / $m
                for ($j = 0; $j -lt $m; $j++) {
                    $slot = $(if ($back) { $m - 1 - $j } else { $j }); $x = $ix0 + $slot * $cwid + 16
                    $set += (Tx ($r[1 + 2 * $j]) 18 '#FFFFFF' $x ($ry + 17) ($cwid - 22) -1)
                    $set += (Tx ($r[2 + 2 * $j]) 15 (Mix '#FFFFFF' $rc[$k] 0.12) $x ($ry + 45) ($cwid - 22) 0)
                    if ($slot) { $set += (Mid (Tx ([string][char]$(if ($back) { 0x2039 } else { 0x203A })) 26 (Mix $rc[$k] '#FFFFFF' 0.5) ($ix0 + $slot * $cwid - 8) $ry 16 0 $Font 2) ($ry - 2) $rh) }
                }
                $bands += (Grp $set)
            }
            foreach ($q in 0, 1) {                                                                                                  # half rings: the ribbon turns back
                $tcx = $(if ($q) { $xa } else { $xb }); $tcy = $y0 + $rh + $gp / 2 + $q * ($rh + $gp)
                $s = Tag ($sl.Shapes.AddShape(20, [single]($tcx - $ro), [single]($tcy - $ro), [single](2 * $ro), [single](2 * $ro)))
                $s.Adjustments.Item(1) = $(if ($q) { 90 } else { 270 }); $s.Adjustments.Item(2) = $(if ($q) { 270 } else { 90 }); $s.Adjustments.Item(3) = [single]($rh / (2 * $ro))
                $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $tc[$q]); $turns += $s
            }
            Anim $bands[0] $WIPE $AFTER 0.7 0 4; Anim $turns[0] $WIPE $AFTER 0.3 0 1; Anim $bands[1] $WIPE $AFTER 0.7 0 2
            Anim $turns[1] $WIPE $AFTER 0.3 0 1; Anim $bands[2] $WIPE $AFTER 0.7 0 4
            if ($Sub) { $cl = Grp @((Box 65 477 4 23 $Accent), (Tx $Sub 19 $Ink 80 476 800 -1)); Anim $cl $FADE $AFTER 0.5 }
            Foot; $secs = 6
        }
        'Dossier' {
            $meta = Rows $Cards
            [void](Poly @(@(0, 0), @(430, 0), @(330, 540), @(0, 540)) $Dark); [void](Poly @(@(430, 0), @(440, 0), @(340, 540), @(330, 540)) $Accent)
            if ($Kicker) { [void](Tx $Kicker 12.5 $soft 65 45 320 -1 $Font 1 3) }
            [void](Tx $Title 30 '#FFFFFF' 65 70 320 $heavyBold $heavy 1 0 1.15)
            $hy0 = 196; $hl = Tx $meta[0][0] 15 $soft 65 $hy0 200 -1
            $hero = Tx $meta[0][1] 108 $Accent 61 ($hy0 + 10) 330 $heavyBold $heavy 1 0 1.0
            $hwid = [double]$hero.TextFrame.TextRange.BoundWidth
            $unit = Tx $(if ($meta[0].Count -gt 2) { $meta[0][2] } else { '' }) 26 '#FFFFFF' (61 + $hwid + 8) ([double]$hero.Top + [double]$hero.Height - 58) 130 -1
            $facts = @()
            for ($q = 1; $q -lt $meta.Count; $q++) {
                $yy = 374 + ($q - 1) * 46; $ex = 430 - 100 * $yy / 540
                $facts += (Grp @((Box 65 ($yy - 12) ($ex - 65 - 36) 1 $dLine), (Tx $meta[$q][0] 15 $soft 65 ($yy + 1) 60 -1), (Tx $meta[$q][1] 17 '#FFFFFF' 122 $yy ($ex - 122 - 24) -1)))
            }
            $lines = @()
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $rt = 58 + $i * 112; $xs = 430 - 100 * ($rt + 50) / 540 + 58
                $set = @((Tx ('{0:00}' -f ($i + 1)) 44 $Accent $xs ($rt - 4) 74 $heavyBold $heavy), (Tx $r[0] 22 $Ink ($xs + 80) ($rt + 2) (895 - $xs - 80) -1))
                if ($r.Count -gt 1) { $set += (Tx (Tail $r) 16.5 $body ($xs + 80) ($rt + 36) (895 - $xs - 80) 0 $Font 1 0 1.4) }
                if ($i -lt $n - 1) { $set += (Box ($xs - 6) ($rt + 99) (895 - $xs + 6) 1 (Mix $Light $Ink 0.16)) }
                $lines += (Grp $set)
            }
            $stamp = $null
            if ($Sub) {
                $scy = $hy0 + 36; $scx = [math]::Min(61 + $hwid + 62, 430 - 100 * $scy / 540 - 56)
                $stamp = Grp @((Ring $scx $scy 44 $Accent 3), (Ring $scx $scy 37 $Accent 1), (Mid (Tx $Sub 16 $Accent ($scx - 40) $scy 80 $heavyBold $heavy 2 0 1.1) ($scy - 30) 60)); $stamp.Rotation = 346
            }
            Anim $hero $WIPE $AFTER 0.6 0 3; Anim $unit $FADE $WITH 0.3 0.4; Anim $hl $FADE $WITH 0.3
            for ($q = 0; $q -lt $facts.Count; $q++) { Anim $facts[$q] $FADE $(if ($q) { $WITH } else { $AFTER }) 0.35 (0.15 * $q) }
            foreach ($g in $lines) { Anim $g $WIPE $AFTER 0.45 0 4 }
            if ($stamp) { Anim $stamp 23 $AFTER 0.25 0.3 }
            Foot $true; $secs = 7
        }
        'HubSpoke' {
            $hx = 150; $hy = 313; $r1 = 150; $r2 = 215; $r3 = 400; $hub = (Rows $Sub)[0]; $com = Rows $Cards; $yA = 148; $yB = 478; $paths = @()
            $dots = @(); for ($k = 0; $k -lt $n; $k++) { $dots += (Disc $hx $hy 5 $Light) }                                          # signals wait under the hub
            $orb = Ring $hx $hy $r3 $dLine 1; $orb.Line.DashStyle = 4
            [void](Disc $hx $hy $r2 $dCard); $core = Disc $hx $hy $r1 $Accent; Glow $core $Accent 0.5 34 0
            if ($hub.Count -gt 2) { [void](Tx $hub[2] 12.5 '#FFFFFF' ($hx - 140) ($hy - 78) 280 -1 $Font 2 3) }
            [void](Tx $hub[0] 28 '#FFFFFF' ($hx - 145) ($hy - 52) 290 $heavyBold $heavy 2 0 1.1)
            if ($hub.Count -gt 1) { [void](Tx $hub[1] 15.5 '#FFFFFF' ($hx - 140) ($hy + 34) 280 -1 $Font 2) }
            if ($com.Count) {
                [void](Tx $com[0][0] 19 $Accent ($hx - 130) ($hy - $r2 + 14) 260 -1 $Font 2)
                if ($com[0].Count -gt 1) { [void](Tx $com[0][1] 15 $dText ($hx - 130) ($hy - $r2 + 40) 260 0 $Font 2) }
                if ($com.Count -gt 1) { [void](Tx $com[1][0] 15 $dText ($hx - 130) ($hy + $r1 + 9) 260 0 $Font 2 0 1.3) }
            }
            if ($Kicker) { [void](Tx $Kicker 12.5 $soft 590 45 310 -1 $Font 1 3) }
            [void](Tx $Title 28 '#FFFFFF' 590 68 320 -1)
            if ($Note) { [void](Tx $Note 9 (Mix $Dark '#FFFFFF' 0.5) 590 516 300) }
            for ($k = 0; $k -lt $n; $k++) {
                $ny = $yA + $k * ($yB - $yA) / [math]::Max(1, $n - 1); $nx = $hx + [math]::Sqrt($r3 * $r3 - ($ny - $hy) * ($ny - $hy)); $ux = ($nx - $hx) / $r3; $uy = ($ny - $hy) / $r3
                $spoke = Seg ($hx + ($r2 + 8) * $ux) ($hy + ($r2 + 8) * $uy) ($hx + ($r3 - 26) * $ux) ($hy + ($r3 - 26) * $uy) $Accent 1.75
                $nd = Disc $nx $ny 22 $Dark; $nd.Line.Visible = -1; $nd.Line.ForeColor.RGB = (Rgb $Accent); $nd.Line.Weight = 2.5
                $set = @($nd, (Mid (Tx ('{0}' -f ($k + 1)) 18 $Accent ($nx - 22) $ny 44 $heavyBold $heavy 2) ($ny - 22) 44), (Tx $rows[$k][0] 20 '#FFFFFF' 590 ($ny - 26) 305 -1))
                if ($rows[$k].Count -gt 1) { $set += (Tx $rows[$k][1] 15.5 $dText 590 ($ny + 3) 305 0) }
                Anim $spoke $WIPE $AFTER 0.3 0 4; Anim (Grp $set) $FADE $WITH 0.3 0.15
                $paths += ('M 0 0 L {0:F4} {1:F4} E' -f (($nx - $hx) / $W), (($ny - $hy) / $H))
            }
            for ($k = 0; $k -lt $n; $k++) { $fx = Glide $dots[$k] $paths[$k] 1.3 $(if ($k) { $WITH } else { $AFTER }) (0.22 * $k); $fx.Timing.RepeatCount = 3 }
            $secs = 10
        }
        'MetroOrg' {
            $root = (Rows $Sub)[0]; $stat = Rows $Cards; $rx = 128; $ry = 300; $endX = 880
            $lc = @($Accent, $Dark, (Mix $Accent '#FFFFFF' 0.38)); $trains = @(); $paths = @()
            for ($k = 0; $k -lt $n; $k++) { $tr = Disc $rx $ry 9.5 "#FFFFFF"; Glow $tr $Ink 0.55 8 2; $tr.Line.Visible = -1; $tr.Line.ForeColor.RGB = (Rgb $lc[$k % 3]); $tr.Line.Weight = 3; $trains += $tr }
            Head
            $tracks = @(); $stops = @()
            for ($k = 0; $k -lt $n; $k++) {
                $ly = $ry + ($k - ($n - 1) / 2) * 130; $dy = $ly - $ry; $col = $lc[$k % 3]; $r = $rows[$k]
                if ([math]::Abs($dy) -lt 1) { $tracks += (Box $rx ($ry - 4.5) ($endX - $rx) 9 $col); $paths += ('M 0 0 L {0:F4} 0 E' -f (($endX - $rx) / $W)) }
                else {
                    $tracks += (Ln @(@($rx, $ry), @(($rx + 52), $ry), @(($rx + 52 + [math]::Abs($dy)), $ly), @($endX, $ly)) $col 9)
                    $paths += ('M 0 0 L {0:F4} 0 L {1:F4} {2:F4} L {3:F4} {2:F4} E' -f (52 / $W), ((52 + [math]::Abs($dy)) / $W), ($dy / $H), (($endX - $rx) / $W))
                }
                $set = @()
                $set += , (Grp (@(Face $(if ($Images.Count -gt $k + 1) { $Images[$k + 1] } else { '' }) 360 $ly 36 $col) + @((Tx $r[0] 18 $Ink 280 ($ly + 44) 160 -1 $Font 2).Name, (Tx $r[1] 15 $body 280 ($ly + 67) 160 0 $Font 2).Name)))
                $m = [math]::Floor(($r.Count - 2) / 2)
                for ($j = 0; $j -lt $m; $j++) {
                    $tx = 525 + $j * 147; $st = Disc $tx $ly 10 '#FFFFFF'; $st.Line.Visible = -1; $st.Line.ForeColor.RGB = (Rgb $col); $st.Line.Weight = 4
                    $set += , (Grp @($st, (Tx ($r[2 + 2 * $j]) 17 $Ink ($tx - 73) ($ly + 18) 146 -1 $Font 2), (Tx ($r[3 + 2 * $j]) 15 $body ($tx - 73) ($ly + 41) 146 0 $Font 2)))
                }
                $set += , (Grp @((Disc $endX $ly 14 $col), (Mid (Tx ('{0}' -f ($k + 1)) 15 '#FFFFFF' ($endX - 14) $ly 28 $heavyBold $heavy 2) ($ly - 14) 28)))
                $stops += , $set
            }
            [void](Disc $rx $ry 60 $Light)
            [void](Grp (@(Face $(if ($Images.Count) { $Images[0] } else { '' }) $rx $ry 52 $Dark) + @((Tx $root[0] 19 $Ink ($rx - 100) ($ry + 64) 200 -1 $Font 2).Name, (Tx $(if ($root.Count -gt 1) { $root[1] } else { '' }) 15 $body ($rx - 100) ($ry + 89) 200 0 $Font 2).Name)))
            $sxp = 65
            foreach ($c in $stat) {
                $num = Tx $c[0] 30 $Accent $sxp 452 80 $heavyBold $heavy; $nw = [double]$num.TextFrame.TextRange.BoundWidth
                [void](Tx $c[1] 15 $body ($sxp + $nw + 3) 467 60 -1); $sxp += $nw + 3 + 15 * $c[1].Length + 16
            }
            for ($k = 0; $k -lt $n; $k++) {
                Anim $tracks[$k] $WIPE $AFTER 0.7 0 4
                for ($j = 0; $j -lt $stops[$k].Count; $j++) { Anim $stops[$k][$j] $FADE $WITH 0.25 (0.2 + 0.12 * $j) }
            }
            for ($k = 0; $k -lt $n; $k++) { [void](Glide $trains[$k] $paths[$k] 2.2 $(if ($k) { $WITH } else { $AFTER }) (0.25 * $k)) }
            Foot; $secs = 10
        }
        'Gantt' {
            Head
            $cols = Rows $Cards; $nc = $cols.Count; $gx0 = 250; $gx1 = 895; $cw = ($gx1 - $gx0) / $nc; $gy0 = 162; $gy1 = 478
            $now = (Rows $Sub)[0]; $pos = [double]$now[0]; $xt = $gx0 + $pos * $cw; $from = 0
            for ($q = 0; $q -lt $nc; $q++) {
                if ($q % 2 -eq 0) { [void](Box ($gx0 + $q * $cw) $gy0 $cw ($gy1 - $gy0) (Mix $Light $Ink 0.05)) }
                [void](Tx $cols[$q][1] 15 $body ($gx0 + $q * $cw) 136 $cw 0 $Font 2)
                if ($q -eq $nc - 1 -or $cols[$q + 1][0] -ne $cols[$q][0]) {
                    [void](Tx $cols[$q][0] 15 $Accent ($gx0 + $from * $cw + 3) 111 200 $heavyBold $heavy); [void](Box ($gx0 + $from * $cw + 3) 131 (($q - $from + 1) * $cw - 6) 2 $Ink); $from = $q + 1
                }
            }
            $pitch = [math]::Min(64, ($gy1 - $gy0) / $n); $plans = @(); $marks = @(); $dones = @()
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $ry = $gy0 + $i * $pitch; $by = $ry + $pitch - 29; $xs = $gx0 + [double]$r[1] * $cw + 3; $xe = $gx0 + [double]$r[2] * $cw - 3
                if ($i) { [void](Box 65 $ry ($gx1 - 65) 1 (Mix $Light $Ink 0.12)) }
                [void](Mid (Tx $r[0] 17.5 $Ink 65 $ry 180 -1) ($by - 10) 42)
                $pl = Box $xs $by ($xe - $xs) 22 $Accent 11; $pl.Fill.Patterned(26); $pl.Fill.ForeColor.RGB = (Rgb (Mix $Accent $Light 0.4)); $pl.Fill.BackColor.RGB = (Rgb (Mix $Accent $Light 0.8)); $plans += $pl
                if ($xt -gt $xs + 22) { $dones += (Box $xs $by ([math]::Min($xe, $xt) - $xs) 22 $Accent 11) }
                if ($r.Count -gt 3 -and $r[3]) {
                    $dm = Tag ($sl.Shapes.AddShape(4, [single]($xe - 12), [single]($by - 1), 24, 24)); $dm.Fill.ForeColor.RGB = (Rgb $Dark); $dm.Line.ForeColor.RGB = (Rgb $Light); $dm.Line.Weight = 2
                    $marks += (Grp @($dm, (Tx $r[3] 15 $AccentDark ($xe + 12 - 240) ($by - 23) 240 -1 $Font 3)))
                } else { $marks += $null }
            }
            $pw = Wd $now[1] 15
            $tg = Grp (@((Box ($xt - 1) ($gy0 - 6) 2 ($gy1 - $gy0 + 10) $Dark).Name) + (Pill ($xt - $pw / 2) ($gy1 + 4) $pw 27 $Dark $now[1] 15 '#FFFFFF'))
            for ($i = 0; $i -lt $n; $i++) { Anim $plans[$i] $WIPE $AFTER 0.5 0 4; if ($marks[$i]) { Anim $marks[$i] $WHEEL $WITH 0.35 0.4 } }
            Anim $tg $FADE $AFTER 0.2; [void](Glide $tg ('M {0:F4} 0 L 0 0 E' -f (($gx0 - $xt) / $W)) 1.3 $WITH)
            for ($i = 0; $i -lt $dones.Count; $i++) { Anim $dones[$i] $WIPE $(if ($i) { $WITH } else { $AFTER }) 0.6 0 4 }
            Foot; $secs = 9
        }
        'Runway' {
            $vx = $W / 2; $vy = 172; $by = 546; $nearHW = 190; $farHW = 14
            function HW($yy) { $farHW + ($nearHW - $farHW) * ($yy - $vy) / ($by - $vy) }
            $dawn = Disc $vx $vy 100 $Accent; $dawn.Fill.Transparency = 0.78; $dawn.SoftEdge.Type = 6
            [void](Box 0 $vy $W 1 $dLine)
            [void](Poly @(@(($vx - $farHW), $vy), @(($vx + $farHW), $vy), @(($vx + $nearHW), $by), @(($vx - $nearHW), $by)) (Mix $Dark '#FFFFFF' 0.1))
            for ($j = 0; $j -lt 7; $j++) {
                $ya = $vy + ($by - $vy) * [math]::Pow(($j + 0.15) / 7, 1.7); $yb = $vy + ($by - $vy) * [math]::Pow(($j + 0.6) / 7, 1.7); $wa = (HW $ya) * 0.035; $wb = (HW $yb) * 0.035
                [void](Poly @(@(($vx - $wa), $ya), @(($vx + $wa), $ya), @(($vx + $wb), $yb), @(($vx - $wb), $yb)) (Mix $Dark '#FFFFFF' 0.45))
            }
            Head $true
            $hero = (Rows $Cards)[0]
            if ($hero) { [void](Tx $hero[0] 62 $Accent 555 58 340 $heavyBold $heavy 3); if ($hero.Count -gt 1) { [void](Tx $hero[1] 15 $dText 555 44 340 -1 $Font 3) } }
            $lamps = @()
            for ($j = 1; $j -le 10; $j++) {
                $tt = [math]::Pow($j / 10, 1.5); $yy = $vy + ($by - $vy) * $tt; if ($yy -gt 530) { continue }
                $off = (HW $yy) + 6 + 9 * $tt; $rr = 1.8 + 4.6 * $tt
                $a = Disc ($vx - $off) $yy $rr $Accent; $b = Disc ($vx + $off) $yy $rr $Accent; Glow $a $Accent 0.3 9 0; Glow $b $Accent 0.3 9 0; $lamps += (Grp @($a, $b))
            }
            for ($j = $lamps.Count - 1; $j -ge 0; $j--) { $d = $lamps.Count - 1 - $j; Anim $lamps[$j] $FADE $(if ($d) { $WITH } else { $AFTER }) 0.15 (0.12 * $d) }
            for ($k = 0; $k -lt $n; $k++) {
                $r = $rows[$k]; $tt = 0.8 - $k * 0.6 / [math]::Max(1, $n - 1); $yy = $vy + ($by - $vy) * $tt; $side = $(if ($k % 2) { 1 } else { -1 }); $th = 2 + 3 * $tt
                $stripe = Poly @(@(($vx - (HW ($yy - $th))), ($yy - $th)), @(($vx + (HW ($yy - $th))), ($yy - $th)), @(($vx + (HW ($yy + $th))), ($yy + $th)), @(($vx - (HW ($yy + $th))), ($yy + $th))) $Accent
                $edge = $vx + $side * ((HW $yy) + 14); $ax = $edge + $side * 36; $bx = $(if ($side -lt 0) { $ax - 270 } else { $ax }); $al = $(if ($side -lt 0) { 3 } else { 1 })
                $set = @((Box ([math]::Min($edge, $edge + $side * 26)) ($yy - 1) 26 2 $Accent), (Tx $r[0] 16.5 $Accent $bx ($yy - 37) 270 $heavyBold $heavy $al), (Tx $r[1] 21 '#FFFFFF' $bx ($yy - 15) 270 -1 $Font $al))
                if ($r.Count -gt 2) { $set += (Tx $r[2] 15.5 $dText $bx ($yy + 14) 270 0 $Font $al) }
                Anim $stripe $WIPE $AFTER 0.25 0 4; Anim (Grp $set) $FADE $WITH 0.35 0.1
            }
            $run = Disc $vx 526 6 $Light; Glow $run $Accent 0.15 14 0
            Anim $run $FADE $AFTER 0.15; [void](Glide $run ('M 0 0 L 0 {0:F4} E' -f (($vy - 526) / $H)) 1.1 $AFTER)
            if ($Sub) {
                $pw = Wd $Sub 17
                $goal = Grp (@((Box ($vx - 1) 150 2 22 $Accent).Name) + (Pill ($vx - $pw / 2) 116 $pw 34 $Accent $Sub 17 $Dark)); Anim $goal $RISE $AFTER 0.5
            }
            Anim $dawn $FADE $WITH 0.9
            Foot $true; $secs = 9
        }
    }
    Save-Deck $Out $Preview $secs
    "saved $Out"
} catch { Close-Deck; throw }
