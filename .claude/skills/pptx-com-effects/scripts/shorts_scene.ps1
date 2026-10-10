# Fourteen scenes grown out of short tutorial videos (2026-10 batch, part 4). Each keeps the video's trick and adds one idea of
# its own (after the dash). Same colour / font parameters as scene_page.ps1. All of them play by themselves.
#   WaveSplit    two colour fields meet along a wave; the title starts alone in the middle - two wave layers slide in at different speeds and a round photo grows
#                -Images photo  -Title 'ORANGE|FUSION' [-Kicker k]  -Items 'line|line'
#   TabDeck      numbered strips lie stacked like a deck, then spread across the slide - their edges peek out first, a progress line runs under them
#                -Title 'FOUR|steps'  -Items 'A;head;line;line|B;head;line|...' (3-5)
#   WaveWindow   a long picture drifts behind a wave-shaped window cut in the paper - several photos joined as one film, captions appear as their part passes
#                -Images p1,p2,p3  -Title T  -Items 'head;line|head;line|...' (2-4)
#   RingWindows  a broken ring is a window onto the undimmed slide background; it turns to the next part - a second, accent ring turns the other way
#                -Images p1,p2,...  -Items 'PART 01;head;line|...' (one photo per part; 16:9 photos, the background is stretched)
#   KeywordLinks keyword circles joined to their explanation by a dotted-end line on a dim photo - a spine is drawn first, then each link draws itself
#                -Images photo  -Title T [-Sub s] [-Kicker watermark]  -Items 'keyword;head;line|...' (3-4)
#   GridCells    a table grid lies over the photo and some cells are filled to carry the title - the photo starts fully tiled and the cells clear in a diagonal wave
#                -Images photo  -Title 'PROFESSIONAL|PRESENTATION' [-Sub s]  -Items 'value;label|...' (0-3 accent cells)
#   MosaicRise   one photo cut into unequal tiles - the tiles arrive from above and below in turn, one tile is a colour block with a figure
#                -Images photo  -Title T [-Kicker k]  -Items 'line|line'  [-Cards 'value;label']
#   EraRoll      the year rolls digit by digit to the next stop - neighbouring digits roll in opposite directions, the colour and a timeline move with it
#                -Title T  -Items '1975;head;line|1991;head;line|...' (2-5, four-digit years)
#   ArchColumns  colour columns, one opens per step - each carries a capsule photo window that grows with its column
#                -Images p1,p2,...  -Items 'label;head;line|...' (3-5, one photo each)  [-Cards '#F8C8C0|#FBE3E8|...' column colours]
#   LensTour     a magnifying lens travels from spot to spot - along a drawn route, and each visited spot stays as a snapshot in the corner
#                -Images photo  -Title T  -Items 'fx;fy;head;line|...' (2-4, fx/fy = 0..1 on the slide)
#   DiamondBricks a tilted brick wall cut out of one photo - the bricks assemble in a wave and some are colour bricks, one with a figure
#                -Images photo  -Title 'REAL ESTATE' [-Sub s]  [-Cards 'value;label']
#   PaperPeel    curved paper sheets with a shadow slide apart and the title is underneath - three sheets in three shades leave in three directions
#                -Title WELCOME [-Sub s] [-Kicker k]
#   TornReveal   a dark sheet is torn down the middle and the photo shows in the gap - the tear opens in two stages and has a pale paper edge
#                -Images photo  -Title T [-Kicker k]  -Items 'line|line'
#   FoldArrows   a timeline of arrows, each tucked under the next with a soft fold shadow - they unfold one by one and the last one is lifted
#                -Title T  -Items '2019;head;line|2020;head;line|...' (3-5)
# Example: .\shorts_scene.ps1 -Layout EraRoll -Title 'History of the PC' -Items '1975;IBM 5100;The first portable|1991;PowerBook;The laptop takes shape' -Out C:\work\era.pptx -Preview C:\work\era.mp4
param(
    [Parameter(Mandatory)][ValidateSet('WaveSplit', 'TabDeck', 'WaveWindow', 'RingWindows', 'KeywordLinks', 'GridCells', 'MosaicRise', 'EraRoll', 'ArchColumns', 'LensTour', 'DiamondBricks', 'PaperPeel', 'TornReveal', 'FoldArrows')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Note = '',
    [string]$Accent = '#E8743B',
    [string]$AccentDark = '',
    [string]$Light = '#FBF3EA',
    [string]$Dark = '#14253D',
    [string]$Ink = '#1F1B2E',
    [string]$Font = 'Malgun Gothic',
    [string]$HeavyFont = '',
    [double]$Duration = 1.1,
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
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function On($slide) { $script:sl = $slide }
function Need([int]$count) { if ($Images.Count -lt $count) { throw "$Layout needs $count photo(s) in -Images" } }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Cut([string]$path, $shape) { $copy = Add-Cover $path; Tag (Merge-Shapes 3 @($copy.Name, $shape.Name)) }     # the slide-filling photo, kept only inside $shape (picture first: stays a picture)
function Fit([string]$path, $x, $y, $bw, $bh, [int]$type = 1) {       # a photo covering the box, cropped to it
    $p = Tag ($sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1)); $p.LockAspectRatio = -1
    $k = [math]::Max($bw / [double]$p.Width, $bh / [double]$p.Height); $p.Width = [single]([double]$p.Width * $k)
    $p.Left = [single]($x + ($bw - [double]$p.Width) / 2); $p.Top = [single]($y + ($bh - [double]$p.Height) / 2)
    $c = $p.PictureFormat.Crop; $c.ShapeLeft = [single]$x; $c.ShapeTop = [single]$y; $c.ShapeWidth = [single]$bw; $c.ShapeHeight = [single]$bh
    if ($type -ne 1) { $p.AutoShapeType = $type }
    $p
}
function Rule($x1, $y1, $x2, $y2, [string]$hex, $wt = 1) { $l = Tag ($sl.Shapes.AddLine([single]$x1, [single]$y1, [single]$x2, [single]$y2)); $l.Line.ForeColor.RGB = (Rgb $hex); $l.Line.Weight = [single]$wt; $l }
function Leave($shape, $dur, $delay) { $e = $sl.TimeLine.MainSequence.AddEffect($shape, $FADE, 0, $WITH); $e.Timing.Duration = [single]$dur; $e.Timing.TriggerDelayTime = [single]$delay; $e.Exit = -1 }
function Show($shape) { Anim $shape 1 $WITH 0.01 0 }                   # a shape whose first effect is a custom one is hidden until it starts
function Hold($sec = 2) { $b = Box -40 -40 10 10 $Light; $b.Fill.Transparency = 1; Anim $b $FADE $AFTER 0.1 $sec }
function Wave($x0, $amp, $phase, [double]$turns = 1.2) {               # points of a wavy vertical edge, top to bottom
    $pts = @(); for ($y = -10; $y -le $H + 10; $y += 14) { $pts += , @(($x0 + $amp * [math]::Sin(2 * [math]::PI * $turns * $y / $H + $phase)), $y) }; , $pts
}

try {
    New-Deck $Light
    $rows = Rows $Items; $n = $rows.Count; $cardRows = Rows $Cards
    switch ($Layout) {
        'WaveSplit' {
            Need 1
            $back = Poly ((Wave 470 34 1.4) + @(, @(($W + 500), ($H + 10))) + @(, @(($W + 500), -10))) (Mix $Accent '#FFFFFF' 0.62); $back.Name = 'WaveBack'
            $front = Poly ((Wave 540 30 0.2) + @(, @(($W + 500), ($H + 10))) + @(, @(($W + 500), -10))) $Accent; $front.Name = 'WaveFront'
            $ph = Fit $Images[0] 580 120 300 300 9; $ph.Line.Visible = -1; $ph.Line.ForeColor.RGB = 0xFFFFFF; $ph.Line.Weight = 6; $ph.Name = 'Photo'; Glow $ph '#000000' 0.7 26 10
            $tt = Tx $Title 56 $Ink 65 170 440 $heavyBold $heavy 1 0 1.05; $tt.Name = 'Title'
            $extra = @(); if ($Kicker) { $extra += (Tx $Kicker 13 $Accent 68 138 400 -1 $Font 1 3) }
            $y = [double]$tt.Top + [double]$tt.Height + 18; $extra += (Box 68 $y 56 5 $Accent); $y += 24
            foreach ($r in $rows) { $b = Tx $r[0] 17 $body 68 $y 400 0 $Font 1 0 1.4; $extra += $b; $y += [double]$b.Height + 8 }
            (Grp $extra).Name = 'Copy'
            Steps 2
            $s1 = $pres.Slides(1); $s1.Shapes('Copy').Delete()
            $s1.Shapes('WaveFront').Left = [single]([double]$s1.Shapes('WaveFront').Left + 420); $s1.Shapes('WaveBack').Left = [single]([double]$s1.Shapes('WaveBack').Left + 250)   # the far layer travels less: depth
            $p1 = $s1.Shapes('Photo'); $p1.Width = 60; $p1.Left = [single]($W + 40); $p1.Top = 240
            $t1 = $s1.Shapes('Title'); $t1.TextFrame.TextRange.Font.Size = 84; $t1.TextFrame.TextRange.ParagraphFormat.Alignment = 2; $t1.Width = [single]($W - 130); $t1.Left = 65; $t1.Top = [single](($H - [double]$t1.Height) / 2)
            Auto $s1 0.9
        }
        'TabDeck' {
            if ($n -lt 3) { throw 'TabDeck needs 3-5 items' }
            $bg = Box 0 0 $W $H $Dark; $bg.Name = 'Bg'
            $sw = ($W - 130) / $n; $top = 128; $bh = 340
            for ($i = $n - 1; $i -ge 0; $i--) {                                                          # last first: the first strip ends up on top
                $r = $rows[$i]; $col = Mix $Accent '#FFFFFF' (0.16 * $i); $x = 65 + $i * $sw
                $set = @((Box $x $top $sw $bh $col))
                if ($i -lt $n - 1) { $tri = Tag ($sl.Shapes.AddShape(7, [single]($x + $sw - 13), [single]($top + 44), 44, 22)); $tri.Rotation = 90; $tri.Line.Visible = 0; $tri.Fill.ForeColor.RGB = (Rgb $col); $set += $tri }
                $set += (Tx $r[0] 62 '#FFFFFF' ($x + 24) ($top + 20) ($sw - 40) $heavyBold $heavy)
                if ($r.Count -gt 1) { $set += (Tx $r[1] 18 '#FFFFFF' ($x + 24) ($top + 118) ($sw - 44) -1) }
                if ($r.Count -gt 2) { $set += (Tx (($r[2..($r.Count - 1)]) -join '|') 13.5 '#FFFFFF' ($x + 24) ($top + 156) ($sw - 44) 0 $Font 1 0 1.5) }
                (Grp $set).Name = "Step$i"
            }
            $tt = Tx $Title 26 '#FFFFFF' 65 52 700 $heavyBold $heavy; $tt.Name = 'Title'
            Steps 2
            $s1 = $pres.Slides(1)
            for ($i = 0; $i -lt $n; $i++) { $s1.Shapes("Step$i").Left = [single](65 + $i * 16) }           # a deck: each strip shows only its edge
            $t1 = $s1.Shapes('Title'); $t1.TextFrame.TextRange.Font.Size = 72; $t1.Left = [single](65 + $sw + 16 * $n + 60); $t1.Top = [single](($H - [double]$t1.Height) / 2)
            Auto $s1 1.0
            $track = Box 65 ($top + $bh + 16) ($W - 130) 4 (Mix $Dark '#FFFFFF' 0.2); $run = Box 65 ($top + $bh + 16) ($W - 130) 4 '#FFFFFF'
            Anim $track $FADE $AFTER 0.3 0.2; Anim $run $WIPE $AFTER 1.6 0 4; Hold 1.5
        }
        'WaveWindow' {
            Need 2
            $fw = 520; $fy = 120; $fh = 300; $film = @()
            for ($i = 0; $i -lt $Images.Count; $i++) { $film += (Fit $Images[$i] ($i * $fw) $fy $fw $fh) }
            $strip = Grp $film; $travel = $Images.Count * $fw - $W
            $top = @(); $bot = @(); for ($x = 40; $x -le 920; $x += 16) { $top += , @($x, (176 + 24 * [math]::Sin($x / 92.0))); $bot += , @($x, (372 + 24 * [math]::Sin($x / 92.0 + 1.9))) }
            [array]::Reverse($bot); $hole = Poly ($top + $bot) $Ink
            $paper = Box -10 -10 ($W + 20) ($H + 20) $Light; $paper = Merge-Shapes 4 @($paper.Name, $hole.Name); $paper.Fill.ForeColor.RGB = (Rgb $Light); $paper.Line.Visible = 0
            Head; Foot
            $caps = @()
            for ($i = 0; $i -lt $n; $i++) {
                $x = 90 + $i * (($W - 180) / $n); $up = ($i % 2 -eq 0); $y = $(if ($up) { 122 } else { 420 })
                $set = @((Disc ($x + 6) $(if ($up) { 160 } else { 408 }) 5 $Accent), (Tx $rows[$i][0] 17 $Ink ($x + 20) ($y - 6) 200 -1))
                if ($rows[$i].Count -gt 1) { $set += (Tx $rows[$i][1] 13 $body ($x + 20) ($y + 20) 200) }
                $caps += Grp $set
            }
            $secs = 3.2 * $Images.Count
            Show $strip
            $e = $sl.TimeLine.MainSequence.AddEffect($strip, 0, 0, $AFTER)
            $e.Behaviors.Add(1).MotionEffect.Path = ('M 0 0 L {0} 0 E' -f ([math]::Round(-$travel / $W, 4)).ToString([cultureinfo]::InvariantCulture))
            $e.Timing.Duration = [single]$secs; $e.Timing.SmoothStart = 0; $e.Timing.SmoothEnd = 0; $e.Timing.TriggerDelayTime = 0.3      # even speed, like a camera pan
            for ($i = 0; $i -lt $n; $i++) { Anim $caps[$i] $FADE $WITH 0.5 (0.6 + $i * ($secs - 1.2) / $n) }
            Hold 1.5
        }
        'RingWindows' {
            Need $n
            $scrim = Box 0 0 $W $H $Dark; $scrim.Fill.Transparency = 0.2; $scrim.Name = 'Scrim'
            $cx = 700; $cy = 300; $ro = 310
            $d = $sl.Shapes.AddShape(18, [single]($cx - $ro), [single]($cy - $ro), [single](2 * $ro), [single](2 * $ro)); $d.Adjustments.Item(1) = 0.3
            $v = New-Rect ($cx - 9) ($cy - $ro - 5) 18 (2 * $ro + 10); $hbar = New-Rect ($cx - $ro - 5) ($cy - 9) (2 * $ro + 10) 18
            $ring = Merge-Shapes 4 @($d.Name, $v.Name, $hbar.Name); $ring.Name = 'Ring'; $ring.Line.Visible = 0; $ring.Fill.Background()            # the slide background shows through: a window in the scrim
            $ri = 196; $d2 = $sl.Shapes.AddShape(18, [single]($cx - $ri), [single]($cy - $ri), [single](2 * $ri), [single](2 * $ri)); $d2.Adjustments.Item(1) = 0.035
            $g1 = New-Rect ($cx - 30) ($cy - $ri - 5) 60 60; $g2 = New-Rect ($cx - $ri - 5) ($cy + 40) 70 50
            $inner = Merge-Shapes 4 @($d2.Name, $g1.Name, $g2.Name); $inner.Name = 'Inner'; $inner.Line.Visible = 0; $inner.Fill.ForeColor.RGB = (Rgb $Accent)
            (Tx 'PART' 15 (Mix $Accent '#FFFFFF' 0.3) 68 150 300 -1 $Font 1 4).Name = 'Part'
            (Tx 'HEAD' 44 '#FFFFFF' 65 178 380 $heavyBold $heavy 1 0 1.08).Name = 'Head'
            (Tx 'LINE' 15 (Mix $Dark '#FFFFFF' 0.82) 68 300 320 0 $Font 1 0 1.5).Name = 'Line'
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $s = $pres.Slides($k + 1); $r = $rows[$k]
                $s.FollowMasterBackground = 0; $s.Background.Fill.UserPicture($Images[$k])
                $s.Shapes('Ring').Rotation = [single](90 * $k); $s.Shapes('Inner').Rotation = [single](360 - 120 * $k)                               # the two rings turn against each other
                $s.Shapes('Part').TextFrame.TextRange.Text = $r[0]
                $hd = $s.Shapes('Head'); $hd.TextFrame.TextRange.Text = ($(if ($r.Count -gt 1) { $r[1] } else { '' }) -replace '//', "`r")
                $ln = $s.Shapes('Line'); $ln.TextFrame.TextRange.Text = ($(if ($r.Count -gt 2) { $r[2] } else { '' }) -replace '//', "`r"); $ln.Top = [single]([double]$hd.Top + [double]$hd.Height + 16)
                if ($k -lt $n - 1) { Auto $s 1.7 }
            }
        }
        'KeywordLinks' {
            Need 1; if ($n -lt 2) { throw 'KeywordLinks needs 3-4 items' }
            $ph = Add-Cover $Images[0]; $ph.PictureFormat.Brightness = 0.42; $scrim = Box 0 0 $W $H $Dark; $scrim.Fill.Transparency = 0.4
            if ($Kicker) { $wm = Tx $Kicker 64 '#FFFFFF' 60 360 420 $heavyBold $heavy 1 0 1.0; $wm.TextFrame2.TextRange.Font.Fill.Transparency = 0.88 }
            [void](Tx $Title 34 '#FFFFFF' 65 56 400 $heavyBold $heavy 1 0 1.1); if ($Sub) { [void](Tx $Sub 14 (Mix $Dark '#FFFFFF' 0.8) 67 150 330 0 $Font 1 0 1.45) }
            Foot $true
            $gap = 420 / $n; $spineX = 470; $rr = [math]::Min(54, $gap / 2 - 8); $beats = @()
            $spine = Rule $spineX 70 $spineX 490 (Mix $Dark '#FFFFFF' 0.45) 1; $spine.Line.DashStyle = 3
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $cy = 70 + $gap * ($i + 0.5); $cx = $spineX + $(if ($i % 2) { 44 } else { -44 })
                $c = Disc $cx $cy $rr $Dark; $c.Line.Visible = -1; $c.Line.ForeColor.RGB = 0xFFFFFF; $c.Line.Weight = 1.5
                $k = Mid (Tx $r[0] 14.5 '#FFFFFF' ($cx - $rr) ($cy - $rr) (2 * $rr) -1 $Font 2 0 1.2) ($cy - $rr) (2 * $rr)
                $ln = Rule ($cx + $rr) $cy 628 $cy '#FFFFFF' 1.25; $ln.Line.EndArrowheadStyle = 6                     # a round end, as in the arrow styles
                $pw = Wd $r[1] 15; $set = @(Pill 640 ($cy - 30) $pw 30 '#FFFFFF' $r[1] 15 $Ink)
                if ($r.Count -gt 2) { $set += (Tx $r[2] 12.5 (Mix $Dark '#FFFFFF' 0.85) 644 ($cy + 8) 270 0 $Font 1 0 1.4).Name }
                $beats += , @((Grp @($c, $k)), $ln, (Grp $set))
            }
            Anim $spine $WIPE $AFTER 0.7 0.3 1
            foreach ($b in $beats) { Anim $b[0] $FADE $AFTER 0.3 0.05; Anim $b[1] $WIPE $AFTER 0.35 0 4; Anim $b[2] $FADE $WITH 0.4 0.2 }
            Hold 2
        }
        'GridCells' {
            Need 1; [void](Add-Cover $Images[0])
            $cols = 6; $rws = 5; $cw = $W / $cols; $ch = $H / $rws; $keep = @{}
            foreach ($c in 0, 1, 2) { $keep["$c,4"] = 'title' }
            $spots = @('4,0', '5,1', '5,3'); for ($i = 0; $i -lt [math]::Min(3, $n); $i++) { $keep[$spots[$i]] = "stat$i" }
            $gone = @()
            for ($r = 0; $r -lt $rws; $r++) { for ($c = 0; $c -lt $cols; $c++) {
                    $what = $keep["$c,$r"]; $cell = Box ($c * $cw) ($r * $ch) $cw $ch $(if ($what -like 'stat*') { $Accent } else { '#FFFFFF' })
                    if (-not $what) { $gone += , @($cell, ($c + $r)) }
                } }
            for ($c = 1; $c -lt $cols; $c++) { [void](Rule ($c * $cw) 0 ($c * $cw) $H '#FFFFFF' 0.75) }; for ($r = 1; $r -lt $rws; $r++) { [void](Rule 0 ($r * $ch) $W ($r * $ch) '#FFFFFF' 0.75) }
            $tt = Tx $Title 27 $Ink 28 ($ch * 4 + 12) ($cw * 3 - 50) $heavyBold $heavy 1 0 1.08
            if ($Sub) { [void](Tx $Sub 12.5 $body 30 ([double]$tt.Top + [double]$tt.Height + 8) ($cw * 3 - 60)) }
            for ($i = 0; $i -lt [math]::Min(3, $n); $i++) {
                $c, $r = $spots[$i] -split ','; $x = [int]$c * $cw; $y = [int]$r * $ch
                [void](Tx $rows[$i][0] 30 '#FFFFFF' ($x + 14) ($y + 16) ($cw - 20) $heavyBold $heavy); if ($rows[$i].Count -gt 1) { [void](Tx $rows[$i][1] 11.5 '#FFFFFF' ($x + 15) ($y + 62) ($cw - 24)) }
            }
            $first = $true
            foreach ($g in $gone) { if ($first) { $e = $sl.TimeLine.MainSequence.AddEffect($g[0], $FADE, 0, $AFTER); $e.Timing.Duration = 0.45; $e.Timing.TriggerDelayTime = 0.6; $e.Exit = -1; $first = $false } else { Leave $g[0] 0.45 (0.6 + 0.13 * $g[1]) } }
            Hold 2
        }
        'MosaicRise' {
            Need 1
            $gx = 40; $gy = 40; $cw = 162; $gap = 8; $full = 460; $cut = 0.62, 0.36, 0.7; $tiles = @(); $block = $null
            for ($c = 0; $c -lt 3; $c++) {
                $x = $gx + $c * ($cw + $gap); $h1 = [math]::Round(($full - $gap) * $cut[$c])
                foreach ($part in @(@($gy, $h1), @(($gy + $h1 + $gap), ($full - $gap - $h1)))) {
                    if ($c -eq 1 -and $part[0] -eq $gy -and $cardRows.Count) { $block = Box $x $part[0] $cw $part[1] $Accent 4; $tiles += $block; continue }
                    $tiles += (Cut $Images[0] (Box $x $part[0] $cw $part[1] $Ink 4))
                }
            }
            if ($block) {
                $num = Tx $cardRows[0][0] 40 '#FFFFFF' ([double]$block.Left + 16) ([double]$block.Top + 18) ($cw - 24) $heavyBold $heavy
                $set = @($block, $num); if ($cardRows[0].Count -gt 1) { $set += (Tx $cardRows[0][1] 12 '#FFFFFF' ([double]$block.Left + 17) ([double]$block.Top + 74) ($cw - 30) 0 $Font 1 0 1.35) }
                $bg = Grp $set; $tiles = @($tiles | ForEach-Object { if ($_.Name -eq $block.Name) { $bg } else { $_ } })
            }
            $tx = $gx + 3 * ($cw + $gap) + 34; $copy = @()
            if ($Kicker) { $copy += (Tx $Kicker 12.5 $Accent $tx 150 330 -1 $Font 1 3) }
            $tt = Tx $Title 34 $Ink $tx 174 340 $heavyBold $heavy 1 0 1.12; $copy += $tt; $y = [double]$tt.Top + [double]$tt.Height + 18
            $copy += (Box $tx $y 48 4 $Accent); $y += 22
            foreach ($r in $rows) { $b = Tx $r[0] 14.5 $body $tx $y 330 0 $Font 1 0 1.5; $copy += $b; $y += [double]$b.Height + 8 }
            Foot
            $i = 0
            foreach ($t in $tiles) { $e = $sl.TimeLine.MainSequence.AddEffect($t, 2, 0, $(if ($i) { $WITH } else { $AFTER })); $e.EffectParameters.Direction = $(if ($i % 2) { 1 } else { 3 }); $e.Timing.Duration = 0.7; $e.Timing.TriggerDelayTime = [single](0.3 + 0.12 * $i); $i++ }   # fly in, from below and from above in turn
            $first = $true; foreach ($c in $copy) { Anim $c $FADE $(if ($first) { $AFTER } else { $WITH }) 0.4 $(if ($first) { 0.1 } else { 0.2 }); $first = $false }
            Hold 2
        }
        'EraRoll' {
            if ($n -lt 2) { throw 'EraRoll needs 2-5 items' }
            $bg = Box 0 0 $W $H $Dark; $bg.Name = 'Bg'
            $size = 150; $up = "0`r1`r2`r3`r4`r5`r6`r7`r8`r9"; $down = "9`r8`r7`r6`r5`r4`r3`r2`r1`r0"; $wy = 150
            $probe = Tx $up $size '#FFFFFF' 0 0 200 $heavyBold $heavy 2 0 1.0; $lh = [double]$probe.Height / 10; $dw = $size * 0.74; $probe.Delete()
            $x0 = 65
            for ($d = 0; $d -lt 4; $d++) { (Tx $(if ($d % 2) { $down } else { $up }) $size '#FFFFFF' ($x0 + $d * $dw) $wy $dw $heavyBold $heavy 2 0 1.0).Name = "Digit$d" }
            (Box 0 -900 $W (900 + $wy) $Dark).Name = 'MaskTop'; (Box 0 ($wy + $lh) $W 1400 $Dark).Name = 'MaskBottom'
            [void](Tx $Title 15 '#FFFFFF' 68 56 600 -1 $Font 1 3)
            (Box 68 ($wy + $lh + 6) 80 5 $Accent).Name = 'Bar'
            (Tx 'HEAD' 30 '#FFFFFF' 520 ($wy + 26) 380 $heavyBold $heavy 1 0 1.1).Name = 'Head'; (Tx 'LINE' 15 '#FFFFFF' 522 ($wy + 80) 370 0 $Font 1 0 1.5).Name = 'Line'
            $ty = 462; [void](Rule 68 $ty ($W - 68) $ty '#FFFFFF' 1)
            for ($i = 0; $i -lt $n; $i++) { $x = 68 + $i * (($W - 136) / ($n - 1)); (Disc $x $ty 5 '#FFFFFF').Name = "Dot$i"; [void](Tx $rows[$i][0] 11.5 '#FFFFFF' ($x - 40) ($ty + 14) 80 0 $Font 2) }
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $s = $pres.Slides($k + 1); $r = $rows[$k]; $yr = $r[0].Trim(); if ($yr -notmatch '^\d{4}$') { throw "EraRoll: '$yr' is not a four-digit year" }
                $tone = Mix $Dark $AccentDark (0.75 * $k / ($n - 1))
                foreach ($nm in 'Bg', 'MaskTop', 'MaskBottom') { $s.Shapes($nm).Fill.ForeColor.RGB = (Rgb $tone) }
                for ($d = 0; $d -lt 4; $d++) { $v = [int]$yr.Substring($d, 1); $s.Shapes("Digit$d").Top = [single]($wy - $(if ($d % 2) { 9 - $v } else { $v }) * $lh) }   # odd columns are written 9..0: they roll the other way
                $s.Shapes('Head').TextFrame.TextRange.Text = $(if ($r.Count -gt 1) { $r[1] } else { '' }); $s.Shapes('Line').TextFrame.TextRange.Text = ($(if ($r.Count -gt 2) { $r[2] } else { '' }) -replace '//', "`r")
                $dot = $s.Shapes("Dot$k"); $cxd = [double]$dot.Left + 5; $dot.Width = 22; $dot.Height = 22; $dot.Left = [single]($cxd - 11); $dot.Top = [single]($ty - 11); $dot.Fill.ForeColor.RGB = (Rgb $Accent)
                $s.Shapes('Bar').Width = [single](80 + 60 * $k)
                if ($k -lt $n - 1) { Auto $s 1.5 }
            }
        }
        'ArchColumns' {
            if ($n -lt 3) { throw 'ArchColumns needs 3-5 items' }; Need $n
            $pal = @('#F8C8C0', '#FBE3E8', '#FBF0A8', '#DDEBD8', '#E3E6F3'); if ($cardRows.Count) { $pal = @($cardRows | ForEach-Object { $_[0] }) }
            $thin = 96; $wide = $W - ($n - 1) * $thin; $bw = 300; $bh = 400
            for ($i = 0; $i -lt $n; $i++) {
                (Box 0 0 $wide $H $pal[$i % $pal.Count]).Name = "Col$i"
                $p = Fit $Images[$i] 0 0 $bw $bh 5; $p.Adjustments.Item(1) = 0.5; $p.Name = "Pic$i"
                (Tx $rows[$i][0] 13 $Ink 0 34 $thin -1 $Font 2 0 1.2).Name = "Lab$i"
                $set = @((Tx $rows[$i][1] 34 $Ink 0 0 300 $heavyBold $heavy 1 0 1.1)); $hgt = [double]$set[0].Height
                if ($rows[$i].Count -gt 2) { $set += (Tx $rows[$i][2] 14 $body 2 ($hgt + 14) 290 0 $Font 1 0 1.5) }
                (Grp $set).Name = "Txt$i"
            }
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $s = $pres.Slides($k + 1); $x = 0
                for ($i = 0; $i -lt $n; $i++) {
                    $open = ($i -eq $k); $cwid = $(if ($open) { $wide } else { $thin })
                    $col = $s.Shapes("Col$i"); $col.Left = [single]$x; $col.Width = [single]$cwid
                    $lab = $s.Shapes("Lab$i"); $lab.Left = [single]$x
                    $pic = $s.Shapes("Pic$i"); $txt = $s.Shapes("Txt$i")
                    if ($open) { $pic.Left = [single]($x + $cwid - $bw - 46); $pic.Top = [single](($H - $bh) / 2 + 20); $txt.Left = [single]($x + 48); $txt.Top = [single](($H - [double]$txt.Height) / 2 + 10) }
                    else {
                        $c = $pic.PictureFormat.Crop; $c.ShapeLeft = [single]([double]$c.ShapeLeft + ($bw - 64) / 2); $c.ShapeWidth = 64; $c.ShapeTop = [single]([double]$c.ShapeTop + ($bh - 210) / 2); $c.ShapeHeight = 210   # the window narrows around the middle of the photo
                        $pic.Left = [single]($x + ($thin - 64) / 2); $pic.Top = [single]($H - 210 - 60); $txt.Delete()
                    }
                    $x += $cwid
                }
                if ($k -lt $n - 1) { Auto $s 1.7 }
            }
        }
        'LensTour' {
            Need 1; if ($n -lt 2) { throw 'LensTour needs 2-4 items' }
            (Add-Cover $Images[0]).Name = 'Base'; $scrim = Box 0 0 $W $H $Dark; $scrim.Fill.Transparency = 0.45; $scrim.Name = 'Scrim'
            $pt = @($rows | ForEach-Object { , @(([double]$_[0] * $W), ([double]$_[1] * $H)) })
            for ($i = 0; $i -lt $n - 1; $i++) { $l = Rule $pt[$i][0] $pt[$i][1] $pt[$i + 1][0] $pt[$i + 1][1] '#FFFFFF' 1.5; $l.Line.DashStyle = 4; $l.Line.Transparency = 0.3 }
            for ($i = 0; $i -lt $n; $i++) { $dd = Disc $pt[$i][0] $pt[$i][1] 6 '#FFFFFF'; $dd.Line.Visible = -1; $dd.Line.ForeColor.RGB = (Rgb $Accent); $dd.Line.Weight = 2.5 }
            if ($Kicker) { [void](Tx $Kicker 12.5 '#FFFFFF' 65 40 500 -1 $Font 1 3) }; if ($Title) { [void](Tx $Title 28 '#FFFFFF' 65 62 600 $heavyBold $heavy) }
            $mx = ($pt | ForEach-Object { $_[0] } | Measure-Object -Average).Average; $my = ($pt | ForEach-Object { $_[1] } | Measure-Object -Average).Average
            $kx = $(if ($mx -lt $W / 2) { $W - 395 } else { 65 }); $ky = $(if ($my -gt $H / 2) { 124 } else { 400 })                 # the card keeps to the corner away from the spots
            (Box $kx $ky 330 96 '#FFFFFF' 10).Name = 'Card'; (Tx 'HEAD' 19 $Ink ($kx + 20) ($ky + 16) 290 -1).Name = 'Head'; (Tx 'LINE' 12.5 $body ($kx + 20) ($ky + 46) 290 0 $Font 1 0 1.4).Name = 'Line'
            Steps $n
            $zoom = 2.2; $rad = 92
            for ($k = 0; $k -lt $n; $k++) {
                On ($pres.Slides($k + 1)); $r = $rows[$k]
                $sl.Shapes('Head').TextFrame.TextRange.Text = $r[2]; $sl.Shapes('Line').TextFrame.TextRange.Text = ($(if ($r.Count -gt 3) { $r[3] } else { '' }) -replace '//', "`r")
                for ($j = 0; $j -le $k; $j++) {                                                           # the lens at this stop, and a small snapshot of every stop so far
                    $isLens = ($j -eq $k); $z = $(if ($isLens) { $zoom } else { 1.5 }); $rr = $(if ($isLens) { $rad } else { 28 })
                    $cx = $(if ($isLens) { $pt[$j][0] } else { $W - 60 - ($n - 1 - $j) * 66 }); $cy = $(if ($isLens) { $pt[$j][1] } else { 62 })
                    $p = Add-Cover $Images[0]; $bx = [double]$p.Left; $by = [double]$p.Top; $pw = [double]$p.Width; $phh = [double]$p.Height
                    $fx = ($pt[$j][0] - $bx) / $pw; $fy = ($pt[$j][1] - $by) / $phh                       # the spot as a share of the picture
                    $p.Width = [single]($pw * $z); $p.Left = [single]($cx - $fx * $pw * $z); $p.Top = [single]($cy - $fy * $phh * $z)
                    $c = $p.PictureFormat.Crop; $c.ShapeLeft = [single]($cx - $rr); $c.ShapeTop = [single]($cy - $rr); $c.ShapeWidth = [single](2 * $rr); $c.ShapeHeight = [single](2 * $rr)
                    $p.AutoShapeType = 9; $p.Line.Visible = -1; $p.Line.ForeColor.RGB = 0xFFFFFF; $p.Line.Weight = [single]$(if ($isLens) { 5 } else { 2.5 })
                    $p.Name = $(if ($isLens) { '!!Lens' } else { "!!Snap$j" }); if ($isLens) { Glow $p '#000000' 0.45 24 8 }
                }
                if ($k -lt $n - 1) { Auto $sl 1.8 }
            }
        }
        'DiamondBricks' {
            Need 1
            $bw = 118; $bh = 78; $gap = 9; $ox = 270; $oy = 270; $cs = [math]::Cos([math]::PI / 4); $list = @()
            for ($r = -3; $r -le 3; $r++) { for ($c = -3; $c -le 3; $c++) {
                    $ux = $c * ($bw + $gap) + $(if ($r % 2) { ($bw + $gap) / 2 } else { 0 }); $uy = $r * ($bh + $gap)
                    $cx = $ox + ($ux - $uy) * $cs; $cy = $oy + ($ux + $uy) * $cs                          # the wall turned 45 degrees about its middle
                    if ($cx -lt -20 -or $cx -gt 470 -or $cy -lt -30 -or $cy -gt $H + 30) { continue }
                    $pts = @(); foreach ($q in @(@(-1, -1), @(1, -1), @(1, 1), @(-1, 1))) { $px = $q[0] * $bw / 2; $py = $q[1] * $bh / 2; $pts += , @(($cx + ($px - $py) * $cs), ($cy + ($px + $py) * $cs)) }
                    $list += , @($pts, $cx, $cy)
                } }
            $colour = @{ 2 = $Accent; 7 = (Mix $Accent '#FFFFFF' 0.45); 11 = $AccentDark }; $bricks = @(); $figure = $null; $i = 0
            foreach ($b in $list) {
                if ($colour.ContainsKey($i)) { $s = Poly $b[0] $colour[$i]; if ($i -eq 7) { $figure = $b } } else { $s = Cut $Images[0] (Poly $b[0] $Ink) }
                $bricks += , @($s, ($b[1] + $b[2])); $i++
            }
            $extra = @()
            if ($figure -and $cardRows.Count) { $extra += (Tx $cardRows[0][0] 24 $Ink ($figure[1] - 50) ($figure[2] - 26) 100 $heavyBold $heavy 2); if ($cardRows[0].Count -gt 1) { $extra += (Tx $cardRows[0][1] 10 $Ink ($figure[1] - 50) ($figure[2] + 6) 100 0 $Font 2) } }
            [void](Poly @(@(($W - 70), $H), @($W, ($H - 70)), @($W, $H)) $Accent); [void](Poly @(@(($W - 150), $H), @(($W - 80), $H), @($W, ($H - 80)), @($W, ($H - 150))) (Mix $Accent '#FFFFFF' 0.55))
            if ($Kicker) { $extra += (Tx $Kicker 12.5 $Accent 540 168 340 -1 $Font 1 3) }
            $tt = Tx $Title 40 $Ink 540 192 360 $heavyBold $heavy 1 0 1.1; $extra += $tt
            if ($Sub) { $extra += (Tx $Sub 14 $body 542 ([double]$tt.Top + [double]$tt.Height + 16) 340 0 $Font 1 0 1.55) }
            Foot
            $far = ($bricks | ForEach-Object { $_[1] } | Measure-Object -Minimum -Maximum); $first = $true
            foreach ($b in ($bricks | Sort-Object { $_[1] })) { Anim $b[0] $FADE $(if ($first) { $AFTER } else { $WITH }) 0.4 (0.3 + 1.4 * ($b[1] - $far.Minimum) / [math]::Max(1, $far.Maximum - $far.Minimum)); $first = $false }
            $first = $true; foreach ($x in $extra) { Anim $x $FADE $(if ($first) { $AFTER } else { $WITH }) 0.4 0.15; $first = $false }
            Hold 2
        }
        'PaperPeel' {
            $sky = Mix $Accent '#FFFFFF' 0.72; $bg = Box 0 0 $W $H $sky
            $tt = Tx $Title 96 $AccentDark 0 0 ($W - 100) $heavyBold $heavy 2 0 1.0; $tt.Left = 50; $tt.Top = [single](($H - [double]$tt.Height) / 2 - 10)
            $sh = $tt.TextFrame2.TextRange.Font.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = (Rgb $Dark); $sh.Transparency = 0.6; $sh.Blur = 8; $sh.OffsetX = 4; $sh.OffsetY = 5
            $late = @(); if ($Kicker) { $late += (Tx $Kicker 14 $AccentDark 50 ([double]$tt.Top - 26) ($W - 100) -1 $Font 2 4) }; if ($Sub) { $late += (Tx $Sub 18 $Ink 50 ([double]$tt.Top + [double]$tt.Height + 10) ($W - 100) 0 $Font 2) }
            $curve = { param($off) $pts = @(); for ($x = -20; $x -le $W + 20; $x += 20) { $t = [math]::Min(1.0, [math]::Max(0.0, [double]$x / $W)); $pts += , @($x, ($off + 330 * $t * $t * (3 - 2 * $t))) }; , $pts }
            $up = & $curve 150; $lo = & $curve 60; [array]::Reverse($up)
            $mid = Poly ((& $curve 20) + $up) (Mix $Accent '#FFFFFF' 0.35)                                                       # the band between the two sheets, at the back
            $top = Poly ((& $curve 110) + @(, @(($W + 20), -20)) + @(, @(-20, -20))) (Mix $Accent '#FFFFFF' 0.55)               # everything above its curve
            $bot = Poly ($lo + @(, @(($W + 20), ($H + 20))) + @(, @(-20, ($H + 20)))) $Accent                                    # everything below its curve
            foreach ($s in $mid, $top, $bot) { Glow $s '#000000' 0.55 18 0; Show $s }
            $go = { param($shape, $path, $delay, $trig) $e = $sl.TimeLine.MainSequence.AddEffect($shape, 0, 0, $trig); $e.Behaviors.Add(1).MotionEffect.Path = $path; $e.Timing.Duration = 1.7; $e.Timing.TriggerDelayTime = [single]$delay }
            & $go $bot 'M 0 0 L -0.45 1.05 E' 0.6 $AFTER; & $go $top 'M 0 0 L 0.45 -1.05 E' 0.25 $WITH; & $go $mid 'M 0 0 L 1.1 0.35 E' 0.5 $WITH
            $first = $true; foreach ($x in $late) { Anim $x $FADE $(if ($first) { $AFTER } else { $WITH }) 0.5 0; $first = $false }
            Hold 2
        }
        'TornReveal' {
            Need 1; (Add-Cover $Images[0]).Name = 'Photo'
            $rnd = New-Object System.Random 7; $edge = @(); for ($y = -12; $y -le $H + 12; $y += 15) { $edge += , @((575 + 0.12 * $y + $rnd.Next(-13, 14)), $y) }
            $rev = @($edge); [array]::Reverse($rev)
            $mk = { param($pts, [string]$hex, [string]$name, $dx) $s = Poly $pts $hex; $s.Left = [single]([double]$s.Left + $dx); $s.Name = $name; $s }
            $lp = $edge + @(, @(-520, ($H + 12))) + @(, @(-520, -12)); $rp = $rev + @(, @(($W + 620), -12)) + @(, @(($W + 620), ($H + 12)))
            [void](& $mk $lp '#F4EFE6' 'LeftEdge' 4); [void](& $mk $rp '#F4EFE6' 'RightEdge' -4)                                # a pale rim: the torn paper's thickness
            $l = & $mk $lp $Dark 'Left' 0; $r = & $mk $rp $Dark 'Right' 0; Glow $l '#000000' 0.5 14 0; Glow $r '#000000' 0.5 14 0
            $set = @(); if ($Kicker) { $set += (Tx $Kicker 12.5 (Mix $Accent '#FFFFFF' 0.25) 65 150 400 -1 $Font 1 3) }
            $tt = Tx $Title 40 '#FFFFFF' 65 176 420 $heavyBold $heavy 1 0 1.1; $set += $tt; $y = [double]$tt.Top + [double]$tt.Height + 16; $set += (Box 67 $y 50 4 $Accent); $y += 22
            foreach ($row in $rows) { $b = Tx $row[0] 14.5 (Mix $Dark '#FFFFFF' 0.85) 67 $y 380 0 $Font 1 0 1.5; $set += $b; $y += [double]$b.Height + 8 }
            (Grp $set).Name = 'Copy'
            Steps 3
            $move = { param($slide, $dl, $dr) foreach ($nm in 'Left', 'LeftEdge') { $q = $slide.Shapes($nm); $q.Left = [single]([double]$q.Left + $dl) }; foreach ($nm in 'Right', 'RightEdge') { $q = $slide.Shapes($nm); $q.Left = [single]([double]$q.Left + $dr) } }
            $s1 = $pres.Slides(1); $s1.Shapes('Copy').Delete(); & $move $s1 3 -3; Auto $s1 0.9                                  # a hairline of the photo
            & $move $pres.Slides(2) -70 120; Auto $pres.Slides(2) 1.4; & $move $pres.Slides(3) -70 470
        }
        'FoldArrows' {
            if ($n -lt 3) { throw 'FoldArrows needs 3-5 items' }
            Head; Foot
            $aw = ($W - 130) / $n; $ah = 92; $ay = 236; $beats = @()
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $x = 65 + $i * $aw; $last = ($i -eq $n - 1); $lift = $(if ($last) { -12 } else { 0 })
                $sd = Tag ($sl.Shapes.AddShape(9, [single]($x - 14), [single]($ay - 26 + $lift), 30, [single]($ah + 52))); $sd.Line.Visible = 0; $sd.Fill.ForeColor.RGB = 0; $sd.Fill.Transparency = 0.5; $sd.Rotation = 24; $sd.SoftEdge.Radius = 9   # the fold: a soft dark blade under the left edge
                $ar = Tag ($sl.Shapes.AddShape(33, [single]$x, [single]($ay + $lift), [single]($aw + 22), [single]$ah)); $ar.Line.Visible = 0; $ar.Fill.ForeColor.RGB = (Rgb $(if ($last) { $AccentDark } else { Mix $Accent '#FFFFFF' (0.3 - 0.3 * $i / [math]::Max(1, $n - 2)) }))
                $ar.Adjustments.Item(1) = 0.62; $ar.Adjustments.Item(2) = 0.42
                $yr = Mid (Tx $r[0] 21 '#FFFFFF' $x ($ay + $lift) $aw $heavyBold $heavy 2) ($ay + $lift) $ah
                $txt = @(); if ($r.Count -gt 1) { $txt += (Tx $r[1] 15 $Ink ($x + 6) ($ay - 62 + $lift) ($aw - 4) -1 $Font 2) }; if ($r.Count -gt 2) { $txt += (Tx $r[2] 11.5 $body ($x + 10) ($ay + $ah + 22) ($aw - 12) 0 $Font 2 0 1.45) }
                $beats += , @($sd, (Grp @($ar, $yr)), $txt)
            }
            foreach ($b in $beats) { Anim $b[1] $WIPE $AFTER 0.45 0.1 4; Anim $b[0] $FADE $WITH 0.3 0.3; foreach ($t in $b[2]) { Anim $t $FADE $WITH 0.35 0.3 } }
            Hold 2
        }
    }
    Save-Deck $Out $Preview 2.5
    "saved $Out"
} catch { Close-Deck; throw }
