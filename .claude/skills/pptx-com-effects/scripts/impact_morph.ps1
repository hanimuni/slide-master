# More Morph scenes (second set; same colour / font parameters and helpers as scene_page.ps1 and impact_scene.ps1):
#   Flip        cards that turn over one by one: question on the front, answer on the back     -Items 'front head;front line;back head;back line|...' (2-4)
#   Parallax    panels scroll sideways while three depth layers move at different speeds       -Items 'head;line|...' (dark page)
#   CutAway     a divider sweeps across two pictures of the same frame (before / after)        -Images before.png after.png  -Items 'BEFORE|AFTER'
#   Accordion   photo strips; the current one stretches wide, the others stay narrow           -Images a.jpg b.jpg ...  -Items 'head;line|...' (one per photo)
#   Wave        two wave-edged panels part and uncover the title (and a photo if given)        -Title 'a|b'  -Sub line  [-Images photo.jpg]
#   Orbit       pictures on a ring; the current one comes to the front and grows               -Images one per item  -Items 'head;line|...'
#   ArcList     a list riding the edge of a big circle; the current line is large              -Items 'head;line;line|...'
#   StatSeries  one huge number per slide, the page colour changes with every number           -Items 'big;unit;line|...'
# Flip, CutAway and Wave play through by themselves after the first click; the others wait for a click at every step.
# Example: .\impact_morph.ps1 -Layout StatSeries -Items '6;people;use it every day|55;coupons;sent in five weeks' -Out C:\work\s.pptx -Preview C:\work\s.mp4
param(
    [Parameter(Mandatory)][ValidateSet('Flip', 'Parallax', 'CutAway', 'Accordion', 'Wave', 'Orbit', 'ArcList', 'StatSeries')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Note = '',
    [string]$Accent = '#7C6BFF',
    [string]$AccentDark = '',
    [string]$Light = '#FCEFE6',
    [string]$Dark = '#241C45',
    [string]$Ink = '#1F1B2E',
    [string]$Font = 'Malgun Gothic',
    [string]$HeavyFont = '',
    [double]$Duration = 0.7,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent
. (Join-Path $PSScriptRoot 'scene_helpers.ps1')
function Steps([int]$count) { for ($i = 1; $i -lt $count; $i++) { Add-MorphSlide $Duration } }      # build everything on slide 1, then copy it
function Auto($slide, $sec) { $slide.SlideShowTransition.AdvanceOnTime = -1; $slide.SlideShowTransition.AdvanceTime = [single]$sec }
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Place($shape, $x, $y) { $shape.Left = [single]$x; $shape.Top = [single]$y }
function Dots($slide, $n, $k, $cx, [string]$onHex, [string]$offHex) {                                # progress dots; the current one is a short bar
    for ($i = 0; $i -lt $n; $i++) {
        $d = $slide.Shapes("Dot$i"); $d.Width = [single]$(if ($i -eq $k) { 28 } else { 8 }); $d.Fill.ForeColor.RGB = (Rgb $(if ($i -eq $k) { $onHex } else { $offHex }))
        $d.Left = [single]($cx - (16 * $n + 20) / 2 + 16 * $i + $(if ($i -gt $k) { 20 } else { 0 }))
    }
}
function MakeDots($n, $y, [string]$hex) { for ($i = 0; $i -lt $n; $i++) { $d = Box 0 $y 8 8 $hex 4; $d.Name = "Dot$i" } }

try {
    $darkPage = $Layout -in 'Parallax', 'CutAway'
    New-Deck $(if ($darkPage) { $Dark } elseif ($Layout -in 'Flip', 'Orbit') { '#FFFFFF' } else { $Light })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count

    switch ($Layout) {
        'Flip' {
            [void](Disc 960 560 250 $Light); Head
            $gap = 24; $cw = (830 - $gap * ($n - 1)) / $n; $ch = 300; $y = 152
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $x = 65 + $i * ($cw + $gap); $mid = $x + $cw / 2
                $bk = Box ($mid - 1) $y 2 $ch $Accent 20; $bk.Name = "B$i"; $bk.Fill.TwoColorGradient(1, 1); $bk.Fill.ForeColor.RGB = (Rgb (Mix $Accent '#FFFFFF' 0.12)); $bk.Fill.BackColor.RGB = (Rgb $AccentDark); $bk.Fill.GradientAngle = 315
                $fr = Box $x $y $cw $ch '#FFFFFF' 20; Glow $fr (Mix $Accent '#FFFFFF' 0.45) 0.72 20 7; $fr.Name = "F$i"
                (Tx ('{0:00}' -f ($i + 1)) 44 $Accent ($x + 26) ($y + 26) 120 $heavyBold $heavy).Name = "FN$i"
                (Tx $r[0] 25 $Ink ($x + 26) ($y + 112) ($cw - 50) -1 $Font 1 0 1.2).Name = "FH$i"
                (Tx $r[1] 17 $body ($x + 26) ($y + 190) ($cw - 50) 0 $Font 1 0 1.3).Name = "FL$i"
                (Tx ([string][char]0x2192) 26 $Accent ($x + $cw - 60) ($y + $ch - 56) 36 -1 $Font 3).Name = "FA$i"
                (Tx $r[2] 25 '#FFFFFF' ($x + 26) ($y + 84) ($cw - 50) -1 $Font 1 0 1.25).Name = "BH$i"
                (Tx $r[3] 18 $tint ($x + 26) ($y + 176) ($cw - 50) 0 $Font 1 0 1.35).Name = "BL$i"
            }
            $Duration = 0.25; Steps (1 + 2 * $n)
            for ($s = 0; $s -le 2 * $n; $s++) {
                $slide = $pres.Slides($s + 1)
                for ($i = 0; $i -lt $n; $i++) {
                    $x = 65 + $i * ($cw + $gap); $mid = $x + $cw / 2
                    $phase = $(if ($s -ge 2 + 2 * $i) { 2 } elseif ($s -eq 1 + 2 * $i) { 1 } else { 0 })      # 0 front, 1 edge-on, 2 back
                    if ($phase -ne 2) { foreach ($nm in 'BH', 'BL') { $slide.Shapes("$nm$i").Delete() } }
                    if ($phase -ne 0) { foreach ($nm in 'FN', 'FH', 'FL', 'FA') { $slide.Shapes("$nm$i").Delete() } }
                    if ($phase -eq 1) { $f = $slide.Shapes("F$i"); $f.Width = 2; $f.Left = [single]($mid - 1) }
                    if ($phase -eq 2) { $slide.Shapes("F$i").Delete(); $b = $slide.Shapes("B$i"); $b.Width = [single]$cw; $b.Left = [single]$x; $b.Adjustments.Item(1) = [single](20 / [math]::Min($cw, $ch)); Glow $b $Accent 0.55 22 8 }
                }
                if ($s -ge 1 -and $s -lt 2 * $n) { Auto $slide $(if ($s % 2) { 0.01 } else { 0.9 }) }       # after the first click the cards turn by themselves
            }
        }
        'Parallax' {
            $far = @(); $midL = @(); $near = @(); $span = $W * ($n - 1)
            for ($j = 0; $j -lt 3 + $n; $j++) { $d = Disc 0 0 (180 + 40 * ($j % 3)) $dDeco; $d.Name = "Far$j"; $far += , @((120 + $j * 420), (90 + 170 * ($j % 3))) }
            for ($j = 0; $j -lt 4 + 2 * $n; $j++) { $d = Disc 0 0 (34 + 16 * ($j % 3)) $dCard; $d.Name = "Mid$j"; $midL += , @((200 + $j * 330), (110 + 150 * (($j * 2) % 3))) }
            for ($j = 0; $j -lt 6 + 3 * $n; $j++) { $d = Disc 0 0 (5 + 3 * ($j % 3)) $(if ($j % 2) { $Accent } else { $soft }); $d.Name = "Near$j"; $near += , @((90 + $j * 290), (70 + 105 * (($j * 3) % 5))) }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]
                $num = Tx ('{0:00}' -f ($i + 1)) 150 (Mix $Dark $Accent 0.55) 0 0 300 $heavyBold $heavy 1 -4; $num.TextFrame.WordWrap = 0; $num.Name = "PNum$i"
                $hd = Tx $r[0] 64 '#FFFFFF' 0 0 760 -1; $hd.TextFrame.WordWrap = 0; $hd.Name = "PHead$i"
                if ($r.Count -gt 1) { (Tx (($r[1..($r.Count - 1)]) -join "`r") 24 $dText 0 0 700 0 $Font 1 0 1.3).Name = "PLine$i" }
            }
            if ($Kicker) { [void](Tx $Kicker 12.5 $soft 65 45 600 -1 $Font 1 3) }; $rail = Box 65 492 830 3 $dCard 1.5; $bar = Box 65 491 (830 / $n) 5 $Accent 2.5; $bar.Name = 'Bar'
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $slide = $pres.Slides($k + 1)
                for ($j = 0; $j -lt $far.Count; $j++) { $d = $slide.Shapes("Far$j"); Place $d ($far[$j][0] - $k * $W * 0.18 - [double]$d.Width / 2) ($far[$j][1] - [double]$d.Height / 2) }
                for ($j = 0; $j -lt $midL.Count; $j++) { $d = $slide.Shapes("Mid$j"); Place $d ($midL[$j][0] - $k * $W * 0.5 - [double]$d.Width / 2) ($midL[$j][1] - [double]$d.Height / 2) }
                for ($j = 0; $j -lt $near.Count; $j++) { $d = $slide.Shapes("Near$j"); Place $d ($near[$j][0] - $k * $W * 1.35) $near[$j][1] }
                for ($i = 0; $i -lt $n; $i++) {
                    $ox = ($i - $k) * $W
                    Place $slide.Shapes("PNum$i") (110 + $ox) 96; Place $slide.Shapes("PHead$i") (116 + $ox) 258
                    if ($rows[$i].Count -gt 1) { Place $slide.Shapes("PLine$i") (120 + $ox) 352 }
                }
                $slide.Shapes('Bar').Left = [single](65 + $k * 830 / $n)
            }
        }
        'CutAway' {
            if ($Images.Count -lt 2) { throw 'CutAway needs -Images <before> <after>' }
            $labels = @($Items -split '\|'); if ($labels.Count -lt 2) { $labels = 'BEFORE', 'AFTER' }
            [void](Add-Cover $Images[1]); $top = Add-Cover $Images[0]; $top.Name = 'Before'; $fullW = [double]$top.Width; $leftEdge = [double]$top.Left
            $bar = Box 0 0 4 $H '#FFFFFF'; $bar.Name = 'Divider'
            $knob = Disc 0 ($H / 2) 24 '#FFFFFF'; Glow $knob $Dark 0.6 14 3; $arrows = Mid (Tx ([string][char]0x2039 + ' ' + [char]0x203A) 20 $AccentDark 0 ($H / 2 - 24) 48 -1 $Font 2) ($H / 2 - 24) 48
            (Grp @($knob, $arrows)).Name = 'Knob'
            (Grp (Pill 40 36 (Wd $labels[0] 15) 34 $Dark $labels[0] 15 '#FFFFFF')).Name = 'TagA'
            (Grp (Pill ($W - 40 - (Wd $labels[1] 15)) 36 (Wd $labels[1] 15) 34 $Accent $labels[1] 15 '#FFFFFF')).Name = 'TagB'
            $Duration = 1.6; Steps 2
            $cuts = 0.86, 0.14
            for ($k = 0; $k -lt 2; $k++) {
                $slide = $pres.Slides($k + 1); $cutX = $W * $cuts[$k]; $vis = $cutX - $leftEdge
                $c = $slide.Shapes('Before').PictureFormat.Crop; $c.ShapeWidth = [single]$vis; $c.ShapeLeft = [single]$leftEdge; $c.PictureOffsetX = [single](($fullW - $vis) / 2)   # the picture stays put, only its frame narrows
                $slide.Shapes('Divider').Left = [single]($cutX - 2); $slide.Shapes('Knob').Left = [single]($cutX - 24)
            }
            Auto $pres.Slides(1) 0.8
        }
        'Accordion' {
            if (-not $Images.Count) { throw 'Accordion needs -Images' }
            [void](Disc 900 20 250 $lilac); Head
            $n = $Images.Count; $narrow = 86; $gap = 10; $wide = 830 - ($n - 1) * ($narrow + $gap); $top = 124; $hgt = 318
            for ($i = 0; $i -lt $n; $i++) {
                $pic = $sl.Shapes.AddPicture($Images[$i], 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = "Strip$i"
                $pic.Height = [single]$hgt; if ([double]$pic.Width -lt $wide) { $pic.Width = [single]$wide }      # large enough to fill the wide state
                $c = $pic.PictureFormat.Crop; $c.ShapeHeight = [single]$hgt; $c.ShapeWidth = [single]$narrow; $pic.Top = [single]$top
                if ($i -lt $rows.Count) {
                    $r = $rows[$i]
                    $cap = @((Tx ('{0:00}' -f ($i + 1)) 40 $Accent 65 456 80 $heavyBold $heavy), (Tx $r[0] 28 $Ink 140 458 380 -1))
                    if ($r.Count -gt 1) { $cap += (Tx (($r[1..($r.Count - 1)]) -join ' ') 17 $body 142 498 700) }
                    (Grp $cap).Name = "Cap$i"
                }
            }
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $slide = $pres.Slides($k + 1); $x = 65
                for ($i = 0; $i -lt $n; $i++) {
                    $sw = $(if ($i -eq $k) { $wide } else { $narrow }); $pic = $slide.Shapes("Strip$i")
                    $pic.PictureFormat.Crop.ShapeWidth = [single]$sw; $pic.Left = [single]$x; $pic.Top = [single]$top; $x += $sw + $gap
                    if ($i -ne $k -and $i -lt $rows.Count) { $slide.Shapes("Cap$i").Delete() }
                }
            }
        }
        'Wave' {
            $photo = $Images.Count -gt 0
            if ($photo) {
                [void](Add-Cover $Images[0]); $sc = Box 0 0 $W $H $Dark; $sc.Fill.TwoColorGradient(1, 1); $sc.Fill.GradientAngle = 270
                $gs = $sc.Fill.GradientStops; $gs.Item(1).Color.RGB = (Rgb $Dark); $gs.Item(1).Transparency = 0.15; $gs.Item(2).Color.RGB = (Rgb $Dark); $gs.Item(2).Position = 0.7; $gs.Item(2).Transparency = 1
            } else { [void](Disc ($W / 2) ($H / 2) 250 $lilac) }
            $ty = $(if ($photo) { 300 } else { 150 }); $tc = $(if ($photo) { '#FFFFFF' } else { $Ink })
            if ($Kicker) { [void](Tx $Kicker 13 $(if ($photo) { '#FFFFFF' } else { $Accent }) 0 ($ty - 34) $W -1 $Font 2 3) }
            $hd = Tx $Title 66 $tc 0 $ty $W -1 $Font 2 0 1.12
            if ($Sub) { [void](Tx $Sub 22 $(if ($photo) { $tint } else { $body }) 0 ([double]$hd.Top + [double]$hd.Height + 10) $W 0 $Font 2) }
            $mid = $W / 2; $amp = 34; $pts = @(); for ($yy = 0; $yy -le $H; $yy += 18) { $pts += , @(($mid + $amp * [math]::Sin($yy / $H * 3 * [math]::PI)), $yy) }
            foreach ($side in 'L', 'R') {
                $edge = $(if ($side -eq 'L') { -30 } else { $W + 30 })
                $fb = $sl.Shapes.BuildFreeform(0, [single]$edge, -20); $fb.AddNodes(0, 0, [single]$pts[0][0], -20)
                foreach ($p in $pts) { $fb.AddNodes(0, 0, [single]$p[0], [single]$p[1]) }
                $fb.AddNodes(0, 0, [single]$pts[-1][0], [single]($H + 20)); $fb.AddNodes(0, 0, [single]$edge, [single]($H + 20)); $fb.AddNodes(0, 0, [single]$edge, -20)
                $pn = $fb.ConvertToShape(); $pn.Line.Visible = 0; $pn.Fill.ForeColor.RGB = (Rgb $(if ($side -eq 'L') { $Accent } else { $AccentDark })); $pn.Name = "Panel$side"
            }
            $Duration = 1.2; Steps 2
            $s2 = $pres.Slides(2); $shift = $W / 2 - 26                                                # the panels stay as thin wavy edges at both sides
            $s2.Shapes('PanelL').Left = [single]([double]$s2.Shapes('PanelL').Left - $shift); $s2.Shapes('PanelR').Left = [single]([double]$s2.Shapes('PanelR').Left + $shift)
            Auto $pres.Slides(1) 0.7
        }
        'Orbit' {
            [void](Disc 960 560 250 $Light); Head
            $cx = 300; $cy = 300; $rad = 165
            $ring = Disc $cx $cy $rad $Accent; $ring.Fill.Visible = 0; $ring.Line.Visible = -1; $ring.Line.ForeColor.RGB = (Rgb (Mix $Accent '#FFFFFF' 0.6)); $ring.Line.Weight = 2; $ring.Line.DashStyle = 4
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $parts = @((Disc 0 0 44 '#FFFFFF'))
                if ($i -lt $Images.Count) { $parts += (Pic $Images[$i] -40 -40 80) } else { $parts += (Disc 0 0 40 $tint); $parts += (Mid (Tx "$($i + 1)" 30 $AccentDark -40 -40 80 -1 $Font 2) -40 80) }
                $orb = Grp $parts; $orb.LockAspectRatio = -1; $orb.Name = "Orb$i"; Glow $orb (Mix $Accent '#FFFFFF' 0.4) 0.65 16 5
                $cap = @((Tx ('{0:00}' -f ($i + 1)) 60 $Accent 560 180 200 $heavyBold $heavy), (Tx $r[0] 40 $Ink 560 262 340 -1))
                if ($r.Count -gt 1) { $cap += (Tx (($r[1..($r.Count - 1)]) -join "`r") 20 $body 562 326 330 0 $Font 1 0 1.35) }
                (Grp $cap).Name = "Cap$i"
            }
            MakeDots $n 478 (Mix $Accent '#FFFFFF' 0.7)
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $slide = $pres.Slides($k + 1)
                for ($i = 0; $i -lt $n; $i++) {
                    $a = ($i - $k) * 2 * [math]::PI / $n; $d = $(if ($i -eq $k) { 156 } else { 74 }); $o = $slide.Shapes("Orb$i")
                    $o.Width = [single]$d; Place $o ($cx + $rad * [math]::Cos($a) - $d / 2) ($cy + $rad * [math]::Sin($a) - $d / 2)
                    if ($i -eq $k) { $o.ZOrder(0) } else { $slide.Shapes("Cap$i").Delete() }
                }
                Dots $slide $n $k 725 $Accent (Mix $Accent '#FFFFFF' 0.7)
            }
        }
        'ArcList' {
            $cx = -150; $cy = $H / 2 - 20; $rad = 450; $stepDeg = 17
            $disc = Disc $cx $cy $rad $lilac; $rim = Disc $cx $cy ($rad + 14) $Accent; $rim.Fill.Visible = 0; $rim.Line.Visible = -1; $rim.Line.ForeColor.RGB = (Rgb (Mix $Accent '#FFFFFF' 0.55)); $rim.Line.Weight = 1.5
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 45 400 -1 $Font 1 3) }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]
                (Disc 0 0 7 $Accent).Name = "Pin$i"
                $t = Tx $r[0] 24 $gray 0 0 420 -1; $t.TextFrame.WordWrap = 0; $t.Name = "Item$i"
                if ($r.Count -gt 1) { (Grp @((Tx ('{0:00}' -f ($i + 1)) 190 (Mix $Accent '#FFFFFF' 0.7) 640 140 300 $heavyBold $heavy 1 -6), (Tx (($r[1..($r.Count - 1)]) -join "`r") 23 $body ($cx + $rad + 46) ($cy + 36) 430 0 $Font 1 0 1.4))).Name = "Detail$i" }   # the text sits right under the current line, the number is a faint backdrop
            }
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $slide = $pres.Slides($k + 1)
                for ($i = 0; $i -lt $n; $i++) {
                    $a = ($i - $k) * $stepDeg * [math]::PI / 180; $on = $i -eq $k; $ex = $cx + ($rad + 14) * [math]::Cos($a); $ey = $cy + ($rad + 14) * [math]::Sin($a)
                    $pin = $slide.Shapes("Pin$i"); $pd = $(if ($on) { 22 } else { 12 }); $pin.Width = [single]$pd; $pin.Height = [single]$pd; Place $pin ($ex - $pd / 2) ($ey - $pd / 2)
                    $pin.Fill.ForeColor.RGB = (Rgb $(if ($on) { $Accent } else { Mix $Accent '#FFFFFF' 0.55 }))
                    $t = $slide.Shapes("Item$i"); $f = $t.TextFrame.TextRange.Font; $f.Size = [single]$(if ($on) { 44 } else { 24 }); $f.Color.RGB = (Rgb $(if ($on) { $Ink } else { Mix $Ink '#FFFFFF' 0.6 }))
                    Place $t ($ex + $(if ($on) { 30 } else { 22 })) ($ey - [double]$t.Height / 2)
                    if (-not $on -and $rows[$i].Count -gt 1) { $slide.Shapes("Detail$i").Delete() }
                }
            }
        }
        'StatSeries' {
            $halo = Disc ($W / 2) ($H / 2) 150 $lilac; $halo.Name = 'Halo'
            $kick = $null; if ($Kicker) { $kick = Tx $Kicker 13 $Accent 0 70 $W -1 $Font 2 3; $kick.Name = 'Kick' }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]
                $num = Tx $r[0] 220 $Accent 0 118 $W $heavyBold $heavy 2 -6; $num.TextFrame.WordWrap = 0; $num.Name = "Num$i"
                (Tx $r[1] 60 $AccentDark 0 250 160 -1).Name = "Unit$i"
                (Tx $r[2] 32 $Ink 0 392 $W -1 $Font 2).Name = "Say$i"
            }
            MakeDots $n 478 $Accent
            Steps $n
            $pages = @($Light, $Accent, $Dark)
            for ($k = 0; $k -lt $n; $k++) {
                $slide = $pres.Slides($k + 1); $pg = $pages[$k % 3]; $onColour = $pg -ne $Light
                $slide.FollowMasterBackground = 0; $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = (Rgb $pg)
                $hl = $slide.Shapes('Halo'); $hr = 170 + 70 * $k; $hl.Width = [single](2 * $hr); $hl.Height = [single](2 * $hr); Place $hl ($W / 2 - $hr) ($H / 2 - $hr)
                $hl.Fill.ForeColor.RGB = (Rgb $(if ($pg -eq $Light) { $lilac } elseif ($pg -eq $Accent) { Mix $Accent '#FFFFFF' 0.12 } else { $dDeco }))
                if ($kick) { $slide.Shapes('Kick').TextFrame.TextRange.Font.Color.RGB = (Rgb $(if ($onColour) { '#FFFFFF' } else { $Accent })) }
                for ($i = 0; $i -lt $n; $i++) {
                    if ($i -ne $k) { foreach ($nm in 'Num', 'Unit', 'Say') { $slide.Shapes("$nm$i").Delete() }; continue }
                    $num = $slide.Shapes("Num$i"); $unit = $slide.Shapes("Unit$i"); $say = $slide.Shapes("Say$i")
                    $num.TextFrame.TextRange.Font.Color.RGB = (Rgb $(if ($pg -eq $Light) { $Accent } else { '#FFFFFF' }))
                    $num.Left = [single](($W - [double]$num.Width - [double]$unit.Width * 0.6) / 2); Place $unit ([double]$num.Left + [double]$num.Width + 4) 268
                    $unit.TextFrame.TextRange.Font.Color.RGB = (Rgb $(if ($pg -eq $Light) { $AccentDark } elseif ($pg -eq $Accent) { $tint } else { $soft }))
                    $say.TextFrame.TextRange.Font.Color.RGB = (Rgb $(if ($onColour) { '#FFFFFF' } else { $Ink }))
                }
                Dots $slide $n $k ($W / 2) $(if ($onColour) { '#FFFFFF' } else { $Accent }) $(if ($pg -eq $Light) { Mix $Accent '#FFFFFF' 0.7 } elseif ($pg -eq $Accent) { Mix $Accent '#FFFFFF' 0.35 } else { $dLine })
            }
        }
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
