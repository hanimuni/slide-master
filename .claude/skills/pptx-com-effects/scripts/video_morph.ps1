# More scenes rebuilt from short tutorial videos (2026-10 batch, part 2). Same colour / font parameters as scene_page.ps1.
#   Cast        cut-out people stand in the dark; each click lights one of them up and shows the line beside them
#               -Images cut1.png cut2.png ... (3-6, transparent)  -Title T  -Items 'role;line;line|...'
#   FoldPanels  contents that stays on the page: full-height arrow panels wait folded at the left edge and come out one per chapter
#               -Title T  [-Sub s]  -Items 'head;point;point|...' (3-5)
#   RingDial    a split ring half below the slide turns a quarter per click; the photo in the top window stays upright
#               -Images a.jpg b.jpg c.jpg d.jpg (3-4)  -Items 'name;line|...'
#   StepFan     sectors of growing radius swirl out of one white dot, with the list beside them   -Title hub  -Items 'head;line|...' (4-6)
#   YearReel    every digit of the year rolls on its own strip; the text and the marker on the time line follow
#               -Items 'year;head;line|...' (3-6, same number of digits)
# Cast, FoldPanels, RingDial and YearReel move one step per click; StepFan plays by itself.
# Example: .\video_morph.ps1 -Layout YearReel -Items '2019;idea;first sketch|2022;pilot;six people|2026;launch;whole company' -Out C:\work\reel.pptx -Preview C:\work\reel.mp4
param(
    [Parameter(Mandatory)][ValidateSet('Cast', 'FoldPanels', 'RingDial', 'StepFan', 'YearReel')][string]$Layout,
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
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Place($shape, $x, $y) { $shape.Left = [single]$x; $shape.Top = [single]$y }
function Side($s, [string]$hex) { $h = $s.Shadow; $h.Visible = -1; $h.ForeColor.RGB = (Rgb $hex); $h.Transparency = 0.6; $h.Blur = 14; $h.OffsetX = 6; $h.OffsetY = 0 }
function Tail($r) { if ($r.Count -gt 1) { ($r[1..($r.Count - 1)]) -join "`r" } else { '' } }

try {
    New-Deck $(if ($Layout -in 'FoldPanels', 'YearReel') { $Light } else { $Dark })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count

    switch ($Layout) {
        'Cast' {
            if ($Images.Count -lt 2) { throw 'Cast needs -Images (2 or more cut-outs)' }
            $n = $Images.Count; $mid = ($n - 1) / 2; $pitch = [math]::Min(175, 740 / $n)
            $halo = Disc ($W / 2) ($H - 120) 170 $Accent; $halo.Fill.Transparency = 1; $halo.Name = 'Halo'; $halo.SoftEdge.Type = 6
            $spot = @{}
            foreach ($i in (0..($n - 1) | Sort-Object { - [math]::Abs($_ - $mid) })) {                                           # the far ones first: the middle stands in front
                $ph = 330 - 36 * [math]::Abs($i - $mid); $p = Tag ($sl.Shapes.AddPicture($Images[$i], 0, -1, 0, 0, -1, -1)); $p.LockAspectRatio = -1; $p.Height = [single]$ph
                $cxi = $W / 2 + ($i - $mid) * $pitch; Place $p ($cxi - [double]$p.Width / 2) ($H - $ph + 26); $p.PictureFormat.Brightness = 0.1; $p.Name = "Cast$i"; $spot[$i] = $cxi
            }
            if ($Kicker) { (Tx $Kicker 13 $soft 0 52 $W -1 $Font 2 3).Name = 'Kick' }
            (Tx $Title 58 '#FFFFFF' 0 80 $W -1 $Font 2 0 1.1).Name = 'Head'
            for ($i = 0; $i -lt $n; $i++) {
                $r = $(if ($i -lt $rows.Count) { $rows[$i] } else { @("$($i + 1)") }); $left = ($spot[$i] -ge $W / 2); $x = $(if ($left) { 65 } else { $W - 65 - 430 }); $al = $(if ($left) { 1 } else { 3 })
                $cap = @((Tx $r[0] 15 $soft $x 54 430 -1 $Font $al 2))
                if ($r.Count -gt 1) { $cap += (Tx (Tail $r) 34 '#FFFFFF' $x 82 430 -1 $Font $al 0 1.2) }
                (Grp $cap).Name = "Cap$i"
            }
            $Duration = 0.7; Steps ($n + 1)
            for ($s = 0; $s -le $n; $s++) {
                $slide = $pres.Slides($s + 1)
                for ($i = 0; $i -lt $n; $i++) {
                    if ($i -ne $s - 1) { $slide.Shapes("Cap$i").Delete(); continue }
                    $p = $slide.Shapes("Cast$i"); $p.PictureFormat.Brightness = 0.5; $bottom = [double]$p.Top + [double]$p.Height; $p.Height = [single]([double]$p.Height * 1.1)   # lit and a little closer
                    Place $p ($spot[$i] - [double]$p.Width / 2) ($bottom - [double]$p.Height)
                    $hl = $slide.Shapes('Halo'); $hl.Fill.Transparency = 0.45; $hl.Left = [single]($spot[$i] - 170)
                }
                if ($s) { $slide.Shapes('Head').Delete(); if ($Kicker) { $slide.Shapes('Kick').Delete() } }
            }
        }
        'FoldPanels' {
            $pw = 250; $strip = 16; $notch = 30; $x0 = ($n - 1) * $strip + $pw + $notch + 54; $cw = $W - $x0 - 60
            for ($i = 1; $i -le $n; $i++) {
                $col = Mix (Mix $Accent '#FFFFFF' 0.12) $Dark (($i - 1) / [math]::Max(1, $n - 1))
                $fb = $sl.Shapes.BuildFreeform(0, 0, 0); foreach ($q in @($pw, 0), @($pw, 112), @(($pw + $notch), 146), @($pw, 180), @($pw, $H), @(0, $H), @(0, 0)) { $fb.AddNodes(0, 0, [single]$q[0], [single]$q[1]) }
                $pane = Tag ($fb.ConvertToShape()); $pane.Line.Visible = 0; $pane.Fill.ForeColor.RGB = (Rgb $col); Side $pane $Dark
                $num = Tx ('{0:00}' -f $i) 66 '#FFFFFF' 34 208 180 $heavyBold $heavy; $hd = Tx $rows[$i - 1][0] 23 '#FFFFFF' 36 292 ($pw - 62) -1 $Font 1 0 1.25
                (Grp @($pane, $num, $hd)).Name = "Pan$i"
            }
            $intro = @((Tx $Title 46 $AccentDark $x0 196 $cw -1 $Font 1 0 1.18)); if ($Sub) { $intro += (Tx $Sub 21 $body $x0 ([double]$intro[0].Top + [double]$intro[0].Height + 16) $cw 0 $Font 1 0 1.4) }
            if ($Kicker) { $intro += (Tx $Kicker 12.5 $Accent $x0 164 $cw -1 $Font 1 3) }
            (Grp $intro).Name = 'Intro'
            for ($i = 1; $i -le $n; $i++) {
                $r = $rows[$i - 1]; $set = @((Tx ('{0:00}' -f $i) 210 $tint ($W - 60 - 320) 296 320 $heavyBold $heavy 3), (Tx ('CHAPTER {0:00}' -f $i) 12.5 $Accent $x0 92 $cw -1 $Font 1 3), (Tx ($r[0] -replace '//', ' ') 40 $Ink $x0 118 $cw -1 $Font 1 0 1.15))
                $y = [double]$set[2].Top + [double]$set[2].Height + 30
                for ($b = 1; $b -lt $r.Count; $b++) { $set += (Disc ($x0 + 7) ($y + 15) 6 $Accent); $t = Tx $r[$b] 21 $body ($x0 + 30) $y ($cw - 30) 0 $Font 1 0 1.3; $set += $t; $y += [double]$t.Height + 20 }
                (Grp $set).Name = "Body$i"
            }
            $Duration = 0.75; Steps ($n + 1)
            for ($s = 0; $s -le $n; $s++) {
                $slide = $pres.Slides($s + 1)
                for ($i = 1; $i -le $n; $i++) {
                    $slide.Shapes("Pan$i").Left = [single]$(if ($i -le $s) { ($n - $s) * $strip } else { ($n - $i + 1) * $strip - $pw })     # out, or folded with only a strip showing
                    if ($i -ne $s) { $slide.Shapes("Body$i").Delete() }
                }
                if ($s) { $slide.Shapes('Intro').Delete() }
            }
        }
        'RingDial' {
            if ($Images.Count -lt 2) { throw 'RingDial needs -Images (3 or 4 photos)' }
            $n = [math]::Min(4, $Images.Count); $cx = $W / 2; $cy = $H + 46; $rad = 336; $inner = 150; $gap = 2.2
            $bw = 2 * $rad * [math]::Sin([math]::PI / 4)
            for ($i = 0; $i -lt 4; $i++) {
                $a = Tag ($sl.Shapes.AddShape(20, [single]($cx - $rad), [single]($cy - $rad), [single](2 * $rad), [single](2 * $rad)))
                $a.Adjustments.Item(1) = [single](225 + $gap); $a.Adjustments.Item(2) = [single](315 - $gap); $a.Adjustments.Item(3) = [single](($rad - $inner) / (2 * $rad)); $a.Line.Visible = 0
                if ($i -lt $n) { $a.Fill.UserPicture($Images[$i]); $a.Fill.RotateWithObject = 0; $c = $a.PictureFormat.Crop; $c.PictureWidth = [single]($bw * 1.04); $c.PictureHeight = [single]($bw * 1.04 * 9 / 16) }   # the photo keeps its own angle while the piece turns
                else { $a.Fill.ForeColor.RGB = (Rgb $dCard) }
                $a.Rotation = [single]($i * 90); $a.Name = "Seg$i"
            }
            $hub = Disc $cx $cy ($inner - 16) (Mix $Dark '#000000' 0.3); $hub.Line.Visible = -1; $hub.Line.ForeColor.RGB = (Rgb $dLine); $hub.Line.Weight = 1.5
            if ($Kicker) { [void](Tx $Kicker 12.5 $soft 65 45 500 -1 $Font 1 3) }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $(if ($i -lt $rows.Count) { $rows[$i] } else { @("$($i + 1)") })
                $cap = @((Tx $r[0] 44 '#FFFFFF' 64 70 420 -1 $Font 1 0 1.12)); if ($r.Count -gt 1) { $cap += (Tx (Tail $r) 20 $dText 66 ([double]$cap[0].Top + [double]$cap[0].Height + 10) 380 0 $Font 1 0 1.35) }
                (Grp $cap).Name = "Cap$i"
                (Tx ('{0:00}' -f ($i + 1)) 54 '#FFFFFF' ($cx - 80) ($H - 92) 160 $heavyBold $heavy 2).Name = "Num$i"
            }
            [void](Tx ('/ {0:00}' -f $n) 15 $dText ($W - 165) 52 100 -1 $Font 3)
            $Duration = 1.1; Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $slide = $pres.Slides($k + 1)
                for ($i = 0; $i -lt 4; $i++) { $slide.Shapes("Seg$i").Rotation = [single]((($i - $k) * 90 + 360) % 360) }
                for ($i = 0; $i -lt $n; $i++) { if ($i -ne $k) { $slide.Shapes("Cap$i").Delete(); $slide.Shapes("Num$i").Delete() } }
            }
        }
        'StepFan' {
            $cx = 290; $cy = $H / 2 + 14; $ang = [math]::Min(60, 320 / $n); $r0 = 128; $r1 = 236
            if ($Kicker) { [void](Tx $Kicker 12.5 $soft 65 45 600 -1 $Font 1 3) }
            $rowH = [math]::Min(84, 400 / $n); $ly = $cy - $n * $rowH / 2
            for ($i = 0; $i -lt $n; $i++) {
                $f = $i / [math]::Max(1, $n - 1); $rad = $r0 + ($r1 - $r0) * $f; $col = Mix (Mix $Accent '#FFFFFF' 0.55) $AccentDark $f
                $wedge = Tag ($sl.Shapes.AddShape(142, [single]($cx - $rad), [single]($cy - $rad), [single](2 * $rad), [single](2 * $rad)))
                $wedge.Adjustments.Item(1) = [single]((270 + $i * $ang) % 360); $wedge.Adjustments.Item(2) = [single]((270 + ($i + 1) * $ang) % 360)
                $wedge.Line.Visible = 0; $wedge.Fill.ForeColor.RGB = (Rgb $col); Glow $wedge '#000000' 0.55 16 0; $wedge.Name = "Fan$i"
                $m = (270 + ($i + 0.5) * $ang) * [math]::PI / 180; $lr = ($rad + 84) / 2 + 8
                (Tx ('{0:00}' -f ($i + 1)) 21 $(if ($f -lt 0.35) { $Ink } else { '#FFFFFF' }) ($cx + $lr * [math]::Cos($m) - 30) ($cy + $lr * [math]::Sin($m) - 14) 60 $heavyBold $heavy 2).Name = "No$i"
                $row = @((Disc 596 ($ly + $i * $rowH + 17) 8 $col), (Tx $rows[$i][0] 25 '#FFFFFF' 620 ($ly + $i * $rowH) 290 -1))
                if ($rows[$i].Count -gt 1) { $row += (Tx (Tail $rows[$i]) 16 $dText 620 ($ly + $i * $rowH + 36) 290) }
                (Grp $row).Name = "Row$i"
            }
            $ring = Disc $cx $cy 84 '#FFFFFF'; $ring.Fill.Transparency = 0.55; $ring.Name = 'Halo'
            $hub = Disc $cx $cy 62 '#FFFFFF'; Glow $hub '#000000' 0.6 14 3; $hub.Name = 'Hub'
            (Mid (Tx $Title 17 $Ink ($cx - 54) ($cy - 40) 108 -1 $Font 2 0 1.15) ($cy - 40) 80).Name = 'HubText'
            $Duration = 1.2; Steps 2
            $s1 = $pres.Slides(1)
            for ($i = 0; $i -lt $n; $i++) {
                $wedge = $s1.Shapes("Fan$i"); $wedge.Width = 60; $wedge.Height = 60; Place $wedge ($cx - 30) ($cy - 30); $wedge.Rotation = [single]((360 - $i * $ang) % 360)   # folded onto the first sector, hidden behind the dot
                $s1.Shapes("No$i").Delete(); $s1.Shapes("Row$i").Delete()
            }
            $s1.Shapes('HubText').Delete(); $d = $s1.Shapes('Hub'); $d.Width = 44; $d.Height = 44; Place $d ($cx - 22) ($cy - 22); $d = $s1.Shapes('Halo'); $d.Width = 44; $d.Height = 44; Place $d ($cx - 22) ($cy - 22)
            Auto $s1 0.6
        }
        'YearReel' {
            $digits = $rows[0][0].Length; $size = 178; $cw = $size * 0.63; $xl = 58; $winTop = 124
            $strips = @(); $tops = $null; $lineH = 0
            for ($d = 0; $d -lt $digits; $d++) {
                $t = Tx ((0..9) -join '|') $size $Ink ($xl + $d * $cw) $winTop $cw $heavyBold $heavy 2 0 0.92; $t.Name = "Reel$d"; $strips += $t
                if (-not $tops) { $tr = $t.TextFrame.TextRange; $tops = @(1..10 | ForEach-Object { [double]$tr.Lines($_).BoundTop - [double]$t.Top }); $lineH = ($tops[9] - $tops[0]) / 9 }
            }
            [void](Box 0 0 $W $winTop $Light); [void](Box 0 ($winTop + $lineH) $W ($H - $winTop - $lineH) $Light)                 # only one row of digits shows between the two covers
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 45 600 -1 $Font 1 3) }
            $tx = $xl + $digits * $cw + 46; $tw = $W - $tx - 60
            for ($k = 0; $k -lt $n; $k++) {
                $r = $rows[$k]; $cap = @((Tx $r[1] 40 $AccentDark $tx ($winTop + 34) $tw -1 $Font 1 0 1.15))
                if ($r.Count -gt 2) { $cap += (Tx (($r[2..($r.Count - 1)]) -join "`r") 21 $body $tx ([double]$cap[0].Top + [double]$cap[0].Height + 14) $tw 0 $Font 1 0 1.4) }
                (Grp $cap).Name = "Cap$k"
            }
            $ty = 452; $x1 = 78; $x2 = $W - 78; [void](Box $x1 ($ty - 1.5) ($x2 - $x1) 3 $lilac 1.5)
            for ($k = 0; $k -lt $n; $k++) { $px = $x1 + ($x2 - $x1) * $k / [math]::Max(1, $n - 1); [void](Disc $px $ty 6 $lilac); [void](Tx $rows[$k][0] 16 $gray ($px - 50) ($ty + 18) 100 -1 $Font 2) }
            $mark = Disc $x1 $ty 12 $Accent; Glow $mark $Accent 0.5 12 3; $mark.Name = 'Mark'
            $Duration = 1.1; Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $slide = $pres.Slides($k + 1); $yr = $rows[$k][0]
                for ($d = 0; $d -lt $digits; $d++) { $slide.Shapes("Reel$d").Top = [single]($winTop - $tops[[int]"$($yr[$d])"]) }
                for ($i = 0; $i -lt $n; $i++) { if ($i -ne $k) { $slide.Shapes("Cap$i").Delete() } }
                $slide.Shapes('Mark').Left = [single]($x1 + ($x2 - $x1) * $k / [math]::Max(1, $n - 1) - 12)
            }
        }
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
