# Impact scenes built on Morph (one slide per step; the same shapes move, grow and recolour between slides):
#   Spotlight  a dimmed photo with a round light that travels from point to point        -Images photo.jpg  -Items 'x;y;r;head;line|...' (x, y = 0..1 of the slide, r = radius in pt)
#   Kinetic    a stack of words; the current one swells to full size in the accent colour -Items 'WORD|WORD|WORD'  -Cards 'line for word 1|line for word 2|...'
#   Carousel   cards on a rail; the current one is large in the middle, neighbours small  -Items 'head;line;line|...' (3-7)
#   Bento      four tiles of different weight that rise in turn (one slide, animations)   -Items 'big;unit;label|big;label|emoji;label|head;line'
# Kinetic advances by itself and stops on the last word; Spotlight and Carousel wait for a click at every step.
# Colours, fonts and text helpers are the ones of scene_page.ps1 (scene_helpers.ps1).
# Example: .\impact_scene.ps1 -Layout Kinetic -Items 'SEE|PRAISE|GROW' -Cards 'what was invisible|what went unsaid|every single day' -Out C:\work\k.pptx -Preview C:\work\k.mp4
param(
    [Parameter(Mandatory)][ValidateSet('Spotlight', 'Kinetic', 'Carousel', 'Bento')][string]$Layout,
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
    [double]$Duration = 0.7,                      # Morph seconds between steps
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent
. (Join-Path $PSScriptRoot 'scene_helpers.ps1')
function Steps([int]$n) { for ($i = 1; $i -lt $n; $i++) { Add-MorphSlide $Duration } }     # everything is built on slide 1, then copied
function Center($s, $cx) { $s.Left = [single]($cx - [double]$s.Width / 2) }

try {
    New-Deck $(if ($Layout -eq 'Spotlight') { $Dark } elseif ($Layout -eq 'Carousel') { '#FFFFFF' } else { $Light })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items

    switch ($Layout) {
        'Spotlight' {
            if (-not $Images.Count) { throw 'Spotlight needs -Images <photo>' }
            $n = $rows.Count; $size = 2800
            [void](Add-Cover $Images[0])
            $veil = Tag ($sl.Shapes.AddShape(18, 0, 0, $size, $size)); $veil.Line.Visible = 0; $veil.Fill.ForeColor.RGB = (ConvertTo-Bgr $Dark); $veil.Fill.Transparency = 0.2; $veil.Name = 'Veil'   # a huge donut: its hole is the light
            $ring = Disc 0 0 10 $Accent; $ring.Fill.Visible = 0; $ring.Line.Visible = -1; $ring.Line.ForeColor.RGB = 0xFFFFFF; $ring.Line.Weight = 2.5; $ring.Line.Transparency = 0.25; $ring.Name = 'Ring'
            if ($Kicker) { [void](Tx $Kicker 12.5 '#FFFFFF' 65 45 600 -1 $Font 1 3) }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $cx = $W * [double]$r[0]; $cy = $H * [double]$r[1]; $rad = [double]$r[2]; $left = $cx -gt $W / 2      # the caption goes to the roomier side
                $capW = 330; $capX = $(if ($left) { $cx - $rad - 28 - $capW } else { $cx + $rad + 28 }); $al = $(if ($left) { 3 } else { 1 })
                $no = Tx ('{0:00}' -f ($i + 1)) 20 $soft $capX ($cy - 62) $capW $heavyBold $heavy $al 2
                $hd = Tx $r[3] 34 '#FFFFFF' $capX ($cy - 34) $capW -1 $Font $al 0 1.15
                $parts = @($no, $hd); if ($r.Count -gt 4) { $parts += (Tx (($r[4..($r.Count - 1)]) -join "`r") 18 $dText $capX ([double]$hd.Top + [double]$hd.Height + 8) $capW 0 $Font $al 0 1.3) }
                (Grp $parts).Name = "Cap$i"
            }
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $s = $pres.Slides($k + 1); $r = $rows[$k]; $cx = $W * [double]$r[0]; $cy = $H * [double]$r[1]; $rad = [double]$r[2]
                $v = $s.Shapes('Veil'); $v.Left = [single]($cx - $size / 2); $v.Top = [single]($cy - $size / 2); $v.Adjustments.Item(1) = [single](($size - 2 * $rad) / (2 * $size))
                $g = $s.Shapes('Ring'); $g.Width = [single](2 * $rad + 14); $g.Height = [single](2 * $rad + 14); $g.Left = [single]($cx - $rad - 7); $g.Top = [single]($cy - $rad - 7)
                for ($i = 0; $i -lt $n; $i++) { if ($i -ne $k) { $s.Shapes("Cap$i").Delete() } }          # only this step's caption stays, so captions cross-fade
            }
        }
        'Kinetic' {
            $n = $rows.Count; $lines = @($Cards -split '\|')
            [void](Disc 900 20 250 $lilac)
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 45 600 -1 $Font 1 3) }
            for ($i = 0; $i -lt $n; $i++) { $t = Tx $rows[$i][0] 40 $gray 0 0 900 $heavyBold $heavy 2; $t.TextFrame.WordWrap = 0; $t.Name = "Word$i" }
            for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i]) { $c = Tx $lines[$i] 24 $body 0 ($H - 96) $W -1 $Font 2; $c.Name = "Line$i" } }
            Steps $n
            $big = 132; $small = 38; $bigH = $big * 1.18; $smallH = $small * 1.5; $total = $bigH + ($n - 1) * $smallH; $y0 = ($H - 70 - $total) / 2 + 10
            for ($k = 0; $k -lt $n; $k++) {
                $s = $pres.Slides($k + 1); $y = $y0
                for ($i = 0; $i -lt $n; $i++) {
                    $on = $i -eq $k; $t = $s.Shapes("Word$i"); $f = $t.TextFrame.TextRange.Font
                    $f.Size = [single]$(if ($on) { $big } else { $small }); $f.Color.RGB = (ConvertTo-Bgr $(if ($on) { $Accent } else { Mix $Ink '#FFFFFF' 0.72 }))
                    if ($on -and [double]$t.Width -gt $W - 100) { $f.Size = [single][math]::Floor($big * ($W - 100) / [double]$t.Width) }
                    $t.Top = [single]$y; Center $t ($W / 2); $y += $(if ($on) { $bigH } else { $smallH })
                    if ($i -lt $lines.Count -and $lines[$i] -and -not $on) { $s.Shapes("Line$i").Delete() }
                }
                if ($k -lt $n - 1) { $s.SlideShowTransition.AdvanceOnTime = -1; $s.SlideShowTransition.AdvanceTime = 1.2 }
            }
        }
        'Carousel' {
            $n = $rows.Count
            [void](Disc 960 560 250 $Light); Head
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]
                $b = Box 0 0 300 340 '#FFFFFF' 22; Glow $b (Mix $Accent '#FFFFFF' 0.45) 0.7 22 8; $b.Name = "Box$i"
                $nm = Tx ('{0:00}' -f ($i + 1)) 54 $Accent 0 0 200 $heavyBold $heavy; $nm.TextFrame.WordWrap = 0; $nm.Name = "Num$i"
                $hd = Tx $r[0] 26 $Ink 0 0 250 -1; $hd.Name = "Head$i"
                if ($r.Count -gt 1) { $bd = Tx (($r[1..($r.Count - 1)]) -join "`r") 17 $body 0 0 250 0 $Font 1 0 1.35; $bd.Name = "Body$i" }
                $d = Box 0 478 8 8 (Mix $Accent '#FFFFFF' 0.7) 4; $d.Name = "Dot$i"
            }
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $s = $pres.Slides($k + 1)
                for ($i = 0; $i -lt $n; $i++) {
                    $o = $i - $k; $ab = [math]::Abs($o); $on = $ab -eq 0
                    $cw = $(if ($on) { 300 } elseif ($ab -eq 1) { 210 } else { 150 }); $chh = $(if ($on) { 286 } elseif ($ab -eq 1) { 214 } else { 156 })
                    $cx = $W / 2 + [math]::Sign($o) * $(if ($ab -eq 0) { 0 } elseif ($ab -eq 1) { 290 } else { 290 + 290 * ($ab - 1) + 30 }); $top = 156 + (286 - $chh) / 2
                    $b = $s.Shapes("Box$i"); $b.Width = [single]$cw; $b.Height = [single]$chh; $b.Left = [single]($cx - $cw / 2); $b.Top = [single]$top
                    $b.Fill.ForeColor.RGB = (ConvertTo-Bgr $(if ($on) { '#FFFFFF' } else { Mix $Accent '#FFFFFF' 0.9 }))
                    $pad = $cw * 0.1
                    $nm = $s.Shapes("Num$i"); $nm.TextFrame.TextRange.Font.Size = [single]$(if ($on) { 54 } elseif ($ab -eq 1) { 34 } else { 24 }); $nm.Left = [single]($cx - $cw / 2 + $pad); $nm.Top = [single]($top + $pad)
                    $nm.TextFrame.TextRange.Font.Color.RGB = (ConvertTo-Bgr $(if ($on) { $Accent } else { Mix $Accent '#FFFFFF' 0.45 }))
                    $hd = $s.Shapes("Head$i"); $hd.TextFrame.TextRange.Font.Size = [single]$(if ($on) { 26 } elseif ($ab -eq 1) { 18 } else { 14 }); $hd.Width = [single]($cw - 2 * $pad); $hd.Left = [single]($cx - $cw / 2 + $pad)
                    $hd.Top = [single]([double]$nm.Top + [double]$nm.Height + $(if ($on) { 14 } else { 6 })); $hd.TextFrame.TextRange.Font.Color.RGB = (ConvertTo-Bgr $(if ($on) { $Ink } else { $gray }))
                    if ($rows[$i].Count -gt 1) { if ($on) { $bd = $s.Shapes("Body$i"); $bd.Left = [single]($cx - $cw / 2 + $pad); $bd.Top = [single]([double]$hd.Top + [double]$hd.Height + 12) } else { $s.Shapes("Body$i").Delete() } }
                    $d = $s.Shapes("Dot$i"); $d.Width = [single]$(if ($on) { 28 } else { 8 }); $d.Fill.ForeColor.RGB = (ConvertTo-Bgr $(if ($on) { $Accent } else { Mix $Accent '#FFFFFF' 0.7 }))
                    $d.Left = [single]($W / 2 - (16 * $n + 20) / 2 + 16 * $i + $(if ($i -gt $k) { 20 } else { 0 }))
                }
            }
        }
        'Bento' {
            [void](Disc 900 20 250 $lilac); Head
            if ($rows.Count -lt 4) { throw 'Bento needs four items' }
            $a = $rows[0]; $t1 = Box 65 128 400 362 $Accent 24; Glow $t1 $Accent 0.55 22 8
            $t1.Fill.TwoColorGradient(1, 1); $t1.Fill.ForeColor.RGB = (ConvertTo-Bgr (Mix $Accent '#FFFFFF' 0.14)); $t1.Fill.BackColor.RGB = (ConvertTo-Bgr $AccentDark); $t1.Fill.GradientAngle = 315
            $n1 = Tx $a[0] 132 '#FFFFFF' 92 196 340 $heavyBold $heavy 1 -4; $n1.TextFrame.WordWrap = 0
            $p1 = @($t1, $n1, (Tx $a[1] 40 $tint ([double]$n1.Left + [double]$n1.Width + 6) 290 100 -1), (Tx $a[2] 22 '#FFFFFF' 95 388 340 -1 $Font 1 0 1.25))
            Anim (Grp $p1) $RISE $AFTER 0.5
            $b = $rows[1]; $t2 = Box 485 128 205 171 $Dark 22
            Anim (Grp @($t2, (Tx $b[0] 50 '#FFFFFF' 507 152 170 $heavyBold $heavy), (Tx $b[1] 16 $dText 508 232 165 0 $Font 1 0 1.25))) $RISE $AFTER 0.4
            $c = $rows[2]; $t3 = Box 710 128 185 171 '#FFFFFF' 22; Glow $t3 (Mix $Accent '#FFFFFF' 0.45) 0.75 18 6
            Anim (Grp @($t3, (Disc 764 184 34 $tint), (Mid (Tx $c[0] 30 $Ink 730 150 68 0 'Segoe UI Emoji' 2) 150 68), (Tx $c[1] 17 $Ink 732 236 150 -1 $Font 1 0 1.2))) $RISE $AFTER 0.4
            $d = $rows[3]; $t4 = Box 485 319 410 171 $tint 22
            $p4 = @($t4, (Tx $d[0] 26 $AccentDark 512 346 360 -1)); if ($d.Count -gt 1) { $p4 += (Tx (($d[1..($d.Count - 1)]) -join "`r") 17 $body 513 390 360 0 $Font 1 0 1.3) }
            Anim (Grp $p4) $RISE $AFTER 0.4
            Foot
        }
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
