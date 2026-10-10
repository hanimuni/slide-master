# Scenes rebuilt from short tutorial videos (2026-10 batch), each restyled to the scene_page standard
# (big type, three-colour system, plays by itself where that helps). Same colour / font parameters as scene_page.ps1.
#   ZoomThrough  a pale panel with the title punched out; the panel rushes at the viewer until the slide passes
#                through one letter and the whole photo is uncovered            -Images photo.jpg  -Title WORD  [-Sub caption]  [-Items 'fx;fy' point inside a stroke]
#   Bubbles      keyword bubbles of different sizes burst out of the title circle  -Title 'centre|text'  -Items 'label;1-3|...' (8-16)
#   Showcase     full-bleed photo slider: the next picture grows out of its thumbnail on the right rail  -Images a.jpg b.jpg ...  -Items 'head;line|...'
#   CircleFan    contents page: circles larger than the slide fan out from the left  -Title T  -Items 'head|head|...' (3-5)
#   FanTimeline  coloured sectors swing out from behind a half disc, one per step   -Title T  -Items 'year;line|...' (3-6)
# ZoomThrough, Bubbles, CircleFan and FanTimeline play through by themselves; Showcase waits for a click per photo.
# Example: .\video_scene.ps1 -Layout FanTimeline -Title 'Roadmap' -Items '2024;idea|2025;pilot|2026;launch' -Out C:\work\fan.pptx -Preview C:\work\fan.mp4
param(
    [Parameter(Mandatory)][ValidateSet('ZoomThrough', 'Bubbles', 'Showcase', 'CircleFan', 'FanTimeline')][string]$Layout,
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
    [double]$Zoom = 70,                           # ZoomThrough: how many times the panel grows
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
function Side($s, [string]$hex) { $h = $s.Shadow; $h.Visible = -1; $h.ForeColor.RGB = (Rgb $hex); $h.Transparency = 0.6; $h.Blur = 14; $h.OffsetX = 6; $h.OffsetY = 0 }   # shadow thrown to the right

try {
    New-Deck $(if ($Layout -in 'ZoomThrough', 'Showcase') { $Dark } else { $Light })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count

    switch ($Layout) {
        'ZoomThrough' {
            if (-not $Images.Count) { throw 'ZoomThrough needs -Images <photo>' }
            [void](Add-Cover $Images[0])
            $panel = New-Rect -10 -10 ($W + 20) ($H + 20); $panel.Line.Visible = 0; $panel.Fill.ForeColor.RGB = (Rgb $Light)
            $t = New-Text $Title 190 0 $heavy; if (-not $HeavyFont) { $t.TextFrame.TextRange.Font.Bold = -1 }
            if ([double]$t.Width -gt $W - 150) { $t.TextFrame.TextRange.Font.Size = [single][math]::Floor(190 * ($W - 150) / [double]$t.Width) }
            Set-Center $t; $tr = $t.TextFrame.TextRange
            $fx = 0.07; $fy = 0.5; if ($n -and $rows[0].Count -ge 2) { $fx = [double]$rows[0][0]; $fy = [double]$rows[0][1] }       # a point that lies inside a stroke (default: the first letter's left stem)
            $px = [double]$tr.BoundLeft + $fx * [double]$tr.BoundWidth; $py = [double]$tr.BoundTop + $fy * [double]$tr.BoundHeight
            $mask = Merge-Shapes 4 @($panel.Name, $t.Name); $mask.Name = 'Mask'                                                 # panel first, text second: the letters become holes
            if ($Kicker) { (Tx $Kicker 13 $AccentDark 0 62 $W -1 $Font 2 3).Name = 'Kick' }
            $Duration = 1.5; Steps 2
            $s2 = $pres.Slides(2); $m = $s2.Shapes('Mask'); $l0 = [double]$m.Left; $t0 = [double]$m.Top
            $m.LockAspectRatio = -1; $m.Width = [single]([double]$m.Width * $Zoom)
            Place $m ($W / 2 - ($px - $l0) * $Zoom) ($H / 2 - ($py - $t0) * $Zoom)                                               # the chosen stroke now covers the whole slide
            if ($Kicker) { $s2.Shapes('Kick').Delete() }
            if ($Sub) { $script:sl = $s2; [void](Grp (Pill 40 ($H - 84) (Wd $Sub 20) 44 $Dark $Sub 20 '#FFFFFF')) }
            Auto $pres.Slides(1) 0.9
        }
        'Bubbles' {
            $cx = $W / 2; $cy = $H / 2 + 6; $fills = @($tint, '#FFFFFF', $Accent, $Dark)
            $spot = @()
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $wt = 2; if ($r.Count -gt 1) { $wt = [int]$r[1] }; $rad = 34 + 14 * $wt; $kind = $i % 4
                $b = Disc $cx $cy $rad $fills[$kind]; $b.Name = "Bub$i"; if ($kind -eq 1) { Glow $b (Mix $Accent '#FFFFFF' 0.45) 0.7 16 5 }
                $f = $b.TextFrame; $f.WordWrap = 0; $f.MarginLeft = 0; $f.MarginRight = 0; $f.TextRange.Text = ($r[0] -replace '//', "`r")
                $f.TextRange.Font.Name = $Font; $f.TextRange.Font.NameFarEast = $Font; $f.TextRange.Font.Bold = -1; $f.TextRange.Font.Size = [single](13 + 2.5 * $wt)
                $f.TextRange.Font.Color.RGB = (Rgb $(if ($kind -ge 2) { '#FFFFFF' } elseif ($kind -eq 1) { $AccentDark } else { $Ink }))
                $ang = (2 * [math]::PI * $i / $n) + $(if ($i % 2) { 0.18 } else { -0.1 }); $dist = $(if ($i % 2) { 228 } else { 150 }) + 10 * $wt
                $spot += , @(($cx + 1.62 * $dist * [math]::Cos($ang)), ($cy + 0.84 * $dist * [math]::Sin($ang)), $rad)
            }
            for ($pass = 0; $pass -lt 300; $pass++) {                                                                          # push overlapping bubbles apart, keep them on the slide and off the title circle
                for ($i = 0; $i -lt $n; $i++) {
                    for ($j = -1; $j -lt $n; $j++) {
                        if ($j -eq $i) { continue }
                        if ($j -lt 0) { $ox = $cx; $oy = $cy; $or = 96 } else { $ox = $spot[$j][0]; $oy = $spot[$j][1]; $or = $spot[$j][2] }
                        $dx = $spot[$i][0] - $ox; $dy = $spot[$i][1] - $oy; $d = [math]::Max(0.01, [math]::Sqrt($dx * $dx + $dy * $dy)); $need = $spot[$i][2] + $or + 12
                        if ($d -lt $need) { $push = ($need - $d) * $(if ($j -lt 0) { 1 } else { 0.5 }); $spot[$i][0] += $dx / $d * $push; $spot[$i][1] += $dy / $d * $push }
                    }
                    $rr = $spot[$i][2] + 14
                    $spot[$i][0] = [math]::Min($W - $rr, [math]::Max($rr, $spot[$i][0])); $spot[$i][1] = [math]::Min($H - $rr, [math]::Max($rr, $spot[$i][1]))
                }
            }
            $core = Disc $cx $cy 96 $Accent; $core.Fill.TwoColorGradient(1, 1); $core.Fill.ForeColor.RGB = (Rgb (Mix $Accent '#FFFFFF' 0.12)); $core.Fill.BackColor.RGB = (Rgb $AccentDark); $core.Fill.GradientAngle = 315; Glow $core $Accent 0.5 26 8
            [void](Mid (Tx $Title 25 '#FFFFFF' ($cx - 86) ($cy - 60) 172 -1 $Font 2 0 1.2) ($cy - 60) 120)
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 45 600 -1 $Font 1 3) }
            $Duration = 1.1; Steps 2
            $s2 = $pres.Slides(2); for ($i = 0; $i -lt $n; $i++) { Place $s2.Shapes("Bub$i") ($spot[$i][0] - $spot[$i][2]) ($spot[$i][1] - $spot[$i][2]) }
            Auto $pres.Slides(1) 0.6
        }
        'Showcase' {
            if (-not $Images.Count) { throw 'Showcase needs -Images' }
            $n = $Images.Count; $tw = 168; $th = $tw * $H / $W; $railX = $W - 44 - $tw
            for ($i = 0; $i -lt $n; $i++) { $p = Add-Cover $Images[$i]; $p.Name = "Shot$i"; $p.Line.ForeColor.RGB = 0xFFFFFF; $p.Line.Weight = 2.5; $p.Line.Visible = 0 }
            $sc = Box 0 0 $W $H $Dark; $sc.Fill.TwoColorGradient(1, 1); $sc.Fill.GradientAngle = 0
            $gs = $sc.Fill.GradientStops; $gs.Item(1).Color.RGB = (Rgb $Dark); $gs.Item(1).Position = 0.12; $gs.Item(1).Transparency = 0.12; $gs.Item(2).Color.RGB = (Rgb $Dark); $gs.Item(2).Position = 0.72; $gs.Item(2).Transparency = 1
            if ($Kicker) { [void](Tx $Kicker 12.5 '#FFFFFF' 65 52 500 -1 $Font 1 3) }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $(if ($i -lt $rows.Count) { $rows[$i] } else { @("$($i + 1)") })
                $cap = @((Tx ('{0:00}' -f ($i + 1)) 26 $soft 66 188 200 $heavyBold $heavy), (Tx (' / {0:00}' -f $n) 16 $dText 104 199 100 -1), (Tx $r[0] 54 '#FFFFFF' 64 228 500 -1 $Font 1 0 1.12))
                if ($r.Count -gt 1) { $cap += (Tx (($r[1..($r.Count - 1)]) -join "`r") 21 $dText 66 ([double]$cap[2].Top + [double]$cap[2].Height + 12) 460 0 $Font 1 0 1.35) }
                (Grp $cap).Name = "Cap$i"
            }
            $rail = Box 65 ($H - 62) 220 3 (Mix $Dark '#FFFFFF' 0.35) 1.5; $bar = Box 65 ($H - 63) (220 / $n) 5 '#FFFFFF' 2.5; $bar.Name = 'Bar'
            Steps $n
            for ($k = 0; $k -lt $n; $k++) {
                $slide = $pres.Slides($k + 1)
                for ($i = 0; $i -lt $n; $i++) {
                    $p = $slide.Shapes("Shot$i")
                    if ($i -gt $k) { $slot = $i - $k - 1; $p.LockAspectRatio = -1; $p.Width = [single]$tw; Place $p $railX (150 + $slot * ($th + 16)); $p.Line.Visible = -1 }     # waiting on the rail
                    if ($i -ne $k) { $slide.Shapes("Cap$i").Delete() }
                }
                $slide.Shapes('Bar').Left = [single](65 + $k * 220 / $n)
            }
        }
        'CircleFan' {
            $dia = 690; $cyy = $H / 2; $pitch = [math]::Min(170, (860 - 215) / $n); $x0 = -($dia - 215)
            for ($i = 0; $i -le $n; $i++) {
                $col = $(if ($i -eq 0) { $Dark } else { Mix $AccentDark (Mix $Accent '#FFFFFF' 0.35) (($i - 1) / [math]::Max(1, $n - 1)) })
                $c = Disc ($x0 + $dia / 2) $cyy ($dia / 2) $col; $c.Name = "Ring$i"; Side $c $Dark; $c.ZOrder(1)                 # each new circle goes to the back: the first one stays on top
            }
            if ($Kicker) { [void](Tx $Kicker 12.5 $soft 44 196 170 -1 $Font 1 3) }
            [void](Tx $Title 32 '#FFFFFF' 44 $(if ($Kicker) { 222 } else { 208 }) 168 -1 $Font 1 0 1.18)
            for ($i = 1; $i -le $n; $i++) {
                $edge = 215 + $i * $pitch
                (Grp @((Tx ('{0:00}' -f $i) 52 '#FFFFFF' ($edge - $pitch + 34) 214 120 $heavyBold $heavy), (Tx $rows[$i - 1][0] 20 '#FFFFFF' ($edge - $pitch + 36) 286 ($pitch - 44) -1 $Font 1 0 1.25))).Name = "Lab$i"
            }
            $Duration = 1.1; Steps 2
            $s1 = $pres.Slides(1); for ($i = 1; $i -le $n; $i++) { $s1.Shapes("Lab$i").Delete() }                                 # labels fade in as the circles open
            $s2 = $pres.Slides(2); for ($i = 1; $i -le $n; $i++) { $s2.Shapes("Ring$i").Left = [single]($x0 + $i * $pitch) }
            Auto $s1 0.6
        }
        'FanTimeline' {
            $cx = $W / 2; $cyy = $H; $rad = 404; $ang = 180 / $n
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 45 600 -1 $Font 1 3) }
            for ($i = 0; $i -lt $n; $i++) {
                $lightness = $i / [math]::Max(1, $n - 1); $col = Mix $AccentDark (Mix $Accent '#FFFFFF' 0.42) $lightness
                $wedge = Tag ($sl.Shapes.AddShape(142, [single]($cx - $rad), [single]($cyy - $rad), [single](2 * $rad), [single](2 * $rad)))     # pie: a sector from 180 deg, turned into place
                $wedge.Line.Visible = -1; $wedge.Line.ForeColor.RGB = (Rgb $Light); $wedge.Line.Weight = 3; $wedge.Fill.ForeColor.RGB = (Rgb $col)
                $wedge.Adjustments.Item(1) = 180; $wedge.Adjustments.Item(2) = [single](180 + $ang); $wedge.Name = "Fan$i"
                $mid = (180 + ($i + 0.5) * $ang) * [math]::PI / 180; $lx = $cx + 292 * [math]::Cos($mid); $ly = $cyy + 292 * [math]::Sin($mid)
                $ink2 = $(if ($lightness -gt 0.7) { $Ink } else { '#FFFFFF' })
                $lab = @((Tx $rows[$i][0] 24 $ink2 ($lx - 80) ($ly - 34) 160 $heavyBold $heavy 2))
                if ($rows[$i].Count -gt 1) { $lab += (Tx (($rows[$i][1..($rows[$i].Count - 1)]) -join "`r") 15 $ink2 ($lx - 80) ($ly - 2) 160 -1 $Font 2 0 1.25) }
                (Grp $lab).Name = "Lab$i"
            }
            $hub = Disc $cx $cyy 168 $Dark; Glow $hub $Dark 0.5 24 0
            [void](Tx $Title 26 '#FFFFFF' ($cx - 140) ($cyy - 118) 280 -1 $Font 2 0 1.2)
            $Duration = 0.6; Steps ($n + 1)
            for ($s = 0; $s -le $n; $s++) {
                $slide = $pres.Slides($s + 1)
                for ($i = 0; $i -lt $n; $i++) {
                    if ($i -lt $s) { $slide.Shapes("Fan$i").Rotation = [single]($i * $ang) } else { $slide.Shapes("Fan$i").Rotation = [single](360 - $ang); $slide.Shapes("Lab$i").Delete() }   # still folded away below the slide edge
                }
                if ($s -lt $n) { Auto $slide $(if ($s) { 0.35 } else { 0.5 }) }
            }
        }
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
