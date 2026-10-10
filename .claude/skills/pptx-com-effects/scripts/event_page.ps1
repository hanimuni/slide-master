# Big-screen slides for official events (expo, forum, ceremony, award). Same colour / font parameters as scene_page.ps1.
#   Backdrop   stage main screen: huge title in a slowly turning ring, two stage lights swinging, host strip. Loops for ever (3 morph slides)
#              -Kicker 'EN NAME' -Title 'line//line' -Sub slogan -Cards 'date|venue' -Items 'label;orgs|label;orgs|label;orgs'
#   Standby    waiting screen: a ring fills while the clock hand goes round once, the event name drifts behind as a step-and-repeat wall
#              -Kicker 'EVENT NAME' -Title 'line//line' -Sub line -Items 'notice|notice|notice' [-Note hosts] [-Duration seconds of one sweep, default 12]
#   Ticket     what / when / where / who on an admission ticket: it rises, the stub tears off, a stamp lands
#              -Title T -Items 'big date;caption|LABEL;value|LABEL;value|LABEL;value' -Cards 'stub big|stub line|stub line' -Sub 'stamp//text'
#   Programme  order of ceremony: each click moves the highlight, shows the item large and turns the clock hands to its time
#              -Title T -Items 'HH:MM;name;detail;detail|...' (4-8)
#   Welcome    the guest name is the hero: rays turn slowly, a folded ribbon unrolls from the middle, the name rises
#              -Kicker event -Title name -Sub 'ribbon text' -Cards 'line under the name' -Items 'role;name|role;name|...' (other guests, 0-4)
#   Speaker    keynote card: portrait in a coloured arch, talk title, broadcast-style name bar that wipes in, blinking on-stage dot
#              -Images portrait.png[,stage.jpg] -Kicker k -Sub 'time and place' -Title 'talk//title' -Items 'point|point|point' -Cards 'name|role'
#   Award      a stage light swings left and right, settles in the middle, then laurel leaves grow and the winner's name rises
#              -Kicker 'award name' -Sub 'question' -Title winner -Cards 'rank|organisation|reason'
#   Sessions   session guide for a tech expo: tracks x time slots, a scan line sweeps and the cells switch on behind it
#              -Title T -Sub 'time|time|time' -Items 'A;track name;hall|...' (2-4) -Cards 'title;who|...' (row by row; a title starting with * is the featured one) [-Note day]
# Programme waits for a click; Backdrop loops by itself; the others play once by themselves.
# Example: .\event_page.ps1 -Layout Programme -Title 'Order of ceremony' -Items '14:00;Opening;MC|14:10;Address;Minister' -Out C:\work\p.pptx -Preview C:\work\p.mp4
param(
    [Parameter(Mandatory)][ValidateSet('Backdrop', 'Standby', 'Ticket', 'Programme', 'Welcome', 'Speaker', 'Award', 'Sessions')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Note = '',
    [string]$Accent = '#E6B655',
    [string]$AccentDark = '',
    [string]$Light = '#F6F1E7',
    [string]$Dark = '#0B1F4B',
    [string]$Ink = '#1B1F2A',
    [string]$Font = 'Malgun Gothic',
    [string]$HeavyFont = '',
    [double]$Duration = 0,
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
function Turn($shape, $deg) { $shape.Rotation = [single]((($deg % 360) + 360) % 360) }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Grad($s, [string]$from, [string]$to, $angle, $clearFrom = 0, $clearTo = 0) {      # two-stop gradient; 90 = top to bottom
    $s.Fill.TwoColorGradient(1, 1); $s.Fill.GradientAngle = [single]$angle; $g = $s.Fill.GradientStops
    $g.Item(1).Color.RGB = (Rgb $from); $g.Item(1).Transparency = [single]$clearFrom; $g.Item(2).Color.RGB = (Rgb $to); $g.Item(2).Transparency = [single]$clearTo
}
function Ring($cx, $cy, $r, [string]$hex, $weight, [int]$dash = 0) {   # circle outline, no fill
    $s = Tag ($sl.Shapes.AddShape(9, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r))); $s.Fill.Visible = 0
    $s.Line.ForeColor.RGB = (Rgb $hex); $s.Line.Weight = [single]$weight; if ($dash) { $s.Line.DashStyle = $dash }; $s
}
function Arc($cx, $cy, $r, $from, $to, $thick, [string]$hex) {         # block arc: angles clockwise from east (top = 270)
    $s = Tag ($sl.Shapes.AddShape(20, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r))); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex)
    $s.Adjustments.Item(1) = [single]$from; $s.Adjustments.Item(2) = [single]$to; $s.Adjustments.Item(3) = [single]($thick / (2 * $r)); $s
}
# A hand that turns about ($cx,$cy): the hairline tail makes the shape's own centre fall on the pivot.
function Needle([double]$cx, [double]$cy, [double]$len, [double]$wid, [string]$hex) {
    $hw = $wid / 2
    Poly @(@(($cx - $hw), $cy), @(($cx - $hw * 0.45), ($cy - $len)), @(($cx + $hw * 0.45), ($cy - $len)), @(($cx + $hw), $cy), @(($cx + 0.4), $cy), @($cx, ($cy + $len)), @(($cx - 0.4), $cy)) $hex
}
# A stage light hanging at ($ax,$ay): the thin spike above keeps the pivot at the lamp, so Rotation swings the beam.
function Beam([double]$ax, [double]$ay, [double]$len, [double]$w0, [double]$w1, [string]$hex, [double]$clearTop) {
    $s = Poly @(@(($ax - $w0), $ay), @($ax, ($ay - $len)), @(($ax + $w0), $ay), @(($ax + $w1), ($ay + $len)), @(($ax - $w1), ($ay + $len))) $hex
    Grad $s $hex $hex 90 ([math]::Max(0, 2 * $clearTop - 1)) 1; $s.SoftEdge.Type = 3; $s
}
function Tm($e, $dur, $delay) { $e.Timing.Duration = [single]$dur; $e.Timing.TriggerDelayTime = [single]$delay; $e.Timing.SmoothStart = 0; $e.Timing.SmoothEnd = 0 }
# Custom effects at constant speed. Add the behaviour BEFORE touching the timing: the other order leaves the shape hidden while it plays.
function Spin($shape, $deg, [int]$trig, $dur, $delay = 0) { $e = $sl.TimeLine.MainSequence.AddEffect($shape, 0, 0, $trig); $e.Behaviors.Add(4).RotationEffect.By = [single]$deg; Tm $e $dur $delay }
function Glide($shape, $dx, $dy, [int]$trig, $dur, $delay = 0) { $e = $sl.TimeLine.MainSequence.AddEffect($shape, 0, 0, $trig); $e.Behaviors.Add(1).MotionEffect.Path = ('M 0 0 L {0:F4} {1:F4} E' -f ($dx / $W), ($dy / $H)); Tm $e $dur $delay }
function Blink($shape, $delay, [int]$times = 8) { $e = $sl.TimeLine.MainSequence.AddEffect($shape, $FADE, 0, $WITH); $e.Timing.Duration = 0.6; $e.Timing.TriggerDelayTime = [single]$delay; $e.Timing.AutoReverse = -1; $e.Timing.RepeatCount = $times }
function Ems([string]$text) { $m = 0; foreach ($ch in $text.ToCharArray()) { $m += $(if ([int]$ch -lt 128) { 0.58 } else { 1.0 }) }; $m }
function Clock([string]$hhmm) { $p = $hhmm -split ':'; $hh = [int]$p[0]; $mm = [int]$p[1]; @((($hh % 12) * 30 + $mm * 0.5), ($mm * 6)) }   # hour hand, minute hand in degrees

try {
    $cl = @($Cards -split '\|' | Where-Object { $_ })
    New-Deck $(switch ($Layout) { 'Ticket' { $Accent } { $_ -in 'Standby', 'Programme', 'Welcome' } { $Light } default { $Dark } })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count; $secs = 3

    switch ($Layout) {
        'Backdrop' {
            $cx = $W / 2; $cy = 228; $rr = 205
            $halo = Disc $cx $cy 330 (Mix $Dark '#FFFFFF' 0.1); $halo.SoftEdge.Type = 6; $halo.Fill.Transparency = 0.25
            (Beam 150 -24 640 12 96 $Light 0.3).Name = 'BeamA'; (Beam 810 -24 640 12 96 $Light 0.3).Name = 'BeamB'
            $parts = @((Ring $cx $cy ($rr + 18) (Mix $Dark $Accent 0.55) 1 3), (Arc $cx $cy $rr 200 340 5 $Accent), (Arc $cx $cy $rr 20 160 5 $Accent),
                (Arc $cx $cy ($rr - 16) 250 290 2 (Mix $Dark $Accent 0.7)), (Arc $cx $cy ($rr - 16) 70 110 2 (Mix $Dark $Accent 0.7)),
                (Disc ($cx + $rr - 2.5) $cy 7 $Accent), (Disc ($cx - $rr + 2.5) $cy 7 $Accent))
            (Grp $parts).Name = 'Ring'
            if ($Kicker) { [void](Tx $Kicker 16 $Accent 0 100 $W -1 $Font 2 7) }
            $lines = @($Title -split '//'); $big = [math]::Min(72, [math]::Floor(700 / (($lines | ForEach-Object { Ems $_ } | Measure-Object -Maximum).Maximum)))
            $tt = Tx $Title $big '#FFFFFF' 0 132 $W $heavyBold $heavy 2 0 1.12; Glow $tt (Mix $Dark '#000000' 0.5) 0.35 14 4; $y = [double]$tt.Top + [double]$tt.Height + 10
            if ($Sub) { [void](Tx $Sub 25 $Light 0 $y $W -1 $Font 2 1) }
            $y = $cy + $rr - 20
            if ($cl.Count) {
                $line = $cl -join "   $([char]0xB7)   "; $pw = (Ems $line) * 19 + 56
                $pill = Box (($W - $pw) / 2) $y $pw 40 $Accent 20; [void](Mid (Tx $line 19 (Mix $Dark '#000000' 0.3) (($W - $pw) / 2) $y $pw -1 $Font 2) $y 40)
            }
            $band = Box 0 458 $W 82 (Mix $Dark '#000000' 0.45); $band.Fill.Transparency = 0.25; [void](Box 0 457 $W 1.5 (Mix $Dark $Accent 0.6))
            $cw = ($W - 120) / [math]::Max(1, $n)
            for ($i = 0; $i -lt $n; $i++) {
                $x = 60 + $cw * $i
                [void](Tx $rows[$i][0] 15 $Accent $x 471 $cw -1 $Font 2 4); [void](Tx $rows[$i][1] 18 '#FFFFFF' $x 495 $cw -1 $Font 2)
                if ($i) { [void](Box $x 474 1 46 (Mix $Dark '#FFFFFF' 0.3)) }
            }
            if (-not $Duration) { $Duration = 7 }
            Steps 3; $s1 = $pres.Slides(1); $s2 = $pres.Slides(2); $s3 = $pres.Slides(3)
            $s1.SlideShowTransition.EntryEffect = 0
            foreach ($s in $s1, $s3) { Turn $s.Shapes('BeamA') -24; Turn $s.Shapes('BeamB') 24 }
            Turn $s2.Shapes('BeamA') 6; Turn $s2.Shapes('BeamB') -6
            Turn $s2.Shapes('Ring') 90; Turn $s3.Shapes('Ring') 180                                            # the ring looks the same every 180 deg, so slide 3 -> slide 1 is seamless
            foreach ($s in $s1, $s2, $s3) { Auto $s 0.1 }
            $pres.SlideShowSettings.LoopUntilStopped = -1; $pres.SlideShowSettings.AdvanceMode = 2
        }
        'Standby' {
            if (-not $Duration) { $Duration = 12 }
            $wall = @(); $word = ($Kicker + '      ') * 6
            for ($r = 0; $r -lt 7; $r++) { $t = New-Text $word 30 (Rgb (Mix $Light $Accent 0.09)) $heavy; Place $t $(if ($r % 2) { -330 } else { -90 }) (4 + 80 * $r); $wall += (Tag $t) }
            $wallG = Grp $wall
            $cx = 262; $cy = 272; $rr = 150
            $face = Disc $cx $cy ($rr + 34) (Mix $Light '#FFFFFF' 0.75); Glow $face $Dark 0.8 30 10
            [void](Ring $cx $cy $rr (Mix $Light $Accent 0.16) 16)
            $prog = Ring $cx $cy $rr $Accent 16
            for ($k = 0; $k -lt 12; $k++) { $a = $k * [math]::PI / 6; [void](Disc ($cx + 116 * [math]::Sin($a)) ($cy - 116 * [math]::Cos($a)) $(if ($k % 3) { 2.5 } else { 5 }) $(if ($k % 3) { $gray } else { $Dark })) }
            $hand = Needle $cx $cy 104 9 $Dark; [void](Disc $cx $cy 11 $Dark); [void](Disc $cx $cy 4.5 $Accent)
            Glide $wallG -150 0 $AFTER ($Duration + 3)
            Anim $prog $WHEEL $WITH $Duration 0.3; $fx = $sl.TimeLine.MainSequence.Item($sl.TimeLine.MainSequence.Count); $fx.Timing.SmoothStart = 0; $fx.Timing.SmoothEnd = 0
            Spin $hand 360 $WITH $Duration 0.3
            # a soft veil keeps the drifting wall out from under the text
            $veil = Box 462 44 ($W - 442) ($H - 88) $Light; $veil.Fill.Transparency = 0.08; $veil.SoftEdge.Type = 5
            $x = 500; $tw = 400
            $tt = Tx $Title 78 $Ink $x 92 $tw $heavyBold $heavy 1 0 1.08; Anim $tt $RISE $WITH 0.7 0.3; $y = [double]$tt.Top + [double]$tt.Height + 14
            if ($Sub) { $bar = Box $x $y 56 5 $Accent; Anim $bar $WIPE $WITH 0.4 0.8 4; $st = Tx $Sub 23 $Accent $x ($y + 18) $tw -1; Anim $st $FADE $WITH 0.5 0.9; $y += 18 + [double]$st.Height + 22 }
            for ($i = 0; $i -lt $n; $i++) {
                $dot = Box $x ($y + 9) 8 8 $Accent; $b = Tx $rows[$i][0] 19 $body ($x + 22) $y ($tw - 22) 0 $Font 1 0 1.3
                Anim (Grp @($dot, $b)) $FADE $WITH 0.5 (1.6 + 1.2 * $i); $y += [double]$b.Height + 12
            }
            if ($Note) { [void](Tx $Note 15 $gray $x 488 $tw) }
            $secs = [int]($Duration + 4)
        }
        'Ticket' {
            $paper = $Light; $rule = Mix $Light $Ink 0.28; $top = 138; $th = 334; $mx = 58; $mw = 590; $sx = $mx + $mw; $sw = 254
            if ($Kicker) { [void](Tx $Kicker 13 $Light 62 40 700 -1 $Font 1 4) }
            [void](Tx $Title 31 '#FFFFFF' 62 64 840 -1)
            $main = @((Box $mx $top $mw $th $paper 20)); Glow $main[0] $Dark 0.6 26 12
            $perf = Tag ($sl.Shapes.AddLine([single]$sx, [single]($top + 22), [single]$sx, [single]($top + $th - 22))); $perf.Line.ForeColor.RGB = (Rgb $rule); $perf.Line.Weight = 1.5; $perf.Line.DashStyle = 4; $main += $perf
            $r0 = $rows[0]
            $main += (Tx 'WHEN' 13 $Accent ($mx + 34) ($top + 26) 200 -1 $Font 1 4)
            $main += (Tx $r0[0] 118 $Ink ($mx + 28) ($top + 40) 340 $heavyBold $heavy 1 -3)
            if ($r0.Count -gt 1) { $main += (Tx $r0[1] 18 $body ($mx + 34) ($top + 190) 320 -1) }
            $main += (Box ($mx + 348) ($top + 28) 1.5 ($th - 56) $rule)
            for ($i = 1; $i -lt $n; $i++) {
                $yy = $top + 26 + ($i - 1) * (($th - 40) / ($n - 1))
                $main += (Tx $rows[$i][0] 13 $Accent ($mx + 374) $yy 200 -1 $Font 1 4); $main += (Tx $rows[$i][1] 19 $Ink ($mx + 374) ($yy + 22) 204 -1 $Font 1 0 1.25)
            }
            (Grp $main).Name = 'Main'
            $stub = @((Box $sx $top $sw $th $paper 20)); Glow $stub[0] $Dark 0.6 26 12
            $stub += (Tx 'ADMIT ONE' 13 $Accent ($sx + 28) ($top + 26) 200 -1 $Font 1 4)
            if ($cl.Count) { $stub += (Tx $cl[0] 40 $Ink ($sx + 26) ($top + 50) ($sw - 40) $heavyBold $heavy 1 0 1.05) }
            $yy = $top + 50 + $(if ($cl.Count) { [double]$stub[-1].Height + 10 } else { 0 })
            for ($i = 1; $i -lt $cl.Count; $i++) { $t = Tx $cl[$i] 17 $body ($sx + 28) $yy ($sw - 50); $stub += $t; $yy += [double]$t.Height + 4 }
            $bx = $sx + 28; $pat = 3, 1, 2, 1, 4, 1, 1, 2, 3, 1, 2, 4, 1, 1, 3, 2, 1, 2, 1, 4, 2, 1, 3, 1, 1, 2, 4, 1, 2, 1, 3
            for ($k = 0; $k -lt $pat.Count -and $bx -lt $sx + $sw - 30; $k++) { if ($k % 2 -eq 0) { $stub += (Box $bx ($top + $th - 86) ($pat[$k] * 2.2) 58 $Ink) }; $bx += $pat[$k] * 2.2 + $(if ($k % 2) { 2.4 } else { 0 }) }
            (Grp $stub).Name = 'Stub'
            if (-not $Duration) { $Duration = 1.0 }
            Steps 3; $s1 = $pres.Slides(1); $s2 = $pres.Slides(2); $s3 = $pres.Slides(3)
            foreach ($nm in 'Main', 'Stub') { $g = $s1.Shapes($nm); $g.Top = [single]([double]$g.Top + 470) }
            Turn $s1.Shapes('Main') -5; Turn $s1.Shapes('Stub') -5
            $g = $s3.Shapes('Stub'); Place $g ([double]$g.Left + 24) ([double]$g.Top + 12); Turn $g 5                # torn off
            $s3.SlideShowTransition.Duration = 0.5
            if ($Sub) {
                $script:sl = $s3; $kx = $mx + 276; $ky = $top + 276; $ink2 = $Accent
                $st = @((Ring $kx $ky 52 $ink2 3.5), (Ring $kx $ky 44 $ink2 1.25))
                $st += (Mid (Tx $Sub 20 $ink2 ($kx - 44) ($ky - 30) 88 $heavyBold $heavy 2 1 1.1) ($ky - 34) 68)
                $g = Grp $st; Turn $g -14
                $e = $sl.TimeLine.MainSequence.AddEffect($g, $FADE, 0, $AFTER); $e.Timing.Duration = 0.28; $e.Timing.TriggerDelayTime = 0.5
                $sc = $e.Behaviors.Add(3).ScaleEffect; $sc.FromX = 280; $sc.FromY = 280; $sc.ToX = 100; $sc.ToY = 100
            }
            Auto $s1 0.3; Auto $s2 0.7
        }
        'Programme' {
            $pw = 372; [void](Box 0 0 $pw $H $Dark)
            if ($Kicker) { [void](Tx $Kicker 12.5 $soft 46 40 300 -1 $Font 1 3) }
            [void](Tx $Title 29 '#FFFFFF' 46 62 300 -1)
            $y0 = 132; $step = [math]::Min(54, 372 / $n)
            (Box 28 ($y0 - ($step - 26) / 2) ($pw - 56) ($step - 6) $Accent 9).Name = 'Hi'
            for ($i = 0; $i -lt $n; $i++) {
                (Tx $rows[$i][0] 16 $dText 48 ($y0 + $i * $step + 1) 70 -1).Name = "T$i"; (Tx $rows[$i][1] 18 $dText 122 ($y0 + $i * $step) 220 -1).Name = "N$i"
            }
            $rx = 424; $ccx = 806; $ccy = 150; $cr = 80
            [void](Disc $ccx $ccy $cr (Mix $Light $Ink 0.05)); [void](Ring $ccx $ccy $cr $Accent 3.5)
            for ($k = 0; $k -lt 12; $k++) { $a = $k * [math]::PI / 6; [void](Disc ($ccx + 66 * [math]::Sin($a)) ($ccy - 66 * [math]::Cos($a)) $(if ($k % 3) { 1.8 } else { 3.6 }) $(if ($k % 3) { $gray } else { $Ink })) }
            (Needle $ccx $ccy 40 8 $Ink).Name = 'HourH'; (Needle $ccx $ccy 60 5 $Accent).Name = 'MinH'; [void](Disc $ccx $ccy 6 $Ink)
            for ($k = 0; $k -lt $n; $k++) {
                $r = $rows[$k]
                $set = @((Tx ('{0:00} / {1:00}' -f ($k + 1), $n) 14 $Accent $rx 58 260 -1 $Font 1 4), (Tx $r[0] 98 $Ink ($rx - 5) 78 300 $heavyBold $heavy 1 -2))
                $hd = Tx $r[1] 46 $AccentDark $rx 236 480 -1 $Font 1 0 1.1; $set += $hd
                if ($r.Count -gt 2) { $set += (Tx (($r[2..($r.Count - 1)]) -join "`r") 21 $body $rx ([double]$hd.Top + [double]$hd.Height + 16) 480 0 $Font 1 0 1.45) }
                if ($k -lt $n - 1) {
                    $set += (Box $rx 398 3 46 $Accent); $set += (Tx 'NEXT' 13 $Accent ($rx + 16) 397 200 -1 $Font 1 4)
                    $set += (Tx "$($rows[$k + 1][0])   $($rows[$k + 1][1])" 21 $Ink ($rx + 16) 417 440 -1)
                }
                (Grp $set).Name = "Info$k"
            }
            [void](Box $rx 470 482 4 (Mix $Light $Ink 0.12) 2); (Box $rx 470 482 4 $Accent 2).Name = 'Prog'
            [void](Tx $rows[0][0] 15 $gray $rx 482 200); [void](Tx $rows[$n - 1][0] 15 $gray ($rx + 282) 482 200 0 $Font 3)
            if (-not $Duration) { $Duration = 0.9 }
            Steps $n
            for ($s = 0; $s -lt $n; $s++) {
                $slide = $pres.Slides($s + 1); $hi = $slide.Shapes('Hi'); $hi.Top = [single]([double]$hi.Top + $s * $step)
                foreach ($nm in "T$s", "N$s") { $f = $slide.Shapes($nm).TextFrame.TextRange.Font; $f.Color.RGB = 0xFFFFFF; $f.Bold = -1 }
                $ang = Clock $rows[$s][0]; Turn $slide.Shapes('HourH') $ang[0]; Turn $slide.Shapes('MinH') $ang[1]
                $slide.Shapes('Prog').Width = [single](482 * ($s + 1) / $n)
                for ($k = 0; $k -lt $n; $k++) { if ($k -ne $s) { $slide.Shapes("Info$k").Delete() } }
            }
        }
        'Welcome' {
            $cx = $W / 2; $cy = 250; $ray = @(); $far = 760
            for ($k = 0; $k -lt 12; $k++) {
                $a = $k * [math]::PI / 6; $b = $a + [math]::PI / 12
                $ray += (Poly @(@($cx, $cy), @(($cx + $far * [math]::Cos($a)), ($cy + $far * [math]::Sin($a))), @(($cx + $far * [math]::Cos($b)), ($cy + $far * [math]::Sin($b)))) (Mix $Light $Accent 0.08))
            }
            $rays = Grp $ray; Spin $rays 30 $AFTER 16
            $glowD = Disc $cx $cy 250 $Light; $glowD.SoftEdge.Type = 6
            if ($Kicker) { $k1 = Tx $Kicker 15 $Accent 0 52 $W -1 $Font 2 5; Anim $k1 $FADE $WITH 0.5 0.1 }
            $by = 92; $bh = 52; $bw = 430; $bl = $cx - $bw / 2; $br = $cx + $bw / 2; $shade = Mix $Accent $Dark 0.3; $fold = Mix $Accent $Dark 0.65
            $tails = @((Poly @(@(($bl - 54), ($by + 13)), @(($bl + 26), ($by + 13)), @(($bl + 26), ($by + $bh + 13)), @(($bl - 54), ($by + $bh + 13)), @(($bl - 32), ($by + $bh / 2 + 13))) $shade),
                (Poly @(@(($br + 54), ($by + 13)), @(($br - 26), ($by + 13)), @(($br - 26), ($by + $bh + 13)), @(($br + 54), ($by + $bh + 13)), @(($br + 32), ($by + $bh / 2 + 13))) $shade),
                (Poly @(@($bl, ($by + $bh)), @(($bl + 26), ($by + $bh)), @(($bl + 26), ($by + $bh + 13))) $fold), (Poly @(@($br, ($by + $bh)), @(($br - 26), ($by + $bh)), @(($br - 26), ($by + $bh + 13))) $fold))
            Anim (Grp $tails) $FADE $WITH 0.4 1.0
            $h1 = Box $bl $by ($bw / 2 + 0.5) $bh $Accent; $h2 = Box $cx $by ($bw / 2) $bh $Accent
            Anim $h1 $WIPE $WITH 0.6 0.4 2; Anim $h2 $WIPE $WITH 0.6 0.4 4
            $bt = Mid (Tx $Sub 25 '#FFFFFF' $bl $by $bw -1 $Font 2 8) $by $bh; Anim $bt $FADE $WITH 0.5 0.9
            $big = [math]::Min(104, [math]::Floor(800 / (Ems $Title)))
            $nm = Tx $Title $big $Ink 0 ($by + $bh + 44) $W $heavyBold $heavy 2; Anim $nm $RISE $WITH 0.9 1.3; $y = [double]$nm.Top + [double]$nm.Height + 10
            $l1 = Box ($cx - 170) $y 170 2.5 $Accent; $l2 = Box $cx $y 170 2.5 $Accent; Anim $l1 $WIPE $WITH 0.5 2.0 2; Anim $l2 $WIPE $WITH 0.5 2.0 4
            $gem = Tag ($sl.Shapes.AddShape(4, [single]($cx - 8), [single]($y - 7), 16, 16)); $gem.Line.Visible = 0; $gem.Fill.ForeColor.RGB = (Rgb $Accent); Anim $gem $FADE $WITH 0.3 1.9
            if ($cl.Count) { $c1 = Tx $cl[0] 28 $AccentDark 0 ($y + 22) $W -1 $Font 2; Anim $c1 $FADE $WITH 0.6 2.3 }
            foreach ($p in @(150, 150, 20), @(812, 138, 14), @(108, 300, 12), @(850, 292, 18), @(760, 60, 10), @(204, 66, 11)) {               # four-point sparks that twinkle
                $sp = Tag ($sl.Shapes.AddShape(91, [single]($p[0] - $p[2]), [single]($p[1] - $p[2]), [single](2 * $p[2]), [single](2 * $p[2]))); $sp.Line.Visible = 0; $sp.Fill.ForeColor.RGB = (Rgb $Accent); $sp.Adjustments.Item(1) = 0.14
                Blink $sp (1.6 + 0.011 * $p[0]) 5
            }
            if ($n) {
                $gy = 424; $lab = Tx 'WITH US' 13 $Accent 0 ($gy - 2) $W -1 $Font 2 5; $rl = Box ($cx - 380) ($gy + 24) 760 1 (Mix $Light $Ink 0.2)
                Anim (Grp @($lab, $rl)) $FADE $WITH 0.5 2.8
                $cw = [math]::Min(230, 800 / $n); $gx = $cx - $cw * $n / 2
                for ($i = 0; $i -lt $n; $i++) {
                    $g = Grp @((Tx $rows[$i][0] 16 $body ($gx + $cw * $i) ($gy + 36) $cw 0 $Font 2), (Tx $rows[$i][1] 21 $Ink ($gx + $cw * $i) ($gy + 58) $cw -1 $Font 2))
                    Anim $g $FADE $WITH 0.5 (3.0 + 0.2 * $i)
                }
            }
            $secs = 9
        }
        'Speaker' {
            if (-not $Images.Count) { throw 'Speaker needs -Images <portrait.png>' }
            if ($Images.Count -gt 1) { $bg = Add-Cover $Images[1]; $bg.PictureFormat.Brightness = 0.56; $sc = Box 0 0 $W $H $Dark; Grad $sc $Dark $Dark 0 0.12 0.7 }
            $ax = 706; $ay = 286; $ar = 196
            $arch = Grp @((Disc $ax $ay $ar $Accent), (Box ($ax - $ar) $ay (2 * $ar) ($H - $ay + 4) $Accent)); Anim $arch $WIPE $AFTER 0.7 0.1 3
            $halo = Ring $ax $ay ($ar + 26) (Mix $Dark $Accent 0.6) 1.5 3; Anim $halo $FADE $WITH 0.6 0.5; Spin $halo 40 $WITH 14 0.5
            $p = Tag ($sl.Shapes.AddPicture($Images[0], 0, -1, 0, 0, -1, -1)); $p.LockAspectRatio = -1; $p.Height = 540; Place $p ($ax - [double]$p.Width / 2) 58
            Anim $p $RISE $WITH 0.8 0.5
            $x = 64; $tw = 430
            $dot = Disc ($x + 7) 60 7 $Accent; Blink $dot 1.0 10
            if ($Kicker) { $k1 = Tx $Kicker 14 $Accent ($x + 24) 51 400 -1 $Font 1 3; Anim $k1 $FADE $WITH 0.4 1.0 }
            $y = 80; if ($Sub) { $s1 = Tx $Sub 17 $dText $x $y $tw; Anim $s1 $FADE $WITH 0.4 1.1; $y += 36 }
            $tt = Tx $Title 46 '#FFFFFF' $x $y $tw $heavyBold $heavy 1 0 1.14; Anim $tt $RISE $WITH 0.7 1.2; $y += [double]$tt.Height + 16
            for ($i = 0; $i -lt $n; $i++) {
                $b = Grp @((Box $x ($y + 11) 14 3 $Accent), (Tx $rows[$i][0] 18 $dText ($x + 26) $y ($tw - 26) 0 $Font 1 0 1.25)); Anim $b $FADE $WITH 0.4 (1.8 + 0.25 * $i); $y += 31
            }
            $ly = 428; $nmW = (Ems $cl[0]) * 30 + 64; $sk = 16
            $bar1 = Poly @(@($x, $ly), @(($x + $nmW + $sk), $ly), @(($x + $nmW), ($ly + 54)), @($x, ($ly + 54))) $Accent; Anim $bar1 $WIPE $WITH 0.35 2.8 4
            $n1 = Mid (Tx $cl[0] 30 (Mix $Dark '#000000' 0.3) ($x + 22) $ly $nmW $heavyBold $heavy) $ly 54; Anim $n1 $FADE $WITH 0.3 3.05
            if ($cl.Count -gt 1) {
                $rw = (Ems $cl[1]) * 17 + 60
                $bar2 = Poly @(@($x, ($ly + 54)), @(($x + $rw + $sk * 0.6), ($ly + 54)), @(($x + $rw), ($ly + 90)), @($x, ($ly + 90))) '#FFFFFF'; Anim $bar2 $WIPE $WITH 0.35 3.15 4
                $n2 = Mid (Tx $cl[1] 17 $Ink ($x + 22) ($ly + 54) $rw -1) ($ly + 54) 36; Anim $n2 $FADE $WITH 0.3 3.4
            }
            $secs = 8
        }
        'Award' {
            $cx = $W / 2; $floorY = 478
            [void](Box 0 452 $W 90 (Mix $Dark '#000000' 0.45))
            (Beam $cx -30 660 14 120 '#FFFFFF' 0.5).Name = 'Beam'
            $pool = Tag ($sl.Shapes.AddShape(9, [single]($cx - 150), [single]($floorY - 26), 300, 52)); $pool.Line.Visible = 0; $pool.Fill.ForeColor.RGB = 0xFFFFFF; $pool.Fill.Transparency = 0.72; $pool.SoftEdge.Type = 5; $pool.Name = 'Pool'
            foreach ($side in -1, 1) {
                $fold = @(); for ($k = 0; $k -lt 6; $k++) { $fx0 = $(if ($side -lt 0) { $k * 28 - 8 } else { $W - ($k + 1) * 28 + 8 }); $fd = Box $fx0 -10 29 ($H + 20) $Dark; Grad $fd (Mix $Dark $Accent $(if ($k % 2) { 0.2 } else { 0.06 })) (Mix $Dark '#000000' 0.5) 90; $fold += $fd }
                (Grp $fold).Name = $(if ($side -lt 0) { 'CurtL' } else { 'CurtR' })
            }
            $ask = @(); if ($Kicker) { $ask += (Tx $Kicker 17 $Accent 0 150 $W -1 $Font 2 5) }
            $ask += (Tx $Sub 60 $Light 0 188 $W $heavyBold $heavy 2); for ($k = -1; $k -le 1; $k++) { $ask += (Disc ($cx + 30 * $k) 300 6 $Accent) }
            (Grp $ask).Name = 'Ask'
            if (-not $Duration) { $Duration = 1.15 }
            Steps 4; $swing = 27, -27, 13, 0
            for ($s = 0; $s -lt 4; $s++) {
                $slide = $pres.Slides($s + 1); Turn $slide.Shapes('Beam') $swing[$s]
                $pl = $slide.Shapes('Pool'); $px = $cx - ($floorY + 30) * [math]::Tan($swing[$s] * [math]::PI / 180)
                if ($s -eq 3) { $pl.Width = 720; $pl.Height = 70; $pl.Fill.Transparency = 0.8 }; Place $pl ($px - [double]$pl.Width / 2) ($floorY - [double]$pl.Height / 2)
                if ($s -lt 3) { Auto $slide 0.15 }
            }
            $script:sl = $pres.Slides(4); $sl.Shapes('Ask').Delete()
            $c = $sl.Shapes('CurtL'); $c.Left = [single]([double]$c.Left - 96); $c = $sl.Shapes('CurtR'); $c.Left = [single]([double]$c.Left + 96)
            $lcy = 256; $rx = 356; $ry = 190; $nodes = 16
            for ($q = 0; $q -lt $nodes; $q++) {                                                                  # two laurel branches, leaf pairs from the bottom up
                $ph = (240 - $q * 7.6) * [math]::PI / 180; $ph2 = $ph - 0.04; $x = $cx + $rx * [math]::Cos($ph); $y = $lcy - $ry * [math]::Sin($ph)
                $ang = [math]::Atan2((- $ry * [math]::Sin($ph2)) - (- $ry * [math]::Sin($ph)), $rx * [math]::Cos($ph2) - $rx * [math]::Cos($ph)) * 180 / [math]::PI
                $f = 1 - 0.4 * $q / ($nodes - 1); $ll = 36 * $f; $lw = 13 * $f; $set = @()
                foreach ($off in -34, 34) {
                    $a = ($ang + $off) * [math]::PI / 180; $lx = $x + $ll * 0.5 * [math]::Cos($a); $ly = $y + $ll * 0.5 * [math]::Sin($a)
                    foreach ($side in 1, -1) {
                        $lf = Tag ($sl.Shapes.AddShape(9, [single]($(if ($side -gt 0) { $lx } else { 2 * $cx - $lx }) - $ll / 2), [single]($ly - $lw / 2), [single]$ll, [single]$lw)); $lf.Line.Visible = 0
                        $lf.Fill.ForeColor.RGB = (Rgb $(if ($off -lt 0) { $Accent } else { Mix $Accent $Dark 0.25 })); Turn $lf $(if ($side -gt 0) { $ang + $off } else { 180 - $ang - $off }); $set += $lf
                    }
                }
                Anim (Grp $set) $FADE $(if ($q) { $WITH } else { $AFTER }) 0.25 (0.5 + 0.07 * $q)
            }
            $medal = @((Poly @(@(($cx - 30), 20), @(($cx - 8), 20), @(($cx + 10), 78), @(($cx - 12), 78)) (Mix $Accent $Dark 0.35)), (Poly @(@(($cx + 30), 20), @(($cx + 8), 20), @(($cx - 10), 78), @(($cx + 12), 78)) (Mix $Accent $Dark 0.15)),
                (Disc $cx 92 30 $Accent), (Ring $cx 92 23 (Mix $Accent $Dark 0.45) 1.5))
            $star = Tag ($sl.Shapes.AddShape(92, [single]($cx - 15), [single](92 - 16), 30, 30)); $star.Line.Visible = 0; $star.Fill.ForeColor.RGB = (Rgb (Mix $Accent $Dark 0.6)); $medal += $star
            Anim (Grp $medal) $RISE $WITH 0.6 0.6
            $y = 136
            if ($Kicker) { $t0 = Tx $Kicker 16 $Accent 0 $y $W -1 $Font 2 5; Anim $t0 $FADE $WITH 0.4 0.9; $y += 28 }
            if ($cl.Count) { $t1 = Tx $cl[0] 34 $Accent 0 $y $W $heavyBold $heavy 2 10; Anim $t1 $FADE $WITH 0.4 1.1; $y += 48 }
            $big = [math]::Min(96, [math]::Floor(520 / (Ems $Title)))
            $nm = Tx $Title $big '#FFFFFF' 0 $y $W $heavyBold $heavy 2 4; Glow $nm $Accent 0.45 22 0; Anim $nm $RISE $WITH 1.0 1.9; $y += [double]$nm.Height + 6
            if ($cl.Count -gt 1) { $t2 = Tx $cl[1] 25 $Light 0 $y $W -1 $Font 2; Anim $t2 $FADE $WITH 0.5 2.7; $y += 42 }
            if ($cl.Count -gt 2) { $t3 = Tx $cl[2] 18 (Mix $Dark '#FFFFFF' 0.85) 0 $y $W 0 $Font 2 0 1.35; Anim $t3 $FADE $WITH 0.5 3.0 }
            $secs = 6
        }
        'Sessions' {
            Add-Grid (Mix $Dark '#FFFFFF' 0.04)
            $mono = 'Consolas'; $slots = @($Sub -split '\|' | Where-Object { $_ }); $nc = $slots.Count
            $k1 = Tx $(if ($Kicker) { "> $Kicker" } else { '> SESSION GUIDE' }) 15 $Accent 65 40 500 -1 $mono 1 2
            $cur = Box (65 + [double]$k1.TextFrame.TextRange.BoundWidth + 8) 41 10 18 $Accent; Blink $cur 0.2 12
            [void](Tx $Title 31 '#FFFFFF' 65 66 700 -1)
            if ($Note) { $nw = (Ems $Note) * 15 + 34; $np = Box ($W - 65 - $nw) 60 $nw 34 $Dark 17; $np.Line.Visible = -1; $np.Line.ForeColor.RGB = (Rgb $Accent); $np.Line.Weight = 1.25; [void](Mid (Tx $Note 15 $Accent ($W - 65 - $nw) 60 $nw -1 $Font 2) 60 34) }
            $gx = 252; $gr = $W - 65; $gap = 10; $cw = ($gr - $gx - $gap * ($nc - 1)) / $nc; $gy = 186; $gb = 492; $ch = ($gb - $gy - $gap * ($n - 1)) / $n
            $rail = Box $gx 168 ($gr - $gx) 1.5 $dLine; Anim $rail $WIPE $WITH 0.8 0.2 4
            for ($j = 0; $j -lt $nc; $j++) {
                $x = $gx + ($cw + $gap) * $j; $g = Grp @((Disc ($x + 5) 168.5 5 $Accent), (Tx $slots[$j] 18 $Accent ($x + 18) 134 ($cw - 18) -1 $mono)); Anim $g $FADE $WITH 0.3 (0.3 + 0.2 * $j)
            }
            for ($i = 0; $i -lt $n; $i++) {
                $y = $gy + ($ch + $gap) * $i; $r = $rows[$i]
                $hex = Tag ($sl.Shapes.AddShape(10, 65, [single]($y + $ch / 2 - 24), 54, 48)); $hex.Line.Visible = 0; $hex.Fill.ForeColor.RGB = (Rgb $Accent)
                $set = @($hex, (Mid (Tx $r[0] 24 $Dark 65 ($y + $ch / 2 - 24) 54 $heavyBold $heavy 2) ($y + $ch / 2 - 24) 48), (Tx $r[1] 17 '#FFFFFF' 130 ($y + $ch / 2 - 30) 116 -1 $Font 1 0 1.15))
                if ($r.Count -gt 2) { $set += (Tx $r[2] 15 $soft 130 ([double]$set[2].Top + [double]$set[2].Height + 3) 116 0 $mono) }
                $g = Grp $set; $g.Top = [single]($y + ($ch - [double]$g.Height) / 2); Anim $g $FADE $WITH 0.4 (0.5 + 0.15 * $i)
            }
            $scanT = 2.6; $t0 = 1.3
            $scan = Box ($gx - 6) ($gy - 6) 3 ($gb - $gy + 12) $Accent; Glow $scan $Accent 0.2 14 0
            Anim $scan $FADE $WITH 0.2 ($t0 - 0.2); Glide $scan ($gr - $gx + 10) 0 $WITH $scanT $t0
            $sess = Rows $Cards
            for ($i = 0; $i -lt $n; $i++) {
                for ($j = 0; $j -lt $nc; $j++) {
                    $q = $i * $nc + $j; if ($q -ge $sess.Count) { continue }
                    $c = $sess[$q]; $hot = $c[0].StartsWith('*'); $head = $c[0].TrimStart('*'); $x = $gx + ($cw + $gap) * $j; $y = $gy + ($ch + $gap) * $i
                    $card = Box $x $y $cw $ch $(if ($hot) { $Accent } else { Mix $Dark '#FFFFFF' 0.1 }) 10
                    if (-not $hot) { $card.Line.Visible = -1; $card.Line.ForeColor.RGB = (Rgb (Mix $Dark $Accent 0.35)); $card.Line.Weight = 0.75 }
                    $set = @($card, (Tx $head 16.5 $(if ($hot) { Mix $Dark '#000000' 0.3 } else { '#FFFFFF' }) ($x + 14) ($y + 11) ($cw - 24) -1 $Font 1 0 1.18))
                    if ($c.Count -gt 1) { $wt = Tx $c[1] 15 $(if ($hot) { Mix $Dark '#000000' 0.1 } else { $soft }) ($x + 14) ($y + $ch - 31) ($cw - 24); $set += $wt }
                    Anim (Grp $set) $FADE $WITH 0.35 ($t0 + $scanT * (($x - $gx) + $cw * 0.5) / ($gr - $gx) + 0.08 * $i)
                }
            }
            $e = $sl.TimeLine.MainSequence.AddEffect($scan, $FADE, 0, $WITH); $e.Exit = -1; $e.Timing.Duration = 0.3; $e.Timing.TriggerDelayTime = [single]($t0 + $scanT)
            $secs = 7
        }
    }
    Save-Deck $Out $Preview $secs
    "saved $Out"
} catch { Close-Deck; throw }
