# Government performance / briefing slides: every chart is drawn from native shapes and plays by itself. Same colour / font parameters as scene_page.ps1.
#   BeforeAfter  two race lanes: the "after" bar overtakes to a chequered finish line, a slanted stamp lands last
#                -Items 'before label;value|after label;value'  -Sub unit  -Cards 'gap text;stamp line//stamp line'
#   Gauges       three speedometers: the needle sweeps to the actual value past a target tick   -Items 'label;actual;target;unit;badge|...' (3)
#   Board        dark split-flap scoreboard, tiles drop in one by one, a ticker runs along the bottom   -Items 'label;value;delta|...' (4)  -Sub 'ticker text'
#   Radar        two polygons (before / after) grow out of the centre of a radar that is scanned in   -Items 'axis;before;after|...' (5-6, 0..100)  -Cards 'before name|after name'
#   Slope        slope chart between two years, the largest rise is drawn last with a dot riding the line
#                -Items 'name;v1;v2|...' (4-5)  -Sub 'year1;year2'  -Cards 'big delta;caption'
#   Keynote      one message under three stage lights, three proofs, a citizen quote   -Title 'line//line'  -Items 'number;label|...' (3)  -Cards 'quote;who'
#   TaskBoard    task rows with a three-lamp status signal, progress bars and one "planned progress" line across all rows
#                -Items 'task;detail;percent;status(1 ok, 2 watch, 3 late);owner|...' (4-5)  -Plan '78;label'  -Cards 'head|head|head|head|ok|watch|late'  -Sub 'date'
#   AIStack      a chip outline wired to four stacked layers by circuit traces that draw themselves   -Items 'layer;programme;figure;unit|...' (4)  -Sub 'chip//label'  -Cards 'line|line'
#   Ramp         one huge figure, and stacks of unit tiles that pile up year by year under a dashed trajectory
#                -Title 'hero figure'  -Sub 'message'  -Items 'year;tiles;shown value;caption|...' (3-4)  -Cards 'value;label|value;label'
# Example: .\gov_data.ps1 -Layout Gauges -Title 'Three indicators' -Items 'A;86;80;%;+6|B;71;75;%;-4|C;93;90;%;+3' -Out C:\work\g.pptx -Preview C:\work\g.mp4
param(
    [Parameter(Mandatory)][ValidateSet('BeforeAfter', 'Gauges', 'Board', 'Radar', 'Slope', 'Keynote', 'TaskBoard', 'AIStack', 'Ramp')][string]$Layout,
    [string]$Kicker = '',
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Plan = '',                            # TaskBoard: 'percent;label' of the planned-progress line
    [string]$Note = '',
    [string]$Accent = '#2447B8',
    [string]$AccentDark = '',
    [string]$Light = '#F3F5F9',
    [string]$Dark = '#0F1A3F',
    [string]$Ink = '#1B2233',
    [string]$Font = 'Malgun Gothic',
    [string]$HeavyFont = '',
    [double]$Duration = 0.9,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent
. (Join-Path $PSScriptRoot 'scene_helpers.ps1')
$ZOOM = 53                                                              # faded zoom: grows out of the shape's own centre
function Rgb([string]$hex) { ConvertTo-Bgr $hex }
function Poly($pts, [string]$hex) {                                    # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
    $fb.AddNodes(0, 0, [single]$pts[0][0], [single]$pts[0][1]); $s = Tag ($fb.ConvertToShape()); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex); $s
}
function Ln($pts, [string]$hex, $wt = 1.5, [int]$dash = 1) {           # open line through the points (2 points = a straight line)
    if ($pts.Count -eq 2) { $s = Tag ($sl.Shapes.AddLine([single]$pts[0][0], [single]$pts[0][1], [single]$pts[1][0], [single]$pts[1][1])) }
    else {
        $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1]); for ($q = 1; $q -lt $pts.Count; $q++) { $fb.AddNodes(0, 0, [single]$pts[$q][0], [single]$pts[$q][1]) }
        $s = Tag ($fb.ConvertToShape()); $s.Fill.Visible = 0
    }
    $s.Line.ForeColor.RGB = (Rgb $hex); $s.Line.Weight = [single]$wt; $s.Line.DashStyle = $dash; $s
}
function Arc($cx, $cy, $rad, $a1, $a2, $thick, [string]$hex) {        # block arc, angles clockwise from east (top = 270)
    $s = Tag ($sl.Shapes.AddShape(20, [single]($cx - $rad), [single]($cy - $rad), [single](2 * $rad), [single](2 * $rad))); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (Rgb $hex)
    $s.Adjustments.Item(1) = [single]($a1 % 360); $s.Adjustments.Item(2) = [single]($a2 % 360); $s.Adjustments.Item(3) = [single]($thick / (2 * $rad)); $s
}
function Ghost($cx, $cy, $rad) { $d = Disc $cx $cy $rad '#FFFFFF'; $d.Fill.Transparency = 1; $d }   # invisible disc: grouped with a shape it moves the group's centre to (cx, cy)
function Spin($shape, $deg, [int]$trig, $dur, $delay = 0) { $e = $sl.TimeLine.MainSequence.AddEffect($shape, 0, 0, $trig); $e.Behaviors.Add(4).RotationEffect.By = [single]$deg; $e.Timing.Duration = [single]$dur; $e.Timing.TriggerDelayTime = [single]$delay }
function Glide($shape, $dx, $dy, [int]$trig, $dur, $delay = 0) {       # straight motion path, in points
    $e = $sl.TimeLine.MainSequence.AddEffect($shape, 0, 0, $trig); $e.Behaviors.Add(1).MotionEffect.Path = ('M 0 0 L {0:F4} {1:F4} E' -f ($dx / $W), ($dy / $H))
    $e.Timing.Duration = [single]$dur; $e.Timing.TriggerDelayTime = [single]$delay
}
function GrowIn($shape, [int]$trig, $dur, $delay = 0) {                 # fades in while it scales up about its own centre
    Anim $shape $FADE $trig 0.3 $delay
    $e = $sl.TimeLine.MainSequence.AddEffect($shape, 0, 0, $WITH); $b = $e.Behaviors.Add(3).ScaleEffect; $b.FromX = 4; $b.FromY = 4; $b.ToX = 100; $b.ToY = 100
    $e.Timing.Duration = [single]$dur; $e.Timing.TriggerDelayTime = [single]$delay
}
function Spread($ys, $gap) {                                           # push label positions apart, keeping their order
    $idx = @(0..($ys.Count - 1) | Sort-Object { [double]$ys[$_] }); $o = @($ys | ForEach-Object { [double]$_ })
    for ($k = 1; $k -lt $idx.Count; $k++) { if ($o[$idx[$k]] - $o[$idx[$k - 1]] -lt $gap) { $o[$idx[$k]] = $o[$idx[$k - 1]] + $gap } }
    , $o
}

try {
    $onDark = $Layout -in 'Board', 'Radar', 'Keynote', 'AIStack'
    New-Deck $(if ($onDark) { $Dark } else { $Light })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6
    $rows = Rows $Items; $n = $rows.Count; $cds = Rows $Cards; $secs = 8
    $rail = Mix $Light $Ink 0.09

    switch ($Layout) {
        'BeforeAfter' {
            Head; $rb = $rows[0]; $ra = $rows[1]; $bv = [double]$rb[1]; $av = [double]$ra[1]; $vmax = [math]::Max($bv, $av)
            [void](Tx $rb[0] 17 $gray 65 176 250 -1)
            $t1 = Tx "$($rb[1])$Sub" 84 (Mix $Ink '#FFFFFF' 0.55) 62 196 300 $heavyBold $heavy
            if ($Sub) { $t1.TextFrame.TextRange.Characters($rb[1].Length + 1, $Sub.Length).Font.Size = 40 }
            $arrow = Poly @(@(322, 234), @(372, 234), @(372, 218), @(404, 244), @(372, 270), @(372, 254), @(322, 254)) $Accent
            [void](Tx $ra[0] 17 $AccentDark 436 146 330 -1)
            $t2 = Tx "$($ra[1])$Sub" 150 $Accent 430 134 20 $heavyBold $heavy; $t2.TextFrame.WordWrap = 0
            if ($Sub) { $t2.TextFrame.TextRange.Characters($ra[1].Length + 1, $Sub.Length).Font.Size = 64 }
            if ([double]$t2.Width -gt 400) { $t2.TextFrame.TextRange.Font.Size = [single][math]::Floor(150 * 400 / [double]$t2.Width) }
            $lx = 190; $tw = 705; $lh = 42; $y1 = 352; $y2 = 410; $full = 0.86 * $tw; $w1 = $full * $bv / $vmax; $w2 = $full * $av / $vmax; $fin = $lx + $full
            foreach ($y in $y1, $y2) { [void](Box $lx $y $tw $lh $rail 6) }
            [void](Mid (Tx $rb[0] 17 $body 65 $y1 118 -1) $y1 $lh); [void](Mid (Tx $ra[0] 17 $AccentDark 65 $y2 118 -1) $y2 $lh)
            $flag = @(); $sq = 7                                                                                                  # chequered finish line across both lanes
            for ($q = 0; $q * $sq -lt ($y2 + $lh - $y1 + 16); $q++) { for ($c = 0; $c -lt 2; $c++) { $flag += (Box ($fin + $c * $sq) ($y1 - 8 + $q * $sq) $sq $sq $(if (($q + $c) % 2) { $Ink } else { '#FFFFFF' })) } }
            [void](Grp $flag)
            $b1 = Box $lx $y1 $w1 $lh (Mix $Ink '#FFFFFF' 0.55) 6; $b2 = Box $lx $y2 $w2 $lh $Accent 6
            $v1 = Mid (Tx "$($rb[1])$Sub" 17 '#FFFFFF' $lx $y1 ($w1 - 12) -1 $Font 3) $y1 $lh; $v2 = Mid (Tx "$($ra[1])$Sub" 19 '#FFFFFF' $lx $y2 ($w2 - 12) -1 $Font 3) $y2 $lh
            $lo = [math]::Min($w1, $w2); $gy = $(if ($bv -le $av) { $y1 } else { $y2 }) + $lh / 2
            $gap = Ln @(@(($lx + $lo + 10), $gy), @(($fin - 8), $gy)) $AccentDark 1.5 4; $gap.Line.BeginArrowheadStyle = 2; $gap.Line.EndArrowheadStyle = 2
            $set = @($gap)
            if ($n -ge 1 -and $cds.Count) { $gt = $cds[0][0]; $pw = Wd $gt 19; $set += Pill ($lx + $lo + ($full - $lo) / 2 - $pw / 2) ($gy - 16) $pw 32 $AccentDark $gt 19 '#FFFFFF' }
            $gapG = Grp $set
            Anim $arrow $WIPE $AFTER 0.4 0.2 4; Anim $t2 $FADE $AFTER 0.5
            Anim $b1 $WIPE $AFTER 0.7 0.1 4; Anim $v1 $FADE $WITH 0.3 0.6; Anim $b2 $WIPE $WITH 1.7 0 4; Anim $v2 $FADE $AFTER 0.3
            Anim $gapG $WIPE $AFTER 0.5 0 4
            if ($cds.Count -and $cds[0].Count -gt 1) {
                $sx = 842; $sy = 108; $o1 = Disc $sx $sy 52 $Light; $o1.Fill.Visible = 0; $o1.Line.Visible = -1; $o1.Line.ForeColor.RGB = (Rgb $AccentDark); $o1.Line.Weight = 3.5
                $o2 = Disc $sx $sy 44 $Light; $o2.Fill.Visible = 0; $o2.Line.Visible = -1; $o2.Line.ForeColor.RGB = (Rgb $AccentDark); $o2.Line.Weight = 1
                $st = Mid (Tx $cds[0][1] 19 $AccentDark ($sx - 44) ($sy - 44) 88 -1 $Font 2 0 1.1) ($sy - 44) 88
                $stamp = Grp @($o1, $o2, $st); $stamp.Rotation = 348; Anim $stamp $ZOOM $AFTER 0.3 0.2
            }
            Foot; $secs = 8
        }
        'Gauges' {
            Head
            for ($i = 0; $i -lt [math]::Min(3, $n); $i++) {
                $r = $rows[$i]; $act = [double]$r[1]; $tgt = [double]$r[2]; $u = $(if ($r.Count -gt 3) { $r[3] } else { '' }); $cx = 185 + 295 * $i; $cy = 322; $R0 = 124; $th = 24
                $mx = [math]::Max($act, $tgt); $mag = [math]::Pow(10, [math]::Floor([math]::Log10($mx * 1.2)))
                $top = $(if ($mx -le 100 -and $mx * 1.2 -ge 100) { 100 } else { [math]::Ceiling($mx * 1.2 / $mag * 2) / 2 * $mag }); $fa = $act / $top; $ft = $tgt / $top
                [void](Arc $cx $cy $R0 180 360 $th $rail)
                $fill = Arc $cx $cy $R0 180 (180 + 180 * $fa) $th $Accent; $over = $null
                if ($act -gt $tgt) { $over = Arc $cx $cy $R0 (180 + 180 * $ft) (180 + 180 * $fa) $th $AccentDark }
                $ang = (180 + 180 * $ft) * [math]::PI / 180; $ca = [math]::Cos($ang); $sa = [math]::Sin($ang)
                [void](Ln @(@(($cx + ($R0 - $th - 6) * $ca), ($cy + ($R0 - $th - 6) * $sa)), @(($cx + ($R0 + 9) * $ca), ($cy + ($R0 + 9) * $sa))) $Ink 2.5)
                [void](Tx "$($r[2])$u" 15 $Ink ($cx + ($R0 + 30) * $ca - 45) ($cy + ($R0 + 30) * $sa - 10) 90 -1 $Font 2)
                [void](Tx '0' 15 $gray ($cx - $R0 - 8) ($cy + 7) 40 0 $Font 2); [void](Tx "$top" 15 $gray ($cx + $R0 - 32) ($cy + 7) 40 0 $Font 2)
                $needle = Grp @((Poly @(@(($cx - $R0 + $th + 8), $cy), @($cx, ($cy - 5)), @($cx, ($cy + 5))) $Ink), (Ghost $cx $cy ($R0 - $th)))     # points west; spun about the pivot
                [void](Disc $cx $cy 10 $Ink); [void](Disc $cx $cy 4 $Light)
                $val = Tx "$($r[1])$u" 46 $(if ($act -ge $tgt) { $AccentDark } else { $Ink }) ($cx - 125) ($cy + 26) 250 $heavyBold $heavy 2
                if ($u) { $val.TextFrame.TextRange.Characters($r[1].Length + 1, $u.Length).Font.Size = 24 }
                [void](Tx $r[0] 19 $Ink ($cx - 135) ($cy + 88) 270 -1 $Font 2)
                $dur = 0.5 + 1.2 * $fa
                Anim $fill $WIPE $AFTER 1.5 0.15 4; if ($over) { Anim $over $WIPE $WITH 1.5 0.15 4 }; Spin $needle (180 * $fa) $WITH $dur 0.15; Anim $val $FADE $WITH 0.4 ($dur - 0.2)
                if ($r.Count -gt 4) {
                    $pw = Wd $r[4] 15; $on = $act -ge $tgt
                    $badge = Grp (Pill ($cx - $pw / 2) ($cy + 122) $pw 30 $(if ($on) { $Accent } else { $rail }) $r[4] 15 $(if ($on) { '#FFFFFF' } else { $Ink })); Anim $badge $FADE $AFTER 0.3
                }
            }
            Foot; $secs = 9
        }
        'Board' {
            Head $true; [void](Box 50 122 860 316 $dCard 14); $pit = (Mix $Dark '#000000' 0.5); $k = 0
            $tk = Tx ((1..3 | ForEach-Object { $Sub }) -join "   $([char]0x25C6)   ") 18 $Dark 40 0 20 -1; $tk.TextFrame.WordWrap = 0
            $band = Box 0 456 $W 44 $Accent; $tk.ZOrder(0); [void](Mid $tk 456 44)
            Glide $tk (-0.5 * $W) 0 $AFTER 15
            for ($i = 0; $i -lt [math]::Min(4, $n); $i++) {
                $r = $rows[$i]; $x0 = 74 + ($i % 2) * 430; $y0 = 138 + [math]::Floor($i / 2) * 150; $tw = 62; $tH = 92
                [void](Tx $r[0] 17 $dText $x0 $y0 250 -1); if ($r.Count -gt 2) { [void](Tx $r[2] 17 $Accent ($x0 + 190) $y0 200 -1 $Font 3) }
                $j = 0
                foreach ($ch in $r[1].ToCharArray()) {                                                                             # one flap per character, hinge line across the middle
                    $x = $x0 + $j * ($tw + 5); $y = $y0 + 32
                    $flap = Grp @((Box $x $y $tw $tH $pit 7), (Mid (Tx "$ch" 60 $Accent $x $y $tw $heavyBold $heavy 2) $y $tH), (Box $x ($y + $tH / 2 - 1) $tw 2 $Dark))
                    Anim $flap $WIPE $WITH 0.16 (0.4 + 0.13 * $k) 1; $j++; $k++
                }
            }
            Foot $true; $secs = 16
        }
        'Radar' {
            Head $true; $cx = 292; $cy = 312; $R0 = 150; $pts = @{ b = @(); a = @() }; $grid = @(); $best = 0
            for ($ring = 1; $ring -le 4; $ring++) { $g = @(); for ($i = 0; $i -le $n; $i++) { $t = (-90 + 360 * $i / $n) * [math]::PI / 180; $g += , @(($cx + $R0 * $ring / 4 * [math]::Cos($t)), ($cy + $R0 * $ring / 4 * [math]::Sin($t))) }; $grid += (Ln $g $dLine 0.75) }
            for ($i = 0; $i -lt $n; $i++) {
                $t = (-90 + 360 * $i / $n) * [math]::PI / 180; $c = [math]::Cos($t); $s = [math]::Sin($t); $r = $rows[$i]
                $grid += (Ln @(@($cx, $cy), @(($cx + $R0 * $c), ($cy + $R0 * $s))) $dLine 0.75)
                $pts.b += , @(($cx + $R0 * [double]$r[1] / 100 * $c), ($cy + $R0 * [double]$r[1] / 100 * $s)); $pts.a += , @(($cx + $R0 * [double]$r[2] / 100 * $c), ($cy + $R0 * [double]$r[2] / 100 * $s))
                $al = $(if ([math]::Abs($c) -lt 0.3) { 2 } elseif ($c -gt 0) { 1 } else { 3 }); $lx = $cx + ($R0 + 14) * $c; $ly = $cy + ($R0 + 14) * $s
                [void](Tx $r[0] 16 '#FFFFFF' $(if ($al -eq 2) { $lx - 80 } elseif ($al -eq 1) { $lx } else { $lx - 160 }) $(if ($s -lt -0.3) { $ly - 22 } elseif ($s -gt 0.3) { $ly } else { $ly - 11 }) 160 -1 $Font $al)
                if (([double]$r[2] - [double]$r[1]) -gt ([double]$rows[$best][2] - [double]$rows[$best][1])) { $best = $i }
            }
            $gridG = Grp $grid; $cold = Mix $Dark '#FFFFFF' 0.6
            $pb = Poly $pts.b $cold; $pb.Fill.Transparency = 0.72; $pb.Line.Visible = -1; $pb.Line.ForeColor.RGB = (Rgb $cold); $pb.Line.Weight = 1.5; $pb.Line.DashStyle = 4
            $gb = Grp @($pb, (Ghost $cx $cy $R0))
            $pa = Poly $pts.a $Accent; $pa.Fill.Transparency = 0.6; $pa.Line.Visible = -1; $pa.Line.ForeColor.RGB = (Rgb $Accent); $pa.Line.Weight = 2.5
            $dots = @($pa, (Ghost $cx $cy $R0)); foreach ($p in $pts.a) { $dots += (Disc $p[0] $p[1] 4.5 $Accent) }
            $ga = Grp $dots
            Anim $gridG $WHEEL $AFTER 0.9 0.2; GrowIn $gb $AFTER 0.8; GrowIn $ga $AFTER 1.1 0.2
            $px = 548; $py = 148; $pitch = [math]::Min(52, 300 / $n)
            if ($cds.Count -gt 1) {
                [void](Tx $cds[0][0] 15 $cold ($px + 172) $py 70 -1 $Font 3); [void](Tx $cds[1][0] 15 $Accent ($px + 250) $py 70 -1 $Font 3)
            }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $y = $py + 30 + $i * $pitch; $d = [double]$r[2] - [double]$r[1]; $set = @()
                if ($i -eq $best) { $hl = Box ($px - 14) ($y - 5) 372 ($pitch - 6) $dCard 8; $hl.Line.Visible = -1; $hl.Line.ForeColor.RGB = (Rgb $Accent); $hl.Line.Weight = 1.25; $set += $hl }
                $rh = $pitch - 16
                $set += (Mid (Tx $r[0] 17 '#FFFFFF' $px $y 170 -1) $y $rh); $set += (Mid (Tx $r[1] 17 $cold ($px + 172) $y 70 0 $Font 3) $y $rh)
                $set += (Mid (Tx $r[2] 22 $Accent ($px + 250) $y 70 $heavyBold $heavy 3) $y $rh)
                $set += (Mid (Tx ('{0}{1}' -f $(if ($d -ge 0) { '+' } else { '' }), $d) 15 $(if ($i -eq $best) { $Accent } else { $dText }) ($px + 322) $y 30 -1 $Font 3) $y $rh)
                Anim (Grp $set) $FADE $(if ($i) { $WITH } else { $AFTER }) 0.35 (0.12 * $i)
            }
            Foot $true; $secs = 8
        }
        'Slope' {
            $yrs = @($Sub -split ';'); $x1 = 548; $x2 = 768; $cTop = 158; $cBot = 452
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 50 300 -1 $Font 1 3) }
            $tt = Tx $Title 30 $Ink 65 76 300 -1 $Font 1 0 1.2; $y = [double]$tt.Top + [double]$tt.Height + 26
            [void](Box 65 $y 46 4 $Accent 2); $hero = @()
            if ($cds.Count) { $hero += (Tx $cds[0][0] 88 $Accent 62 ($y + 16) 310 $heavyBold $heavy); if ($cds[0].Count -gt 1) { $hero += (Tx $cds[0][1] 17 $body 65 ($y + 124) 290 0 $Font 1 0 1.4) } }
            $all = @($rows | ForEach-Object { [double]$_[1]; [double]$_[2] }); $lo = ($all | Measure-Object -Minimum).Minimum; $hi = ($all | Measure-Object -Maximum).Maximum
            $ya = @(); $yb = @(); $best = 0
            for ($i = 0; $i -lt $n; $i++) {
                $ya += $cBot - ([double]$rows[$i][1] - $lo) / ($hi - $lo) * ($cBot - $cTop); $yb += $cBot - ([double]$rows[$i][2] - $lo) / ($hi - $lo) * ($cBot - $cTop)
                if (([double]$rows[$i][2] - [double]$rows[$i][1]) -gt ([double]$rows[$best][2] - [double]$rows[$best][1])) { $best = $i }
            }
            $la = Spread $ya 28; $lb = Spread $yb 30; $mute = Mix $Ink $Light 0.6
            foreach ($p in @($x1, $yrs[0]), @($x2, $yrs[1])) { [void](Ln @(@($p[0], ($cTop - 22)), @($p[0], ($cBot + 22))) (Mix $Light $Ink 0.2) 1.25); [void](Tx $p[1] 21 $Ink ($p[0] - 60) 100 120 -1 $Font 2) }
            $order = @(0..($n - 1) | Where-Object { $_ -ne $best }) + $best
            foreach ($i in $order) {
                $r = $rows[$i]; $on = ($i -eq $best); $col = $(if ($on) { $Accent } else { $mute })
                [void](Tx "$($r[0])   $($r[1])" 16 $(if ($on) { $Ink } else { $body }) ($x1 - 214) ($la[$i] - 11) 200 $(if ($on) { -1 } else { 0 }) $Font 3)
                [void](Disc $x1 $ya[$i] $(if ($on) { 7 } else { 5 }) $col)
                $line = Ln @(@($x1, $ya[$i]), @($x2, $yb[$i])) $col $(if ($on) { 5 } else { 2 })
                $end = Grp @((Disc $x2 $yb[$i] $(if ($on) { 8 } else { 5 }) $col), $(if ($on) { Tx $r[2] 32 $Accent ($x2 + 16) ($lb[$i] - 22) 110 $heavyBold $heavy } else { Tx $r[2] 17 $body ($x2 + 16) ($lb[$i] - 11) 110 }))
                $dirn = 4
                if ($on) {
                    $rider = Disc $x1 $ya[$i] 10 $AccentDark; Glow $rider $Accent 0.4 10 0
                    Anim $line $WIPE $AFTER 1.3 0.3 $dirn; Glide $rider ($x2 - $x1) ($yb[$i] - $ya[$i]) $WITH 1.3 0.3; Anim $end $FADE $AFTER 0.3
                } else { Anim $line $WIPE $AFTER 0.45 0.05 $dirn; Anim $end $FADE $WITH 0.3 0.35 }
            }
            if ($hero.Count) { Anim (Grp $hero) $RISE $AFTER 0.6 }
            Foot; $secs = 8
        }
        'Keynote' {
            $xs = 190, 480, 770; $lamp = 372, 480, 588; $cones = @()
            for ($i = 0; $i -lt 3; $i++) {                                                                                       # three stage lights from the rig above
                $c = Poly @(@(($lamp[$i] - 12), -6), @(($lamp[$i] + 12), -6), @(($xs[$i] + 135), 352), @(($xs[$i] - 135), 352)) '#FFFFFF'
                $c.Fill.TwoColorGradient(1, 1); $c.Fill.GradientAngle = 90; $gs = $c.Fill.GradientStops
                $gs.Item(1).Color.RGB = (Rgb (Mix $Accent '#FFFFFF' 0.6)); $gs.Item(1).Transparency = 0.74; $gs.Item(2).Color.RGB = 0xFFFFFF; $gs.Item(2).Transparency = 1; $cones += $c
                [void](Box ($lamp[$i] - 20) 0 40 9 (Mix $Dark '#FFFFFF' 0.3) 3)
            }
            if ($Kicker) { [void](Tx $Kicker 13 $soft 0 34 $W -1 $Font 2 4) }
            $tt = Tx $Title 44 '#FFFFFF' 60 60 ($W - 120) $heavyBold $heavy 2 0 1.18
            for ($i = 0; $i -lt 3; $i++) { Anim $cones[$i] $WIPE $(if ($i) { $WITH } else { $AFTER }) 0.5 (0.2 + 0.22 * $i) 1 }
            Anim $tt $FADE $AFTER 0.6
            for ($i = 0; $i -lt [math]::Min(3, $n); $i++) {
                $r = $rows[$i]; $pool = Tag ($sl.Shapes.AddShape(9, [single]($xs[$i] - 120), 338, 240, 24)); $pool.Line.Visible = 0; $pool.Fill.ForeColor.RGB = (Rgb $Accent); $pool.Fill.Transparency = 0.7; $pool.SoftEdge.Type = 4
                $set = @($pool, (Tx $r[0] 66 $Accent ($xs[$i] - 140) 190 280 $heavyBold $heavy 2)); if ($r.Count -gt 1) { $set += (Tx $r[1] 17 '#FFFFFF' ($xs[$i] - 130) 276 260 -1 $Font 2 0 1.3) }
                Anim (Grp $set) $RISE $AFTER 0.5
            }
            if ($cds.Count) {
                $q = @((Box 110 388 740 104 $dCard 16), (Tx ([string][char]0x201C) 120 $Accent 134 372 80 $heavyBold $heavy), (Tx $cds[0][0] 20 '#FFFFFF' 214 404 610 0 $Font 1 0 1.3))
                if ($cds[0].Count -gt 1) { $q += (Tx $cds[0][1] 15 $soft 214 462 610 -1 $Font 3) }
                Anim (Grp $q) $WIPE $AFTER 0.7 0.2 4
            }
            Foot $true; $secs = 8
        }
        'TaskBoard' {
            [void](Box 0 0 $W 104 $Dark); if ($Kicker) { [void](Tx $Kicker 12.5 $soft 65 22 600 -1 $Font 1 3) }
            [void](Tx $Title 30 '#FFFFFF' 65 46 640 $heavyBold $heavy); if ($Sub) { [void](Tx $Sub 15 $dText 655 62 240 0 $Font 3) }
            [void](Box 65 100 120 4 $Accent)
            $heads = @($cds | ForEach-Object { $_[0] }); $hx = 112, 448, 736, 812
            for ($c = 0; $c -lt [math]::Min(4, $heads.Count); $c++) { [void](Tx $heads[$c] 15 $gray $hx[$c] 116 200 -1) }
            $pitch = [math]::Min(84, 340 / $n); $rh = $pitch - 8; $y0 = 144; $bx = 448; $bw = 214; $bars = @()
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $y = $y0 + $i * $pitch; $pc = [double]$r[2]; $stt = [int]$r[3]; $mid = $y + $rh / 2
                $set = @((Box 50 $y 860 $rh '#FFFFFF' 8), (Mid (Tx ('{0:00}' -f ($i + 1)) 22 $Accent 66 $y 40 $heavyBold $heavy) $y $rh))
                $set += (Tx $r[0] 18 $Ink 112 ($mid - 25) 326 -1); $set += (Tx $r[1] 15 $body 112 ($mid + 2) 326)
                $set += (Box $bx ($mid - 7) $bw 14 $rail 7); $set += (Mid (Tx "$($r[2])%" 19 $Ink 664 $y 62 $heavyBold $heavy 3) $y $rh)
                for ($q = 1; $q -le 3; $q++) { $set += (Disc (734 + 20 * $q) $mid 7 $(if ($q -ne $stt) { $rail } elseif ($q -eq 1) { $Accent } elseif ($q -eq 2) { $soft } else { $Ink })) }   # three lamps, one lit
                if ($r.Count -gt 4) { $set += (Mid (Tx $r[4] 15 $body 812 $y 90 -1 $Font 1 0 1.1) $y $rh) }
                $bar = Box $bx ($mid - 7) ($bw * $pc / 100) 14 $(if ($stt -eq 1) { $Accent } elseif ($stt -eq 2) { $soft } else { Mix $Ink '#FFFFFF' 0.4 }) 7
                Anim (Grp $set) $RISE $AFTER 0.4; Anim $bar $WIPE $WITH 0.6 0.3 4
            }
            $bot = $y0 + $n * $pitch - 4
            if ($Plan) {
                $pl = @($Plan -split ';'); $px = $bx + $bw * [double]$pl[0] / 100
                $pln = Ln @(@($px, ($y0 - 6)), @($px, ($bot + 2))) $AccentDark 1.75 4
                $lab = "$($pl[1]) $($pl[0])%"; $pw = Wd $lab 15
                $set = @($pln) + (Pill ($px - $pw / 2) ($bot + 2) $pw 28 $AccentDark $lab 15 '#FFFFFF'); Anim (Grp $set) $WIPE $AFTER 0.6 0.2 1
            }
            $x = 112
            for ($q = 1; $q -le 3; $q++) { if ($heads.Count -ge 4 + $q) { [void](Disc ($x + 7) ($bot + 17) 7 $(if ($q -eq 1) { $Accent } elseif ($q -eq 2) { $soft } else { $Ink })); [void](Tx $heads[3 + $q] 15 $body ($x + 20) ($bot + 7) 90); $x += (Wd $heads[3 + $q] 15) + 14 } }
            Foot; $secs = 3 + 1.0 * $n + 2
        }
        'AIStack' {
            Head $true; $cxc = 190; $cyc = 300; $half = 85; $pin = Mix $Accent $Dark 0.35; $chip = @()
            for ($q = 0; $q -lt 6; $q++) {                                                                                       # pins on all four sides
                $o = -62.5 + 25 * $q
                $chip += (Box ($cxc + $o - 2.5) ($cyc - $half - 13) 5 13 $pin); $chip += (Box ($cxc + $o - 2.5) ($cyc + $half) 5 13 $pin)
                $chip += (Box ($cxc - $half - 13) ($cyc + $o - 2.5) 13 5 $pin); $chip += (Box ($cxc + $half) ($cyc + $o - 2.5) 13 5 $pin)
            }
            $o1 = Box ($cxc - $half) ($cyc - $half) (2 * $half) (2 * $half) $dCard 14; $o1.Line.Visible = -1; $o1.Line.ForeColor.RGB = (Rgb $Accent); $o1.Line.Weight = 1.75
            $o2 = Box ($cxc - $half + 18) ($cyc - $half + 18) (2 * $half - 36) (2 * $half - 36) $Dark 6; $o2.Line.Visible = -1; $o2.Line.ForeColor.RGB = (Rgb $pin); $o2.Line.Weight = 0.75
            $chip += $o1; $chip += $o2; $chip += (Disc ($cxc - $half + 11) ($cyc - $half + 11) 3 $Accent)
            $chip += (Mid (Tx $Sub 30 $Accent ($cxc - $half + 18) ($cyc - $half + 18) (2 * $half - 36) $heavyBold $heavy 2 0 1.1) ($cyc - $half + 18) (2 * $half - 36))
            $chipG = Grp $chip; Glow $o1 $Accent 0.55 26 0; Anim $chipG $ZOOM $AFTER 0.6 0.2
            $y = $cyc + $half + 36; $foot = @(); foreach ($c in $cds) { $foot += (Tx $c[0] 15 $dText 50 $y 280 -1 $Font 2); $y += 25 }
            $sx = 430; $swid = 465; $pitch = [math]::Min(86, 344 / $n); $sh = $pitch - 12; $y0 = 136; $bend = 330, 372, 372, 330
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $y = $y0 + $i * $pitch; $ty = $y + $sh / 2; $sy = $cyc + ($i - ($n - 1) / 2) * 25; $bx = $(if ($n -eq 4) { $bend[$i] } else { 350 })
                $trace = Ln @(@(($cxc + $half + 13), $sy), @($bx, $sy), @($bx, $ty), @(($sx - 4), $ty)) $Accent 1.5
                $set = @((Box $sx $y $swid $sh $dCard 10), (Box $sx ($y + 10) 4 ($sh - 20) $Accent), (Disc ($sx - 4) $ty 4.5 $Accent))
                $set += (Tx ('{0:00}' -f ($i + 1)) 15 $Accent ($sx + 22) ($ty - 28) 30 $heavyBold $heavy); $set += (Tx $r[0] 21 '#FFFFFF' ($sx + 54) ($ty - 31) 240 -1)
                $set += (Tx $r[1] 15 $dText ($sx + 54) ($ty + 3) 250)
                if ($r.Count -gt 2) { $set += (Tx $r[2] 32 $Accent ($sx + $swid - 24 - 150) ($ty - 30) 150 $heavyBold $heavy 3) }
                if ($r.Count -gt 3) { $set += (Tx $r[3] 15 $dText ($sx + $swid - 24 - 150) ($ty + 9) 150 0 $Font 3) }
                Anim $trace $WIPE $AFTER 0.45 0.05 4; Anim (Grp $set) $WIPE $AFTER 0.4 0 4
            }
            if ($foot.Count) { Anim (Grp $foot) $FADE $AFTER 0.5 }
            Foot $true; $secs = 4 + 1.0 * $n
        }
        'Ramp' {
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 48 420 -1 $Font 1 3) }
            $hero = Tx $Title 112 $Accent 60 66 20 $heavyBold $heavy; $hero.TextFrame.WordWrap = 0
            if ([double]$hero.Width -gt 430) { $hero.TextFrame.TextRange.Font.Size = [single][math]::Floor(112 * 430 / [double]$hero.Width) }
            $y = [double]$hero.Top + [double]$hero.Height + 2; $msg = Tx $Sub 25 $Ink 65 $y 400 -1 $Font 1 0 1.25
            [void](Box 65 372 380 1.5 (Mix $Light $Ink 0.2)); $x = 65
            foreach ($c in $cds) { [void](Tx $c[0] 34 $AccentDark $x 384 185 $heavyBold $heavy); if ($c.Count -gt 1) { [void](Tx $c[1] 15 $body $x 430 180 0 $Font 1 0 1.3) }; $x += 195 }
            $vals = @($rows | ForEach-Object { [double]$_[1] }); $vmax = ($vals | Measure-Object -Maximum).Maximum; $unit = [math]::Ceiling($vmax / 52)
            $tw = 24; $tH = 13; $gp = 4; $per = 4; $cw = $per * ($tw + $gp) - $gp; $base = 446; $x0 = 512; $step = (904 - $x0 - $cw) / [math]::Max(1, $n - 1); $tops = @()
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $cnt = [int][math]::Round($vals[$i] / $unit); $x = $x0 + $i * $step; $last = ($i -eq $n - 1)
                $col = Mix $Light $Accent (0.3 + 0.7 * $i / [math]::Max(1, $n - 1)); $set = @()
                for ($t = 0; $t -lt $cnt; $t++) { $set += (Box ($x + ($t % $per) * ($tw + $gp)) ($base - $tH - [math]::Floor($t / $per) * ($tH + $gp)) $tw $tH $col 3) }
                $nrow = [math]::Ceiling($cnt / $per); $top = $base - $nrow * ($tH + $gp); $tops += , @(($x + $cw / 2), $top)
                $stack = Grp $set; Anim $stack $WIPE $AFTER (0.35 + 0.07 * $nrow) 0.1 3
                $v = Tx $(if ($r.Count -gt 2) { $r[2] } else { $r[1] }) $(if ($last) { 28 } else { 22 }) $(if ($last) { $AccentDark } else { $Ink }) ($x + $cw / 2 - 85) ($top - $(if ($last) { 40 } else { 33 })) 170 $heavyBold $heavy 2
                Anim $v $FADE $WITH 0.3 (0.3 + 0.07 * $nrow)
                [void](Tx $r[0] 19 $Ink ($x + $cw / 2 - 80) ($base + 8) 160 -1 $Font 2); if ($r.Count -gt 3) { [void](Tx $r[3] 15 $body ($x + $cw / 2 - 80) ($base + 34) 160 0 $Font 2) }
            }
            $curve = @(); $lift = 58                                                                                              # dashed trajectory over the stacks
            for ($q = 0; $q -le 24; $q++) {
                $u = $q / 24 * ($n - 1); $k = [math]::Min($n - 2, [math]::Floor($u)); $f = $u - $k; $e = $f * $f * (3 - 2 * $f)
                $curve += , @(($tops[$k][0] + ($tops[$k + 1][0] - $tops[$k][0]) * $f), ($tops[$k][1] + ($tops[$k + 1][1] - $tops[$k][1]) * $e - $lift))
            }
            $tr = Ln $curve $AccentDark 2 4; $tr.Line.EndArrowheadStyle = 2; Anim $tr $WIPE $AFTER 0.9 0.1 4
            Foot; $secs = 8
        }
    }
    Save-Deck $Out $Preview ([int][math]::Ceiling($secs))
    "saved $Out"
} catch { Close-Deck; throw }
