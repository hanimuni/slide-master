# Presentation-grade scene pages: big type, soft-shadow cards, one accent colour, and entrance animations that
# play by themselves when the slide opens (no click per item). One script, one layout per page:
#   DayFlow    cards in a row that rise in turn, the last one lit            -Items 'chip;emoji;head;line;line|...' (3-5)  -Sub 'footer line'  [-Images last.png]
#   Dashboard  dark score card + bars that grow from the left                -Cards 'label;big;unit;badge;line;line'  -Items 'emoji;name;value;#colour;tag|...'
#   HeroChart  one huge number + an area chart that draws itself (dark page) -Sub 'big;unit;caption'  -Items 'label;value|...'  -Cards 'chip|chip'
#   Ring       headline at the left, a ring that sweeps to the percentage    -Items '83'  -Sub 'caption'  -Cards 'chip|chip'
#   People     ten figures per row lighting up one by one                    -Items 'count;line;line|count;line;line' (count out of ten)
#   Voice      full-bleed photo that drifts closer while a quote comes in    -Images photo.jpg  -Title 'quote|line two'  -Sub 'who said it'   (2 slides, Morph)
#   Roadmap    a line that draws itself, stops popping up along it (dark)    -Items 'date;line;line;#TAG|...'  [-Images one per stop, growing]
# Colours come from four inputs (-Accent, -Light, -Dark and optionally -AccentDark); every tint is mixed from them.
# Example: .\scene_page.ps1 -Layout Ring -Kicker 'WEEKLY ACTIVE' -Title 'Eight in ten|open it weekly' -Items 83 -Sub 'five-week average' -Out C:\work\ring.pptx -Preview C:\work\ring.mp4
param(
    [Parameter(Mandatory)][ValidateSet('DayFlow', 'Dashboard', 'HeroChart', 'Ring', 'People', 'Voice', 'Roadmap')][string]$Layout,
    [string]$Kicker = '',                         # small spaced line above the title
    [string]$Title = '',                          # '|' = line break
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Note = '',                           # small footnote at the lower left (source, "sample data" ...)
    [string]$Accent = '#7C6BFF',
    [string]$AccentDark = '',                     # '' = mixed from -Accent
    [string]$Light = '#FCEFE6',                   # background of the light pages
    [string]$Dark = '#241C45',                    # background of the dark pages and of the dark card
    [string]$Ink = '#1F1B2E',
    [string]$Font = 'Malgun Gothic',
    [string]$HeavyFont = '',                      # for the huge numbers; '' = -Font in bold
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

. (Join-Path $PSScriptRoot 'scene_helpers.ps1')

try {
    $onDark = $Layout -in 'HeroChart', 'Roadmap', 'Voice'
    New-Deck $(if ($onDark) { $Dark } elseif ($Layout -in 'Dashboard', 'People') { '#FFFFFF' } else { $Light })
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.6            # fade in
    $rows = Rows $Items

    switch ($Layout) {
        'DayFlow' {
            [void](Disc 900 20 250 $lilac); Head
            $n = $rows.Count; $gap = 20; $cw = (830 - $gap * ($n - 1)) / $n; $cy0 = 138; $ch = 282
            for ($i = 0; $i -lt $n; $i++) {
                $x = 65 + $i * ($cw + $gap); $st = $rows[$i]; $hot = $i -eq $n - 1; $parts = @()
                $card = Box $x $cy0 $cw $ch $(if ($hot) { $Accent } else { '#FFFFFF' }) 20; Glow $card $(if ($hot) { $Accent } else { Mix $Accent '#FFFFFF' 0.55 }) $(if ($hot) { 0.55 } else { 0.75 }) 18 6; $parts += $card
                if ($hot) { $card.Fill.TwoColorGradient(1, 1); $card.Fill.ForeColor.RGB = (ConvertTo-Bgr (Mix $Accent '#FFFFFF' 0.12)); $card.Fill.BackColor.RGB = (ConvertTo-Bgr $AccentDark); $card.Fill.GradientAngle = 315 }
                $parts += Pill ($x + 18) ($cy0 + 18) (Wd $st[0] 13) 28 $(if ($hot) { '#FFFFFF' } else { $tint }) $st[0] 13 $AccentDark
                if ($hot -and $Images.Count) { $parts += (Disc ($x + 62) ($cy0 + 112) 46 $Light); $parts += (Pic $Images[0] ($x + 20) ($cy0 + 70) 84) }
                else { $parts += (Disc ($x + 58) ($cy0 + 108) 40 $(if ($hot) { '#FFFFFF' } else { $tint })); $parts += (Mid (Tx $st[1] 34 $Ink ($x + 18) ($cy0 + 68) 80 0 'Segoe UI Emoji' 2) ($cy0 + 68) 80) }
                $parts += (Tx $st[2] 21 $(if ($hot) { '#FFFFFF' } else { $Ink }) ($x + 18) ($cy0 + 170) ($cw - 30) -1)
                if ($st.Count -gt 3) { $parts += (Tx (($st[3..($st.Count - 1)]) -join "`r") 15 $(if ($hot) { $tint } else { $body }) ($x + 18) ($cy0 + 206) ($cw - 28) 0 $Font 1 0 1.3) }
                Anim (Grp $parts) $RISE $AFTER 0.45 0.05
                if (-not $hot) { $a = Mid (Tx ([string][char]0x203A) 30 $Accent ($x + $cw - 2) ($cy0 + 110) 24 -1 $Font 2) ($cy0 + 110) 40; Anim $a $FADE $WITH 0.3 0.2 }
            }
            if ($Sub) { Anim (Grp (Pill 65 444 830 50 '#FFFFFF' $Sub 17 $AccentDark)) $FADE $AFTER 0.5 0.1 }
            Foot
        }
        'Dashboard' {
            [void](Disc 960 560 250 $Light); Head
            $c = @($Cards -split ';'); $sc = @(); $c1 = Box 65 132 300 350 $Dark 22; Glow $c1 $Accent 0.6 22 8; $sc += $c1
            $sc += (Tx $c[0] 12.5 $soft 93 162 240 -1 $Font 1 3)
            $big = Tx $c[1] 78 '#FFFFFF' 90 190 260 $heavyBold $heavy; $big.TextFrame.WordWrap = 0; $sc += $big
            if ($c.Count -gt 2 -and $c[2]) { $sc += (Tx $c[2] 30 $soft ([double]$big.Left + [double]$big.Width + 4) 240 60 -1) }
            if ($c.Count -gt 3 -and $c[3]) { $sc += Pill 93 300 (Wd $c[3] 14) 34 $Accent $c[3] 14 '#FFFFFF' }
            $sc += (Box 93 362 244 1.5 $dLine)
            if ($c.Count -gt 4) { $sc += (Tx $c[4] 16 $dText 93 380 250) }
            if ($c.Count -gt 5) { $sc += (Tx $c[5] 14 (Mix $Dark '#FFFFFF' 0.55) 93 416 250) }
            Anim (Grp $sc) $RISE $AFTER 0.5
            $n = $rows.Count; $pitch = [math]::Min(71, 355 / $n); $top = [double]0; foreach ($r in $rows) { $top = [math]::Max($top, [double]$r[2]) }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $y = 132 + $i * $pitch; $lab = @()
                $lab += (Tx $r[0] 20 $Ink 405 $y 30 0 'Segoe UI Emoji'); $nm = Tx $r[1] 19 $Ink 438 ($y + 1) 300 -1; $nm.TextFrame.WordWrap = 0; $lab += $nm
                if ($r.Count -gt 4 -and $r[4]) { $best = [double]$r[2] -eq $top; $lab += Pill ([double]$nm.Left + [double]$nm.Width + 14) ($y + 1) (Wd $r[4] 12) 24 $(if ($best) { $Accent } else { $tint }) $r[4] 12 $(if ($best) { '#FFFFFF' } else { $AccentDark }) }
                $lab += (Box 405 ($y + 36) 420 16 $lilac 8)
                Anim (Grp $lab) $FADE $(if ($i) { $WITH } else { $AFTER }) 0.3 ($i * 0.08)
            }
            for ($i = 0; $i -lt $n; $i++) {
                $r = $rows[$i]; $y = 132 + $i * $pitch
                $bar = Box 405 ($y + 36) ([math]::Max(16, 420 * [double]$r[2] / 100)) 16 $(if ($r.Count -gt 3 -and $r[3]) { $r[3] } else { $Accent }) 8; Anim $bar $WIPE $AFTER 0.45 0 4
                $v = Tx $r[2] 24 $Ink 838 ($y + 24) 56 -1 $Font 3; Anim $v $FADE $WITH 0.3 0.25
            }
            Foot
        }
        'HeroChart' {
            [void](Disc 60 560 240 $dDeco); Head $true
            $h3 = @($Sub -split ';'); $num = Tx $h3[0] 190 '#FFFFFF' 58 128 340 $heavyBold $heavy 1 -6; $num.TextFrame.WordWrap = 0; $hero = @($num)
            if ($h3.Count -gt 1 -and $h3[1]) { $hero += (Tx $h3[1] 54 $soft ([double]$num.Left + [double]$num.Width + 6) 262 120 -1) }
            Anim (Grp $hero) $RISE $AFTER 0.6
            if ($h3.Count -gt 2) { $cap = Tx $h3[2] 22 $dText 68 362 340 -1; Anim $cap $FADE $AFTER 0.4 }
            if ($Cards) { Anim (Grp (Chips $Cards 68 410 $dCard '#FFFFFF')) $FADE $WITH 0.4 0.15 }
            $n = $rows.Count; $x0 = 440; $x1 = 880; $yb = 440; $yt = 150; $max = [double]0; foreach ($r in $rows) { $max = [math]::Max($max, [double]$r[1]) }; $max *= 1.09
            $px = @(); $py = @(); for ($i = 0; $i -lt $n; $i++) { $px += $x0 + ($x1 - $x0) * $i / ($n - 1); $py += $yb - ($yb - $yt) * [double]$rows[$i][1] / $max }
            foreach ($gy in 0.25, 0.5, 0.75, 1) { [void](Box $x0 ($yb - ($yb - $yt) * $gy) ($x1 - $x0) 1 $dCard) }
            $fb = $sl.Shapes.BuildFreeform(0, [single]$x0, [single]$yb); for ($i = 0; $i -lt $n; $i++) { $fb.AddNodes(0, 0, [single]$px[$i], [single]$py[$i]) }; $fb.AddNodes(0, 0, [single]$x1, [single]$yb); $fb.AddNodes(0, 0, [single]$x0, [single]$yb)
            $area = Tag ($fb.ConvertToShape()); $area.Line.Visible = 0; $area.Fill.TwoColorGradient(1, 1); $area.Fill.GradientAngle = 90
            $gs = $area.Fill.GradientStops; $gs.Item(1).Color.RGB = (ConvertTo-Bgr $Accent); $gs.Item(1).Transparency = 0.1; $gs.Item(2).Color.RGB = (ConvertTo-Bgr $Accent); $gs.Item(2).Transparency = 0.92
            $lb = $sl.Shapes.BuildFreeform(0, [single]$px[0], [single]$py[0]); for ($i = 1; $i -lt $n; $i++) { $lb.AddNodes(0, 0, [single]$px[$i], [single]$py[$i]) }
            $line = Tag ($lb.ConvertToShape()); $line.Fill.Visible = 0; $line.Line.ForeColor.RGB = (ConvertTo-Bgr $soft); $line.Line.Weight = 3.5
            Anim $area $WIPE $AFTER 1.1 0 4; Anim $line $WIPE $WITH 1.1 0 4
            for ($i = 0; $i -lt $n; $i++) {
                $last = $i -eq $n - 1; $dot = Disc $px[$i] $py[$i] $(if ($last) { 10 } else { 5.5 }) $(if ($last) { $Accent } else { '#FFFFFF' })
                $wk = Tx $rows[$i][0] 16 $(if ($last) { '#FFFFFF' } else { Mix $Dark '#FFFFFF' 0.5 }) ($px[$i] - 40) ($yb + 12) 80 $(if ($last) { -1 } else { 0 }) $Font 2
                if ($last) {
                    $dot.Line.Visible = -1; $dot.Line.ForeColor.RGB = 0xFFFFFF; $dot.Line.Weight = 3; Glow $dot $Accent 0.3 22 0
                    $lab = "$($rows[$i][1])$(if ($h3.Count -gt 1) { $h3[1] })"; $lw = Wd $lab 18
                    Anim (Grp (@($dot, $wk) + (Pill ($px[$i] - $lw - 6) ($py[$i] - 54) $lw 36 '#FFFFFF' $lab 18 $AccentDark))) $RISE $AFTER 0.4
                } else { $val = Tx $rows[$i][1] 17 $dText ($px[$i] - 30) ($py[$i] - 32) 60 -1 $Font 2; Anim (Grp @($dot, $wk, $val)) $FADE $(if ($i) { $WITH } else { $AFTER }) 0.25 ($i * 0.12) }
            }
            Foot $true
        }
        'Ring' {
            [void](Disc 40 40 210 $lilac)
            $pct = [double]$rows[0][0]
            if ($Kicker) { [void](Tx $Kicker 12.5 $Accent 65 150 400 -1 $Font 1 3) }
            $hd = Tx $Title 44 $Ink 65 178 440 -1 $Font 1 0 1.2; $below = [double]$hd.Top + [double]$hd.Height + 14
            if ($Sub) { $c4 = Tx $Sub 19 $body 67 $below 420; Anim $c4 $FADE $AFTER 0.4; $below += 46 }
            if ($Cards) { Anim (Grp (Chips $Cards 65 $below '#FFFFFF' $AccentDark)) $FADE $WITH 0.4 0.2 }
            $rc = 705; $ry = 280; $rr = 178; $th = 36
            $halo = Disc $rc $ry ($rr + 22) '#FFFFFF'; Glow $halo (Mix $Accent '#FFFFFF' 0.55) 0.7 30 8
            $track = Tag ($sl.Shapes.AddShape(18, [single]($rc - $rr), [single]($ry - $rr), [single](2 * $rr), [single](2 * $rr))); $track.Line.Visible = 0; $track.Fill.ForeColor.RGB = (ConvertTo-Bgr $lilac); $track.Adjustments.Item(1) = [single]($th / (2 * $rr))
            $arc = Tag ($sl.Shapes.AddShape(20, [single]($rc - $rr), [single]($ry - $rr), [single](2 * $rr), [single](2 * $rr))); $arc.Line.Visible = 0     # block arc: starts at 12 o'clock, runs clockwise
            $arc.Adjustments.Item(1) = 270; $arc.Adjustments.Item(2) = [single]((270 + 3.6 * [math]::Min(99.9, $pct)) % 360); $arc.Adjustments.Item(3) = [single]($th / (2 * $rr))
            $arc.Fill.TwoColorGradient(1, 1); $arc.Fill.ForeColor.RGB = (ConvertTo-Bgr (Mix $Accent '#FFFFFF' 0.2)); $arc.Fill.BackColor.RGB = (ConvertTo-Bgr $AccentDark); $arc.Fill.GradientAngle = 45
            Anim $arc $WHEEL $AFTER 1.3
            $big = Tx $rows[0][0] 128 $Accent ($rc - 132) ($ry - 82) 210 $heavyBold $heavy 3 -4
            Anim (Grp @($big, (Tx '%' 44 $AccentDark ($rc + 82) ($ry + 6) 60 -1))) $RISE $WITH 0.7 0.5
            Foot
        }
        'People' {
            [void](Disc 960 0 230 $Light); Head
            $n = $rows.Count; $rh = [math]::Min(162, (360 - 20 * ($n - 1)) / $n)
            for ($q = 0; $q -lt $n; $q++) {
                $y = 136 + $q * ($rh + 20); $it = $rows[$q]; $col = $(if ($q % 2) { $AccentDark } else { $Accent }); $mid = $y + $rh / 2
                $bg = Box 65 $y 830 $rh $Light 22; Anim $bg $FADE $AFTER 0.3
                $hn = Grp @((Tx "$([int]$it[0] * 10)" 76 $col 88 ($mid - 47) 120 $heavyBold $heavy 3 -2), (Tx '%' 30 $col 210 ($mid - 1) 40 -1)); Anim $hn $RISE $WITH 0.5 0.1
                for ($p = 0; $p -lt 10; $p++) {
                    $pxx = 280 + $p * 36; $c = $(if ($p -lt [int]$it[0]) { $col } else { Mix $Accent '#FFFFFF' 0.72 })
                    Anim (Grp @((Disc ($pxx + 13) ($mid - 29) 10 $c), (Box $pxx ($mid - 15) 26 46 $c 11))) $FADE $AFTER 0.12
                }
                if ($it.Count -gt 1) { $say = Tx (($it[1..($it.Count - 1)]) -join "`r") 21 $Ink 655 0 230 -1 $Font 1 0 1.25; $say.Top = [single]($mid - [double]$say.Height / 2); Anim $say $FADE $AFTER 0.4 }
            }
            Foot
        }
        'Voice' {
            if (-not $Images.Count) { throw 'Voice needs -Images <photo>' }
            $ph = Add-Cover $Images[0]; $ph.Name = 'VoicePhoto'
            $scrim = Box 0 0 $W $H $Dark; $scrim.Name = 'VoiceScrim'; $scrim.Fill.TwoColorGradient(1, 1); $scrim.Fill.GradientAngle = 0      # dark at the left, clear at the right
            $ss = $scrim.Fill.GradientStops; $ss.Item(1).Color.RGB = (ConvertTo-Bgr $Dark); $ss.Item(1).Position = 0.18; $ss.Item(1).Transparency = 0.04; $ss.Item(2).Color.RGB = (ConvertTo-Bgr $Dark); $ss.Item(2).Position = 0.78; $ss.Item(2).Transparency = 1
            $sl.SlideShowTransition.AdvanceOnTime = -1; $sl.SlideShowTransition.AdvanceTime = 0.4                                    # the photo alone for a moment, then it moves on by itself
            Add-MorphSlide 2.2
            $sl.SlideShowTransition.AdvanceOnTime = 0
            $p2 = $sl.Shapes('VoicePhoto'); $p2.Width = [single]([double]$p2.Width * 1.1); $p2.Left = [single](-0.07 * $W); $p2.Top = [single](-0.05 * $H)
            [void](Tx ([string][char]0x201C) 190 $Accent 58 46 200 $heavyBold $heavy)
            $q = Tx $Title 50 '#FFFFFF' 66 196 540 -1 $Font 1 0 1.18; $qb = [double]$q.Top + [double]$q.Height
            [void](Box 68 ($qb + 14) 56 4 $Accent 2)
            if ($Sub) { [void](Tx $Sub 18 $dText 68 ($qb + 34) 520) }
            Foot $true
        }
        'Roadmap' {
            [void](Disc 930 10 230 $dDeco); Head $true
            $n = $rows.Count; $ly = 270; $lx0 = 105; $pitch = 750 / ($n - 1); $now = 0
            for ($i = 0; $i -lt $n; $i++) { if (@($rows[$i] | Where-Object { $_ -like '#*' }).Count -and -not $now) { $now = $i } }   # the line is coloured up to the first tagged stop
            $tr = Box $lx0 ($ly - 2) 750 4 $dLine 2; Anim $tr $WIPE $AFTER 0.9 0 4
            if ($now) { $pr = Box $lx0 ($ly - 3) ($pitch * $now) 6 $Accent 3; Anim $pr $WIPE $AFTER 0.4 0 4 }
            $tags = 0
            for ($i = 0; $i -lt $n; $i++) {
                $s = $rows[$i]; $tagText = "$(@($s | Where-Object { $_ -like '#*' })[0])".TrimStart('#'); $lines = @($s[1..($s.Count - 1)] | Where-Object { $_ -notlike '#*' })
                $x = $lx0 + $i * $pitch; $d = [math]::Min($pitch - 30, 62 + 34 * $i / ($n - 1) + $(if ($tagText) { 12 } else { 0 })); $parts = @()
                $ring = Disc $x $ly ($d / 2 + 5) $(if ($tagText) { $Accent } else { $dCard }); if ($tagText) { Glow $ring $Accent 0.35 26 0 }; $parts += $ring
                if ($i -lt $Images.Count) { $parts += (Pic $Images[$i] ($x - $d / 2) ($ly - $d / 2) $d) }
                else { $parts += (Disc $x $ly ($d / 2) $(if ($tagText) { $AccentDark } else { $Dark })); $parts += (Mid (Tx "$($i + 1)" ($d * 0.36) '#FFFFFF' ($x - $d / 2) ($ly - $d / 2) $d -1 $Font 2) ($ly - $d / 2) $d) }
                if ($tagText) { $tags++; $tw = Wd $tagText 13; $parts += Pill ($x - $tw / 2) ($ly - $d / 2 - 44) $tw 28 $(if ($tags -gt 1) { '#FFD93D' } else { '#FFFFFF' }) $tagText 13 $(if ($tags -gt 1) { $Dark } else { $AccentDark }) }
                $parts += (Tx $s[0] 21 $(if ($tagText) { '#FFFFFF' } else { $soft }) ($x - $pitch / 2 + 3) ($ly + 72) ($pitch - 6) -1 $Font 2)
                if ($lines.Count) { $parts += (Tx ($lines -join "`r") 15.5 $dText ($x - $pitch / 2 + 3) ($ly + 104) ($pitch - 6) 0 $Font 2 0 1.3) }
                Anim (Grp $parts) $RISE $AFTER 0.35 0.05
            }
            Foot $true
        }
    }
    Save-Deck $Out $Preview 4
    "saved $Out"
} catch { Close-Deck; throw }
