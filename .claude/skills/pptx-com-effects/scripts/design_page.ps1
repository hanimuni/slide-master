# Finished page designs seen in passing in the tutorial videos - one script, one layout per page:
#   Window    retro "window" cards: title bar with - [] x, thick outline, hard shadow      -Items 'head;text|...'
#   Lines     concept cover: two thick slanted lines pass behind a big title
#   Index     contents list laid over a wide colour band, dotted leaders to page numbers  -Items 'name;page|...'
#   Dots      dot-grid schedule: one row per task, twelve dots, the active months coloured -Items 'task;3-6|...'
#   Radar     radar chart drawn with shapes, value labels at the corners                  -Items 'label;0-100|...'
#   HalfDonut half ring split by share, labels around it, title inside                    -Items 'label;value|...'
#   Leaders   one photo with dots on it and elbow lines out to labels                     -Items 'label;x;y|...' (x,y = 0-1 inside the photo)
#   Process   steps as circles inside one long capsule outline, arrows between            -Items 'head;text|...'
#   Plus      two cards joined by "+", conclusion on a ribbon below (-Sub)                -Items 'head;text|head;text'
#   Stack     photo on a pile of offset sheets, outlined title across its edge
#   Pills     dark photo, five pill-shaped windows where it stays bright
#   Ring      photo with a huge translucent ring running off the slide
#   GlassFade two glass panels over a photo, outline fading out at the ends               -Items 'head;text|head;text'
#   Runner    plastic-model parts frame: parts hang off a spine inside a frame            -Items 'label|label|...'
#   Scribble  neon loops between a portrait and its cut-out (-Cutout), so they circle the head
#   Sparkle   an animated GIF (second image) between a darkened photo and its cut-out
#   Bento     cover in blocks: title card with a cut-out figure rising above it, two colour cells, outlined contents card
#   NeonCard  tall card with a glowing outline and one huge word wider than the card
#   Rows      content page: title band, then rows of a colour card + an outlined note     -Items 'head;text;note|...'
#   Select    cards in a row, one lit in neon and larger (-Pick, or -Steps to light them in turn)   -Items 'text|text|text'
#   Focus     list where one line is bright and large, the rest dim (-Pick, or -Steps)    -Items 'key;text|...'
#   FadePhoto photos whose lower part dissolves into the page, caption below              -Items 'caption|caption'
#   Magazine  blurred dark photo, big title, a paragraph (-Sub) running behind the cut-out figure
#   Slash     white page, diagonal colour field on the right, emoji stickers round the title
#   Poster    one word over a dark photo, last letter in the accent colour, rules above and below, cut-out in front
#   Emblem    section divider: dark photo, huge title over a large translucent gear, short text
#   SideBands section divider: dark photo, title, translucent vertical bands at both sides
#   Pastel    three tall pastel cards with a thin outline and a hard grey shadow          -Items 'light;bold;foot|...'
#   Diamonds  numbered diamonds in a row with a label and a line below                    -Items 'label;text|...'
#   GlassBar  photo, glass text box (-Sub), row of gradient buttons along the bottom      -Items 'label|label|...'
#   Split     two halves answering each other (photo | colour)                            -Items 'small;BIG;big|small;BIG;big'
#   Pierce    three round photos in a row with one cut-out object across them (-Cutout)   -Items 'head;text|...'
#   Drawer    contents panel that slides in from the right edge (2 slides, Morph; use an .mp4 preview)   -Items 'line|line|...'
#   TiltCover tilted title in several colours ('word|word|word') with 3-D icons laid over and under the letters;
#             with -Images the icons drop in (2 slides, Morph; .mp4 preview). Use icons from ONE set (same angle, same light)
#   SlantPage content page under a slanted white band: small icon group top left, tilted title, arrow bars  -Items 'label;value|...'
#   Swot      four panels that do not meet in the middle, slightly rotated, a cut-out portrait on the crossing
#             -Items 'HEAD;text|HEAD;text|HEAD;text|HEAD;text' (top-left, top-right, bottom-left, bottom-right)  -Sub 'MARK|role'
#   Profile   person page: slanted colour panel at the left, the cut-out twice (a large faint copy behind the sharp one),
#             big name at the right  -Sub 'English name|affiliation'  -Items 'line|line|line' (one-sentence introduction)
#   Launch    product reveal in three cuts (Morph; .mp4 preview): big word, word + lines, feature cards
#             -Images product.png  -Items 'word;line;line|card|card|card|card'
# Example: .\design_page.ps1 -Layout Radar -Title 'Skills' -Items 'Plan;80|Design;95|Code;60|Talk;70|Lead;85' -Out C:\work\radar.pptx -Preview C:\work\radar.png
param(
    [Parameter(Mandatory)][ValidateSet('Window', 'Lines', 'Index', 'Dots', 'Radar', 'HalfDonut', 'Leaders', 'Process', 'Plus', 'Stack', 'Pills', 'Ring', 'GlassFade', 'Runner', 'Scribble', 'Sparkle', 'Bento', 'NeonCard', 'Rows', 'Select', 'Focus', 'FadePhoto', 'Magazine', 'Slash', 'Poster', 'Emblem', 'SideBands', 'Pastel', 'Diamonds', 'GlassBar', 'Split', 'Pierce', 'Drawer', 'Launch', 'TiltCover', 'SlantPage', 'Swot', 'Profile')][string]$Layout,
    [string]$Title = '',                          # '|' = line break
    [string]$Sub = '',
    [string]$Items = '',
    [string[]]$Images = @(),
    [string]$Cutout = '',
    [string]$Accent = '#2F6BFF',
    [string]$Color = '#111111',
    [string]$Background = '#F4F3EF',
    [string]$Font = 'Arial Black',
    [string[]]$Palette = @('#E87CB4', '#00D840', '#5B6B7C', '#D09650'),   # TiltCover: word colours; Swot: the four panels
    [int]$Seed = 7,
    [int]$Pick = 0,                               # Select / Focus: which item is lit (1-based; 0 = the last one)
    [switch]$Steps,                               # Select / Focus: light the items one after another (one slide each, Morph; use an .mp4 preview)
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in @($Images) + $Cutout) { if ($p -and -not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

. (Join-Path $PSScriptRoot 'page_helpers.ps1')

# -Steps: a fresh slide with a Morph transition; shapes keep the same names on every slide, so Morph pairs them
function Next-Step {
    $script:sl = $pres.Slides.Add($pres.Slides.Count + 1, 12); $sl.FollowMasterBackground = 0
    $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = (ConvertTo-Bgr $Background)
    $sl.SlideShowTransition.EntryEffect = 3954; $sl.SlideShowTransition.Duration = 0.5
}
function Build-Select([int]$on) {                   # cards in a row; card $on is lit, larger, and drawn last so it overlaps its neighbour
    $cw = 210; $gap = 12; $x0 = ($W - $cw * $n - $gap * ($n - 1)) / 2; $cy = $H * 0.56; $dim = ConvertTo-Bgr '#4A4A4A'
    if ($Title) { $t = Put (New-FitText $Title 34 $ink $Font ($W - 2 * $m)) $m 50; $t.Name = 'Title' }
    foreach ($i in @(0..($n - 1) | Where-Object { $_ -ne $on }) + $on) {
        $lit = ($i -eq $on); $size = if ($lit) { $cw * 1.12 } else { $cw }
        $x = $x0 + ($cw + $gap) * $i + ($cw - $size) / 2 - $(if ($lit -and $i -gt 0) { 16 } else { 0 })
        if ($lit) { $card = Round $x ($cy - $size / 2) $size $size $acc 0.08; $card.Glow.Radius = 20; $card.Glow.Color.RGB = $acc; $card.Glow.Transparency = 0.4 }
        else { $card = Round $x ($cy - $size / 2) $size $size $bg 0.08 $dim 1.5 }
        $card.Name = "Card$($i + 1)"; $tone = if ($lit) { 0 } else { $dim }
        $num = Put (New-Text "$($i + 1)" $(if ($lit) { 18 } else { 16 }) $tone $Font) ($x + 18) ($cy - $size / 2 + 14); $num.Name = "Num$($i + 1)"
        $lab = Mid (New-FitText $list[$i][0] $(if ($lit) { 24 } else { 20 }) $tone $Font ($size - 30)) ($x + $size / 2) $cy; $lab.Name = "Label$($i + 1)"
    }
}
function Build-Focus([int]$on) {                    # list; line $on is bright and large, the others dim
    $dim = ConvertTo-Bgr '#4A4A4A'
    if ($Title) { $tag = Round ($W * 0.33) 46 ($W * 0.34) 40 $acc 0.5; $tag.Name = 'Tag'; $tt = Mid (New-FitText $Title 17 0 $Font ($W * 0.3)) ($W / 2) 66; $tt.Name = 'Title' }
    $top = 120; $rh = ($H - $top - 40) / [math]::Max(1, $n); $vx = $W * 0.42
    for ($i = 0; $i -lt $n; $i++) {
        $y = $top + $rh * ($i + 0.5); $isOn = ($i -eq $on)
        $k = New-FitText $list[$i][0] $(if ($isOn) { 44 } else { 32 }) $(if ($isOn) { 0xFFFFFF } else { $dim }) $Font ($vx - $m - 30); [void](Put $k ($vx - 24 - [double]$k.Width) ($y - [double]$k.Height / 2)); $k.Name = "Key$($i + 1)"
        $rule = Line $vx ($y - 20) $vx ($y + 20) $(if ($isOn) { 0xFFFFFF } else { $dim }) 1.25; $rule.Name = "Rule$($i + 1)"
        if ($list[$i].Count -gt 1) { $d = New-FitText $list[$i][1] 15 $(if ($isOn) { $acc } else { $dim }) $Font ($W - $vx - $m - 30); [void](Put $d ($vx + 24) ($y - [double]$d.Height / 2)); $d.Name = "Text$($i + 1)" }
    }
}

try {
    New-Deck $Background
    $ink = ConvertTo-Bgr $Color; $acc = ConvertTo-Bgr $Accent; $bg = ConvertTo-Bgr $Background; $rng = New-Object System.Random $Seed
    $list = @($Items -split '\|' | Where-Object { $_ } | ForEach-Object { , @($_ -split ';') }); $n = $list.Count; $m = 60
    $head = $true                                   # most layouts: title at the top left
    switch ($Layout) {
        'Window' {
            $gap = 30; $cw = ($W - 2 * $m - $gap * ($n - 1)) / $n; $top = 170; $ch = $H - $top - 60
            for ($i = 0; $i -lt $n; $i++) {
                $x = $m + ($cw + $gap) * $i
                $body = Box 1 $x $top $cw $ch 0xFFFFFF 0 2.5; Hard $body 8
                [void](Box 1 $x $top $cw 34 $acc 0 2.5)
                [void](Put (New-FitText $list[$i][0] 14 0xFFFFFF $Font ($cw - 90)) ($x + 12) ($top + 8))
                $sym = New-Text '-  []  x' 13 0xFFFFFF 'Consolas'; $sym.TextFrame.TextRange.Font.Bold = -1; [void](Put $sym ($x + $cw - [double]$sym.Width - 12) ($top + 8))
                if ($list[$i].Count -gt 1) { [void](Put (New-FitText (($list[$i][1..($list[$i].Count - 1)]) -join '|') 15 0 'Arial' ($cw - 36)) ($x + 18) ($top + 58)) }
            }
        }
        'Lines' {
            foreach ($q in @(@(-0.08, 0.3), @(0.2, 0.86))) {
                $bar = Box 1 ($W * $q[0]) ($H * $q[1]) ($W * 0.9) 26 $acc; $bar.Rotation = -16
            }
            $t = Mid (New-FitText $Title 76 $ink $Font ($W - 240)) ($W / 2) ($H / 2 - 10); $t.Name = 'Title'
            if ($Sub) { [void](Mid (New-FitText $Sub 16 $ink 'Arial' ($W - 300)) ($W / 2) ([double]$t.Top + [double]$t.Height + 26)) }
            $head = $false
        }
        'Index' {
            $bh = $H * 0.46; $by = ($H - $bh) / 2; [void](Box 1 0 $by $W $bh $acc)
            $t = Put (New-FitText $(if ($Title) { $Title } else { 'INDEX' }) 72 $ink $Font ($W * 0.36)) $m ($by + ($bh - 90) / 2); $t.Name = 'Title'
            $x1 = $W * 0.46; $x2 = $W - $m; $step = ($bh - 40) / [math]::Max(1, $n)
            for ($i = 0; $i -lt $n; $i++) {
                $y = $by + 20 + $step * ($i + 0.5)
                $a = Put (New-FitText $list[$i][0] 18 $ink $Font ($W * 0.3)) $x1 ($y - 13)
                $pg = New-Text $(if ($list[$i].Count -gt 1) { $list[$i][1] } else { '{0:00}' -f ($i + 1) }) 18 $ink $Font; [void](Put $pg ($x2 - [double]$pg.Width) ($y - 13))
                $dl = Line ($x1 + [double]$a.Width + 14) $y ($x2 - [double]$pg.Width - 14) $y $ink 1.25; $dl.Line.DashStyle = 3
            }
            $head = $false
        }
        'Dots' {
            $x0 = $W * 0.3; $top = 170; $cw = ($W - $x0 - $m) / 12; $rh = [math]::Min(52, ($H - $top - 50) / [math]::Max(1, $n))
            for ($c = 0; $c -lt 12; $c++) { [void](Mid (New-Text ("$($c + 1)") 12 $ink 'Arial') ($x0 + $cw * ($c + 0.5)) ($top - 26)) }
            for ($i = 0; $i -lt $n; $i++) {
                $y = $top + $rh * ($i + 0.5); $rngM = @($list[$i][1] -split '-'); $from = [int]$rngM[0]; $to = [int]$rngM[-1]
                $lab = New-FitText $list[$i][0] 15 $ink $Font ($x0 - $m - 20); [void](Put $lab $m ($y - [double]$lab.Height / 2))
                for ($c = 1; $c -le 12; $c++) {
                    $on = ($c -ge $from -and $c -le $to); $d = if ($on) { 18 } else { 7 }
                    $dot = Box 9 ($x0 + $cw * ($c - 0.5) - $d / 2) ($y - $d / 2) $d $d $(if ($on) { $acc } else { $ink })
                    if (-not $on) { $dot.Fill.Transparency = 0.65 }
                }
            }
        }
        'Radar' {
            $cx = $W * 0.62; $cy = $H * 0.54; $R = $H * 0.34
            $pt = { param($k, $r) $a = -[math]::PI / 2 + 2 * [math]::PI * $k / $n; , @(($cx + $r * [math]::Cos($a)), ($cy + $r * [math]::Sin($a))) }
            foreach ($lv in 0.25, 0.5, 0.75, 1) { $g = Poly @(0..($n - 1) | ForEach-Object { & $pt $_ ($R * $lv) }) $null $ink 0.75; $g.Line.Transparency = 0.7 }
            for ($k = 0; $k -lt $n; $k++) { $e = & $pt $k $R; $ax = Line $cx $cy $e[0] $e[1] $ink 0.75; $ax.Line.Transparency = 0.7 }
            $data = Poly @(0..($n - 1) | ForEach-Object { & $pt $_ ($R * [math]::Min(100, [double]$list[$_][1]) / 100) }) $acc $acc 2; $data.Fill.Transparency = 0.6; $data.Name = 'Data'
            for ($k = 0; $k -lt $n; $k++) {
                $e = & $pt $k ($R + 34); $v = & $pt $k ($R * [double]$list[$k][1] / 100)
                [void](Box 9 ($v[0] - 5) ($v[1] - 5) 10 10 $acc)
                [void](Mid (New-Text "$($list[$k][0])`r$($list[$k][1])" 13 $ink $Font) $e[0] $e[1])
            }
        }
        'HalfDonut' {
            $cx = $W / 2; $cy = $H * 0.86; $R = $H * 0.62; $sum = 0; foreach ($it in $list) { $sum += [double]$it[1] }
            $shade = 0; $a0 = 180
            for ($i = 0; $i -lt $n; $i++) {
                $a1 = $a0 + 180 * [double]$list[$i][1] / $sum
                $pie = Box 142 ($cx - $R) ($cy - $R) (2 * $R) (2 * $R) $(if ($i % 2) { $ink } else { 0xFFFFFF }) $bg 3     # 142 = pie; angles clockwise from 3 o'clock
                $pie.Adjustments.Item(1) = [single]$a0; $pie.Adjustments.Item(2) = [single]$a1; if ($i % 2) { $pie.Fill.Transparency = 0.25 }
                $mid = ($a0 + $a1) / 2 * [math]::PI / 180; $lr = $R * 0.78
                [void](Mid (New-Text "$($list[$i][0])`r$($list[$i][1])" 15 $(if ($i % 2) { 0xFFFFFF } else { $ink }) $Font) ($cx + $lr * [math]::Cos($mid)) ($cy + $lr * [math]::Sin($mid)))
                $a0 = $a1
            }
            $r2 = $R * 0.56; [void](Box 9 ($cx - $r2) ($cy - $r2) (2 * $r2) (2 * $r2) $bg)
            [void](Box 1 0 $cy $W ($H - $cy) $bg)
            if ($Title) { $t = Mid (New-FitText $Title 34 0xFFFFFF $Font ($r2 * 1.5)) $cx ($cy - $r2 * 0.42); $t.Name = 'Title' }
            if ($Sub) { [void](Mid (New-FitText $Sub 14 0xFFFFFF 'Arial' ($W - 200)) $cx ($cy + ($H - $cy) / 2)) }
            $head = $false
        }
        'Leaders' {
            $pw = $W * 0.44; $ph = $H * 0.66; $px = ($W - $pw) / 2; $py = $H - $ph - 50
            [void](Photo $Images[0] $px $py $pw $ph)
            for ($i = 0; $i -lt $n; $i++) {
                $dx = $px + $pw * [double]$list[$i][1]; $dy = $py + $ph * [double]$list[$i][2]; $left = ([double]$list[$i][1] -lt 0.5)
                $lab = New-FitText $list[$i][0] 15 $ink $Font ($px - $m - 40); $ly = $py + 30 + ($ph - 60) * [math]::Floor($i / 2) / [math]::Max(1, [math]::Ceiling($n / 2) - 1)
                $lx = if ($left) { $px - 40 - [double]$lab.Width } else { $px + $pw + 40 }; [void](Put $lab $lx ($ly - [double]$lab.Height / 2))
                $ex = if ($left) { $px - 30 } else { $px + $pw + 30 }
                $con = $sl.Shapes.AddConnector(2, [single]$dx, [single]$dy, [single]$ex, [single]$ly); $con.Line.ForeColor.RGB = $ink; $con.Line.Weight = 1.25   # 2 = elbow
                $dot = Box 9 ($dx - 7) ($dy - 7) 14 14 $acc 0xFFFFFF 2.5
            }
        }
        'Process' {
            $ch = 150; $cy = $H * 0.5; $cap = Box 5 $m ($cy - $ch / 2) ($W - 2 * $m) $ch $null $acc 2.5; $cap.Adjustments.Item(1) = 0.5
            $d = 96; $span = ($W - 2 * $m - 60 - $d) / [math]::Max(1, $n - 1)
            for ($i = 0; $i -lt $n; $i++) {
                $x = $m + 30 + $d / 2 + $span * $i
                [void](Box 9 ($x - $d / 2) ($cy - $d / 2) $d $d $acc)
                [void](Mid (New-FitText $list[$i][0] 15 0xFFFFFF $Font ($d - 14)) $x $cy)
                if ($i -lt $n - 1) { [void](Box 33 ($x + $span / 2 - 16) ($cy - 10) 32 20 $ink) }                       # 33 = right arrow
                if ($list[$i].Count -gt 1) { [void](Mid (New-FitText $list[$i][1] 13 $ink 'Arial' ($span - 20)) $x ($cy + $ch / 2 + 34)) }
            }
        }
        'Plus' {
            $cw = $W * 0.32; $ch = 190; $top = 160; $gap = 90; $x0 = ($W - 2 * $cw - $gap) / 2
            for ($i = 0; $i -lt 2; $i++) {
                $x = $x0 + ($cw + $gap) * $i; $card = Box 5 $x $top $cw $ch 0xFFFFFF $ink 1.5; $card.Adjustments.Item(1) = 0.08
                [void](Mid (New-FitText $list[$i][0] 22 $ink $Font ($cw - 40)) ($x + $cw / 2) ($top + 58))
                if ($list[$i].Count -gt 1) { [void](Mid (New-FitText $list[$i][1] 14 $ink 'Arial' ($cw - 40)) ($x + $cw / 2) ($top + 122)) }
            }
            [void](Mid (New-Text '+' 54 $acc $Font) ($W / 2) ($top + $ch / 2))
            $rb = Box 1 $x0 ($top + $ch + 46) (2 * $cw + $gap) 64 $acc
            if ($Sub) { [void](Mid (New-FitText $Sub 22 0xFFFFFF $Font (2 * $cw + $gap - 40)) ($W / 2) ($top + $ch + 78)) }
        }
        'Stack' {
            $pw = $W * 0.4; $ph = $H * 0.68; $px = $W * 0.5; $py = ($H - $ph) / 2
            for ($k = 4; $k -ge 1; $k--) { $sheet = Box 1 ($px + 12 * $k) ($py - 9 * $k) $pw $ph 0xFFFFFF 0x999999 0.75; $sheet.Rotation = [single](1.5 * $k) }
            $shot = Photo $Images[0] $px $py $pw $ph; $shot.Line.Visible = -1; $shot.Line.ForeColor.RGB = 0xFFFFFF; $shot.Line.Weight = 6
            $t = New-FitText $Title 84 $ink $Font ($W * 0.62); $f = $t.TextFrame2.TextRange.Font; $f.Fill.Visible = 0; $f.Line.Visible = -1; $f.Line.ForeColor.RGB = $ink; $f.Line.Weight = 1.75
            [void](Put $t $m ($H / 2 - [double]$t.Height / 2)); $t.Name = 'Title'
            if ($Sub) { [void](Put (New-FitText $Sub 15 $ink 'Arial' ($W * 0.36)) $m ([double]$t.Top + [double]$t.Height + 14)) }
            $head = $false
        }
        'Pills' {
            [void](Add-Cover $Images[0])
            $veil = Box 1 0 0 $W $H 0; $veil.Name = 'Veil'
            $pw = 78; $gap = 22; $x0 = $W * 0.42; $hs = 300, 380, 440, 380, 300
            for ($i = 0; $i -lt 5; $i++) {
                $hole = $sl.Shapes.AddShape(5, [single]($x0 + ($pw + $gap) * $i), [single](($H - $hs[$i]) / 2), [single]$pw, [single]$hs[$i]); $hole.Adjustments.Item(1) = 0.5; $hole.Name = "Hole$i"
                $veil = Merge-Shapes 4 @('Veil', "Hole$i"); $veil.Name = 'Veil'
            }
            $veil.Fill.ForeColor.RGB = 0; $veil.Fill.Transparency = 0.35; $veil.Line.Visible = 0
            $t = Put (New-FitText $Title 44 0xFFFFFF $Font ($W * 0.34)) $m ($H * 0.36); $t.Name = 'Title'
            if ($Sub) { [void](Put (New-FitText $Sub 15 0xFFFFFF 'Arial' ($W * 0.32)) $m ([double]$t.Top + [double]$t.Height + 14)) }
            $head = $false
        }
        'Ring' {
            [void](Add-Cover $Images[0])
            $d = $H * 1.5; $ring = Box 18 ($W * 0.42) (($H - $d) / 2) $d $d 0x808080; $ring.Adjustments.Item(1) = 0.16; $ring.Fill.Transparency = 0.4   # 18 = donut
            $t = Put (New-FitText $Title 52 0xFFFFFF $Font ($W * 0.4)) $m ($H * 0.34); $t.Name = 'Title'
            $s = $t.TextFrame2.TextRange.Font.Shadow; $s.Visible = -1; $s.Blur = 12; $s.Transparency = 0.4; $s.OffsetX = 0; $s.OffsetY = 2
            if ($Sub) { [void](Put (New-FitText $Sub 16 0xFFFFFF 'Arial' ($W * 0.36)) $m ([double]$t.Top + [double]$t.Height + 14)) }
            $head = $false
        }
        'GlassFade' {
            [void](Add-Cover $Images[0])
            $veil = Box 1 0 0 $W $H (ConvertTo-Bgr '#0B1B3A'); $veil.Fill.Transparency = 0.35
            $cw = $W * 0.36; $ch = $H * 0.5; $gap = 40; $x0 = ($W - 2 * $cw - $gap) / 2; $top = $H * 0.34
            for ($i = 0; $i -lt [math]::Min(2, $n); $i++) {
                $x = $x0 + ($cw + $gap) * $i
                $glass = Box 5 $x $top $cw $ch 0xFFFFFF; $glass.Adjustments.Item(1) = 0.06; $glass.Fill.Transparency = 0.86
                $o = $sl.Shapes.AddShape(5, [single]$x, [single]$top, [single]$cw, [single]$ch); $o.Adjustments.Item(1) = 0.06; $o.Name = "O$i"
                $q = $sl.Shapes.AddShape(5, [single]($x + 1.5), [single]($top + 1.5), [single]($cw - 3), [single]($ch - 3)); $q.Adjustments.Item(1) = 0.06; $q.Name = "I$i"
                $rim = Merge-Shapes 4 @("O$i", "I$i"); $rim.Line.Visible = 0; $rim.Fill.TwoColorGradient(1, 1); $rim.Fill.GradientAngle = 45   # the outline is a shape, so it can carry a gradient
                $st = $rim.Fill.GradientStops; $st.Insert(0xFFFFFF, 0.5, 0, 2); foreach ($k in 1, 3) { $st.Item($k).Color.RGB = 0xFFFFFF; $st.Item($k).Transparency = 0.9 }
                [void](Put (New-FitText $list[$i][0] 24 0xFFFFFF $Font ($cw - 60)) ($x + 30) ($top + 34))
                if ($list[$i].Count -gt 1) { [void](Put (New-FitText (($list[$i][1..($list[$i].Count - 1)]) -join '|') 14 0xFFFFFF 'Arial' ($cw - 60)) ($x + 30) ($top + 86)) }
            }
            if ($Title) { $t = Put (New-FitText $Title 38 0xFFFFFF $Font ($W - 2 * $m)) $x0 ($H * 0.13); $t.Name = 'Title' }
            $head = $false
        }
        'Runner' {
            $fx = $m; $fy = 150; $fw = $W - 2 * $m; $fh = $H - $fy - 50; $grey = ConvertTo-Bgr '#8A8F98'
            $frame = Box 5 $fx $fy $fw $fh $null $grey 7; $frame.Adjustments.Item(1) = 0.06
            $sy = $fy + $fh / 2; [void](Line $fx $sy ($fx + $fw) $sy $grey 7)
            $cols = [math]::Ceiling($n / 2); $cw = $fw / $cols
            for ($i = 0; $i -lt $n; $i++) {
                $up = ($i -lt $cols); $c = if ($up) { $i } else { $i - $cols }; $x = $fx + $cw * ($c + 0.5); $pw = $cw * 0.66; $ph = $fh * 0.26
                $py = if ($up) { $sy - 30 - $ph } else { $sy + 30 }
                [void](Line $x $sy $x $(if ($up) { $py + $ph } else { $py }) $grey 5)
                $part = Box 5 ($x - $pw / 2) $py $pw $ph $acc; $part.Adjustments.Item(1) = 0.2
                [void](Mid (New-FitText $list[$i][0] 15 0xFFFFFF $Font ($pw - 16)) $x ($py + $ph / 2))
                $tag = Box 1 ($x + 8) $(if ($up) { $sy - 22 } else { $sy + 6 }) 30 16 $grey; [void](Mid (New-Text ('{0}{1}' -f $(if ($up) { 'A' } else { 'B' }), ($c + 1)) 9 0xFFFFFF 'Arial') ($x + 23) $(if ($up) { $sy - 14 } else { $sy + 14 }))
            }
        }
        'Scribble' {
            $ph = $H * 0.92; $probe = $sl.Shapes.AddPicture($Images[0], 0, -1, 0, 0, -1, -1); $pw = $ph * [double]$probe.Width / [double]$probe.Height; $probe.Delete()
            $px = ($W - $pw) / 2; $py = ($H - $ph) / 2
            $back = $sl.Shapes.AddPicture($Images[0], 0, -1, [single]$px, [single]$py, [single]$pw, [single]$ph); $back.Name = 'Portrait'
            for ($k = 0; $k -lt 7; $k++) {                                 # loose loops around the upper part of the picture
                $ew = $pw * (0.7 + $rng.NextDouble() * 0.5); $eh = $ph * (0.22 + $rng.NextDouble() * 0.2)
                $loop = Box 9 ($px + $pw / 2 - $ew / 2 + $rng.Next(-20, 20)) ($py + $ph * 0.28 - $eh / 2 + $rng.Next(-30, 30)) $ew $eh $null $acc 5; $loop.Rotation = [single]$rng.Next(-35, 35)
            }
            if ($Cutout) { $front = $sl.Shapes.AddPicture($Cutout, 0, -1, [single]$px, [single]$py, [single]$pw, [single]$ph); $front.Name = 'Cutout' }
            $head = $false
        }
        'Sparkle' {
            $cover = Add-Cover $Images[0]; $cover.PictureFormat.Brightness = 0.3
            if ($Images.Count -gt 1) { $g = $sl.Shapes.AddPicture($Images[1], 0, -1, 0, 0, [single]$W, [single]$H); $g.Name = 'Sparkle' }
            if ($Cutout) { $front = Add-Cover $Cutout; $front.Name = 'Cutout' }
            $head = $false
        }
        'Bento' {
            $top = 80; $bw = $W * 0.6; $ch = $H * 0.36; $gap = 8; $cream = ConvertTo-Bgr '#FBEBD5'
            [void](Round $m $top $bw $ch $cream 0.12)
            [void](Put (New-FitText $Title 34 0 $Font ($bw * 0.5)) ($m + 30) ($top + 34))
            if ($Images.Count) { $fig = Pic $Images[0] ($ch + 56); [void](Put $fig ($m + $bw - [double]$fig.Width - 26) ($top + $ch - [double]$fig.Height)) }   # head rises above the card
            $cy = $top + $ch + $gap; $cw = ($bw - $gap) / 2; $c2 = $H - $cy - 50; $note = @($Sub -split '\|')
            [void](Round $m $cy $cw $c2 (ConvertTo-Bgr '#F79A3E') 0.12); [void](Round ($m + $cw + $gap) $cy $cw $c2 $acc 0.12)
            if ($note[0]) { [void](Put (New-FitText $note[0] 15 0 $Font ($cw - 40)) ($m + 20) ($cy + 20)) }
            if ($note.Count -gt 1) { $nt = New-FitText (Lines $note[1..($note.Count - 1)]) 13 0 'Arial' ($cw - 40); $nt.TextFrame.TextRange.ParagraphFormat.Alignment = 3; [void](Put $nt ($m + $bw - 20 - [double]$nt.Width) ($cy + $c2 - 20 - [double]$nt.Height)) }
            $rx = $m + $bw + $gap; $rw = $W - $rx - $m; [void](Round $rx $top $rw ($H - $top - 50) $null 0.12 0xFFFFFF 1.25)
            $step = ($H - $top - 110) / [math]::Max(1, $n)
            for ($i = 0; $i -lt $n; $i++) {
                $no = Put (New-Text "$($i + 1)." 22 $acc $Font) ($rx + 26) ($top + 30 + $step * $i + ($step - 30) / 2)
                [void](Put (New-FitText $list[$i][0] 22 0xFFFFFF $Font ($rw - 90)) ($rx + 62) ([double]$no.Top))
            }
            $head = $false
        }
        'NeonCard' {
            $cw = $W * 0.45; $card = Round (($W - $cw) / 2) 24 $cw ($H - 48) 0 0.08 $acc 2.5
            $card.Glow.Radius = 16; $card.Glow.Color.RGB = $acc; $card.Glow.Transparency = 0.45
            $t = Mid (New-FitText $Title 120 $acc $Font ($cw * 1.25)) ($W / 2) ($H / 2); $t.Name = 'Title'
            if ($Sub) { [void](Mid (New-FitText $Sub 15 0xFFFFFF 'Arial' ($cw - 60)) ($W / 2) ($H - 80)) }
            $head = $false
        }
        'Rows' {
            $bandW = $W * 0.72; $band = Box 1 (($W - $bandW) / 2) 86 $bandW 38 (ConvertTo-Bgr '#FBEBD5')
            if ($Sub) { [void](Mid (New-FitText $Sub 15 0 'Arial' ($bandW - 40)) ($W / 2) 105) }
            if ($Title) { [void](Put (New-FitText $Title 15 $ink $Font ($W - 2 * $m)) (($W - $bandW) / 2) 50) }
            $top = 150; $rh = ($H - $top - 40) / [math]::Max(1, $n); $x0 = ($W - $bandW) / 2; $lw = $bandW * 0.6
            for ($i = 0; $i -lt $n; $i++) {
                $y = $top + $rh * $i; $bh = $rh * 0.82
                [void](Round $x0 $y $lw $bh $acc 0.14)
                $hd = Put (New-FitText $list[$i][0] 18 0 $Font ($lw - 40)) ($x0 + 22) ($y + 14)
                if ($list[$i].Count -gt 1) { [void](Put (New-FitText $list[$i][1] 12 0 'Arial' ($lw - 40)) ($x0 + 22) ([double]$hd.Top + [double]$hd.Height + 6)) }
                [void](Round ($x0 + $lw + 14) $y ($bandW - $lw - 14) $bh $null 0.14 $ink 1)
                if ($list[$i].Count -gt 2) { [void](Mid (New-FitText $list[$i][2] 13 $ink 'Arial' ($bandW - $lw - 50)) ($x0 + $lw + 14 + ($bandW - $lw - 14) / 2) ($y + $bh / 2)) }
            }
            $head = $false
        }
        'Select' {
            $on = if ($Pick -ge 1) { $Pick - 1 } else { $n - 1 }
            foreach ($k in $(if ($Steps) { 0..($n - 1) } else { $on })) { if ($Steps -and $k -gt 0) { Next-Step }; Build-Select $k }
            $head = $false
        }
        'Focus' {
            $on = if ($Pick -ge 1) { $Pick - 1 } else { $n - 1 }
            foreach ($k in $(if ($Steps) { 0..($n - 1) } else { $on })) { if ($Steps -and $k -gt 0) { Next-Step }; Build-Focus $k }
            $head = $false
        }
        'FadePhoto' {
            $k = [math]::Max(1, $Images.Count); $pw = $W * 0.28; $gap = $W * 0.1; $x0 = ($W - $pw * $k - $gap * ($k - 1)) / 2; $top = 150
            for ($i = 0; $i -lt $k; $i++) {
                $x = $x0 + ($pw + $gap) * $i; [void](Photo $Images[$i] $x $top $pw $pw)
                $fog = Box 1 ($x - 1) ($top + $pw * 0.5) ($pw + 2) ($pw * 0.5 + 1) $bg; $fog.Fill.TwoColorGradient(1, 1)      # page colour: clear at the top, solid at the bottom
                $fog.Fill.GradientStops.Item(1).Color.RGB = $bg; $fog.Fill.GradientStops.Item(2).Color.RGB = $bg
                $a = $fog.Fill.GradientStops.Item(1); $b = $fog.Fill.GradientStops.Item(2); if ([double]$a.Position -gt [double]$b.Position) { $a, $b = $b, $a }
                $a.Transparency = 1; $b.Transparency = 0; $fog.Fill.GradientAngle = 90
                if ($i -lt $n) { [void](Mid (New-FitText $list[$i][0] 14 $ink 'Arial' $pw) ($x + $pw / 2) ($top + $pw + 24)) }
            }
        }
        'Magazine' {
            $cover = Add-Cover $Images[0]; $cover.PictureFormat.Brightness = 0.32; $fx = $cover.Fill.PictureEffects.Insert(2); $fx.EffectParameters(1).Value = 18
            $t = Mid (New-FitText $Title 78 0xFFFFFF $Font ($W * 0.86)) ($W / 2) 110; $t.Name = 'Title'
            if ($Sub) {
                $para = $sl.Shapes.AddTextbox(1, [single]($W * 0.08), 190, [single]($W * 0.84), 260); $r = $para.TextFrame.TextRange; $r.Text = $Sub
                $r.Font.Name = 'Arial'; $r.Font.NameFarEast = 'Malgun Gothic'; $r.Font.Size = 17; $r.Font.Color.RGB = 0xFFFFFF; $r.ParagraphFormat.Alignment = 4; $r.ParagraphFormat.SpaceWithin = 1.35
            }
            if ($Cutout) { $front = Add-Cover $Cutout; $front.Name = 'Cutout' }
            $t.ZOrder(0)                                                 # paragraph behind the figure, title in front of it
            $head = $false
        }
        'Slash' {
            $field = Poly @(@(($W * 0.52), 0), @($W, 0), @($W, $H), @(($W * 0.9), $H)) $acc $null 0; $field.Fill.TwoColorGradient(1, 1)
            $field.Fill.ForeColor.RGB = (ConvertTo-Bgr '#7CC4FF'); $field.Fill.BackColor.RGB = $acc; $field.Fill.GradientAngle = 315
            $wedge = Poly @(@(($W * 0.72), ($H * 0.5)), @($W, ($H * 0.5)), @($W, $H), @(($W * 0.9), $H)) $acc $null 0; $wedge.Fill.Transparency = 0.25
            $t = Put (New-FitText $Title 54 $ink $Font ($W * 0.5)) ($W * 0.16) ($H * 0.36); $t.Name = 'Title'
            if ($Sub) { $parts = @($Sub -split '\|'); [void](Put (New-FitText $parts[0] 16 $ink 'Arial' ($W * 0.4)) ($W * 0.16) ([double]$t.Top + [double]$t.Height + 10))
                if ($parts.Count -gt 1) { [void](Put (New-Text $parts[1] 13 $acc $Font) ($W * 0.16) ([double]$t.Top + [double]$t.Height + 40)) } }
            $faces = 0x1F600, 0x1F60E, 0x1F929, 0x1F914, 0x1F973, 0x1F609, 0x1F62E, 0x1F642
            foreach ($q in @(@(0.1, 0.2), @(0.4, 0.14), @(0.62, 0.3), @(0.08, 0.66), @(0.36, 0.78), @(0.6, 0.62), @(0.78, 0.2), @(0.84, 0.72))) {
                $em = New-Text ([char]::ConvertFromUtf32($faces[$script:shots % 8])) 30 0 'Segoe UI Emoji'; $script:shots++
                [void](Put $em ($W * $q[0]) ($H * $q[1])); $em.Rotation = [single]$rng.Next(-20, 20)
            }
            $script:shots = 0; $head = $false
        }
        'Poster' {
            $cover = Add-Cover $Images[0]; $cover.PictureFormat.Brightness = 0.3
            $t = New-FitText $Title 190 0xFFFFFF $Font ($W * 0.6); $c = $t.TextFrame.TextRange.Characters($Title.Length, 1); $c.Font.Color.RGB = $acc
            [void](Mid $t ($W / 2) ($H / 2)); $t.Name = 'Title'; $tl = [double]$t.Left; $tw = [double]$t.Width; $tt = [double]$t.Top; $th = [double]$t.Height
            [void](Box 1 $tl ($tt - 4) ($tw * 0.46) 8 $acc); [void](Box 1 ($tl + $tw * 0.54) ($tt - 4) ($tw * 0.46) 8 $acc)
            [void](Box 1 $tl ($tt + $th) $tw 8 0xFFFFFF)
            if ($Sub) { [void](Put (New-Text $Sub 11 $acc $Font) $tl ($tt + $th + 14)) }
            if ($Cutout) { $front = Add-Cover $Cutout; $front.Name = 'Cutout' }
            $head = $false
        }
        'Emblem' {
            $cover = Add-Cover $Images[0]; $cover.PictureFormat.Brightness = 0.3
            $d = $H * 0.62; $gear = Box 172 ($W / 2 - $d / 2) ($H * 0.42 - $d / 2) $d $d $acc; $gear.Fill.Transparency = 0.6; $gear.Name = 'Emblem'    # 172 = gear with six teeth
            $small = Box 172 ($W / 2 - 22) 70 44 44 $null 0xFFFFFF 2
            $t = Mid (New-FitText $Title 84 0xFFFFFF $Font ($W * 0.8)) ($W / 2) ($H * 0.42); $t.Name = 'Title'
            if ($Sub) { $b = New-FitText $Sub 13 0xFFFFFF 'Arial' ($W * 0.4); $b.TextFrame2.TextRange.Font.Spacing = 1.5; [void](Mid $b ($W / 2) ($H * 0.84)) }
            $head = $false
        }
        'SideBands' {
            $cover = Add-Cover $Images[0]; $cover.PictureFormat.Brightness = 0.32
            $bwid = $W * 0.042; $bh = $H * 0.56
            for ($i = 0; $i -lt 7; $i++) {
                $x = if ($i -lt 4) { $W * 0.05 + $bwid * $i } else { $W * 0.95 - $bwid * (7 - $i) }
                $band = Box 1 $x (($H - $bh) / 2 + (($i % 2) * 18)) $bwid $bh 0xFFFFFF; $band.Fill.Transparency = [single](0.92 - 0.03 * ($i % 4))
            }
            $t = Mid (New-FitText $Title 84 0xFFFFFF $Font ($W * 0.5)) ($W / 2) ($H * 0.42); $t.Name = 'Title'
            if ($Sub) { $b = New-FitText $Sub 13 0xFFFFFF 'Arial' ($W * 0.36); $b.TextFrame2.TextRange.Font.Spacing = 1.5; [void](Mid $b ($W / 2) ($H * 0.66)) }
            $head = $false
        }
        'Pastel' {
            $tones = '#F6B8A8', '#F7B37A', '#F6DE8D', '#BFE3C0', '#B9D4F1'; $cw = $W * 0.18; $ch = $H * 0.5; $gap = 46; $x0 = ($W - $cw * $n - $gap * ($n - 1)) / 2; $top = 170
            for ($i = 0; $i -lt $n; $i++) {
                $x = $x0 + ($cw + $gap) * $i; $card = Round $x $top $cw $ch (ConvertTo-Bgr $tones[$i % 5]) 0.1 0 1
                $sh = $card.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = (ConvertTo-Bgr '#9A9A9A'); $sh.Transparency = 0; $sh.Blur = 0; $sh.OffsetX = 7; $sh.OffsetY = 7
                $a = Put (New-FitText $list[$i][0] 24 0xFFFFFF 'Arial' ($cw - 36)) ($x + 18) ($top + 22)
                if ($list[$i].Count -gt 1) { [void](Put (New-FitText $list[$i][1] 20 0xFFFFFF $Font ($cw - 36)) ($x + 18) ([double]$a.Top + [double]$a.Height + 4)) }
                if ($list[$i].Count -gt 2) { $f = New-FitText $list[$i][2] 15 0 $Font ($cw - 36); [void](Put $f ($x + 18) ($top + $ch - 20 - [double]$f.Height)) }
            }
        }
        'Diamonds' {
            $d = 120; $span = ($W - 2 * $m - $d) / [math]::Max(1, $n - 1); $cy = $H * 0.52
            [void](Box 1 ($W / 2 - 22) 44 44 6 $acc)
            if ($Title) { $t = Mid (New-FitText $Title 30 $ink $Font ($W - 240)) ($W / 2) 84; $t.Name = 'Title' }
            if ($Sub) { [void](Mid (New-FitText $Sub 14 $ink 'Arial' ($W - 240)) ($W / 2) 122) }
            for ($i = 0; $i -lt $n; $i++) {
                $x = if ($n -eq 1) { $W / 2 } else { $m + $d / 2 + $span * $i }
                $dia = Box 4 ($x - $d / 2) ($cy - $d / 2) $d $d 0xFFFFFF (ConvertTo-Bgr '#BDBDBD') 1                               # 4 = diamond
                [void](Mid (New-Text "$($i + 1)" 30 $acc $Font) $x ($cy - 16)); [void](Mid (New-FitText $list[$i][0] 14 $ink $Font ($d * 0.62)) $x ($cy + 18))
                if ($list[$i].Count -gt 1) { [void](Mid (New-FitText $list[$i][1] 12 (ConvertTo-Bgr '#6B6B6B') 'Arial' ($span - 20)) $x ($cy + $d / 2 + 34)) }
            }
            $head = $false
        }
        'GlassBar' {
            [void](Add-Cover $Images[0]); [void](Veil '#0B1B3A' 0.7)
            $t = Put (New-FitText $Title 38 0xFFFFFF $Font ($W * 0.6)) $m 54; $t.Name = 'Title'
            if ($Sub) { $box = Round $m 140 ($W * 0.4) 130 0xFFFFFF 0.08 0xFFFFFF 1; $box.Fill.Transparency = 0.88; [void](Put (New-FitText $Sub 14 0xFFFFFF 'Arial' ($W * 0.4 - 40)) ($m + 20) 160) }
            $bs = 88; $gap = 16; $y = $H - $bs - 70; $x0 = $W - $m - ($bs * $n + $gap * ($n - 1))
            for ($i = 0; $i -lt $n; $i++) {
                $x = $x0 + ($bs + $gap) * $i; $btn = Round $x $y $bs $bs $acc 0.18; $btn.Fill.TwoColorGradient(1, 1); $btn.Fill.ForeColor.RGB = (ConvertTo-Bgr '#8FC2FF'); $btn.Fill.BackColor.RGB = $acc
                [void](Mid (New-Text ('{0:00}' -f ($i + 1)) 24 0xFFFFFF $Font) ($x + $bs / 2) ($y + $bs / 2)); [void](Mid (New-FitText $list[$i][0] 12 0xFFFFFF 'Arial' ($bs + 10)) ($x + $bs / 2) ($y + $bs + 16))
            }
            $head = $false
        }
        'Split' {
            if ($Images.Count) { $half = Photo $Images[0] 0 0 ($W / 2) $H; $half.PictureFormat.Brightness = 0.3 } else { [void](Box 1 0 0 ($W / 2) $H 0) }
            [void](Box 1 ($W / 2) 0 ($W / 2) $H $acc)
            for ($i = 0; $i -lt [math]::Min(2, $n); $i++) {
                $cx = $W * (0.25 + 0.5 * $i)
                [void](Mid (New-FitText $list[$i][0] 15 0xFFFFFF 'Arial' ($W * 0.4)) $cx ($H * 0.24))
                if ($list[$i].Count -gt 1) { [void](Mid (New-FitText $list[$i][1] 54 0xFFFFFF $Font ($W * 0.4)) $cx ($H * 0.4)) }
                if ($list[$i].Count -gt 2) { [void](Mid (New-FitText $list[$i][2] 34 0xFFFFFF $Font ($W * 0.4)) $cx ($H * 0.66)) }
            }
            [void](Line ($W * 0.42) ($H * 0.4) ($W * 0.58) ($H * 0.4) 0xFFFFFF 1)
            if ($Title) { [void](Mid (New-FitText $Title 15 0xFFFFFF 'Arial' 200) ($W / 2) ($H * 0.54)) }
            if ($Sub) { [void](Put (New-FitText $Sub 12 0xFFFFFF 'Arial' ($W * 0.4)) $m ($H - 50)) }
            $head = $false
        }
        'Pierce' {
            $d = $W * 0.2; $gap = $d * 0.4; $x0 = ($W - 3 * $d - 2 * $gap) / 2; $top = 130
            for ($i = 0; $i -lt 3; $i++) {
                $x = $x0 + ($d + $gap) * $i; [void](Photo $Images[$i % $Images.Count] $x $top $d $d 9)
                if ($i -lt $n) {
                    [void](Mid (New-FitText $list[$i][0] 15 $ink $Font ($d + 40)) ($x + $d / 2) ($top + $d + 34))
                    if ($list[$i].Count -gt 1) { [void](Mid (New-FitText $list[$i][1] 12 (ConvertTo-Bgr '#6B6B6B') 'Arial' ($d + 50)) ($x + $d / 2) ($top + $d + 66)) }
                }
            }
            if ($Cutout) { $obj = $sl.Shapes.AddPicture($Cutout, 0, -1, 0, 0, -1, -1); $obj.LockAspectRatio = -1; $obj.Width = [single]((3 * $d + 2 * $gap) * 0.8)
                [void](Put $obj (($W - [double]$obj.Width) / 2) ($top + $d / 2 - [double]$obj.Height / 2)); $obj.Name = 'Object' }
            if ($Title) { $t = Mid (New-FitText $Title 30 $ink $Font ($W - 240)) ($W / 2) 76; $t.Name = 'Title' }
            $head = $false
        }
        'Drawer' {
            $t = Put (New-FitText $Title 54 $ink $Font ($W * 0.6)) $m ($H * 0.36); $t.Name = 'Title'
            if ($Sub) { [void](Put (New-FitText $Sub 16 $ink 'Arial' ($W * 0.5)) $m ([double]$t.Top + [double]$t.Height + 12)) }
            $pw = $W * 0.24; $px = $W - $pw; $made = @()
            $made += Box 1 $px 0 $pw $H 0x16181D
            $made += Put (New-Text 'Contents' 26 $acc $Font) ($px + 30) 70
            for ($i = 0; $i -lt $n; $i++) { $made += Put (New-FitText $list[$i][0] 17 0xFFFFFF 'Arial' ($pw - 60)) ($px + 30) (140 + 46 * $i) }
            $k = 0; foreach ($sh in $made) { $k++; $sh.Name = "Drawer$k" }
            Add-MorphSlide 0.6                                           # the copy is the "open" slide; push the drawer out on the first one
            foreach ($j in 1..$k) { $q = $pres.Slides(1).Shapes("Drawer$j"); $q.Left = [single]([double]$q.Left + $pw + 10) }
            $head = $false
        }
        'Launch' {
            $word = Mid (New-FitText $Title 150 $acc $Font ($W * 0.5)) ($W / 2) ($H * 0.42); $word.Name = 'Word'
            if ($Sub) { $lab = Mid (New-Text $Sub 18 0xFFFFFF 'Arial') ($W / 2) ($H * 0.2); $lab.Name = 'Label' }
            $prod = Pic $Images[0] ($H * 0.9); [void](Put $prod (($W - [double]$prod.Width) / 2) ($H * 0.5)); $prod.Name = 'Product'
            $ratio = [double]$prod.Width / [double]$prod.Height
            Add-MorphSlide 0.9                                           # cut 2: product small in the middle, a word left, lines right
            $q = $sl.Shapes('Product'); $q.Height = [single]($H * 0.42); [void](Put $q (($W - [double]$q.Width) / 2) ($H * 0.3))
            $w2 = $sl.Shapes('Word'); $w2.TextFrame.TextRange.Font.Size = 40; [void](Mid $w2 ($W / 2) 60)
            if ($n) {
                [void](Put (New-FitText $list[0][0] 44 $acc $Font ($W * 0.26)) $m ($H * 0.42))
                if ($list[0].Count -gt 1) { [void](Put (New-FitText (Lines $list[0][1..($list[0].Count - 1)]) 20 $acc $Font ($W * 0.26)) ($W * 0.7) ($H * 0.42)) }
            }
            Add-MorphSlide 0.9                                           # cut 3: product smaller still, feature cards on both sides
            foreach ($old in @($sl.Shapes | Where-Object { $_.Name -notin 'Product', 'Word', 'Label' })) { $old.Delete() }
            $q = $sl.Shapes('Product'); $q.Height = [single]($H * 0.26); [void](Put $q (($W - [double]$q.Width) / 2) ($H * 0.4))
            $cw = $W * 0.16; $ch = 110
            for ($i = 1; $i -lt [math]::Min(5, $n); $i++) {
                $col = if ($i -le 2) { $i - 1 } else { $i }; $x = $m + ($W - 2 * $m - $cw) * $col / 4.0
                [void](Round $x ($H * 0.4) $cw $ch (ConvertTo-Bgr '#1E1E1E') 0.12 $acc 1); [void](Mid (New-FitText $list[$i][0] 14 0xFFFFFF $Font ($cw - 20)) ($x + $cw / 2) ($H * 0.4 + $ch / 2))
            }
            $head = $false
        }
        'TiltCover' {
            $words = @($Title -split '\|'); $t = New-FitText ($words -join '') 110 $ink $Font ($W * 0.62); $r = $t.TextFrame.TextRange; $pos = 1
            for ($i = 0; $i -lt $words.Count; $i++) { $r.Characters($pos, $words[$i].Length).Font.Color.RGB = (ConvertTo-Bgr $Palette[$i % $Palette.Count]); $pos += $words[$i].Length }
            [void](Mid $t ($W / 2) ($H * 0.54)); $t.Rotation = -15; $t.Name = 'Title'
            if ($Sub) { $s2 = New-FitText $Sub 22 (ConvertTo-Bgr '#8A8A8A') $Font ($W * 0.4); [void](Put $s2 ($W * 0.56) ($H * 0.16)); $s2.Rotation = -15; $s2.Name = 'Sub' }
            $spots = @(@(0.3, 0.3, 0.3), @(0.86, 0.42, 0.4), @(0.15, 0.66, 0.24), @(0.6, 0.82, 0.24), @(0.52, 0.2, 0.2), @(0.34, 0.8, 0.2))   # centre x, centre y, height
            for ($i = 0; $i -lt [math]::Min($Images.Count, $spots.Count); $i++) {
                $ic = Pic $Images[$i] ($H * $spots[$i][2]); [void](Put $ic ($W * $spots[$i][0] - [double]$ic.Width / 2) ($H * $spots[$i][1] - [double]$ic.Height / 2)); $ic.Name = "Icon$($i + 1)"
            }
            if ($Images.Count) {
                Add-MorphSlide 0.9                                       # the copy is the finished cover; on the first slide the icons wait above the slide
                for ($i = 1; $i -le [math]::Min($Images.Count, $spots.Count); $i++) { $q = $pres.Slides(1).Shapes("Icon$i"); $q.Top = [single](-[double]$q.Height - 30 * $i) }
            }
            $head = $false
        }
        'SlantPage' {
            $field = Poly @(@(0, ($H * 0.42)), @($W, ($H * 0.16)), @($W, $H), @(0, $H)) $acc $null 0; $field.Name = 'Field'
            $ix = $m
            foreach ($path in @($Images | Select-Object -First 5)) { $ic = Pic $path 64; [void](Put $ic $ix 34); $ix += [double]$ic.Width - 6 }
            if ($Title) { $t = New-FitText $Title 36 $acc $Font ($W * 0.48); [void](Put $t ($W - $m - [double]$t.Width) 40); $t.Rotation = -13; $t.Name = 'Title' }
            if ($Sub) { $s2 = New-FitText $Sub 20 0xFFFFFF $Font ($W * 0.44); [void](Put $s2 ($W - $m - [double]$s2.Width) 150); $s2.Rotation = -13 }
            $peak = 1; foreach ($it in $list) { $peak = [math]::Max($peak, [double]$it[1]) }
            $base = $H - 70; $bw = 50; $x0 = $m + 30
            $ln = Line $m $base ($m + ($bw + 28) * $n + 40) $base 0xFFFFFF 1; $ln.Line.DashStyle = 3
            for ($i = 0; $i -lt $n; $i++) {
                $bh = 150 * [double]$list[$i][1] / $peak; $x = $x0 + ($bw + 28) * $i
                $bar = Box 35 $x ($base - $bh) $bw $bh 0xFFFFFF; $bar.Fill.Transparency = [single]$(if ([double]$list[$i][1] -eq $peak) { 0 } else { 0.55 })      # 35 = up arrow
                [void](Mid (New-Text $list[$i][0] 11 0xFFFFFF 'Arial') ($x + $bw / 2) ($base + 14))
            }
            $top1 = @($list | Sort-Object { -[double]$_[1] })[0]
            [void](Put (New-Text "$($top1[1])%" 60 0xFFFFFF $Font) ($W * 0.56) ($H * 0.5)); [void](Put (New-Text $top1[0] 18 0xFFFFFF $Font) ($W * 0.56) ($H * 0.5 + 78))
            $head = $false
        }
        'Swot' {
            $vx = $W * 0.56; $hy = $H * 0.57; $far = $W * 0.5; $made = @()
            foreach ($q in @(@(-$far, -$far, $vx, $hy, 0), @($vx, -$far, ($W + $far), $hy, 1), @(-$far, $hy, $vx, ($H + $far), 2), @($vx, $hy, ($W + $far), ($H + $far), 3))) {
                $pn = Box 1 $q[0] $q[1] ($q[2] - $q[0]) ($q[3] - $q[1]) (ConvertTo-Bgr $Palette[$q[4] % $Palette.Count]); $pn.Name = "Panel$($q[4] + 1)"; $made += $pn.Name
            }
            $grid = $sl.Shapes.Range($made).Group(); $grid.Name = 'Panels'; $grid.Rotation = 5       # by hand a table cannot be rotated: cut it, Paste Special > Picture (Enhanced Metafile), then rotate
            $spots = @(@($m, 60, ($vx - $m - 150)), @(($vx + 110), 60, ($W - $vx - 110 - $m)), @($m, ($hy + 60), ($vx - $m - 150)), @(($vx + 110), ($hy + 60), ($W - $vx - 110 - $m)))
            for ($i = 0; $i -lt [math]::Min(4, $n); $i++) {
                $hd = Put (New-FitText $list[$i][0] 22 0xFFFFFF $Font $spots[$i][2]) $spots[$i][0] $spots[$i][1]
                if ($list[$i].Count -gt 1) { [void](Put (New-FitText (Lines $list[$i][1..($list[$i].Count - 1)]) 12 0xFFFFFF 'Arial' $spots[$i][2]) $spots[$i][0] ([double]$hd.Top + [double]$hd.Height + 8)) }
            }
            if ($Images.Count) {                                         # portrait on the crossing: frame, inner block, then the cut-out standing out of the frame
                $fw = 150; $fh = 170; $fx = $vx - $fw / 2 + 10; $fy = $hy - $fh / 2 - 6
                [void](Box 1 ($fx + 10) ($fy + 10) $fw $fh (ConvertTo-Bgr $Palette[3 % $Palette.Count])); [void](Box 1 $fx $fy $fw $fh $null (ConvertTo-Bgr $Palette[0]) 3)
                $fig = Pic $Images[0] ($fh + 40); [void](Put $fig ($fx + ($fw - [double]$fig.Width) / 2 + 5) ($fy + $fh + 10 - [double]$fig.Height)); $fig.Name = 'Portrait'
                $mk = @($Sub -split '\|'); if ($mk[0]) { [void](Mid (New-Text $mk[0] 34 0xFFFFFF $Font) ($fx + $fw / 2 + 5) ($fy + $fh + 44)) }
                if ($mk.Count -gt 1) { [void](Mid (New-Text $mk[1] 12 0xFFFFFF 'Arial') ($fx + $fw / 2 + 5) ($fy + $fh + 76)) }
            }
            $head = $false
        }
        'Profile' {
            $panel = Poly @(@(0, 0), @(($W * 0.327), 0), @(($W * 0.383), $H), @(0, $H)) $acc $null 0; $panel.Name = 'Panel'
            if ($Images.Count) {
                $probe = Pic $Images[0] ($H * 0.84); $ratio = [double]$probe.Width / [double]$probe.Height
                $gh = $H * 1.18; $ghost = Box 1 ($W * 0.04) ($H * 0.01) ($gh * $ratio) $gh 0xFFFFFF     # the faint copy: a picture-filled rectangle, because fill transparency is scriptable
                $ghost.Fill.UserPicture($Images[0]); $ghost.Fill.Transparency = 0.9; $ghost.Name = 'Ghost'
                [void](Put $probe ($W * 0.28 - [double]$probe.Width / 2) ($H - [double]$probe.Height)); $probe.ZOrder(0); $probe.Name = 'Portrait'
            }
            $mk = @($Sub -split '\|')
            $t = Put (New-FitText $Title 110 $acc $Font ($W * 0.44)) ($W * 0.5) ($H * 0.16); $t.Name = 'Title'
            if ($mk[0]) { $en = Put (New-Text $mk[0] 22 $acc 'Arial') ($W * 0.508) ([double]$t.Top + [double]$t.Height + 2); $en.TextFrame2.TextRange.Font.Spacing = -0.5 }
            if ($mk.Count -gt 1) { $org = New-Text $mk[1] 11 (ConvertTo-Bgr '#8A8A8A') 'Arial'; [void](Put $org ($W * 0.955 - [double]$org.Width) ($H * 0.08)) }
            if ($n) { $tag = New-FitText (($list | ForEach-Object { $_[0] }) -join '|') 24 (ConvertTo-Bgr '#1A1A1A') $Font ($W * 0.4); $tag.TextFrame.TextRange.ParagraphFormat.Alignment = 3
                [void](Put $tag ($W * 0.955 - [double]$tag.Width) ($H * 0.9 - [double]$tag.Height)) }
            $head = $false
        }
    }
    if ($head) {
        $y = 50
        if ($Title) { $t = Put (New-FitText $Title 34 $ink $Font ($W - 2 * $m)) $m $y; $t.Name = 'Title'; $y += [double]$t.Height + 6 }
        if ($Sub -and $Layout -ne 'Plus') { [void](Put (New-FitText $Sub 15 $ink 'Arial' ($W - 2 * $m)) $m $y) }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
