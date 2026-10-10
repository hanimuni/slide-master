# One-page data slides built from plain shapes (the "draw it yourself" infographics of the short tutorials):
#   Pictogram     two cards, ten people each, as many filled as the share        -Items 'pct;head;line;line|pct;head;line;line'
#   DiamondGrid   numbered side band, big photo diamond, four small diamonds     -Items 'label|label|label|label' (first = lit)  -Sub 'heading|line|line'
#   Phone         phone mock-up at the right, two poll cards with five units     -Items 'label;pct|label;pct'  -Images screen.png [icon.png]
#   Scale         balance scale, the heavier side down                           -Items 'label;pct;line;line|label;pct;line;line'
#   AreaChart     area graph drawn as one freeform, gradient fill, cards below   -Items 'year;value[;top label]|...'  -Cards 'big;text|big;text|big;text'
#   FadeBars      bars that fade into the dark page, photo panel under a veil    -Items 'label;value;caption|...'  -Cards 'head;line|line|line'
#   ProcessArrow  gradient arrow band with ringed circles, the goal after it     -Items 'head;line;line|...|goal;line' (last = goal)  -Images icon.png ...
#   SlantChart    gradient parallelogram + photo cut to the same slant, chart card, two KPI boxes
#                 -Items 'category;v1;v2;v3|...'  -Cards 'big number|label|label'
#   OrgChart      head in a thick ring, three teams around it                    -Items 'team;lead;member;member|...'  -Sub 'name|role'  -Images head.jpg lead1.jpg ...
#   Donut3D       two tilted 3-D rings, one coloured block arc each              -Items 'label;pct;line;line|label;pct;line;line'
#   PhotoTimeline side band with a vertical word, years on a line, photo columns -Items 'year;head;line;line|...'  -Images a.jpg b.jpg ...
#   ProCon        two tinted panels with a notch for the side's pill, big faint VS       -Items 'head;line;line|...' (left)  -Cards '...' (right)  -Sub 'Pro|Con'
#   SignPost      contents as coloured bands wrapped round a grey post               -Items 'head;text|head;text|head;text'  -Kicker CONTENTS
#   FadeCards     four tall cards whose photo appears from under the card colour     -Items 'head;english;line;line|...'  -Images a.jpg b.jpg ...
#   StepArrows    three upward arrows in steps with a round head, the middle one lit -Items 'head;line;line|...'  -Kicker '2026|DATA REPORT'  -Title 'a|b|c'
#   Compare       mirrored share bars either side of a middle column                 -Items 'pct;caption;keyword|...' (left)  -Cards 'pct;caption|...' (right)  -Sub 'A|B'
#   Folders       four folders; on the next slide a picture rises out of each (2 slides, Morph; use an .mp4 preview)
#                 -Items 'small;BOLD|...'  -Images cutout.png ...
# Positions follow the videos as percentages of the slide. Photos and icons are optional everywhere: without them the
# layout falls back to plain shapes. The chart in SlantChart is drawn with rectangles (no chart object, nothing to link).
# Example: .\info_page.ps1 -Layout Scale -Title 'Work compared' -Items 'Automated;75|By hand;25' -Out C:\work\scale.pptx -Preview C:\work\scale.png
param(
    [Parameter(Mandatory)][ValidateSet('Pictogram', 'DiamondGrid', 'Phone', 'Scale', 'AreaChart', 'FadeBars', 'ProcessArrow', 'SlantChart', 'OrgChart', 'Donut3D', 'PhotoTimeline', 'ProCon', 'SignPost', 'FadeCards', 'StepArrows', 'Compare', 'Folders')][string]$Layout,
    [string]$Title = '',                          # '|' = line break
    [string]$Kicker = '',                         # small line above the title (PhotoTimeline: the word on the band)
    [string]$Sub = '',
    [string]$Items = '',
    [string]$Cards = '',
    [string[]]$Images = @(),
    [string]$Accent = '',                         # default depends on the layout
    [string]$Background = '',
    [string]$Font = 'Arial Black',
    [string[]]$Palette = @(),
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
. (Join-Path $PSScriptRoot 'page_helpers.ps1')
foreach ($p in $Images) { if ($p -and -not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

$defaults = @{                                    # accent, background, palette
    Pictogram = '#0A5BC4', '#F5F5F5', @('#0A5BC4', '#F8766D'); DiamondGrid = '#0F43F5', '#FFFFFF', @('#9E9E9E')
    Phone = '#1F4FD8', '#F2F2F4', @('#D0D0D0'); Scale = '#00844A', '#FFFFFF', @('#00844A', '#707070')
    AreaChart = '#FF5722', '#0B0B0B', @('#FFA04D'); FadeBars = '#0F9D58', '#000000', @('#C41208', '#1F86D8', '#1E9E57')
    ProcessArrow = '#0A66F0', '#FFFFFF', @('#0A3A8C'); SlantChart = '#0A66D0', '#FFFFFF', @('#0A5CC8', '#707070', '#14A0F0')
    OrgChart = '#0A38D6', '#FFFFFF', @('#F2494B', '#F25A10', '#17A04F'); Donut3D = '#0A2E9E', '#FFFFFF', @('#0A44F5', '#E8161B')
    PhotoTimeline = '#1A63F5', '#FFFFFF', @('#1A63F5'); ProCon = '#0082F0', '#FAFAFA', @('#0082F0', '#FA6743')
    SignPost = '#86AB35', '#F3F3F3', @('#86AB35', '#00859D', '#1794E0'); FadeCards = '#001F3C', '#F2F2F2', @('#001F3C')
    StepArrows = '#0062E6', '#EEF0F3', @('#0062E6'); Compare = '#FD7982', '#F0F0F0', @('#FD7982', '#00984D')
    Folders = '#4DD3DC', '#FFFFFF', @('#E42C73', '#00B5C0', '#6E34A0', '#DBC600', '#FE58AE', '#00D2E1', '#B56BD3', '#FCEC50', '#F4639F')      # four backs, four fronts, the bottom of the page
}
if (-not $Accent) { $Accent = $defaults[$Layout][0] }
if (-not $Background) { $Background = $defaults[$Layout][1] }
if (-not $Palette.Count) { $Palette = $defaults[$Layout][2] }

function PX($p) { $W * $p / 100 }
function PY($p) { $H * $p / 100 }
function Shade([string]$hex, $f) {             # lighter ($f > 0: toward white) or darker ($f < 0: toward black) version of a colour
    $v = [Convert]::ToInt32($hex.TrimStart('#'), 16); $hx = ''
    foreach ($sh in 16, 8, 0) { $c = ($v -shr $sh) -band 255; $c = if ($f -ge 0) { $c + (255 - $c) * $f } else { $c * (1 + $f) }; $hx += '{0:X2}' -f [int]$c }
    "#$hx"
}
function Tone([int]$i) { ConvertTo-Bgr $Palette[$i % $Palette.Count] }
# text placed by its vertical centre; $align 1 = $x is the left edge, 2 = the centre, 3 = the right edge
function Say([string]$text, $size, $rgb, $x, $cy, [string]$face = 'Arial', [int]$align = 1) {
    $t = New-Text ($text -replace '\|', "`r") $size ([int]$rgb) $face; $t.TextFrame.TextRange.ParagraphFormat.Alignment = $align
    $left = if ($align -eq 2) { $x - [double]$t.Width / 2 } elseif ($align -eq 3) { $x - [double]$t.Width } else { $x }
    $t.Left = [single]$left; $t.Top = [single]($cy - [double]$t.Height / 2); $t
}
function Grad($s, [string]$from, [string]$to, $angle, $clearFrom = 0, $clearTo = 0) {      # two-stop linear gradient; 90 = top to bottom
    $s.Fill.TwoColorGradient(1, 1); $s.Fill.GradientAngle = [single]$angle; $g = $s.Fill.GradientStops
    $g.Item(1).Color.RGB = (ConvertTo-Bgr $from); $g.Item(1).Transparency = [single]$clearFrom
    $g.Item(2).Color.RGB = (ConvertTo-Bgr $to); $g.Item(2).Transparency = [single]$clearTo
}
function Soft($s, $blur = 20, $drop = 4, $clear = 0.85) { $sh = $s.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Transparency = [single]$clear; $sh.Blur = [single]$blur; $sh.OffsetX = 0; $sh.OffsetY = [single]$drop }
# pictogram person: head + body. $solid = $false draws the outline version (white fill, coloured line)
function Person($cx, $top, $high, $rgb, [bool]$solid = $true, $paper = 0xFFFFFF) {
    $hd = $high * 0.28; $bw = $high * 0.46
    foreach ($q in @(@(9, ($cx - $hd / 2), $top, $hd, $hd), @(5, ($cx - $bw / 2), ($top + $high * 0.34), $bw, ($high * 0.66)))) {
        $part = if ($solid) { Box $q[0] $q[1] $q[2] $q[3] $q[4] $rgb } else { Box $q[0] $q[1] $q[2] $q[3] $q[4] $paper $rgb 1 }
        if ($q[0] -eq 5) { $part.Adjustments.Item(1) = 0.38 }
    }
}
function Heading($x, $cy, $kickRgb, $titleRgb, $size = 28) {       # kicker, title and (grey) sub stacked from $cy down; returns the y below them
    $y = $cy
    if ($Kicker) { $k = Say $Kicker 12 $kickRgb $x $y $Font; $y += [double]$k.Height / 2 + 8 }
    if ($Title) { $t = New-FitText $Title $size ([int]$titleRgb) $Font ($W * 0.6); $t.Left = [single]$x; $t.Top = [single]$y; $t.Name = 'Title'; $y += [double]$t.Height + 8 }
    $y
}

try {
    New-Deck $Background
    $ink = ConvertTo-Bgr '#111111'; $grey = ConvertTo-Bgr '#777777'; $acc = ConvertTo-Bgr $Accent; $bg = ConvertTo-Bgr $Background
    $list = @($Items -split '\|' | Where-Object { $_ } | ForEach-Object { , @($_ -split ';') }); $n = $list.Count
    $notes = @($Cards -split '\|' | Where-Object { $_ } | ForEach-Object { , @($_ -split ';') })
    switch ($Layout) {
        'Pictogram' {
            $band = Box 1 0 (PY 95.4) $W (PY 4.6) $acc; Grad $band $Palette[0] $Palette[1 % $Palette.Count] 0
            [void](Heading (PX 6.6) (PY 10.5) $acc $ink)
            if ($Sub) { [void](Say $Sub 10 $grey (PX 88.6) (PY 17.8) 'Arial' 3) }
            for ($i = 0; $i -lt [math]::Min(2, $n); $i++) {
                $it = $list[$i]; $tone = Tone $i; $top = PY (30.2 + 29.9 * $i); $ch = PY 26.8; $mid = $top + $ch / 2
                $card = Round (PX -2.6) $top (PX 96.9) $ch 0xFFFFFF 0.3; Soft $card
                [void](Say "$($it[0])%" 36 $tone (PX 6.6) $mid $Font); [void](Box 9 (PX 17.2) ($mid - 6) 12 12 $tone)
                if ($it.Count -gt 1) { [void](Say $it[1] 14 $ink (PX 22.5) ($mid - 22) $Font) }
                if ($it.Count -gt 2) { [void](Say (Lines $it[2..($it.Count - 1)]) 10.5 $grey (PX 22.5) ($mid + 18)) }
                $full = [math]::Round([double]$it[0] / 10); $ph = PY 13
                for ($k = 0; $k -lt 10; $k++) { Person (PX (43.1 + 4.7 * $k)) ($mid - $ph / 2) $ph $tone ($k -lt $full) }
            }
        }
        'DiamondGrid' {
            [void](Box 1 0 0 (PX 7.6) $H $acc)
            for ($i = 0; $i -lt 5; $i++) { $no = Say ('{0:00}' -f ($i + 1)) 12 0xFFFFFF (PX 3.8) (PY (15.2 + 7.5 * $i)) $Font 2; if ($i) { $no.TextFrame2.TextRange.Font.Fill.Transparency = 0.5 } }
            [void](Heading (PX 12.6) (PY 15.5) $acc $ink)
            $mk = @($Sub -split '\|')
            if ($mk[0]) { [void](Say $mk[0] 14 $ink (PX 12.6) (PY 68.1) $Font) }
            if ($mk.Count -gt 1) { $body = Say (Lines $mk[1..($mk.Count - 1)]) 11 $grey (PX 12.6) 0; $body.Top = [single](PY 72.5); $body.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.4 }
            $big = 384; $bx = (PX 60.5) - $big / 2; $by = (PY 52.4) - $big / 2
            if ($Images.Count) { $ph = Photo $Images[0] $bx $by $big $big 4 } else { $ph = Box 4 $bx $by $big $big $acc; Grad $ph '#0A1E6E' $Accent 90 }
            $spots = @(@(61.6, 52.5), @(73.8, 30.6), @(85.5, 51.6), @(73.4, 73.3)); $d = 157
            for ($i = 0; $i -lt [math]::Min(4, $n); $i++) {
                $cx = PX $spots[$i][0]; $cy = PY $spots[$i][1]
                [void](Box 4 ($cx - $d / 2) ($cy - $d / 2) $d $d $(if ($i) { Tone 0 } else { $acc }))
                [void](Say $list[$i][0] 12 0xFFFFFF $cx $cy $Font 2)
            }
        }
        'Phone' {
            if ($Kicker) { [void](Say $Kicker 12 (ConvertTo-Bgr '#999999') (PX 8.4) (PY 11.9)) }
            if ($Title) { $t = New-FitText $Title 30 $acc $Font (PX 50); $t.Left = [single](PX 8.4); $t.Top = [single](PY 15.5); $t.Name = 'Title' }
            [void](Line (PX 8.7) (PY 36.2) (PX 17.2) (PY 36.2) $acc 3)
            if ($Sub) { [void](Say $Sub 12 $grey (PX 8.5) (PY 43)) }
            for ($i = 0; $i -lt [math]::Min(2, $n); $i++) {
                $it = $list[$i]; $top = PY (52.2 + 20.7 * $i); $ch = PY 17.4; $on = [bool]$i
                $card = Round (PX 7.8) $top (PX 46.5) $ch $(if ($on) { $acc } else { 0xFFFFFF }) 0.3; Soft $card 16 4 0.8
                $tint = if ($on) { 0xFFFFFF } else { $acc }
                [void](Say $it[0] 14 $tint (PX 10.2) ($top + $ch * 0.3) $Font); [void](Say "$($it[1])%" 30 $tint (PX 10.2) ($top + $ch * 0.68) $Font)
                $full = [math]::Round([double]$it[1] / 20); $ph = PY 9
                for ($k = 0; $k -lt 5; $k++) {
                    if ($on) { Person (PX (30.7 + 5.6 * $k)) ($top + ($ch - $ph) / 2) $ph $(if ($k -lt $full) { 0xFFFFFF } else { ConvertTo-Bgr '#8FB0F2' }) }
                    else { Person (PX (30.7 + 5.6 * $k)) ($top + ($ch - $ph) / 2) $ph $(if ($k -lt $full) { $acc } else { Tone 0 }) }
                }
            }
            $pl = PX 70.5; $pt = PY 9.8; $pw = PX 21.2; $phh = PY 80.2
            $body = Round $pl $pt $pw $phh 0x111111 0.14; Soft $body 18 6 0.75
            if ($Images.Count) { [void](Photo $Images[0] ($pl + 8) ($pt + 8) ($pw - 16) ($phh - 16) 5) }
            else {                                                   # no screenshot: a plain app screen
                $scr = Round ($pl + 8) ($pt + 8) ($pw - 16) ($phh - 16) $acc 0.12; Grad $scr '#55BEF5' $Accent 90
                foreach ($q in 0.2, 0.42, 0.64) { $c2 = Round ($pl + 22) ($pt + $phh * $q) ($pw - 44) ($phh * 0.17) 0xFFFFFF 0.18; $c2.Fill.Transparency = 0.1 }
            }
            [void](Round ($pl + $pw / 2 - 28) ($pt + 16) 56 14 0x111111 0.5)
            if ($Images.Count -gt 1) {
                $tile = Round (PX 61.5) (PY 61) (PX 16) (PX 16) 0xFFFFFF 0.2; Soft $tile 16 4 0.75
                $ic = Pic $Images[1] (PX 12); [void](Put $ic ((PX 69.5) - [double]$ic.Width / 2) ((PY 61) + (PX 2)))
            }
        }
        'Scale' {
            [void](Box 1 0 0 $W (PY 2) $acc); [void](Box 1 0 (PY 85.3) $W (PY 14.7) (ConvertTo-Bgr '#F0F0F0'))
            if ($Title) { [void](Say $Title 28 $acc (PX 5.9) (PY 16.5) $Font) }
            $mk = @($Sub -split '\|'); if ($mk[0]) { [void](Say $mk[0] 14 (ConvertTo-Bgr '#666666') (PX 5.9) (PY 26.2)) }
            if ($mk.Count -gt 1) { [void](Say $mk[1] 14 $acc (PX 5.9) (PY 31.1) $Font) }
            $metal = ConvertTo-Bgr '#BDBDBD'; $px = PX 53.3; $py = PY 32.4
            [void](Box 1 (PX 44.8) (PY 90.3) (PX 17.5) (PY 2) $metal)
            $dome = Box 142 (PX 47.8) (PY 86.4) (PX 11.5) (PY 7.8) $metal; $dome.Adjustments.Item(1) = 180; $dome.Adjustments.Item(2) = 0      # pie: upper half only
            [void](Box 1 (PX 51.5) (PY 80) (PX 3.7) (PY 6.4) $metal); [void](Box 1 (PX 52.8) (PY 32.4) (PX 1.3) (PY 48) $metal)
            $a0 = [double]$list[0][1]; $b0 = [double]$list[1][1]; $tilt = 30 * [math]::Max(-1, [math]::Min(1, ($a0 - $b0) / 50))     # degrees, left side down when positive
            $len = PX 36.4; $beam = Box 1 ($px - $len / 2) ($py - 4.7) $len 9.4 $metal; $beam.Rotation = [single](-$tilt)
            [void](Box 9 ($px - (PX 2.45)) ($py - (PX 2.45)) (PX 4.9) (PX 4.9) $metal)
            $rad = $tilt * [math]::PI / 180
            for ($i = 0; $i -lt 2; $i++) {
                $it = $list[$i]; $side = if ($i) { 1 } else { -1 }; $tone = Tone $i
                $ex = $px + $side * $len / 2 * [math]::Cos($rad); $ey = $py - $side * $len / 2 * [math]::Sin($rad); $ay = $ey + (PY 6.3)
                [void](Line $ex $ey $ex $ay (ConvertTo-Bgr '#9E9E9E') 1.25)
                $tw = PX 14.9; $th = PY 24; [void](Box 7 ($ex - $tw / 2) $ay $tw $th $null (ConvertTo-Bgr '#9E9E9E') 1.5)      # 7 = isosceles triangle
                [void](Round ($ex - $tw / 2) ($ay + $th) $tw (PY 2.4) $metal 0.5); [void](Box 9 ($ex - 6) ($ay - 6) 12 12 $tone)
                $dd = PX 7.4; $cyc = $ay + $th - $dd / 2 - 8; [void](Box 9 ($ex - $dd / 2) ($cyc - $dd / 2) $dd $dd $tone); [void](Say ($it[0] -replace ' ', '|') 11 0xFFFFFF $ex $cyc $Font 2)      # a space in the label = line break inside the disc
                $nx = $ex + $side * ($tw / 2 + 14); $al = if ($i) { 1 } else { 3 }
                $num = Say "$($it[1])%" 28 $tone $nx ($ay + $th * 0.45) $Font $al
                if ($it.Count -gt 2) { [void](Say (Lines $it[2..($it.Count - 1)]) 9 $grey $nx ($ay + $th * 0.45 + 40) 'Arial' $al) }
            }
        }
        'AreaChart' {
            $y = Heading (PX 6.9) (PY 9.7) $acc 0xFFFFFF 22
            if ($Sub) { $sb = Say $Sub 11 0xFFFFFF (PX 6.9) 0; $sb.Top = [single]$y }
            $x0 = PX 6.9; $x1 = PX 92.6; $base = PY 64.9; $peakY = PY 21.7; $peak = 1; foreach ($it in $list) { $peak = [math]::Max($peak, [double]$it[1]) }
            [void](Line $x0 $base $x1 $base (ConvertTo-Bgr '#444444') 0.75)
            $pts = @(); for ($i = 0; $i -lt $n; $i++) { $pts += , @(($x0 + ($x1 - $x0) * $i / [math]::Max(1, $n - 1)), ($base - ($base - $peakY) * [double]$list[$i][1] / $peak)) }
            $area = Poly ($pts + , @($x1, $base) + , @($x0, $base)) $acc $null 0; Grad $area $Palette[0] $Accent 0; $area.Name = 'Area'       # by hand: Shapes > Freeform, click every point, close on the baseline
            for ($i = 0; $i -lt $n; $i++) {
                $last = ($i -eq $n - 1); $tint = if ($last) { $acc } else { ConvertTo-Bgr '#8A8A8A' }
                [void](Box 9 ($pts[$i][0] - 5.4) ($pts[$i][1] - 5.4) 10.8 10.8 $tint)
                [void](Say $list[$i][0] 9 $(if ($last) { $acc } else { ConvertTo-Bgr '#CFCFCF' }) $pts[$i][0] (PY 67.9) $(if ($last) { $Font } else { 'Arial' }) 2)
                if ($list[$i].Count -gt 2) { [void](Say $list[$i][2] 20 $acc $pts[$i][0] ($pts[$i][1] - 30) $Font $(if ($last) { 3 } else { 2 })) }
            }
            $cw = PX 27.8
            for ($i = 0; $i -lt [math]::Min(3, $notes.Count); $i++) {
                $cl = PX (6.9 + 28.95 * $i); [void](Box 1 $cl (PY 72.3) $cw (PY 20.8) $(if ($i -eq 2) { $acc } else { ConvertTo-Bgr '#262626' }))
                [void](Say $notes[$i][0] 28 0xFFFFFF ($cl + $cw - (PX 2.3)) (PY 79.9) $Font 3)
                if ($notes[$i].Count -gt 1) { [void](Say $notes[$i][1] 8 0xFFFFFF ($cl + (PX 2.1)) (PY 88)) }
            }
        }
        'FadeBars' {
            $mint = ConvertTo-Bgr '#26E08A'; $pl = PX 58.3
            if ($Images.Count) { [void](Photo $Images[0] $pl 0 ($W - $pl) $H) } else { $pn = Box 1 $pl 0 ($W - $pl) $H 0; Grad $pn '#1B1B1B' '#0A0A0A' 90 }
            $veil = Box 1 $pl (PY 38) ($W - $pl) (PY 62) $acc; Grad $veil $Accent $Accent 90 1 0; $veil.Name = 'Veil'       # clear at the top, solid at the bottom
            if ($notes.Count) {
                [void](Say $notes[0][0] 14 $mint (PX 67.5) (PY 71.2) $Font)
                $rest = @($notes[0][1..($notes[0].Count - 1)]) + @($notes[1..($notes.Count)] | ForEach-Object { $_[0] }) | Where-Object { $_ }
                if ($rest) { $bd = Say (Lines $rest) 11 0xFFFFFF (PX 67.5) 0; $bd.Top = [single](PY 76) }
            }
            $tl = @($Title -split '\|'); if ($tl[0]) { [void](Say $tl[0] 26 0xFFFFFF (PX 6.9) (PY 12.6) $Font) }
            if ($tl.Count -gt 1) { [void](Say $tl[1] 26 $mint (PX 6.9) (PY 19.7) $Font) }
            $peak = 1; foreach ($it in $list) { $peak = [math]::Max($peak, [double]$it[1]) }
            $bw = PX 10.2; $foot = PY 87
            for ($i = 0; $i -lt $n; $i++) {
                $it = $list[$i]; $bl = PX (12.3 + 12 * $i); $top = $foot - (PY 57.3) * [double]$it[1] / $peak
                $bar = Box 152 $bl $top $bw ($foot - $top) 0; Grad $bar $Palette[$i % $Palette.Count] $Background 90; $bar.Name = "Bar$($i + 1)"      # 152 = two top corners rounded
                if ($it.Count -gt 2) { [void](Say $it[2] 9 0xFFFFFF ($bl + $bw / 2) ($top + 16) 'Arial' 2) }
                [void](Say ($it[0] -replace ' ', '|') 10 0xFFFFFF ($bl + $bw / 2) (PY 78) $Font 2)
                $tile = Round ($bl + $bw / 2 - 16) ($top - 40) 32 32 0xFFFFFF 0.22; [void](Box 9 ($bl + $bw / 2 - 9) ($top - 33) 18 18 (Tone $i))
            }
            if ($Sub) { [void](Say $Sub 11 $mint (PX 30.1) (PY 93) $Font 2) }
        }
        'ProcessArrow' {
            $hd = Box 1 0 0 $W (PY 22.6) $acc; Grad $hd $Palette[0] $Accent 0
            if ($Kicker) { [void](Say $Kicker 14 (ConvertTo-Bgr '#1CA4F0') (PX 5.6) (PY 7.7)) }
            if ($Title) { [void](Say $Title 32 0xFFFFFF (PX 5.6) (PY 14.2) $Font) }
            $band = Box 51 0 (PY 40.6) (PX 78.4) (PY 22.6) $acc; Grad $band $Accent $Palette[0] 0; $band.Name = 'Arrow'       # 51 = pentagon arrow
            $cy = PY 51.9; $d = PX 14.5; $steps = [math]::Max(1, $n - 1)
            for ($i = 0; $i -lt $n; $i++) {
                $it = $list[$i]; $goal = ($i -eq $n - 1 -and $n -gt 1); $cx = if ($goal) { PX 86.7 } else { PX (16.7 + 20.15 * $i * 2 / [math]::Max(1, $steps - 1)) }
                if ($steps -eq 1) { $cx = if ($goal) { PX 86.7 } else { PX 37 } }
                if (-not $goal) { $ring = Box 9 ($cx - $d / 2) ($cy - $d / 2) $d $d $acc 0xFFFFFF 5; Grad $ring $Accent $Palette[0] 90; $ring.Line.ForeColor.RGB = 0xFFFFFF; $ring.Line.Weight = 5 }
                $iy = if ($goal) { $cy - 30 } else { $cy }
                if ($i -lt $Images.Count) { $ic = Pic $Images[$i] ($d * 0.6); [void](Put $ic ($cx - [double]$ic.Width / 2) ($iy - [double]$ic.Height / 2)) }
                elseif (-not $goal) { [void](Say "$($i + 1)" 36 0xFFFFFF $cx $cy $Font 2) }
                $ty = if ($goal) { $cy + 44 } else { PY 71.5 }
                [void](Say $it[0] 12 $ink $cx $ty $Font 2)
                if ($it.Count -gt 1) { [void](Say (Lines $it[1..($it.Count - 1)]) 9 (ConvertTo-Bgr '#555555') $cx ($ty + 34) 'Arial' 2) }
            }
        }
        'SlantChart' {
            $pg = Box 2 (PX 44.4) (PY 24.4) (PX 45.1) (PY 75.6) $acc; Grad $pg '#55BEF5' $Accent 90; $pg.Name = 'Slant'       # 2 = parallelogram
            if ($Images.Count) { [void](Photo $Images[0] (PX 53) 0 (PX 45) (PY 77) 2) } else { [void](Box 2 (PX 53) 0 (PX 45) (PY 77) (ConvertTo-Bgr '#3A4A60')) }
            $bar0 = Box 1 (PX 6) (PY 10.5) (PX 6) (PY 1.1) $acc; Grad $bar0 $Accent '#55BEF5' 0
            if ($Title) { [void](Say $Title 24 $ink (PX 6) (PY 18.4) $Font) }
            if ($Sub) { [void](Say $Sub 10.5 (ConvertTo-Bgr '#666666') (PX 6) (PY 25.8)) }
            $card = Round (PX 7) (PY 34) (PX 67) (PY 51) 0xFFFFFF 0.07; Soft $card 14 3 0.8
            $cl = PX 10; $cr = PX 50; $ct = PY 38; $cb = PY 77; $peak = 1
            foreach ($it in $list) { foreach ($v in $it[1..($it.Count - 1)]) { $peak = [math]::Max($peak, [double]$v) } }
            for ($g = 0; $g -le 4; $g++) { $gy = $cb - ($cb - $ct) * $g / 4; [void](Line $cl $gy $cr $gy (ConvertTo-Bgr '#E3E3E3') 0.75) }
            $gw = ($cr - $cl) / [math]::Max(1, $n)
            for ($i = 0; $i -lt $n; $i++) {                               # ponytail: bars are rectangles, not a chart object; Insert > Chart when the numbers must stay editable as data
                $it = $list[$i]; $series = $it.Count - 1; $bw = $gw * 0.62 / $series
                for ($k = 0; $k -lt $series; $k++) { $bh = ($cb - $ct) * [double]$it[$k + 1] / $peak * 0.92; [void](Box 1 ($cl + $gw * $i + $gw * 0.19 + $bw * $k) ($cb - $bh) ($bw - 2) $bh (Tone $k)) }
                [void](Say $it[0] 8 $grey ($cl + $gw * ($i + 0.5)) ($cb + 12) 'Arial' 2)
            }
            $kx = PX 53.6; $kw = PX 18.4
            $k1 = Round $kx (PY 38.5) $kw (PY 19.5) (ConvertTo-Bgr '#505050') 0.12; $k2 = Round $kx (PY 60.8) $kw (PY 19.8) $acc 0.12; Grad $k2 '#55BEF5' $Accent 90
            if ($notes.Count) { [void](Say $notes[0][0] 20 0xFFFFFF ($kx + $kw / 2) (PY 48.2) $Font 2) }
            if ($notes.Count -gt 1) { [void](Say (Lines ($notes[1..($notes.Count - 1)] | ForEach-Object { $_[0] })) 12 0xFFFFFF ($kx + $kw / 2) (PY 70.7) $Font 2) }
        }
        'OrgChart' {
            if ($Kicker) { [void](Say $Kicker 14 (ConvertTo-Bgr '#8A8A8A') (PX 7.1) (PY 11.5)) }
            if ($Title) { [void](Say $Title 28 $ink (PX 7.1) (PY 18.5) $Font) }
            $hx = PX 49.4; $hy = PY 53.9; $spots = @(@(15.7, 61.2), @(70.7, 31.3), @(74.9, 73.7)); $teams = [math]::Min(3, $n)
            for ($i = 0; $i -lt $teams; $i++) { [void](Line $hx $hy (PX $spots[$i][0]) (PY $spots[$i][1]) (Tone $i) 1) }
            for ($i = 0; $i -lt $teams; $i++) {
                $it = $list[$i]; $tone = Tone $i; $cx = PX $spots[$i][0]; $cy = PY $spots[$i][1]; $d = PX 16.1; $pd = PX 10.8
                $halo = Box 9 ($cx - $d / 2) ($cy - $d / 2) $d $d $tone $tone 1; $halo.Fill.ForeColor.RGB = [int]$bg; $halo.Name = "Halo$($i + 1)"
                $tint = Box 9 ($cx - $d / 2) ($cy - $d / 2) $d $d $tone; $tint.Fill.Transparency = 0.72           # fill transparency about 70 %: the pale disc is the team colour itself
                [void](Round ($cx + (PX 5.1)) ($cy - (PY 5.7)) (PX 12.8) (PY 11.4) $tone 0.5)
                [void](Say $it[0] 11 0xFFFFFF ($cx + (PX 12.2)) ($cy - 9) $Font 2)
                if ($it.Count -gt 1) { [void](Say $it[1] 9 0xFFFFFF ($cx + (PX 12.2)) ($cy + 10) 'Arial' 2) }
                if ($i + 1 -lt $Images.Count) { $face = Photo $Images[$i + 1] ($cx - $pd / 2) ($cy - $pd / 2) $pd $pd 9; $face.Line.Visible = -1; $face.Line.ForeColor.RGB = 0xFFFFFF; $face.Line.Weight = 6 }
                else { [void](Box 9 ($cx - $pd / 2) ($cy - $pd / 2) $pd $pd 0xFFFFFF 0xFFFFFF 6); Person $cx ($cy - $pd * 0.3) ($pd * 0.6) $tone }
                $at = @(@(4.9, -10.7), @(3.8, 14.8))
                for ($k = 2; $k -lt [math]::Min(4, $it.Count); $k++) {
                    $mx = $cx + (PX $at[$k - 2][0]); $my = $cy + (PY $at[$k - 2][1]); $md = PX 3.4
                    [void](Box 9 ($mx - $md / 2) ($my - $md / 2) $md $md $tone); Person $mx ($my - $md * 0.32) ($md * 0.64) 0xFFFFFF
                    [void](Say $it[$k] 8 $tone ($mx + $md / 2 + 6) $my)
                }
            }
            $rd = PX 22.7
            if ($Images.Count) { $boss = Photo $Images[0] ($hx - $rd / 2) ($hy - $rd / 2) $rd $rd 9; $boss.Line.Visible = -1; $boss.Line.ForeColor.RGB = [int]$acc; $boss.Line.Weight = 16 }
            else { [void](Box 9 ($hx - $rd / 2) ($hy - $rd / 2) $rd $rd (ConvertTo-Bgr '#E9EDF7') $acc 16); Person $hx ($hy - $rd * 0.3) ($rd * 0.6) $acc }
            $mk = @($Sub -split '\|'); if ($mk[0]) { $nm = Say $mk[0] 14 (ConvertTo-Bgr '#0A2E9E') $hx (PY 79) $Font 2 }
            if ($mk.Count -gt 1) { [void](Say $mk[1] 9 (ConvertTo-Bgr '#0A2E9E') $hx ((PY 79) + 20) 'Arial' 2) }
        }
        'Donut3D' {
            [void](Box 1 0 (PY 24.8) $W (PY 55.2) (ConvertTo-Bgr '#F0F0F0'))
            if ($Kicker) { [void](Say $Kicker 11 (ConvertTo-Bgr '#888888') ($W / 2) (PY 10.7) 'Arial' 2) }
            if ($Title) { [void](Say $Title 26 $ink ($W / 2) (PY 16.6) $Font 2) }
            $dash = Line (PX 50.2) (PY 31) (PX 50.2) (PY 74) (ConvertTo-Bgr '#BDBDBD') 1; $dash.Line.DashStyle = 4
            for ($i = 0; $i -lt [math]::Min(2, $n); $i++) {
                $it = $list[$i]; $tone = Tone $i; $cx = PX @(31, 70.6)[$i]; $cy = PY 57.5; $d = 190
                [void](Say $it[0] 15 $tone $cx (PY 31.3) $Font 2)
                $ring = Box 18 ($cx - $d / 2) ($cy - $d / 2) $d $d (ConvertTo-Bgr '#C8C8C8'); $ring.Adjustments.Item(1) = 0.25; $ring.Name = "Ring$($i + 1)"
                $arc = Box 20 ($cx - $d / 2) ($cy - $d / 2) $d $d $tone; $arc.Name = "Arc$($i + 1)"       # 20 = block arc: start angle, end angle, thickness
                $arc.Adjustments.Item(1) = 270; $arc.Adjustments.Item(2) = [single]((270 + 3.6 * [double]$it[1]) % 360); $arc.Adjustments.Item(3) = 0.25
                foreach ($part in @(@($ring, 6), @($arc, 50))) {         # Format Shape > 3-D Rotation X 314.7 Y 324.6 Z 60, 3-D Format > Depth 50 pt (the grey ring stays almost flat)
                    $td = $part[0].ThreeD; $td.Visible = -1; $td.SetPresetCamera(1); $td.RotationX = 314.7; $td.RotationY = 324.6; $td.RotationZ = 60; $td.Depth = [single]$part[1]
                }
                [void](Say "$($it[1])%" 16 $tone ($cx + $(if ($i) { 1 } else { -1 }) * ($d / 2 + 34)) ($cy - 40) $Font 2)
                if ($it.Count -gt 2) { [void](Say (Lines $it[2..($it.Count - 1)]) 9 (ConvertTo-Bgr '#666666') $cx (PY 88) 'Arial' 2) }
            }
        }
        'PhotoTimeline' {
            [void](Box 1 0 0 (PX 9.4) $H $acc)
            if ($Kicker) { $word = Say $Kicker 20 0xFFFFFF (PX 4.7) (PY 28) $Font 2; $word.TextFrame2.TextRange.Font.Spacing = 2; $word.Rotation = 90 }
            if ($Title) { [void](Say $Title 20 $acc (PX 19) (PY 14.2) $Font) }
            if ($Sub) { [void](Say $Sub 9.5 (ConvertTo-Bgr '#666666') (PX 19) (PY 19.3)) }
            [void](Line (PX 18.7) (PY 29.6) (PX 91.6) (PY 29.6) (ConvertTo-Bgr '#4A86F5') 1)
            for ($i = 0; $i -lt [math]::Min(4, $n); $i++) {
                $it = $list[$i]; $cl = PX (18.9 + 19.37 * $i); $cw = PX 14.3; $mid = $cl + $cw / 2
                [void](Say $it[0] 14 $acc $mid (PY 25.2) $Font 2); [void](Box 1 ($mid - 3.8) ((PY 29.6) - 3.8) 7.6 7.6 $acc)
                if ($i -lt $Images.Count) { [void](Photo $Images[$i] $cl (PY 34) $cw (PY 40)) } else { [void](Box 1 $cl (PY 34) $cw (PY 40) (ConvertTo-Bgr '#D9DEE8')) }
                $sq = PX 7.1; [void](Box 1 $cl (PY 65) $sq $sq $acc); [void](Say ('{0:00}' -f ($i + 1)) 18 0xFFFFFF ($cl + $sq / 2) ((PY 65) + $sq / 2) $Font 2)
                if ($it.Count -gt 1) { [void](Say $it[1] 12 $ink $cl (PY 81.5) $Font) }
                if ($it.Count -gt 2) { [void](Say (Lines $it[2..($it.Count - 1)]) 8 (ConvertTo-Bgr '#555555') $cl (PY 87)) }
            }
        }
        'ProCon' {
            if ($Kicker) { [void](Say $Kicker 16 $ink ($W / 2) (PY 10.5) 'Arial' 2) }
            if ($Title) { [void](Say $Title 32 $ink ($W / 2) (PY 18.5) $Font 2) }
            $mk = @($Sub -split '\|'); $sides = @($list, $notes); $pw = PX 41.5; $tabW = PX 15.2
            for ($i = 0; $i -lt 2; $i++) {                               # panel = two rounded rectangles united; the upper one stops short and leaves the notch (by hand: Merge Shapes > Union, then Edit Points to round the inner corner)
                $pl = PX @(7.1, 51.4)[$i]
                $lo = Round $pl (PY 39.8) $pw (PY 49.7) 0 0.1; $lo.Name = "Lo$i"
                $up = Round $(if ($i) { $pl + $tabW } else { $pl }) (PY 29.2) ($pw - $tabW) (PY 30) 0 0.16; $up.Name = "Up$i"
                $panel = Merge-Shapes 1 @("Lo$i", "Up$i"); $panel.Fill.ForeColor.RGB = (ConvertTo-Bgr (Shade $Palette[$i % $Palette.Count] 0.9)); $panel.Line.Visible = 0; $panel.Name = "Panel$($i + 1)"
            }
            $vs = Say 'VS' 120 0xFFFFFF ($W / 2) (PY 66) $Font 2; $vs.TextFrame2.TextRange.Font.Fill.Transparency = 0.25
            for ($i = 0; $i -lt 2; $i++) {
                $tone = Tone $i; $deep = ConvertTo-Bgr (Shade $Palette[$i % $Palette.Count] (-0.25)); $al = if ($i) { 3 } else { 1 }
                $pill = Round (PX @(33.6, 51.9)[$i]) (PY 30.2) (PX 14.5) (PY 8.8) $tone 0.5
                if ($mk.Count -gt $i) { [void](Say $mk[$i] 22 0xFFFFFF ((PX @(33.6, 51.9)[$i]) + (PX 7.25)) (PY 34.6) $Font 2) }
                for ($k = 0; $k -lt [math]::Min(3, $sides[$i].Count); $k++) {
                    $it = $sides[$i][$k]; $y = PY (45 + 15.5 * $k); $dotX = PX @(11.5, 88.5)[$i]; $tx = PX @(14.5, 85.5)[$i]
                    [void](Box 9 ($dotX - 12) ($y - 12) 24 24 $tone); [void](Say "$($k + 1)" 11 0xFFFFFF $dotX $y $Font 2)
                    [void](Say $it[0] 16 $deep $tx $y $Font $al)
                    if ($it.Count -gt 1) { $d2 = Say (Lines $it[1..($it.Count - 1)]) 9.5 (ConvertTo-Bgr '#555555') $tx 0 'Arial' $al; $d2.Top = [single]($y + 14) }
                }
            }
        }
        'SignPost' {
            $pale = ConvertTo-Bgr '#CFCFCF'; $slope = 0.25
            if ($Kicker) { $word = Say $Kicker 40 $pale (PX 6.5) (PY 30) $Font 2; $word.Rotation = 270 }
            [void](Line (PX 6.6) (PY 58) (PX 6.6) (PY 94) $pale 2)
            $cx = PX 27.8; $hw = PX 6.4; $bw = PX 8.8; $rise = $hw * $slope; $br = $bw * $slope; $fh = (PY 22.1) - 2 * $br; $rows = [math]::Min(3, $n)
            for ($i = 0; $i -lt $rows; $i++) {                            # the dark inside of each band, behind the post
                $t = PY (23.2 + 21.4 * $i)
                [void](Poly @(@(($cx - $bw), ($t + $br)), @($cx, $t), @(($cx + $bw), ($t + $br)), @($cx, ($t + 2 * $br))) (ConvertTo-Bgr (Shade $Palette[$i % $Palette.Count] (-0.85))) $null 0)
            }
            # the post: by hand one square with Format Shape > 3-D Format > Depth 450 pt and a 3-D rotation; here its three visible faces as polygons, so it stands exactly upright
            $t0 = PY 13.8; $foot = PY 94.5
            [void](Poly @(@(($cx - $hw), ($t0 + $rise)), @($cx, ($t0 + 2 * $rise)), @($cx, $foot), @(($cx - $hw), ($foot - $rise))) (ConvertTo-Bgr '#616161') $null 0)
            [void](Poly @(@($cx, ($t0 + 2 * $rise)), @(($cx + $hw), ($t0 + $rise)), @(($cx + $hw), ($foot - $rise)), @($cx, $foot)) (ConvertTo-Bgr '#B9B9B9') $null 0)
            [void](Poly @(@(($cx - $hw), ($t0 + $rise)), @($cx, $t0), @(($cx + $hw), ($t0 + $rise)), @($cx, ($t0 + 2 * $rise))) (ConvertTo-Bgr '#E4E4E4') $null 0)
            for ($i = 0; $i -lt $rows; $i++) {
                $it = $list[$i]; $t = PY (23.2 + 21.4 * $i); $hexI = $Palette[$i % $Palette.Count]; $mid = $t + 1.5 * $br + $fh / 2
                [void](Poly @(@(($cx - $bw), ($t + $br)), @($cx, ($t + 2 * $br)), @($cx, ($t + 2 * $br + $fh)), @(($cx - $bw), ($t + $br + $fh))) (ConvertTo-Bgr $hexI) $null 0)
                [void](Poly @(@($cx, ($t + 2 * $br)), @(($cx + $bw), ($t + $br)), @(($cx + $bw), ($t + $br + $fh)), @($cx, ($t + 2 * $br + $fh))) (ConvertTo-Bgr (Shade $hexI 0.22)) $null 0)
                [void](Say "$($i + 1)" 32 0xFFFFFF ($cx + $bw / 2) $mid $Font 2)
                [void](Box 9 ((PX 39.9) - 10) ($mid - 10) 20 20 (ConvertTo-Bgr (Shade $hexI 0.22)))
                [void](Say "$($i + 1). $($it[0])" 18 $ink (PX 43.2) ($mid - 10) $Font)
                if ($it.Count -gt 1) { [void](Say $it[1] 10 (ConvertTo-Bgr '#666666') (PX 43.2) ($mid + 18)) }
            }
        }
        'FadeCards' {
            if ($Kicker) { [void](Box 1 (PX 35) (PY 8) (PX 30) (PY 5.3) $acc); [void](Say $Kicker 13 0xFFFFFF ($W / 2) (PY 10.65) 'Arial' 2) }
            if ($Title) { [void](Say $Title 30 $ink ($W / 2) (PY 19.5) $Font 2) }
            for ($i = 0; $i -lt [math]::Min(4, $n); $i++) {
                $it = $list[$i]; $deep = ($i % 2 -eq 0); $cl = PX (9.2 + 20.5 * $i); $cw = PX 19.3; $ct = PY $(if ($deep) { 27.9 } else { 32.2 }); $chh = PY 58.5
                if ($i -lt $Images.Count) { [void](Photo $Images[$i] $cl $ct $cw $chh) } else { [void](Box 1 $cl $ct $cw $chh (ConvertTo-Bgr '#8C96A8')) }
                $paper = if ($deep) { $Accent } else { '#FFFFFF' }       # same-colour rectangle over the photo: solid down to the middle, clear at the bottom
                $cv = Box 1 $cl $ct $cw $chh 0; Grad $cv $paper $paper 90 0 1; [void]$cv.Fill.GradientStops.Insert((ConvertTo-Bgr $paper), [single]0.5, 0, 2); $cv.Name = "Cover$($i + 1)"
                $tint = if ($deep) { 0xFFFFFF } else { [int]$acc }; $tx = $cl + (PX 1.6)
                [void](Say $it[0] 18 $tint $tx ($ct + $chh * 0.14) $Font)
                if ($it.Count -gt 1) { [void](Say $it[1] 11 $tint $tx ($ct + $chh * 0.23)) }
                if ($it.Count -gt 2) { $b2 = Say (Lines $it[2..($it.Count - 1)]) 9 $tint $tx 0; $b2.Top = [single]($ct + $chh * 0.32) }
            }
        }
        'StepArrows' {
            [void](Box 1 (PX 3.4) (PY 4.8) (PX 0.7) (PY 6.7) $acc)
            $mk = @($Kicker -split '\|'); if ($mk[0]) { [void](Say $mk[0] 12 $acc (PX 5.4) (PY 6.6) $Font) }
            if ($mk.Count -gt 1) { [void](Say $mk[1] 12 $ink (PX 5.4) (PY 9.8) $Font) }
            $tl = @($Title -split '\|' | Where-Object { $_ })
            for ($k = 0; $k -lt $tl.Count; $k++) { [void](Say $tl[$k] 28 $(if ($k) { $ink } else { $acc }) (PX 14) (PY (31 + 7.6 * $k)) $Font) }
            if ($Sub) { $sb = Say $Sub 11 $grey (PX 14) 0; $sb.Top = [single](PY (30 + 7.6 * $tl.Count)) }
            for ($i = 0; $i -lt [math]::Min(3, $n); $i++) {
                $it = $list[$i]; $lit = ($i -eq 1); $al = PX (52.2 + 10.2 * $i); $aw = PX 9.9; $top = PY (24 + 10.5 * $i); $cx = $al + $aw / 2; $hd = PX 7.5
                $shaft = Poly @(@($al, ($top + 26)), @($cx, $top), @(($al + $aw), ($top + 26)), @(($al + $aw), $H), @($al, $H)) $acc $null 0; $shaft.Name = "Arrow$($i + 1)"
                if (-not $lit) { Grad $shaft '#FFFFFF' '#D9DCE2' 90 }
                $knob = Box 9 ($cx - $hd / 2) ($top - $hd * 0.45) $hd $hd $(if ($lit) { $acc } else { 0xFFFFFF }); Soft $knob 10 4 0.85       # shadow: transparency 90 %, blur 10 pt, distance 4 pt
                [void](Say ('{0:00}' -f ($i + 1)) 16 $(if ($lit) { 0xFFFFFF } else { $acc }) $cx ($top + $hd * 0.05) $Font 2)
                $tint = if ($lit) { 0xFFFFFF } else { [int]$ink }; $ty = $top + $hd * 0.55 + 30
                [void](Say $it[0] 11 $tint $cx $ty $Font 2)
                if ($it.Count -gt 1) { $d3 = Say (Lines $it[1..($it.Count - 1)]) 8 $(if ($lit) { 0xFFFFFF } else { [int]$grey }) $cx 0 'Arial' 2; $d3.Top = [single]($ty + 12) }
            }
        }
        'Compare' {
            if ($Kicker) { [void](Say $Kicker 14 $ink ($W / 2) (PY 9.5) 'Arial' 2) }
            if ($Title) { [void](Say $Title 30 $ink ($W / 2) (PY 15.6) $Font 2) }
            $mk = @($Sub -split '\|'); $sides = @($list, $notes); $bw = PX 32.1; $bh = PY 8.1
            $pn = Box 1 (PX 41.8) (PY 24) (PX 16.6) (PY 71) 0; Grad $pn '#E5E5E7' '#F6F6F6' 0; Soft $pn 9 0 0.8
            for ($i = 0; $i -lt 2; $i++) {
                $tone = Tone $i; $bl = PX @(4.8, 63.2)[$i]
                if ($mk.Count -gt $i) { [void](Say $mk[$i] 18 $tone (PX @(21, 79)[$i]) (PY 23.5) $Font 2) }
                for ($k = 0; $k -lt [math]::Min(4, $sides[$i].Count); $k++) {
                    $it = $sides[$i][$k]; $bt = PY @(29.4, 46.4, 63, 79.7)[$k]; $share = [double]$it[0] / 100
                    $base = Box 1 $bl $bt $bw $bh 0xFFFFFF (ConvertTo-Bgr '#878787') 1; Soft $base 9 4 0.8
                    [void](Box 1 $(if ($i) { $bl } else { $bl + $bw * (1 - $share) }) $bt ($bw * $share) $bh $tone)                      # both sides grow from the middle column outward
                    if ($i) { [void](Say "$($it[0])%" 20 $ink ($bl + $bw - 10) ($bt + $bh / 2) $Font 3) } else { [void](Say "$($it[0])%" 20 $ink ($bl + 10) ($bt + $bh / 2) $Font 1) }
                    if ($it.Count -gt 1) { if ($i) { [void](Say $it[1] 9 $ink $bl ($bt + $bh + 12)) } else { [void](Say $it[1] 9 $ink ($bl + $bw) ($bt + $bh + 12) 'Arial' 3) } }
                    if ($i -eq 0 -and $it.Count -gt 2) {
                        [void](Box 9 (($W / 2) - 16) ($bt + $bh / 2 - 24) 32 32 $null $ink 1.5); [void](Say "$($k + 1)" 12 $ink ($W / 2) ($bt + $bh / 2 - 8) $Font 2)
                        [void](Say $it[2] 8 $ink ($W / 2) ($bt + $bh / 2 + 22) 'Arial' 2)
                    }
                }
            }
        }
        'Folders' {
            $sky = Box 1 0 0 $W $H 0; Grad $sky $Accent $Palette[8 % $Palette.Count] 90; $sky.Name = 'Sky'
            [void]$sky.Fill.GradientStops.Insert(0xFFFFFF, [single]0.35, 0, 2); [void]$sky.Fill.GradientStops.Insert(0xFFFFFF, [single]0.62, 0, 3)
            $teal = ConvertTo-Bgr '#2A575B'
            if ($Kicker) { [void](Say $Kicker 12 $teal ($W / 2) (PY 14) $Font 2) }
            if ($Title) { [void](Say $Title 30 $teal ($W / 2) (PY 21) $Font 2) }
            $count = [math]::Min(4, $n); $fw = PX 18.7; $top = PY 40.8; $fh = PY 30.4; $lip = PY 5.8
            for ($i = 0; $i -lt $count; $i++) {
                $fl = PX (6.5 + 22.8 * $i); $back = ConvertTo-Bgr $Palette[$i % $Palette.Count]
                $b1 = Round $fl $top $fw $fh $back 0.08; $b1.Name = "B$i"; $b2 = Round $fl ($top - (PY 3.2)) ($fw * 0.52) (PY 9) $back 0.3; $b2.Name = "T$i"
                $bk = Merge-Shapes 1 @("B$i", "T$i"); $bk.Fill.ForeColor.RGB = [int]$back; $bk.Line.Visible = 0; $bk.Name = "Back$i"
                if ($i -lt $Images.Count) { $pc = Pic $Images[$i] (PY 6); [void](Put $pc ($fl + ($fw - [double]$pc.Width) / 2) ($top + $fh * 0.45)); $pc.Name = "Pic$i" }      # folded away inside the folder
                $fr = Round $fl ($top + $lip) $fw ($fh - $lip) (ConvertTo-Bgr $Palette[(4 + $i) % $Palette.Count]) 0.08; $fr.Name = "Front$i"
            }
            Add-MorphSlide 1
            for ($i = 0; $i -lt $count; $i++) {
                $fl = PX (6.5 + 22.8 * $i); $drop = PY 4
                foreach ($nm in "Back$i", "Front$i") { $sl.Shapes($nm).Top = [single]([double]$sl.Shapes($nm).Top + $drop) }
                if ($i -lt $Images.Count) {
                    $pc = $sl.Shapes("Pic$i"); $pc.LockAspectRatio = -1; $pc.Height = [single](PY 24)
                    $pc.Left = [single]($fl + ($fw - [double]$pc.Width) / 2); $pc.Top = [single]($top + $lip + $drop + (PY 7) - [double]$pc.Height)
                }
                $it = $list[$i]; [void](Say $it[0] 11 (ConvertTo-Bgr '#77676F') ($fl + $fw / 2) (PY 83.5) 'Arial' 2)
                if ($it.Count -gt 1) { [void](Say $it[1] 13 $ink ($fl + $fw / 2) (PY 88) $Font 2) }
            }
        }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
