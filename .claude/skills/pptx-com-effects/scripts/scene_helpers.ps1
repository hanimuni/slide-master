# Drawing helpers shared by scene_page.ps1 and impact_scene.ps1. Dot-source AFTER ppt_helpers.ps1 and after the caller's
# param block: the colours below are mixed from the caller's $Accent / $AccentDark / $Light / $Dark / $Ink, and the
# text helpers use its $Font / $HeavyFont. Everything draws on the current slide ($sl).
function Mix([string]$a, [string]$b, [double]$f) {                    # $f of the way from colour a to colour b
    $x = [Convert]::ToInt32($a.TrimStart('#'), 16); $y = [Convert]::ToInt32($b.TrimStart('#'), 16); $o = '#'
    foreach ($sh in 16, 8, 0) { $p = ($x -shr $sh) -band 255; $q = ($y -shr $sh) -band 255; $o += '{0:X2}' -f [int]($p + ($q - $p) * $f) }
    $o
}
if (-not $AccentDark) { $AccentDark = Mix $Accent '#000000' 0.22 }
$tint = Mix $Accent '#FFFFFF' 0.86; $lilac = Mix $Accent '#FFFFFF' 0.82; $body = Mix $Ink '#FFFFFF' 0.22; $gray = Mix $Ink '#FFFFFF' 0.5
$soft = Mix $Accent '#FFFFFF' 0.45; $dText = Mix $Dark '#FFFFFF' 0.75; $dLine = Mix $Dark '#FFFFFF' 0.25; $dCard = Mix $Dark '#FFFFFF' 0.09; $dDeco = Mix $Dark '#FFFFFF' 0.05
$heavy = $(if ($HeavyFont) { $HeavyFont } else { $Font }); $heavyBold = $(if ($HeavyFont) { 0 } else { -1 })
$FADE = 10; $WIPE = 22; $WHEEL = 21; $RISE = 42; $AFTER = 3; $WITH = 2      # entrance effects and "after / with previous"
$script:uid = 0

function Tag($s) { $script:uid++; $s.Name = "g$($script:uid)"; $s }
# Text box. "|" or "//" in the text = line break (PowerPoint may wrap Korean inside a word, so give long lines their own breaks).
function Tx([string]$text, $size, [string]$hex, $x, $y, $wd, [int]$bold = 0, [string]$face = $Font, [int]$align = 1, $track = 0, $lead = 1.15) {
    $t = Tag ($sl.Shapes.AddTextbox(1, [single]$x, [single]$y, [single]$wd, 20))
    $f = $t.TextFrame; $f.WordWrap = -1; $f.AutoSize = 1; $f.MarginLeft = 0; $f.MarginRight = 0; $f.MarginTop = 0; $f.MarginBottom = 0
    $r = $f.TextRange; $r.Text = ($text -replace '\s*(\||//)\s*', "`r"); $r.Font.Name = $face; $r.Font.NameFarEast = $face; $r.Font.Size = [single]$size
    $r.Font.Color.RGB = (ConvertTo-Bgr $hex); $r.Font.Bold = $bold; $r.ParagraphFormat.Alignment = $align; $r.ParagraphFormat.SpaceWithin = [single]$lead
    if ($track) { $t.TextFrame2.TextRange.Font.Spacing = [single]$track }
    $t
}
function Mid($t, $y, $bh) { $t.TextFrame.AutoSize = 0; $t.Top = [single]$y; $t.Height = [single]$bh; $t.TextFrame.VerticalAnchor = 3; $t }   # centre a text vertically in a band
function Box($x, $y, $bw, $bh, [string]$hex, $rad = 0) {
    $s = Tag ($sl.Shapes.AddShape($(if ($rad) { 5 } else { 1 }), [single]$x, [single]$y, [single]$bw, [single]$bh)); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (ConvertTo-Bgr $hex)
    if ($rad) { $s.Adjustments.Item(1) = [single][math]::Min(0.5, $rad / [math]::Min($bw, $bh)) }
    $s
}
function Disc($cx, $cy, $r, [string]$hex) { $s = Tag ($sl.Shapes.AddShape(9, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r))); $s.Line.Visible = 0; $s.Fill.ForeColor.RGB = (ConvertTo-Bgr $hex); $s }
function Glow($s, [string]$hex, $transp = 0.72, $blur = 16, $oy = 5) { $h = $s.Shadow; $h.Visible = -1; $h.ForeColor.RGB = (ConvertTo-Bgr $hex); $h.Transparency = [single]$transp; $h.Blur = [single]$blur; $h.OffsetX = 0; $h.OffsetY = [single]$oy }
function Wd([string]$text, $size) { $n = 0; foreach ($ch in $text.ToCharArray()) { $n += $(if ([int]$ch -lt 128) { 0.58 } else { 1.0 }) }; $n * $size + 30 }   # rough width of a pill for this text
function Pill($x, $y, $bw, $bh, [string]$fill, [string]$text, $size, [string]$hex) { $b = Box $x $y $bw $bh $fill ($bh / 2); $t = Mid (Tx $text $size $hex $x $y $bw -1 $Font 2) $y $bh; @($b.Name, $t.Name) }
function Pic([string]$path, $x, $y, $side) { Tag ($sl.Shapes.AddPicture($path, 0, -1, [single]$x, [single]$y, [single]$side, [single]$side)) }
function Grp($names) { Tag ($sl.Shapes.Range([string[]]@($names | ForEach-Object { if ($_ -is [string]) { $_ } else { $_.Name } })).Group()) }
function Anim($shape, [int]$fx, [int]$trig = 3, $dur = 0.5, $delay = 0, [int]$dir = 0) {
    $e = $sl.TimeLine.MainSequence.AddEffect($shape, $fx, 0, $trig); $e.Timing.Duration = [single]$dur; $e.Timing.TriggerDelayTime = [single]$delay
    if ($dir) { $e.EffectParameters.Direction = $dir }           # 4 = from the left
}
function Head([bool]$onDark = $false) {
    if ($Kicker) { [void](Tx $Kicker 12.5 $(if ($onDark) { $soft } else { $Accent }) 65 45 600 -1 $Font 1 3) }
    if ($Title) { [void](Tx ($Title -replace '\|', ' ') 30 $(if ($onDark) { '#FFFFFF' } else { $Ink }) 65 68 830 -1) }
}
function Foot([bool]$onDark = $false) { if ($Note) { [void](Tx $Note 9 $(if ($onDark) { Mix $Dark '#FFFFFF' 0.5 } else { $gray }) 65 514 600) } }
function Rows([string]$spec) { $o = @($spec -split '\|' | Where-Object { $_ } | ForEach-Object { , @($_ -split ';') }); , $o }   # the leading comma keeps a single row from being unrolled
function Chips([string]$spec, $x, $y, [string]$fill, [string]$hex) {  # a row of pills from 'a|b|c'
    $names = @(); foreach ($c in @($spec -split '\|' | Where-Object { $_ })) { $cwid = Wd $c 15; $names += Pill $x $y $cwid 36 $fill $c 15 $hex; $x += $cwid + 10 }
    $names
}
