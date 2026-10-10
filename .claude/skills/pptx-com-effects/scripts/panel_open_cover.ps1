# Cover that opens: a blank panel fills the slide, the photo first shows only through the title letters, then the panel slides
# aside and settles as a title panel beside the full photo. One slide, plays by itself (no click).
# The letter holes start over the part of the photo at (title position + slide width - panel width): pick a photo that is bright there.
# Example: .\panel_open_cover.ps1 -Image C:\work\photo.jpg -Title 'NEW|START' -Kicker 'ANNUAL REPORT' -Sub 'one line under the title' -Out C:\work\cover.pptx -Preview C:\work\cover.mp4
param(
    [Parameter(Mandatory)][string]$Image,          # subject in the half the panel does not cover
    [Parameter(Mandatory)][string]$Title,          # '|' = line break; the last line takes the accent colour
    [string]$Kicker = '',                          # small line above the title
    [string]$Sub = '',                             # one line under the title
    [string]$Foot = '',                            # organisation, year
    [string]$PanelColor = '#FFFFFF',
    [string]$Ink = '#16181D',
    [string]$Accent = '#1E7B34',
    [double]$PanelWidth = 0.46,                    # share of the slide width the panel keeps at the end
    [ValidateSet('Left', 'Right')][string]$Side = 'Left',
    [string]$Font = 'Malgun Gothic',               # a heavy face reads best: the letters are windows
    [double]$FontSize = 80,
    [string]$Sound = '',                           # optional mp3/wav played when the panel opens (level it first, see ppt-craft 29)
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($f in @($Image, $Sound) | Where-Object { $_ }) { if (-not (Test-Path $f)) { throw "not found: $f" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Fx($shape, [int]$fx, [int]$trig, $dur, $delay, [int]$dir = 0, [switch]$Exit) {
    $e = $sl.TimeLine.MainSequence.AddEffect($shape, $fx, 0, $trig); $e.Timing.Duration = [single]$dur; $e.Timing.TriggerDelayTime = [single]$delay
    if ($dir) { $e.EffectParameters.Direction = $dir }
    if ($Exit) { $e.Exit = -1 }
}
function Line([string]$text, $size, [string]$hex, $x, $y) {
    $t = New-Text ($text -replace '\|', "`r") $size (ConvertTo-Bgr $hex) $Font; $t.Left = [single]$x; $t.Top = [single]$y; $t
}
function Heading($x, $y, [string]$hex) {                                   # the title, bold, shrunk to the panel
    $t = Line $Title $FontSize $hex $x $y; $t.TextFrame.TextRange.Font.Bold = -1
    if ([double]$t.Width -gt $room) { $t.TextFrame.TextRange.Font.Size = [single][math]::Floor($FontSize * $room / [double]$t.Width) }
    $t
}

try {
    New-Deck $PanelColor
    $sl.SlideShowTransition.EntryEffect = 0
    $FADE = 10; $WIPE = 22; $AFTER = 3; $WITH = 2
    $pw = $W * $PanelWidth; $shift = $W - $pw; $pad = 60; $room = $pw - 2 * $pad
    $dir = if ($Side -eq 'Left') { -1 } else { 1 }
    $x0 = if ($Side -eq 'Left') { $pad } else { $shift + $pad }             # where the title rests

    $photo = Add-Cover $Image; $photo.Name = 'Photo'
    $probe = Heading 0 0 $Ink; $ty = ($H - [double]$probe.Height) / 2; $probe.Delete()
    # the panel starts shifted so that it covers the whole slide; its holes sit where the title will be once it has moved
    $hole = Heading ($x0 - $dir * $shift) $ty $Ink; $glyph = New-GlyphShape $hole; $hole.Delete()
    $panel = New-Rect $(if ($Side -eq 'Left') { -$shift } else { 0 }) 0 ($W + $shift) $H
    $panel = Merge-Shapes 4 @($panel.Name, $glyph.Name); $panel.Name = 'Panel'
    $panel.Fill.ForeColor.RGB = (ConvertTo-Bgr $PanelColor); $panel.Line.Visible = 0

    $head = Heading $x0 $ty $Ink; $pc = $head.TextFrame.TextRange.Paragraphs().Count
    if ($pc -gt 1) { $head.TextFrame.TextRange.Paragraphs($pc).Font.Color.RGB = (ConvertTo-Bgr $Accent) }
    $bottom = $ty + [double]$head.Height; $extra = @()
    if ($Kicker) { $k = Line $Kicker 13 $Accent ($x0 + 3) ($ty - 34); $k.TextFrame.TextRange.Font.Bold = -1; $extra += $k }
    $rule = New-Rect ($x0 + 3) ($bottom + 14) 64 6; $rule.Fill.ForeColor.RGB = (ConvertTo-Bgr $Accent); $rule.Line.Visible = 0
    if ($Sub) { $extra += Line $Sub 17 $Ink ($x0 + 3) ($bottom + 34) }
    if ($Foot) { $extra += Line $Foot 11 $Ink ($x0 + 3) ($H - 46) }
    $lid = New-Rect 0 0 $W $H; $lid.Fill.ForeColor.RGB = (ConvertTo-Bgr $PanelColor); $lid.Line.Visible = 0      # the first moment is blank

    # a shape whose first effect is a custom one stays hidden until it starts: show the photo and the panel from 0
    foreach ($s in $photo, $panel) { Fx $s 1 $WITH 0.01 0 }
    Fx $lid $FADE $AFTER 1.5 0.4 -Exit                                      # the photo wells up inside the letters
    $e = $sl.TimeLine.MainSequence.AddEffect($panel, 0, 0, $AFTER)          # the panel slides aside
    $e.Behaviors.Add(1).MotionEffect.Path = ('M 0 0 L {0} 0 E' -f ([math]::Round($dir * $shift / $W, 4)).ToString([cultureinfo]::InvariantCulture))
    $e.Timing.Duration = 0.9; $e.Timing.TriggerDelayTime = 0.3
    $e = $sl.TimeLine.MainSequence.AddEffect($photo, 0, 0, $WITH)           # the photo settles from slightly larger
    $b = $e.Behaviors.Add(3).ScaleEffect; $b.FromX = 112; $b.FromY = 112; $b.ToX = 100; $b.ToY = 100
    $e.Timing.Duration = 1.1; $e.Timing.TriggerDelayTime = 0.3
    if ($Sound) {
        $m = $sl.Shapes.AddMediaObject2($Sound, 0, -1, 0, 0, -1, -1); $m.Left = [single](-80); $m.Top = [single](-80)
        $e = $sl.TimeLine.MainSequence.AddEffect($m, 83, 0, $WITH); $e.Timing.TriggerDelayTime = 0.15
    }
    Fx $head $FADE $WITH 0.4 1.15; Fx $rule $WIPE $WITH 0.4 1.5 4
    $d = 1.5; foreach ($s in $extra) { Fx $s $FADE $WITH 0.4 $d; $d += 0.15 }
    $hold = New-Rect -40 -40 10 10; $hold.Fill.Visible = 0; $hold.Line.Visible = 0; Fx $hold $FADE $AFTER 0.1 2   # stay 2 s in the video
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
