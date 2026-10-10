# Picture inside big letters that changes while a thin bar sweeps across (animated WELCOME title).
# Each picture is cut to the letter outlines (picture first, then the letters -> Merge Shapes > Intersect) and
# stacked; the later ones enter with a Wipe while the bar runs the same way on a motion path with the same timing.
# Example: .\text_picture_wipe.ps1 -Text WELCOME -Sub 'TO OUR PRESENTATION' -Images a.jpg,b.jpg -Out C:\work\w.pptx -Preview C:\work\w.mp4
param(
    [Parameter(Mandatory)][string]$Text,
    [Parameter(Mandatory)][string[]]$Images,      # two or more; shown in this order
    [string]$Sub = '',                            # small spaced line under the word
    [string]$BarColor = '#111111',
    [string]$Background = '#E9E9E9',
    [string]$Font = 'Impact',
    [double]$FontSize = 190,
    [double]$Seconds = 1.6,                       # length of one sweep
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $t = New-FitText $Text $FontSize 0 $Font ($W - 160); Set-Center $t
    if ($Sub) { $t.Top = [single]([double]$t.Top - 24) }
    $glyph = New-GlyphShape $t; $glyph.Name = 'Letters'; $t.Delete()
    $g = Get-Box $glyph
    $seq = $sl.TimeLine.MainSequence
    $bar = New-Rect ([single]($g.L - 26)) ([single]($g.T - 20)) 5 ([single]($g.H + 70)); $bar.Fill.ForeColor.RGB = (ConvertTo-Bgr $BarColor); $bar.Line.Visible = 0; $bar.Name = 'Bar'
    $travel = ($g.W + 52) / $W                                       # motion paths are in fractions of the slide size
    for ($i = 0; $i -lt $Images.Count; $i++) {
        $pic = $sl.Shapes.AddPicture($Images[$i], 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = "Src$i"
        $pic.Width = [single]$g.W; if ([double]$pic.Height -lt $g.H) { $pic.Height = [single]$g.H }       # cover the letters
        $pic.Left = [single]($g.L + ($g.W - [double]$pic.Width) / 2); $pic.Top = [single]($g.T + ($g.H - [double]$pic.Height) / 2)
        $cut = $glyph.Duplicate(); $cut.Left = [single]$g.L; $cut.Top = [single]$g.T; $cut.Name = "Cut$i"
        $shot = Merge-Shapes 3 @("Src$i", "Cut$i"); $shot.Name = "Picture$($i + 1)"
        if ($i -eq 0) { continue }
        $toRight = ($i % 2) -eq 1                                    # odd swaps sweep left-to-right, even ones come back
        $e = $seq.AddEffect($shot, 22, 0, 3)                          # Wipe, after previous
        $e.EffectParameters.Direction = $(if ($toRight) { 4 } else { 2 })      # from left / from right
        $e.Timing.Duration = $Seconds; $e.Timing.TriggerDelayTime = 1
        $m = $seq.AddEffect($bar, 0, 0, 2)                            # custom path, with previous
        $path = if ($toRight) { "M 0 0 L $travel 0 E" } else { "M $travel 0 L 0 0 E" }
        $m.Behaviors.Add(1).MotionEffect.Path = $path
        $m.Timing.Duration = $Seconds; $m.Timing.TriggerDelayTime = 1; $m.Timing.SmoothStart = 0; $m.Timing.SmoothEnd = 0
    }
    $glyph.Delete(); $bar.ZOrder(0)
    if ($Sub) {
        $s = New-Text $Sub 18 0x202020 'Georgia'; $s.TextFrame2.TextRange.Font.Spacing = 9
        $s.Left = [single]($g.L + ($g.W - [double]$s.Width) / 2); $s.Top = [single]($g.T + $g.H + 14); $s.Name = 'Sub'
    }
    Save-Deck $Out $Preview ([int](2 + ($Images.Count - 1) * ($Seconds + 1)))
    "saved $Out"
} catch { Close-Deck; throw }
