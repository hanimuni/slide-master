# Award-ceremony / year-end title: gold gradient text on black, optional looping video behind it,
# optional video playing inside the letters (video intersected with the text), optional gold frame.
# Example: .\gold_award.ps1 -Title 'GOODBYE 2025' -Sub 'YEAR-END CEREMONY' -Out C:\work\award.pptx -Preview C:\work\award.png
#          .\gold_award.ps1 -Title 'AWARDS' -Video C:\work\bokeh.mp4 -TextVideo C:\work\glitter.mp4 -Out C:\work\award.pptx -Preview C:\work\award.mp4
param(
    [Parameter(Mandatory)][string]$Title,         # '|' = line break
    [string]$Sub = '',                            # small line under the title (gold gradient too)
    [string]$Video = '',                          # full-slide background video; plays automatically and loops
    [string]$TextVideo = '',                      # video shown inside the title letters (the title is then no longer editable)
    [switch]$Frame,                               # thin gold frame around the title block
    [string[]]$Gold = @('#8C6B1F', '#F7E08B', '#B88A2A', '#FFF4BF'),   # four stops, sampled from a gold gradient image
    [string]$Background = '#000000',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 110,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($f in $Video, $TextVideo) { if ($f -and -not (Test-Path $f)) { throw "not found: $f" } }
if ($Gold.Count -ne 4) { throw '-Gold needs four colours' }
Assert-Output @($Out, $Preview) $Force.IsPresent

# Linear gradient left -> right ("linear right" in the UI): on wide text a top-to-bottom gradient looks flat.
function Set-GoldFill($fill) {
    $fill.TwoColorGradient(1, 1)
    $stops = $fill.GradientStops
    $stops.Insert((ConvertTo-Bgr $Gold[1]), 0.35, 0, 2); $stops.Insert((ConvertTo-Bgr $Gold[2]), 0.65, 0, 3)
    $pos = 0.0, 0.35, 0.65, 1.0
    for ($i = 1; $i -le 4; $i++) { $s = $stops.Item($i); $s.Color.RGB = (ConvertTo-Bgr $Gold[$i - 1]); $s.Position = $pos[$i - 1] }
    $fill.GradientAngle = 0
}
function Add-LoopVideo([string]$path) {
    $v = $sl.Shapes.AddMediaObject2($path, 0, -1, 0, 0, -1, -1); $v.LockAspectRatio = -1
    $v.Width = $W
    if ([double]$v.Height -lt $H) { $v.Height = $H }
    Set-Center $v
    $v
}
function Set-AutoLoop($v) {                       # Playback tab: Start = Automatically, Loop until stopped
    $p = $v.AnimationSettings.PlaySettings
    $p.PlayOnEntry = -1; $p.LoopUntilStopped = -1; $p.HideWhileNotPlaying = 0
    $v.AnimationSettings.AdvanceMode = 2          # ppAdvanceOnTime
}

try {
    New-Deck $Background
    if ($Video) { $bg = Add-LoopVideo $Video; $bg.Name = 'BackVideo'; Set-AutoLoop $bg }

    $t = New-FitText $Title $FontSize 0xFFFFFF $Font ($W - 260)
    $t.TextFrame.TextRange.Font.Bold = -1
    $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    Set-Center $t
    if ($Sub) { $t.Top = [single]([double]$t.Top - 22) }
    $t.Name = 'Title'
    $box = Get-Box $t

    if ($TextVideo) {
        # keep the video, cut it to the letters: video first, glyph shape second, intersect
        $g = New-GlyphShape $t; $t.Delete()
        $tv = Add-LoopVideo $TextVideo; $tv.Name = 'TextVideoRaw'
        $res = Merge-Shapes 3 @('TextVideoRaw', $g.Name); $res.Name = 'TitleVideo'
        Set-AutoLoop $res
    } else {
        Set-GoldFill $t.TextFrame2.TextRange.Font.Fill
    }
    if ($Sub) {
        $s = New-FitText $Sub 22 0xFFFFFF $Font ($W - 260)
        $s.TextFrame2.TextRange.Font.Spacing = 3
        $s.Left = [single](($W - [double]$s.Width) / 2); $s.Top = [single]($box.T + $box.H + 6); $s.Name = 'Sub'
        Set-GoldFill $s.TextFrame2.TextRange.Font.Fill
        $box.H = $box.H + 6 + [double]$s.Height
    }
    if ($Frame) {
        $m = 34; $th = 4
        $o = New-Rect ([single]($box.L - $m)) ([single]($box.T - $m)) ([single]($box.W + 2 * $m)) ([single]($box.H + 2 * $m))
        $i = New-Rect ([single]($box.L - $m + $th)) ([single]($box.T - $m + $th)) ([single]($box.W + 2 * $m - 2 * $th)) ([single]($box.H + 2 * $m - 2 * $th))
        $fr = Merge-Shapes 4 @($o.Name, $i.Name); $fr.Name = 'Frame'; $fr.Line.Visible = 0
        Set-GoldFill $fr.Fill
    }
    # every video starts together with the slide, not one after another
    $seq = $sl.TimeLine.MainSequence
    for ($k = 1; $k -le $seq.Count; $k++) { $seq.Item($k).Timing.TriggerType = $(if ($k -eq 1) { 3 } else { 2 }) }   # after previous / with previous
    Save-Deck $Out $Preview 5
    "saved $Out"
} catch { Close-Deck; throw }
