# Recipe 1: big text whose letters fill up to a ratio (label + number + suffix on one centred row).
# Example: .\text_fill.ps1 -Label TWOSOME -Number 75 -Out C:\work\twosome.pptx -Preview C:\work\twosome.mp4
param(
    [string]$Label = '',
    [Parameter(Mandatory)][string]$Number,
    [string]$Suffix = '%',
    [double]$Ratio = -1,                 # 0..1; default = Number / 100
    [string]$LabelColor = '#FFFFFF',
    [string]$AccentColor = '#D50032',
    [string]$Background = '#000000',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 96,
    [double]$Dome = 0.6,                 # how far the fill edge drops at the word ends, as a share of glyph height; 0 = flat
    [double]$Seconds = 2.5,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Ratio -lt 0) { $Ratio = [double]$Number / 100 }
if ($Ratio -lt 0 -or $Ratio -gt 1) { throw "Ratio must be within 0..1, got $Ratio" }
Assert-Output @($Out, $Preview) $Force.IsPresent
$accent = ConvertTo-Bgr $AccentColor; $labelC = ConvertTo-Bgr $LabelColor
$rise = 0.15                             # fill travel, as a share of slide height

try {
    New-Deck $Background
    $hasLabel = $Label -ne ''

    # 1. one centred row; shrink the font until it fits with side margins
    $fs = $FontSize
    while ($true) {
        $num = New-Text $Number $fs $labelC $Font
        $pct = if ($Suffix) { New-Text $Suffix ($fs * 0.46) $accent $Font } else { $null }
        $lab = if ($hasLabel) { New-Text $Label $fs $labelC $Font } else { $null }
        $gap = 0.32 * $fs
        $total = [double]$num.Width
        if ($lab) { $total += [double]$lab.Width + $gap }
        if ($pct) { $total += [double]$pct.Width + 8 }
        if ($total -le $W - 120) { break }
        foreach ($t in $num, $pct, $lab) { if ($t) { $t.Delete() } }
        $fs = $fs * ($W - 120) / $total * 0.98
    }
    $x = ($W - $total) / 2
    $top = ($H - [double]$num.Height) / 2
    if ($lab) { $lab.Left = $x; $lab.Top = $top; $x += [double]$lab.Width + $gap }
    $num.Left = $x; $num.Top = $top
    if ($pct) { $pct.Left = $x + [double]$num.Width + 8; $pct.Top = $top }

    # 2. outlines and exact glyph bounds
    $numLine = New-GlyphShape $num; $numLine.Name = 'NumberOutline'
    $pairs = @(, @($numLine, $accent))
    if ($lab) { $labLine = New-GlyphShape $lab; $labLine.Name = 'LabelOutline'; $pairs += , @($labLine, $labelC) }
    foreach ($p in $pairs) {
        $p[0].Fill.Visible = 0
        $p[0].Line.Visible = -1; $p[0].Line.ForeColor.RGB = $p[1]; $p[0].Line.Weight = 1.5
    }
    $nb = Get-Box $numLine
    if ($pct) {                              # suffix sits on the digits' baseline
        $pg = New-GlyphShape $pct; $pb = Get-Box $pg; $pg.Delete()
        $pct.Top = [double]$pct.Top + ($nb.T + $nb.H) - ($pb.T + $pb.H)
    }

    # 3. mask: slide-covering rectangle with the texts cut out
    $mask = New-Rect -20 -20 ($W + 40) ($H + 40)
    $cut = @($mask.Name); if ($lab) { $cut += $lab.Name }; $cut += $num.Name
    $mask = Merge-Shapes 4 $cut
    $mask.Name = 'Mask'
    $mask.Fill.Solid(); $mask.Fill.ForeColor.RGB = (ConvertTo-Bgr $Background)
    $mask.Line.Visible = 0

    # 4. fills with a domed top edge; the mean height of the arc is the ratio line
    $bottom = $nb.T + $nb.H + 15
    $level  = $nb.T + $nb.H * (1 - $Ratio)
    function New-Fill($box, $color, $drop, $clipL, $clipR, $name) {
        $apex = $level - $drop / 3
        if ($drop -le 0) {
            $f = New-Rect $clipL $level ($clipR - $clipL) ($bottom - $level)
        } else {
            $a = $box.W / 2; $cx = $box.L + $a
            $A = 1.2 * $a                                    # ellipse semi-axes: the word edge sits $drop below the apex
            $B = $drop / (1 - [math]::Sqrt(1 - [math]::Pow($a / $A, 2)))
            $f = $sl.Shapes.AddShape(9, $cx - $A, $apex, 2 * $A, 2 * $B)   # msoShapeOval
            $bodyH = $bottom - $apex - $B                    # a tall dome already reaches below the letters on its own
            if ($bodyH -gt 0) {
                $body = New-Rect ($cx - $A) ($apex + $B) (2 * $A) $bodyH
                $f = Merge-Shapes 1 @($f.Name, $body.Name)
            }
            $clip = New-Rect $clipL ($apex - 5) ($clipR - $clipL) ($bottom - $apex + 5)
            $f = Merge-Shapes 3 @($f.Name, $clip.Name)
        }
        $f.Name = $name
        $f.Fill.Solid(); $f.Fill.ForeColor.RGB = $color
        $f.Line.Visible = 0
        $f.Top = [double]$f.Top + $rise * $H             # start below the letters; the motion path lifts it into place
        $f.ZOrder(1)                                     # msoSendToBack
        $f
    }
    $fills = @()
    $numLeft = $nb.L - 30
    if ($lab) {
        $lb = Get-Box $labLine
        $numLeft = ($lb.L + $lb.W + $nb.L) / 2
        $fills += New-Fill $lb $labelC ($Dome * $nb.H) ($lb.L - 30) $numLeft 'FillLabel'
    }
    $fills += New-Fill $nb $accent ($Dome * $nb.H * 2 / 3) $numLeft ($nb.L + $nb.W + 30) 'FillNumber'

    # 5. outlines and suffix above the mask
    $numLine.ZOrder(0); if ($lab) { $labLine.ZOrder(0) }; if ($pct) { $pct.ZOrder(0) }

    # 6. motion: fills rise together along a straight path, easing out at the top
    foreach ($f in $fills) {
        $e = $sl.TimeLine.MainSequence.AddEffect($f, 0, 0, 2)        # msoAnimEffectCustom, WithPrevious
        $m = $e.Behaviors.Add(1)                                     # msoAnimTypeMotion
        $m.MotionEffect.Path = "M 0 0 L 0 -$rise E"
        $e.Timing.Duration = $Seconds
        $e.Timing.SmoothEnd = -1
    }

    Save-Deck $Out $Preview ([int][math]::Ceiling($Seconds + 2.5))
    "saved $Out (ratio $Ratio, font size $([math]::Round($fs, 1)))"
} catch { Close-Deck; throw }
