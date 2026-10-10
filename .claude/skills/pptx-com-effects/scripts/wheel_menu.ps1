# Rotating wheel between sections: a big dial sits half off the left edge with the section names written around
# it; every slide turns the dial one step (Morph), the section text rises in from below while the previous one
# leaves at the top, and the background picture changes. The dial is filled with "slide background fill", so it
# shows the un-darkened picture while a gradient rectangle darkens the rest.
# Example: .\wheel_menu.ps1 -Sections 'PART 01|PART 02|PART 03|PART 04' -Images a.jpg,b.jpg,c.jpg,d.jpg -Out C:\work\wheel.pptx -Preview C:\work\wheel.mp4
param(
    [Parameter(Mandatory)][string]$Sections,      # 'a|b|c|d' - one slide each, written around the dial
    [string]$Bodies = '',                         # 'text for a|text for b|...' shown under each heading
    [string[]]$Images = @(),                      # background per section (fewer than sections: reused in turn)
    [string]$Color = '#FFFFFF',
    [string]$Background = '#10141C',              # used when no picture is given
    [string]$Font = 'Arial Black',
    [double]$Duration = 1,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Backdrop([int]$k) { if ($Images.Count) { $sl.Background.Fill.UserPicture($Images[$k % $Images.Count]) } }
function Set-SoftShadow($shape) { $s = $shape.Shadow; $s.Visible = -1; $s.ForeColor.RGB = 0; $s.Transparency = 0.4; $s.Blur = 12; $s.OffsetX = 3; $s.OffsetY = 0; $s.Size = 103 }

try {
    New-Deck $Background
    $names = $Sections -split '\|'; $texts = if ($Bodies) { $Bodies -split '\|' } else { @() }
    $n = $names.Count; $step = 360 / $n; $c = ConvertTo-Bgr $Color
    $cx = 150; $cy = $H / 2; $r = 300; $hubR = 95
    Set-Backdrop 0
    # darkening layer: solid on the right where the text sits, lighter towards the dial
    $scrim = New-Rect 0 0 $W $H; $scrim.Line.Visible = 0; $scrim.Name = 'Scrim'
    $scrim.Fill.TwoColorGradient(1, 1); $scrim.Fill.GradientAngle = 0
    $st = $scrim.Fill.GradientStops
    $st.Item(1).Color.RGB = 0; $st.Item(1).Position = 0.2; $st.Item(1).Transparency = 0.7
    $st.Item(2).Color.RGB = 0; $st.Item(2).Position = 1; $st.Item(2).Transparency = 0.1
    # dial: flowchart "summing junction" (circle with a cross) so the turning is visible
    $dial = $sl.Shapes.AddShape(77, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r)); $dial.Name = 'Dial'
    $dial.Fill.Background(); $dial.Line.ForeColor.RGB = 0xFFFFFF; $dial.Line.Weight = 1; Set-SoftShadow $dial
    $group = @('Dial')
    for ($i = 0; $i -lt $n; $i++) {                                  # label i sits $step*i counter-clockwise from 3 o'clock
        $a = -$step * $i; $rad = $a * [math]::PI / 180
        $t = New-Text $names[$i] 20 $c $Font; $t.Name = "Label$($i + 1)"
        $t.Left = [single]($cx + $r * 0.7 * [math]::Cos($rad) - [double]$t.Width / 2); $t.Top = [single]($cy + $r * 0.7 * [math]::Sin($rad) - [double]$t.Height / 2)
        $t.Rotation = [single]((($a % 360) + 360) % 360)
        $group += $t.Name
    }
    $wheel = $sl.Shapes.Range($group).Group(); $wheel.Name = 'Wheel'
    $hub = $sl.Shapes.AddShape(9, [single]($cx - $hubR), [single]($cy - $hubR), [single](2 * $hubR), [single](2 * $hubR)); $hub.Name = 'Hub'
    $hub.Fill.Background(); $hub.Line.Visible = 0; Set-SoftShadow $hub
    $mark = $sl.Shapes.AddLine([single]$cx, [single]($cy - $hubR), [single]$cx, [single]($cy + $hubR)); $mark.Line.ForeColor.RGB = 0xFFFFFF; $mark.Line.Weight = 1; $mark.Name = 'HubMark'
    $hubGroup = $sl.Shapes.Range(@('Hub', 'HubMark')).Group(); $hubGroup.Name = 'HubWheel'
    $tx = 540; $homeY = @()
    for ($i = 0; $i -lt $n; $i++) {                                  # section text: all built on slide 1, parked below the slide
        $head = New-FitText $names[$i] 60 $c $Font ($W - $tx - 50); $head.Name = "Head$($i + 1)"; $head.Left = [single]$tx
        $y = $cy - [double]$head.Height / 2 - $(if ($i -lt $texts.Count) { 20 } else { 0 }); $homeY += $y
        $head.Top = [single]$(if ($i -eq 0) { $y } else { $y + $H })
        if ($i -lt $texts.Count) {
            $b = New-FitText $texts[$i] 16 $c 'Arial' ($W - $tx - 50); $b.Name = "Body$($i + 1)"; $b.Left = [single]$tx
            $b.Top = [single]([double]$head.Top + [double]$head.Height + 10)
        }
    }
    for ($k = 1; $k -lt $n; $k++) {
        Add-MorphSlide $Duration
        Set-Backdrop $k
        $sl.Shapes('Wheel').Rotation = [single]($step * $k)                # clockwise: the next label arrives at 3 o'clock
        $sl.Shapes('HubWheel').Rotation = [single]((360 - 45 * $k) % 360)  # hub turns the other way
        for ($i = 0; $i -lt $n; $i++) {
            $shift = if ($i -lt $k) { -$H } elseif ($i -gt $k) { $H } else { 0 }
            $head = $sl.Shapes("Head$($i + 1)"); $head.Top = [single]($homeY[$i] + $shift)
            if ($i -lt $texts.Count) { $sl.Shapes("Body$($i + 1)").Top = [single]($homeY[$i] + $shift + [double]$head.Height + 10) }
        }
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
