# Structured content page: one key sentence on top, then a row of identical cards (numbered circle, label, short
# text) inside a frame whose outline is a gradient that fades out at both ends. Build one card, then repeat it at
# equal spacing - never place the pieces one by one. The gradient outline is a real shape (outer rounded rectangle
# minus an inset copy), so it survives in every PowerPoint version.
# Example: .\card_row.ps1 -Key 'Digital twin covers every step' -Cards 'Design|Build|Test|Run' -Bodies 'a|b|c|d' -Out C:\work\cards.pptx -Preview C:\work\cards.png
param(
    [Parameter(Mandatory)][string]$Cards,         # 'label|label|...'
    [string]$Bodies = '',                         # 'text|text|...' (use ';' inside one for a line break)
    [string]$Key = '',                            # the one sentence the slide says
    [string]$Sub = '',
    [string]$Accent = '#2EC4B6',
    [string]$CardColor = '#1B2433',
    [string]$Color = '#FFFFFF',
    [string]$Background = '#0F1420',
    [string]$Font = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $labels = $Cards -split '\|'; $texts = if ($Bodies) { $Bodies -split '\|' } else { @() }
    $n = $labels.Count; $c = ConvertTo-Bgr $Color; $acc = ConvertTo-Bgr $Accent
    $margin = 60; $y = 56
    if ($Key) { $k = New-FitText $Key 30 $c $Font ($W - 2 * $margin); $k.Left = [single]$margin; $k.Top = [single]$y; $k.Name = 'Key'; $y += [double]$k.Height + 6 }
    if ($Sub) { $s2 = New-FitText $Sub 15 (ConvertTo-Bgr '#9AA6B8') 'Arial' ($W - 2 * $margin); $s2.Left = [single]$margin; $s2.Top = [single]$y; $y += [double]$s2.Height }
    $fy = $y + 26; $fh = $H - $fy - 50; $fw = $W - 2 * $margin
    # gradient outline: outer rounded rectangle minus a copy inset by the line width
    $lw = 1.5
    $outer = $sl.Shapes.AddShape(5, [single]$margin, [single]$fy, [single]$fw, [single]$fh); $outer.Adjustments.Item(1) = 0.06; $outer.Name = 'FrameOuter'
    $inner = $sl.Shapes.AddShape(5, [single]($margin + $lw), [single]($fy + $lw), [single]($fw - 2 * $lw), [single]($fh - 2 * $lw)); $inner.Adjustments.Item(1) = 0.06; $inner.Name = 'FrameInner'
    $frame = Merge-Shapes 4 @('FrameOuter', 'FrameInner'); $frame.Name = 'Frame'; $frame.Line.Visible = 0
    $frame.Fill.TwoColorGradient(1, 1); $frame.Fill.GradientAngle = 45
    $stops = $frame.Fill.GradientStops; $stops.Insert($acc, 0.5, 0, 2)
    foreach ($i in 1, 3) { $stops.Item($i).Color.RGB = $acc; $stops.Item($i).Transparency = 0.9 }
    $stops.Item(1).Position = 0; $stops.Item(3).Position = 1
    # cards at equal spacing
    $pad = 26; $gap = 16; $cw = ($fw - 2 * $pad - $gap * ($n - 1)) / $n; $ch = $fh - 2 * $pad
    for ($i = 0; $i -lt $n; $i++) {
        $x = $margin + $pad + ($cw + $gap) * $i; $cy = $fy + $pad
        $card = $sl.Shapes.AddShape(5, [single]$x, [single]$cy, [single]$cw, [single]$ch); $card.Adjustments.Item(1) = 0.08
        $card.Fill.ForeColor.RGB = (ConvertTo-Bgr $CardColor); $card.Line.Visible = 0; $card.Name = "Card$($i + 1)"
        $d = 46
        $dot = $sl.Shapes.AddShape(9, [single]($x + ($cw - $d) / 2), [single]($cy + 26), $d, $d); $dot.Fill.ForeColor.RGB = $acc; $dot.Line.Visible = 0
        $num = New-Text ('{0:00}' -f ($i + 1)) 15 (ConvertTo-Bgr $Background) $Font
        $num.Left = [single]($x + ($cw - [double]$num.Width) / 2); $num.Top = [single]($cy + 26 + ($d - [double]$num.Height) / 2)
        $lab = New-FitText $labels[$i] 20 $c $Font ($cw - 20); $lab.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        $lab.Left = [single]($x + ($cw - [double]$lab.Width) / 2); $lab.Top = [single]($cy + 26 + $d + 16)
        if ($i -lt $texts.Count) {
            $rule = $sl.Shapes.AddLine([single]($x + $cw * 0.3), [single]([double]$lab.Top + [double]$lab.Height + 12), [single]($x + $cw * 0.7), [single]([double]$lab.Top + [double]$lab.Height + 12))
            $rule.Line.ForeColor.RGB = $acc; $rule.Line.Weight = 0.75
            $body = New-FitText ($texts[$i] -replace ';', '|') 13 (ConvertTo-Bgr '#C8D0DC') 'Arial' ($cw - 24); $body.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $body.Left = [single]($x + ($cw - [double]$body.Width) / 2); $body.Top = [single]([double]$lab.Top + [double]$lab.Height + 24)
        }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
