# Neumorphism cards: tall rounded cards in exactly the background colour that seem to rise out of the page.
# A shape carries only one shadow, so each card is two identical shapes stacked: the lower one throws a dark
# shadow away from the light (down-right), the upper one a white shadow toward it (up-left).
# Same rule as neumorph_text.ps1: the page must be a mid-light grey; on white or black nothing shows.
# Example: .\neumorph_cards.ps1 -Title OBJECTIVE -Cards 'Reach;100k visits|Convert;3% sign-up|Keep;40% return' -Out C:\work\neu.pptx -Preview C:\work\neu.png
param(
    [Parameter(Mandatory)][string]$Cards,         # 'head;text|head;text|...' (two to five)
    [string]$Title = '',
    [string]$Base = '#E6E6E6',
    [string]$Accent = '#12B5A6',                  # numbers
    [string]$TitleColor = '#E85D8A',
    [string]$Font = 'Arial Black',
    [double]$Blur = 14,
    [double]$Distance = 8,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Base
    $c = ConvertTo-Bgr $Base; $list = @($Cards -split '\|' | Where-Object { $_ }); $n = $list.Count
    $y = 60
    if ($Title) { $t = New-FitText $Title 26 (ConvertTo-Bgr $TitleColor) $Font ($W - 200); $t.Left = [single](($W - [double]$t.Width) / 2); $t.Top = $y; $y += [double]$t.Height + 40 } else { $y = 90 }
    $gap = 46; $cw = [math]::Min(170, ($W - 160 - $gap * ($n - 1)) / $n); $ch = $H - $y - 70; $x0 = ($W - ($cw * $n + $gap * ($n - 1))) / 2
    for ($i = 0; $i -lt $n; $i++) {
        $x = $x0 + ($cw + $gap) * $i; $part = $list[$i] -split ';', 2
        foreach ($layer in @(@(0, 1, 0.72), @(0xFFFFFF, -1, 0.1))) {     # colour, direction (1 = down-right), transparency
            $s = $sl.Shapes.AddShape(5, [single]$x, [single]$y, [single]$cw, [single]$ch); $s.Adjustments.Item(1) = 0.5
            $s.Fill.ForeColor.RGB = $c; $s.Line.Visible = 0
            $sh = $s.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = $layer[0]; $sh.Transparency = $layer[2]; $sh.Blur = [single]$Blur
            $sh.OffsetX = [single]($Distance * $layer[1]); $sh.OffsetY = [single]($Distance * $layer[1])
        }
        $num = New-Text ("$($i + 1)") 40 (ConvertTo-Bgr $Accent) $Font; $num.Left = [single]($x + ($cw - [double]$num.Width) / 2); $num.Top = [single]($y + $ch * 0.16)
        $head = New-FitText $part[0] 18 (ConvertTo-Bgr '#2B2B2B') $Font ($cw - 30); $head.Left = [single]($x + ($cw - [double]$head.Width) / 2); $head.Top = [single]($y + $ch * 0.42)
        if ($part.Count -gt 1) {
            $b = New-FitText $part[1] 12 (ConvertTo-Bgr '#5A5A5A') 'Arial' ($cw - 30); $b.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $b.Left = [single]($x + ($cw - [double]$b.Width) / 2); $b.Top = [single]([double]$head.Top + [double]$head.Height + 10)
        }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
