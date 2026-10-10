# Circle reveal: a light cover turns into a dark content page as a small dot grows until it fills the slide.
# Two slides + Morph. Slide 1 = cover title and a tiny dark circle; slide 2 = the same circle, now larger than the
# slide, with the content on top. Shapes that exist only on slide 2 fade in while the circle grows; the cover title
# is moved off the slide so it flies out instead of just vanishing.
# Example: .\circle_reveal.ps1 -Title 'CURIOUS|DESIGNER' -Heading 'How I work' -Items 'Ask first;why before how|Draw fast;three rough options' -Out C:\work\reveal.pptx -Preview C:\work\reveal.mp4
param(
    [Parameter(Mandatory)][string]$Title,         # cover; '|' = line break
    [string]$Heading = '',
    [string]$Items = '',                          # 'head;text|head;text|...' (up to four, laid out in two columns)
    [double[]]$Spot = @(0.5, 0.5),                # where the dot sits on the cover, as a fraction of width and height
    [string]$Cover = '#ECECEC',
    [string]$Dark = '#111111',
    [string]$Accent = '#B6FF3B',
    [string]$Font = 'Arial Black',
    [double]$Duration = 0.6,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Cover
    $ink = ConvertTo-Bgr $Dark; $cx = $W * $Spot[0]; $cy = $H * $Spot[1]
    $t = New-FitText $Title 96 $ink $Font ($W - 160); $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2; Set-Center $t; $t.Name = 'Cover'
    $disc = $sl.Shapes.AddShape(9, [single]($cx - 5), [single]($cy - 5), 10, 10); $disc.Fill.ForeColor.RGB = $ink; $disc.Line.Visible = 0; $disc.Name = 'Disc'

    Add-MorphSlide $Duration
    $d = 2.4 * $W                                                    # larger than the slide diagonal from any spot
    $disc = $sl.Shapes('Disc'); $disc.Width = [single]$d; $disc.Height = [single]$d; $disc.Left = [single]($cx - $d / 2); $disc.Top = [single]($cy - $d / 2)
    $sl.Shapes('Cover').Top = [single](-[double]$sl.Shapes('Cover').Height - 40)
    $m = 70; $y = 70
    if ($Heading) { $h2 = New-FitText $Heading 34 (ConvertTo-Bgr $Accent) $Font ($W - 2 * $m); $h2.Left = [single]$m; $h2.Top = [single]$y; $y += [double]$h2.Height + 34 }
    $list = @($Items -split '\|' | Where-Object { $_ } | Select-Object -First 4); $colW = ($W - 2 * $m - 40) / 2; $rowH = ($H - $y - 50) / 2
    for ($i = 0; $i -lt $list.Count; $i++) {
        $part = $list[$i] -split ';', 2; $x = $m + ($colW + 40) * ($i % 2); $iy = $y + $rowH * [math]::Floor($i / 2)
        $head = New-FitText $part[0] 24 0xFFFFFF $Font $colW; $head.Left = [single]$x; $head.Top = [single]$iy
        if ($part.Count -gt 1) { $b = New-FitText $part[1] 14 (ConvertTo-Bgr '#BDBDBD') 'Arial' $colW; $b.Left = [single]$x; $b.Top = [single]($iy + [double]$head.Height + 8) }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
