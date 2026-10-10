# Glassmorphism card: a copy of the photo is cut to a rounded card (intersect), blurred, and given a
# white inner shadow, so the card looks like frosted glass over the picture.
# Example: .\glass_card.ps1 -Image C:\work\photo.jpg -Title 'GLASS' -Body 'line one|line two' -Out C:\work\glass.pptx -Preview C:\work\glass.png
param(
    [Parameter(Mandatory)][string]$Image,
    [string]$Title = '',
    [string]$Body = '',                           # '|' = line break
    [double[]]$Card = @(540, 110, 360, 320),      # left, top, width, height in pt (slide is 960x540)
    [double]$Blur = 30,
    [double]$Transparency = 0.69,                 # of the white inner shadow; lower = milkier glass
    [string]$Font = 'Arial',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Image)) { throw "not found: $Image" }
if ($Card.Count -ne 4) { throw '-Card needs left,top,width,height' }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck '#000000'
    $bg = Add-Cover $Image; $bg.Name = 'Photo'
    $copy = Add-Cover $Image; $copy.Name = 'PhotoCopy'
    $r = $sl.Shapes.AddShape(5, $Card[0], $Card[1], $Card[2], $Card[3]); $r.Name = 'CardShape'   # rounded rectangle
    $r.Adjustments.Item(1) = 0.08
    $glass = Merge-Shapes 3 @('PhotoCopy', 'CardShape')                                          # picture first = picture survives
    if ($glass.Type -ne 13) { throw "intersect did not keep the picture (type $($glass.Type))" }
    $glass.Name = 'Glass'
    $e = $glass.Fill.PictureEffects.Insert(2); $e.EffectParameters(1).Value = $Blur              # msoEffectBlur
    $s = $glass.Shadow
    $s.Visible = -1; $s.Style = 1                                                                # msoShadowStyleInnerShadow
    $s.ForeColor.RGB = 0xFFFFFF; $s.Transparency = $Transparency; $s.Blur = 9; $s.OffsetX = 0; $s.OffsetY = 0

    $x = $Card[0] + 28; $y = $Card[1] + 28; $inner = $Card[2] - 56
    if ($Title) {
        $t = New-FitText $Title 34 0xFFFFFF $Font $inner
        $t.TextFrame.TextRange.Font.Bold = -1; $t.Left = $x; $t.Top = $y
        $y += [double]$t.Height + 12
    }
    if ($Body) {
        $b = $sl.Shapes.AddTextbox(1, $x, $y, $inner, 100)
        $b.TextFrame.MarginLeft = 0; $b.TextFrame.MarginRight = 0; $b.TextFrame.WordWrap = -1
        $br = $b.TextFrame.TextRange
        $br.Text = ($Body -replace '\|', "`r"); $br.Font.Name = $Font; $br.Font.NameFarEast = $Font; $br.Font.Size = 15; $br.Font.Color.RGB = 0xFFFFFF
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
