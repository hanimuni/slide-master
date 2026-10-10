# Neumorphism text: two identical texts in the background colour, stacked. One carries a white shadow
# toward the light, the other a dark shadow away from it, so the word looks pressed out of the surface.
# Example: .\neumorph_text.ps1 -Text PPTPRO -Out C:\work\neu.pptx -Preview C:\work\neu.png
param(
    [Parameter(Mandatory)][string]$Text,          # '|' = line break
    [string]$Base = '#E4E4E4',                    # background = text colour; a mid-light grey works, pure white or black does not
    [string]$Font = 'Arial Black',
    [double]$FontSize = 130,
    [double]$Blur = 10,
    [double]$Distance = 5,
    [double]$LightAngle = 230,                    # degrees clockwise from the right; 230 = light from the upper left
    [double]$LightTransparency = 0.2,
    [double]$DarkTransparency = 0.7,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Shadow($t, [int]$rgb, [double]$angle, [double]$transparency) {
    $s = $t.TextFrame2.TextRange.Font.Shadow
    $rad = $angle * [math]::PI / 180
    $s.Visible = -1; $s.ForeColor.RGB = $rgb; $s.Transparency = $transparency; $s.Blur = $Blur; $s.Size = 100
    $s.OffsetX = [single]($Distance * [math]::Cos($rad)); $s.OffsetY = [single]($Distance * [math]::Sin($rad))
}

try {
    New-Deck $Base
    $c = ConvertTo-Bgr $Base
    foreach ($layer in 'Dark', 'Light') {
        $t = New-FitText $Text $FontSize $c $Font ($W - 200)
        $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        Set-Center $t; $t.Name = "Text$layer"
        if ($layer -eq 'Light') { Set-Shadow $t 0xFFFFFF $LightAngle $LightTransparency }
        else { Set-Shadow $t 0 ($LightAngle - 180) $DarkTransparency }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
