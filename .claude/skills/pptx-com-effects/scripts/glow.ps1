# Glow by shadow: a centred, blurred shadow in the object's own colour makes text or a picture look lit.
# Example: .\glow.ps1 -Text 'PPTPRO' -Out C:\work\glow.pptx -Preview C:\work\glow.png
#          .\glow.ps1 -Image C:\work\cover.jpg -Color '#FF4FA0' -Out C:\work\glow.pptx
param(
    [string]$Text = '',                           # glow text ('|' = line break) ...
    [string]$Image = '',                          # ... or a picture (give exactly one)
    [string]$Color = '#00E000',                   # text colour = glow colour; for a picture, its dominant colour
    [string]$Background = '#000000',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 120,
    [double]$Blur = 30,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (($Text -ne '') -eq ($Image -ne '')) { throw 'give exactly one of -Text or -Image' }
if ($Image -and -not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $c = ConvertTo-Bgr $Color
    if ($Text) {
        $t = New-FitText $Text $FontSize $c $Font ($W - 200)
        $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        Set-Center $t
        $s = $t.TextFrame2.TextRange.Font.Shadow
    } else {
        $p = $sl.Shapes.AddPicture($Image, 0, -1, 0, 0, -1, -1); $p.LockAspectRatio = -1
        $p.Height = $H - 200
        if ([double]$p.Width -gt $W - 200) { $p.Width = $W - 200 }
        Set-Center $p
        $s = $p.Shadow
    }
    $s.Visible = -1; $s.ForeColor.RGB = $c; $s.Transparency = 0
    $s.Blur = $Blur; $s.OffsetX = 0; $s.OffsetY = 0; $s.Size = 102
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
