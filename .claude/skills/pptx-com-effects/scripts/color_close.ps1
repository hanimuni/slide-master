# Colour panels close behind the subject: slide 1 is the photo with the title behind the cut-out subject; on slide 2
# two colour panels have come in from above and below and cover the photo, the title has grown and turned dark, and
# the subject has grown and stepped aside (Morph). Stack on both slides: photo, panels, title, cut-out.
# Example: .\color_close.ps1 -Image p.jpg -Cutout cut.png -Text 'JUMP' -Out C:\work\close.pptx -Preview C:\work\close.mp4
param(
    [Parameter(Mandatory)][string]$Image,
    [Parameter(Mandatory)][string]$Cutout,        # same frame as -Image, transparent background
    [Parameter(Mandatory)][string]$Text,
    [string]$PanelColor = '#FFD400',
    [string]$Color = '#FFFFFF',                   # title on slide 1
    [string]$ToColor = '#111111',                 # title on slide 2
    [string]$Font = 'Arial Black',
    [double]$FontSize = 200,
    [double]$Grow = 1.25,
    [double]$Shift = 150,                         # how far the subject steps to the right on slide 2
    [double]$Duration = 0.75,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Image, $Cutout) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck '#000000'
    $photo = Add-Cover $Image; $photo.Name = 'Photo'
    foreach ($side in 'Top', 'Bottom') {                               # parked outside the slide
        $p = New-Rect 0 $(if ($side -eq 'Top') { -$H / 2 - 4 } else { $H + 4 }) $W ($H / 2 + 2); $p.Name = "Panel$side"
        $p.Fill.ForeColor.RGB = (ConvertTo-Bgr $PanelColor); $p.Line.Visible = 0
    }
    $t = New-FitText $Text $FontSize (ConvertTo-Bgr $Color) $Font ($W - 160); Set-Center $t
    $title = New-GlyphShape $t; $title.Name = 'Title'; $title.Fill.ForeColor.RGB = (ConvertTo-Bgr $Color); $title.Line.Visible = 0; $t.Delete()
    $cut = Add-Cover $Cutout; $cut.Name = 'Cutout'
    $tb = Get-Box $title; $cb = Get-Box $cut
    Add-MorphSlide $Duration
    $sl.Shapes('PanelTop').Top = 0; $sl.Shapes('PanelBottom').Top = [single]($H / 2 - 2)
    $x = $sl.Shapes('Title'); $x.LockAspectRatio = 0; $x.Fill.ForeColor.RGB = (ConvertTo-Bgr $ToColor)
    $x.Width = [single]($tb.W * $Grow); $x.Height = [single]($tb.H * $Grow); $x.Left = [single](($W - $tb.W * $Grow) / 2 - $Shift / 2); $x.Top = [single](($H - $tb.H * $Grow) / 2)
    $x = $sl.Shapes('Cutout'); $x.LockAspectRatio = 0
    $x.Width = [single]($cb.W * $Grow); $x.Height = [single]($cb.H * $Grow); $x.Left = [single]($W / 2 - ($W / 2 - $cb.L) * $Grow + $Shift); $x.Top = [single]($H - ($H - $cb.T) * $Grow)
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
