# Colour panel with the text punched out (the photo shows through the letters); Morph slides the panel in.
# Slide 1 = photo only (panel parked off-slide), slide 2 = panel in place.
# Example: .\cutout_panel.ps1 -Image C:\work\photo.jpg -Text 'YOU|HAD ME|AT HELLO' -Out C:\work\panel.pptx -Preview C:\work\panel.mp4
param(
    [Parameter(Mandatory)][string]$Image,
    [Parameter(Mandatory)][string]$Text,          # '|' = line break
    [string]$PanelColor = '#FFC800',
    [double]$PanelWidth = 0.42,                   # share of the slide width
    [ValidateSet('Left', 'Right')][string]$Side = 'Left',
    [string]$Font = 'Impact',
    [double]$FontSize = 92,
    [double]$Duration = 1,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck '#000000'
    $bg = Add-Cover $Image; $bg.Name = 'Photo'
    $pw = $W * $PanelWidth
    $px = if ($Side -eq 'Left') { 0 } else { $W - $pw }
    $panel = New-Rect $px 0 $pw $H
    $panel.Fill.ForeColor.RGB = (ConvertTo-Bgr $PanelColor); $panel.Line.Visible = 0
    $t = New-FitText $Text $FontSize 0 $Font ($pw - 80)
    $t.Left = [single]($px + 40); $t.Top = [single](($H - [double]$t.Height) / 2)
    $g = New-GlyphShape $t; $t.Delete()
    $res = Merge-Shapes 4 @($panel.Name, $g.Name)                  # panel first: the letters become holes
    $res.Name = 'Panel'

    $first = $sl
    Add-MorphSlide $Duration
    $first.Shapes('Panel').Left = if ($Side -eq 'Left') { -$pw - 10 } else { $W + 10 }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
