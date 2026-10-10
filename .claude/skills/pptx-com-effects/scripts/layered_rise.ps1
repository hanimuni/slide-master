# "Living picture": one photo cut into layers (far skyline, near buildings ...) that rise into place one after
# another, with the title sitting between two layers. Slide 1 holds every layer below the slide, each a bit
# further down than the one before; slide 2 is the finished picture with a Morph transition. All layers
# travel for the same time, so the ones that start further away move faster and seem to arrive in turn.
# Layers are transparent PNGs in the same frame as the slide (PowerPoint: duplicate the photo, Picture Format >
# Remove Background, mark what to keep - once per layer. This script does not cut them.)
# Example: .\layered_rise.ps1 -Background sky.jpg -Layers far.png,near.png -Text TOKYO -TextAfter 1 -Out C:\work\rise.pptx -Preview C:\work\rise.mp4
param(
    [Parameter(Mandatory)][string]$Background,    # stays put (sky)
    [Parameter(Mandatory)][string[]]$Layers,      # back to front, transparent PNGs covering the whole slide
    [string]$Text = '',                           # '|' = line break
    [int]$TextAfter = 0,                          # the text goes in front of this many layers (0 = right above the background)
    [double]$TextTop = 90,
    [string]$Color = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 150,
    [double]$Stagger = 160,                       # extra start distance per layer, in pt
    [double]$Duration = 1.5,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in @($Background) + $Layers) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck '#000000'
    $bg = Add-Cover $Background; $bg.Name = 'Sky'
    $moving = @()                                                   # names in stacking order, back to front
    for ($i = 0; $i -le $Layers.Count; $i++) {
        if ($Text -and $i -eq $TextAfter) {
            $t = New-FitText $Text $FontSize (ConvertTo-Bgr $Color) $Font ($W - 80)
            $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $t.Left = [single](($W - [double]$t.Width) / 2); $t.Top = [single]$TextTop; $t.Name = 'Title'; $moving += 'Title'
        }
        if ($i -lt $Layers.Count) { $l = Add-Cover $Layers[$i]; $l.Name = "Layer$($i + 1)"; $moving += $l.Name }
    }
    Add-MorphSlide $Duration                                         # slide 2 = finished picture
    $first = $pres.Slides(1)
    for ($k = 0; $k -lt $moving.Count; $k++) {                       # slide 1: everything below the slide, staggered
        $s = $first.Shapes($moving[$k]); $s.Top = [single]([double]$s.Top + $H + $Stagger * $k)
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
