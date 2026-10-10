# Photo + background-removed copy of its subject, layered for depth. Two looks from one stack:
#   text behind the subject:  .\subject_layers.ps1 -Image p.jpg -Cutout p_cut.png -Text RUN -Out o.pptx -Preview o.png
#   blurred backdrop focus:   .\subject_layers.ps1 -Image p.jpg -Cutout p_cut.png -Blur 20 -Brightness -0.2 -CutoutScale 1.05 -Out o.pptx
# -Cutout must be the same frame as -Image (same pixel size, subject in the same place, transparent elsewhere).
param(
    [Parameter(Mandatory)][string]$Image,
    [Parameter(Mandatory)][string]$Cutout,
    [string]$Text = '',                           # sits between photo and cutout ('|' = line break)
    [string]$Color = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 230,
    [double]$TextTop = -1,                        # pt from the slide top; -1 = vertically centred
    [double]$TextLeft = -1,                       # pt from the slide left; -1 = horizontally centred
    [switch]$TextOnTop,                           # put the text in front of the subject instead of behind it
    [double]$Blur = 0,                            # blur radius of the photo, 0 = sharp
    [double]$Brightness = 0,                      # -1..1 applied to the photo (negative = darker)
    [double]$CutoutScale = 1.0,                   # >1 lifts the subject off the backdrop
    [double]$MorphZoom = 0,                       # e.g. 1.3 adds a second slide: subject grows by this, photo by a quarter of it (parallax)
    [double]$Duration = 1.5,                      # Morph seconds for -MorphZoom
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($f in $Image, $Cutout) { if (-not (Test-Path $f)) { throw "not found: $f" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck '#000000'
    $bg = Add-Cover $Image; $bg.Name = 'Photo'
    if ($Blur -gt 0) { $e = $bg.Fill.PictureEffects.Insert(2); $e.EffectParameters(1).Value = $Blur }   # msoEffectBlur
    if ($Brightness -ne 0) { $bg.PictureFormat.Brightness = 0.5 + $Brightness / 2 }
    if ($Text) {
        $t = New-FitText $Text $FontSize (ConvertTo-Bgr $Color) $Font ($W - 120)
        $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        Set-Center $t
        if ($TextTop -ge 0) { $t.Top = [single]$TextTop }
        if ($TextLeft -ge 0) { $t.TextFrame.TextRange.ParagraphFormat.Alignment = 1; $t.Left = [single]$TextLeft }
        $t.Name = 'TextBehind'
    }
    $cut = Add-Cover $Cutout; $cut.Name = 'Subject'
    if ($Text -and $TextOnTop) { $sl.Shapes('TextBehind').Name = 'TextFront'; $sl.Shapes('TextFront').ZOrder(0) }   # msoBringToFront
    if ($CutoutScale -ne 1) { $cut.Width = [single]([double]$cut.Width * $CutoutScale); Set-Center $cut }
    if ($MorphZoom -gt 1) {
        # the subject grows a lot, the photo only a little: the difference reads as the subject coming closer
        Add-MorphSlide $Duration
        foreach ($pair in @('Subject', $MorphZoom), @('Photo', (1 + ($MorphZoom - 1) / 4))) {
            $s = $sl.Shapes($pair[0]); $s.Width = [single]([double]$s.Width * $pair[1]); Set-Center $s
        }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
