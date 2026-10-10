# Curtain intro: a coloured panel with a pointed front sweeps across the photo (optionally pulled by a cut-out
# object such as a plane seen from above), shows the lead-in words one after another, then leaves and uncovers
# the final title on the photo. Slides: photo -> panel covering with word 1 -> word 2 ... -> photo with title.
# Example: .\curtain_sweep.ps1 -Image beach.jpg -Words 'TRAVEL|TO|THAILAND' -Object plane.png -Out C:\work\intro.pptx -Preview C:\work\intro.mp4
param(
    [Parameter(Mandatory)][string]$Image,
    [Parameter(Mandatory)][string]$Words,         # 'a|b|c' - the last one is the title left on the photo
    [string]$Object = '',                         # transparent PNG at the tip of the panel, pointing right
    [string]$PanelColor = '#C2185B',
    [string]$Color = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 120,
    [double]$Duration = 1,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Image, $Object) { if ($p -and -not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Panel([double]$left) {                # the object rides on the tip
    $sl.Shapes('Panel').Left = [single]$left
    if ($Object) { $o = $sl.Shapes('Object'); $o.Left = [single]($left + $panelW - [double]$o.Width * 0.75) }
}
function Set-Word([int]$i, [double]$top) { $x = $sl.Shapes("Word$i"); $x.Left = [single](($W - [double]$x.Width) / 2); $x.Top = [single]$top }

try {
    New-Deck '#000000'
    $list = $Words -split '\|'; $lead = $list.Count - 1; $c = ConvertTo-Bgr $Color
    $photo = Add-Cover $Image; $photo.Name = 'Photo'
    $final = New-FitText $list[$lead] $FontSize $c $Font ($W - 140); $final.Name = 'FinalTitle'; Set-Center $final
    $sh = $final.TextFrame2.TextRange.Font.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Blur = 14; $sh.OffsetX = 0; $sh.OffsetY = 3; $sh.Transparency = 0.55
    $finalTop = [double]$final.Top; $final.Top = [single]($H + 40)          # arrives with the last sweep
    $tip = 170; $panelW = $W + 2 * $tip
    $panel = $sl.Shapes.AddShape(51, 0, -20, [single]$panelW, [single]($H + 40)); $panel.Name = 'Panel'    # pentagon arrow, tip to the right
    $panel.Adjustments.Item(1) = [single]($tip / ($H + 40)); $panel.Fill.ForeColor.RGB = (ConvertTo-Bgr $PanelColor); $panel.Line.Visible = 0
    if ($Object) { $o = $sl.Shapes.AddPicture($Object, 0, -1, 0, 0, -1, -1); $o.LockAspectRatio = -1; $o.Height = 300; $o.Top = [single](($H - 300) / 2); $o.Name = 'Object' }
    for ($i = 0; $i -lt $lead; $i++) { $t = New-FitText $list[$i] $FontSize $c $Font ($W - 200); $t.Name = "Word$i" }
    $mid = ($H - [double]$sl.Shapes('Word0').Height) / 2
    # 1: photo only, panel waiting on the left
    Set-Panel (-$panelW - 320); for ($i = 0; $i -lt $lead; $i++) { Set-Word $i ($H + 60) }
    # 2..: panel covers the slide, lead-in words change (previous leaves upward, next comes from below)
    for ($k = 0; $k -lt $lead; $k++) {
        Add-MorphSlide $(if ($k -eq 0) { $Duration } else { 0.5 })
        Set-Panel (-$tip)
        for ($i = 0; $i -lt $lead; $i++) { Set-Word $i $(if ($i -lt $k) { -220 } elseif ($i -eq $k) { $mid } else { $H + 60 }) }
    }
    # last: panel leaves to the right, the title is on the photo
    Add-MorphSlide $Duration
    Set-Panel ($W + 340); for ($i = 0; $i -lt $lead; $i++) { Set-Word $i (-220) }
    $sl.Shapes('FinalTitle').Top = [single]$finalTop
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
