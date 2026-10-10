# Long-shadow title: text 3D format with an oblique camera, flat light, matte material and a huge depth.
# Example: .\long_shadow.ps1 -Text 'PPT|PROJECT' -Out C:\work\cover.pptx -Preview C:\work\cover.png
param(
    [Parameter(Mandatory)][string]$Text,          # '|' = line break
    [string]$Color = '#8BFF2E',
    [string]$ShadowColor = '#000000',
    [string]$Background = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 110,
    [ValidateSet('BottomRight', 'BottomLeft', 'TopRight', 'TopLeft')][string]$Direction = 'BottomRight',
    [double]$Depth = 1584,                        # pt; the UI maximum, long enough to leave the slide
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent
$camera = @{ TopLeft = 40; TopRight = 42; BottomLeft = 45; BottomRight = 47 }[$Direction]    # msoCameraOblique*

try {
    New-Deck $Background
    $t = New-FitText $Text $FontSize (ConvertTo-Bgr $Color) $Font ($W - 240)
    $t.TextFrame.TextRange.Font.Bold = -1
    $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    Set-Center $t
    $d = $t.TextFrame2.ThreeD
    # Order matters: a fresh text 3D starts in "legacy" mode, where the light rig is stored as legacyFlat
    # and resets the camera. Camera -> light -> camera again leaves both as asked (checked in the saved XML).
    $d.PresetMaterial = 5                         # msoMaterialMatte2
    $d.ExtrusionColor.RGB = (ConvertTo-Bgr $ShadowColor)
    $d.Depth = $Depth
    $d.SetPresetCamera($camera)
    $d.PresetLighting = 24                        # msoLightRigFlat
    $d.SetPresetCamera($camera)
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
