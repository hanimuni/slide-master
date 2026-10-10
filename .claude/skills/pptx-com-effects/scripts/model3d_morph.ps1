# 3D model that turns and travels between two slides (Morph interpolates position, size and rotation).
# Example: .\model3d_morph.ps1 -Model C:\work\phone.glb -Title 'NEW|PHONE' -Out C:\work\3d.pptx -Preview C:\work\3d.mp4
param(
    [Parameter(Mandatory)][string]$Model,         # .glb (also .fbx/.obj/.stl/.3mf)
    [string]$Title = '',                          # '|' = line break
    [string]$Color = '#FFFFFF',
    [string]$Background = '#0B0B0F',
    [string]$Font = 'Arial Black',
    [double[]]$From = @(110, 150, 240, 20, 30),   # left, top, size, rotationX, rotationY on slide 1
    [double[]]$To = @(520, 70, 400, -10, 205),    # the same on slide 2
    [double]$Duration = 0.5,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Model)) { throw "not found: $Model" }
if ($From.Count -ne 5 -or $To.Count -ne 5) { throw '-From / -To need left,top,size,rotationX,rotationY' }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Pose($m, [double[]]$pose) {
    $m.LockAspectRatio = -1
    $m.Height = $pose[2]; $m.Left = $pose[0]; $m.Top = $pose[1]
    $m.Model3D.RotationX = $pose[3]; $m.Model3D.RotationY = $pose[4]
}

try {
    New-Deck $Background
    $m = $sl.Shapes.Add3DModel($Model, 0, -1, 0, 0, -1, -1); $m.Name = 'Model'
    Set-Pose $m $From
    if ($Title) {
        $t = New-FitText $Title 64 (ConvertTo-Bgr $Color) $Font 420
        $t.Left = $W - [double]$t.Width - 90; $t.Top = ($H - [double]$t.Height) / 2; $t.Name = 'Title'
    }
    Add-MorphSlide $Duration
    Set-Pose $sl.Shapes('Model') $To
    if ($Title) { $sl.Shapes('Title').Left = 80 }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
