# Picture window that opens up: the same photo is shown through a small pill-shaped window, then full bleed, then
# as the right half of a split page - Morph animates the crop. An optional cut-out object (plane, product ...)
# sits across the edge of the window so it seems to come out of the frame.
# The crop is changed with Picture Format > Crop > Crop to Shape and by dragging the crop frame; the picture
# itself is never resized, which is why Morph can glide between the three states.
# Example: .\window_morph.ps1 -Image sky.jpg -Object plane.png -Title AVIATION -Sub 'Your subtitle' -Bullets 'one|two|three' -Out C:\work\w.pptx -Preview C:\work\w.mp4
param(
    [Parameter(Mandatory)][string]$Image,
    [string]$Object = '',                         # transparent PNG that overlaps the window edge
    [string]$Title = 'TITLE',
    [string]$Sub = '',
    [string]$Bullets = '',                        # 'a|b|c' on the split page
    [string]$Kicker = 'WELCOME',                  # small word above the pill on the first slide
    [string]$Color = '#1B1B1B',
    [string]$Background = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [double]$Duration = 1,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Image, $Object) { if ($p -and -not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Window([int]$shapeType, [double]$round, $left, $top, $width, $height) {
    $p = $sl.Shapes('Window'); $p.AutoShapeType = $shapeType
    if ($shapeType -eq 5) { $p.Adjustments.Item(1) = [single]$round }
    $c = $p.PictureFormat.Crop
    $c.ShapeWidth = [single]$width; $c.ShapeHeight = [single]$height; $c.ShapeLeft = [single]$left; $c.ShapeTop = [single]$top
}
function Set-Object($cx, $cy, $width) {
    if (-not $Object) { return }
    $o = $sl.Shapes('Object'); $o.Width = [single]$width; $o.Left = [single]($cx - $width / 2); $o.Top = [single]($cy - [double]$o.Height / 2)
}
function Set-Place([string]$name, $left, $top) { $x = $sl.Shapes($name); $x.Left = [single]$left; $x.Top = [single]$top }

try {
    New-Deck $Background
    $ink = ConvertTo-Bgr $Color
    $pic = Add-Cover $Image; $pic.Name = 'Window'
    if ($Object) { $o = $sl.Shapes.AddPicture($Object, 0, -1, 0, 0, -1, -1); $o.LockAspectRatio = -1; $o.Name = 'Object' }
    # every text is made once and parked off-slide where it is not needed, so Morph moves it instead of fading
    $k = New-Text $Kicker 16 $ink 'Arial'; $k.TextFrame2.TextRange.Font.Spacing = 8; $k.Name = 'Kicker'
    $big = New-FitText $Title 110 0xFFFFFF $Font ($W - 160); $big.Name = 'BigTitle'
    $sh = $big.TextFrame2.TextRange.Font.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Blur = 18; $sh.OffsetX = 0; $sh.OffsetY = 2; $sh.Transparency = 0.6
    $st = New-FitText $Title 34 $ink $Font 330; $st.Name = 'SideTitle'
    $ss = New-FitText $(if ($Sub) { $Sub } else { ' ' }) 16 $ink 'Arial' 330; $ss.Name = 'SideSub'
    $sb = New-FitText $(if ($Bullets) { ($Bullets -split '\|' | ForEach-Object { "- $_" }) -join '|' } else { ' ' }) 14 $ink 'Arial' 330; $sb.Name = 'SideBody'
    $off = $H + 60
    # 1: pill window
    Set-Window 5 0.5 270 205 420 150; Set-Object ($W / 2) 280 560
    Set-Place 'Kicker' (($W - [double]$k.Width) / 2) 150; Set-Place 'BigTitle' (($W - [double]$big.Width) / 2) $off
    Set-Place 'SideTitle' 70 $off; Set-Place 'SideSub' 70 ($off + 60); Set-Place 'SideBody' 70 ($off + 100)
    # 2: full bleed
    Add-MorphSlide $Duration
    Set-Window 1 0 0 0 $W $H; Set-Object ($W / 2) (-260) 900
    Set-Place 'Kicker' (($W - [double]$k.Width) / 2) (-60); Set-Place 'BigTitle' (($W - [double]$big.Width) / 2) (($H - [double]$big.Height) / 2)
    # 3: split page, object across the boundary
    Add-MorphSlide $Duration
    $edge = $W * 0.46
    Set-Window 1 0 $edge 0 ($W - $edge) $H; Set-Object ($edge + 150) ($H / 2 + 40) 640
    Set-Place 'BigTitle' (($W - [double]$big.Width) / 2) (-220)
    $y = 150; Set-Place 'SideTitle' 70 $y; $y += [double]$sl.Shapes('SideTitle').Height + 8
    Set-Place 'SideSub' 70 $y; $y += [double]$sl.Shapes('SideSub').Height + 18; Set-Place 'SideBody' 70 $y
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
