# Impact punch-in: the slide sits calm, then slams towards the viewer in a fraction of a second while a jagged
# burst flashes behind the title and shards fly out from the centre, then everything settles. Three slides with
# Morph: 1 calm -> 2 hit (very short transition, advances by itself) -> 3 settled.
# With -Cutout (same frame as -Image, transparent background) the subject stays in front of the title; erase part
# of the cutout (e.g. one leg) and only the remaining part overlaps the letters.
# Example: .\impact_punch.ps1 -Image run.jpg -Cutout leg.png -Text RUN -Out C:\work\punch.pptx -Preview C:\work\punch.mp4
param(
    [Parameter(Mandatory)][string]$Image,
    [Parameter(Mandatory)][string]$Text,
    [string]$Cutout = '',
    [string]$Color = '#FFFFFF',
    [string]$Accent = '#FF2D55',                  # burst and half of the shards
    [string]$Font = 'Arial Black',
    [double]$FontSize = 210,
    [double]$Punch = 1.22,                        # how far the picture jumps forward on the hit
    [int]$Shards = 18,
    [double]$HitTime = 0.12,                      # length of the hit transition in seconds
    [int]$Seed = 3,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Image, $Cutout) { if ($p -and -not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Scaled($shape, $box, [double]$k) {    # scale about the slide centre, from the slide-1 box
    $shape.LockAspectRatio = 0
    $shape.Width = [single]($box.W * $k); $shape.Height = [single]($box.H * $k)
    $shape.Left = [single]($W / 2 - ($W / 2 - $box.L) * $k); $shape.Top = [single]($H / 2 - ($H / 2 - $box.T) * $k)
}
function Set-Shard($shape, [double]$angle, [double]$dist, [double]$len) {
    $shape.Width = [single]($len / 6); $shape.Height = [single]$len
    $rad = $angle * [math]::PI / 180
    $shape.Left = [single]($W / 2 + $dist * [math]::Cos($rad) - $len / 12); $shape.Top = [single]($H / 2 + $dist * [math]::Sin($rad) - $len / 2)
    $shape.Rotation = [single]((($angle + 90) % 360 + 360) % 360)       # the tip points away from the centre
}
function Set-Burst($shape, [double]$bw, [double]$bh, [double]$transparency) {
    $shape.Width = [single]$bw; $shape.Height = [single]$bh; $shape.Left = [single](($W - $bw) / 2); $shape.Top = [single](($H - $bh) / 2)
    $shape.Fill.Transparency = [single]$transparency
}

try {
    New-Deck '#000000'
    $rng = New-Object System.Random $Seed
    $acc = ConvertTo-Bgr $Accent
    $photo = Add-Cover $Image; $photo.Name = 'Photo'
    $burst = $sl.Shapes.AddShape(89, 0, 0, 10, 10); $burst.Name = 'Burst'; $burst.Fill.ForeColor.RGB = $acc; $burst.Line.Visible = 0     # msoShapeExplosion1
    Set-Burst $burst 60 36 1
    $t = New-FitText $Text $FontSize (ConvertTo-Bgr $Color) $Font ($W - 200); Set-Center $t
    $title = New-GlyphShape $t; $title.Name = 'Title'; $t.Delete()           # outlines, so it scales like a picture
    $title.Fill.ForeColor.RGB = (ConvertTo-Bgr $Color); $title.Line.Visible = 0
    $moving = @('Photo', 'Title')
    if ($Cutout) { $cut = Add-Cover $Cutout; $cut.Name = 'Cutout'; $moving += 'Cutout' }
    $angles = @(); $far = @()
    for ($i = 0; $i -lt $Shards; $i++) {
        $s = $sl.Shapes.AddShape(7, 0, 0, 4, 10); $s.Name = "Shard$i"; $s.Line.Visible = 0      # thin triangle
        $s.Fill.ForeColor.RGB = $(if ($i % 2) { $acc } else { 0xFFFFFF })
        $angles += (360 / $Shards * $i + $rng.Next(-8, 9)); $far += $rng.Next(190, 400)
        Set-Shard $s $angles[$i] 6 8; $s.Fill.Transparency = 1
    }
    $boxes = @{}; foreach ($n in $moving) { $boxes[$n] = Get-Box $sl.Shapes($n) }

    Add-MorphSlide $HitTime                                                   # 2: the hit
    foreach ($n in $moving) { Set-Scaled $sl.Shapes($n) $boxes[$n] $Punch }
    Set-Burst $sl.Shapes('Burst') 760 470 0.1
    for ($i = 0; $i -lt $Shards; $i++) { $s = $sl.Shapes("Shard$i"); Set-Shard $s $angles[$i] $far[$i] ($rng.Next(50, 130)); $s.Fill.Transparency = 0 }
    $sl.SlideShowTransition.AdvanceOnTime = -1; $sl.SlideShowTransition.AdvanceTime = 0.15      # goes on by itself

    Add-MorphSlide 0.45                                                       # 3: settled
    foreach ($n in $moving) { Set-Scaled $sl.Shapes($n) $boxes[$n] (1 + ($Punch - 1) * 0.3) }
    Set-Burst $sl.Shapes('Burst') 900 560 1
    for ($i = 0; $i -lt $Shards; $i++) { $s = $sl.Shapes("Shard$i"); Set-Shard $s $angles[$i] 760 60; $s.Fill.Transparency = 1 }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
