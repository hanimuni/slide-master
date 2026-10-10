# Image slider: one slide per picture. The picture is the slide background, a dark gradient covers it, and a row
# of pill-shaped windows filled with "slide background fill" shows it undarkened. A column of round thumbnails on
# the left has the current picture enlarged; Morph moves the pills up and down and scrolls the thumbnails.
# Example: .\image_slider.ps1 -Images a.jpg,b.jpg,c.jpg -Title 'YOUR TITLE' -Body 'two lines|of text' -Out C:\work\slider.pptx -Preview C:\work\slider.mp4
param(
    [Parameter(Mandatory)][string[]]$Images,
    [string]$Title = '',
    [string]$Body = '',                           # '|' = line break
    [int]$Pills = 5,
    [string]$Color = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [double]$Duration = 1,
    [int]$Seed = 4,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck '#000000'
    $n = $Images.Count; $rng = New-Object System.Random $Seed; $c = ConvertTo-Bgr $Color
    $sl.Background.Fill.UserPicture($Images[0])
    $scrim = New-Rect 0 0 $W $H; $scrim.Line.Visible = 0; $scrim.Name = 'Scrim'      # dark on the left, clear on the right
    $scrim.Fill.TwoColorGradient(1, 1); $scrim.Fill.GradientAngle = 0
    $st = $scrim.Fill.GradientStops
    $st.Item(1).Color.RGB = 0; $st.Item(1).Position = 0; $st.Item(1).Transparency = 0.1
    $st.Item(2).Color.RGB = 0; $st.Item(2).Position = 1; $st.Item(2).Transparency = 0.75
    $pw = 78; $gap = 14; $x0 = $W - 60 - $Pills * $pw - ($Pills - 1) * $gap
    for ($i = 0; $i -lt $Pills; $i++) {
        $p = $sl.Shapes.AddShape(5, [single]($x0 + ($pw + $gap) * $i), 120, $pw, 300); $p.Adjustments.Item(1) = 0.5; $p.Name = "Pill$i"
        $p.Fill.Background(); $p.Line.Visible = 0
        $s = $p.Shadow; $s.Visible = -1; $s.ForeColor.RGB = 0; $s.Transparency = 0; $s.Blur = 30; $s.OffsetX = 0; $s.OffsetY = 0
    }
    $thumb = 62; $bigT = 124; $made = 0
    for ($i = 0; $i -lt $n; $i++) {                                   # round thumbnails, cut 1:1 from the centre
        $f = $sl.Shapes.AddShape(9, 0, 0, $thumb, $thumb); $f.Name = "TF$i"
        $pic = $sl.Shapes.AddPicture($Images[$i], 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = "TS$i"
        $pic.Height = $thumb; if ([double]$pic.Width -lt $thumb) { $pic.Width = $thumb }
        $pic.Left = [single](($thumb - [double]$pic.Width) / 2); $pic.Top = 0
        $t = Merge-Shapes 3 @("TS$i", "TF$i"); $t.Name = "Thumb$i"; $t.LockAspectRatio = -1
    }
    $tx = 290
    if ($Title) { $tt = New-FitText $Title 30 $c $Font 300; $tt.Left = [single]$tx; $tt.Top = 120; $tt.Name = 'Title' }
    if ($Body) { $bd = New-FitText $Body 14 $c 'Arial' 300; $bd.Left = [single]$tx; $bd.Top = 170; $bd.Name = 'Body' }
    for ($k = 0; $k -lt $n; $k++) {
        if ($k -gt 0) { Add-MorphSlide $Duration; $sl.Background.Fill.UserPicture($Images[$k]) }
        for ($i = 0; $i -lt $Pills; $i++) { $sl.Shapes("Pill$i").Top = [single](120 + $rng.Next(-70, 71)) }
        for ($i = 0; $i -lt $n; $i++) {                               # current thumbnail big in the middle, the others stacked around it
            $t = $sl.Shapes("Thumb$i"); $d = $i - $k
            $size = if ($d -eq 0) { $bigT } else { $thumb }
            $cy = $H / 2 + $d * ($thumb + 16) + [math]::Sign($d) * ($bigT - $thumb) / 2
            $t.Width = [single]$size; $t.Left = [single](120 - $size / 2); $t.Top = [single]($cy - $size / 2)
        }
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
