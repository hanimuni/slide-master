# One photo seen through a row of rounded bars: draw the bars, unite them (Merge Shapes > Union), then cut the
# picture with the united shape (picture first -> Intersect). Diagonal = long leaning bars stacked on the right
# with the title on the left; Vertical = upright bars of different lengths across the slide with a spaced
# title running over them.
# Example: .\stripe_photo.ps1 -Image city.jpg -Style Diagonal -Title CITY -Body 'one|two' -Out C:\work\city.pptx -Preview C:\work\city.png
param(
    [Parameter(Mandatory)][string]$Image,
    [ValidateSet('Diagonal', 'Vertical')][string]$Style = 'Diagonal',
    [string]$Title = '',
    [string]$Body = '',                           # '|' = line break
    [int]$Bars = 0,                               # 0 = 5 (Diagonal) or 13 (Vertical)
    [double]$Angle = -18,                         # Diagonal: lean of the bars
    [string]$Color = '#1B1B1B',
    [string]$Background = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [int]$Seed = 5,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $rng = New-Object System.Random $Seed; $names = @()
    if ($Style -eq 'Diagonal') {
        $n = if ($Bars) { $Bars } else { 5 }; $bh = 78; $gap = 12
        for ($i = 0; $i -lt $n; $i++) {
            $len = 520 + $rng.Next(0, 160)
            $b = $sl.Shapes.AddShape(5, [single]($W - 470 + $i * 34 - ($len - 560) / 2), [single](96 + ($bh + $gap) * $i), [single]$len, [single]$bh)
            $b.Adjustments.Item(1) = 0.5; $b.Rotation = [single]((($Angle % 360) + 360) % 360); $b.Name = "Bar$i"; $names += $b.Name
        }
    } else {
        $n = if ($Bars) { $Bars } else { 13 }; $gap = 8; $bw = ($W - 160 - $gap * ($n - 1)) / $n
        for ($i = 0; $i -lt $n; $i++) {
            $len = 220 + $rng.Next(0, 230)
            $b = $sl.Shapes.AddShape(5, [single](80 + ($bw + $gap) * $i), [single](($H - $len) / 2 + $rng.Next(-30, 31)), [single]$bw, [single]$len)
            $b.Adjustments.Item(1) = 0.5; $b.Name = "Bar$i"; $names += $b.Name
        }
    }
    $mask = Merge-Shapes 1 $names; $mask.Name = 'Bars'                 # union
    $mb = Get-Box $mask
    $pic = $sl.Shapes.AddPicture($Image, 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = 'Src'
    $pic.Width = [single]$mb.W; if ([double]$pic.Height -lt $mb.H) { $pic.Height = [single]$mb.H }
    $pic.Left = [single]($mb.L + ($mb.W - [double]$pic.Width) / 2); $pic.Top = [single]($mb.T + ($mb.H - [double]$pic.Height) / 2)
    $shot = Merge-Shapes 3 @('Src', 'Bars'); $shot.Name = 'Photo'
    if ($Style -eq 'Diagonal') {
        $y = 170
        if ($Title) { $t = New-FitText $Title 64 (ConvertTo-Bgr $Color) $Font 330; $t.Left = 64; $t.Top = [single]$y; $t.Name = 'Title'; $y += [double]$t.Height + 14 }
        if ($Body) { $bd = New-FitText $Body 15 (ConvertTo-Bgr $Color) 'Arial' 300; $bd.Left = 66; $bd.Top = [single]$y; $bd.Name = 'Body' }
    } elseif ($Title) {
        $t = New-Text $Title 30 0xFFFFFF $Font; $t.TextFrame2.TextRange.Font.Spacing = 22; $t.Name = 'Title'
        if ([double]$t.Width -gt $W - 120) { $t.TextFrame2.TextRange.Font.Spacing = [single](22 * ($W - 120) / [double]$t.Width - 4) }
        Set-Center $t
        $sh = $t.TextFrame2.TextRange.Font.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Blur = 10; $sh.OffsetX = 0; $sh.OffsetY = 0; $sh.Transparency = 0.4
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
