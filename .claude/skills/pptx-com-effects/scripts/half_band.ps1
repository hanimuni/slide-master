# Half-slide band flip: a solid band fills half the slide with the pictures small and close together on it and the
# headline under them; on the next slide the band has dropped to the other half, the headline has moved to the
# empty half, and the pictures have grown and spread out with a caption under each (Morph).
# Example: .\half_band.ps1 -Images a.png,b.png,c.png -Title 'Soft and comfortable' -Sub 'fruit slippers' -Captions 'Strawberry|Apple|Melon' -Out C:\work\band.pptx -Preview C:\work\band.mp4
param(
    [Parameter(Mandatory)][string[]]$Images,      # transparent PNGs work best
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Captions = '',                       # 'a|b|c' one per picture, shown on slide 2
    [string]$BandColor = '#000000',
    [string]$Background = '#F1F1F1',
    [string]$Font = 'Arial Black',
    [double]$Duration = 1,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Pic([int]$i, $cx, $cy, $height) {
    $p = $sl.Shapes("Pic$i"); $p.Height = [single]$height; $p.Left = [single]($cx - [double]$p.Width / 2); $p.Top = [single]($cy - $height / 2)
}
try {
    New-Deck $Background
    $n = $Images.Count; $caps = if ($Captions) { $Captions -split '\|' } else { @() }
    $band = New-Rect 0 0 $W ($H / 2); $band.Fill.ForeColor.RGB = (ConvertTo-Bgr $BandColor); $band.Line.Visible = 0; $band.Name = 'Band'
    for ($i = 0; $i -lt $n; $i++) { $p = $sl.Shapes.AddPicture($Images[$i], 0, -1, 0, 0, -1, -1); $p.LockAspectRatio = -1; $p.Name = "Pic$i" }
    $t = New-FitText "$Title" 22 0xFFFFFF $Font ($W - 200); $t.Name = 'Title'
    $s = New-FitText $(if ($Sub) { $Sub } else { ' ' }) 14 0xFFFFFF 'Arial' ($W - 200); $s.Name = 'Sub'
    for ($i = 0; $i -lt $caps.Count; $i++) { $c = New-FitText $caps[$i] 15 0xFFFFFF $Font 220; $c.Name = "Cap$i" }
    # 1: band on top, pictures small and bunched, headline under them (white on the band)
    for ($i = 0; $i -lt $n; $i++) { Set-Pic $i ($W / 2 + ($i - ($n - 1) / 2) * 62) 110 70 }
    $t.Left = [single](($W - [double]$t.Width) / 2); $t.Top = 165; $s.Left = [single](($W - [double]$s.Width) / 2); $s.Top = [single](165 + [double]$t.Height + 2)
    for ($i = 0; $i -lt $caps.Count; $i++) { $c = $sl.Shapes("Cap$i"); $c.Left = [single](($W - [double]$c.Width) / 2); $c.Top = [single]($H + 40) }
    # 2: band below, headline up in dark, pictures large across the boundary, captions on the band
    Add-MorphSlide $Duration
    $sl.Shapes('Band').Top = [single]($H / 2)
    $ink = ConvertTo-Bgr '#1B1B1B'
    foreach ($nm in 'Title', 'Sub') { $sl.Shapes($nm).TextFrame.TextRange.Font.Color.RGB = $ink }
    $sl.Shapes('Title').Top = 70; $sl.Shapes('Sub').Top = [single](70 + [double]$sl.Shapes('Title').Height + 2)
    $slot = ($W - 160) / $n
    for ($i = 0; $i -lt $n; $i++) {
        $cx = 80 + $slot * ($i + 0.5); Set-Pic $i $cx ($H / 2 + 10) 190
        if ($i -lt $caps.Count) { $c = $sl.Shapes("Cap$i"); $c.Left = [single]($cx - [double]$c.Width / 2); $c.Top = [single]($H / 2 + 130) }
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
