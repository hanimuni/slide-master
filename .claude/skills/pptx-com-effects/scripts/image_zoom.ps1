# One photo, several slides: each slide zooms the same picture onto another spot, Morph flies between them.
# Example: .\image_zoom.ps1 -Image C:\work\city.jpg -Title 'NEW YORK' -Stops '0.72,0.5,2.4,SUN','0.2,0.85,3,CITY' -Out C:\work\zoom.pptx -Preview C:\work\zoom.mp4
param(
    [Parameter(Mandatory)][string]$Image,
    [Parameter(Mandatory)][string[]]$Stops,       # each 'fx,fy,zoom,label': focus point 0..1 inside the photo, zoom factor, caption
    [string]$Title = '',                          # caption of the opening full view
    [string]$Color = '#00E5FF',
    [string]$Font = 'Impact',
    [double]$FontSize = 60,
    [double]$Duration = 1,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent
$inv = [Globalization.CultureInfo]::InvariantCulture

function Set-Label([string]$text) {
    foreach ($n in @($sl.Shapes | Where-Object { $_.Name -eq 'Label' } | ForEach-Object { $_.Name })) { $sl.Shapes($n).Delete() }
    if (-not $text) { return }
    $t = New-FitText $text $FontSize (ConvertTo-Bgr $Color) $Font ($W - 200)
    $t.TextFrame.TextRange.Font.Italic = -1
    Set-Center $t; $t.Name = 'Label'
}

try {
    New-Deck '#000000'
    $p = Add-Cover $Image; $p.Name = 'Photo'
    $baseW = [double]$p.Width
    Set-Label $Title
    foreach ($stop in $Stops) {
        $f = $stop.Split(',', 4)
        if ($f.Count -lt 3) { throw "bad stop '$stop' (want fx,fy,zoom[,label])" }
        $fx = [double]::Parse($f[0], $inv); $fy = [double]::Parse($f[1], $inv); $zoom = [double]::Parse($f[2], $inv)
        Add-MorphSlide $Duration
        $p = $sl.Shapes('Photo')
        $p.Width = [single]($baseW * $zoom)
        $pw = [double]$p.Width; $ph = [double]$p.Height
        # focus point to the slide centre, but never show the area outside the photo
        $p.Left = [math]::Max($W - $pw, [math]::Min(0, $W / 2 - $fx * $pw))
        $p.Top = [math]::Max($H - $ph, [math]::Min(0, $H / 2 - $fy * $ph))
        Set-Label $(if ($f.Count -eq 4) { $f[3] } else { '' })
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
