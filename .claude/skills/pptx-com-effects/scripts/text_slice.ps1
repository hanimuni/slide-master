# Sliced text: the word is cut in two along a slanted line; a Morph transition slides the halves apart.
# Slide 1 = whole word, slide 2 = split.
# Example: .\text_slice.ps1 -Text SALE -Out C:\work\slice.pptx -Preview C:\work\slice.mp4
param(
    [Parameter(Mandatory)][string]$Text,
    [string]$Color = '#00E000',
    [string]$Background = '#0B0B0B',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 190,
    [double]$Angle = -12,                         # degrees, clockwise; negative = cut rises to the right
    [double]$Shift = 26,                          # pt each half travels along the cut
    [double]$Gap = 5,                             # pt each half moves away from the cut
    [double]$Duration = 0.3,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $c = ConvertTo-Bgr $Color
    $t = New-FitText $Text $FontSize $c $Font ($W - 240)
    Set-Center $t
    $g1 = New-GlyphShape $t; $t.Delete()
    $g1.Fill.ForeColor.RGB = $c; $g1.Line.Visible = 0
    $box = Get-Box $g1
    $g2 = $g1.Duplicate(); $g2.Left = [single]$box.L; $g2.Top = [single]$box.T

    # a big square whose lower edge runs through the centre of the word at -Angle
    $rad = $Angle * [math]::PI / 180; $cos = [math]::Cos($rad); $sin = [math]::Sin($rad)
    $cx = $box.L + $box.W / 2; $cy = $box.T + $box.H / 2; $side = 2 * $W; $half = $side / 2
    $names = foreach ($i in 1, 2) {
        $r = New-Rect ($cx + $half * $sin - $half) ($cy - $half * $cos - $half) $side $side
        $r.Rotation = $Angle; $r.Name
    }
    $top = Merge-Shapes 3 @($g1.Name, $names[0]); $top.Name = 'SliceTop'
    $bottom = Merge-Shapes 4 @($g2.Name, $names[1]); $bottom.Name = 'SliceBottom'

    Add-MorphSlide $Duration
    $dx = $Shift * $cos + $Gap * $sin; $dy = $Shift * $sin - $Gap * $cos
    $a = $sl.Shapes('SliceTop'); $a.Left = [double]$a.Left + $dx; $a.Top = [double]$a.Top + $dy
    $b = $sl.Shapes('SliceBottom'); $b.Left = [double]$b.Left - $dx; $b.Top = [double]$b.Top - $dy
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
