# Text looks made from three ingredients only - fill, outline, shadow - with no font change:
#   Outline            outline only (fill off)
#   SoftShadow         light text made readable by a blurred, centred shadow (distance 0, blur up)
#   HardShadow         crisp offset shadow (blur 0, distance ~5pt, angle ~50 = lower right)
#   FillOutline        two colours: fill + outline
#   OutlineShadow      outline only + crisp shadow -> looks like two thin outlines
#   FillOutlineShadow  fill + outline + crisp shadow in the outline colour
#   DoubleShadow       two stacked text boxes: the back one is the "shadow" and so can carry its own outline
#   DoubleOutline      filled+outlined text with an outline-only copy shifted on top
#   All                one sample of each on a single slide
# The two Double styles are two text boxes: an edit has to be made twice.
# Example: .\text_style.ps1 -Text 'PPTPRO' -Style All -Out C:\work\styles.pptx -Preview C:\work\styles.png
#          .\text_style.ps1 -Text 'NEW YORK' -Style SoftShadow -Image C:\work\city.jpg -Color '#FFFFFF' -Out C:\work\t.pptx
param(
    [Parameter(Mandatory)][string]$Text,          # '|' = line break
    [ValidateSet('Outline', 'SoftShadow', 'HardShadow', 'FillOutline', 'OutlineShadow', 'FillOutlineShadow', 'DoubleShadow', 'DoubleOutline', 'All')]
    [string]$Style = 'FillOutlineShadow',
    [string]$Image = '',                          # optional photo behind the text
    [string]$Color = '#D7F205',                   # fill
    [string]$LineColor = '#1B1B1B',               # outline (and crisp shadow)
    [string]$Background = '#F2F2F0',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 96,
    [double]$LineWeight = 0,                      # 0 = about 1/60 of the font size (thick outlines hurt legibility)
    [double]$Offset = 0,                          # shadow / second copy distance in pt; 0 = about 1/18 of the font size
    [double]$Angle = 50,                          # shadow direction, 50 = lower right
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Image -and -not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Look($t, [bool]$fill, [bool]$line, [string]$shadow, [double]$size) {   # $shadow: '', 'soft', 'hard'
    $f = $t.TextFrame2.TextRange.Font
    $lw = if ($LineWeight -gt 0) { $LineWeight } else { [math]::Max(0.75, $size / 60) }
    $d = if ($Offset -gt 0) { $Offset } else { $size / 18 }
    if ($fill) { $f.Fill.ForeColor.RGB = (ConvertTo-Bgr $Color) } else { $f.Fill.Visible = 0 }
    if ($line) { $f.Line.Visible = -1; $f.Line.ForeColor.RGB = (ConvertTo-Bgr $LineColor); $f.Line.Weight = [single]$lw }
    if ($shadow) {
        $s = $f.Shadow; $s.Visible = -1
        if ($shadow -eq 'soft') { $s.ForeColor.RGB = 0; $s.Blur = [single]($size / 4); $s.OffsetX = 0; $s.OffsetY = 0; $s.Transparency = 0.25 }
        else {
            $s.ForeColor.RGB = (ConvertTo-Bgr $LineColor); $s.Blur = 0; $s.Transparency = 0
            $s.OffsetX = [single]($d * [math]::Cos($Angle * [math]::PI / 180)); $s.OffsetY = [single]($d * [math]::Sin($Angle * [math]::PI / 180))
        }
    }
}
# Draws one style with its centre at ($cx,$cy); returns nothing.
function Add-Styled([string]$style, [double]$size, $cx, $cy, $maxW) {
    $t = New-FitText $Text $size 0 $Font $maxW
    $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $size = [double]$t.TextFrame.TextRange.Font.Size
    $t.Left = [single]($cx - [double]$t.Width / 2); $t.Top = [single]($cy - [double]$t.Height / 2); $t.Name = "Text_$style"
    $d = if ($Offset -gt 0) { $Offset } else { $size / 18 }
    $dx = [single]($d * [math]::Cos($Angle * [math]::PI / 180)); $dy = [single]($d * [math]::Sin($Angle * [math]::PI / 180))
    switch ($style) {
        'Outline'           { Set-Look $t $false $true '' $size }
        'SoftShadow'        { Set-Look $t $true $false 'soft' $size; if (-not $Image) { $t.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF } }
        'HardShadow'        { Set-Look $t $true $false 'hard' $size }
        'FillOutline'       { Set-Look $t $true $true '' $size }
        'OutlineShadow'     { Set-Look $t $false $true 'hard' $size }
        'FillOutlineShadow' { Set-Look $t $true $true 'hard' $size }
        'DoubleShadow'      {                    # back copy = the shadow, with an outline of its own
            Set-Look $t $true $true '' $size
            $b = $t.Duplicate(); $b.Name = "Back_$style"; $b.Left = [single]([double]$t.Left + $dx); $b.Top = [single]([double]$t.Top + $dy)
            $b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (ConvertTo-Bgr $Background); $b.ZOrder(3)      # msoSendBackward
        }
        'DoubleOutline'     {                    # outline-only copy shifted on top
            Set-Look $t $true $true '' $size
            $b = $t.Duplicate(); $b.Name = "Front_$style"; $b.Left = [single]([double]$t.Left + $dx); $b.Top = [single]([double]$t.Top + $dy)
            $b.TextFrame2.TextRange.Font.Fill.Visible = 0
        }
    }
}

try {
    New-Deck $Background
    if ($Image) { $bg = Add-Cover $Image; $bg.Name = 'Photo' }
    if ($Style -eq 'All') {
        $list = 'Outline', 'SoftShadow', 'HardShadow', 'FillOutline', 'OutlineShadow', 'FillOutlineShadow', 'DoubleShadow', 'DoubleOutline'
        for ($i = 0; $i -lt 8; $i++) {
            $cx = 240 + 480 * ($i % 2); $cy = 96 + 122 * [math]::Floor($i / 2)
            $cap = New-Text $list[$i] 11 (ConvertTo-Bgr '#D0342C') 'Arial'; $cap.Left = [single]($cx - [double]$cap.Width / 2); $cap.Top = [single]($cy - 52)
            Add-Styled $list[$i] 52 $cx ($cy + 6) 420
        }
    } else { Add-Styled $Style $FontSize ($W / 2) ($H / 2) ($W - 120) }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
