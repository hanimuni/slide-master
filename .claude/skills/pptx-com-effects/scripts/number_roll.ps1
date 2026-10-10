# Rolling number (2024 -> 2025): the changing part is a two-line text (old on line 1, new on line 2).
# Slide 2 moves it up one line with Morph; background-coloured blocks hide whatever is outside the line.
# Example: .\number_roll.ps1 -Prefix 202 -From 4 -To 5 -Out C:\work\year.pptx -Preview C:\work\year.mp4
param(
    [string]$Prefix = '',                         # the part that stays
    [Parameter(Mandatory)][string]$From,
    [Parameter(Mandatory)][string]$To,
    [string]$Color = '#FFFFFF',
    [string]$ToColor = '',                        # colour of the changing part on slide 2 ('' = unchanged)
    [string]$Background = '#000000',
    [string]$Font = 'Arial Black',
    [double]$FontSize = 170,
    [double]$Duration = 1,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $c = ConvertTo-Bgr $Color
    $line = New-Text ($Prefix + $From) $FontSize $c $Font       # only to measure the full width
    $fullW = [double]$line.Width; $lineH = [double]$line.Height; $line.Delete()
    $left = ($W - $fullW) / 2; $top = ($H - $lineH) / 2
    $x = $left
    if ($Prefix) { $p = New-Text $Prefix $FontSize $c $Font; $p.Left = [single]$left; $p.Top = [single]$top; $p.Name = 'Prefix'; $x = $left + [double]$p.Width }
    $roll = New-Text "$From`r$To" $FontSize $c $Font
    $roll.Left = [single]$x; $roll.Top = [single]$top; $roll.Name = 'Roll'
    $step = [double]$roll.Height / 2                             # one line of the two-line box
    $rw = [double]$roll.Width + 20
    foreach ($m in @('MaskTop', 0, $top), @('MaskBottom', ($top + $step), ($H - $top - $step))) {
        $r = New-Rect ([single]($x - 10)) ([single]$m[1]) ([single]$rw) ([single]$m[2])
        $r.Fill.ForeColor.RGB = (ConvertTo-Bgr $Background); $r.Line.Visible = 0; $r.Name = $m[0]
    }
    Add-MorphSlide $Duration
    $r2 = $sl.Shapes('Roll'); $r2.Top = [single]($top - $step)
    if ($ToColor) { $r2.TextFrame.TextRange.Font.Color.RGB = (ConvertTo-Bgr $ToColor) }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
