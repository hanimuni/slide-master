# Negative-space lettering: one letter of a word is shown only as a hole (W[A]TER), its counter stays solid.
# Works best on letters with an inner counter (A, O, D, P, R ...).
# Example: .\negative_letter.ps1 -Text WATER -Index 2 -Out C:\work\water.pptx -Preview C:\work\water.png
param(
    [Parameter(Mandatory)][string]$Text,          # one line
    [Parameter(Mandatory)][int]$Index,            # 1-based position of the letter to hollow out
    [string]$Color = '#FFFFFF',
    [string]$Background = '#0B0B0B',
    [string]$Image = '',                          # optional full-slide picture behind the word
    [string]$Font = 'Arial Black',
    [double]$FontSize = 170,
    [double]$Pad = 0.05,                          # block overlap into the neighbours, as a share of the font size
    [double]$Cover = 1.0,                         # share of the letter height the block covers (A reads best near 0.7)
    [ValidateSet('Top', 'Bottom')][string]$Anchor = 'Top',   # which end the block starts from (Bottom for V, Y, U)
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Index -lt 1 -or $Index -gt $Text.Length) { throw "-Index must be 1..$($Text.Length)" }
if ($Image -and -not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    if ($Image) { [void](Add-Cover $Image) }
    $c = ConvertTo-Bgr $Color
    $t = New-FitText $Text $FontSize $c $Font ($W - 200)
    Set-Center $t
    $size = [double]$t.TextFrame.TextRange.Font.Size
    $ch = $t.TextFrame.TextRange.Characters($Index, 1)
    $chL = [double]$ch.BoundLeft; $chW = [double]$ch.BoundWidth

    # the letter on its own, at exactly the place it has inside the word
    $one = New-Text $Text.Substring($Index - 1, 1) $size $c $Font
    $one.Left = [single]$chL; $one.Top = [single]$t.Top
    $letter = New-GlyphShape $one; $one.Delete()

    $word = New-GlyphShape $t; $t.Delete()
    $box = Get-Box $word
    $bh = $box.H * $Cover
    $bt = $box.T + $(if ($Anchor -eq 'Top') { 0.0 } else { $box.H - $bh })
    $block = New-Rect ($chL - $size * $Pad) ([single]$bt) ($chW + 2 * $size * $Pad) $bh
    $solid = Merge-Shapes 1 @($word.Name, $block.Name)        # word + block over the letter
    $res = Merge-Shapes 4 @($solid.Name, $letter.Name)        # punch the letter out of it
    $res.Name = 'NegativeWord'
    $res.Fill.ForeColor.RGB = $c; $res.Line.Visible = 0
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
