# Neo-brutalism content page: saturated flat background, thick black outlines, and hard shadows (no blur, offset
# down-right) on every block. Left: a white panel with up to three items - hollow circle + bold head, then a box
# with one snipped corner holding the detail lines. Right: one photo with the same outline and shadow.
# By hand: Shape Format > Shape Effects > Shadow Options: blur 0, transparency 0, distance 6-8pt, angle 45.
# Example: .\brutal_panel.ps1 -Title 'SITUATION' -Items 'Cost;rent up 12%;staff short|Time;two launches|Risk;one supplier' -Image p.jpg -Out C:\work\brutal.pptx -Preview C:\work\brutal.png
param(
    [Parameter(Mandatory)][string]$Title,
    [Parameter(Mandatory)][string]$Items,         # 'head;line;line|head;line|...' (up to three items)
    [string]$Image = '',                          # right-hand photo (omit for a full-width panel)
    [string]$Note = '',                           # short handwritten remark above the photo
    [string]$Background = '#1EA7FF',
    [string]$Accent = '',                         # detail boxes; default = background colour
    [string]$Font = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Image -and -not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

# thick black outline + hard offset shadow, the two marks of the style
function Set-Brutal($shape, $offset) {
    $shape.Line.Visible = -1; $shape.Line.ForeColor.RGB = 0; $shape.Line.Weight = 3
    $sh = $shape.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Transparency = 0; $sh.Blur = 0
    $sh.OffsetX = [single]$offset; $sh.OffsetY = [single]$offset
}

try {
    New-Deck $Background
    if (-not $Accent) { $Accent = $Background }
    $m = 50
    $t = New-FitText $Title 54 0xFFFFFF $Font ($W - 2 * $m); $t.Left = $m; $t.Top = 30; $t.Name = 'Title'
    $f = $t.TextFrame2.TextRange.Font
    $f.Line.Visible = -1; $f.Line.ForeColor.RGB = 0; $f.Line.Weight = 2
    $f.Shadow.Visible = -1; $f.Shadow.ForeColor.RGB = 0; $f.Shadow.Transparency = 0; $f.Shadow.Blur = 0; $f.Shadow.OffsetX = 4; $f.Shadow.OffsetY = 4

    $top = 126; $ph = $H - $top - 44
    $pw = if ($Image) { $W * 0.5 } else { $W - 2 * $m - 8 }
    $panel = New-Rect $m $top ([single]$pw) ([single]$ph); $panel.Fill.ForeColor.RGB = 0xFFFFFF; $panel.Name = 'Panel'; Set-Brutal $panel 8
    $list = @($Items -split '\|' | Select-Object -First 3); $pad = 22; $gh = ($ph - 2 * $pad) / $list.Count
    for ($i = 0; $i -lt $list.Count; $i++) {
        $part = $list[$i] -split ';'; $gy = $top + $pad + $gh * $i; $gx = $m + $pad
        $ring = $sl.Shapes.AddShape(9, [single]$gx, [single]($gy + 3), 18, 18); $ring.Fill.Visible = 0; $ring.Line.ForeColor.RGB = 0; $ring.Line.Weight = 3
        $head = New-FitText $part[0] 18 0 $Font ($pw - 2 * $pad - 30); $head.Left = [single]($gx + 28); $head.Top = [single]$gy
        if ($part.Count -lt 2) { continue }
        $by = $gy + 30; $bh = $gh - 44
        $box = $sl.Shapes.AddShape(155, [single]$gx, [single]$by, [single]($pw - 2 * $pad), [single]$bh)   # 155 = rectangle with one snipped corner
        $box.Flip(1); $box.Fill.ForeColor.RGB = (ConvertTo-Bgr $Accent); Set-Brutal $box 5; $box.Line.Weight = 2   # flipped: the cut moves to the bottom right
        $body = New-FitText ((($part[1..($part.Count - 1)]) | ForEach-Object { "- $_" }) -join '|') 13 0xFFFFFF 'Arial' ($pw - 2 * $pad - 50)
        $body.TextFrame.TextRange.Font.Bold = -1; $body.Left = [single]($gx + 16); $body.Top = [single]($by + ($bh - [double]$body.Height) / 2)
    }
    if ($Image) {
        $fx = $m + $pw + 46; $fw = $W - $fx - $m - 8
        $frame = New-Rect ([single]$fx) $top ([single]$fw) ([single]$ph); $frame.Name = 'Frame'
        $pic = $sl.Shapes.AddPicture($Image, 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = 'Src'
        $pic.Width = [single]$fw; if ([double]$pic.Height -lt $ph) { $pic.Height = [single]$ph }
        $pic.Left = [single]($fx + ($fw - [double]$pic.Width) / 2); $pic.Top = [single]($top + ($ph - [double]$pic.Height) / 2)
        $shot = Merge-Shapes 3 @('Src', 'Frame'); $shot.Name = 'Photo'; Set-Brutal $shot 8
        if ($Note) { $n = New-FitText $Note 15 0 'Segoe Print' $fw; $n.Left = [single]$fx; $n.Top = [single]($top - [double]$n.Height - 8) }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
