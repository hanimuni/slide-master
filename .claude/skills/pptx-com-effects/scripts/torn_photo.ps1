# Torn photo cuts: the picture is ripped by hand rather than cropped - ragged edges, pieces that no longer line up.
#   Edge    one photo with a torn right edge, text beside it
#   Split   one photo torn in two across the middle, the halves pulled apart and tilted, the title in the gap
#   Strips  one photo torn into vertical strips, each shifted and tilted by a different amount
#   Scraps  up to three different photos as torn scraps, overlapping at odd angles, taped down
# By hand: draw the tear with Shapes > Freeform (many short clicks, never a straight run), select the photo first and
# the freeform second, Merge Shapes > Intersect. For pieces of one photo, copy the photo once per piece BEFORE cutting,
# so every piece keeps its place in the picture; move and tilt them only after all cuts are done.
# The white paper edge is the picture outline (white, 3 pt); the lift is a soft shadow.
# Example: .\torn_photo.ps1 -Style Split -Images C:\work\city.jpg -Title 'BEFORE / AFTER' -Out C:\work\torn.pptx -Preview C:\work\torn.png
param(
    [ValidateSet('Edge', 'Split', 'Strips', 'Scraps')][string]$Style = 'Split',
    [Parameter(Mandatory)][string[]]$Images,
    [string]$Title = '',                          # '|' = line break
    [string]$Sub = '',
    [int]$Pieces = 5,                             # Strips: how many
    [double]$Rough = 10,                          # how far the tear wanders, in points
    [double]$Gap = 90,                            # Split: distance between the halves
    [int]$Seed = 7,                               # another number = another tear
    [string]$Background = '#EFEAE0',
    [string]$Color = '#1B1B1B',
    [string]$Font = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
. (Join-Path $PSScriptRoot 'page_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent
$rng = New-Object System.Random $Seed

# ragged line from A to B: a slow drift plus a fine jitter, pinned at both ends so neighbouring pieces still meet
function Tear($ax, $ay, $bx, $by) {
    $len = [math]::Sqrt(($bx - $ax) * ($bx - $ax) + ($by - $ay) * ($by - $ay)); $n = [math]::Max(6, [int]($len / 11))
    $nx = - ($by - $ay) / $len; $ny = ($bx - $ax) / $len; $drift = 0; $pts = @()
    for ($i = 0; $i -le $n; $i++) {
        $drift = $drift * 0.7 + ($rng.NextDouble() - 0.5) * $Rough
        $o = if ($i -eq 0 -or $i -eq $n) { 0 } else { $drift + ($rng.NextDouble() - 0.5) * $Rough * 0.5 }
        $pts += , @(($ax + ($bx - $ax) * $i / $n + $nx * $o), ($ay + ($by - $ay) * $i / $n + $ny * $o))
    }
    , $pts
}
function Flip($pts) { $r = @($pts); [array]::Reverse($r); , $r }
# the photo (covering the box, never stretched) cut to a polygon; white paper edge and a soft shadow
function Cut([string]$path, $l, $t, $w, $h, $pts) {
    $script:cuts++; $pic = $sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = "Src$script:cuts"
    $pic.Width = [single]$w; if ([double]$pic.Height -lt $h) { $pic.Height = [single]$h }
    $pic.Left = [single]($l + ($w - [double]$pic.Width) / 2); $pic.Top = [single]($t + ($h - [double]$pic.Height) / 2)
    $frame = Poly $pts 0 $null 0; $frame.Name = "Cut$script:cuts"
    $piece = Merge-Shapes 3 @("Src$script:cuts", "Cut$script:cuts"); $piece.Name = "Piece$script:cuts"
    $piece.Line.Visible = -1; $piece.Line.ForeColor.RGB = 0xFFFFFF; $piece.Line.Weight = 3
    $sh = $piece.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Transparency = 0.65; $sh.Blur = 9; $sh.OffsetX = 2; $sh.OffsetY = 4
    $piece
}
function Say([string]$text, $size, $x, $y, [string]$face = $Font) { $t = New-FitText $text $size (ConvertTo-Bgr $Color) $face ($W * 0.5); $t.Left = [single]$x; $t.Top = [single]$y; $t }
$script:cuts = 0

try {
    New-Deck $Background
    $m = 30                                                           # pieces run this far past the slide, so tilting them shows no corner
    switch ($Style) {
        'Edge' {
            $xt = $W * 0.56; $rip = Tear $xt (-$m) ($xt + 34) ($H + $m)
            [void](Cut $Images[0] (-$m) (-$m) ($xt + 70 + $m) ($H + 2 * $m) (@(, @((-$m), (-$m))) + $rip + , @((-$m), ($H + $m))))
            if ($Title) { $t = Say $Title 40 ($W * 0.65) ($H * 0.3); $t.Name = 'Title' }
            if ($Sub) { [void](Say $Sub 15 ($W * 0.65) ($H * 0.62) 'Arial') }
        }
        'Split' {
            $rip = Tear (-$m) ($H * 0.42) ($W + $m) ($H * 0.58)
            $top = Cut $Images[0] (-$m) (-$m) ($W + 2 * $m) ($H + 2 * $m) (@(, @((-$m), (-$m))) + , @(($W + $m), (-$m)) + (Flip $rip))
            $low = Cut $Images[0] (-$m) (-$m) ($W + 2 * $m) ($H + 2 * $m) ($rip + , @(($W + $m), ($H + $m)) + , @((-$m), ($H + $m)))
            $top.Top = [single]([double]$top.Top - $Gap * 0.55); $top.Rotation = -1.4; $low.Top = [single]([double]$low.Top + $Gap * 0.45); $low.Rotation = 1.1
            if ($Title) { $t = New-FitText $Title 34 (ConvertTo-Bgr $Color) $Font ($W * 0.8); $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                $t.Left = [single](($W - [double]$t.Width) / 2); $t.Top = [single]($H / 2 - [double]$t.Height / 2 - 4); $t.Rotation = 5; $t.Name = 'Title' }
        }
        'Strips' {
            $l = $W * 0.07; $r = $W * 0.93; $t0 = $H * 0.09; $b0 = $H * 0.74; $step = ($r - $l) / $Pieces
            $edges = @(); $edges += , @(@($l, $t0), @($l, $b0))                 # left side straight, then one tear per boundary, right side straight
            for ($i = 1; $i -lt $Pieces; $i++) { $x = $l + $step * $i + ($rng.NextDouble() - 0.5) * $step * 0.3; $edges += , (Tear $x $t0 ($x + ($rng.NextDouble() - 0.5) * 40) $b0) }
            $edges += , @(@($r, $t0), @($r, $b0))
            $made = @()
            for ($i = 0; $i -lt $Pieces; $i++) { $made += Cut $Images[0] $l $t0 ($r - $l) ($b0 - $t0) ($edges[$i] + (Flip $edges[$i + 1])) }
            for ($i = 0; $i -lt $Pieces; $i++) {                      # every strip gets its own shift and tilt: nothing lines up any more
                $made[$i].Left = [single]([double]$made[$i].Left + ($i - ($Pieces - 1) / 2) * 9)
                $made[$i].Top = [single]([double]$made[$i].Top + ($rng.NextDouble() - 0.5) * 46); $made[$i].Rotation = [single](($rng.NextDouble() - 0.5) * 4.5)
            }
            if ($Title) { $t = Say $Title 30 $l ($H * 0.82); $t.Name = 'Title' }
            if ($Sub) { $s2 = New-Text $Sub 14 (ConvertTo-Bgr $Color) 'Arial'; $s2.Left = [single]($r - [double]$s2.Width); $s2.Top = [single]($H * 0.86) }
        }
        'Scraps' {
            $spots = @(@(0.05, 0.1, 0.31, 0.5, -6), @(0.27, 0.42, 0.31, 0.48, 5), @(0.37, 0.06, 0.27, 0.42, -2))
            for ($i = 0; $i -lt [math]::Min(3, $Images.Count); $i++) {
                $q = $spots[$i]; $l = $W * $q[0]; $t0 = $H * $q[1]; $pw = $W * $q[2]; $ph = $H * $q[3]      # not $w / $h: PowerShell names ignore case, they would overwrite the slide size
                $pts = (Tear $l $t0 ($l + $pw) $t0) + (Tear ($l + $pw) $t0 ($l + $pw) ($t0 + $ph)) + (Tear ($l + $pw) ($t0 + $ph) $l ($t0 + $ph)) + (Tear $l ($t0 + $ph) $l $t0)
                $piece = Cut $Images[$i] ($l - 14) ($t0 - 14) ($pw + 28) ($ph + 28) $pts; $piece.Rotation = [single]$q[4]
                $tape = Box 1 ($l + $pw / 2 - 38) ($t0 - 12) 76 24 (ConvertTo-Bgr '#F2E3A0'); $tape.Fill.Transparency = 0.25; $tape.Rotation = [single]($q[4] + ($rng.NextDouble() - 0.5) * 16)
            }
            if ($Title) { $t = Say $Title 38 ($W * 0.68) ($H * 0.34); $t.Name = 'Title' }
            if ($Sub) { [void](Say $Sub 15 ($W * 0.68) ($H * 0.64) 'Arial') }
        }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
