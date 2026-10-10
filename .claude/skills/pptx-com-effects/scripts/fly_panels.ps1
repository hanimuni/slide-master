# Card panels that fly in one after another: a dark lead card (title, text, date, a slider) and light cards
# beside it, each with an icon, a heading, a short text and progress bars. Every card is one group with a soft
# side shadow (Format Shape > Shadow: transparency 60 %, blur 4 pt, angle 0, distance 3 pt), and the groups get
# Animations > Fly In > Effect Options > From Left, each starting after the one before.
# Example: .\fly_panels.ps1 -Title 'MARKETING|OBJECTIVES' -Body 'short text' -Date 'May 17, 2024' -Lead 88 -Items 'USA;text;Reach:83;Sales:42|EU;text;Reach:64;Sales:51' -Out C:\work\panels.pptx -Preview C:\work\panels.mp4
param(
    [Parameter(Mandatory)][string]$Title,         # lead card; '|' = line break
    [Parameter(Mandatory)][string]$Items,         # 'heading;text;label:percent;label:percent|...' (two to four cards)
    [string]$Body = '',
    [string]$Date = '',
    [int]$Lead = -1,                              # slider value on the lead card (0-100); -1 = none
    [string[]]$Images = @(),                      # one icon (transparent PNG / SVG) per card, in order
    [string]$Dark = '#0B2239',
    [string]$Card = '#F1F1F1',
    [string]$Background = '#E9ECEF',
    [string]$Font = 'Arial',
    [double]$Duration = 0.5,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent
$script:names = @()

function Part($shape, [string]$name) { $shape.Name = $name; $script:names += $name; $shape }
function Bar($x, $y, $w, [int]$pct, [int]$fill, [int]$track, [string]$tag) {
    [void](Part (New-Rect ([single]$x) ([single]$y) ([single]$w) 5) "${tag}Track"); $sl.Shapes("${tag}Track").Fill.ForeColor.RGB = $track; $sl.Shapes("${tag}Track").Line.Visible = 0
    [void](Part (New-Rect ([single]$x) ([single]$y) ([single]($w * $pct / 100)) 5) "${tag}Fill"); $sl.Shapes("${tag}Fill").Fill.ForeColor.RGB = $fill; $sl.Shapes("${tag}Fill").Line.Visible = 0
}
function Close-Group([string]$name) {              # group what was added since the last call, add the shadow, return the group
    $g = $sl.Shapes.Range($script:names).Group(); $g.Name = $name; $script:names = @(); $g
}

try {
    New-Deck $Background
    $ink = ConvertTo-Bgr $Dark; $paper = ConvertTo-Bgr $Card; $grey = ConvertTo-Bgr '#B8BDC4'
    $list = @($Items -split '\|' | Where-Object { $_ } | ForEach-Object { , @($_ -split ';') }); $n = $list.Count
    $m = 50; $gap = 26; $top = 60; $ch = $H - 2 * $top; $cw = ($W - 2 * $m - $gap * $n) / ($n + 1); $groups = @()

    $x = $m; $pad = 22                                               # lead card
    $base = Part (New-Rect ([single]$x) $top ([single]$cw) ([single]$ch)) 'LeadBase'; $base.Fill.ForeColor.RGB = $ink; $base.Line.Visible = 0
    $base.Shadow.Visible = -1; $base.Shadow.Transparency = 0.6; $base.Shadow.Blur = 4; $base.Shadow.OffsetX = 3; $base.Shadow.OffsetY = 0
    $t = Part (New-FitText $Title 24 0xFFFFFF $Font ($cw - 2 * $pad)) 'LeadTitle'; $t.Left = [single]($x + $pad); $t.Top = [single]($top + 40); $y = [double]$t.Top + [double]$t.Height + 12
    $rule = Part (New-Rect ([single]($x + $pad)) ([single]$y) 60 3) 'LeadRule'; $rule.Fill.ForeColor.RGB = 0xFFFFFF; $rule.Line.Visible = 0; $y += 20
    if ($Body) { $b = Part (New-FitText $Body 11 0xFFFFFF $Font ($cw - 2 * $pad)) 'LeadBody'; $b.Left = [single]($x + $pad); $b.Top = [single]$y }
    if ($Date) { $d = Part (New-Text $Date 12 0xFFFFFF $Font) 'LeadDate'; $d.Left = [single]($x + $pad); $d.Top = [single]($top + $ch * 0.62) }
    if ($Lead -ge 0) {
        $by = $top + $ch * 0.8; $bw = $cw - 2 * $pad
        $pc = Part (New-Text "$Lead%" 12 0xFFFFFF $Font) 'LeadPct'; $pc.Left = [single]($x + $pad + $bw * $Lead / 100 - [double]$pc.Width / 2); $pc.Top = [single]($by - 24)
        Bar ($x + $pad) $by $bw $Lead 0xFFFFFF (ConvertTo-Bgr '#4A5D73') 'Lead'
        $knob = Part ($sl.Shapes.AddShape(9, [single]($x + $pad + $bw * $Lead / 100 - 5), [single]($by - 3), 10, 10)) 'LeadKnob'; $knob.Fill.ForeColor.RGB = 0xFFFFFF; $knob.Line.Visible = 0
    }
    $groups += Close-Group 'Lead'

    for ($i = 0; $i -lt $n; $i++) {                                  # light cards
        $x = $m + ($cw + $gap) * ($i + 1); $it = $list[$i]; $k = "C$($i + 1)"
        $base = Part (New-Rect ([single]$x) $top ([single]$cw) ([single]$ch)) "${k}Base"; $base.Fill.ForeColor.RGB = $paper; $base.Line.Visible = 0
        $base.Shadow.Visible = -1; $base.Shadow.Transparency = 0.6; $base.Shadow.Blur = 4; $base.Shadow.OffsetX = 3; $base.Shadow.OffsetY = 0
        $y = $top + 34
        if ($i -lt $Images.Count) {
            $ic = $sl.Shapes.AddPicture($Images[$i], 0, -1, 0, 0, -1, -1); $ic.LockAspectRatio = -1; $ic.Height = [single]($ch * 0.24)
            if ([double]$ic.Width -gt $cw - 2 * $pad) { $ic.Width = [single]($cw - 2 * $pad) }
            $ic.Left = [single]($x + ($cw - [double]$ic.Width) / 2); $ic.Top = [single]($y + ($ch * 0.24 - [double]$ic.Height) / 2); [void](Part $ic "${k}Icon")
        }
        $y += $ch * 0.24 + 22
        $hd = Part (New-FitText $it[0] 13 $ink $Font ($cw - 2 * $pad)) "${k}Head"; $hd.Left = [single]($x + $pad); $hd.Top = [single]$y; $y += [double]$hd.Height + 8
        if ($it.Count -gt 1 -and $it[1]) { $tx = Part (New-FitText $it[1] 10 $ink $Font ($cw - 2 * $pad)) "${k}Text"; $tx.Left = [single]($x + $pad); $tx.Top = [single]$y }
        $by = $top + $ch * 0.68
        for ($j = 2; $j -lt [math]::Min(4, $it.Count); $j++) {      # up to two progress bars: 'label:percent'
            $kv = $it[$j] -split ':'; $pct = [int]$kv[-1]; $bw = $cw - 2 * $pad
            $pt = Part (New-Text "$pct%" 10 $ink $Font) "${k}Pct$j"; $pt.Left = [single]($x + $pad + $bw * $pct / 100 - [double]$pt.Width); $pt.Top = [single]($by - 16)
            Bar ($x + $pad) $by $bw $pct $ink $grey "${k}B$j"
            $lb = Part (New-FitText $kv[0] 9 $ink $Font $bw) "${k}Lab$j"; $lb.Left = [single]($x + $pad); $lb.Top = [single]($by + 9)
            $by += 46
        }
        $groups += Close-Group $k
    }

    foreach ($g in $groups) {                                        # Fly In, From Left, one after another
        $fx = $sl.TimeLine.MainSequence.AddEffect($g, 2, 0, 3)       # 2 = Fly, 3 = after previous
        $fx.EffectParameters.Direction = 4; $fx.Timing.Duration = [single]$Duration      # 4 = from the left
    }
    Save-Deck $Out $Preview 5
    "saved $Out"
} catch { Close-Deck; throw }
