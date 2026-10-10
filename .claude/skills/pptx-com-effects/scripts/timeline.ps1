# Vertical timeline: one line down the right half, a dot per entry at equal spacing, the year to the left of the
# dot and one short line to the right. Only the entries worth remembering are coloured - the rest stay white.
# Example: .\timeline.ps1 -Title 'Ten years|of results' -Items '2016;first product|2019;series A|2023;10M users' -Highlight 3 -Out C:\work\time.pptx -Preview C:\work\time.png
param(
    [Parameter(Mandatory)][string]$Items,         # 'year;text|year;text|...' (five to eight read best)
    [string]$Title = '',                          # '|' = line break
    [int[]]$Highlight = @(),                      # 1-based entries drawn in the accent colour
    [string]$Image = '',                          # optional transparent PNG (person, product) standing at the lower left
    [string]$Accent = '#B6FF3B',
    [string]$Color = '#FFFFFF',
    [string]$Background = '#0B0B0B',
    [string]$Font = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Image -and -not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $ink = ConvertTo-Bgr $Color; $acc = ConvertTo-Bgr $Accent; $list = @($Items -split '\|' | Where-Object { $_ }); $n = $list.Count
    if ($Image) {
        $pic = $sl.Shapes.AddPicture($Image, 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Height = [single]($H * 0.66)
        $pic.Left = 40; $pic.Top = [single]($H - [double]$pic.Height); $pic.Name = 'Figure'
    }
    if ($Title) { $t = New-FitText $Title 34 $acc $Font ($W * 0.42); $t.Left = 60; $t.Top = 54; $t.Name = 'Title' }
    $lx = $W * 0.62; $top = 50; $bottom = $H - 50; $step = ($bottom - $top) / [math]::Max(1, $n - 1)
    $line = $sl.Shapes.AddLine([single]$lx, [single]$top, [single]$lx, [single]$bottom); $line.Line.ForeColor.RGB = $ink; $line.Line.Weight = 1.5; $line.Name = 'Axis'
    for ($i = 0; $i -lt $n; $i++) {
        $part = $list[$i] -split ';', 2; $y = $top + $step * $i; $c = if ($Highlight -contains ($i + 1)) { $acc } else { $ink }
        $dot = $sl.Shapes.AddShape(9, [single]($lx - 7), [single]($y - 7), 14, 14); $dot.Fill.ForeColor.RGB = $c; $dot.Line.Visible = 0; $dot.Name = "Dot$($i + 1)"
        $yr = New-Text $part[0] 20 $c $Font; $yr.Left = [single]($lx - 26 - [double]$yr.Width); $yr.Top = [single]($y - [double]$yr.Height / 2); $yr.Name = "Year$($i + 1)"
        if ($part.Count -gt 1) { $d = New-FitText $part[1] 13 $ink 'Arial' ($W - $lx - 80); $d.Left = [single]($lx + 26); $d.Top = [single]($y - [double]$d.Height / 2); $d.Name = "Text$($i + 1)" }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
