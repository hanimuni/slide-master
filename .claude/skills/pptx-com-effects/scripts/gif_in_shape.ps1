# Recipe 2: an animated GIF playing inside a logo (SVG) or inside text.
# Example: .\gif_in_shape.ps1 -Svg C:\work\logo.svg -Gif C:\work\pattern.gif -Copy 'IMPOSSIBLE IS NOTHING' -Out C:\work\logo.pptx -Preview C:\work\logo.mp4
param(
    [string]$Svg = '',                   # clipping shape from an SVG file ...
    [string]$Text = '',                  # ... or from text (give exactly one of the two)
    [Parameter(Mandatory)][string]$Gif,
    [string]$Copy = '',                  # optional line under the shape
    [string]$Notes = '',                 # credits / licence lines for the speaker notes
    [string]$Outline = '#F5F06E',        # '' = no outline
    [string]$Background = '#000000',
    [string]$Grid = '#1E1E1E',           # '' = no grid
    [string]$Font = 'Arial Black',
    [double]$ShapeHeight = 270,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (($Svg -ne '') -eq ($Text -ne '')) { throw 'give exactly one of -Svg or -Text' }
foreach ($f in $Svg, $Gif) { if ($f -and -not (Test-Path $f)) { throw "not found: $f" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    if ($Grid) { Add-Grid $Grid }
    $areaH = if ($Copy) { $H - 90 } else { $H }          # vertical room for the shape

    # 1. clipping shape
    if ($Svg) {
        $existing = @($sl.Shapes | ForEach-Object { $_.Name })
        $pic = $sl.Shapes.AddPicture($Svg, 0, -1, 0, 0, -1, -1)
        $pic.LockAspectRatio = -1; $pic.Height = $ShapeHeight
        if ([double]$pic.Width -gt $W - 160) { $pic.Width = $W - 160 }
        $pic.Left = ($W - [double]$pic.Width) / 2; $pic.Top = ($areaH - [double]$pic.Height) / 2
        # "Convert to Shape" has no object-model method: select the graphic and fire the ribbon command.
        $pres.Windows(1).Activate(); $pic.Select()
        $app.CommandBars.ExecuteMso('SVGEdit')
        Start-Sleep -Seconds 2
        while ($true) {                                  # flatten nested groups down to freeforms
            $g = @($sl.Shapes | Where-Object { $_.Type -eq 6 })
            if ($g.Count -eq 0) { break }
            $g[0].Ungroup() | Out-Null
        }
        $parts = @($sl.Shapes | Where-Object { $existing -notcontains $_.Name -and $_.Type -ne 13 } | ForEach-Object { $_.Name })
        if ($parts.Count -eq 0) { throw 'SVG was not converted to shapes' }
        $shape = if ($parts.Count -gt 1) { Merge-Shapes 1 $parts } else { $sl.Shapes($parts[0]) }
    } else {
        $t = New-Text $Text 150 0xFFFFFF $Font
        if ([double]$t.Width -gt $W - 120) { $t.TextFrame.TextRange.Font.Size = 150 * ($W - 120) / [double]$t.Width }
        $t.Left = ($W - [double]$t.Width) / 2; $t.Top = ($areaH - [double]$t.Height) / 2
        $shape = New-GlyphShape $t
        $t.Delete()
    }
    $shape.Name = 'ClipShape'

    # 2. outline copy
    if ($Outline) {
        $edge = $shape.Duplicate(); $edge.Left = $shape.Left; $edge.Top = $shape.Top
        $edge.Name = 'Outline'
        $edge.Fill.Visible = 0
        $edge.Line.Visible = -1; $edge.Line.ForeColor.RGB = (ConvertTo-Bgr $Outline); $edge.Line.Weight = 1.25
    }
    $box = Get-Box $shape

    # 3. GIF covering the shape, then intersect with the GIF as the primary shape
    $g = $sl.Shapes.AddPicture($Gif, 0, -1, 0, 0, -1, -1)
    $g.LockAspectRatio = -1
    $g.Height = $box.H + 30
    if ([double]$g.Width -lt $box.W + 30) { $g.Width = $box.W + 30 }
    $g.Left = $box.L + ($box.W - [double]$g.Width) / 2; $g.Top = $box.T + ($box.H - [double]$g.Height) / 2
    $filled = Merge-Shapes 3 @($g.Name, $shape.Name)
    $filled.Name = 'ShapeWithGif'
    if ($filled.Type -ne 13) { throw "intersect did not keep the picture (type $($filled.Type))" }
    if ($Outline) { $edge.ZOrder(0) }

    # 4. copy line and notes
    if ($Copy) {
        $c = New-Text $Copy 34 0xFFFFFF $Font
        $c.TextFrame2.TextRange.Font.Spacing = 4
        if ([double]$c.Width -gt $W - 120) { $c.TextFrame.TextRange.Font.Size = 34 * ($W - 120) / [double]$c.Width }
        $c.Left = ($W - [double]$c.Width) / 2
        $c.Top = $box.T + $box.H + 44
    }
    if ($Notes) { $sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = $Notes }

    Save-Deck $Out $Preview 5
    "saved $Out"
} catch { Close-Deck; throw }
