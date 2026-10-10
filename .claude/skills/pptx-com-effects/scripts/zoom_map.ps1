# Clickable overview that zooms into its topics - the look of PowerPoint's Zoom feature, built from hyperlinks and
# Morph because Zoom objects cannot be created through the object model. Slide 1 is a photo with round nodes joined
# by dashed lines; clicking a node jumps to that topic's slide, which is the same picture enlarged around the node
# (Morph makes the camera fly in) with the topic text underneath; clicking anywhere there flies back to the overview.
# Example: .\zoom_map.ps1 -Image forest.jpg -Title 'ROADMAP' -Nodes 'IDEA|MARKET|SOLUTION|MODEL' -Bodies 'a|b|c|d' -Out C:\work\map.pptx -Preview C:\work\map.mp4
param(
    [Parameter(Mandatory)][string]$Image,
    [Parameter(Mandatory)][string]$Nodes,         # 'a|b|c' - 2 to 6 topics
    [string]$Bodies = '',                         # one text per topic (';' = line break)
    [string]$Title = '',
    [string]$Color = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [double]$Zoom = 2.6,
    [double]$Duration = 0.9,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck '#000000'
    $names = $Nodes -split '\|'; $texts = if ($Bodies) { $Bodies -split '\|' } else { @() }
    $n = $names.Count; $c = ConvertTo-Bgr $Color; $d = 150
    $photo = Add-Cover $Image; $photo.Name = 'Photo'
    $dim = New-Rect 0 0 $W $H; $dim.Fill.ForeColor.RGB = 0; $dim.Fill.Transparency = 0.45; $dim.Line.Visible = 0; $dim.Name = 'Dim'
    $scaled = @('Photo', 'Dim')                                       # everything that takes part in the zoom
    if ($Title) {
        $tt = New-FitText $Title 40 $c $Font ($W - 200); $tt.Left = [single](($W - [double]$tt.Width) / 2); $tt.Top = 34
        $g = New-GlyphShape $tt; $g.Name = 'MapTitle'; $g.Fill.ForeColor.RGB = $c; $g.Line.Visible = 0; $tt.Delete(); $scaled += 'MapTitle'
    }
    $cx = @(); $cy = @()
    for ($i = 0; $i -lt $n; $i++) { $cx += (120 + ($W - 240) / [math]::Max(1, $n - 1) * $i); $cy += $(if ($i % 2) { 250 } else { 370 }) }
    for ($i = 0; $i -lt $n - 1; $i++) {                               # dashed links first, so the circles cover their ends
        $ln = $sl.Shapes.AddLine([single]$cx[$i], [single]$cy[$i], [single]$cx[$i + 1], [single]$cy[$i + 1]); $ln.Name = "Link$i"
        $ln.Line.ForeColor.RGB = $c; $ln.Line.Weight = 3; $ln.Line.DashStyle = 4; $scaled += $ln.Name
    }
    for ($i = 0; $i -lt $n; $i++) {
        $o = $sl.Shapes.AddShape(9, [single]($cx[$i] - $d / 2), [single]($cy[$i] - $d / 2), $d, $d); $o.Name = "Node$i"
        $o.Fill.ForeColor.RGB = 0x202020; $o.Fill.Transparency = 0.25; $o.Line.ForeColor.RGB = $c; $o.Line.Weight = 3; $scaled += $o.Name
        $t = New-FitText $names[$i] 15 $c $Font ($d - 26); $t.Left = [single]($cx[$i] - [double]$t.Width / 2); $t.Top = [single]($cy[$i] - [double]$t.Height / 2)
        $g = New-GlyphShape $t; $g.Name = "Label$i"; $g.Fill.ForeColor.RGB = $c; $g.Line.Visible = 0; $t.Delete(); $scaled += $g.Name
    }
    $base = @{}; foreach ($s in $scaled) { $x = $sl.Shapes($s); $base[$s] = @{ L = [double]$x.Left; T = [double]$x.Top; W = [double]$x.Width; H = [double]$x.Height } }
    $map = $sl
    for ($i = 0; $i -lt $n; $i++) {
        [void]$map.Duplicate()                                        # a copy lands right after the original (index 2)
        $pres.Slides(2).MoveTo($i + 2); $script:sl = $pres.Slides($i + 2)       # ... so move it behind the earlier topics
        foreach ($s in $scaled) {                                     # enlarge around node i and bring it to the upper middle
            $x = $sl.Shapes($s); $b = $base[$s]
            $x.Left = [single]($W / 2 + ($b.L - $cx[$i]) * $Zoom); $x.Top = [single](200 + ($b.T - $cy[$i]) * $Zoom)
            $x.Width = [single]($b.W * $Zoom); if ($s -notlike 'Link*' -or $b.H -gt 0.5) { $x.Height = [single]($b.H * $Zoom) }
            if ($s -like 'Link*' -or $s -like 'Node*') { $x.Line.Weight = [single](3 * $Zoom / 1.6) }
        }
        if ($i -lt $texts.Count) {
            $bd = New-FitText ($texts[$i] -replace ';', '|') 20 $c 'Arial' ($W - 240); $bd.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $bd.Left = [single](($W - [double]$bd.Width) / 2); $bd.Top = 420; $bd.Name = 'Body'
        }
        $back = New-Rect 0 0 $W $H; $back.Fill.Transparency = 1; $back.Line.Visible = 0; $back.Name = 'Back'
        $a = $back.ActionSettings(1); $a.Action = 7; $a.Hyperlink.SubAddress = "$($map.SlideID),1,"          # click anywhere -> overview
        $sl.SlideShowTransition.EntryEffect = 3954; $sl.SlideShowTransition.Duration = $Duration
    }
    $map.SlideShowTransition.EntryEffect = 3954; $map.SlideShowTransition.Duration = $Duration
    for ($i = 0; $i -lt $n; $i++) {                                   # nodes on the overview: click -> topic slide
        foreach ($nm in "Node$i", "Label$i") { $a = $map.Shapes($nm).ActionSettings(1); $a.Action = 7; $a.Hyperlink.SubAddress = "$($pres.Slides($i + 2).SlideID),$($i + 2)," }
    }
    $script:sl = $pres.Slides($pres.Slides.Count)
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
