# Choropleth map slide: the map from make_choropleth.py is drawn as native freeforms, one per country and named after
# it (by hand: insert the SVG, Graphics Format > Convert to Shape, ungroup with Ctrl+Shift+G). The slide also gets a
# title with a tag, a source line, value labels on the countries listed in the sidecar .json, a legend bar and an
# optional remark box. PowerPoint drops SVG patterns and filters, so the map uses flat fills only ("no data" = grey).
# Example:
#   python make_choropleth.py countries-110m.json co2.csv C:\work\co2 --label China India
#   .\map_chart.ps1 -Svg C:\work\co2.svg -Title 'CO2 emissions by country' -Tag Total -Sub 'Mt, 2023' -Unit Mt -Out C:\work\map.pptx -Preview C:\work\map.png
param(
    [Parameter(Mandatory)][string]$Svg,           # the .json written next to it is read too
    [string]$Title = '',
    [string]$Tag = '',                            # short word on a coloured pill after the title
    [string]$Sub = '',                            # unit and source
    [string]$Unit = '',
    [string]$Names = '',                          # display names: 'China=CHN|United States of America=USA'
    [string]$Note = '',                           # remark in a light box ('|' = line break)
    [double[]]$NoteAt = @(0.5, 0.42),             # top-left corner of the remark box, as fractions of the slide
    [string]$Accent = '#C8261B',
    [string]$Background = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [switch]$KeepPicture,                         # insert the SVG as one picture instead (faster; not editable country by country)
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
$side = [IO.Path]::ChangeExtension($Svg, '.json')
foreach ($p in $Svg, $side) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $info = Get-Content $side -Raw -Encoding UTF8 | ConvertFrom-Json
    $alias = @{}; foreach ($pair in ($Names -split '\|' | Where-Object { $_ })) { $kv = $pair -split '=', 2; $alias[$kv[0]] = $kv[1] }
    $m = 44; $acc = ConvertTo-Bgr $Accent; $ink = ConvertTo-Bgr '#1B1B1B'

    $y = 34
    if ($Title) {
        $t = New-FitText $Title 24 $ink $Font ($W - 2 * $m - 140); $t.Left = $m; $t.Top = $y; $t.Name = 'Title'
        if ($Tag) {
            $tg = New-Text $Tag 14 0xFFFFFF $Font; $pill = $sl.Shapes.AddShape(5, [single]($m + [double]$t.Width + 12), [single]($y + 3), [single]([double]$tg.Width + 26), [single]([double]$t.Height - 6))
            $pill.Adjustments.Item(1) = 0.3; $pill.Fill.ForeColor.RGB = $acc; $pill.Line.Visible = 0; $pill.Name = 'Tag'
            $tg.Left = [single]([double]$pill.Left + 13); $tg.Top = [single]([double]$pill.Top + ([double]$pill.Height - [double]$tg.Height) / 2); $tg.ZOrder(0)
        }
        $y += [double]$t.Height + 2
    }
    if ($Sub) { $s2 = New-FitText $Sub 11 (ConvertTo-Bgr '#8A8A8A') 'Arial' ($W - 2 * $m); $s2.Left = $m; $s2.Top = [single]$y; $y += [double]$s2.Height }

    # map: fit the picture into the area between the heading and the legend
    $top = $y + 10; $areaH = $H - $top - 62; $areaW = $W - 2 * $m
    $scale = [math]::Min($areaW / [double]$info.width, $areaH / [double]$info.height)
    $box = @{ W = [double]$info.width * $scale; H = [double]$info.height * $scale }; $box.L = $m + ($areaW - $box.W) / 2; $box.T = $top + ($areaH - $box.H) / 2
    if ($KeepPicture) {
        $pic = $sl.Shapes.AddPicture($Svg, 0, -1, [single]$box.L, [single]$box.T, [single]$box.W, [single]$box.H); $pic.Name = 'Map'
    } else {
        # one freeform per outline, named after the country, so each can be recoloured or deleted by hand
        $count = 0
        foreach ($c in $info.shapes) {
            $k = 0
            foreach ($ring in $c.rings) {
                $k++; $fb = $sl.Shapes.BuildFreeform(0, [single]($box.L + $box.W * $ring[0][0]), [single]($box.T + $box.H * $ring[0][1]))
                for ($q = 1; $q -lt $ring.Count; $q++) { $fb.AddNodes(0, 0, [single]($box.L + $box.W * $ring[$q][0]), [single]($box.T + $box.H * $ring[$q][1])) }
                $f = $fb.ConvertToShape(); $f.Fill.ForeColor.RGB = (ConvertTo-Bgr $c.fill); $f.Line.ForeColor.RGB = 0xFFFFFF; $f.Line.Weight = 0.25
                $f.Name = if ($k -eq 1) { $c.name } else { "$($c.name) $k" }; $count++
            }
        }
        "map: $count editable shapes"
    }

    foreach ($prop in $info.labels.PSObject.Properties) {            # country name (small) over the value (bold)
        $lab = $prop.Value; $shown = if ($alias.ContainsKey($prop.Name)) { $alias[$prop.Name] } else { $prop.Name }
        $tone = if ($lab.dark) { 0xFFFFFF } else { $ink }
        $cx = $box.L + $box.W * [double]$lab.x; $cy = $box.T + $box.H * [double]$lab.y
        $a = New-Text $shown 8 $tone 'Arial'; $a.Left = [single]($cx - [double]$a.Width / 2); $a.Top = [single]($cy - [double]$a.Height)
        $b = New-Text ('{0:N0}' -f [double]$lab.value) 11 $tone $Font; $b.Left = [single]($cx - [double]$b.Width / 2); $b.Top = [single]($cy - 2)
    }

    # legend: one gradient bar through every colour of the ramp, then the "no data" swatch
    $ly = $H - 40; $lw = 220; $bar = New-Rect $m ([single]$ly) $lw 10; $bar.Line.Visible = 0; $bar.Name = 'Legend'
    $cols = @($info.colors); $bar.Fill.TwoColorGradient(1, 1); $bar.Fill.GradientAngle = 0
    $st = $bar.Fill.GradientStops
    for ($i = 1; $i -lt $cols.Count - 1; $i++) { $st.Insert((ConvertTo-Bgr $cols[$i]), [single]($i / ($cols.Count - 1)), 0, $i + 1) }
    $st.Item(1).Color.RGB = (ConvertTo-Bgr $cols[0]); $st.Item($st.Count).Color.RGB = (ConvertTo-Bgr $cols[-1])
    $l0 = New-Text '0' 9 $ink 'Arial'; $l0.Left = $m; $l0.Top = [single]($ly + 13)
    $l1 = New-Text ("{0:N0} $Unit" -f [double]$info.max).Trim() 9 $ink 'Arial'; $l1.Left = [single]($m + $lw - [double]$l1.Width); $l1.Top = [single]($ly + 13)
    $sw = New-Rect ([single]($m + $lw + 24)) ([single]$ly) 14 10; $sw.Fill.ForeColor.RGB = (ConvertTo-Bgr $info.nodata); $sw.Line.Visible = 0
    $nd = New-Text 'No data' 9 $ink 'Arial'; $nd.Left = [single]($m + $lw + 44); $nd.Top = [single]($ly - 2)

    if ($Note) {
        $n = New-FitText $Note 10 $acc 'Arial' ($W * 0.26)
        $nb = $sl.Shapes.AddShape(5, [single]($W * $NoteAt[0]), [single]($H * $NoteAt[1]), [single]([double]$n.Width + 24), [single]([double]$n.Height + 16)); $nb.Adjustments.Item(1) = 0.12
        $nb.Fill.ForeColor.RGB = (ConvertTo-Bgr '#F1F1F1'); $nb.Line.Visible = 0; $n.Left = [single]([double]$nb.Left + 12); $n.Top = [single]([double]$nb.Top + 8); $n.ZOrder(0)
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
