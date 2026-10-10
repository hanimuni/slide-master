# A table that reads at a glance: no grid, only horizontal rules (heavy at the top and under the header, hairlines
# between rows), header in bold, one row picked out. Tables do not have to be boxes of equal lines.
# Example: .\clean_table.ps1 -Rows 'Plan|Price|Users','Basic|9,900|1','Team|29,000|10','Company|99,000|100' -Highlight 2 -Title 'PRICING' -Out C:\work\table.pptx -Preview C:\work\table.png
param(
    [Parameter(Mandatory)][string[]]$Rows,        # 'a|b|c' per row; the first row is the header
    [int]$Highlight = 0,                          # 1-based data row to pick out (0 = none)
    [string]$Title = '',
    [string]$Accent = '#E8604C',
    [string]$TextColor = '#1B1B1B',
    [string]$Background = '#FFFFFF',
    [string]$Font = 'Arial',
    [string]$BoldFont = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

function Set-Rule($cell, [int]$side, [double]$weight, [int]$bgr) {   # 1 top, 3 bottom
    $b = $cell.Borders($side); $b.Visible = -1; $b.ForeColor.RGB = $bgr; $b.Weight = [single]$weight; $b.Transparency = 0
}

try {
    New-Deck $Background
    $ink = ConvertTo-Bgr $TextColor; $acc = ConvertTo-Bgr $Accent; $hair = ConvertTo-Bgr '#C9C9C9'
    $grid = @($Rows | ForEach-Object { , ($_ -split '\|') })
    $nr = $grid.Count; $nc = ($grid | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    $top = 110
    if ($Title) { $tt = New-FitText $Title 34 $ink $BoldFont ($W * 0.6); $tt.Left = 70; $tt.Top = 48; $tt.Name = 'Title'; $top = 130 }
    $rowH = [math]::Min(52, ($H - $top - 60) / $nr)
    $shape = $sl.Shapes.AddTable($nr, $nc, 70, [single]$top, [single]($W - 140), [single]($rowH * $nr)); $shape.Name = 'Table'
    $tbl = $shape.Table
    $tbl.ApplyStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307C}')        # "No Style, No Grid": start from nothing
    for ($r = 1; $r -le $nr; $r++) {
        $tbl.Rows($r).Height = [single]$rowH
        $picked = ($Highlight -gt 0 -and $r -eq $Highlight + 1)
        for ($c = 1; $c -le $nc; $c++) {
            $cell = $tbl.Cell($r, $c); $tr = $cell.Shape.TextFrame.TextRange
            $tr.Text = $(if ($c -le $grid[$r - 1].Count) { $grid[$r - 1][$c - 1] } else { '' })
            $bold = ($r -eq 1 -or $picked)
            $face = if ($bold) { $BoldFont } else { $Font }
            $tr.Font.Name = $face; $tr.Font.NameFarEast = $face; $tr.Font.Size = $(if ($r -eq 1) { 14 } else { 16 })
            $tr.Font.Color.RGB = $(if ($picked) { $acc } else { $ink })
            $tr.ParagraphFormat.Alignment = $(if ($c -eq 1) { 1 } else { 2 })
            $cell.Shape.TextFrame.VerticalAnchor = 3                 # middle
            $cell.Shape.Fill.Visible = 0
            if ($r -eq 1) { Set-Rule $cell 1 2.25 $ink; Set-Rule $cell 3 1.5 $ink }
            elseif ($r -eq $nr) { Set-Rule $cell 3 1.5 $ink }
            else { Set-Rule $cell 3 0.5 $hair }
        }
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
