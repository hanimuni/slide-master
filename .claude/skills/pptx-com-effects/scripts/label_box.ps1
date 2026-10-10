# Label box: a thin outlined box whose title sits on the top edge and breaks the line - for an order of events,
# a programme or a short contents list. By hand: draw the outline, then put the title in a text box filled with
# the background colour so it hides the piece of line behind it. Here the outline is drawn as one open path that
# leaves a gap for the title, so it also works over a photo. -Double adds a second, inner outline (frame look).
# Example: .\label_box.ps1 -Label PROGRAM -Items 'Opening|Toast|Awards|Concert|Closing' -Out C:\work\order.pptx -Preview C:\work\order.png
param(
    [Parameter(Mandatory)][string]$Label,
    [Parameter(Mandatory)][string]$Items,         # 'line|line|...'
    [string]$Image = '',                          # optional full-slide background photo (keep it dark)
    [string]$Accent = '#D8B25A',                  # outline and label
    [string]$Color = '#FFFFFF',
    [string]$Background = '#0A0A0A',
    [string]$Font = 'Arial Black',
    [double]$Width = 0.3,                         # box width as a fraction of the slide
    [switch]$Double,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Image -and -not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    if ($Image) { [void](Add-Cover $Image) }
    $acc = ConvertTo-Bgr $Accent; $bw = $W * $Width; $bh = $H * 0.76; $bx = ($W - $bw) / 2; $by = ($H - $bh) / 2 + 12
    $lab = New-FitText $Label 26 $acc $Font ($bw - 60)
    $g1 = ($W - [double]$lab.Width) / 2 - 18; $g2 = ($W + [double]$lab.Width) / 2 + 18      # the gap the label sits in
    $fb = $sl.Shapes.BuildFreeform(0, [single]$g2, [single]$by)
    foreach ($q in @(@(($bx + $bw), $by), @(($bx + $bw), ($by + $bh)), @($bx, ($by + $bh)), @($bx, $by), @($g1, $by))) { $fb.AddNodes(0, 0, [single]$q[0], [single]$q[1]) }
    $box = $fb.ConvertToShape(); $box.Fill.Visible = 0; $box.Line.ForeColor.RGB = $acc; $box.Line.Weight = 1.5; $box.Name = 'Box'
    if ($Double) {
        $in = New-Rect ([single]($bx + 10)) ([single]($by + 10)) ([single]($bw - 20)) ([single]($bh - 20)); $in.Fill.Visible = 0; $in.Line.ForeColor.RGB = $acc; $in.Line.Weight = 0.75
    }
    $lab.Left = [single](($W - [double]$lab.Width) / 2); $lab.Top = [single]($by - [double]$lab.Height / 2); $lab.ZOrder(0)
    $list = @($Items -split '\|' | Where-Object { $_ }); $step = ($bh - 90) / $list.Count
    for ($i = 0; $i -lt $list.Count; $i++) {
        $t = New-FitText $list[$i] 18 (ConvertTo-Bgr $Color) 'Arial' ($bw - 50); $t.Left = [single](($W - [double]$t.Width) / 2); $t.Top = [single]($by + 55 + $step * $i + ($step - [double]$t.Height) / 2)
    }
    Save-Deck $Out $Preview 3
    "saved $Out"
} catch { Close-Deck; throw }
