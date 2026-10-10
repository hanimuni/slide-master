# Morph instead of animations: slide 1 shows a title block, slide 2 the same block moved aside while a
# list panel that was parked outside the slide travels in. The pattern for any before/after pair.
# Example: .\morph_reveal.ps1 -Title 'POWER|POINT|TITLE' -Items 'SITUATION|STRATEGY|CONCEPT|PROGRAM' -Out C:\work\reveal.pptx -Preview C:\work\reveal.mp4
param(
    [Parameter(Mandatory)][string]$Title,         # '|' = line break
    [Parameter(Mandatory)][string]$Items,         # '|' separated; numbered automatically
    [string]$Note = '',                           # small line in the second tile
    [string]$Image = '',                          # optional photo for the right half of the title tile
    [string[]]$Palette = @('#FCEBD3', '#F59A3C', '#B6F04A'),   # title tile, lower-left tile, lower-right tile
    [string]$Background = '#000000',
    [string]$Font = 'Arial Black',
    [double]$Duration = 1,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Palette.Count -ne 3) { throw '-Palette needs three colours' }
if ($Image -and -not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out, $Preview) $Force.IsPresent

function New-Tile($l, $t, $w, $h, [string]$hex, [string]$name) {
    $r = New-Rect $l $t $w $h; $r.Fill.ForeColor.RGB = (ConvertTo-Bgr $hex); $r.Line.Visible = 0; $r.Name = $name
}

try {
    New-Deck $Background
    $bw = 300; $x = ($W - $bw) / 2; $y = 110
    New-Tile $x $y $bw 170 $Palette[0] 'TileTitle'
    New-Tile $x ($y + 174) 148 146 $Palette[1] 'TileA'
    New-Tile ($x + 152) ($y + 174) 148 146 $Palette[2] 'TileB'
    if ($Image) {
        # cover the right half of the title tile with the photo: scale to fill, then intersect (picture first)
        $slot = New-Rect ($x + $bw / 2) $y ($bw / 2) 170; $slot.Name = 'PhotoSlot'
        $pic = $sl.Shapes.AddPicture($Image, 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = 'PhotoRaw'
        $pic.Height = 170
        if ([double]$pic.Width -lt $bw / 2) { $pic.Width = [single]($bw / 2) }
        $pic.Left = [single]($x + $bw / 2 + ($bw / 2 - [double]$pic.Width) / 2); $pic.Top = [single]($y + (170 - [double]$pic.Height) / 2)
        $ph = Merge-Shapes 3 @('PhotoRaw', 'PhotoSlot'); $ph.Name = 'TilePhoto'
    }
    $t = New-FitText $Title 34 0 $Font ($(if ($Image) { $bw / 2 } else { $bw }) - 40); $t.Left = $x + 18; $t.Top = $y + 170 - [double]$t.Height - 14; $t.Name = 'TitleText'
    if ($Note) { $n = New-FitText $Note 12 0 $Font 128; $n.Left = $x + 162; $n.Top = $y + 174 + 146 - [double]$n.Height - 12; $n.Name = 'NoteText' }

    # the panel starts outside the right edge; Morph brings it in because slide 2 has it inside
    $panelW = 280; $park = $W + 40
    $p = $sl.Shapes.AddShape(5, $park, $y, $panelW, 320); $p.Name = 'Panel'
    $p.Adjustments.Item(1) = 0.05; $p.Fill.Visible = 0; $p.Line.ForeColor.RGB = 0xFFFFFF; $p.Line.Weight = 1.25
    $i = 0
    $list = ($Items.Split('|') | ForEach-Object { $i++; "$i. $_" }) -join '|'
    $lt = New-FitText $list 26 0xFFFFFF $Font ($panelW - 60)
    $lt.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 14
    $lt.Left = $park + 30; $lt.Top = $y + (320 - [double]$lt.Height) / 2; $lt.Name = 'PanelText'

    Add-MorphSlide $Duration
    $shift = ($panelW + 16) / 2
    foreach ($s in $sl.Shapes) {
        if ($s.Name -like 'Panel*') { $s.Left = [single]([double]$s.Left - ($park - ($x + $bw + 16 - $shift))) }
        else { $s.Left = [single]([double]$s.Left - $shift) }
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
