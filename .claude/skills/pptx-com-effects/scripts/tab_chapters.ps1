# Folder-tab chapters: cards with a tab on the right edge lie stacked like index dividers. The cover shows them as
# steps; on each following slide one more card is pushed to the left, so the next card's content comes into view.
# One slide per state, Morph between them (the cards and their texts keep the same "!!" names on every slide).
# By hand: rectangle + small rounded rectangle on its right edge -> Merge Shapes > Union (the shadow then follows the
# tab), one card per chapter with the tab one step lower each time, lighter colour toward the front, and a shadow to
# the right (transparency 60 %, blur 10 pt, distance 5 pt). Put each card's content on its right half only.
# Example: .\tab_chapters.ps1 -Title SLIDES -Kicker chapter -Items 'Plan;short text|Build;short text|Test;short text|Ship;short text' -Out C:\work\tabs.pptx -Preview C:\work\tabs.mp4
param(
    [Parameter(Mandatory)][string]$Items,         # 'heading;text|heading;text|...' (three to five chapters)
    [string]$Title = 'SLIDES',
    [string]$Kicker = 'chapter',                  # small script word above the title
    [string[]]$Images = @(),                      # optional icon per chapter (transparent PNG / SVG)
    [string[]]$Colors = @('#DCCBF2', '#D3A2F2', '#8B45C0', '#4F2A90', '#35206B'),   # front (light) to back (dark)
    [string]$Background = '#1A0F3D',
    [string]$Accent = '#F5A623',                  # kicker
    [string]$Font = 'Arial Black',
    [string]$ScriptFont = 'Segoe Script',
    [double]$Duration = 0.8,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Images) { if (-not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

# Right edge of card $k (1 = front) when chapter $c is open (0 = cover), as a fraction of the slide width.
function Edge([int]$k, [int]$c) {
    if ($c -eq 0) { return 0.3 + 0.09 * $k }      # cover: even steps
    if ($k -lt $c) { return 0.05 + 0.06 * ($k - 1) }   # already read: pushed left, only a strip and its tab stay
    0.58 + 0.07 * ($k - $c)                       # open card, then the ones still waiting behind it
}
function Draw([int]$c) {
    $cardW = $W * 0.92; $tabW = 30; $tabH = $H * 0.13
    for ($k = $n; $k -ge 1; $k--) {               # back to front
        $right = $W * (Edge $k $c); $tone = ConvertTo-Bgr $Colors[($k - 1) % $Colors.Count]; $lightCard = ($k -le 2)
        $ty = $H * 0.16 + ($k - 1) * ($H * 0.68 - $tabH) / [math]::Max(1, $n - 1)
        $body = New-Rect ([single]($right - $cardW)) 0 ([single]$cardW) ([single]$H); $body.Name = 'B'
        $tab = $sl.Shapes.AddShape(5, [single]($right - 20), [single]$ty, [single]($tabW + 20), [single]$tabH); $tab.Adjustments.Item(1) = 0.3; $tab.Name = 'T'
        $card = Merge-Shapes 1 @('B', 'T'); $card.Name = "!!Card$k"; $card.Fill.ForeColor.RGB = $tone; $card.Line.Visible = 0
        $sh = $card.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Transparency = 0.6; $sh.Blur = 10; $sh.OffsetX = 5; $sh.OffsetY = 0
        $ink = if ($lightCard) { ConvertTo-Bgr $Colors[[math]::Min(3, $Colors.Count - 1)] } else { 0xFFFFFF }
        $num = New-Text "$k" 22 $ink $Font; $num.Left = [single]($right + ($tabW - [double]$num.Width) / 2 - 2); $num.Top = [single]($ty + ($tabH - [double]$num.Height) / 2); $num.Name = "!!Tab$k"
        # content sits on the right part of the card, so it shows only while this card is the open one
        $cx = $right - $W * 0.2; $cw = $W * 0.3; $it = $list[$k - 1]
        $big = New-Text ('{0:00}' -f $k) 96 $ink $Font; $big.TextFrame2.TextRange.Font.Fill.Transparency = 0.75
        $big.Left = [single]($cx - [double]$big.Width / 2); $big.Top = [single]($H * 0.06); $big.Name = "!!Big$k"
        $hd = New-FitText $it[0] 20 $ink $Font $cw; $hd.TextFrame.TextRange.ParagraphFormat.Alignment = 2; $hd.Left = [single]($cx - [double]$hd.Width / 2); $hd.Top = [single]($H * 0.3); $hd.Name = "!!Head$k"
        if ($it.Count -gt 1) {
            $tx = $sl.Shapes.AddTextbox(1, [single]($cx - $cw / 2), [single]($H * 0.39), [single]$cw, 80); $r = $tx.TextFrame.TextRange; $r.Text = $it[1]
            $r.Font.Name = 'Arial'; $r.Font.NameFarEast = 'Malgun Gothic'; $r.Font.Size = 12; $r.Font.Color.RGB = $ink; $r.ParagraphFormat.Alignment = 2; $tx.Name = "!!Text$k"
        }
        if ($k -le $Images.Count) {
            $ic = $sl.Shapes.AddPicture($Images[$k - 1], 0, -1, 0, 0, -1, -1); $ic.LockAspectRatio = -1; $ic.Height = [single]($H * 0.24)
            $ic.Left = [single]($cx - [double]$ic.Width / 2); $ic.Top = [single]($H * 0.64); $ic.Name = "!!Icon$k"
        }
    }
    # title on the dark area at the right: large on the cover, small once a card is open
    $size = if ($c -eq 0) { 60 } else { 26 }; $tx0 = $W * (Edge $n $c) + 46; $room = $W - $tx0 - 20
    $t = New-FitText $Title $size 0xFFFFFF $Font $room; $t.Left = [single]($tx0 + ($room - [double]$t.Width) / 2); $t.Top = [single]($H * 0.46); $t.Name = '!!Title'
    $kk = New-FitText $Kicker ($size * 0.6) (ConvertTo-Bgr $Accent) $ScriptFont $room; $kk.Left = [single]($tx0 + ($room - [double]$kk.Width) / 2); $kk.Top = [single]([double]$t.Top - [double]$kk.Height + 6); $kk.Name = '!!Kicker'
}

try {
    New-Deck $Background
    $list = @($Items -split '\|' | Where-Object { $_ } | ForEach-Object { , @($_ -split ';', 2) }); $n = $list.Count
    Draw 0
    for ($c = 1; $c -le $n; $c++) {
        $script:sl = $pres.Slides.Add($c + 1, 12); $sl.FollowMasterBackground = 0; $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = (ConvertTo-Bgr $Background)
        $sl.SlideShowTransition.EntryEffect = 3954; $sl.SlideShowTransition.Duration = $Duration      # Morph
        Draw $c
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
