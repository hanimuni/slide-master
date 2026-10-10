# Hover glossary: resting the mouse on a marked word slides a note card in; moving off it puts the page back.
# Slide 1 = the page (the card waits outside the right edge). Slide 2 = the same page with the card in view and an
# invisible shape over everything except the word and the card; hovering that shape returns to slide 1. Both
# directions use Morph, so the card glides instead of popping. Works in the slide show only.
# Example: .\hover_note.ps1 -Title 'Report' -Before 'We apply' -Word 'digital twin' -After 'to every line.' -Note 'A virtual copy of the factory|used to test changes first.' -Out C:\work\note.pptx -Preview C:\work\note.png
param(
    [Parameter(Mandatory)][string]$Word,          # the marked word
    [Parameter(Mandatory)][string]$Note,          # card text, '|' = line break
    [string]$Before = '', [string]$After = '',    # sentence around the word
    [string]$Title = '',
    [string]$Accent = '#D0342C',
    [string]$Background = '#F4F3EF',
    [string]$Font = 'Arial Black',
    [double]$Duration = 0.5,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',                        # *.png = picture of slide 2 (card shown)
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck $Background
    $ink = ConvertTo-Bgr '#1B1B1B'; $acc = ConvertTo-Bgr $Accent
    if ($Title) { $tt = New-FitText $Title 34 $ink $Font ($W - 200); $tt.Left = 70; $tt.Top = 60 }
    $x = 70; $y = 190
    if ($Before) { $b = New-Text "$Before " 22 $ink 'Arial'; $b.Left = [single]$x; $b.Top = [single]$y; $x += [double]$b.Width }
    $wd = New-Text $Word 22 $acc $Font; $wd.Left = [single]$x; $wd.Top = [single]$y; $wd.Name = 'Word'; $wd.TextFrame.TextRange.Font.Underline = -1
    $wx = $x; $ww = [double]$wd.Width; $wh = [double]$wd.Height; $x += $ww
    if ($After) { $a = New-Text " $After" 22 $ink 'Arial'; $a.Left = [single]$x; $a.Top = [single]$y }
    $cw = 330; $chh = 150; $cardTop = $y + $wh + 26
    $card = $sl.Shapes.AddShape(5, [single]($W + 30), [single]$cardTop, $cw, $chh); $card.Adjustments.Item(1) = 0.08; $card.Name = 'Card'
    $card.Fill.ForeColor.RGB = (ConvertTo-Bgr '#16181D'); $card.Line.Visible = 0
    $nt = New-FitText $Note 15 0xFFFFFF 'Arial' ($cw - 40); $nt.Name = 'NoteText'; $nt.Left = [single]($W + 50); $nt.Top = [single]($cardTop + 22)
    $first = $sl
    Add-MorphSlide $Duration                                           # 2: card in view
    $cl = [math]::Min($W - $cw - 40, [math]::Max(40, $wx))
    $sl.Shapes('Card').Left = [single]$cl; $sl.Shapes('NoteText').Left = [single]($cl + 20)
    # the "leave" area = slide minus (word + card): four rectangles united into one invisible shape
    $l = [math]::Min($wx, $cl) - 6; $r = [math]::Max($wx + $ww, $cl + $cw) + 6; $t = $y - 6; $bt = $cardTop + $chh + 6
    $parts = @(@(0, 0, $W, $t), @(0, $bt, $W, ($H - $bt)), @(0, $t, $l, ($bt - $t)), @($r, $t, ($W - $r), ($bt - $t)))
    $names = @(); $k = 0
    foreach ($p in $parts) { $q = New-Rect ([single]$p[0]) ([single]$p[1]) ([single]$p[2]) ([single]$p[3]); $q.Name = "Leave$k"; $names += $q.Name; $k++ }
    $leave = Merge-Shapes 1 $names; $leave.Name = 'Leave'; $leave.Line.Visible = 0; $leave.Fill.Transparency = 1
    $act = $leave.ActionSettings(2); $act.Action = 7; $act.Hyperlink.SubAddress = "$($first.SlideID),1,"
    $go = $first.Shapes('Word').ActionSettings(2); $go.Action = 7; $go.Hyperlink.SubAddress = "$($sl.SlideID),2,"
    $first.SlideShowTransition.EntryEffect = 3954; $first.SlideShowTransition.Duration = $Duration
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
