# Website-style navigation inside a slide show (Insert > Action). Every section slide has a hot spot in the
# top-right corner: resting the mouse on it jumps to a menu slide (Mouse Over > Hyperlink to > Slide...),
# clicking a menu entry jumps to that section (Mouse Click), and moving the mouse off the menu returns to the
# slide you came from (an invisible shape around the menu, Mouse Over > Last Slide Viewed). The menu panel sits
# above each section slide, so Morph makes it drop down instead of popping up.
# Hover only works in the slide show, not in the editor and not in an exported video.
# Example: .\hover_menu.ps1 -Title 'PLAN 2026' -Sections 'Market|Competitors|Strategy|Schedule' -Out C:\work\nav.pptx -Preview C:\work\nav.png
param(
    [Parameter(Mandatory)][string]$Sections,      # 'a|b|c' - one slide each
    [string]$Title = '',
    [string]$Accent = '#E8604C',
    [string]$PanelColor = '#16181D',
    [string]$Background = '#F4F3EF',
    [string]$Font = 'Arial Black',
    [double]$Duration = 0.5,
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',                        # *.png = picture of the menu slide
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

function Add-Panel([double]$top) {                # menu strip with one entry per section; returns the entry shapes
    $p = New-Rect 0 ([single]$top) $W $panelH; $p.Fill.ForeColor.RGB = (ConvertTo-Bgr $PanelColor); $p.Line.Visible = 0; $p.Name = 'Menu'
    $slot = ($W - 120) / $names.Count; $made = @()
    for ($i = 0; $i -lt $names.Count; $i++) {
        $t = New-Text ('{0:00}  {1}' -f ($i + 1), $names[$i]) 16 0xFFFFFF $Font
        $t.Left = [single](60 + $slot * $i); $t.Top = [single]($top + ($panelH - [double]$t.Height) / 2); $t.Name = "Entry$($i + 1)"
        $made += $t
    }
    , $made
}
function Set-Morph($slide) { $slide.SlideShowTransition.EntryEffect = 3954; $slide.SlideShowTransition.Duration = $Duration }

try {
    New-Deck $Background
    $names = $Sections -split '\|'; $panelH = 110
    $ink = ConvertTo-Bgr '#1B1B1B'; $acc = ConvertTo-Bgr $Accent
    for ($i = 0; $i -lt $names.Count; $i++) {
        if ($i -gt 0) {
            $script:sl = $pres.Slides.Add($i + 1, 12); $sl.FollowMasterBackground = 0
            $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = (ConvertTo-Bgr $Background)
        }
        if ($Title) { $d = New-Text $Title 13 $ink 'Arial'; $d.Left = 60; $d.Top = 40 }
        $head = New-FitText ('{0:00}. {1}' -f ($i + 1), $names[$i]) 44 $ink $Font ($W - 240); $head.Left = 60; $head.Top = 70; $head.Name = 'Heading'   # not $h: PowerShell variables ignore case and $H is the slide height
        $hot = $sl.Shapes.AddShape(9, ($W - 96), 36, 44, 44); $hot.Fill.ForeColor.RGB = $acc; $hot.Line.Visible = 0; $hot.Name = 'Hot'
        [void](Add-Panel (-$panelH))                                  # parked above the slide: Morph drops it down
        Set-Morph $sl
    }
    $n = $names.Count
    $script:sl = $pres.Slides.Add($n + 1, 12); $sl.FollowMasterBackground = 0
    $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = (ConvertTo-Bgr $Background)
    $hint = New-Text 'MENU' 44 (ConvertTo-Bgr '#C9C6BD') $Font; $hint.Left = 60; $hint.Top = [single]($panelH + 40)
    # invisible "leave the menu" area: the slide minus the menu strip. A fully transparent fill still receives the mouse.
    $back = New-Rect 0 ([single]($panelH + 20)) $W ([single]($H - $panelH - 20)); $back.Line.Visible = 0; $back.Fill.Transparency = 1; $back.Name = 'Leave'
    $back.ActionSettings(2).Action = 5                                # ppMouseOver -> ppActionLastSlideViewed
    $entries = Add-Panel 0
    for ($i = 0; $i -lt $n; $i++) {
        $a = $entries[$i].ActionSettings(1)                           # ppMouseClick -> that section's slide
        $a.Action = 7; $a.Hyperlink.SubAddress = "$($pres.Slides($i + 1).SlideID),$($i + 1),"
    }
    Set-Morph $sl
    for ($i = 1; $i -le $n; $i++) {                                   # hot spots: mouse over -> menu slide
        $a = $pres.Slides($i).Shapes('Hot').ActionSettings(2)
        $a.Action = 7; $a.Hyperlink.SubAddress = "$($sl.SlideID),$($n + 1),"
    }
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
