# A small deck built on deck_helpers.ps1: copy this file to start a whole deck (ppt-craft.md, chapter 31).
# Six pages in one frame: contents, figures on a dark page, an emphasised before/after (two-step Morph), rows beside a photo
# (or beside a figure card when no -Image is given), a timeline, and a stepped card row. Cover and closing are left to a recipe
# or to the deck's own scene: they carry no frame.
# Example: .\deck_demo.ps1 -Out C:\work\demo.pptx -PreviewDir C:\work\demo -Image C:\work\photo.jpg
param(
    [string]$Image = '',
    [string]$Accent = '#2E6BFF',
    [string]$AccentDark = '',
    [string]$Light = '#F3F6FB',
    [string]$Dark = '#0A1B3D',
    [string]$Ink = '#111A2E',
    [string]$Font = 'Malgun Gothic',
    [string]$HeavyFont = '',
    [Parameter(Mandatory)][string]$Out,
    [string]$PreviewDir = '',
    [string]$Video = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if ($Image -and -not (Test-Path $Image)) { throw "not found: $Image" }
Assert-Output @($Out) $Force.IsPresent
. (Join-Path $PSScriptRoot 'scene_helpers.ps1')
. (Join-Path $PSScriptRoot 'deck_helpers.ps1')

try {
    Set-DeckFrame -Sections 'Results', 'Vision', 'Tasks', 'Roadmap' -Footer 'Sample Agency  -  Work Plan 2026' -Pages 6

    # 01 contents: tall cards with large numerals
    New-Deck $Light; $sl.SlideShowTransition.EntryEffect = 3849
    Chrome -1 'CONTENTS' 'What this deck covers'
    $toc = @(@('01', 'Results', 'What changed|last year'), @('02', 'Vision', 'Four goals|ten tasks'), @('03', 'Tasks', 'Key figures|per goal'), @('04', 'Roadmap', '2026 to 2030|by year'))
    $cols = Cols 4
    for ($i = 0; $i -lt 4; $i++) { Rise (TCard $cols[$i][0] 140 $cols[$i][1] 300 $toc[$i][0] $toc[$i][1] $toc[$i][2] '' 64) }

    # 02 figures on a dark page (one of the few dark slides)
    NewSlide $Dark; Chrome 0 'RESULTS 2025' 'The year in four figures' $true
    $st = @(@('39.9k', 'Talent pool', 'up 16,747 in a year'), @('11.1%', 'Old practice', 'down from 18.1%'), @('10d', 'Family leave', 'new in July'), @('5-7d', 'Service leave', 'after 10 / 20 years'))
    $half = Cols 2
    for ($i = 0; $i -lt 4; $i++) { Rise (StatWide $half[$i % 2][0] (136 + [math]::Floor($i / 2) * 170) $half[$i % 2][1] 152 $st[$i][0] $st[$i][1] $st[$i][2] $true) }

    # 03 emphasis: the "before" figure alone, then Morph to the full slide
    NewSlide $Light; Chrome 2 'TASKS - INCENTIVES' 'Reward what serves the public'
    $script:static = $true
    $ba = BA 56 130 440 340 'Share of top performers with a bold reward' '50%' '75' '%' 'Promotion, pay step, top bonus|priority for training abroad' 100
    $rc = @(@('01', 'More awards', 'At least 50% more commendations'), @('02', 'Wider insurance', 'Counsel for criminal cases too'), @('03', 'One council', 'Progress checks and tailored advice'))
    $rest = @(); for ($i = 0; $i -lt 3; $i++) { $rest += (RowCard 514 (130 + $i * 119) 390 102 $rc[$i][0] $rc[$i][1] $rc[$i][2]) }
    Emph $ba $rest 120; $script:static = $false

    # 04 rows with small before -> after figures, beside a photo (or a figure card)
    NewSlide $Light; Chrome 2 'TASKS - SAFETY' 'The state takes responsibility'
    $rb = @(@('10', '17', 'Care centres', 'One per region by 2030'), @('100', '200', 'Review panel', 'More specialists on the panel'), @('30', '45', 'Rehab allowance', 'Easier to qualify'))
    for ($i = 0; $i -lt 3; $i++) { Rise (RowBA 56 (130 + $i * 119) 480 102 $rb[$i][0] $rb[$i][1] $rb[$i][2] $rb[$i][3] 16) }
    if ($Image) { Anim (Photo $Image 554 130 350 340 0.5) $FADE $AFTER 0.6 }
    else { Rise (BA 554 130 350 340 'Care centres' '10' '17' '' 'One per region|by 2030' 84) }

    # 05 timeline with a governing message under the title
    NewSlide $Light; Chrome 3 '2026 - 2030' 'Milestones to 2030' $false 'The figures we commit to, year by year'
    Timeline @(@('2026', 'Study', 'Design the|new test'), @('2027', 'Plan published', 'Two years|of notice'), @('2028', 'Stage one', 'Civil service and|large agencies'), @('2030', 'Centre opens', 'One place for|all hiring')) ($script:BodyY + 90)
    $parts = @((Card 56 396 848 84), (Box 56 412 5 52 $Accent), (Tx 'Still to come' 19 $Ink 82 408 400 -1), (Tx '5,000 more homes for staff  -  new digital specialists every year' 15 $body 82 440 800))
    Rise $parts

    # 06 a stepped row: one dominant card instead of four equal text cards
    NewSlide $Light; Chrome 2 'TASKS - PROMOTION' 'Grow experts, promote on results'
    [void](Tx 'Skill and results lead to promotion' 18 $body 56 132 420)
    $tc = @(@('01', 'Track', 'Expert track', 'Stay longer in|expert fields'), @('02', 'G5', 'Fast promotion', 'Picked by the|ministries'), @('03', 'G6', 'Open posts', 'Open to|grade six'), @('04', 'Always', 'Ongoing review', 'Feedback all|year round'))
    for ($i = 0; $i -lt 4; $i++) {
        $x = $cols[$i][0]; $cwid = $cols[$i][1]; $top = 256 - $i * 38; $r = $tc[$i]; $hot = ($i -eq 3)
        $c = Card $x $top $cwid (470 - $top) $(if ($hot) { $Accent } else { '#FFFFFF' }); $parts = @($c)
        $parts += (Tx $r[0] 15 $(if ($hot) { Mix $Accent '#FFFFFF' 0.8 } else { $Accent }) ($x + 20) ($top + 18) 60 $heavyBold $heavy)
        $parts += (Tx $r[1] 46 $(if ($hot) { '#FFFFFF' } else { $Accent }) ($x + 18) ($top + 36) ($cwid - 30) $heavyBold $heavy)
        $parts += (Tx $r[2] 20 $(if ($hot) { '#FFFFFF' } else { $Ink }) ($x + 20) ($top + 102) ($cwid - 30) -1)
        $parts += (Tx $r[3] 16 $(if ($hot) { Mix $Accent '#FFFFFF' 0.88 } else { $body }) ($x + 20) ($top + 134) ($cwid - 30) 0 $Font 1 0 1.45)
        Rise $parts
    }

    Save-DeckAll $Out $PreviewDir $Video
} catch { Close-Deck; throw }
