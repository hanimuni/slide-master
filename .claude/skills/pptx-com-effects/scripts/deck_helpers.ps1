# Whole-deck helpers: one frame (header / body / footer) and a handful of body archetypes, so that every content slide
# of a deck shares its structure, margins, colours and type scale (ppt-craft.md, chapter 31).
# Dot-source AFTER ppt_helpers.ps1 and scene_helpers.ps1 (the caller's $Accent / $Light / $Dark / $Ink / $Font / $HeavyFont
# must exist first), then call Set-DeckFrame once. deck_demo.ps1 is a working example to copy.
# ASCII only. All sizes are for the 960 x 540 slide. Do not use $w / $h / $s as variables in a deck script (see SKILL.md).
#
#   Set-DeckFrame -Sections 'A','B','C' -Footer 'Org  -  Document' -Pages 20     once per deck
#   NewSlide <bg hex>                      next slide (fade in); the first slide comes from New-Deck
#   Chrome <section index or -1> <kicker> <title> [onDark] [lead]                 header + footer, counts the page number
#   Cols <count> [gap]                     @(@(x, width), ...) for equal columns across the body
#   Card / Rise                            the one card style / the one entrance (groups the parts and lets them rise)
#   StatWide  BA  RowCard  RowBA  TCard  Timeline  Photo                          body archetypes (each returns its shapes)
#   Emph <BA shapes> <other shapes> <size> two-step Morph emphasis: "before" figure alone, then the full slide
#   Auto <slide> <seconds>                 advance by itself
#   Save-DeckAll <pptx> [preview folder] [mp4]                                    save, one PNG per slide, optional video, close
$script:Deck = @{ Sections = @(); Footer = ''; Pages = 0; MarginX = 56; BodyTop = 128; BodyBottom = 488; Gap = 18; Radius = 16 }
$script:page = 0; $script:static = $false; $script:BodyY = 128
$ARROW = [string][char]0x2192

function Set-DeckFrame([string[]]$Sections = @(), [string]$Footer = '', [int]$Pages = 0, $MarginX = 56, $BodyTop = 128, $BodyBottom = 488, $Gap = 18, $Radius = 16) {
    $script:Deck = @{ Sections = $Sections; Footer = $Footer; Pages = $Pages; MarginX = $MarginX; BodyTop = $BodyTop; BodyBottom = $BodyBottom; Gap = $Gap; Radius = $Radius }
    $script:page = 0; $script:BodyY = $BodyTop
}
function NewSlide([string]$hex) {
    $script:sl = $pres.Slides.Add($pres.Slides.Count + 1, 12); $sl.FollowMasterBackground = 0
    $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = (ConvertTo-Bgr $hex)
    $sl.SlideShowTransition.EntryEffect = 3849; $sl.SlideShowTransition.Duration = 0.5
}
function Auto($slide, $sec) { $slide.SlideShowTransition.AdvanceOnTime = -1; $slide.SlideShowTransition.AdvanceTime = [single]$sec }
function NoWrap($t) { $t.TextFrame.WordWrap = 0; $t }
function Skip-Page { $script:page++ }                                    # a slide without the frame (cover, closing) still takes a page number

# The frame. $sec = index into Sections of the current section (-1 = none lit). $lead = optional one-line governing
# message under the title; with it the body starts lower: read $script:BodyY after the call.
function Chrome([int]$sec, [string]$kicker, [string]$title, [bool]$onDark = $false, [string]$lead = '') {
    $script:page++; $mx = $Deck.MarginX; $right = $W - $mx
    $fg = $(if ($onDark) { '#FFFFFF' } else { $Ink }); $mute = $(if ($onDark) { Mix $Dark '#FFFFFF' 0.45 } else { $gray })
    if ($kicker) { [void](Tx $kicker 12.5 $(if ($onDark) { $soft } else { $Accent }) $mx 38 480 -1 $Font 1 3) }
    [void](Tx $title 30 $fg $mx 60 ($right - $mx - 40) -1)
    $n = $Deck.Sections.Count; $x = $right - $n * 50
    for ($i = 0; $i -lt $n; $i++) {
        $on = ($i -eq $sec)
        [void](Tx $Deck.Sections[$i] 11 $(if ($on) { $(if ($onDark) { '#FFFFFF' } else { $Accent }) } else { $mute }) $x 30 50 $(if ($on) { -1 } else { 0 }) $Font 2)
        if ($on) { [void](Box ($x + 14) 49 22 2.5 $Accent) }
        $x += 50
    }
    $script:BodyY = $Deck.BodyTop
    if ($lead) { [void](Tx $lead 17 $(if ($onDark) { $dText } else { $body }) $mx $Deck.BodyTop ($right - $mx)); $script:BodyY = $Deck.BodyTop + 40 }
    $fy = $Deck.BodyBottom + 16
    [void](Box $mx $fy ($right - $mx) 1 $(if ($onDark) { $dLine } else { Mix $Light $Ink 0.12 }))
    if ($Deck.Footer) { [void](Tx $Deck.Footer 10.5 $mute $mx ($fy + 8) 500) }
    $no = $(if ($Deck.Pages) { '{0:00} / {1}' -f $script:page, $Deck.Pages } else { '{0:00}' -f $script:page })
    [void](Tx $no 10.5 $mute ($right - 100) ($fy + 8) 100 0 $Font 3)
}
function Cols([int]$count, $gap = -1) {
    if ($gap -lt 0) { $gap = $Deck.Gap }
    $cwid = ($W - 2 * $Deck.MarginX - $gap * ($count - 1)) / $count; $o = @()
    for ($i = 0; $i -lt $count; $i++) { $o += , @(($Deck.MarginX + $i * ($cwid + $gap)), $cwid) }
    , $o
}
function Card($x, $y, $bw, $bh, [string]$fill = '#FFFFFF', [bool]$shadow = $true) {
    $c = Box $x $y $bw $bh $fill $Deck.Radius; if ($shadow) { Glow $c (Mix $Accent '#FFFFFF' 0.5) 0.78 18 6 }; $c
}
function Rise($parts, $delay = 0.05) { if (-not $script:static) { Anim (Grp $parts) $RISE $AFTER 0.4 $delay } }

# Wide figure card: the figure at the left, name and description at the right.
function StatWide($x, $y, $bw, $bh, [string]$big, [string]$label, [string]$desc, [bool]$onDark = $false) {
    $parts = @((Card $x $y $bw $bh $(if ($onDark) { $dCard } else { '#FFFFFF' }) (-not $onDark)))
    $n = NoWrap (Tx $big 60 $(if ($onDark) { '#FFFFFF' } else { $Accent }) ($x + 26) $y 20 $heavyBold $heavy)
    if ([double]$n.Width -gt 186) { $n.TextFrame.TextRange.Font.Size = [single][math]::Floor(60 * 186 / [double]$n.Width) }
    $n.Top = [single]($y + ($bh - [double]$n.Height) / 2); $parts += $n
    $parts += (Tx $label 19 $(if ($onDark) { '#FFFFFF' } else { $Ink }) ($x + 226) ($y + $bh / 2 - 40) ($bw - 240) -1)
    $parts += (Tx $desc 15 $(if ($onDark) { $dText } else { $body }) ($x + 226) ($y + $bh / 2 - 10) ($bw - 240) 0 $Font 1 0 1.4)
    $parts
}
# Before -> after card. Returns card, label, before, arrow, after, [unit], description (Emph relies on that order).
function BA($x, $y, $bw, $bh, [string]$label, [string]$from, [string]$to, [string]$unit, [string]$desc, $toSize) {
    $parts = @((Card $x $y $bw $bh))
    $parts += (Tx $label 16 $Accent ($x + 26) ($y + 22) ($bw - 52) -1)
    $t = NoWrap (Tx $to $toSize $Accent 0 ($y + 52) 20 $heavyBold $heavy)
    $f = NoWrap (Tx $from ($toSize * 0.42) $gray ($x + 26) 0 20 $heavyBold $heavy)
    $a = NoWrap (Tx $ARROW ($toSize * 0.4) $Accent 0 0 20 -1)
    $base = [double]$t.Top + [double]$t.Height
    $f.Top = [single]($base - [double]$f.Height - $toSize * 0.13)
    $a.Left = [single]([double]$f.Left + [double]$f.Width + 10); $a.Top = [single]($base - [double]$a.Height - $toSize * 0.14)
    $t.Left = [single]([double]$a.Left + [double]$a.Width + 10)
    $parts += $f; $parts += $a; $parts += $t
    if ($unit) { $u = NoWrap (Tx $unit ($toSize * 0.36) $AccentDark ([double]$t.Left + [double]$t.Width + 5) 0 20 -1); $u.Top = [single]($base - [double]$u.Height - $toSize * 0.15); $parts += $u }
    $parts += (Tx $desc 15 $body ($x + 26) ($base + 10) ($bw - 52) 0 $Font 1 0 1.45)
    $parts
}
# Row card: number + head + text ('' as number = the text starts at the left edge). $ts = text size.
function RowCard($x, $y, $bw, $bh, [string]$no, [string]$head, [string]$text, $ts = 15) {
    $parts = @((Card $x $y $bw $bh)); $tx = $x + 22
    if ($no) { $parts += (Tx $no 22 $Accent ($x + 22) ($y + 16) 40 $heavyBold $heavy); $tx = $x + 66 }
    $parts += (Tx $head 19 $Ink $tx ($y + 18) ($x + $bw - $tx - 16) -1)
    $parts += (Tx $text $ts $body $tx ($y + 48) ($x + $bw - $tx - 16) 0 $Font 1 0 1.4)
    $parts
}
# Row card with a small before -> after figure at the left.
function RowBA($x, $y, $bw, $bh, [string]$from, [string]$to, [string]$head, [string]$text, $ts = 15) {
    $parts = @((Card $x $y $bw $bh))
    $t = NoWrap (Tx $to 38 $Accent 0 0 20 $heavyBold $heavy); $t.Top = [single]($y + ($bh - [double]$t.Height) / 2)
    $f = NoWrap (Tx $from 20 $gray ($x + 22) 0 20 $heavyBold $heavy); $base = [double]$t.Top + [double]$t.Height
    $f.Top = [single]($base - [double]$f.Height - 6)
    $a = NoWrap (Tx $ARROW 18 $Accent ([double]$f.Left + [double]$f.Width + 6) 0 20 -1); $a.Top = [single]($base - [double]$a.Height - 7)
    $t.Left = [single]([double]$a.Left + [double]$a.Width + 6)
    $parts += $f; $parts += $a; $parts += $t
    $parts += (Tx $head 19 $Ink ($x + 226) ($y + $bh / 2 - 34) ($bw - 240) -1)
    $parts += (Tx $text $ts $body ($x + 226) ($y + $bh / 2 - 4) ($bw - 240) 0 $Font 1 0 1.4)
    $parts
}
# Tall card: number (or an icon picture) + head + lines ('|' = line break).
function TCard($x, $y, $bw, $bh, [string]$no, [string]$head, [string]$text, [string]$icon = '', $noSize = 26) {
    $parts = @((Card $x $y $bw $bh))
    if ($icon) { $parts += (Pic $icon ($x + 14) ($y + 14) 100); $hy = $y + 124 }
    else { $parts += (Tx $no $noSize $Accent ($x + 22) ($y + 20) ($bw - 40) $heavyBold $heavy); $hy = $y + 20 + $noSize * 1.25 + 14 }
    $parts += (Tx $head 22 $Ink ($x + 20) $hy ($bw - 30) -1)
    $parts += (Tx $text 17 $body ($x + 20) ($hy + 42) ($bw - 30) 0 $Font 1 0 1.55)
    $parts
}
# Horizontal timeline: the line is drawn first, then the columns rise in turn. $rows = @(@(date, head, text), ...); the last stop is lit.
function Timeline($rows, $y0) {
    $n = $rows.Count; $mx = $Deck.MarginX; $span = $W - 2 * $mx; $cwid = $span / $n
    $track = Box $mx ($y0 - 1.5) $span 3 (Mix $Light $Accent 0.35); if (-not $script:static) { Anim $track $WIPE $AFTER 0.6 0 4 }
    for ($i = 0; $i -lt $n; $i++) {
        $r = $rows[$i]; $x = $mx + $i * $cwid; $last = ($i -eq $n - 1)
        $node = Disc ($x + 11) $y0 $(if ($last) { 11 } else { 8 }) $(if ($last) { $Accent } else { '#FFFFFF' })
        if (-not $last) { $node.Line.Visible = -1; $node.Line.ForeColor.RGB = (ConvertTo-Bgr $Accent); $node.Line.Weight = 3 }
        $parts = @($node)
        $parts += Pill $x ($y0 - 58) (Wd $r[0] 16) 32 $(if ($last) { $Accent } else { $tint }) $r[0] 16 $(if ($last) { '#FFFFFF' } else { $AccentDark })
        $parts += (Tx $r[1] 20 $Ink $x ($y0 + 24) ($cwid - 14) -1)
        $parts += (Tx $r[2] 16.5 $body $x ($y0 + 56) ($cwid - 14) 0 $Font 1 0 1.4)
        Rise $parts
    }
}
# A photo cropped to a rounded frame. $ax = 0 (keep the left of the photo) .. 1 (keep the right).
function Photo([string]$path, $x, $y, $bw, $bh, $ax = 0.5) {
    $p = Tag ($sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1)); $p.LockAspectRatio = -1
    $k = [math]::Max($bw / [double]$p.Width, $bh / [double]$p.Height); $p.Width = [single]([double]$p.Width * $k)
    $p.Left = [single]($x - ([double]$p.Width - $bw) * $ax); $p.Top = [single]($y - ([double]$p.Height - $bh) / 2)
    $r = Box $x $y $bw $bh '#FFFFFF' $Deck.Radius
    Tag (Merge-Shapes 3 @($p.Name, $r.Name))
}
# Emphasis in two steps (Morph). Build the slide with $script:static = $true (no entrance effects, nothing grouped), then call this:
# the slide is duplicated; on the first copy everything in $others and the BA's arrow / after / unit / description go, and the
# "before" figure stands alone at $toSize. The first copy advances by itself.
function Emph($ba, $others, $toSize, $hold = 1.3) {
    Add-MorphSlide 1.0
    $first = $ba[0].Parent
    foreach ($p in @($ba[3..($ba.Count - 1)]) + @($others)) { $p.Delete() }
    $r = $ba[2].TextFrame.TextRange; $r.Font.Size = [single]$toSize; $r.Font.Color.RGB = (ConvertTo-Bgr $Ink)
    $ba[2].Top = [single]([double]$ba[0].Top + 56)
    Auto $first $hold; $first.SlideShowTransition.EntryEffect = 3849
}
function Save-DeckAll([string]$out, [string]$previewDir = '', [string]$video = '') {
    $pres.SaveAs($out)
    if ($previewDir) {
        [void](New-Item -ItemType Directory -Force $previewDir)
        foreach ($q in $pres.Slides) { $q.Export((Join-Path $previewDir ('slide_{0:00}.png' -f $q.SlideIndex)), 'PNG', 1280, 720) }
    }
    if ($video) {
        if (Test-Path $video) { Remove-Item $video }
        $pres.CreateVideo($video, $true, 3, 720, 30, 85)
        while ($pres.CreateVideoStatus -in 1, 2) { Start-Sleep -Seconds 2 }
        if ($pres.CreateVideoStatus -ne 3) { throw "video export failed, status $($pres.CreateVideoStatus)" }
    }
    "saved $out ($($pres.Slides.Count) slides, $($script:page) pages)"
    Close-Deck
}
