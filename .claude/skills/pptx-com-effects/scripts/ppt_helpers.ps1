# Shared helpers for building slides through PowerPoint COM. Dot-source from a recipe script.
# ASCII only: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.
$ErrorActionPreference = 'Stop'
$script:W = 960; $script:H = 540

function ConvertTo-Bgr([string]$hex) {           # '#D50032' -> COM RGB (which is BGR)
    $h = $hex.TrimStart('#')
    [Convert]::ToInt32($h.Substring(0, 2), 16) + [Convert]::ToInt32($h.Substring(2, 2), 16) * 256 + [Convert]::ToInt32($h.Substring(4, 2), 16) * 65536
}

function Assert-Output([string[]]$paths, [bool]$force) {
    foreach ($p in $paths) {
        if (-not $p -or -not (Test-Path $p)) { continue }
        if (-not $force) { throw "already exists (pass -Force to replace): $p" }
        Remove-Item $p
    }
}

function New-Deck([string]$backgroundHex) {
    # A PowerPoint left over from the previous build is still exiting; attaching to it gives random
    # "Specified cast is not valid" errors. Give it a moment to go (a PowerPoint the user keeps open stays).
    for ($w = 0; $w -lt 10 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $w++) { Start-Sleep -Seconds 1 }
    for ($try = 1; ; $try++) {           # a PowerPoint that is still shutting down rejects the first calls
        try {
            $script:app = New-Object -ComObject PowerPoint.Application
            $script:pres = $script:app.Presentations.Add(-1)
            break
        } catch { if ($try -ge 5) { throw }; Start-Sleep -Seconds 2 }
    }
    $script:pres.PageSetup.SlideWidth = $script:W; $script:pres.PageSetup.SlideHeight = $script:H
    $script:sl = $script:pres.Slides.Add(1, 12)                      # ppLayoutBlank
    $script:sl.FollowMasterBackground = 0
    $script:sl.Background.Fill.Solid(); $script:sl.Background.Fill.ForeColor.RGB = (ConvertTo-Bgr $backgroundHex)
}

# 1 union, 3 intersect, 4 subtract. The first name is the primary shape (it survives / gives the formatting).
function Merge-Shapes([int]$mode, [string[]]$names) {
    $before = @($script:sl.Shapes | ForEach-Object { $_.Id })
    $script:sl.Shapes.Range($names).MergeShapes($mode, $script:sl.Shapes($names[0]))
    # the result is not always the top-most shape (it can take the primary's z-position), so find it by id
    $new = @($script:sl.Shapes | Where-Object { $before -notcontains $_.Id })
    if ($new.Count -eq 1) { return $new[0] }
    $script:sl.Shapes($script:sl.Shapes.Count)
}

function New-Rect($l, $t, $w, $h) { $script:sl.Shapes.AddShape(1, $l, $t, $w, $h) }

function New-Text([string]$text, $size, [int]$colorBgr, [string]$font) {
    $t = $script:sl.Shapes.AddTextbox(1, 0, 0, 100, 100)
    $t.TextFrame.WordWrap = 0; $t.TextFrame.AutoSize = 1
    $t.TextFrame.MarginLeft = 0; $t.TextFrame.MarginRight = 0
    $t.TextFrame.MarginTop = 0; $t.TextFrame.MarginBottom = 0
    $r = $t.TextFrame.TextRange
    # Font.Name only sets the Latin face; Korean glyphs take NameFarEast (without it they fall back to the theme font)
    $r.Text = $text; $r.Font.Name = $font; $r.Font.NameFarEast = $font; $r.Font.Size = [single]$size; $r.Font.Color.RGB = $colorBgr
    $t
}

# Glyph outlines of a text box as a real shape: exact bounds, usable as outline or as a clipping shape.
function New-GlyphShape($text) {
    $dup = $text.Duplicate(); $dup.Left = $text.Left; $dup.Top = $text.Top
    $dup.Name = "GlyphSource$($dup.Id)"                       # a copy keeps a custom name; Range() would then grab the original
    $probe = New-Rect -20 -20 ($script:W + 40) ($script:H + 40)
    Merge-Shapes 3 @($probe.Name, $dup.Name)
}

# Text box whose font shrinks until it fits $maxW. '|' in the text starts a new line.
function New-FitText([string]$text, $size, [int]$colorBgr, [string]$font, $maxW) {
    $t = New-Text ($text -replace '\|', "`r") $size $colorBgr $font
    $limit = [double]$maxW
    if ([double]$t.Width -gt $limit) { $t.TextFrame.TextRange.Font.Size = [single][math]::Floor([double]$size * $limit / [double]$t.Width) }
    $t
}

function Set-Center($shape) {
    $shape.Left = ($script:W - [double]$shape.Width) / 2; $shape.Top = ($script:H - [double]$shape.Height) / 2
}

# Picture scaled to cover the whole slide (overflow is clipped by the slide edge).
function Add-Cover([string]$path) {
    $p = $script:sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1); $p.LockAspectRatio = -1
    $p.Width = $script:W
    if ([double]$p.Height -lt $script:H) { $p.Height = $script:H }
    Set-Center $p
    $p
}

# Duplicate the current slide, give the copy a Morph transition and make it the current slide.
# Morph pairs shapes by name; the copy keeps every name, so just move/resize/recolor shapes on either slide.
function Add-MorphSlide([double]$duration) {
    $i = $script:sl.SlideIndex
    [void]$script:sl.Duplicate()
    $script:sl = $script:pres.Slides($i + 1)
    $script:sl.SlideShowTransition.EntryEffect = 3954                # ppEffectMorphByObject
    $script:sl.SlideShowTransition.Duration = $duration
}

function Get-Box($shape) { @{ L = [double]$shape.Left; T = [double]$shape.Top; W = [double]$shape.Width; H = [double]$shape.Height } }

function Add-Grid([string]$hex) {
    $c = ConvertTo-Bgr $hex
    for ($x = 0; $x -le $script:W; $x += 40) { $l = $script:sl.Shapes.AddLine($x, 0, $x, $script:H); $l.Line.ForeColor.RGB = $c; $l.Line.Weight = 0.75 }
    for ($y = 20; $y -le $script:H; $y += 40) { $l = $script:sl.Shapes.AddLine(0, $y, $script:W, $y); $l.Line.ForeColor.RGB = $c; $l.Line.Weight = 0.75 }
}

# Save, optionally render a preview, then close. Preview *.png = still of the last slide (static effects);
# anything else = slide show as mp4, $seconds per slide (the only way to check motion).
function Save-Deck([string]$out, [string]$preview, [int]$seconds) {
    $script:pres.SaveAs($out)
    if ($preview -like '*.png') {
        $script:pres.Slides($script:pres.Slides.Count).Export($preview, 'PNG', 1920, 1080)
    } elseif ($preview -like '*.gif') {
        # File > Export > Create an Animated GIF. Each slide shows for its "advance after" time (1 s when none is set);
        # transitions such as Morph are rendered. The export runs in the background: wait until the file stops growing.
        $script:pres.SaveCopyAs($preview, 40)                        # ppSaveAsAnimatedGIF
        $size = -1
        for ($i = 0; $i -lt 120; $i++) {
            Start-Sleep -Seconds 1
            $now = if (Test-Path $preview) { (Get-Item $preview).Length } else { 0 }
            if ($now -gt 0 -and $now -eq $size) { break }
            $size = $now
        }
        if (-not $size) { throw "gif export produced nothing: $preview" }
    } elseif ($preview) {
        $script:pres.CreateVideo($preview, $true, $seconds, 1080, 30, 85)
        while ($script:pres.CreateVideoStatus -in 1, 2) { Start-Sleep -Seconds 1 }
        if ($script:pres.CreateVideoStatus -ne 3) { throw "video export failed, status $($script:pres.CreateVideoStatus)" }
    }
    Close-Deck
}

function Close-Deck {
    try { if ($script:pres) { $script:pres.Saved = -1; $script:pres.Close() } } catch {}
    $script:pres = $null
    if (-not $script:app) { return }
    $quit = $false
    try { if ($script:app.Presentations.Count -eq 0) { $script:app.Quit(); $quit = $true } } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:app); $script:app = $null
    # PowerPoint only exits once every COM reference is gone, i.e. when this PowerShell process ends.
    # Run one build per PowerShell process: a second build in the same session attaches to the
    # half-closed instance and fails with random "Specified cast is not valid" errors.
}
