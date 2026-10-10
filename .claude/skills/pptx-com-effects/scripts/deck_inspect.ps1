# Read an existing deck before proposing effects for it: one PNG per slide plus deck.json with what must be kept
# (slide size, theme colours and fonts, the colours and fonts actually used) and what each slide holds (texts, pictures,
# charts, tables, and which shapes recur from slide to slide). The deck is opened read-only and is not changed.
# Example: .\deck_inspect.ps1 -Deck C:\work\draft.pptx -OutDir C:\work\draft_inspect
param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$OutDir,
    [int]$Width = 1280
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Deck)) { throw "not found: $Deck" }
[void](New-Item -ItemType Directory -Force $OutDir)
function Hex([int]$bgr) { '#{0:X2}{1:X2}{2:X2}' -f ($bgr -band 255), (($bgr -shr 8) -band 255), (($bgr -shr 16) -band 255) }
function Spot($sh) { '{0}|{1}|{2}|{3}|{4}' -f $sh.Type, [math]::Round([double]$sh.Left / 2), [math]::Round([double]$sh.Top / 2), [math]::Round([double]$sh.Width / 2), [math]::Round([double]$sh.Height / 2) }
function Top($table, [int]$n) { @($table.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First $n | ForEach-Object { "$($_.Key) x$($_.Value)" }) }

try {
    $script:app = New-Object -ComObject PowerPoint.Application
    $script:pres = $app.Presentations.Open((Resolve-Path $Deck).Path, -1, -1, 0)      # read-only, untitled copy, no window
    $w = [double]$pres.PageSetup.SlideWidth; $h = [double]$pres.PageSetup.SlideHeight
    $theme = $pres.SlideMaster.Theme; $fonts = @{}; $inks = @{}; $fills = @{}; $slides = @()
    # shapes at the same place on at least half of the slides = the deck's furniture (title band, logo, page number); deck_replace.ps1 -KeepRecurring uses the same rule
    $seen = @{}; foreach ($s in $pres.Slides) { foreach ($k in @($s.Shapes | ForEach-Object { Spot $_ } | Sort-Object -Unique)) { $seen[$k]++ } }
    $need = [math]::Max(2, [math]::Ceiling($pres.Slides.Count / 2))
    foreach ($s in $pres.Slides) {
        $png = Join-Path (Resolve-Path $OutDir).Path ('slide_{0:00}.png' -f $s.SlideIndex); $s.Export($png, 'PNG', $Width, [int]($Width * $h / $w))
        $texts = @(); $rec = @(); $pics = 0; $charts = 0; $tables = 0; $title = ''
        foreach ($sh in $s.Shapes) {
            if ($sh.Type -eq 13 -or $sh.Type -eq 11) { $pics++ }                      # picture, linked picture
            if ($seen[(Spot $sh)] -ge $need) { $rec += $sh.Name }
            if ($sh.HasChart) { $charts++ }; if ($sh.HasTable) { $tables++ }
            if ($sh.Fill.Visible -and $sh.Fill.Type -eq 1 -and $sh.Type -ne 13) { $fills[(Hex $sh.Fill.ForeColor.RGB)]++ }
            if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
                $r = $sh.TextFrame.TextRange; $txt = "$($r.Text)".Trim() -replace '[\r\n\v]+', ' / '
                if ($txt.Length -gt 200) { $txt = $txt.Substring(0, 200) + '...' }
                $texts += [ordered]@{ text = $txt; size = [double]$r.Font.Size; font = $r.Font.NameFarEast; name = $sh.Name; left = [int]$sh.Left; top = [int]$sh.Top; width = [int]$sh.Width }
                $fonts["$($r.Font.Name) / $($r.Font.NameFarEast)"]++; $inks[(Hex $r.Font.Color.RGB)]++
            }
        }
        try { if ($s.Shapes.HasTitle) { $title = "$($s.Shapes.Title.TextFrame.TextRange.Text)".Trim() } } catch {}
        if (-not $title -and $texts.Count) { $title = @($texts | Sort-Object { - $_.size })[0].text }       # no title placeholder: the largest text
        $back = try { Hex $s.Background.Fill.ForeColor.RGB } catch { '' }
        $slides += [ordered]@{ index = $s.SlideIndex; id = $s.SlideID; layout = $s.CustomLayout.Name; title = $title; background = $back; image = $png
            shapes = $s.Shapes.Count; pictures = $pics; charts = $charts; tables = $tables; transition = [int]$s.SlideShowTransition.EntryEffect; recurring = $rec; texts = $texts }
    }
    $scheme = @(); for ($i = 1; $i -le 12; $i++) { $scheme += Hex $theme.ThemeColorScheme.Colors($i).RGB }
    $info = [ordered]@{ deck = (Resolve-Path $Deck).Path; slideWidth = $w; slideHeight = $h; slideCount = $pres.Slides.Count
        themeColors = $scheme                                                          # dark1, light1, dark2, light2, accent1-6, hyperlink, followed
        themeFonts = [ordered]@{ headingLatin = $theme.ThemeFontScheme.MajorFont(1).Name; headingEastAsian = $theme.ThemeFontScheme.MajorFont(2).Name
            bodyLatin = $theme.ThemeFontScheme.MinorFont(1).Name; bodyEastAsian = $theme.ThemeFontScheme.MinorFont(2).Name }
        masterShapes = @($pres.SlideMaster.Shapes | Where-Object { $_.Type -ne 14 } | ForEach-Object { "$($_.Name) (type $($_.Type))" })      # drawn by the master on every slide: deck_replace.ps1 -MasterShapes keeps them
        usedFonts = @(Top $fonts 6); usedTextColors = @(Top $inks 6); usedFillColors = @(Top $fills 8); slides = $slides }
    $json = Join-Path (Resolve-Path $OutDir).Path 'deck.json'
    [IO.File]::WriteAllText($json, ($info | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
    "$($pres.Slides.Count) slides ($w x $h) -> $json"
    $pres.Close(); $script:pres = $null; if ($app.Presentations.Count -eq 0) { $app.Quit() }
} catch { Close-Deck; throw }
