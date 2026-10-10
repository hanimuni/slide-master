# Put rebuilt slides into a COPY of a deck: slide N of the copy is replaced by the slide(s) of another .pptx
# (a recipe's output). The original file is never opened for writing. A recipe that makes several slides (Morph)
# replaces one slide by all of them; later slide numbers shift, which is why replacements run from the back.
# What the old slide carried besides its content can come along:
#   -KeepRecurring  shapes that sit at the same place on at least half of the deck's slides (title band, logo, page
#                   number, footer line) are copied from the old slide onto the new one
#   -KeepTitle      the old slide's title text too (build the recipe WITHOUT its own title then, or there are two)
#   -Keep           more shapes by name: '3=Picture 4|TextBox 9' (names as deck_inspect lists them)
#   -MasterShapes   the new slide takes the old slide's layout and shows the master's shapes (logo, slide number)
# Copying goes through the clipboard: whatever the user had on it is replaced.
# Example: .\deck_replace.ps1 -Deck C:\work\draft.pptx -Replace '3=C:\work\s3.pptx','7=C:\work\s7.pptx' -KeepRecurring -KeepTitle -Out C:\work\draft_fx.pptx -PreviewDir C:\work\after
param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string[]]$Replace,     # 'slide number=path to the new .pptx'
    [switch]$KeepRecurring,
    [switch]$KeepTitle,
    [string[]]$Keep = @(),
    [switch]$MasterShapes,
    [Parameter(Mandatory)][string]$Out,
    [string]$PreviewDir = '',                     # one PNG per slide of the result
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Deck)) { throw "not found: $Deck" }
if ((Test-Path $Out) -and ((Resolve-Path $Out).Path -eq (Resolve-Path $Deck).Path)) { throw 'Out must be a different file: the original stays as it is' }
Assert-Output @($Out) $Force.IsPresent
$jobs = @($Replace | ForEach-Object { $kv = $_ -split '=', 2; if (-not (Test-Path $kv[1])) { throw "not found: $($kv[1])" }; [pscustomobject]@{ At = [int]$kv[0]; File = (Resolve-Path $kv[1]).Path } } | Sort-Object At -Descending)
$named = @{}; foreach ($k in $Keep) { $kv = $k -split '=', 2; $named[[int]$kv[0]] = @($kv[1] -split '\|') }
Copy-Item $Deck $Out -Force
function Spot($sh) { '{0}|{1}|{2}|{3}|{4}' -f $sh.Type, [math]::Round([double]$sh.Left / 2), [math]::Round([double]$sh.Top / 2), [math]::Round([double]$sh.Width / 2), [math]::Round([double]$sh.Height / 2) }

# Copy shapes from one slide to another through the clipboard, and make sure it is really those shapes that arrive:
# a copy that has not landed yet would paste whatever was on the clipboard before (seen with a page-sized picture).
function Carry($from, $to, [string[]]$names) {
    Add-Type -AssemblyName System.Windows.Forms
    $before = @($to.Shapes | ForEach-Object { $_.Id })
    for ($try = 1; ; $try++) {
        [System.Windows.Forms.Clipboard]::Clear()
        $from.Shapes.Range($names).Copy()
        for ($w = 0; $w -lt 50 -and -not [System.Windows.Forms.Clipboard]::GetDataObject().GetFormats().Count; $w++) { Start-Sleep -Milliseconds 100 }
        [void]$to.Shapes.Paste()
        $got = @($to.Shapes | Where-Object { $before -notcontains $_.Id })
        # same number of shapes, same kinds and sizes (names can change on arrival when the new slide already uses one)
        $want = @($names | ForEach-Object { $o = $from.Shapes($_); '{0}|{1}|{2}' -f $o.Type, [math]::Round([double]$o.Width), [math]::Round([double]$o.Height) } | Sort-Object)
        $have = @($got | ForEach-Object { '{0}|{1}|{2}' -f $_.Type, [math]::Round([double]$_.Width), [math]::Round([double]$_.Height) } | Sort-Object)
        if ($got.Count -eq $names.Count -and -not (Compare-Object $want $have)) { break }
        "copy attempt ${try}: wanted $($names.Count) shapes, got $($got.Count) ($(($got | ForEach-Object { $_.Name }) -join ', '))"
        foreach ($g in $got) { $g.Delete() }
        if ($try -ge 3) { throw "could not copy $($names -join ', ') from slide $($from.SlideIndex): the clipboard returned something else" }
        Start-Sleep -Seconds 1
    }
    # pasted shapes keep their place and come out on top - except page-sized ones (background picture, frame), which go behind the new content
    foreach ($q in @($got | Sort-Object { [double]$_.Width * [double]$_.Height })) { if ([double]$q.Width * [double]$q.Height -gt 0.5 * $dw * $dh) { $q.ZOrder(1) } }
}

try {
    $script:app = New-Object -ComObject PowerPoint.Application
    $script:pres = $app.Presentations.Open((Resolve-Path $Out).Path, 0, 0, 0)
    $dw = [double]$pres.PageSetup.SlideWidth; $dh = [double]$pres.PageSetup.SlideHeight
    $seen = @{}; foreach ($s in $pres.Slides) { foreach ($k in @($s.Shapes | ForEach-Object { Spot $_ } | Sort-Object -Unique)) { $seen[$k]++ } }
    $need = [math]::Max(2, [math]::Ceiling($pres.Slides.Count / 2))                    # "recurring" = same kind, place and size on at least half of the slides
    foreach ($j in $jobs) {
        if ($j.At -lt 1 -or $j.At -gt $pres.Slides.Count) { throw "the deck has no slide $($j.At)" }
        $old = $pres.Slides($j.At); $carry = @()
        if ($KeepRecurring) { $carry += @($old.Shapes | Where-Object { $seen[(Spot $_)] -ge $need } | ForEach-Object { $_.Name }) }
        if ($KeepTitle) {                                                              # the title placeholder, else the largest text in the top quarter
            $tt = $null; try { if ($old.Shapes.HasTitle) { $tt = $old.Shapes.Title } } catch {}
            if (-not $tt) { $tt = @($old.Shapes | Where-Object { $_.HasTextFrame -and $_.TextFrame.HasText -and [double]$_.Top -lt $dh * 0.25 } | Sort-Object { - [double]$_.TextFrame.TextRange.Font.Size })[0] }
            if ($tt) { $carry += $tt.Name }
        }
        if ($named.ContainsKey($j.At)) { $carry += $named[$j.At] }
        [string[]]$carry = @($carry | Sort-Object -Unique | ForEach-Object { "$_" })      # Shapes.Range wants plain strings
        $src = $app.Presentations.Open($j.File, -1, -1, 0)
        $count = $src.Slides.Count; $sw = [double]$src.PageSetup.SlideWidth; $sh2 = [double]$src.PageSetup.SlideHeight
        $backs = @($src.Slides | ForEach-Object { $_.Background.Fill.ForeColor.RGB }); $moves = @($src.Slides | ForEach-Object { @{ E = $_.SlideShowTransition.EntryEffect; D = $_.SlideShowTransition.Duration } })
        $src.Close()
        if ([math]::Abs($sw / $sh2 - $dw / $dh) -gt 0.01) { "warning: slide $($j.At) - the new slide is $sw x $sh2, the deck is $dw x $dh (shapes will not fill the page)" }
        [void]$pres.Slides.InsertFromFile($j.File, $j.At, 1, $count)                   # lands after the old slide
        for ($k = 0; $k -lt $count; $k++) {                                            # InsertFromFile takes the deck's master: put back what the recipe set on the slide itself
            $ns = $pres.Slides($j.At + 1 + $k)
            if ($MasterShapes) {
                $ns.CustomLayout = $old.CustomLayout; $ns.DisplayMasterShapes = -1
                foreach ($ph in @($ns.Shapes | Where-Object { $_.Type -eq 14 -and $_.HasTextFrame -and -not $_.TextFrame.HasText })) { $ph.Delete() }      # empty "click to add" boxes the layout brought
                try { $ns.HeadersFooters.SlideNumber.Visible = $old.HeadersFooters.SlideNumber.Visible; $ns.HeadersFooters.Footer.Visible = $old.HeadersFooters.Footer.Visible } catch {}
            } else { $ns.DisplayMasterShapes = 0 }
            $ns.FollowMasterBackground = 0; $ns.Background.Fill.Solid(); $ns.Background.Fill.ForeColor.RGB = $backs[$k]
            if ($k -gt 0) { $ns.SlideShowTransition.EntryEffect = $moves[$k].E; $ns.SlideShowTransition.Duration = $moves[$k].D }
            if ($carry.Count) { Carry $old $ns $carry }
        }
        $old.Delete()
        "slide $($j.At) <- $($j.File) ($count slide(s))" + $(if ($carry.Count) { "; kept: $($carry -join ', ')" } else { '' })
    }
    $pres.Save()
    if ($PreviewDir) {
        [void](New-Item -ItemType Directory -Force $PreviewDir)
        foreach ($s in $pres.Slides) { $s.Export((Join-Path (Resolve-Path $PreviewDir).Path ('slide_{0:00}.png' -f $s.SlideIndex)), 'PNG', 1280, [int](1280 * $dh / $dw)) }
    }
    "saved $Out ($($pres.Slides.Count) slides)"
    $pres.Close(); $script:pres = $null; if ($app.Presentations.Count -eq 0) { $app.Quit() }
} catch { Close-Deck; throw }
