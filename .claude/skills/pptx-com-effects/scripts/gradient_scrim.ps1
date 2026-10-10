# Text over a photo without darkening the whole picture: a black gradient that fades to fully transparent,
# so it has no visible edge. Stops as in the tutorial: 100% / 70% / 20% transparency.
# Example: .\gradient_scrim.ps1 -Image C:\work\photo.jpg -Title 'TITLE' -Body 'one line' -Out C:\work\scrim.pptx -Preview C:\work\scrim.png
# Vignette (radial, dark edges, eye drawn to the centre: 70% / 25% / 0%): add -Side Center
param(
    [Parameter(Mandatory)][string]$Image,
    [string]$Title = '',
    [string]$Body = '',                           # '|' = line break
    [ValidateSet('Bottom', 'Top', 'Left', 'Right', 'Center')][string]$Side = 'Bottom',   # where the text sits; Center = vignette
    [double]$Cover = 0.65,                        # share of the slide the gradient spans; smaller keeps more of the photo bright
    [string]$ScrimColor = '#000000',
    [double[]]$Transparency = @(),                # clear end, middle, dark end; default 1,0.7,0.2 (Center: 0.7,0.25,0)
    [string]$Font = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
if (-not (Test-Path $Image)) { throw "not found: $Image" }
if ($Transparency.Count -eq 0) { $Transparency = if ($Side -eq 'Center') { 0.7, 0.25, 0.0 } else { 1.0, 0.7, 0.2 } }
if ($Transparency.Count -ne 3) { throw '-Transparency needs three values' }
Assert-Output @($Out, $Preview) $Force.IsPresent

try {
    New-Deck '#000000'
    $bg = Add-Cover $Image; $bg.Name = 'Photo'
    $box = switch ($Side) {
        'Bottom' { @(0, ($H * (1 - $Cover)), $W, ($H * $Cover), 90) }
        'Top'    { @(0, 0, $W, ($H * $Cover), 270) }
        'Left'   { @(0, 0, ($W * $Cover), $H, 180) }
        'Right'  { @(($W * (1 - $Cover)), 0, ($W * $Cover), $H, 0) }
        'Center' { @(0, 0, $W, $H, -1) }                              # whole slide, radial
    }
    $r = New-Rect ([single]$box[0]) ([single]$box[1]) ([single]$box[2]) ([single]$box[3]); $r.Name = 'Scrim'
    $r.Line.Visible = 0
    $c = ConvertTo-Bgr $ScrimColor
    if ($Side -eq 'Center') { $r.Fill.TwoColorGradient(7, 1) }       # msoGradientFromCenter
    else { $r.Fill.TwoColorGradient(1, 1) }                         # horizontal style; the angle below decides the direction
    $stops = $r.Fill.GradientStops
    $stops.Insert($c, 0.5, $Transparency[1], 2)
    $pos = 0.0, 0.5, 1.0
    for ($i = 1; $i -le 3; $i++) { $s = $stops.Item($i); $s.Color.RGB = $c; $s.Position = $pos[$i - 1]; $s.Transparency = $Transparency[$i - 1] }
    if ($Side -ne 'Center') { $r.Fill.GradientAngle = $box[4] }

    $pad = 56
    $align = if ($Side -eq 'Right') { 3 } elseif ($Side -eq 'Center') { 2 } else { 1 }
    $items = @()
    if ($Title) { $t = New-FitText $Title 54 0xFFFFFF $Font ($W * 0.6); $t.Name = 'Title'; $items += $t }
    if ($Body) { $b = New-FitText $Body 20 0xFFFFFF 'Arial' ($W * 0.6); $b.Name = 'Body'; $items += $b }
    $total = 0.0; foreach ($it in $items) { $total += [double]$it.Height + 8 }
    $y = if ($Side -eq 'Top') { $pad } elseif ($Side -eq 'Bottom') { $H - $pad - $total } else { ($H - $total) / 2 }
    foreach ($it in $items) {
        $it.TextFrame.TextRange.ParagraphFormat.Alignment = $align
        $x = if ($Side -eq 'Right') { $W - $pad - [double]$it.Width } elseif ($Side -eq 'Center') { ($W - [double]$it.Width) / 2 } else { $pad }
        $it.Left = [single]$x; $it.Top = [single]$y
        $y += [double]$it.Height + 8
    }
    $tmp = ''
    if ($Side -eq 'Center') {
        # COM can only make the rectangular "from centre" gradient (path="rect"); the tutorial's radial type is
        # path="circle" and has no object-model switch. Save, patch that one attribute in the slide XML, reopen.
        $tmp = Join-Path $env:TEMP ('scrim_' + [guid]::NewGuid().ToString('N') + '.pptx')
        $pres.SaveAs($tmp); $pres.Close()
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::Open($tmp, 'Update')
        $entry = $zip.GetEntry('ppt/slides/slide1.xml')
        $reader = New-Object IO.StreamReader($entry.Open()); $xml = $reader.ReadToEnd(); $reader.Close()
        if ($xml -notmatch '<a:path path="rect">') { $zip.Dispose(); throw 'vignette gradient not found in the slide XML' }
        $entry.Delete()
        $writer = New-Object IO.StreamWriter($zip.CreateEntry('ppt/slides/slide1.xml').Open())
        $writer.Write($xml.Replace('<a:path path="rect">', '<a:path path="circle">')); $writer.Close(); $zip.Dispose()
        $script:pres = $app.Presentations.Open($tmp); $script:sl = $pres.Slides(1)
    }
    Save-Deck $Out $Preview 3
    if ($tmp) { Remove-Item $tmp -ErrorAction SilentlyContinue }
    "saved $Out"
} catch { Close-Deck; throw }
