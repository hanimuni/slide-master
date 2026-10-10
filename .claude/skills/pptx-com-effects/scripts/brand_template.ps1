# In-house template from one graphic image: slide 1 = cover (image full bleed, small subtitle and bold title in
# the top-left where the eye lands first, contact details on a white bar at the bottom), slide 2 = body page
# (only the top strip of the same image is kept as a header, title on the left, logo on the right).
# A logo that comes with a flat background gets it removed with Picture Format > Color > Set Transparent Color.
# Example: .\brand_template.ps1 -Image C:\work\graphic.jpg -Sub 'Mobile coupon service' -Title 'SERVICE GUIDE 2026' `
#            -Contact 'Name | Team|010-0000-0000|mail@example.com|www.example.com' -BodyTitle '01. New service' -Out C:\work\tpl.pptx -Preview C:\work\tpl.png
param(
    [Parameter(Mandatory)][string]$Image,         # dark graphic works best (white text sits on it)
    [string]$Title = '',
    [string]$Sub = '',
    [string]$Logo = '',                           # optional picture
    [string]$LogoTransparent = '',                # hex colour of the logo's flat background to knock out, e.g. '#000000'
    [string]$Contact = '',                        # 'a|b|c' - spread across the bottom bar of the cover
    [string]$BodyTitle = 'Title',
    [string]$Color = '#FFFFFF',
    [string]$Font = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',                        # *.png = picture of the body slide (the cover goes to *_cover.png)
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
foreach ($p in $Image, $Logo) { if ($p -and -not (Test-Path $p)) { throw "not found: $p" } }
Assert-Output @($Out, $Preview) $Force.IsPresent

function Add-Logo([double]$right, [double]$cy, [double]$height) {
    if (-not $Logo) { return }
    $l = $sl.Shapes.AddPicture($Logo, 0, -1, 0, 0, -1, -1); $l.LockAspectRatio = -1; $l.Height = [single]$height; $l.Name = 'Logo'
    if ($LogoTransparent) { $l.PictureFormat.TransparentBackground = -1; $l.PictureFormat.TransparencyColor = (ConvertTo-Bgr $LogoTransparent) }
    $l.Left = [single]($right - [double]$l.Width); $l.Top = [single]($cy - $height / 2)
}

try {
    New-Deck '#FFFFFF'
    $c = ConvertTo-Bgr $Color; $ink = ConvertTo-Bgr '#1B1B1B'; $pad = 56
    # --- cover ---
    $bg = Add-Cover $Image; $bg.Name = 'Graphic'
    $y = 70
    if ($Sub) { $s = New-FitText $Sub 20 $c 'Arial' ($W * 0.7); $s.Left = [single]$pad; $s.Top = [single]$y; $s.Name = 'Sub'; $y += [double]$s.Height + 6 }
    if ($Title) { $t = New-FitText $Title 28 $c $Font ($W * 0.7); $t.Left = [single]$pad; $t.Top = [single]$y; $t.Name = 'Title' }
    Add-Logo ($W - $pad) 92 34
    if ($Contact) {
        $barH = 44
        $bar = New-Rect 0 ([single]($H - $barH)) $W $barH; $bar.Fill.ForeColor.RGB = 0xFFFFFF; $bar.Line.Visible = 0; $bar.Name = 'ContactBar'
        $parts = $Contact -split '\|'; $slot = ($W - 2 * $pad) / $parts.Count
        for ($i = 0; $i -lt $parts.Count; $i++) {
            $p = New-Text $parts[$i] 11 $ink 'Arial'
            $p.Left = [single]($pad + $slot * $i); $p.Top = [single]($H - $barH + ($barH - [double]$p.Height) / 2)
        }
    }
    if ($Preview -like '*.png') { $sl.Export(($Preview -replace '\.png$', '_cover.png'), 'PNG', 1920, 1080) }
    # --- body page: keep only the top strip of the same picture ---
    $script:sl = $pres.Slides.Add(2, 12)
    $strip = 64
    $pic = Add-Cover $Image; $pic.Top = 0; $pic.Name = 'HeaderSource'
    $cut = New-Rect 0 0 $W $strip; $cut.Name = 'HeaderCut'
    $head = Merge-Shapes 3 @('HeaderSource', 'HeaderCut'); $head.Name = 'Header'      # picture first: the result stays a picture
    $bt = New-FitText $BodyTitle 20 $c $Font ($W * 0.7); $bt.Left = [single]$pad; $bt.Top = [single](($strip - [double]$bt.Height) / 2); $bt.Name = 'BodyTitle'
    Add-Logo ($W - $pad) ($strip / 2) 24
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
