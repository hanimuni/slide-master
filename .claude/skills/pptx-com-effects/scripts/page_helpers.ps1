# Drawing helpers shared by design_page.ps1 and info_page.ps1 (dot-source after ppt_helpers.ps1; they draw on $sl).
# shape with fill / outline in one call ($fill or $line = $null turns that part off)
function Box([int]$type, $l, $t, $w, $h, $fill, $line = $null, $weight = 1.5) {
    $s = $sl.Shapes.AddShape($type, [single]$l, [single]$t, [single]$w, [single]$h)
    if ($null -eq $fill) { $s.Fill.Visible = 0 } else { $s.Fill.ForeColor.RGB = [int]$fill }
    if ($null -eq $line) { $s.Line.Visible = 0 } else { $s.Line.ForeColor.RGB = [int]$line; $s.Line.Weight = [single]$weight }
    $s
}
function Put($t, $l, $top) { $t.Left = [single]$l; $t.Top = [single]$top; $t }
function Mid($t, $cx, $cy) { $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2; $t.Left = [single]($cx - [double]$t.Width / 2); $t.Top = [single]($cy - [double]$t.Height / 2); $t }
function Line($x1, $y1, $x2, $y2, $rgb, $weight) { $l = $sl.Shapes.AddLine([single]$x1, [single]$y1, [single]$x2, [single]$y2); $l.Line.ForeColor.RGB = [int]$rgb; $l.Line.Weight = [single]$weight; $l }
function Poly($pts, $fill, $line, $weight) {      # closed polygon from @(@(x,y),...)
    $fb = $sl.Shapes.BuildFreeform(0, [single]$pts[0][0], [single]$pts[0][1])
    foreach ($q in @($pts[1..($pts.Count - 1)]) + , $pts[0]) { $fb.AddNodes(0, 0, [single]$q[0], [single]$q[1]) }
    $s = $fb.ConvertToShape()
    if ($null -eq $fill) { $s.Fill.Visible = 0 } else { $s.Fill.ForeColor.RGB = [int]$fill }
    if ($null -eq $line) { $s.Line.Visible = 0 } else { $s.Line.ForeColor.RGB = [int]$line; $s.Line.Weight = [single]$weight }
    $s
}
function Photo([string]$path, $l, $t, $w, $h, [int]$type = 1) {   # picture cut to a shape, never stretched
    $script:shots++; $f = $sl.Shapes.AddShape($type, [single]$l, [single]$t, [single]$w, [single]$h); $f.Name = "F$script:shots"
    $pic = $sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1); $pic.LockAspectRatio = -1; $pic.Name = "S$script:shots"
    $pic.Width = [single]$w; if ([double]$pic.Height -lt $h) { $pic.Height = [single]$h }
    $pic.Left = [single]($l + ($w - [double]$pic.Width) / 2); $pic.Top = [single]($t + ($h - [double]$pic.Height) / 2)
    $r = Merge-Shapes 3 @("S$script:shots", "F$script:shots"); $r.Name = "Photo$script:shots"; $r
}
function Hard($s, $d) { $sh = $s.Shadow; $sh.Visible = -1; $sh.ForeColor.RGB = 0; $sh.Transparency = 0; $sh.Blur = 0; $sh.OffsetX = [single]$d; $sh.OffsetY = [single]$d }
function Round($l, $t, $w, $h, $fill, $adj = 0.1, $line = $null, $weight = 1.25) { $s = Box 5 $l $t $w $h $fill $line $weight; $s.Adjustments.Item(1) = [single]$adj; $s }
function Pic([string]$path, $height) { $q = $sl.Shapes.AddPicture($path, 0, -1, 0, 0, -1, -1); $q.LockAspectRatio = -1; $q.Height = [single]$height; $q }
function Veil($hex, $alpha) { $v = Box 1 0 0 $W $H (ConvertTo-Bgr $hex); $v.Fill.Transparency = [single](1 - $alpha); $v }
function Lines($parts) { ($parts | Where-Object { $_ }) -join '|' }
$script:shots = 0
