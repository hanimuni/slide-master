# Mouse-over maze game for an ice-breaker: follow the light path with the pointer from START to GOAL.
# Touching the dark wall jumps to a "try again" slide; reaching GOAL jumps to the "clear" slide.
# How it works: the wall is one big shape with a Mouse Over hyperlink; the path shapes lie on top of it and have no
# action, so the pointer is safe while it stays on them. Clicks do not advance the show (Advance on click is off).
# Works only in the slide show, with a real mouse.
# Example: .\hover_maze.ps1 -Title 'MAZE' -Out C:\work\maze.pptx -Preview C:\work\maze.png
param(
    [string]$Title = 'MAZE',
    [string]$Start = 'START',
    [string]$Goal = 'GOAL',
    [string]$Fail = 'TRY AGAIN',
    [string]$Clear = 'CLEAR!',
    [string]$Points = '0.08,0.8 0.08,0.25 0.36,0.25 0.36,0.72 0.64,0.72 0.64,0.3 0.9,0.3',   # corners of the path, as fractions of the slide
    [double]$Width = 46,                          # corridor width in points; narrower = harder
    [string]$Wall = '#16181D',
    [string]$Path = '#F4F3EF',
    [string]$Accent = '#FF4D4D',
    [string]$Font = 'Arial Black',
    [Parameter(Mandatory)][string]$Out,
    [string]$Preview = '',
    [switch]$Force
)
. (Join-Path $PSScriptRoot 'ppt_helpers.ps1')
Assert-Output @($Out, $Preview) $Force.IsPresent

function Page([int]$index, [string]$hex) {
    $script:sl = $pres.Slides.Add($index, 12); $sl.FollowMasterBackground = 0; $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = (ConvertTo-Bgr $hex)
}
function Link($shape, [int]$trigger, [int]$index) { $a = $shape.ActionSettings($trigger); $a.Action = 7; $a.Hyperlink.SubAddress = "$($pres.Slides($index).SlideID),$index," }   # 1 click, 2 mouse over
function Centre($t, $cy) { $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2; $t.Left = [single](($W - [double]$t.Width) / 2); $t.Top = [single]($cy - [double]$t.Height / 2) }

try {
    New-Deck $Wall                                                   # slide 1: start screen
    $light = ConvertTo-Bgr $Path; $acc = ConvertTo-Bgr $Accent; $dark = ConvertTo-Bgr $Wall
    $pts = @($Points -split '\s+' | Where-Object { $_ } | ForEach-Object { $q = $_ -split ','; , @(($W * [double]$q[0]), ($H * [double]$q[1])) }); $hw = $Width / 2
    # the button that opens the maze sits exactly where the start pad will be: the pointer is then already on the path
    # (anywhere else it would be over the wall, and the first mouse move would end the game)
    function Starter([string]$name) { $b = $sl.Shapes.AddShape(9, [single]($pts[0][0] - 34), [single]($pts[0][1] - 34), 68, 68); $b.Fill.ForeColor.RGB = $light; $b.Line.Visible = 0; $b.Name = $name
        $b.TextFrame.TextRange.Text = $Start; $b.TextFrame.TextRange.Font.Name = $Font; $b.TextFrame.TextRange.Font.Size = 11; $b.TextFrame.TextRange.Font.Color.RGB = $dark; $b.TextFrame.WordWrap = 0 }
    Centre (New-FitText $Title 96 $light $Font ($W - 200)) ($H * 0.4)
    Starter 'Go'
    Page 2 $Wall                                                     # slide 2: the maze
    $wallShape = New-Rect 0 0 $W $H; $wallShape.Fill.ForeColor.RGB = $dark; $wallShape.Line.Visible = 0; $wallShape.Name = 'Wall'
    for ($i = 0; $i -lt $pts.Count - 1; $i++) {                      # one rectangle per leg, overlapping at the corners
        $x1 = [math]::Min($pts[$i][0], $pts[$i + 1][0]) - $hw; $y1 = [math]::Min($pts[$i][1], $pts[$i + 1][1]) - $hw
        $leg = New-Rect ([single]$x1) ([single]$y1) ([single]([math]::Abs($pts[$i][0] - $pts[$i + 1][0]) + $Width)) ([single]([math]::Abs($pts[$i][1] - $pts[$i + 1][1]) + $Width))
        $leg.Fill.ForeColor.RGB = $light; $leg.Line.Visible = 0; $leg.Name = "Leg$($i + 1)"
    }
    $s0 = $pts[0]; $pad = $sl.Shapes.AddShape(9, [single]($s0[0] - 34), [single]($s0[1] - 34), 68, 68); $pad.Fill.ForeColor.RGB = $light; $pad.Line.Visible = 0; $pad.Name = 'StartPad'
    $st = New-Text $Start 11 $dark $Font; $st.Left = [single]($s0[0] - [double]$st.Width / 2); $st.Top = [single]($s0[1] - [double]$st.Height / 2)
    $e = $pts[-1]; $goalShape = $sl.Shapes.AddShape(9, [single]($e[0] - 34), [single]($e[1] - 34), 68, 68); $goalShape.Fill.ForeColor.RGB = $acc; $goalShape.Line.Visible = 0; $goalShape.Name = 'Goal'
    $goalShape.TextFrame.TextRange.Text = $Goal; $goalShape.TextFrame.TextRange.Font.Name = $Font; $goalShape.TextFrame.TextRange.Font.Size = 11; $goalShape.TextFrame.WordWrap = 0
    Page 3 $Accent                                                   # slide 3: touched the wall
    Centre (New-FitText $Fail 84 0xFFFFFF $Font ($W - 200)) ($H * 0.45)
    Starter 'Again'
    Page 4 $Path                                                     # slide 4: reached the goal
    Centre (New-FitText $Clear 110 $dark $Font ($W - 200)) ($H * 0.45)
    $restart = New-Rect 0 0 $W $H; $restart.Fill.Transparency = 1; $restart.Line.Visible = 0; $restart.Name = 'Restart'

    Link $pres.Slides(1).Shapes('Go') 1 2
    Link $pres.Slides(2).Shapes('Wall') 2 3
    Link $pres.Slides(2).Shapes('Goal') 2 4
    Link $pres.Slides(3).Shapes('Again') 1 2
    Link $pres.Slides(4).Shapes('Restart') 1 1
    foreach ($k in 1..4) { $pres.Slides($k).SlideShowTransition.AdvanceOnClick = 0 }
    $script:sl = $pres.Slides(2)
    Save-Deck $Out $Preview 2
    "saved $Out"
} catch { Close-Deck; throw }
