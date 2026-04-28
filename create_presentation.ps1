# create_presentation.ps1
# PowerPoint Creator Skill - Setup and Usage Guide
# Requires Microsoft Office / PowerPoint installed. No pip packages needed.

function rgb($r, $g, $b) { return [int]($b * 65536 + $g * 256 + $r) }

$cBg      = rgb 0x0D 0x11 0x17
$cSurface = rgb 0x16 0x1B 0x22
$cAccent  = rgb 0x58 0xA6 0xFF
$cGreen   = rgb 0x3F 0xB9 0x50
$cText    = rgb 0xE6 0xED 0xF3
$cMuted   = rgb 0x89 0x92 0x9B
$cWhite   = rgb 0xFF 0xFF 0xFF
$cBorder  = rgb 0x30 0x36 0x3D

$outputPath = Join-Path $PSScriptRoot "PowerPointCreatorSkill_Guide.pptx"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$prs = $ppt.Presentations.Add()
$prs.PageSetup.SlideWidth  = 960
$prs.PageSetup.SlideHeight = 540

function Add-Bg($slide, $color) {
    $r = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
    $r.Fill.ForeColor.RGB = $color; $r.Fill.Solid(); $r.Line.Visible = 0; $r.ZOrder(1)
}
function Add-Rect($slide, $x, $y, $w, $h, $color) {
    $r = $slide.Shapes.AddShape(1, $x, $y, $w, $h)
    $r.Fill.ForeColor.RGB = $color; $r.Fill.Solid(); $r.Line.Visible = 0; return $r
}
function Add-Text($slide, $x, $y, $w, $h, $text, $size, $color, $bold=0, $align=1) {
    $tb = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $tb.TextFrame.WordWrap = 1
    $tb.TextFrame.MarginLeft = 0; $tb.TextFrame.MarginRight = 0
    $tb.TextFrame.MarginTop  = 0; $tb.TextFrame.MarginBottom = 0
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $text; $tr.Font.Name = "Segoe UI"; $tr.Font.Size = $size
    $tr.Font.Bold = $bold; $tr.Font.Color.RGB = $color
    $tr.ParagraphFormat.Alignment = $align
}
function Add-Bullets($slide, $x, $y, $w, $h, $items, $size, $color, $spacing=8) {
    $tb = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $tb.TextFrame.WordWrap = 1
    $tb.TextFrame.MarginLeft = 0; $tb.TextFrame.MarginRight = 0
    $tb.TextFrame.MarginTop  = 0; $tb.TextFrame.MarginBottom = 0
    $tf = $tb.TextFrame
    for ($i = 0; $i -lt $items.Count; $i++) {
        if ($i -eq 0) { $tf.TextRange.Paragraphs(1).Text = $items[$i] }
        else { $tf.TextRange.InsertAfter("`r" + $items[$i]) | Out-Null }
        $p = $tf.TextRange.Paragraphs($i + 1)
        $p.Font.Name = "Segoe UI"; $p.Font.Size = $size
        $p.Font.Color.RGB = $color; $p.ParagraphFormat.SpaceBefore = $spacing
    }
}

# ===========================================================================
# SLIDE 1 - Title
# ===========================================================================
$s1 = $prs.Slides.Add(1, 12)
Add-Bg $s1 $cBg
Add-Rect $s1 0 0 6 540 $cAccent | Out-Null

$tag = $s1.Shapes.AddShape(1, 40, 48, 260, 32)
$tag.Fill.ForeColor.RGB = $cSurface; $tag.Fill.Solid()
$tag.Line.ForeColor.RGB = $cBorder; $tag.Line.Visible = 1; $tag.Line.Weight = 1
Add-Text $s1 50 52 240 26 "GITHUB COPILOT / AGENT SKILLS" 10 $cAccent 0 1

Add-Text $s1 40 108 700 110 "PowerPoint Creator" 46 $cText 1 1
Add-Text $s1 40 220 700 58  "Skill Setup and Usage Guide" 32 $cAccent 1 1
Add-Text $s1 40 300 700 40  "Generate professional decks from Copilot agent mode with no manual work" 15 $cMuted 0 1
Add-Rect $s1 40 396 560 2 $cBorder | Out-Null
Add-Text $s1 40 408 560 28 "User-level skill | Available in every workspace" 13 $cMuted 0 1

Add-Rect $s1 730 80 190 190 $cSurface | Out-Null
Add-Rect $s1 734 84 182 182 $cBg | Out-Null
Add-Text $s1 734 136 182 76 ".pptx" 34 $cAccent 1 2

# ===========================================================================
# SLIDE 2 - What Is This Skill?
# ===========================================================================
$s2 = $prs.Slides.Add(2, 12)
Add-Bg $s2 $cBg
Add-Rect $s2 0 0 6 540 $cAccent | Out-Null
Add-Rect $s2 0 0 960 80 $cSurface | Out-Null
Add-Rect $s2 0 78 960 2 $cBorder | Out-Null
Add-Text $s2 36 18 880 46 "What Is This Skill?" 30 $cText 1 1

# Description card
Add-Rect $s2 36 104 888 130 $cSurface | Out-Null
Add-Rect $s2 36 104 4 130 $cAccent | Out-Null
Add-Text $s2 56 116 840 28 "About" 14 $cAccent 1 1
Add-Text $s2 56 148 840 76 "The powerpoint-creator skill teaches GitHub Copilot agent mode to generate professional PowerPoint presentations (.pptx) on demand. It auto-selects the best generation method for your environment." 15 $cText 0 1

# Skill path card
Add-Rect $s2 36 250 888 60 $cSurface | Out-Null
Add-Rect $s2 36 250 4 60 $cGreen | Out-Null
Add-Text $s2 56 260 200 40 "Skill Location" 13 $cGreen 1 1
Add-Text $s2 200 264 680 30 "C:\Users\<you>\.agents\skills\powerpoint-creator\" 14 $cText 0 1

# Three capability pills
$caps = @(
    @{ title="Modern Layouts";     desc="Consistent spacing,`ntypography and themes" },
    @{ title="All Slide Types";    desc="Title, bullets, charts,`ntables, two-column" },
    @{ title="Zero Dependencies";  desc="Works via PowerShell COM`nwhen pip is blocked" }
)
$cx = 36
foreach ($c in $caps) {
    Add-Rect $s2 $cx 328 278 170 $cSurface | Out-Null
    Add-Rect $s2 $cx 328 278 4  $cAccent  | Out-Null
    Add-Text $s2 ($cx+16) 346 246 32 $c.title 17 $cAccent 1 1
    Add-Text $s2 ($cx+16) 382 246 96 $c.desc  14 $cMuted  0 1
    $cx += 294
}

# ===========================================================================
# SLIDE 3 - Requirements
# ===========================================================================
$s3 = $prs.Slides.Add(3, 12)
Add-Bg $s3 $cBg
Add-Rect $s3 0 0 6 540 $cGreen | Out-Null
Add-Rect $s3 0 0 960 80 $cSurface | Out-Null
Add-Rect $s3 0 78 960 2 $cBorder | Out-Null
Add-Text $s3 36 18 880 46 "Requirements" 30 $cText 1 1

# Option A header
Add-Rect $s3 36 104 430 42 $cSurface | Out-Null
Add-Rect $s3 36 104 430 4  $cAccent  | Out-Null
Add-Text $s3 52 112 400 28 "Option A - Python + python-pptx" 15 $cAccent 1 1

Add-Rect $s3 36 150 430 272 $cSurface | Out-Null
Add-Rect $s3 36 150 4   272 $cAccent  | Out-Null
$optA = @(
    "Python 3.8+",
    "python-pptx 1.0.0+",
    "",
    "pip install python-pptx",
    "",
    "For corporate proxy issues, add:",
    "--trusted-host pypi.org",
    "--trusted-host files.pythonhosted.org"
)
Add-Bullets $s3 56 162 400 248 $optA 13 $cText 6

# Option B header
Add-Rect $s3 494 104 430 42 $cSurface | Out-Null
Add-Rect $s3 494 104 430 4  $cGreen   | Out-Null
Add-Text $s3 510 112 400 28 "Option B - PowerShell COM (Recommended)" 15 $cGreen 1 1

Add-Rect $s3 494 150 430 272 $cSurface | Out-Null
Add-Rect $s3 494 150 4   272 $cGreen   | Out-Null
$optB = @(
    "Windows 10 or 11",
    "Microsoft Office / PowerPoint installed",
    "PowerShell 5.1+ (built into Windows)",
    "",
    "No pip install needed",
    "No proxy or SSL issues",
    "Works in locked-down corporate environments"
)
Add-Bullets $s3 514 162 400 248 $optB 13 $cText 6

# Recommendation banner
Add-Rect $s3 36 434 888 60 $cSurface | Out-Null
Add-Rect $s3 36 434 4   60 $cGreen   | Out-Null
Add-Text $s3 56 447 840 34 "Recommended for Progressive: Option B requires no network access and no package installation." 14 $cMuted 0 1

# ===========================================================================
# SLIDE 4 - How to Use
# ===========================================================================
$s4 = $prs.Slides.Add(4, 12)
Add-Bg $s4 $cBg
Add-Rect $s4 0 0 6 540 $cAccent | Out-Null
Add-Rect $s4 0 0 960 80 $cSurface | Out-Null
Add-Rect $s4 0 78 960 2 $cBorder | Out-Null
Add-Text $s4 36 18 880 46 "How to Use the Skill" 30 $cText 1 1

$steps = @(
    @{
        num="01"; title="Open Copilot in Agent Mode";
        body="In VS Code, open the Copilot Chat panel and switch to agent mode using the dropdown next to the chat input box."
        col=$cAccent
    },
    @{
        num="02"; title="Ask for a Presentation";
        body="Describe your presentation topic, audience, and slide count. Example: 'Create a 10-slide executive summary based on the README in this project using the Midnight Dark theme.'"
        col=$cGreen
    },
    @{
        num="03"; title="Run the Generated Script";
        body="Copilot generates a .ps1 (or .py) file in your workspace. Run it with: powershell -ExecutionPolicy Bypass -File '.\create_presentation.ps1'"
        col=$cAccent
    }
)

$sy = 100
foreach ($step in $steps) {
    Add-Rect $s4 36 $sy 888 118 $cSurface | Out-Null
    Add-Rect $s4 36 $sy 4 118 $step.col | Out-Null
    Add-Text $s4 46  ($sy+16) 60  86 $step.num  34 $step.col 1 2
    Add-Text $s4 116 ($sy+14) 220 28 $step.title 15 $step.col 1 1
    Add-Text $s4 116 ($sy+46) 748 62 $step.body  14 $cMuted   0 1
    $sy += 128
}

# ===========================================================================
# SLIDE 5 - Troubleshooting
# ===========================================================================
$s5 = $prs.Slides.Add(5, 12)
Add-Bg $s5 $cBg
Add-Rect $s5 0 0 6 540 $cGreen | Out-Null
Add-Rect $s5 0 0 960 80 $cSurface | Out-Null
Add-Rect $s5 0 78 960 2 $cBorder | Out-Null
Add-Text $s5 36 18 880 46 "Troubleshooting" 30 $cText 1 1

$issues = @(
    @{ issue="SSL Certificate Error";     fix="Add --trusted-host flags to pip, or switch to Option B (PowerShell COM)";  col=$cGreen },
    @{ issue="Hash Mismatch on Download"; fix="Corporate proxy rewrites packages. Run pip cache purge then retry, or use Option B"; col=$cGreen },
    @{ issue="pip Upgrade Fails";         fix="Use: python -m pip install --upgrade pip --trusted-host pypi.org --trusted-host files.pythonhosted.org"; col=$cGreen },
    @{ issue="Encoding / Token Error";    fix="Special characters (em dashes, smart quotes) in .ps1 files. Replace with plain ASCII equivalents (- instead of a dash character)"; col=$cGreen }
)

$iy = 100
foreach ($item in $issues) {
    Add-Rect $s5 36 $iy 888 88 $cSurface | Out-Null
    Add-Rect $s5 36 $iy 4   88 $item.col  | Out-Null
    Add-Text $s5 56 ($iy+10) 840 26 $item.issue 15 $item.col 1 1
    Add-Text $s5 56 ($iy+38) 840 40 $item.fix   13 $cMuted   0 1
    $iy += 98
}

# ===========================================================================
# SLIDE 6 - Color Themes
# ===========================================================================
$s6 = $prs.Slides.Add(6, 12)
Add-Bg $s6 $cBg
Add-Rect $s6 0 0 6 540 $cAccent | Out-Null
Add-Rect $s6 0 0 960 80 $cSurface | Out-Null
Add-Rect $s6 0 78 960 2 $cBorder | Out-Null
Add-Text $s6 36 18 880 46 "Color Themes" 30 $cText 1 1

$themes = @(
    @{ name="Corporate Blue"; primary=(rgb 0x1F 0x35 0x64); accent=(rgb 0x2E 0x75 0xB6); bg=(rgb 0xFF 0xFF 0xFF); pHex="#1F3564"; aHex="#2E75B6" },
    @{ name="Slate and Teal"; primary=(rgb 0x2C 0x3E 0x50); accent=(rgb 0x1A 0xBC 0x9C); bg=(rgb 0xF8 0xF9 0xFA); pHex="#2C3E50"; aHex="#1ABC9C" },
    @{ name="Crimson Pro";    primary=(rgb 0xC0 0x39 0x2B); accent=(rgb 0xE7 0x4C 0x3C); bg=(rgb 0xFF 0xFF 0xFF); pHex="#C0392B"; aHex="#E74C3C" },
    @{ name="Forest Green";   primary=(rgb 0x1E 0x57 0x4B); accent=(rgb 0x27 0xAE 0x60); bg=(rgb 0xF5 0xF5 0xF5); pHex="#1E574B"; aHex="#27AE60" },
    @{ name="Midnight Dark";  primary=(rgb 0x0D 0x11 0x17); accent=(rgb 0x58 0xA6 0xFF); bg=(rgb 0x16 0x1B 0x22); pHex="#0D1117"; aHex="#58A6FF" }
)

$ty = 102
foreach ($t in $themes) {
    Add-Rect $s6 36 $ty 888 68 $cSurface | Out-Null
    # Primary swatch
    Add-Rect $s6 36 $ty 54 68 $t.primary | Out-Null
    # Accent swatch
    Add-Rect $s6 90 $ty 54 68 $t.accent | Out-Null
    Add-Text $s6 162 ($ty+8)  260 28 $t.name    17 $cText  1 1
    Add-Text $s6 430 ($ty+10) 200 24 $t.pHex    14 $cMuted 0 1
    Add-Text $s6 640 ($ty+10) 200 24 $t.aHex    14 $cMuted 0 1
    $ty += 76
}

Add-Rect $s6 36 494 888 2 $cBorder | Out-Null
Add-Text $s6 36 500 888 28 "Specify a theme name when prompting Copilot: 'Create a presentation using the Midnight Dark theme'" 13 $cMuted 0 1

# ===========================================================================
# SLIDE 7 - Closing
# ===========================================================================
$s7 = $prs.Slides.Add(7, 12)
Add-Bg $s7 $cBg
Add-Rect $s7 0 0 6 540 $cAccent | Out-Null

$gridColor = rgb 0x1C 0x22 0x2B
for ($gx = 80; $gx -lt 960; $gx += 80) { Add-Rect $s7 $gx 0 1 540 $gridColor | Out-Null }
for ($gy = 60; $gy -lt 540; $gy += 60) { Add-Rect $s7 0 $gy 960 1 $gridColor | Out-Null }

Add-Text $s7 140 148 680 80  "Ready to Build Decks" 46 $cText   1 2
Add-Text $s7 140 236 680 40  "Open agent mode and describe your presentation" 20 $cAccent 0 2
Add-Rect $s7 360 296 240 2 $cBorder | Out-Null
Add-Text $s7 140 312 680 30  "Python or PowerShell COM - the skill handles both" 14 $cMuted  0 2

Add-Rect $s7 310 388 340 60 $cSurface | Out-Null
Add-Rect $s7 310 388 340 4  $cAccent  | Out-Null
Add-Text $s7 310 406 340 30  "powerpoint-creator skill" 14 $cAccent 0 2
Add-Text $s7 310 430 340 22  "~\.agents\skills\powerpoint-creator\" 12 $cMuted  0 2

# ===========================================================================
# Save and close
# ===========================================================================
$prs.SaveAs($outputPath, 24)
$prs.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null

Write-Host ""
Write-Host "Saved: $outputPath"
