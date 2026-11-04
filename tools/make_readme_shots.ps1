# Re-generates the screenshots used by README.md.
#
# The pictures are part of the public page, so they must show the ENGLISH interface and
# ENGLISH sample documents: a screenshot is read before a single line of the README, and a
# foreign file name in it makes the tool look like it is not meant for the reader. The
# samples are therefore built here, from scratch, every time - nothing from the machine
# this runs on ever reaches the images.
#
# Settings are redirected to a throw-away folder for the same reason: the real ones carry
# the last used paths and the conversion history.
#
# Usage: powershell -NoProfile -STA -File tools\make_readme_shots.ps1
param(
    [string]$ExePath = '',
    [ValidateSet('en', 'ru')][string]$Language = 'en',
    [string]$OutputDir = '',
    [string]$SampleDirectory = ''
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$exe = if ([string]::IsNullOrWhiteSpace($ExePath)) {
    Join-Path $root 'dist\iwoHelperDesktop.exe'
} else {
    [IO.Path]::GetFullPath($ExePath)
}
if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) {
    throw "Application executable not found: $exe"
}
$outDir = if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    Join-Path $root 'docs\screenshots'
} else {
    [IO.Path]::GetFullPath($OutputDir)
}
[void](New-Item -ItemType Directory -Path $outDir -Force)
# The fixed default keeps visible paths and PNGs reproducible. Refuse an existing directory:
# cleanup is allowed to remove only the directory this invocation created.
$work = if ([string]::IsNullOrWhiteSpace($SampleDirectory)) {
    Join-Path ([Environment]::GetFolderPath('CommonDocuments')) 'iwo-readme-shot-samples'
} else {
    [IO.Path]::GetFullPath($SampleDirectory)
}
if (Test-Path -LiteralPath $work) {
    throw "Sample directory already exists; move it or pass an unused -SampleDirectory: $work"
}
[void](New-Item -ItemType Directory -Path $work)
$createdWork = $true
$books = Join-Path $work 'Workbooks'
[void](New-Item -ItemType Directory -Path $books -Force)

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'build\PdfSharp.dll'))
[void][System.Reflection.Assembly]::LoadFrom($exe)

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Shot {
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
}
"@

# ---------- sample documents ----------

function Write-Col($g, $font, $text, [double]$x, [double]$y, [double]$w, [double]$maxY) {
    $brush = [PdfSharp.Drawing.XBrushes]::Black
    $words = $text.Split(' ')
    $line = ''
    foreach ($word in $words) {
        $try = $(if ($line.Length -gt 0) { $line + ' ' + $word } else { $word })
        $sz = $g.MeasureString($try, $font)
        if ($sz.Width -gt $w -and $line.Length -gt 0) {
            if ($y -gt $maxY) { return $y }
            $g.DrawString($line, $font, $brush, (New-Object PdfSharp.Drawing.XPoint($x, $y)))
            $y += 12.5
            $line = $word
        } else { $line = $try }
    }
    if ($line.Length -gt 0 -and $y -le $maxY) {
        $g.DrawString($line, $font, $brush, (New-Object PdfSharp.Drawing.XPoint($x, $y)))
        $y += 12.5
    }
    return $y
}

function New-Report([string]$path, [int]$pages, [string]$title, [bool]$revised = $false) {
    $doc = New-Object PdfSharp.Pdf.PdfDocument
    $mast = New-Object PdfSharp.Drawing.XFont('Times New Roman', 18.0, [PdfSharp.Drawing.XFontStyle]::Bold)
    $sub = New-Object PdfSharp.Drawing.XFont('Times New Roman', 9.0, [PdfSharp.Drawing.XFontStyle]::Italic)
    $body = New-Object PdfSharp.Drawing.XFont('Times New Roman', 8.5)
    $small = New-Object PdfSharp.Drawing.XFont('Times New Roman', 8.0)
    $bold = New-Object PdfSharp.Drawing.XFont('Times New Roman', 8.5, [PdfSharp.Drawing.XFontStyle]::Bold)
    $colA = 'The quarter closed ahead of the plan in three of five regions. The two that lagged did so for one reason: a supplier change that moved deliveries into the next month. Revenue is therefore shown both as booked and as delivered, so the gap is visible rather than argued about. Costs were flat. Logistics grew by less than the volume it carried.'
    $colB = 'Head count is unchanged. The routing change that was meant to cut empty miles is now the default for the north and west depots. Two risk-register entries were closed; one was added for the remaining single-supplier dependency and is reviewed weekly until a second source is qualified. Internal audit sampled forty invoices and found no exception that changed the total.'
    $colC = 'What follows on this page is the layout the converter must keep: three newspaper columns, a nested numbered list, and a ruled table. A short last line of a paragraph is not a new story. A wrap inside a column is not a signature line. The right-hand list continues on the next page when the paper is A5, so nesting must survive the page break.'
    if ($revised) {
        $colA = 'The quarter closed ahead of the plan in four of five regions. The revised forecast is the one in the table. Deliveries that slipped last month have been caught up, and booked revenue now matches delivered revenue in every region except East.'
    }
    $items = @(
        '1. Close the East gap in booked versus delivered revenue.',
        '1.1 Confirm the second supplier before the next board.',
        'a) Keep the weekly risk review until then.',
        '2. Keep logistics cost below the volume it carries.',
        '2.1 Do not reopen the two closed register entries.'
    )
    $rows = @(
        @('Region', 'Booked', 'Delivered'),
        @('North', '12.4', '12.1'),
        @('West', '9.8', '9.8'),
        @('East', $(if ($revised) { '8.9' } else { '7.1' }), '8.6'),
        @('South', '6.2', '6.0')
    )
    for ($p = 1; $p -le $pages; $p++) {
        $page = $doc.AddPage()
        $g = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
        $g.DrawString(("{0}  ·  p.{1}" -f $title, $p), $mast, [PdfSharp.Drawing.XBrushes]::Black,
            (New-Object PdfSharp.Drawing.XPoint(48.0, 52.0)))
        $g.DrawString('Operations bulletin  ·  three columns, nested list, ruled table', $sub,
            [PdfSharp.Drawing.XBrushes]::DarkGray, (New-Object PdfSharp.Drawing.XPoint(48.0, 70.0)))
        $g.DrawLine([PdfSharp.Drawing.XPens]::Black, 48.0, 78.0, $page.Width.Point - 48.0, 78.0)
        $gap = 12.0
        $left = 48.0
        $usable = $page.Width.Point - 96.0
        $cw = ($usable - 2.0 * $gap) / 3.0
        $top = 92.0
        $bot = $page.Height.Point - 200.0
        [void](Write-Col $g $body $colA $left $top $cw $bot)
        [void](Write-Col $g $body $colB ($left + $cw + $gap) $top $cw $bot)
        [void](Write-Col $g $body $colC ($left + 2.0 * ($cw + $gap)) $top $cw $bot)
        $ty = $bot + 16.0
        $g.DrawString('Results by region', $bold, [PdfSharp.Drawing.XBrushes]::Black,
            (New-Object PdfSharp.Drawing.XPoint($left, $ty)))
        $ty += 14.0
        $tw = 240.0
        $th = 14.0
        for ($r = 0; $r -lt $rows.Count; $r++) {
            $rx = $left
            for ($c = 0; $c -lt 3; $c++) {
                $cell = New-Object PdfSharp.Drawing.XRect($rx, $ty, 80.0, $th)
                $g.DrawRectangle([PdfSharp.Drawing.XPens]::Black, $cell)
                $g.DrawString($rows[$r][$c], $(if ($r -eq 0) { $bold } else { $small }),
                    [PdfSharp.Drawing.XBrushes]::Black,
                    (New-Object PdfSharp.Drawing.XPoint(($rx + 4.0), ($ty + 10.0))))
                $rx += 80.0
            }
            $ty += $th
        }
        $ly = $bot + 16.0
        $lx = $left + $tw + 28.0
        $g.DrawString('Actions', $bold, [PdfSharp.Drawing.XBrushes]::Black,
            (New-Object PdfSharp.Drawing.XPoint($lx, $ly)))
        $ly += 14.0
        foreach ($it in $items) {
            $g.DrawString($it, $small, [PdfSharp.Drawing.XBrushes]::Black,
                (New-Object PdfSharp.Drawing.XPoint($lx, $ly)))
            $ly += 12.0
        }
        $g.Dispose()
    }
    $doc.Save($path)
    $doc.Dispose()
}

function New-Deck([string]$path, [string[]]$titles) {
    $doc = New-Object PdfSharp.Pdf.PdfDocument
    $titleFont = New-Object PdfSharp.Drawing.XFont('Segoe UI', 22.0, [PdfSharp.Drawing.XFontStyle]::Bold)
    $bodyFont = New-Object PdfSharp.Drawing.XFont('Segoe UI', 12.0)
    $small = New-Object PdfSharp.Drawing.XFont('Segoe UI', 11.0)
    $leftBits = @(
        @('1. Keep text as text, not a snapshot', '1.1 Nested numbering stays a list', 'a) Markers survive a page break'),
        @('North closed 4% above plan', 'West matched booked to delivered', 'East is the remaining gap'),
        @('Week 1  qualify second supplier', 'Week 3  reroute East depot', 'Week 6  board review'),
        @('Ops  14  logistics and audit', 'Finance  6  forecast owners', 'Risk  3  weekly watch'),
        @('Do not reopen closed risks', 'Keep cost below carried volume', 'Publish the two-column pack'),
        @('Columns, list and table on one slide', 'Background and bars stay pictures', 'No PowerPoint required to build')
    )
    $i = 0
    foreach ($t in $titles) {
        $page = $doc.AddPage()
        $page.Width = New-Object PdfSharp.Drawing.XUnit(720.0)
        $page.Height = New-Object PdfSharp.Drawing.XUnit(405.0)
        $g = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
        $band = New-Object PdfSharp.Drawing.XSolidBrush ([PdfSharp.Drawing.XColor]::FromArgb(15, 108, 189))
        $g.DrawRectangle($band, 0.0, 0.0, 720.0, 56.0)
        $g.DrawString($t, $titleFont, [PdfSharp.Drawing.XBrushes]::White, (New-Object PdfSharp.Drawing.XPoint(36.0, 38.0)))
        $bits = $leftBits[$i % $leftBits.Count]
        $y = 88.0
        foreach ($line in $bits) {
            $g.DrawString($line, $bodyFont, [PdfSharp.Drawing.XBrushes]::Black,
                (New-Object PdfSharp.Drawing.XPoint(40.0, $y)))
            $y += 28.0
        }
        $tx = 340.0
        $ty = 88.0
        $g.DrawString('Region', $small, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint($tx, $ty)))
        $g.DrawString('Booked', $small, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(($tx + 90), $ty)))
        $g.DrawString('Done', $small, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(($tx + 170), $ty)))
        $ty += 8.0
        $g.DrawLine([PdfSharp.Drawing.XPens]::Black, $tx, $ty, ($tx + 240), $ty)
        foreach ($pair in @(@('North','12.4','12.1'), @('West','9.8','9.8'), @('East','7.1','8.6'))) {
            $ty += 22.0
            $g.DrawString($pair[0], $small, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint($tx, $ty)))
            $g.DrawString($pair[1], $small, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(($tx + 90), $ty)))
            $g.DrawString($pair[2], $small, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(($tx + 170), $ty)))
        }
        $bar = New-Object PdfSharp.Drawing.XSolidBrush ([PdfSharp.Drawing.XColor]::FromArgb(16, 124, 65))
        $h = 24.0
        for ($k = 0; $k -lt 5; $k++) {
            $g.DrawRectangle($bar, (500.0 + $k * 36), (360.0 - $h), 22.0, $h)
            $h += 16.0
        }
        $g.Dispose()
        $i++
    }
    $doc.Save($path)
    $doc.Dispose()
}

# The Excel window lists a folder by extension and does not open the books, but an empty
# file with an .xlsx name is still a lie to anyone who repeats this by hand.
function New-Workbook([string]$path, [string]$sheet) {
    $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $parts = @{
            '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
            '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
            'xl/workbook.xml' = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{0}" sheetId="1" r:id="rId1"/></sheets></workbook>' -f $sheet)
            'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
            'xl/worksheets/sheet1.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row></sheetData></worksheet>'
        }
        foreach ($name in $parts.Keys) {
            $entry = $zip.CreateEntry($name)
            $writer = New-Object System.IO.StreamWriter($entry.Open())
            try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose() }
}

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$report = Join-Path $work 'Quarterly report.pdf'
$appendix = Join-Path $work 'Appendix A.pdf'
$reviewEarlier = Join-Path $work 'Earlier report.pdf'
$reviewLater = Join-Path $work 'Later report.pdf'
$deck = Join-Path $work 'Quarterly deck.pdf'
New-Report $report 11 'Quarterly report'
New-Report $appendix 4 'Appendix A'
New-Report $reviewEarlier 3 'Quarterly report'
New-Report $reviewLater 3 'Quarterly report' $true
New-Deck $deck @('Quarterly report', 'Revenue by region', 'Delivery timeline', 'Team and roles', 'Next steps', 'Summary')
foreach ($b in @('Revenue by region', 'Costs by department', 'Head count', 'Logistics routes', 'Risk register')) {
    New-Workbook (Join-Path $books ($b + '.xlsx')) 'Sheet1'
}

# ---------- application ----------

$asm = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'iwoHelperDesktop' }
$flags = [Reflection.BindingFlags]'NonPublic,Static'
[string]$appData = Join-Path $work 'AppData'
$asm.GetType('ExcelMerger.AppPaths').GetMethod('SetRootForTests', $flags).Invoke($null, [object[]]@($appData))
$loc = $asm.GetType('ExcelMerger.Loc')
[void]$loc.GetMethod('Init').Invoke($null, @($loc.GetMethod('Parse').Invoke($null, @($Language))))

# A plain backdrop: the desktop and other windows must not show through a rounded corner.
$backdrop = New-Object System.Windows.Forms.Form
$backdrop.FormBorderStyle = 'None'
$backdrop.BackColor = [System.Drawing.Color]::FromArgb(232, 234, 238)
$backdrop.ShowInTaskbar = $false
$backdrop.StartPosition = 'Manual'
$backdrop.Bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$backdrop.Show()

function Pump([int]$ms) {
    for ($i = 0; $i -lt $ms; $i += 40) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 40
    }
}

function Save-Shot($form, [string]$name) {
    $bmp = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    [void][Shot]::PrintWindow($form.Handle, $hdc, 2)
    $g.ReleaseHdc($hdc)
    $g.Dispose()
    $path = Join-Path $outDir ($name + '.png')
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Host ("{0}.png {1}x{2}" -f $name, $bmp.Width, $bmp.Height)
    $bmp.Dispose()
}

function Show-Form($form, [int]$w, [int]$h) {
    $form.StartPosition = 'Manual'
    $form.Location = New-Object System.Drawing.Point(40, 40)
    if ($w -gt 0) { $form.Size = New-Object System.Drawing.Size($w, $h) }
    $form.Show()
    $form.Activate()
    [void][Shot]::SetForegroundWindow($form.Handle)
    Pump 700
}

function Show-Loaded([string]$type, [string[]]$files, [int]$w, [int]$h, [int]$wait = 4000) {
    $t = $asm.GetType($type)
    $ctor = $t.GetConstructor([type[]]@([Action]))
    $form = if ($ctor) { $ctor.Invoke(@($null)) } else { [Activator]::CreateInstance($t, $true) }
    Show-Form $form $w $h
    $accept = $t.GetInterface('ExcelMerger.IFileAcceptor').GetMethod('AcceptFiles')
    [void]$accept.Invoke($form, @(, [string[]]$files))
    if ($type -eq 'ExcelMerger.OcrForm' -or $type -eq 'ExcelMerger.PptxForm') {
        Pump 800
        $inst = [Reflection.BindingFlags]'NonPublic,Instance'
        $grid = $t.GetField('_grid', $inst).GetValue($form)
        $grid.SetTileWidth(200)
    }
    if ($type -eq 'ExcelMerger.PdfReviewForm') {
        Pump 3000
        $inst = [Reflection.BindingFlags]'NonPublic,Instance'
        $compare = $t.GetField('_compare', $inst).GetValue($form)
        $compare.PerformClick()
    }
    Pump $wait
    return $form
}

# 1. The start screen itself - the first thing anyone sees.
$hub = [Activator]::CreateInstance($asm.GetType('ExcelMerger.StartForm'), $true)
Show-Form $hub 0 0
Pump 500
Save-Shot $hub 'hub-main'

# 2. The PDF section with the six tools.
$level = $asm.GetType('ExcelMerger.HubLevel')
$show = $hub.GetType().GetMethod('ShowLevel', [Reflection.BindingFlags]'NonPublic,Instance')
[void]$show.Invoke($hub, @([Enum]::Parse($level, 'Pdf')))
Pump 700
Save-Shot $hub 'hub'
$hub.Close(); $hub.Dispose(); Pump 300

# 3. Compare PDF (beta). The pair is generated above and contains no private or root-fixture path.
$f = Show-Loaded 'ExcelMerger.PdfReviewForm' @($reviewEarlier, $reviewLater) 1120 760 7000
Save-Shot $f 'review'
$f.Close(); $f.Dispose(); Pump 300

# 4. Merge PDF.
$f = Show-Loaded 'ExcelMerger.PdfMergeForm' @($report, $appendix) 782 692
Save-Shot $f 'pdf-merge'
$f.Close(); $f.Dispose(); Pump 300

# 5. Split PDF.
$f = Show-Loaded 'ExcelMerger.PdfSplitForm' @($report) 802 692
Save-Shot $f 'pdf-split'
$f.Close(); $f.Dispose(); Pump 300

# 6. PDF to Word.
$f = Show-Loaded 'ExcelMerger.OcrForm' @($report, $appendix) 802 692
Save-Shot $f 'pdf-word'
$f.Close(); $f.Dispose(); Pump 300

# 7. PDF to PowerPoint.
$f = Show-Loaded 'ExcelMerger.PptxForm' @($deck) 816 699
Save-Shot $f 'pdf-pptx'
$f.Close(); $f.Dispose(); Pump 300

# 8. Merge Excel.
$t = $asm.GetType('ExcelMerger.MainForm')
$form = $t.GetConstructor([type[]]@([Action])).Invoke(@($null))
Show-Form $form 782 760
$inst = [Reflection.BindingFlags]'NonPublic,Instance'
$t.GetField('_txtInput', $inst).GetValue($form).Text = $books
$t.GetField('_txtOutDir', $inst).GetValue($form).Text = $work
Pump 3000
Save-Shot $form 'excel'
$form.Close(); $form.Dispose()

$backdrop.Close(); $backdrop.Dispose()
if ($createdWork -and (Test-Path -LiteralPath $work)) {
    Remove-Item -LiteralPath $work -Recurse -Force
}
Write-Host ('cleaned samples: ' + $work)
$asm.GetType('ExcelMerger.FastExit').GetMethod('Now').Invoke($null, [object[]]@(0))
