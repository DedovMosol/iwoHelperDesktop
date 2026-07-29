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
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'
$outDir = Join-Path $root 'docs\screenshots'
# The samples live under Public Documents, not under the temp folder: their path is VISIBLE
# in the Excel window, and a scratch path with a machine's drive letters in it says more
# about this computer than about the program. Falls back to the temp folder if that folder
# cannot be written to.
$work = 'C:\Users\Public\Documents\Reports'
try {
    if (Test-Path $work) { Remove-Item $work -Recurse -Force }
    [void](New-Item -ItemType Directory -Path $work -Force)
} catch {
    $work = Join-Path $env:TEMP 'iwo_readme_shots'
    if (Test-Path $work) { Remove-Item $work -Recurse -Force }
}
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

function New-Report([string]$path, [int]$pages, [string]$title) {
    $doc = New-Object PdfSharp.Pdf.PdfDocument
    $head = New-Object PdfSharp.Drawing.XFont('Times New Roman', 16.0, [PdfSharp.Drawing.XFontStyle]::Bold)
    $body = New-Object PdfSharp.Drawing.XFont('Times New Roman', 10.0)
    $lines = @(
        'The quarter closed ahead of the plan in three of the five regions, and the two that',
        'lagged did so for the same reason: a supplier change that moved deliveries into the',
        'following month. Revenue is therefore reported both ways below - as booked and as',
        'delivered - so that the gap is visible rather than argued about.',
        '',
        'Costs were flat against the previous quarter. The one line that grew, logistics, grew',
        'by less than the volume it carried, which is the result the new routing was meant to',
        'produce. Head count is unchanged.',
        '',
        'The risk register lost two entries and gained one. The entry gained concerns the',
        'single-supplier dependency described above and is tracked weekly until a second',
        'source is qualified.'
    )
    for ($p = 1; $p -le $pages; $p++) {
        $page = $doc.AddPage()
        $g = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
        $g.DrawString(("{0} - section {1}" -f $title, $p), $head, [PdfSharp.Drawing.XBrushes]::Black,
            (New-Object PdfSharp.Drawing.XPoint(60.0, 80.0)))
        $y = 120.0
        for ($k = 0; $k -lt 4; $k++) {
            foreach ($line in $lines) {
                if ($line.Length -gt 0) {
                    $g.DrawString($line, $body, [PdfSharp.Drawing.XBrushes]::Black,
                        (New-Object PdfSharp.Drawing.XPoint(60.0, $y)))
                }
                $y += 14.0
            }
            $y += 8.0
        }
        $g.Dispose()
    }
    $doc.Save($path)
    $doc.Dispose()
}

function New-Deck([string]$path, [string[]]$titles) {
    $doc = New-Object PdfSharp.Pdf.PdfDocument
    $titleFont = New-Object PdfSharp.Drawing.XFont('Segoe UI', 28.0, [PdfSharp.Drawing.XFontStyle]::Bold)
    $bodyFont = New-Object PdfSharp.Drawing.XFont('Segoe UI', 15.0)
    foreach ($t in $titles) {
        $page = $doc.AddPage()
        $page.Width = New-Object PdfSharp.Drawing.XUnit(720.0)
        $page.Height = New-Object PdfSharp.Drawing.XUnit(405.0)
        $g = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
        $band = New-Object PdfSharp.Drawing.XSolidBrush ([PdfSharp.Drawing.XColor]::FromArgb(15, 108, 189))
        $g.DrawRectangle($band, 0.0, 0.0, 720.0, 72.0)
        $g.DrawString($t, $titleFont, [PdfSharp.Drawing.XBrushes]::White, (New-Object PdfSharp.Drawing.XPoint(48.0, 48.0)))
        $y = 130.0
        foreach ($line in @('Editable text, not a picture of text', 'Charts and background stay as they were',
                            'Tables become slide tables', 'No PowerPoint required')) {
            $g.DrawString(('- ' + $line), $bodyFont, [PdfSharp.Drawing.XBrushes]::Black,
                (New-Object PdfSharp.Drawing.XPoint(60.0, $y)))
            $y += 38.0
        }
        $bar = New-Object PdfSharp.Drawing.XSolidBrush ([PdfSharp.Drawing.XColor]::FromArgb(16, 124, 65))
        $h = 30.0
        for ($i = 0; $i -lt 5; $i++) {
            $g.DrawRectangle($bar, (450.0 + $i * 44), (350.0 - $h), 28.0, $h)
            $h += 18.0
        }
        $g.Dispose()
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
$deck = Join-Path $work 'Quarterly deck.pdf'
New-Report $report 11 'Quarterly report'
New-Report $appendix 4 'Appendix A'
New-Deck $deck @('Quarterly report', 'Revenue by region', 'Delivery timeline', 'Team and roles', 'Next steps', 'Summary')
foreach ($b in @('Revenue by region', 'Costs by department', 'Head count', 'Logistics routes', 'Risk register')) {
    New-Workbook (Join-Path $books ($b + '.xlsx')) 'Sheet1'
}

# ---------- application ----------

$asm = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'iwoHelperDesktop' }
$flags = [Reflection.BindingFlags]'NonPublic,Static'
[string]$appData = Join-Path $env:TEMP 'iwo_readme_appdata'
$asm.GetType('ExcelMerger.AppPaths').GetMethod('SetRootForTests', $flags).Invoke($null, [object[]]@($appData))
$loc = $asm.GetType('ExcelMerger.Loc')
[void]$loc.GetMethod('Init').Invoke($null, @($loc.GetMethod('Parse').Invoke($null, @('en'))))

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

function Show-Loaded([string]$type, [string[]]$files, [int]$w, [int]$h) {
    $t = $asm.GetType($type)
    $ctor = $t.GetConstructor([type[]]@([Action]))
    $form = if ($ctor) { $ctor.Invoke(@($null)) } else { [Activator]::CreateInstance($t, $true) }
    Show-Form $form $w $h
    $accept = $t.GetInterface('ExcelMerger.IFileAcceptor').GetMethod('AcceptFiles')
    [void]$accept.Invoke($form, @(, [string[]]$files))
    Pump 4000
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

# 3. Merge PDF.
$f = Show-Loaded 'ExcelMerger.PdfMergeForm' @($report, $appendix) 782 692
Save-Shot $f 'pdf-merge'
$f.Close(); $f.Dispose(); Pump 300

# 4. Split PDF.
$f = Show-Loaded 'ExcelMerger.PdfSplitForm' @($report) 802 692
Save-Shot $f 'pdf-split'
$f.Close(); $f.Dispose(); Pump 300

# 5. PDF to Word.
$f = Show-Loaded 'ExcelMerger.OcrForm' @($report, $appendix) 802 692
Save-Shot $f 'pdf-word'
$f.Close(); $f.Dispose(); Pump 300

# 6. PDF to PowerPoint.
$f = Show-Loaded 'ExcelMerger.PptxForm' @($deck) 816 699
Save-Shot $f 'pdf-pptx'
$f.Close(); $f.Dispose(); Pump 300

# 7. Merge Excel.
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
Write-Host ('samples: ' + $work)
