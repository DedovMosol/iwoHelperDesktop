# Integration test for born-digital PDF -> Word: a digital PDF is created via PdfSharp,
# text is extracted (PdfTextExtract + OcrLayout) and written to .docx (WordDocxWriter),
# then read back via Word COM. Requires installed Word.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$fails = @()
New-Item -ItemType Directory -Force (Join-Path $PSScriptRoot 'out') | Out-Null

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'build\PdfSharp.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'dist\iwoHelperDesktop.exe'))

# 1) A born-digital PDF with known text (two lines).
$pdf = Join-Path $PSScriptRoot 'out\wordsrc.pdf'
Remove-Item $pdf -Force -ErrorAction SilentlyContinue
$doc = New-Object PdfSharp.Pdf.PdfDocument
$page = $doc.AddPage()
$g = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
$font = New-Object PdfSharp.Drawing.XFont('Times New Roman', 14)
$g.DrawString('Привет, мир! Hello world.', $font, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(50, 100)))
$g.DrawString('Вторая строка текста.', $font, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(50, 130)))
$g.Dispose(); $doc.Save($pdf); $doc.Dispose()

# 2) Extract -> .docx
$pages = [ExcelMerger.PdfTextExtract]::Extract($pdf)
if ($pages.Count -ne 1) { $fails += "страниц $($pages.Count), ожидалась 1" }
$docx = Join-Path $PSScriptRoot 'out\extracted.docx'
Remove-Item $docx -Force -ErrorAction SilentlyContinue
[ExcelMerger.WordDocxWriter]::Write($pages, $docx)
if (-not (Test-Path $docx)) { $fails += 'docx не создан' }

# 3) Read the .docx back via Word.
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $wdoc = $word.Documents.Open($docx, $false, $true)
    $text = $wdoc.Content.Text
    if ($text -notmatch 'Hello world') { $fails += 'нет латиницы в docx' }
    if ($text -notmatch 'Привет') { $fails += 'нет кириллицы в docx' }
    if ($text -notmatch 'Вторая строка') { $fails += 'нет второй строки в docx' }
    $wdoc.Close($false)
}
finally {
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
}

if ($fails.Count -eq 0) {
    Write-Host 'VERIFY PDFWORD OK'
    exit 0
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    exit 1
}
