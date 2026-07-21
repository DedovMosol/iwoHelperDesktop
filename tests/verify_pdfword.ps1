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

# 2) Convert (PdfToWordService: extract -> .docx)
$docx = Join-Path $PSScriptRoot 'out\extracted.docx'
Remove-Item $docx -Force -ErrorAction SilentlyContinue
$res = [ExcelMerger.PdfToWordService]::Convert($pdf, $docx)
if ($res.Pages -ne 1) { $fails += "страниц $($res.Pages), ожидалась 1" }
if ($res.PagesWithText -ne 1) { $fails += "страниц с текстом $($res.PagesWithText), ожидалась 1" }
if (-not (Test-Path $docx)) { $fails += 'docx не создан' }

# 2b) A justified PDF whose paragraphs are separated only by a first-line indent
#     (красная строка) at uniform spacing — the real Word-export case.
$pdf2 = Join-Path $PSScriptRoot 'out\wordsrc_indent.pdf'
Remove-Item $pdf2 -Force -ErrorAction SilentlyContinue
$doc2 = New-Object PdfSharp.Pdf.PdfDocument
$p2 = $doc2.AddPage(); $p2.Width = 595; $p2.Height = 842
$g2 = [PdfSharp.Drawing.XGraphics]::FromPdfPage($p2)
$f2 = New-Object PdfSharp.Drawing.XFont('Times New Roman', 12)
$bk = [PdfSharp.Drawing.XBrushes]::Black
$mL = 70.0; $mInd = 105.0; $mR = 525.0   # левое поле, красная строка (+35pt), правое поле
function PutS([string]$t,[double]$x,[double]$y) { $g2.DrawString($t,$f2,$bk,(New-Object PdfSharp.Drawing.XPoint($x,$y))) }
function PutR([string]$t,[double]$y) { $w=$g2.MeasureString($t,$f2).Width; PutS $t ($mR-$w) $y }
PutS 'Абзац' $mInd 80;  PutR 'первыйй' 80    # абзац 1: красная строка + полные justified-строки
PutS 'текст' $mL 98;    PutR 'полнаяя' 98
PutS 'ещёее' $mL 116;   PutR 'строкаа' 116
PutS 'конец.' $mL 134                          # короткая последняя строка
PutS 'Второй' $mInd 158; PutR 'абзацц' 158    # абзац 2: снова красная строка
PutS 'текст' $mL 176;   PutR 'дальшее' 176
PutS 'снова' $mL 194;   PutR 'полнаяя' 194
PutS 'точка.' $mL 212
$g2.Dispose(); $doc2.Save($pdf2); $doc2.Dispose()

$docx2 = Join-Path $PSScriptRoot 'out\extracted_indent.docx'
Remove-Item $docx2 -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert($pdf2, $docx2)

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

    # Красная строка + сегментация: два абзаца и заметный отступ первой строки (~35pt).
    $wdoc2 = $word.Documents.Open($docx2, $false, $true)
    $nonEmpty = 0; $maxIndent = 0.0
    foreach ($par in $wdoc2.Paragraphs) {
        if ((($par.Range.Text).Trim()).Length -gt 0) { $nonEmpty++ }
        $fi = [double]$par.Format.FirstLineIndent
        if ($fi -gt $maxIndent) { $maxIndent = $fi }
    }
    if ($nonEmpty -lt 2) { $fails += "отступный документ дал абзацев: $nonEmpty (ожидалось >=2)" }
    if ($maxIndent -le 10) { $fails += "красная строка не применена (FirstLineIndent=$maxIndent pt)" }
    $wdoc2.Close($false)
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
