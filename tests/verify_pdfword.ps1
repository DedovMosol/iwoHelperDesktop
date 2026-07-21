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
$f2  = New-Object PdfSharp.Drawing.XFont('Times New Roman', 12)
$f2i = New-Object PdfSharp.Drawing.XFont('Times New Roman', 16, [PdfSharp.Drawing.XFontStyle]::Italic)  # абзац 2 — курсив 16
$bk = [PdfSharp.Drawing.XBrushes]::Black
$mL = 70.0; $mInd = 105.0; $mR = 525.0   # левое поле, красная строка (+35pt), правое поле
function PutS([string]$t,[double]$x,[double]$y,$fnt) { $g2.DrawString($t,$fnt,$bk,(New-Object PdfSharp.Drawing.XPoint($x,$y))) }
function PutR([string]$t,[double]$y,$fnt) { $w=$g2.MeasureString($t,$fnt).Width; PutS $t ($mR-$w) $y $fnt }
PutS 'Абзац' $mInd 80 $f2;  PutR 'первыйй' 80 $f2    # абзац 1: красная строка + полные justified-строки, обычный 12
PutS 'текст' $mL 98 $f2;    PutR 'полнаяя' 98 $f2
PutS 'ещёее' $mL 116 $f2;   PutR 'строкаа' 116 $f2
PutS 'конец.' $mL 134 $f2                              # короткая последняя строка
PutS 'Второй' $mInd 166 $f2i; PutR 'абзацц' 166 $f2i  # абзац 2: снова красная строка, курсив 16
PutS 'текст' $mL 189 $f2i;   PutR 'дальшее' 189 $f2i
PutS 'снова' $mL 212 $f2i;   PutR 'полнаяя' 212 $f2i
PutS 'точкаа.' $mL 235 $f2i
$g2.Dispose(); $doc2.Save($pdf2); $doc2.Dispose()

$docx2 = Join-Path $PSScriptRoot 'out\extracted_indent.docx'
Remove-Item $docx2 -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert($pdf2, $docx2)

# 2c) Центрированная строка + пословный формат: красное слово и полужирное слово среди обычных.
$pdf3 = Join-Path $PSScriptRoot 'out\wordsrc_fmt.pdf'
Remove-Item $pdf3 -Force -ErrorAction SilentlyContinue
$doc3 = New-Object PdfSharp.Pdf.PdfDocument
$p3 = $doc3.AddPage(); $p3.Width = 595; $p3.Height = 842
$g3 = [PdfSharp.Drawing.XGraphics]::FromPdfPage($p3)
$reg = New-Object PdfSharp.Drawing.XFont('Times New Roman', 12)
$bld = New-Object PdfSharp.Drawing.XFont('Times New Roman', 12, [PdfSharp.Drawing.XFontStyle]::Bold)
$blk = [PdfSharp.Drawing.XBrushes]::Black
$red = New-Object PdfSharp.Drawing.XSolidBrush ([PdfSharp.Drawing.XColor]::FromArgb(220,0,0))
$cL = 70.0; $cR = 525.0
$sev = 'Семь'; $w7 = $g3.MeasureString($sev,$reg).Width
$g3.DrawString($sev,$reg,$blk,(New-Object PdfSharp.Drawing.XPoint((($cL+$cR)/2 - $w7/2), 80)))   # центрированная строка
$x3 = $cL
function Seq3([string]$t,$fnt,$br,[double]$y) { $g3.DrawString($t,$fnt,$br,(New-Object PdfSharp.Drawing.XPoint($script:x3,$y))); $script:x3 += $g3.MeasureString($t,$fnt).Width + 4 }
Seq3 'обычный' $reg $blk 110
Seq3 'КРАСНЫЙ' $reg $red 110      # красный ран
Seq3 'жирный' $bld $blk 110       # полужирный ран
$rf='добор'; $wf=$g3.MeasureString($rf,$reg).Width; $g3.DrawString($rf,$reg,$blk,(New-Object PdfSharp.Drawing.XPoint(($cR-$wf),110)))
$g3.DrawString('вторая',$reg,$blk,(New-Object PdfSharp.Drawing.XPoint($cL,130)))
$rf2='строка'; $wf2=$g3.MeasureString($rf2,$reg).Width; $g3.DrawString($rf2,$reg,$blk,(New-Object PdfSharp.Drawing.XPoint(($cR-$wf2),130)))
$g3.Dispose(); $doc3.Save($pdf3); $doc3.Dispose()

$docx3 = Join-Path $PSScriptRoot 'out\extracted_fmt.docx'
Remove-Item $docx3 -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert($pdf3, $docx3)

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

    # Красная строка + сегментация + формат: два абзаца, заметный отступ первой строки (~35pt),
    # и второй абзац перенесён курсивом кеглем 16.
    $wdoc2 = $word.Documents.Open($docx2, $false, $true)
    $nonEmpty = 0; $maxIndent = 0.0; $anyItalic16 = $false
    foreach ($par in $wdoc2.Paragraphs) {
        if ((($par.Range.Text).Trim()).Length -eq 0) { continue }
        $nonEmpty++
        $fi = [double]$par.Format.FirstLineIndent
        if ($fi -gt $maxIndent) { $maxIndent = $fi }
        if (($par.Range.Font.Italic -ne 0) -and ([math]::Round([double]$par.Range.Font.Size) -eq 16)) { $anyItalic16 = $true }
    }
    if ($nonEmpty -lt 2) { $fails += "отступный документ дал абзацев: $nonEmpty (ожидалось >=2)" }
    if ($maxIndent -le 10) { $fails += "красная строка не применена (FirstLineIndent=$maxIndent pt)" }
    if (-not $anyItalic16) { $fails += 'курсив 16pt не перенесён в docx' }
    $wdoc2.Close($false)

    # Центрирование + пословный формат: центрированный абзац, красное слово (BGR 220), полужирное слово.
    $wdoc3 = $word.Documents.Open($docx3, $false, $true)
    $centered = $false
    foreach ($par in $wdoc3.Paragraphs) { if ([int]$par.Alignment -eq 1) { $centered = $true } }
    $hasRed = $false; $hasBold = $false
    foreach ($ww in $wdoc3.Words) {
        if ([int]$ww.Font.Bold -ne 0) { $hasBold = $true }
        if ([int]$ww.Font.Color -eq 220) { $hasRed = $true }
    }
    if (-not $centered) { $fails += 'центрированная строка не по центру в docx' }
    if (-not $hasRed) { $fails += 'красный цвет (пословно) не перенесён в docx' }
    if (-not $hasBold) { $fails += 'полужирный (пословно) не перенесён в docx' }
    $wdoc3.Close($false)
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
