# Integration test for born-digital PDF -> Word: a digital PDF is created via PdfSharp,
# text is extracted (PdfTextExtract + OcrLayout) and written to .docx (WordDocxWriter),
# then read back via Word COM. Requires installed Word.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$root = Split-Path $PSScriptRoot
$fails = @()
New-Item -ItemType Directory -Force (Join-Path $PSScriptRoot 'out') | Out-Null

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'build\PdfSharp.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'dist\iwoHelperDesktop.exe'))

# Порядок страниц одного файла для Convert (актуальный API: список ссылок источник+индекс).
# ",$order" — обёртка от разворачивания коллекции пайплайном PowerShell.
function New-PageOrder([string]$path, [int]$rotation = 0) {
    $pages = [ExcelMerger.PdfMergeService]::LoadPages($path)
    $order = New-Object 'System.Collections.Generic.List[ExcelMerger.PdfPageRef]'
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $ref = New-Object ExcelMerger.PdfPageRef
        $ref.SourcePath = $path
        $ref.PageIndex = $i
        $ref.Rotation = $rotation
        $order.Add($ref)
    }
    ,$order
}

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
$res = [ExcelMerger.PdfToWordService]::Convert((New-PageOrder $pdf), $docx, $null)
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
$mL = 70.0; $mInd = 105.0   # левое поле и красная строка (+35pt)
function PutS([string]$t,[double]$x,[double]$y,$fnt) { $g2.DrawString($t,$fnt,$bk,(New-Object PdfSharp.Drawing.XPoint($x,$y))) }
# НЕПРЕРЫВНАЯ строка из повторов слова с обычными пробелами: широких внутристрочных
# зазоров нет, иначе GridDetector (формы «метка … значение») честно увёз бы её таблицей.
function Row2([double]$x,[double]$y,$fnt,[int]$n) {
    $step = $g2.MeasureString('словоо ', $fnt).Width
    for ($k = 0; $k -lt $n; $k++) { PutS 'словоо' $x $y $fnt; $x += $step }
}
Row2 $mInd 80 $f2 8     # абзац 1 (12pt): красная строка + равнодлинные justified-строки
Row2 $mL 98 $f2 9
Row2 $mL 116 $f2 9
PutS 'конец.' $mL 134 $f2                              # короткая последняя строка
Row2 $mInd 166 $f2i 6   # абзац 2: снова красная строка, курсив 16
Row2 $mL 189 $f2i 7
Row2 $mL 212 $f2i 7
PutS 'точкаа.' $mL 235 $f2i
$g2.Dispose(); $doc2.Save($pdf2); $doc2.Dispose()

$docx2 = Join-Path $PSScriptRoot 'out\extracted_indent.docx'
Remove-Item $docx2 -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert((New-PageOrder $pdf2), $docx2, $null)

# 2d) ОДИН абзац в шесть строк: равный шаг, общий левый край, никаких отступов и широких
#     зазоров. Разделить его не на чем — и в .docx он обязан приехать одним абзацем.
#     Это мера ДРОБЛЕНИЯ для пути в Word: у пути в презентации она попиксельная, а здесь
#     достаточно счёта, потому что правильный ответ ровно один.
$pdfFrag = Join-Path $PSScriptRoot 'out\wordsrc_onepara.pdf'
Remove-Item $pdfFrag -Force -ErrorAction SilentlyContinue
$docF = New-Object PdfSharp.Pdf.PdfDocument
$pF = $docF.AddPage(); $pF.Width = 595; $pF.Height = 842
$gF = [PdfSharp.Drawing.XGraphics]::FromPdfPage($pF)
$fF = New-Object PdfSharp.Drawing.XFont('Times New Roman', 12)
$bkF = [PdfSharp.Drawing.XBrushes]::Black
$stepF = $gF.MeasureString('словоо ', $fF).Width
for ($row = 0; $row -lt 6; $row++) {
    $x = 70.0
    for ($k = 0; $k -lt 9; $k++) {
        $gF.DrawString('словоо', $fF, $bkF, (New-Object PdfSharp.Drawing.XPoint($x, (80.0 + $row * 18.0))))
        $x += $stepF
    }
}
$gF.Dispose(); $docF.Save($pdfFrag); $docF.Dispose()

$docxFrag = Join-Path $PSScriptRoot 'out\extracted_onepara.docx'
Remove-Item $docxFrag -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert((New-PageOrder $pdfFrag), $docxFrag, $null)

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
$cL = 70.0; $cR = 525.0; $cInd = 105.0
# Кластер красных строк (два абзаца с отступом первой строки): по текущим правилам
# центр одиночной строки применяется только в документах с таким кластером
# (без него ложные центры блоков не срабатывают — осознанное поведение 1.16.7).
function Row3([double]$x,[double]$y,[int]$n) {
    $step = $g3.MeasureString('словоо ', $reg).Width
    for ($k = 0; $k -lt $n; $k++) { $g3.DrawString('словоо',$reg,$blk,(New-Object PdfSharp.Drawing.XPoint($x,$y))); $x += $step }
}
Row3 $cInd 60 8
Row3 $cL 78 9
Row3 $cInd 110 8
Row3 $cL 128 9
$sev = 'Семь'; $w7 = $g3.MeasureString($sev,$reg).Width
# Центр — относительно ТЕЛА текста (правый край 9-словных строк), а не листа:
# правила центрирования меряют зазоры до рамки контента страницы.
$stepC = $g3.MeasureString('словоо ', $reg).Width
$rowsRight = $cL + $stepC * 8 + $g3.MeasureString('словоо', $reg).Width
$g3.DrawString($sev,$reg,$blk,(New-Object PdfSharp.Drawing.XPoint((($cL+$rowsRight)/2 - $w7/2), 168)))   # центрированная строка
$x3 = $cL
function Seq3([string]$t,$fnt,$br,[double]$y) { $g3.DrawString($t,$fnt,$br,(New-Object PdfSharp.Drawing.XPoint($script:x3,$y))); $script:x3 += $g3.MeasureString($t,$fnt).Width + 4 }
Seq3 'обычный' $reg $blk 200
Seq3 'КРАСНЫЙ' $reg $red 200      # красный ран
Seq3 'жирный' $bld $blk 200       # полужирный ран
Seq3 'дальшеее' $reg $blk 200
$g3.DrawString('вторая',$reg,$blk,(New-Object PdfSharp.Drawing.XPoint($cL,220)))
$g3.DrawString('строка.',$reg,$blk,(New-Object PdfSharp.Drawing.XPoint(($cL+45),220)))
$g3.Dispose(); $doc3.Save($pdf3); $doc3.Dispose()

$docx3 = Join-Path $PSScriptRoot 'out\extracted_fmt.docx'
Remove-Item $docx3 -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert((New-PageOrder $pdf3), $docx3, $null)

# 2d) Изображение из PDF должно попасть в .docx как встроенная картинка.
$png = Join-Path $PSScriptRoot 'out\sq.png'
$bmp = New-Object System.Drawing.Bitmap(80, 60); $gr = [System.Drawing.Graphics]::FromImage($bmp)
$gr.Clear([System.Drawing.Color]::Blue)
# Двухцветная: одноцветный растр конвейер бракует НАМЕРЕННО (признак сбоя декодера).
$gr.FillRectangle([System.Drawing.Brushes]::Red, 0, 0, 40, 60)
$gr.Dispose(); $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
$pdf4 = Join-Path $PSScriptRoot 'out\wordsrc_img.pdf'
Remove-Item $pdf4 -Force -ErrorAction SilentlyContinue
$doc4 = New-Object PdfSharp.Pdf.PdfDocument; $p4 = $doc4.AddPage(); $p4.Width = 595; $p4.Height = 842
$g4 = [PdfSharp.Drawing.XGraphics]::FromPdfPage($p4)
$g4.DrawString('Рисунок ниже', (New-Object PdfSharp.Drawing.XFont('Times New Roman', 12)), [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(70, 90)))
$xi = [PdfSharp.Drawing.XImage]::FromFile($png); $g4.DrawImage($xi, 70, 120, 80, 60); $xi.Dispose()
$g4.Dispose(); $doc4.Save($pdf4); $doc4.Dispose()
$docx4 = Join-Path $PSScriptRoot 'out\extracted_img.docx'
Remove-Item $docx4 -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert((New-PageOrder $pdf4), $docx4, $null)

# 2e) Гиперссылка из PDF должна перенестись в .docx как Word Hyperlink.
$pdf5 = Join-Path $PSScriptRoot 'out\wordsrc_link.pdf'
Remove-Item $pdf5 -Force -ErrorAction SilentlyContinue
$doc5 = New-Object PdfSharp.Pdf.PdfDocument; $p5 = $doc5.AddPage(); $p5.Width = 595; $p5.Height = 842
$g5 = [PdfSharp.Drawing.XGraphics]::FromPdfPage($p5)
$g5.DrawString('Ссылка тут', (New-Object PdfSharp.Drawing.XFont('Times New Roman', 12)), [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(70, 110)))
$xr5 = New-Object PdfSharp.Drawing.XRect(0, 0, 595, 842)   # рамка ссылки на всю страницу — слово точно внутри
[void]$p5.AddWebLink((New-Object PdfSharp.Pdf.PdfRectangle($xr5)), 'https://example.com')
$g5.Dispose(); $doc5.Save($pdf5); $doc5.Dispose()
$docx5 = Join-Path $PSScriptRoot 'out\extracted_link.docx'
Remove-Item $docx5 -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert((New-PageOrder $pdf5), $docx5, $null)

# 2f) Боковая страница + поворот в сетке: текст, нарисованный по часовой, выправляется
#     поворотом 270° — в .docx уходит ландшафтная страница с нормальными строками.
$pdf6 = Join-Path $PSScriptRoot 'out\wordsrc_rot.pdf'
Remove-Item $pdf6 -Force -ErrorAction SilentlyContinue
$doc6 = New-Object PdfSharp.Pdf.PdfDocument; $p6 = $doc6.AddPage(); $p6.Width = 595; $p6.Height = 842
$g6 = [PdfSharp.Drawing.XGraphics]::FromPdfPage($p6)
$g6.RotateAtTransform(90, (New-Object PdfSharp.Drawing.XPoint(297, 420)))
$g6.DrawString('ROTATEDLINE', (New-Object PdfSharp.Drawing.XFont('Times New Roman', 14)), [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(200, 420)))
$g6.Dispose(); $doc6.Save($pdf6); $doc6.Dispose()
$docx6 = Join-Path $PSScriptRoot 'out\extracted_rot.docx'
Remove-Item $docx6 -Force -ErrorAction SilentlyContinue
[void][ExcelMerger.PdfToWordService]::Convert((New-PageOrder $pdf6 270), $docx6, $null)

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
    # РОВНО два, а не «хотя бы два»: нижняя граница ловит слипание абзацев в один, но
    # молчит про обратную беду — когда один абзац приезжает четырьмя кусками. Дробление
    # заметно только точным числом.
    if ($nonEmpty -ne 2) { $fails += "отступный документ дал абзацев: $nonEmpty (ожидалось ровно 2)" }
    if ($maxIndent -le 10) { $fails += "красная строка не применена (FirstLineIndent=$maxIndent pt)" }
    if (-not $anyItalic16) { $fails += 'курсив 16pt не перенесён в docx' }
    # Поля страницы унаследованы: исходник A4 (595) с полями 70.
    $ps2 = $wdoc2.PageSetup
    $pw2 = [math]::Round([double]$ps2.PageWidth)
    if ($pw2 -lt 590 -or $pw2 -gt 600) { $fails += "ширина страницы не унаследована: $pw2" }
    $lm2 = [double]$ps2.LeftMargin
    if ($lm2 -lt 60 -or $lm2 -gt 80) { $fails += "левое поле не унаследовано: $lm2" }
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

    # Дробление: шесть строк одного абзаца обязаны остаться ОДНИМ абзацем.
    $wdocF = $word.Documents.Open($docxFrag, $false, $true)
    $paraF = 0
    foreach ($par in $wdocF.Paragraphs) {
        if ((($par.Range.Text).Trim()).Length -gt 0) { $paraF++ }
    }
    Write-Host ("дробление: абзацев в docx = " + $paraF + " на 1 абзац источника (6 строк)")
    if ($paraF -ne 1) { $fails += "сплошной абзац из 6 строк разбит на $paraF абзацев" }
    $wdocF.Close($false)

    # Изображение вставлено в .docx.
    $wdoc4 = $word.Documents.Open($docx4, $false, $true)
    if ([int]$wdoc4.InlineShapes.Count -lt 1) { $fails += 'изображение не вставлено в docx' }
    $wdoc4.Close($false)

    # Гиперссылка перенесена в .docx.
    $wdoc5 = $word.Documents.Open($docx5, $false, $true)
    if ([int]$wdoc5.Hyperlinks.Count -lt 1) { $fails += 'гиперссылка не перенесена в docx' }
    $wdoc5.Close($false)

    # Повёрнутая страница: текст выправлен, страница стала ландшафтной (ширина > высоты).
    $wdoc6 = $word.Documents.Open($docx6, $false, $true)
    $text6 = ($wdoc6.Content.Text) -replace ' ', ''
    if ($text6 -notmatch 'ROTATEDLINE') { $fails += 'повёрнутый текст не выправлен в docx' }
    $ps6 = $wdoc6.PageSetup
    if ([double]$ps6.PageWidth -le [double]$ps6.PageHeight) { $fails += 'страница не стала ландшафтной после поворота' }
    $wdoc6.Close($false)
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
