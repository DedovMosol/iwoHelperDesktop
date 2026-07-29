# Живой тест «PDF → PowerPoint»: собираем настоящий .pptx и разбираем его так же, как это
# сделает читатель формата — распаковкой архива и разбором XML. PowerPoint для проверки НЕ
# нужен (и в CI его нет); если он всё-таки установлен, в конце файл дополнительно открывается
# через COM — это единственный способ убедиться, что он не считает файл повреждённым.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
[void][System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\..\build\PdfSharp.dll")
[void][System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\..\dist\iwoHelperDesktop.exe")

$out = Join-Path $PSScriptRoot 'out'
if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }
$fails = New-Object System.Collections.ArrayList

function Fail([string]$what) { [void]$fails.Add($what); Write-Host "FAIL: $what" }

function New-PageOrder([string]$path, [int]$rotation = 0) {
    $pages = [ExcelMerger.PdfMergeService]::LoadPages($path)
    $order = New-Object 'System.Collections.Generic.List[ExcelMerger.PdfPageRef]'
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $r = New-Object ExcelMerger.PdfPageRef
        $r.SourcePath = $path
        $r.PageIndex = $i
        $r.Rotation = $rotation
        $order.Add($r)
    }
    return ,$order
}

function Get-Entry([string]$archive, [string]$name) {
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try {
        $entry = $zip.GetEntry($name)
        if ($null -eq $entry) { return $null }
        $reader = New-Object IO.StreamReader($entry.Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
}

function Get-Names([string]$archive) {
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try { return @($zip.Entries | ForEach-Object { $_.FullName }) } finally { $zip.Dispose() }
}

function Get-Xml([string]$archive, [string]$name) {
    $text = Get-Entry $archive $name
    if ($null -eq $text) { return $null }
    return [xml]$text
}

# Возвращать ОДИН объект: без [void] и запятой PowerShell отдаёт массив из результатов
# всех вызовов, и SelectNodes потом не находит перегрузку.
function New-Ns([xml]$doc) {
    $ns = New-Object -TypeName Xml.XmlNamespaceManager -ArgumentList $doc.NameTable
    [void]$ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    [void]$ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
    [void]$ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    return ,$ns
}

# ---------- 1. Текст, формат, гиперссылка, картинка ----------

$pdf = Join-Path $out 'pptxsrc.pdf'
$doc = New-Object PdfSharp.Pdf.PdfDocument
$page = $doc.AddPage()
$page.Width = New-Object PdfSharp.Drawing.XUnit(720.0)
$page.Height = New-Object PdfSharp.Drawing.XUnit(540.0)
$g = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
$title = New-Object PdfSharp.Drawing.XFont('Arial', 24.0, [PdfSharp.Drawing.XFontStyle]::Bold)
$body  = New-Object PdfSharp.Drawing.XFont('Arial', 12.0)
$g.DrawString('Slide title', $title, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(60.0, 90.0)))
# Второй блок — далеко от заголовка: две строки подряд разбор законно считает одним
# абзацем (единственный зазор не с чем сравнить), а на слайде это разные надписи.
$g.DrawString('Привет, мир', $body, [PdfSharp.Drawing.XBrushes]::Red, (New-Object PdfSharp.Drawing.XPoint(60.0, 320.0)))
$png = Join-Path $out 'pptx_dot.png'
$bmp = New-Object System.Drawing.Bitmap(40, 30)
$gr = [System.Drawing.Graphics]::FromImage($bmp)
$gr.Clear([System.Drawing.Color]::CornflowerBlue)
# Одноцветную картинку разбор отбраковывает намеренно (пустая — значит показывать нечего),
# поэтому в тестовой обязан быть хоть какой-то рисунок.
$gr.FillRectangle([System.Drawing.Brushes]::Orange, 5, 5, 20, 15)
$gr.Dispose()
$bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
$img = [PdfSharp.Drawing.XImage]::FromFile($png)
$g.DrawImage($img, 400.0, 300.0, 80.0, 60.0)
$img.Dispose()
$g.Dispose()
# Рамка ссылки — на всю страницу: слово точно внутри, и тест не зависит от того, как
# PdfSharp пересчитывает координаты рисования в координаты страницы.
$linkRect = New-Object PdfSharp.Drawing.XRect(0.0, 0.0, 720.0, 540.0)
[void]$page.AddWebLink((New-Object PdfSharp.Pdf.PdfRectangle($linkRect)), 'https://example.org/a?b=1&c=2')
$doc.Save($pdf)
$doc.Dispose()

$pptx = Join-Path $out 'pptxout.pptx'
if (Test-Path $pptx) { Remove-Item $pptx -Force }
$res = [ExcelMerger.PdfToPptxService]::Convert((New-PageOrder $pdf), $pptx, $null, $null, $null)

if ($res.Pages -ne 1) { Fail "страниц в результате: $($res.Pages)" }
if (-not (Test-Path $pptx)) { Fail 'файл презентации не создан' }

$names = Get-Names $pptx
if ($names[0] -ne '[Content_Types].xml') { Fail "первая запись архива: $($names[0])" }
$slides = @($names | Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' })
if ($slides.Count -ne 1) { Fail "слайдов в пакете: $($slides.Count)" }

$slide = Get-Xml $pptx 'ppt/slides/slide1.xml'
$ns = New-Ns $slide
$texts = @($slide.SelectNodes('//a:t', $ns) | ForEach-Object { $_.InnerText })
$all = [string]::Join(' ', $texts)
if ($all -notmatch 'Slide title') { Fail 'заголовок не перенесён' }
if ($all -notmatch 'Привет') { Fail 'кириллица не перенесена' }

$shapes = @($slide.SelectNodes('//p:sp', $ns))
if ($shapes.Count -lt 2) { Fail "надписей на слайде: $($shapes.Count)" }
$pics = @($slide.SelectNodes('//p:pic', $ns))
$bg = @($slide.SelectNodes('//p:bg', $ns))
if ($pics.Count -lt 1 -and $bg.Count -lt 1) { Fail 'ни картинки, ни подложки на слайде' }

# кегль и цвет доехали
$sizes = @($slide.SelectNodes('//a:rPr/@sz', $ns) | ForEach-Object { [int]$_.Value })
if (-not ($sizes -contains 2400)) { Fail "кегль заголовка не 24 pt: $($sizes -join ',')" }
$colors = @($slide.SelectNodes('//a:srgbClr/@val', $ns) | ForEach-Object { $_.Value })
if (-not ($colors -contains 'FF0000')) { Fail "красный цвет не перенесён: $($colors -join ',')" }

# у надписи нулевые поля и выключенный автоподбор — иначе текст уезжает от подложки
$bodyPr = $slide.SelectSingleNode('//p:sp/p:txBody/a:bodyPr', $ns)
if ($null -eq $bodyPr) { Fail 'нет свойств текстового тела' }
elseif ($bodyPr.lIns -ne '0' -or $bodyPr.tIns -ne '0') { Fail "поля надписи не нулевые: lIns=$($bodyPr.lIns) tIns=$($bodyPr.tIns)" }

# ---------- 2. Согласованность пакета: типы, связи, ссылки ----------

$types = Get-Xml $pptx '[Content_Types].xml'
$tns = New-Object -TypeName Xml.XmlNamespaceManager -ArgumentList $types.NameTable
[void]$tns.AddNamespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types')
foreach ($ov in $types.SelectNodes('//ct:Override', $tns)) {
    $part = $ov.PartName
    if (-not $part.StartsWith('/')) { Fail "Override без ведущего слэша: $part" }
    elseif ($names -notcontains $part.Substring(1)) { Fail "Override указывает в никуда: $part" }
}

$rels = Get-Xml $pptx 'ppt/slides/_rels/slide1.xml.rels'
$rns = New-Object -TypeName Xml.XmlNamespaceManager -ArgumentList $rels.NameTable
[void]$rns.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
$declared = @($rels.SelectNodes('//rel:Relationship', $rns) | ForEach-Object { $_.Id })
foreach ($node in $slide.SelectNodes('//*[@r:id or @r:embed]', $ns)) {
    foreach ($attr in @($node.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
                        $node.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'))) {
        if ($attr -and $declared -notcontains $attr) { Fail "ссылка $attr не объявлена в связях слайда" }
    }
}
$external = @($rels.SelectNodes('//rel:Relationship[@TargetMode="External"]', $rns))
if ($external.Count -lt 1) { Fail 'гиперссылка не стала внешней связью' }
$relsText = Get-Entry $pptx 'ppt/slides/_rels/slide1.xml.rels'
if ($relsText -notmatch 'b=1&amp;c=2') { Fail 'амперсанд в адресе ссылки не экранирован' }

# ---------- 3. Повёрнутая страница ----------

$pdfRot = Join-Path $out 'pptxsrc_rot.pdf'
$doc2 = New-Object PdfSharp.Pdf.PdfDocument
$p2 = $doc2.AddPage()
$p2.Width = New-Object PdfSharp.Drawing.XUnit(595.0)
$p2.Height = New-Object PdfSharp.Drawing.XUnit(842.0)
$g2 = [PdfSharp.Drawing.XGraphics]::FromPdfPage($p2)
$g2.RotateAtTransform(90, (New-Object PdfSharp.Drawing.XPoint(300.0, 400.0)))
$g2.DrawString('Rotated page', $body, [PdfSharp.Drawing.XBrushes]::Black, (New-Object PdfSharp.Drawing.XPoint(120.0, 400.0)))
$g2.Dispose()
$doc2.Save($pdfRot)
$doc2.Dispose()

$pptxRot = Join-Path $out 'pptxout_rot.pptx'
if (Test-Path $pptxRot) { Remove-Item $pptxRot -Force }
[void][ExcelMerger.PdfToPptxService]::Convert((New-PageOrder $pdfRot 270), $pptxRot, $null, $null, $null)
$pres = Get-Xml $pptxRot 'ppt/presentation.xml'
$pns = New-Ns $pres
$sz = $pres.SelectSingleNode('//p:sldSz', $pns)
if ($null -eq $sz) { Fail 'нет размера слайда' }
elseif ([long]$sz.cx -le [long]$sz.cy) { Fail "повёрнутая страница дала книжный слайд: $($sz.cx)x$($sz.cy)" }
$rotSlide = Get-Xml $pptxRot 'ppt/slides/slide1.xml'
$rotNs = New-Ns $rotSlide
$rotText = [string]::Join(' ', @($rotSlide.SelectNodes('//a:t', $rotNs) | ForEach-Object { $_.InnerText }))
if ($rotText -notmatch 'Rotated') { Fail 'текст повёрнутой страницы потерян' }

# ---------- 4. Отмена не оставляет файла ----------

$pptxCancel = Join-Path $out 'pptxout_cancel.pptx'
if (Test-Path $pptxCancel) { Remove-Item $pptxCancel -Force }
$cancel = [Func[bool]] { return $true }
try {
    [void][ExcelMerger.PdfToPptxService]::Convert((New-PageOrder $pdf), $pptxCancel, $null, $cancel, $null)
    Fail 'отмена не сработала'
} catch {
    $root = $_.Exception
    while ($root.InnerException) { $root = $root.InnerException }
    if ($root -isnot [OperationCanceledException]) { Fail "при отмене прилетело $($root.GetType().Name)" }
}
if (Test-Path $pptxCancel) { Fail 'после отмены остался файл результата' }
if (Test-Path ($pptxCancel + '.tmp')) { Fail 'после отмены остался временный файл' }

# ---------- 5. Если PowerPoint установлен — пусть скажет своё слово ----------

$ppt = $null
try { $ppt = New-Object -ComObject PowerPoint.Application } catch { }
if ($ppt) {
    try {
        $opened = $ppt.Presentations.Open($pptx, -1, 0, 0)
        if ($opened.Slides.Count -ne 1) { Fail "PowerPoint насчитал слайдов: $($opened.Slides.Count)" }
        if ($opened.Slides(1).Shapes.Count -lt 2) { Fail "PowerPoint насчитал фигур: $($opened.Slides(1).Shapes.Count)" }
        $opened.Close()
    } catch {
        Fail "PowerPoint не открыл файл: $($_.Exception.Message)"
    } finally {
        $ppt.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
} else {
    Write-Host 'PowerPoint не установлен — проверка открытия пропущена (файл проверен разбором архива)'
}

if ($fails.Count -eq 0) { Write-Host 'VERIFY PDFPPTX OK'; exit 0 }
Write-Host "VERIFY PDFPPTX FAILED: $($fails.Count)"
exit 1
