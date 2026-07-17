# Интеграционный тест объединения PDF. Эталонные PDF со страницами разных
# размеров создаёт сам PdfSharp (вшит в exe) — без Word, детерминированно;
# порядок и дубль страницы в склейке проверяются по габаритам, битый файл —
# по понятной ошибке.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$out = Join-Path $PSScriptRoot 'out'
$fails = @()

# Логика под тестом — во вшитой в exe копии PdfSharp; отдельная копия из build/
# нужна лишь тесту, чтобы СОЗДАТЬ входные PDF (наш AssemblyResolve срабатывает
# только при вызове кода приложения).
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'build\PdfSharp.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'dist\iwoHelperDesktop.exe'))

function New-Pdf([string]$path, [object[]]$pages) {
    # Размеры в пунктах (XUnit неявно принимает double): без enum-типов PdfSharp.
    $doc = New-Object PdfSharp.Pdf.PdfDocument
    foreach ($p in $pages) {
        $page = $doc.AddPage()
        $page.Width = [double]$p.W
        $page.Height = [double]$p.H
    }
    $doc.Save($path)
    $doc.Dispose()
}

$pdfA = Join-Path $out 'A.pdf'   # стр.1 — A4 книжная (595x842), стр.2 — A5 (420x595)
$pdfB = Join-Path $out 'B.pdf'   # A4 альбомная (842x595)
Remove-Item $pdfA, $pdfB -Force -ErrorAction SilentlyContinue
New-Pdf $pdfA @(
    @{ W = 595; H = 842 },
    @{ W = 420; H = 595 })
New-Pdf $pdfB @(@{ W = 842; H = 595 })

function New-PageRef([string]$path, [int]$index) {
    $ref = New-Object ExcelMerger.PdfPageRef
    $ref.SourcePath = $path
    $ref.PageIndex = $index
    return $ref
}
function Describe($w, $h) {
    if ($w -gt $h) { return 'landscape' }
    if ($w -lt 500) { return 'a5' }
    return 'a4' }

$pagesA = [ExcelMerger.PdfMergeService]::LoadPages($pdfA)
if ($pagesA.Count -ne 2) { $fails += "в A.pdf $($pagesA.Count) страниц, ожидалось 2" }

# Склейка с перестановкой и дублем: [B:1, A:2, A:1, A:1]
$orderList = New-Object 'System.Collections.Generic.List[ExcelMerger.PdfPageRef]'
$orderList.Add((New-PageRef $pdfB 0))
$orderList.Add((New-PageRef $pdfA 1))
$orderList.Add((New-PageRef $pdfA 0))
$orderList.Add((New-PageRef $pdfA 0))
$merged = Join-Path $out 'Объединённый.pdf'
Remove-Item $merged -Force -ErrorAction SilentlyContinue
[ExcelMerger.PdfMergeService]::Merge($orderList, $merged)

$pages = [ExcelMerger.PdfMergeService]::LoadPages($merged)
if ($pages.Count -ne 4) { $fails += "в склейке $($pages.Count) страниц, ожидалось 4" }
$shapes = @($pages | ForEach-Object { Describe $_.WidthPt $_.HeightPt }) -join '|'
if ($shapes -ne 'landscape|a5|a4|a4') { $fails += "порядок страниц: $shapes, ожидалось landscape|a5|a4|a4" }

# Битый PDF — понятная ошибка, а не краш
$broken = Join-Path $out 'битый.pdf'
Set-Content -Path $broken -Value 'это не pdf' -Encoding ASCII
try {
    [void][ExcelMerger.PdfMergeService]::LoadPages($broken)
    $fails += 'битый PDF не дал ошибку'
} catch {
    $msg = $_.Exception.InnerException.Message
    if ($msg -notmatch 'повреждён') { $fails += "неожиданный текст ошибки: $msg" }
}

if ($fails.Count -eq 0) {
    Write-Host "VERIFY PDF OK"
    exit 0
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    exit 1
}
