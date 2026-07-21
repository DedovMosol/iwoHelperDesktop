# Integration test for PDF merge. Reference PDFs with pages of different sizes are
# created by PdfSharp itself (embedded in the exe) - no Word, deterministic; the
# order and a duplicated page in the merge are checked by dimensions, a broken file
# by a clear error.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$out = Join-Path $PSScriptRoot 'out'
$fails = @()

# The logic under test lives in the PdfSharp copy embedded in the exe; a separate
# copy from build/ is only needed by the test to CREATE input PDFs (our
# AssemblyResolve fires only when application code is invoked).
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'build\PdfSharp.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'dist\iwoHelperDesktop.exe'))

function New-Pdf([string]$path, [object[]]$pages) {
    # Sizes in points (XUnit implicitly accepts double): no PdfSharp enum types.
    $doc = New-Object PdfSharp.Pdf.PdfDocument
    foreach ($p in $pages) {
        $page = $doc.AddPage()
        $page.Width = [double]$p.W
        $page.Height = [double]$p.H
    }
    $doc.Save($path)
    $doc.Dispose()
}

$pdfA = Join-Path $out 'A.pdf'   # page 1 - A4 portrait (595x842), page 2 - A5 (420x595)
$pdfB = Join-Path $out 'B.pdf'   # A4 landscape (842x595)
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

# Merge with reordering and a duplicate: [B:1, A:2, A:1, A:1]
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

# Broken PDF - a clear error, not a crash
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
