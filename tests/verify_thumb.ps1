# Интеграционный тест миниатюр PDF: рендер страницы через WinRT
# (Windows.Data.Pdf) даёт непустой Bitmap с правильным соотношением сторон.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$out = Join-Path $PSScriptRoot 'out'
$fails = @()

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'build\PdfSharp.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'dist\iwoHelperDesktop.exe'))

# A4-книжная (595x842) и альбомная (842x595) страницы через PdfSharp
$pdf = Join-Path $out 'thumb.pdf'
Remove-Item $pdf -Force -ErrorAction SilentlyContinue
$doc = New-Object PdfSharp.Pdf.PdfDocument
$p1 = $doc.AddPage(); $p1.Width = [double]595; $p1.Height = [double]842
$p2 = $doc.AddPage(); $p2.Width = [double]842; $p2.Height = [double]595
$doc.Save($pdf); $doc.Dispose()

$r = New-Object ExcelMerger.PdfThumbnailRenderer
try {
    $b0 = $r.Render($pdf, 0, 120)
    if ($b0 -eq $null) { $fails += 'страница 0 не отрендерилась (WinRT недоступен?)' }
    else {
        if ($b0.Width -ne 120) { $fails += "ширина миниатюры $($b0.Width), ожидалось 120" }
        $ratio = $b0.Height / $b0.Width
        if ([math]::Abs($ratio - 1.414) -gt 0.05) { $fails += "соотношение A4 $ratio, ожидалось ~1.414" }
        $b0.Dispose()
    }
    $b1 = $r.Render($pdf, 1, 120)
    if ($b1 -ne $null) {
        if ($b1.Height -ge $b1.Width) { $fails += "альбомная: H=$($b1.Height) >= W=$($b1.Width)" }
        $b1.Dispose()
    }
    # несуществующая страница -> null, без исключения
    $bad = $r.Render($pdf, 99, 120)
    if ($bad -ne $null) { $fails += 'страница 99 должна быть null' }
}
finally {
    $r.Dispose()
}

if ($fails.Count -eq 0) {
    Write-Host "VERIFY THUMB OK"
    # Жёсткий выход без финализаторов: WinRT в STA-хосте PowerShell иначе роняет
    # процесс при выгрузке (в приложении рендер идёт на фоновом MTA-потоке).
    [Environment]::Exit(0)
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    [Environment]::Exit(1)
}
