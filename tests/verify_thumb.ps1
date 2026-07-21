# Integration test for PDF thumbnails: rendering a page via WinRT
# (Windows.Data.Pdf) yields a non-empty Bitmap with the correct aspect ratio.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$out = Join-Path $PSScriptRoot 'out'
$fails = @()

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'build\PdfSharp.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'dist\iwoHelperDesktop.exe'))

# A4 portrait (595x842) and landscape (842x595) pages via PdfSharp
$pdf = Join-Path $out 'thumb.pdf'
Remove-Item $pdf -Force -ErrorAction SilentlyContinue
$doc = New-Object PdfSharp.Pdf.PdfDocument
$p1 = $doc.AddPage(); $p1.Width = [double]595; $p1.Height = [double]842
$p2 = $doc.AddPage(); $p2.Width = [double]842; $p2.Height = [double]595
$doc.Save($pdf); $doc.Dispose()

$r = New-Object ExcelMerger.PdfThumbnailRenderer
try {
    $b0 = $r.Render($pdf, 0, 120)
    if ($b0 -eq $null) { $fails += 'page 0 did not render (WinRT unavailable?)' }
    else {
        if ($b0.Width -ne 120) { $fails += "thumbnail width $($b0.Width), expected 120" }
        $ratio = $b0.Height / $b0.Width
        if ([math]::Abs($ratio - 1.414) -gt 0.05) { $fails += "A4 ratio $ratio, expected ~1.414" }
        $b0.Dispose()
    }
    $b1 = $r.Render($pdf, 1, 120)
    if ($b1 -ne $null) {
        if ($b1.Height -ge $b1.Width) { $fails += "landscape: H=$($b1.Height) >= W=$($b1.Width)" }
        $b1.Dispose()
    }
    # non-existent page -> null, no exception
    $bad = $r.Render($pdf, 99, 120)
    if ($bad -ne $null) { $fails += 'page 99 should be null' }
}
finally {
    $r.Dispose()
}

if ($fails.Count -eq 0) {
    Write-Host "VERIFY THUMB OK"
    # Hard exit without finalizers: WinRT in a PowerShell STA host would otherwise crash
    # the process on unload (in the app rendering runs on a background MTA thread).
    [Environment]::Exit(0)
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    [Environment]::Exit(1)
}
