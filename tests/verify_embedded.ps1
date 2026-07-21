# Verify the embedded PdfSharp resolves from the exe resource with no DLL alongside.
# The exe is copied into an empty folder and run there with --pdfcheck.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$tmp = Join-Path $env:TEMP ('emb_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
try {
    Copy-Item (Join-Path $root 'dist\iwoHelperDesktop.exe') (Join-Path $tmp 'iwoHelperDesktop.exe')
    $p = Start-Process -FilePath (Join-Path $tmp 'iwoHelperDesktop.exe') -ArgumentList '--pdfcheck' `
        -Wait -PassThru -WorkingDirectory $tmp
    if ($p.ExitCode -eq 0) {
        Write-Host "VERIFY EMBEDDED OK"
        exit 0
    }
    Write-Host "FAIL: --pdfcheck exit $($p.ExitCode) - embedded PdfSharp did not resolve"
    exit 1
}
finally {
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
