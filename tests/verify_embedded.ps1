# Verify the embedded assemblies resolve from the exe resource with no DLL alongside:
# the exe is copied into an empty folder and each probe is run there. --pdfcheck covers
# PdfSharp, --reviewcheck covers the first-touch PdfPig path of the Compare tool
# (the one that shipped broken in 1.18.5: JIT failed before any resolver could register).
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$tmp = Join-Path $env:TEMP ('emb_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
try {
    Copy-Item (Join-Path $root 'dist\iwoHelperDesktop.exe') (Join-Path $tmp 'iwoHelperDesktop.exe')
    foreach ($flag in @('--pdfcheck', '--reviewcheck')) {
        $p = Start-Process -FilePath (Join-Path $tmp 'iwoHelperDesktop.exe') -ArgumentList $flag `
            -Wait -PassThru -WorkingDirectory $tmp
        if ($p.ExitCode -ne 0) {
            Write-Host "FAIL: $flag exit $($p.ExitCode) - embedded assembly did not resolve"
            exit 1
        }
    }
    Write-Host "VERIFY EMBEDDED OK"
    exit 0
}
finally {
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
