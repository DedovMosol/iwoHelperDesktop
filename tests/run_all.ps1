# Full test pyramid in one run: build -> unit -> GUI smoke -> corpus generation ->
# two integration runs with checks -> zombie-process check.
# Requires an installed Excel. Exit code 0 = everything passed.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'

function Invoke-Exe([string]$argLine) {
    # GUI-subsystem process: without -Wait, PowerShell would not wait for the exit code.
    $p = Start-Process -FilePath $exe -ArgumentList $argLine -Wait -PassThru
    return $p.ExitCode
}

function Step([string]$name, [scriptblock]$body) {
    Write-Host ("--- " + $name)
    & $body
}

Step 'Build the application' {
    cmd /c "`"$root\build.cmd`""
    if ($LASTEXITCODE) { Write-Host 'BUILD FAILED'; exit 1 }
}

Step 'Unit tests' {
    cmd /c "`"$PSScriptRoot\build_tests.cmd`""
    if ($LASTEXITCODE) { Write-Host 'UNIT TESTS FAILED'; exit 1 }
}

Step 'GUI smoke test' {
    $code = Invoke-Exe '--selftest'
    if ($code) { Write-Host "SELFTEST FAILED ($code)"; exit 1 }
}

Step 'Generate the test corpus' {
    powershell -NoProfile -File "$PSScriptRoot\make_testdata.ps1" | Out-Null
    if ($LASTEXITCODE) { Write-Host 'TESTDATA FAILED'; exit 1 }
}

Remove-Item -Recurse -Force "$PSScriptRoot\out" -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force "$PSScriptRoot\out" | Out-Null

Step 'Base run (corpus contains broken files: expecting code 2)' {
    $code = Invoke-Exe "--cli `"$PSScriptRoot\testdata`" `"$PSScriptRoot\out\Свод.xlsx`""
    if ($code -ne 2) { Write-Host "BASE RUN exit=$code, expected 2"; exit 1 }
    powershell -NoProfile -File "$PSScriptRoot\verify.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'All-sheets mode' {
    powershell -NoProfile -File "$PSScriptRoot\verify_allsheets.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Run with table of contents and formula-to-value' {
    $code = Invoke-Exe "--cli `"$PSScriptRoot\testdata`" `"$PSScriptRoot\out\Свод_toc.xlsx`" --toc --values"
    if ($code -ne 2) { Write-Host "TOC RUN exit=$code, expected 2"; exit 1 }
    powershell -NoProfile -File "$PSScriptRoot\verify_toc.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Run to .xlsb format' {
    $code = Invoke-Exe "--cli `"$PSScriptRoot\testdata`" `"$PSScriptRoot\out\Свод_b.xlsb`""
    if ($code -ne 2) { Write-Host "XLSB RUN exit=$code, expected 2"; exit 1 }
    powershell -NoProfile -File "$PSScriptRoot\verify_format.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Append skipped files into an existing digest' {
    powershell -NoProfile -File "$PSScriptRoot\verify_retry.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Embedded PdfSharp resolves from the exe resource' {
    powershell -NoProfile -File "$PSScriptRoot\verify_embedded.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'PDF merge' {
    powershell -NoProfile -File "$PSScriptRoot\verify_pdf.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'PDF thumbnails (WinRT)' {
    powershell -NoProfile -File "$PSScriptRoot\verify_thumb.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Clean exit after thumbnail render' {
    $code = Invoke-Exe '--thumbcheck'
    if ($code -ne 0) { Write-Host "thumbcheck exit=$code"; exit 1 }
}

Step 'Word cover note' {
    powershell -NoProfile -File "$PSScriptRoot\verify_note.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Born-digital PDF -> Word (PdfPig + OcrLayout + WordDocxWriter)' {
    powershell -NoProfile -File "$PSScriptRoot\verify_pdfword.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Zombie Excel processes' {
    Start-Sleep -Seconds 3
    if (Get-Process EXCEL -ErrorAction SilentlyContinue) { Write-Host 'ZOMBIE EXCEL'; exit 1 }
}

Write-Host 'ALL TESTS PASSED'
exit 0
