# Полная пирамида тестов одним запуском: сборка -> юнит -> смоук GUI ->
# генерация корпуса -> два интеграционных прогона с проверками -> зомби-процессы.
# Требует установленный Excel. Код выхода 0 = всё прошло.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'

function Invoke-Exe([string]$argLine) {
    # GUI-subsystem процесс: PowerShell без -Wait не дождался бы кода выхода.
    $p = Start-Process -FilePath $exe -ArgumentList $argLine -Wait -PassThru
    return $p.ExitCode
}

function Step([string]$name, [scriptblock]$body) {
    Write-Host ("--- " + $name)
    & $body
}

Step 'Сборка приложения' {
    cmd /c "`"$root\build.cmd`""
    if ($LASTEXITCODE) { Write-Host 'BUILD FAILED'; exit 1 }
}

Step 'Юнит-тесты' {
    cmd /c "`"$PSScriptRoot\build_tests.cmd`""
    if ($LASTEXITCODE) { Write-Host 'UNIT TESTS FAILED'; exit 1 }
}

Step 'Смоук-тест GUI' {
    $code = Invoke-Exe '--selftest'
    if ($code) { Write-Host "SELFTEST FAILED ($code)"; exit 1 }
}

Step 'Генерация тестового корпуса' {
    powershell -NoProfile -File "$PSScriptRoot\make_testdata.ps1" | Out-Null
    if ($LASTEXITCODE) { Write-Host 'TESTDATA FAILED'; exit 1 }
}

Remove-Item -Recurse -Force "$PSScriptRoot\out" -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force "$PSScriptRoot\out" | Out-Null

Step 'Базовый прогон (корпус содержит битые файлы: ожидается код 2)' {
    $code = Invoke-Exe "--cli `"$PSScriptRoot\testdata`" `"$PSScriptRoot\out\Свод.xlsx`""
    if ($code -ne 2) { Write-Host "BASE RUN exit=$code, ожидалось 2"; exit 1 }
    powershell -NoProfile -File "$PSScriptRoot\verify.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Прогон с оглавлением и заменой формул' {
    $code = Invoke-Exe "--cli `"$PSScriptRoot\testdata`" `"$PSScriptRoot\out\Свод_toc.xlsx`" --toc --values"
    if ($code -ne 2) { Write-Host "TOC RUN exit=$code, ожидалось 2"; exit 1 }
    powershell -NoProfile -File "$PSScriptRoot\verify_toc.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Прогон в формат .xlsb' {
    $code = Invoke-Exe "--cli `"$PSScriptRoot\testdata`" `"$PSScriptRoot\out\Свод_b.xlsb`""
    if ($code -ne 2) { Write-Host "XLSB RUN exit=$code, ожидалось 2"; exit 1 }
    powershell -NoProfile -File "$PSScriptRoot\verify_format.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Дослияние пропущенных в существующий свод' {
    powershell -NoProfile -File "$PSScriptRoot\verify_retry.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Вшитый PdfSharp резолвится из ресурса exe' {
    powershell -NoProfile -File "$PSScriptRoot\verify_embedded.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Объединение PDF' {
    powershell -NoProfile -File "$PSScriptRoot\verify_pdf.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Записка Word' {
    powershell -NoProfile -File "$PSScriptRoot\verify_note.ps1"
    if ($LASTEXITCODE) { exit 1 }
}

Step 'Зомби-процессы Excel' {
    Start-Sleep -Seconds 3
    if (Get-Process EXCEL -ErrorAction SilentlyContinue) { Write-Host 'ZOMBIE EXCEL'; exit 1 }
}

Write-Host 'ALL TESTS PASSED'
exit 0
