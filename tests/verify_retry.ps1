# Интеграционный тест «Повторить пропущенные»: exe загружается как .NET-сборка,
# сценарий — слияние с пропусками, починка битого файла, дослияние в свод.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$fails = @()

# Add-Type не принимает .exe — загружаем сборку напрямую.
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'dist\iwoHelperDesktop.exe'))

# Отдельная копия корпуса: битый файл будем чинить, общий корпус не трогаем.
$data = Join-Path $PSScriptRoot 'retrydata'
Remove-Item -Recurse -Force $data -ErrorAction SilentlyContinue
Copy-Item (Join-Path $PSScriptRoot 'testdata') $data -Recurse
$out = Join-Path $PSScriptRoot 'out\Свод_retry.xlsx'
Remove-Item $out -Force -ErrorAction SilentlyContinue

$options = New-Object ExcelMerger.MergeOptions
$options.AddToc = $true

$service = New-Object ExcelMerger.MergeService
$files = [ExcelMerger.MergeService]::FindSourceFiles($data, $out)
$first = $service.Merge($files, $out, $options)
if ($first.OkCount -ne 10 -or $first.SkipCount -ne 2) {
    $fails += "первый прогон: ok=$($first.OkCount) skip=$($first.SkipCount), ожидалось 10/2"
}

# Починка: битый файл заменяется валидной книгой (лист получит имя файла).
Copy-Item (Join-Path $data 'Отчет 2.xlsx') (Join-Path $data 'Битый файл.xlsx') -Force

$service2 = New-Object ExcelMerger.MergeService
$second = $service2.RetrySkipped($out, $options, $first)
if ($second.OkCount -ne 11) { $fails += "после повтора ok=$($second.OkCount), ожидалось 11" }
if ($second.SkipCount -ne 1) { $fails += "после повтора skip=$($second.SkipCount), ожидалось 1 (пароль)" }
if ($second.Files.Count -ne 12) { $fails += "записей $($second.Files.Count), ожидалось 12" }
if (-not $second.Files[0].Ok) { $fails += "«Битый файл» не стал перенесённым" }

# Содержимое свода после дослияния.
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($out, 0, $true)
    if ($wb.Sheets.Count -ne 12) { $fails += "листов $($wb.Sheets.Count), ожидалось 12 (11 + Содержание)" }
    if ($wb.Sheets.Item(1).Name -ne 'Содержание') { $fails += "первый лист не «Содержание»" }
    $toc = $wb.Sheets.Item(1)
    if ($toc.Hyperlinks.Count -ne 11) { $fails += "ссылок в оглавлении $($toc.Hyperlinks.Count), ожидалось 11" }
    if ($toc.Range("C2").Value2 -ne 'Битый файл.xlsx') { $fails += "строка 2 оглавления: $($toc.Range('C2').Value2)" }
    if ($toc.Range("B2").Value2 -ne 'Битый файл') { $fails += "починенный файл без листа в оглавлении: $($toc.Range('B2').Value2)" }
    $sheet = $wb.Sheets.Item('Битый файл')
    if ($sheet.Range("A1").Value2 -ne 'второй') { $fails += "содержимое дослитого листа: $($sheet.Range('A1').Value2)" }
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}

Remove-Item -Recurse -Force $data -ErrorAction SilentlyContinue

if ($fails.Count -eq 0) {
    Write-Host "VERIFY RETRY OK"
    exit 0
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    exit 1
}
