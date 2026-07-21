# Integration test for the "all sheets" mode: a file with three sheets is merged
# with --allsheets, and the digest must contain all three named "file · sheet".
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$dir = Join-Path $PSScriptRoot 'allsheets_in'
$out = Join-Path $PSScriptRoot 'out\Свод_all.xlsx'
$fails = @()

Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $dir | Out-Null
Remove-Item $out -Force -ErrorAction SilentlyContinue

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Sheets.Count -lt 3) { $wb.Sheets.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count)) | Out-Null }
    $names = @('Янв','Фев','Мар')
    for ($i = 0; $i -lt 3; $i++) {
        $sh = $wb.Sheets.Item($i + 1)
        $sh.Name = $names[$i]
        $sh.Range('A1').Value2 = "данные " + $names[$i]
    }
    $wb.SaveAs((Join-Path $dir 'Отчет 3л.xlsx'), 51)
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}

$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'
$p = Start-Process -FilePath $exe -ArgumentList "--cli `"$dir`" `"$out`" --allsheets" -Wait -PassThru
if ($p.ExitCode -ne 0) { $fails += "код выхода $($p.ExitCode), ожидался 0" }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($out, 0, $true)
    $sheetNames = @(); foreach ($s in $wb.Sheets) { $sheetNames += $s.Name }
    $expected = @('Отчет 3л · Янв','Отчет 3л · Фев','Отчет 3л · Мар')
    if (($sheetNames -join ';') -ne ($expected -join ';')) {
        $fails += "листы: [$($sheetNames -join ', ')], ожидалось [$($expected -join ', ')]"
    }
    if ($wb.Sheets.Item('Отчет 3л · Фев').Range('A1').Value2 -ne 'данные Фев') {
        $fails += "содержимое листа Фев не совпало"
    }
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}

Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue

if ($fails.Count -eq 0) {
    Write-Host "VERIFY ALLSHEETS OK"
    exit 0
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    exit 1
}
