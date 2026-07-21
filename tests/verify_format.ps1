# Verify the .xlsb binary digest: the file opens, the sheet set is complete,
# and the formula and formatting are present.
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'out\Свод_b.xlsb'
$fails = @()

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($out, 0, $true)
    if ($wb.FileFormat -ne 50) { $fails += "формат файла $($wb.FileFormat), ожидался 50 (xlExcel12)" }
    if ($wb.Sheets.Count -ne 10) { $fails += "листов $($wb.Sheets.Count), ожидалось 10" }
    $wsA = $wb.Sheets.Item('Отчет управления A_2')
    if ($wsA.Range("A4").Value2 -ne 30) { $fails += "A4: $($wsA.Range('A4').Value2), ожидалось 30" }
    if (-not $wsA.Range("A1").Font.Bold) { $fails += "A1 потеряла жирный шрифт" }
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}

if ($fails.Count -eq 0) {
    Write-Host "VERIFY FORMAT OK"
    exit 0
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    exit 1
}
