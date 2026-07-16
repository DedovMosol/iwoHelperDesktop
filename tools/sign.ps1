# Подпись ExcelMerger.exe самоподписанным сертификатом (SHA256 + метка времени).
# Всё встроено в Windows PowerShell 5.1 — внешних инструментов не требуется.
# Сертификат создаётся один раз и переиспользуется (Cert:\CurrentUser\My).
# Примечание: самоподписанная подпись даёт целостность и постоянного издателя
# (репутация для антивирусов); «доверенной» для Windows она станет только после
# добавления сертификата в доверенные корневые на целевой машине.
param(
    [string]$ExePath = (Join-Path (Split-Path $PSScriptRoot) 'ExcelMerger.exe')
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $ExePath)) {
    Write-Host "Файл не найден: $ExePath (сначала build.cmd)"
    exit 1
}

$subject = 'CN=Svod Excel self-signed'
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
    Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
    Sort-Object NotAfter -Descending | Select-Object -First 1
if (-not $cert) {
    Write-Host 'Создаётся новый сертификат подписи кода…'
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject `
        -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5) `
        -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
}

# Метка времени — чтобы подпись жила дольше сертификата; при недоступности
# сервера (прокси) подписываем без неё.
$result = $null
try {
    $result = Set-AuthenticodeSignature -FilePath $ExePath -Certificate $cert `
        -HashAlgorithm SHA256 -TimestampServer 'http://timestamp.digicert.com'
} catch {
    Write-Host 'Сервер меток времени недоступен — подпись без метки.'
    $result = Set-AuthenticodeSignature -FilePath $ExePath -Certificate $cert -HashAlgorithm SHA256
}

if ($result.SignerCertificate -eq $null) {
    Write-Host "ПОДПИСЬ НЕ ПОСТАВЛЕНА: $($result.StatusMessage)"
    exit 1
}
# Для самоподписанного сертификата статус UnknownError («цепочка не доверена») —
# ожидаем; сама подпись при этом корректна и присутствует в файле.
Write-Host ("Подписано: " + $result.SignerCertificate.Subject)
Write-Host ("Статус:    " + $result.Status + " (" + $result.StatusMessage + ")")
Write-Host ("SHA256:    " + (Get-FileHash $ExePath -Algorithm SHA256).Hash)
exit 0
