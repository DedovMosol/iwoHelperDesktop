# Sign iwoHelperDesktop.exe with a self-signed certificate (SHA256 + timestamp).
# Everything is built into Windows PowerShell 5.1 - no external tools required.
# The certificate is created once and reused (Cert:\CurrentUser\My).
# Note: a self-signed signature gives integrity and a stable publisher identity
# (reputation with antivirus); it becomes "trusted" by Windows only after the
# certificate is added to the Trusted Root store on the target machine.
param(
    [string]$ExePath = (Join-Path (Split-Path $PSScriptRoot) 'dist\iwoHelperDesktop.exe')
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $ExePath)) {
    Write-Host "File not found: $ExePath (run build.cmd first)"
    exit 1
}

$subject = 'CN=Svod Excel self-signed'
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
    Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
    Sort-Object NotAfter -Descending | Select-Object -First 1
if (-not $cert) {
    Write-Host 'Creating a new code-signing certificate...'
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject `
        -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5) `
        -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
}

# Timestamp so the signature outlives the certificate; if the server is
# unreachable (proxy), sign without it.
$result = $null
try {
    $result = Set-AuthenticodeSignature -FilePath $ExePath -Certificate $cert `
        -HashAlgorithm SHA256 -TimestampServer 'http://timestamp.digicert.com'
} catch {
    Write-Host 'Timestamp server unavailable - signing without a timestamp.'
    $result = Set-AuthenticodeSignature -FilePath $ExePath -Certificate $cert -HashAlgorithm SHA256
}

if ($result.SignerCertificate -eq $null) {
    Write-Host "SIGNING FAILED: $($result.StatusMessage)"
    exit 1
}
# For a self-signed certificate the status is UnknownError ("chain not trusted") -
# expected; the signature itself is valid and present in the file.
Write-Host ("Signed:  " + $result.SignerCertificate.Subject)
Write-Host ("Status:  " + $result.Status + " (" + $result.StatusMessage + ")")
Write-Host ("SHA256:  " + (Get-FileHash $ExePath -Algorithm SHA256).Hash)
exit 0
