# Local build of the SIGNED iwo Helper Desktop installer with bundled Ghostscript.
# Steps: build.cmd -> sign the exe -> stage_gs.ps1 -> ISCC -> sign setup.exe.
# -Arch x86 builds the 32-bit package (exe from dist\x86, 32-bit Ghostscript, the
# -x86 suffix in the setup name); default is x64. Signing uses a self-signed
# certificate (tools\sign.ps1), locally only (the cert lives in Cert:\CurrentUser\My
# and is not on CI). Requires an installed Inno Setup 6 (ISCC.exe) and a Ghostscript
# of the SAME bitness (for staging). Child .ps1 scripts run in a separate process,
# otherwise their `exit` would terminate this script too.
param(
    [switch]$SkipBuild,
    [string]$Iscc,
    [ValidateSet('x64', 'x86')]
    [string]$Arch = 'x64'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$distDir = if ($Arch -eq 'x86') { 'dist\x86' } else { 'dist' }
$suffix = if ($Arch -eq 'x86') { '-x86' } else { '' }
$exe = Join-Path $root "$distDir\iwoHelperDesktop.exe"
$ps = (Get-Process -Id $PID).Path  # current PowerShell host

function Invoke-Child([string]$script, [string[]]$scriptArgs) {
    & $ps -NoProfile -ExecutionPolicy Bypass -File $script @scriptArgs
    if ($LASTEXITCODE -ne 0) { throw "Error: $script (code $LASTEXITCODE)" }
}

function Find-Iscc {
    # Any Inno Setup version (6/7/...) on any drive. Order: registry
    # (InstallLocation - most reliable, independent of %ProgramFiles%) -> glob under
    # Program Files -> PATH. Pick the newest version.
    $found = @()

    $unins = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($k in $unins) {
        Get-ItemProperty $k -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like 'Inno Setup*' -and $_.InstallLocation } |
            ForEach-Object {
                $iscc = Join-Path $_.InstallLocation 'ISCC.exe'
                if (Test-Path $iscc) { $found += $iscc }
            }
    }

    $bases = @()
    if (${env:ProgramFiles(x86)}) { $bases += ${env:ProgramFiles(x86)} }
    if ($env:ProgramFiles) { $bases += $env:ProgramFiles }
    foreach ($b in $bases) {
        Get-ChildItem -LiteralPath $b -Directory -Filter 'Inno Setup *' -ErrorAction SilentlyContinue | ForEach-Object {
            $iscc = Join-Path $_.FullName 'ISCC.exe'
            if (Test-Path $iscc) { $found += $iscc }
        }
    }

    if ($found.Count) { return ($found | Sort-Object -Descending -Unique | Select-Object -First 1) }
    $cmd = Get-Command 'ISCC.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

# 1. Build the application (per architecture)
if (-not $SkipBuild) {
    & (Join-Path $root 'build.cmd') $Arch
    if ($LASTEXITCODE -ne 0) { throw 'build.cmd failed' }
}
if (-not (Test-Path $exe)) { throw "missing $exe (build first)" }

# 2. Sign the exe (the installer embeds the already-signed file)
Invoke-Child (Join-Path $PSScriptRoot 'sign.ps1') @('-ExePath', $exe)

# 3. Stage the bundled Ghostscript of the matching bitness
Invoke-Child (Join-Path $PSScriptRoot 'stage_gs.ps1') @('-Arch', $Arch)

# 4. Version (3 components) from the exe version - single source of truth
$ver = (Get-Item $exe).VersionInfo.FileVersion   # e.g. 1.13.0.0
$ver3 = ($ver -split '\.')[0..2] -join '.'

# 5. Compile the installer
if (-not $Iscc) { $Iscc = Find-Iscc }
if (-not $Iscc) {
    Write-Host 'ISCC.exe (Inno Setup) not found. Install it: https://jrsoftware.org/isdl.php'
    exit 1
}
& $Iscc "/DAppVersion=$ver3" "/DArch=$Arch" (Join-Path $root 'installer\iwoHelperDesktop.iss')
if ($LASTEXITCODE -ne 0) { throw 'ISCC failed' }

$setup = Join-Path $root ("dist\iwoHelperDesktop-setup-$ver3$suffix.exe")
if (-not (Test-Path $setup)) { throw "installer was not created: $setup" }

# 6. Sign the installer
Invoke-Child (Join-Path $PSScriptRoot 'sign.ps1') @('-ExePath', $setup)

Write-Host ""
Write-Host "OK: $setup"
exit 0
