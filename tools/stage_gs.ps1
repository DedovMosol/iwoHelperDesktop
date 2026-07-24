# Copies a minimal Ghostscript runtime into installer\gs\ (x64) or installer\gs32\ (x86)
# for bundling in the matching installer. Needed: bin\gsdll64.dll|gsdll32.dll (engine) +
# bin\gswin64c.exe|gswin32c.exe (loader) + lib\ + Resource\ + iccprofiles\ (no import
# .lib - not needed at runtime). Source is an installed GS of the SAME bitness
# (Program Files\gs\gs*, Program Files (x86)\gs\gs* or %USERPROFILE%\gs*); can be set
# via -Source. Both staging dirs are in .gitignore (large binaries are not committed).
param(
    [string]$Source,
    [string]$Dest,
    [ValidateSet('x64', 'x86')]
    [string]$Arch = 'x64'
)
$ErrorActionPreference = 'Stop'

# Per-bitness file names and staging dir: the app looks for gswin64c.exe AND gswin32c.exe
# next to itself, so each installer bundles the binaries matching its package.
$dllName = if ($Arch -eq 'x86') { 'gsdll32.dll' } else { 'gsdll64.dll' }
$exeName = if ($Arch -eq 'x86') { 'gswin32c.exe' } else { 'gswin64c.exe' }
if (-not $Dest) {
    $dirName = if ($Arch -eq 'x86') { 'installer\gs32' } else { 'installer\gs' }
    $Dest = Join-Path (Split-Path $PSScriptRoot) $dirName
}

function Find-GsRoot {
    $roots = @()

    # 1) Registry first - reliable regardless of the install drive (software on this
    #    kind of machine may live outside %ProgramFiles%, e.g. on another disk). The
    #    32-bit GS on a 64-bit Windows registers under WOW6432Node; GS_DLL points at
    #    bin\gsdll<NN>.dll, two levels below the root we need.
    $regBases = if ($Arch -eq 'x86') {
        @('HKLM:\SOFTWARE\WOW6432Node\GPL Ghostscript', 'HKLM:\SOFTWARE\WOW6432Node\Artifex Ghostscript')
    } else {
        @('HKLM:\SOFTWARE\GPL Ghostscript', 'HKLM:\SOFTWARE\Artifex Ghostscript')
    }
    foreach ($rb in $regBases) {
        Get-ChildItem $rb -ErrorAction SilentlyContinue | ForEach-Object {
            $dllPath = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).GS_DLL
            if ($dllPath) {
                $r = Split-Path (Split-Path $dllPath)
                if (Test-Path (Join-Path $r "bin\$dllName")) { $roots += $r }
            }
        }
    }

    # 2) Fallback - Program Files on every LOCAL drive (network drives are skipped:
    #    scanning a dead share hangs) plus the user profile.
    $bases = @()
    foreach ($d in ([IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady })) {
        $bases += (Join-Path $d.Name 'Program Files\gs')
        $bases += (Join-Path $d.Name 'Program Files (x86)\gs')
    }
    if ($env:USERPROFILE) { $bases += $env:USERPROFILE }
    foreach ($b in $bases) {
        if (Test-Path $b) {
            Get-ChildItem -LiteralPath $b -Directory -Filter 'gs*' -ErrorAction SilentlyContinue | ForEach-Object {
                if (Test-Path (Join-Path $_.FullName "bin\$dllName")) { $roots += $_.FullName }
            }
        }
    }

    # Newest version by the FOLDER name (gs10.07 > gs10.05 > gs9.56) - a full-path sort
    # would compare drive letters first and could prefer an older install.
    return ($roots | Sort-Object { Split-Path $_ -Leaf } -Descending -Unique | Select-Object -First 1)
}

if (-not $Source) { $Source = Find-GsRoot }
if (-not $Source -or -not (Test-Path $Source)) {
    Write-Host "Ghostscript ($Arch) not found. Install the matching build: https://ghostscript.com/releases/gsdnld.html"
    Write-Host "or pass -Source <gs root, e.g. C:\Program Files\gs\gs10.07>."
    exit 1
}

$dll = Join-Path $Source "bin\$dllName"
$exe = Join-Path $Source "bin\$exeName"
if (-not (Test-Path $dll) -or -not (Test-Path $exe)) {
    Write-Host "Source has no bin\$dllName / bin\$exeName (need a $Arch Ghostscript): $Source"
    exit 1
}

Write-Host "Ghostscript source ($Arch): $Source"
if (Test-Path $Dest) { Remove-Item $Dest -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $Dest 'bin') -Force | Out-Null

Copy-Item $dll (Join-Path $Dest "bin\$dllName") -Force
Copy-Item $exe (Join-Path $Dest "bin\$exeName") -Force
foreach ($sub in @('lib', 'Resource', 'iccprofiles')) {
    $s = Join-Path $Source $sub
    if (Test-Path $s) {
        Copy-Item $s (Join-Path $Dest $sub) -Recurse -Force
    } elseif ($sub -ne 'iccprofiles') {
        Write-Host "WARNING: required directory '$sub' is missing in $Source"
        exit 1
    }
}
# GS license (AGPL) - bundled alongside for compliance.
foreach ($lic in @('LICENSE', 'doc\COPYING', 'COPYING')) {
    $lp = Join-Path $Source $lic
    if (Test-Path $lp) { Copy-Item $lp (Join-Path $Dest 'LICENSE') -Force; break }
}

$size = [math]::Round(((Get-ChildItem $Dest -Recurse -File | Measure-Object Length -Sum).Sum / 1MB), 1)
Write-Host "Ghostscript ($Arch) staged in $Dest ($size MB)."
exit 0
