# Copies a minimal Ghostscript runtime into installer\gs\ (x64) or installer\gs32\ (x86)
# for bundling in the matching installer. Needed: bin\gsdll64.dll|gsdll32.dll (engine) +
# bin\gswin64c.exe|gswin32c.exe (loader) + lib\ + Resource\ + iccprofiles\ (no import
# .lib - not needed at runtime). Source is an installed GS of the SAME bitness
# (Program Files\gs\gs*, Program Files (x86)\gs\gs* or %USERPROFILE%\gs*); can be set
# via -Source. Both staging dirs are in .gitignore (large binaries are not committed).
#
# The installer ships the engine that then reads every PDF the user opens, so an outdated
# build is a security defect, not a cosmetic one: Ghostscript before 10.08.0 has a heap
# overflow in its JPEG 2000 decoder reachable from a crafted PDF (CVE-2026-39919, CVSS 9.8,
# fixed in 10.08.0 of 2026-09-08). Staging therefore refuses any build below -MinVersion.
param(
    [string]$Source,
    [string]$Dest,
    [ValidateSet('x64', 'x86')]
    [string]$Arch = 'x64',
    [version]$MinVersion = '10.8.0'
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

    # Newest by the VERSION of the engine itself. Folder names cannot be trusted for this:
    # a string sort puts "gs9.56" above "gs10.08" ('9' > '1'), and an install may live in a
    # folder without any version at all (e.g. %USERPROFILE%\gs64x).
    return ($roots | Sort-Object -Unique | Sort-Object { Get-GsVersion $_ } -Descending | Select-Object -First 1)
}

# Engine version from the DLL's file version (e.g. 10.07.1); $null when it cannot be read.
function Get-GsVersion([string]$root) {
    $path = Join-Path $root "bin\$dllName"
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $raw = (Get-Item -LiteralPath $path).VersionInfo.FileVersion
    $parsed = $null
    if ($raw -and [version]::TryParse(($raw -replace '[^0-9.].*$', ''), [ref]$parsed)) { return $parsed }
    return $null
}

if (-not $Source) { $Source = Find-GsRoot }
if (-not $Source -or -not (Test-Path -LiteralPath $Source)) {
    Write-Host "Ghostscript ($Arch) not found. Install the matching build: https://ghostscript.com/releases/gsdnld.html"
    Write-Host "or pass -Source <gs root, e.g. C:\Program Files\gs\gs10.08.0>."
    exit 1
}

$dll = Join-Path $Source "bin\$dllName"
$exe = Join-Path $Source "bin\$exeName"
if (-not (Test-Path -LiteralPath $dll) -or -not (Test-Path -LiteralPath $exe)) {
    Write-Host "Source has no bin\$dllName / bin\$exeName (need a $Arch Ghostscript): $Source"
    exit 1
}

$engineVersion = Get-GsVersion $Source
if (-not $engineVersion) {
    Write-Host "Cannot read the Ghostscript version of $dll - refusing to bundle an unidentified engine."
    exit 1
}
if ($engineVersion -lt $MinVersion) {
    Write-Host "Ghostscript $engineVersion ($Source) is older than the required $MinVersion."
    Write-Host "Install the current release: https://ghostscript.com/releases/gsdnld.html"
    exit 1
}

Write-Host "Ghostscript source ($Arch): $Source, version $engineVersion"
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
# GS license (AGPL) - bundled alongside for compliance. Redistribution without it is not
# allowed, so a source without a license file stops the staging instead of shipping silently.
$licensed = $false
foreach ($lic in @('LICENSE', 'doc\COPYING', 'COPYING')) {
    $lp = Join-Path $Source $lic
    if (Test-Path -LiteralPath $lp) { Copy-Item -LiteralPath $lp (Join-Path $Dest 'LICENSE') -Force; $licensed = $true; break }
}
if (-not $licensed) {
    Write-Host "Ghostscript license (LICENSE / doc\COPYING) not found in $Source - cannot redistribute it."
    exit 1
}
$gsVersion = ''
try { $gsVersion = ((& $exe --version) | Out-String).Trim() } catch { }
if (-not $gsVersion) { $gsVersion = (Split-Path $Source -Leaf) -replace '^gs', '' }
$sourceNotice = @"
Ghostscript version bundled by this installer: $gsVersion

Corresponding source code is published by Artifex at:
https://github.com/ArtifexSoftware/ghostpdl-downloads/releases
https://ghostscript.com/releases/

Choose the ghostpdl source archive whose version is exactly $gsVersion. iwo Helper Desktop
invokes Ghostscript as a separate process and does not link it into the application executable.
For redistribution questions or a source copy, open an issue at:
https://github.com/DedovMosol/iwoHelperDesktop
"@
[IO.File]::WriteAllText((Join-Path $Dest 'SOURCE.txt'), $sourceNotice,
    (New-Object Text.UTF8Encoding($false)))

$size = [math]::Round(((Get-ChildItem $Dest -Recurse -File | Measure-Object Length -Sum).Sum / 1MB), 1)
Write-Host "Ghostscript ($Arch) staged in $Dest ($size MB)."
exit 0
