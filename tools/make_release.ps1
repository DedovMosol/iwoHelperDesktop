# Local GitHub release: builds and SIGNS all four artifacts - the portable exe and the
# installer for BOTH architectures (make_installer.ps1 -Arch x64/x86) - extracts the
# notes for the current version from docs\CHANGELOG.md, and creates a release with them.
# Signing uses a self-signed certificate (locally only), so releases are cut here, not
# on CI. Staging the x86 installer needs a 32-bit Ghostscript installed alongside the
# 64-bit one (see docs\RELEASING.md).
#
# Default is a dry run (prepares artifacts and notes, publishes nothing).
# Real publish: tools\make_release.ps1 -Publish (needs git push access for the tag and an
# authenticated gh; the tag vX.Y.Z is derived from the exe version).
param(
    [switch]$Publish,
    [switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$ps = (Get-Process -Id $PID).Path

# 1. Build + sign both architectures (portable exe + installer each)
if (-not $SkipBuild) {
    foreach ($arch in @('x64', 'x86')) {
        & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'make_installer.ps1') -Arch $arch
        if ($LASTEXITCODE -ne 0) { throw "make_installer.ps1 -Arch $arch failed" }
    }
}

# 2. Version and paths (single source - the exe version; both arches must match)
$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'
$exe86src = Join-Path $root 'dist\x86\iwoHelperDesktop.exe'
foreach ($p in @($exe, $exe86src)) { if (-not (Test-Path $p)) { throw "missing $p" } }
$ver = (Get-Item $exe).VersionInfo.FileVersion
$ver86 = (Get-Item $exe86src).VersionInfo.FileVersion
if ($ver -ne $ver86) { throw "version mismatch: x64=$ver x86=$ver86 (rebuild both)" }
$ver3 = ($ver -split '\.')[0..2] -join '.'
$tag = "v$ver3"
$installer = Join-Path $root ("dist\iwoHelperDesktop-setup-$ver3.exe")
$installer86 = Join-Path $root ("dist\iwoHelperDesktop-setup-$ver3-x86.exe")
foreach ($p in @($installer, $installer86)) { if (-not (Test-Path $p)) { throw "missing installer $p" } }

# The portable x86 asset gets an explicit name on the release page (the signature made
# by make_installer survives the copy).
$exe86 = Join-Path $root 'dist\iwoHelperDesktop-x86.exe'
Copy-Item $exe86src $exe86 -Force

# 3. Release notes from the CHANGELOG section for this version
$notes = Join-Path $root ("dist\release-notes-$ver3.md")
$changelog = [IO.File]::ReadAllLines((Join-Path $root 'docs\CHANGELOG.md'), [Text.Encoding]::UTF8)
$section = New-Object System.Collections.Generic.List[string]
$inSection = $false
foreach ($line in $changelog) {
    if ($line -match '^##\s*\[') {
        if ($inSection) { break }                         # next version section - stop
        if ($line -match ("\[" + [regex]::Escape($ver3) + "\]")) { $inSection = $true }
    }
    elseif ($inSection) { $section.Add($line) }
}
if ($section.Count -eq 0) { throw "no [$ver3] section in docs\CHANGELOG.md" }
$body = "# iwo Helper Desktop $tag`r`n" + (($section -join "`r`n").Trim()) + "`r`n`r`n" +
    "**Files:** ``iwoHelperDesktop.exe`` / ``iwoHelperDesktop-x86.exe`` - portable " +
    "(single file, 64-bit / 32-bit); ``iwoHelperDesktop-setup-$ver3.exe`` / " +
    "``iwoHelperDesktop-setup-$ver3-x86.exe`` - installers with bundled Ghostscript " +
    "(per-user install, no administrator rights). Take x64 unless your Windows is 32-bit."
[IO.File]::WriteAllText($notes, $body, (New-Object Text.UTF8Encoding($false)))
Write-Host "Release notes: $notes"

$assets = @($exe, $exe86, $installer, $installer86)

# 4. Publish
if (-not $Publish) {
    Write-Host ""
    Write-Host "DRY RUN. Artifacts and notes are ready:"
    Write-Host "  tag   = $tag"
    foreach ($a in $assets) { Write-Host "  asset = $a" }
    Write-Host "  notes = $notes"
    Write-Host "To publish on GitHub: tools\make_release.ps1 -Publish"
    exit 0
}

# git/gh write to stderr on normal "not found" cases - under ErrorActionPreference=Stop
# that would crash the script. Switch to Continue temporarily and decide by exit code.
$eap = $ErrorActionPreference
$ErrorActionPreference = 'Continue'

# Tag (create and push if it does not exist yet)
& git -C $root rev-parse --verify --quiet "refs/tags/$tag" 1>$null 2>$null
if ($LASTEXITCODE -ne 0) {
    & git -C $root tag $tag
    & git -C $root push origin $tag
    if ($LASTEXITCODE -ne 0) { $ErrorActionPreference = $eap; throw "failed to push tag $tag" }
}

# Release: create if missing; otherwise update the artifacts and notes.
& gh release view $tag 1>$null 2>$null
$releaseExists = ($LASTEXITCODE -eq 0)
if (-not $releaseExists) {
    & gh release create $tag @assets --title $tag --notes-file $notes
} else {
    & gh release upload $tag @assets --clobber
    & gh release edit $tag --notes-file $notes
}
$rc = $LASTEXITCODE
$ErrorActionPreference = $eap
if ($rc -ne 0) { throw 'gh release failed' }
Write-Host "OK: release $tag published (portable + installer, x64 + x86)."
exit 0
