# Local GitHub release: builds and SIGNS the portable exe and the installer
# (make_installer.ps1), extracts the notes for the current version from docs\CHANGELOG.md,
# and creates a release with both artifacts. Signing uses a self-signed certificate
# (locally only), so releases are cut here, not on CI.
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

# 1. Build + sign both artifacts
if (-not $SkipBuild) {
    & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'make_installer.ps1')
    if ($LASTEXITCODE -ne 0) { throw 'make_installer.ps1 failed' }
}

# 2. Version and paths (single source - the exe version)
$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'
if (-not (Test-Path $exe)) { throw "missing $exe" }
$ver = (Get-Item $exe).VersionInfo.FileVersion
$ver3 = ($ver -split '\.')[0..2] -join '.'
$tag = "v$ver3"
$installer = Join-Path $root ("dist\iwoHelperDesktop-setup-$ver3.exe")
if (-not (Test-Path $installer)) { throw "missing installer $installer" }

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
    "**Files:** ``iwoHelperDesktop.exe`` - portable (single file); " +
    "``iwoHelperDesktop-setup-$ver3.exe`` - installer with bundled Ghostscript " +
    "(per-user install, no administrator rights)."
[IO.File]::WriteAllText($notes, $body, (New-Object Text.UTF8Encoding($false)))
Write-Host "Release notes: $notes"

# 4. Publish
if (-not $Publish) {
    Write-Host ""
    Write-Host "DRY RUN. Artifacts and notes are ready:"
    Write-Host "  tag       = $tag"
    Write-Host "  exe       = $exe"
    Write-Host "  installer = $installer"
    Write-Host "  notes     = $notes"
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
    & gh release create $tag $exe $installer --title $tag --notes-file $notes
} else {
    & gh release upload $tag $exe $installer --clobber
    & gh release edit $tag --notes-file $notes
}
$rc = $LASTEXITCODE
$ErrorActionPreference = $eap
if ($rc -ne 0) { throw 'gh release failed' }
Write-Host "OK: release $tag published (portable + installer)."
exit 0
