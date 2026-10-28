# Local GitHub release: builds and SIGNS the exe and the installer for BOTH architectures
# (make_installer.ps1 -Arch x64/x86), packs each signed exe with its .config into a
# portable zip, extracts the
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

# Release assets must be reproducible from the commit carrying the tag. The project compiles
# src\**\*.cs, so even one untracked .cs silently changes the exe without changing HEAD.
$status = @(& git -C $root status --porcelain --untracked-files=all)
if ($LASTEXITCODE -ne 0) { throw 'git status failed' }
if ($status.Count -ne 0) {
    throw "working tree is not clean; commit or remove every release input first:`n$($status -join "`n")"
}

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

# Every asset's file name carries the version, so a downloaded file is self-identifying.
# Installers already come out versioned from ISCC. The portable build is a versioned zip:
# the exe needs iwoHelperDesktop.exe.config beside it (WinForms DPI mode, build\app.config),
# and a zip keeps the pair together. The exe inside keeps the Authenticode signature made
# by make_installer; a zip itself cannot carry one.
# The README buttons link to releases/download/<tag>/<versioned name> — a direct
# download (not the release page), pinned to this version and updated per release.
function Assert-Signed([string]$path) {
    $signature = Get-AuthenticodeSignature $path
    if (-not $signature.SignerCertificate) { throw "not signed: $path" }
    if (-not $signature.TimeStamperCertificate) { throw "no trusted timestamp: $path" }
}
foreach ($p in @($exe, $exe86src, $installer, $installer86)) { Assert-Signed $p }
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
function New-PortableZip([string]$exePath, [string]$zipPath) {
    $files = @($exePath, "$exePath.config")
    foreach ($f in $files) { if (-not (Test-Path $f)) { throw "missing $f" } }
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    $archive = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($f in $files) {
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $f,
                [IO.Path]::GetFileName($f), [IO.Compression.CompressionLevel]::Optimal)
        }
    }
    finally { $archive.Dispose() }
    # Read the zip back: exactly these two entries, byte-identical to the signed build output.
    $archive = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $names = @($archive.Entries | ForEach-Object { $_.FullName }) -join '|'
        $expected = @($files | ForEach-Object { [IO.Path]::GetFileName($_) }) -join '|'
        if ($names -ne $expected) { throw "unexpected entries in ${zipPath}: $names" }
        foreach ($f in $files) {
            $stream = $archive.GetEntry([IO.Path]::GetFileName($f)).Open()
            $sha = [Security.Cryptography.SHA256]::Create()
            try { $packed = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
            finally { $stream.Dispose(); $sha.Dispose() }
            if ($packed -ne (Get-FileHash $f -Algorithm SHA256).Hash) { throw "zip entry differs from ${f}: $zipPath" }
        }
    }
    finally { $archive.Dispose() }
}
$portable64 = Join-Path $root ("dist\iwoHelperDesktop-$ver3.zip")
$portable86 = Join-Path $root ("dist\iwoHelperDesktop-$ver3-x86.zip")
New-PortableZip $exe $portable64
New-PortableZip $exe86src $portable86

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
    "**Files:** ``iwoHelperDesktop-$ver3.zip`` / ``iwoHelperDesktop-$ver3-x86.zip`` - portable " +
    "(64-bit / 32-bit; unpack both files into one folder and run the exe - the .config beside it " +
    "keeps the windows sharp on high-DPI displays); ``iwoHelperDesktop-setup-$ver3.exe`` / " +
    "``iwoHelperDesktop-setup-$ver3-x86.exe`` - installers with bundled Ghostscript " +
    "(per-user install, no administrator rights). Take x64 unless your Windows is 32-bit."
[IO.File]::WriteAllText($notes, $body, (New-Object Text.UTF8Encoding($false)))
Write-Host "Release notes: $notes"

$assets = @($portable64, $portable86, $installer, $installer86)
foreach ($asset in $assets) {
    Write-Host ("SHA256 {0}  {1}" -f (Get-FileHash $asset -Algorithm SHA256).Hash,
        [IO.Path]::GetFileName($asset))
}

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
try {
    # Tag: an existing release may be refreshed only from the exact tagged commit.
    $head = (& git -C $root rev-parse HEAD).Trim()
    & git -C $root rev-parse --verify --quiet "refs/tags/$tag" 1>$null 2>$null
    $tagExists = ($LASTEXITCODE -eq 0)
    if ($tagExists) {
        $tagCommit = (& git -C $root rev-list -n 1 $tag).Trim()
        if ($tagCommit -ne $head) {
            throw "tag $tag points to $tagCommit, but assets were built from $head"
        }
    } else {
        & git -C $root tag $tag
        if ($LASTEXITCODE -ne 0) { throw "failed to create tag $tag" }
        & git -C $root push origin $tag
        if ($LASTEXITCODE -ne 0) { throw "failed to push tag $tag" }
    }

    # Release: create if missing; otherwise update the artifacts and notes.
    & gh release view $tag 1>$null 2>$null
    $releaseExists = ($LASTEXITCODE -eq 0)
    # Every gh call is checked on its own: $LASTEXITCODE holds only the LAST command, so a
    # failed upload followed by a successful edit used to be reported as a published release.
    if (-not $releaseExists) {
        & gh release create $tag @assets --title $tag --notes-file $notes
        if ($LASTEXITCODE -ne 0) { throw "gh release create $tag failed (exit $LASTEXITCODE): release not published" }
    } else {
        & gh release upload $tag @assets --clobber
        if ($LASTEXITCODE -ne 0) { throw "gh release upload $tag failed (exit $LASTEXITCODE): assets were NOT replaced" }
        & gh release edit $tag --notes-file $notes
        if ($LASTEXITCODE -ne 0) { throw "gh release edit $tag failed (exit $LASTEXITCODE): assets uploaded, notes NOT updated" }
    }
}
finally {
    $ErrorActionPreference = $eap
}
Write-Host "OK: release $tag published (portable + installer, x64 + x86)."
exit 0
