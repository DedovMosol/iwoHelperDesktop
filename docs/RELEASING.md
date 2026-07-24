# Releasing

Releases are cut **locally**, because both artifacts are signed with a self-signed
certificate that lives only on the maintainer's machine (`Cert:\CurrentUser\My`).
CI (`.github/workflows/ci.yml`) builds, tests and validates the installer on every
push, but does **not** create GitHub Releases.

Each release publishes four assets — the same tool set for both architectures:

- `iwoHelperDesktop.exe` / `iwoHelperDesktop-x86.exe` — portable single files, 64-bit and
  32-bit (run as-is — PDF compression works if Ghostscript is installed on the machine).
- `iwoHelperDesktop-setup-<version>.exe` / `iwoHelperDesktop-setup-<version>-x86.exe` —
  installers that **bundle Ghostscript** of the matching bitness (compression out of the
  box), install **per-user without admin** by default. Minimum OS is Windows 8.1; the
  installer checks for .NET Framework 4.8 and points to the download when it is missing.

## Prerequisites (maintainer machine)

- .NET SDK (build), [Inno Setup 6/7](https://jrsoftware.org/isdl.php) (installer),
  Ghostscript in **both bitnesses** — the w64 and w32 builds installed side by side
  (`tools\stage_gs.ps1 -Arch x64|x86` copies the matching subset into `installer\gs\`
  and `installer\gs32\`).
- [GitHub CLI](https://cli.github.com/) authenticated (`gh auth login`) with push access.

## Steps

1. Bump the version in `src/AssemblyInfo.cs` (`AssemblyVersion` + `AssemblyFileVersion`).
2. Add a `## [X.Y.Z] — <date>` section to `docs/CHANGELOG.md` (its text becomes the
   release notes verbatim).
3. Commit the changes (explicit paths only).
4. Dry run — builds and signs all four artifacts, writes `dist\release-notes-<ver>.md`,
   prints what would be published:

   ```
   powershell -NoProfile -File tools\make_release.ps1
   ```

5. Publish — creates the tag `vX.Y.Z`, pushes it, and creates the GitHub release with
   the four signed assets and the CHANGELOG-derived notes:

   ```
   powershell -NoProfile -File tools\make_release.ps1 -Publish
   ```

`make_release.ps1` chains `make_installer.ps1` per architecture (x64, then x86; each:
build → sign exe → stage Ghostscript → ISCC → sign installer) and refuses to publish
when the x64 and x86 exe versions differ. Re-running `-Publish` for an existing tag
updates the assets (`gh release upload --clobber`) and the notes.

## Trust (optional)

The self-signed signature gives file integrity and a stable publisher, but Windows
SmartScreen still warns (unknown publisher). For a managed fleet, IT can deploy the
public certificate to **Trusted Publishers / Trusted Root** via Group Policy — then the
signature is trusted org-wide with no warnings, at no cost. For public distribution use
an OV/EV code-signing certificate or Azure Trusted Signing. See `tools/sign.ps1`.
