# Releasing

Releases are cut **locally**, because both artifacts are signed with a self-signed
certificate that lives only on the maintainer's machine (`Cert:\CurrentUser\My`).
CI (`.github/workflows/ci.yml`) builds, tests and validates the installer on every
push, but does **not** create GitHub Releases.

Each release publishes two assets:

- `iwoHelperDesktop.exe` — portable single file (run as-is — PDF compression works if
  Ghostscript is installed on the machine).
- `iwoHelperDesktop-setup-<version>.exe` — installer that **bundles Ghostscript**
  (compression out of the box), installs **per-user without admin** by default.

## Prerequisites (maintainer machine)

- .NET SDK (build), [Inno Setup 6/7](https://jrsoftware.org/isdl.php) (installer),
  Ghostscript (bundled into the installer — `tools\stage_gs.ps1` copies the subset).
- [GitHub CLI](https://cli.github.com/) authenticated (`gh auth login`) with push access.

## Steps

1. Bump the version in `src/AssemblyInfo.cs` (`AssemblyVersion` + `AssemblyFileVersion`).
2. Add a `## [X.Y.Z] — <date>` section to `docs/CHANGELOG.md` (its text becomes the
   release notes verbatim).
3. Commit the changes (explicit paths only).
4. Dry run — builds and signs both artifacts, writes `dist\release-notes-<ver>.md`,
   prints what would be published:

   ```
   powershell -NoProfile -File tools\make_release.ps1
   ```

5. Publish — creates the tag `vX.Y.Z`, pushes it, and creates the GitHub release with
   both signed assets and the CHANGELOG-derived notes:

   ```
   powershell -NoProfile -File tools\make_release.ps1 -Publish
   ```

`make_release.ps1` chains `make_installer.ps1` (build → sign exe → stage Ghostscript →
ISCC → sign installer). Re-running `-Publish` for an existing tag updates the assets
(`gh release upload --clobber`) and the notes.

## Trust (optional)

The self-signed signature gives file integrity and a stable publisher, but Windows
SmartScreen still warns (unknown publisher). For a managed fleet, IT can deploy the
public certificate to **Trusted Publishers / Trusted Root** via Group Policy — then the
signature is trusted org-wide with no warnings, at no cost. For public distribution use
an OV/EV code-signing certificate or Azure Trusted Signing. See `tools/sign.ps1`.
