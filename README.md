# iwo Helper Desktop

![CI](https://github.com/DedovMosol/iwoHelperDesktop/actions/workflows/ci.yml/badge.svg)

Version history — see [docs/CHANGELOG.md](docs/CHANGELOG.md).

A set of office tools in a single application:

- **Excel Digest** — merges the first visible sheet of every Excel file in a folder
  into one combined workbook (`.xlsx`/`.xlsm`/`.xlsb`/`.xls`) without losing any
  formatting (styles, formulas, charts, pivot tables), with a table of contents,
  formula-to-value conversion and a Word cover note.
- **PDF Merge** — builds a single PDF from several files: pick the pages you need
  and their order, no re-conversion.
- **PDF Split** — the reverse: open one PDF and either extract the selected pages
  into a new file or split the document into several — by page ranges (optionally
  **combined into one file**), every N pages, or by top-level bookmarks. Pages are
  copied as-is; the source is untouched.

Both PDF tools have a **“Compression”** dropdown (Acrobat-level): the default keeps
the file untouched; the other levels downsample images while preserving text and
vectors (see [PDF compression](#pdf-compression)).

On launch the program lets you choose a tool; the start screen also has a
**“Check for updates”** button (compares with GitHub Releases and opens the download
page — nothing is downloaded or replaced automatically). Each tool's **Help** menu has
**“Statistics”** (local operation counters, with manual and optional auto-clear).

## Requirements

- Windows 10/11 x64 (needs .NET Framework 4.8 — bundled with Windows 10 1903+);
- Microsoft Excel installed (2007–2024); Microsoft Word for the cover note;
- **Ghostscript** is required only for PDF compression — the installer bundles it;
  the portable exe uses a system-installed Ghostscript if present (otherwise the
  compression option shows a download link and stays a no-op);
- administrator rights and network are **not required** (a per-user install and the
  portable exe both work without admin).

## Repository layout

```
src/       application sources           tools/     exe signing, installer scripts
tests/     unit and integration tests    build/     build inputs (manifest, icons, PdfSharp.dll)
dist/      build output (not in git)     docs/      changelog and documentation
installer/ Inno Setup script + license (bundled Ghostscript staged into installer/gs, not in git)
```

Built with the `dotnet` SDK. The only shipped dependency is `build/PdfSharp.dll`
(MIT); it is embedded into the exe as a resource, so a single file goes out.
PDF thumbnails use the system `Windows.Data.Pdf` (WinRT); the compile-time
projections come from the NuGet package `Microsoft.Windows.SDK.Contracts`
(build only, not shipped).

## Deployment

Two options, both from [Releases](https://github.com/DedovMosol/iwoHelperDesktop/releases):

1. **Portable exe** — copy the single `iwoHelperDesktop.exe` to the target machine.
   That's all. PDF compression works if Ghostscript is installed on the machine
   (otherwise it is a no-op with a download hint).
2. **Installer** (`iwoHelperDesktop-setup-*.exe`) — installs the app **and bundles
   Ghostscript**, so compression works out of the box. Installs **per-user without
   admin** by default (`%LOCALAPPDATA%`), with an option to install for all users
   (Program Files, requires admin).

The portable exe is not packed or obfuscated (an ordinary .NET assembly, ~0.7 MB —
including the resource-embedded PdfSharp), makes no network calls, and writes
only to user-selected folders and to `%APPDATA%\iwo Helper Desktop` (settings
and report history) — nothing for antivirus software to react to.

## PDF compression

Both PDF tools offer a **“Compression”** dropdown, applied as a post-processing step
to the produced PDF:

- **Отлично** — no compression (default): the file is byte-for-byte the merge/extract
  output; fidelity and any digital signatures are preserved.
- **Хорошо** — Ghostscript `/ebook` (~150 DPI): good balance of size and quality.
- **Нормально** — Ghostscript `/screen` (~72 DPI): smallest size.

The compressing levels **downsample images while keeping text and vectors** (the file
stays searchable) — the same approach as Adobe Acrobat / Foxit “Reduce File Size”,
**not** rasterization. Compression is done by **Ghostscript** as a separate process;
the result is validated (correct PDF, strictly smaller) before it replaces the
original — an already-optimized file is left untouched. The output is PDF 1.4, so a
compressed file can still be re-merged or re-split by the app.

Caveat on signatures: any real compression changes the file's bytes, so a **signed**
PDF's signature becomes invalid afterwards (this is true of Acrobat too). Compress
**unsigned** documents, or compress **before** signing. The default level never
touches the file.

Ghostscript is bundled by the installer; the portable exe looks for a system-installed
Ghostscript and, if absent, the compression option opens the official
[download page](https://ghostscript.com/releases/gsdnld.html). Ghostscript is used
under its own AGPL license (invoked as a separate process — mere aggregation; this
app stays MIT).

## Usage (GUI)

On launch the program lets you choose a tool: **“Excel Digest”** (merging Excel
sheets) or **“PDF Merge”**. After a tool is closed the chooser is shown again —
you can switch to the other one without restarting. Both tools can be open at the
same time as independent windows, and closing the start screen does not close
them — the app exits only when the last window is closed. Each tool has a
**“⌂ Home”** button that opens the tool chooser (re-creating it if it was closed).

### Excel Digest

1. Select the folder with the source files (the “Browse…” button or drop the
   folder onto the window). The program immediately shows how many Excel files
   were found.
2. Set the output file name and format (`.xlsx` / `.xlsm` / `.xlsb` / `.xls`).
   Under Options, “Sheets” lets you take only the first sheet of each file or all
   of them.
3. Change the output folder if needed (defaults to the source folder).
4. **Arrange the files** in the “Files to merge” list: reorder them by dragging
   rows or with the “▲ Up” / “▼ Down” buttons, and **exclude** any file by
   clearing its checkbox. “By name” restores the natural order; “Check all” /
   “Uncheck all” select the whole set quickly.
5. Click “Merge”. Progress is shown in the list and on the taskbar button; when
   the window is inactive the button flashes on completion.
6. If the output file already exists, the program asks whether to overwrite it;
   a busy file (open in Excel) is detected up front, before any processing.

**Files to merge** is a single list that serves two purposes. Before the merge it
shows the source files with their order and inclusion state. During and after the
merge each row is filled in with the per-file result: the sheet name in the
digest, a status, and the reason a file was skipped or a warning (for example,
“file contains macros”). Rows can be copied to the clipboard with Ctrl+C or the
context menu — handy for forwarding the reason to a file's owner.

After the merge the following become available: “Open file”, “Open folder”,
“Open report” (a text history kept in `%APPDATA%\iwo Helper Desktop\reports`, at
most the three latest; also reachable via Help → “Reports folder”) and
**“Word note”** — a `.docx` cover note next to the digest: period, counters, a
table of skipped files with reasons, formatted per GOST R 7.0.97-2016 (margins
30/15/20/20 mm, Times New Roman 14, first-line indent, 1.5 line spacing).

If some files were skipped, a **“Retry skipped”** button appears — fixed files
are appended to the existing digest without a full rebuild (the table of contents
is regenerated; the order and the already-merged sheets are left untouched).

Supported source formats: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`. By default the files
follow the natural order of their names, as in Explorer: “Report 2” before
“Report 10” — and you can override it in the list.

Options (the format and “Table of contents” are remembered; “Replace formulas
with values” starts off on every run — the mode is enabled deliberately):

- **“Table of contents” sheet** (on by default) — the first sheet of the digest
  becomes a table of contents: hyperlinks to every sheet and the status of every
  file, including skipped ones with reasons. The header row is frozen.
- **Replace formulas with values** — the digest no longer depends on the source
  files: computed values are stored instead of formulas (formatting and merged
  cells are left intact).

Help is in the menu (F1): “How to use”, “Reports folder”, “About”.

## PDF Merge

The **“PDF Merge”** tool from the start screen is a separate window for stitching
PDFs. You pick documents (with the button or by drag-and-drop), get a **grid of
page thumbnails**, reorder them by dragging or with the buttons, delete the extra
ones, and save to a single PDF. Pages are copied **as-is**, without
re-conversion — scans, stamps and signatures are not distorted. Broken and
password-protected PDFs are skipped with a clear reason.

Stitching is done with PdfSharp (MIT, embedded into the exe as a resource).
Thumbnails are drawn by the system Windows.Data.Pdf engine; if it is unavailable
(for example, on Windows Server) placeholders are shown and the tool keeps
working. A single file is shipped; nothing is installed.

## Edge-case behaviour

- a broken / password-protected / unreadable file is skipped, with the reason in
  the list, the table of contents and the report; corrupt and encrypted files are
  detected by their signature and skipped **before** Excel opens them, so they
  cannot wedge the shared Excel instance;
- if a file still wedges Excel, the instance is restarted automatically without
  that file and the merge continues (no machine reboot);
- if the system, temp or output drive is nearly full, the merge stops up front
  with a clear message instead of cryptic per-file COM errors;
- hidden sheets are not transferred (the first **visible** sheet is taken);
- name clashes (for example `Report.xls` and `Report.xlsx`) — the second sheet
  gets a `_2` suffix;
- a name longer than 31 characters is truncated; the characters `: \ / ? * [ ]`
  are replaced with `_`;
- Excel temporary files (`~$…`) are ignored;
- source macros are **not executed**; files with VBA are flagged with a note.
  Sheet code does not reach an `.xlsx` digest (Excel drops it); in `.xlsm`/`.xls`
  it is transferred together with the sheet;
- if no sheet could be transferred or the merge was cancelled, the digest file is
  not created (and on a retry of skipped files it is left unchanged).

## Command-line mode (for tests and scripts)

```
iwoHelperDesktop.exe --cli <source_folder> <digest_path> [--toc] [--values] [--allsheets]
```

The digest format is derived from the path extension (`.xlsx`/`.xlsm`/`.xlsb`/`.xls`).
`--toc` adds a “Table of contents” sheet, `--values` replaces formulas with
values, `--allsheets` takes every visible sheet (otherwise only the first). In
the CLI the options default to off and are enabled explicitly. The report is
written to `<digest>.report.txt`. Exit codes: `0` — all files transferred,
`2` — some were skipped, `1` — error (the digest file was not created).

## Building from source

```
build.cmd
```

Requires the `dotnet` SDK (6+); builds the SDK project `iwoHelperDesktop.csproj`
(target framework .NET Framework 4.8). The result is `dist\iwoHelperDesktop.exe`,
a single file. Nothing is installed on the target machine (only .NET Framework
4.8 is needed, which ships with Windows 10).

The application icon is `build/app.ico` (multi-size, from `build/logo.ico`).
The interface palette is `src/Theme.cs`.

Signing the exe (an optional step before deployment):

```
powershell -NoProfile -File tools\sign.ps1
```

Creates (once) a self-signed certificate in `Cert:\CurrentUser\My` and signs the
exe (SHA256, with a timestamp when the network is available). This gives file
integrity and a persistent publisher; the signature becomes “trusted” for Windows
only after the certificate is added to the trusted roots on the target machine.

Building the installer (locally; needs [Inno Setup 6](https://jrsoftware.org/isdl.php)
and Ghostscript installed):

```
powershell -NoProfile -File tools\make_installer.ps1
```

It builds and signs the exe, stages the bundled Ghostscript (`tools\stage_gs.ps1`
copies the needed subset into `installer\gs`), compiles `installer\iwoHelperDesktop.iss`
with Inno Setup, and signs the resulting `dist\iwoHelperDesktop-setup-<version>.exe`.
The branded wizard images (`installer\wizard.bmp`, `wizard_small.bmp`) are generated by
`tools\make_wizard_images.ps1` from `build\app.ico`. CI validates the `.iss` on every
push (unsigned artifact); the **signed** installer is produced locally, because the
self-signed certificate lives only on the maintainer's machine.

Cutting a GitHub release (portable exe + installer + CHANGELOG notes) is scripted in
`tools\make_release.ps1` — see [docs/RELEASING.md](docs/RELEASING.md).

## CI

GitHub Actions (`.github/workflows/ci.yml`): on every push — build, unit tests,
GUI smoke test, a Ghostscript compression round-trip (`--gscheck`), and an installer
compile check (Inno Setup + staged Ghostscript, unsigned artifact); artifacts
`iwoHelperDesktop.exe` and `iwoHelperDesktop-setup-*.exe`. On a `v*` tag — the exe is
published to Releases (the signed exe and signed installer are uploaded from the
maintainer's machine). Integration tests require Office and run locally only.

## Tests

The whole pyramid in one run (needs Excel and Word installed):

```
tests\run_all.cmd
```

Sequence: build → unit tests → GUI smoke → corpus generation → base run → run
with table of contents and formula replacement → run to `.xlsb` → retry of
skipped files → embedded PdfSharp → PDF merge → PDF thumbnails → Word note →
Excel zombie-process check.

Separately — unit tests (no Office, a custom mini-runner; a live compression test
uses Ghostscript when it is installed, otherwise it verifies the graceful no-op):

```
tests\build_tests.cmd    # build and run; exit code 0 = all passed
```

Integration checks: `tests\verify.ps1` (base behaviour), `verify_toc.ps1` (table
of contents and formula replacement), `verify_format.ps1` (digest to `.xlsb`),
`verify_retry.ps1` (retry of skipped files), `verify_pdf.ps1` (PDF merge),
`verify_thumb.ps1` (WinRT thumbnail rendering), `verify_note.ps1` (Word note,
including GOST margins and font).

The corpus covers: formatting/formulas/merged cells (including a formula inside a
merged cell and formulas with string results), an empty file, `.xls`/`.xlsm`/`.xlsb`,
a duplicate base name, a hidden first sheet, a name longer than 31 characters,
brackets in the file name, natural order (“Report 2” before “Report 10”), a
password-protected file, a broken file, a temporary `~$` file.

## License

[MIT](LICENSE)

## Important for maintainers

Office COM is used through late binding (`dynamic`) — there is no dependency on
the Office version. The rule: **never perform dynamic operations on a closed COM
object** (store the `dynamic` reference in an `object` variable before
`Close`/`Quit`) — a dynamic bind on a dead object crashes the process with
`COMException 0x80010114` before the method is even entered, bypassing any
try/catch (details are in the comment next to `ComSafe.Release`). Text written
into Excel cells must go through `CellText.EscapeValues` — otherwise a string
like “=x” turns into a formula.
