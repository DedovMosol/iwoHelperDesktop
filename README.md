<div align="center">

<img src="docs/screenshots/banner.png" width="720" alt="iwo Helper Desktop">

<br>

[![CI](https://github.com/DedovMosol/iwoHelperDesktop/actions/workflows/ci.yml/badge.svg)](https://github.com/DedovMosol/iwoHelperDesktop/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/DedovMosol/iwoHelperDesktop?label=release&color=0F6CBD)](https://github.com/DedovMosol/iwoHelperDesktop/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/DedovMosol/iwoHelperDesktop/total?color=107C41)](https://github.com/DedovMosol/iwoHelperDesktop/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Privacy: offline‑only](https://img.shields.io/badge/Privacy-offline--only-107C41)](docs/PRIVACY.md)
![Windows](https://img.shields.io/badge/Windows-10%2F11%20x64-0078D6?logo=windows&logoColor=white)
[![.NET Framework 4.8](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?logo=dotnet&logoColor=white)](https://dotnet.microsoft.com/download/dotnet-framework/net48)

**Free, offline office tools in a single Windows app — merge Excel sheets, merge/split/compress PDFs at Acrobat‑level quality, and turn born‑digital PDFs back into editable Word. No subscription, no admin rights, no network.**

[![Installer](https://img.shields.io/badge/Download-Installer%20x64-0F6CBD?style=for-the-badge&logo=windows&logoColor=white)](https://github.com/DedovMosol/iwoHelperDesktop/releases/latest)
[![Portable](https://img.shields.io/badge/Download-Portable%20x64-107C41?style=for-the-badge&logo=windows&logoColor=white)](https://github.com/DedovMosol/iwoHelperDesktop/releases/latest)

</div>

## What is iwo Helper Desktop?

A small, self‑contained Windows application that bundles the office tasks people do every day with `.xlsx` and `.pdf` files — without a paid suite. It runs **offline**, needs **no administrator rights**, makes **no network calls**, and ships either as a single portable `.exe` or a per‑user installer.

## 🚀 Features

- 📊 **Excel Digest** — merges the first visible sheet of every workbook in a folder into one file (`.xlsx`/`.xlsm`/`.xlsb`/`.xls`), keeping all formatting (styles, formulas, charts, pivots); adds a table of contents, optional formula→value conversion, and a Word cover note (GOST R 7.0.97‑2016).
- 📄 **PDF Merge** — build one PDF from several: a grid of page thumbnails, drag to reorder, delete extras. Pages are copied **as‑is** — scans, stamps and signatures are not distorted.
- ✂️ **PDF Split** — extract selected pages into one file, or split by page ranges, every N pages, or top‑level bookmarks. The source is never modified.
- 📝 **PDF → Word** — extract the text layer of a **born‑digital** PDF (saved from Word, “Microsoft Print to PDF”, exported from a browser) into an editable `.docx`. Inherits the font family, size, bold/italic, colour, super/subscript, paragraph alignment (left/justify/centre) and first‑line indent, page size and margins, images (placed in reading order) and hyperlinks. Scanned documents are not supported yet — a clear message is shown, the file is untouched.
- 🗜️ **PDF Compression** — Acrobat‑level “Reduce File Size”: downsamples images while keeping text and vectors (not rasterization), via bundled **Ghostscript**. Default level leaves the file untouched.
- 🔄 **Update check & statistics** — compares with GitHub Releases (opens the page, downloads nothing); local operation counters with manual/auto clear.
- 🔒 **Safe by design** — no network, no admin, not packed/obfuscated; writes only to user‑selected folders and `%APPDATA%`.

## 📸 Screenshots

|  |  |
|:--:|:--:|
| <img src="https://raw.githubusercontent.com/DedovMosol/iwoHelperDesktop/assets/hub.png" width="400" alt="Start screen"><br>**Start screen** — pick a tool | <img src="https://raw.githubusercontent.com/DedovMosol/iwoHelperDesktop/assets/excel.png" width="400" alt="Excel Digest"><br>**Excel Digest** |
| <img src="https://raw.githubusercontent.com/DedovMosol/iwoHelperDesktop/assets/pdf-merge.png" width="400" alt="PDF Merge"><br>**PDF Merge** — thumbnails & compression | <img src="https://raw.githubusercontent.com/DedovMosol/iwoHelperDesktop/assets/pdf-split.png" width="400" alt="PDF Split"><br>**PDF Split** — modes & compression |

## ⬇️ Download

| OS | Download |
|----|----------|
| **Windows 10 / 11 (x64)** | [![Installer](https://img.shields.io/badge/Installer-x64-0F6CBD?logo=windows&logoColor=white)](https://github.com/DedovMosol/iwoHelperDesktop/releases/latest) &nbsp; [![Portable](https://img.shields.io/badge/Portable-x64-107C41?logo=windows&logoColor=white)](https://github.com/DedovMosol/iwoHelperDesktop/releases/latest) |

- **Installer** *(recommended)* — bundles Ghostscript, so PDF compression works out of the box. Installs **per‑user without admin** by default (choose “for all users” for a machine‑wide install).
- **Portable** — a single `iwoHelperDesktop.exe`; just run it. PDF compression works if Ghostscript is installed on the machine.

> Requirements: Windows 10/11 x64 (with .NET Framework 4.8, bundled since Windows 10 1903). Microsoft Excel is needed for **Excel Digest**, Microsoft Word for the **Excel Digest** cover note and **PDF → Word**. The PDF tools need neither.

## 🖥️ Usage

Launch the app and pick a tool from the start screen. Tools open as independent windows; a **⌂ Home** button returns to the chooser. Long tasks run in the background with progress on the list and the taskbar, and can be cancelled.

- **Excel Digest** — pick the source folder, set the output name/format, arrange/exclude files, click **Merge**. A report and an optional Word cover note are produced next to the digest.
- **PDF Merge / Split** — add PDFs (button or drag‑and‑drop), reorder/select pages on the thumbnail grid, choose a **Compression** level if desired, and save.
- **PDF → Word** — open a born‑digital PDF (button or drag‑and‑drop), then **Convert to Word…** and choose the `.docx` name.

<details>
<summary><b>Full Excel Digest guide, options and edge cases</b></summary>

1. Select the source folder (Browse… or drop it onto the window); the file count is shown immediately.
2. Set the output name and format (`.xlsx`/`.xlsm`/`.xlsb`/`.xls`); “Sheets” takes the first sheet of each file or all of them.
3. Change the output folder if needed (defaults to the source folder).
4. Arrange the **Files to merge** list — reorder by dragging or ▲/▼, exclude via checkboxes; “By name” restores natural order.
5. Click **Merge** — progress on the list and taskbar; the button flashes on completion when the window is inactive.
6. Existing output prompts to overwrite; a file open in Excel is detected up front.

**Files to merge** is one list with two roles: before the merge it shows order/inclusion; during/after it fills in the per‑file result (sheet name, status, skip reason / warning such as “file contains macros”). Rows copy to the clipboard (Ctrl+C).

After the merge: **Open file / folder / report** (a text history in `%APPDATA%\iwo Helper Desktop\reports`, three latest) and **Word note** — a `.docx` cover note (period, counters, a table of skipped files), formatted per GOST R 7.0.97‑2016. If files were skipped, **Retry skipped** appends fixed files without a full rebuild.

Options (format and “Table of contents” are remembered; “Replace formulas with values” starts off each run):
- **Table of contents** (on by default) — the first sheet becomes a TOC with hyperlinks and per‑file status; header row frozen.
- **Replace formulas with values** — the digest no longer depends on the sources.

Edge cases handled: broken/password‑protected files are detected by signature and skipped **before** Excel opens them (so they can’t wedge the shared instance); if Excel still wedges, it is restarted automatically; low disk space stops the run up front; hidden sheets are skipped; name clashes get a `_2` suffix; names > 31 chars are truncated and `: \ / ? * [ ]` become `_`; `~$` temp files are ignored; macros are never executed (VBA files are flagged).

</details>

<details>
<summary><b>PDF compression details & signature caveat</b></summary>

Both PDF tools have a **Compression** dropdown applied to the produced PDF:
- **Отлично** — no compression (default): byte‑for‑byte the merge/extract output; fidelity and signatures preserved.
- **Хорошо** — Ghostscript `/ebook` (~150 DPI).
- **Нормально** — Ghostscript `/screen` (~72 DPI).

The compressing levels downsample images while keeping text and vectors (the same idea as Adobe Acrobat / Foxit “Reduce File Size”), done by **Ghostscript** as a separate process. The result is validated (valid PDF, strictly smaller) before replacing the original; an already‑optimized file is left untouched. Output is PDF 1.4, so a compressed file can still be re‑merged/split by the app.

**Signatures:** any real compression changes the file’s bytes, so a **signed** PDF’s signature becomes invalid afterwards (true of Acrobat too). Compress unsigned documents, or before signing. Ghostscript is used under its own AGPL license (invoked as a separate process — the app stays MIT); the portable exe opens the official [download page](https://ghostscript.com/releases/gsdnld.html) if it is absent.

</details>

<details>
<summary><b>Command‑line mode (Excel Digest, for scripts)</b></summary>

```
iwoHelperDesktop.exe --cli <source_folder> <digest_path> [--toc] [--values] [--allsheets]
```
Format is derived from the path extension. `--toc` adds a table of contents, `--values` replaces formulas with values, `--allsheets` takes every visible sheet. The report is written to `<digest>.report.txt`. Exit codes: `0` all transferred, `2` some skipped, `1` error.

</details>

## 🛠️ Build from source

```
build.cmd
```
Needs the `dotnet` SDK (6+); builds `iwoHelperDesktop.csproj` (target .NET Framework 4.8) to a single `dist\iwoHelperDesktop.exe`. Managed dependencies are embedded as resources: `build/PdfSharp.dll` (MIT) for PDF create/merge/split, and `build/pdfpig/*` (**PdfPig**, Apache 2.0) for born‑digital text extraction in PDF → Word. PDF thumbnails use the system `Windows.Data.Pdf` (WinRT); PDF → Word writes the `.docx` through Word COM.

<details>
<summary><b>Signing, installer, release, CI and tests</b></summary>

- **Sign the exe:** `powershell -NoProfile -File tools\sign.ps1` — self‑signed cert in `Cert:\CurrentUser\My` (SHA256 + timestamp).
- **Installer:** `powershell -NoProfile -File tools\make_installer.ps1` — builds/signs the exe, stages bundled Ghostscript (`tools\stage_gs.ps1`), compiles `installer\iwoHelperDesktop.iss` (Inno Setup), signs the setup. Wizard images come from `tools\make_wizard_images.ps1`.
- **Release:** `powershell -NoProfile -File tools\make_release.ps1 -Publish` — builds/signs both artifacts, notes from the CHANGELOG, creates the GitHub release. See [docs/RELEASING.md](docs/RELEASING.md).
- **CI** (`.github/workflows/ci.yml`): on every push — build, unit tests, GUI smoke, a Ghostscript round‑trip (`--gscheck`), and an installer compile check. Releases are cut locally (self‑signed cert lives only on the maintainer’s machine).
- **Tests:** `tests\build_tests.cmd` (unit, no Office); `tests\run_all.cmd` (full pyramid, needs Excel/Word). Repository layout, corpus and maintainer notes are in [docs/](docs/).

**Maintainer note:** Office COM is used through late binding (`dynamic`). Never perform dynamic operations on a closed COM object (store the reference in an `object` before `Close`/`Quit`) — a dynamic bind on a dead object crashes with `COMException 0x80010114`. Text into Excel cells must go through `CellText.EscapeValues`.

</details>

## 🧩 Built with

Written in **C#** (.NET Framework 4.8, Windows Forms), powered by these open projects:

[![PdfSharp](https://img.shields.io/badge/PdfSharp-MIT-1f6feb?style=for-the-badge)](https://github.com/empira/PDFsharp)
[![PdfPig](https://img.shields.io/badge/PdfPig-Apache%202.0-1f6feb?style=for-the-badge)](https://github.com/UglyToad/PdfPig)
[![Ghostscript](https://img.shields.io/badge/Ghostscript-AGPL-d32f2f?style=for-the-badge)](https://ghostscript.com/)
[![Inno Setup](https://img.shields.io/badge/Inno%20Setup-installer-107C41?style=for-the-badge)](https://jrsoftware.org/isinfo.php)
[![Windows.Data.Pdf](https://img.shields.io/badge/Windows.Data.Pdf-WinRT-0078D6?style=for-the-badge&logo=windows&logoColor=white)](https://learn.microsoft.com/uwp/api/windows.data.pdf)

## 🔒 Privacy

**Your files never leave your computer.** No telemetry, no analytics, no accounts, no background network calls. All processing (Excel, PDF merge/split/compress) runs locally, the only network request is the **manual** update check, which reads the latest version tag from GitHub and sends no file contents or personal data. Full details: **[Privacy Policy](docs/PRIVACY.md)**.

## ⚖️ License

[MIT](LICENSE) © 2026 **Dodonov Andrey** ([DedovMosol](https://github.com/DedovMosol)) · full history in [docs/CHANGELOG.md](docs/CHANGELOG.md) · [Privacy Policy](docs/PRIVACY.md)
