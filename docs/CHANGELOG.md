# Changelog

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
versions follow [SemVer](https://semver.org/).

## [1.13.7] — 2026-07-21

### Added
- **Privacy Policy** ([docs/PRIVACY.md](PRIVACY.md)) making explicit that the app is
  offline‑only — no telemetry, and your files never leave your computer. Linked from the
  README and from the **About** dialog.

### Internal
- **Shared `PdfToolFormBase` for the two PDF tools (DRY).** The Merge and Split forms now
  inherit common state and behaviour — thumbnail grid, zoom slider + throttle timer,
  compression picker, status line, tooltips, the busy‑aware close guard and deterministic
  teardown — instead of duplicating them. No layout change. As a side‑fix, each window’s
  `ToolTip` (a component, not a child control) is now disposed on teardown.
- **Shared bottom‑strip builder.** The zoom slider (+ throttle timer), compression picker
  and status line — previously built identically in both forms — are now created by a single
  `BuildBottomStrip(...)` in the base. Control order, tab order and layout are unchanged.

## [1.13.6] — 2026-07-20

### Performance
- **Scrolling large PDFs is O(log n), not O(n).** The visible‑thumbnail range is found by
  binary search over the (monotonic) item bounds instead of scanning from the top on every
  scroll tick — no more hundreds of `LVM_GETITEMRECT` calls per tick on long documents.
- **Applying a rendered page is O(1) in the item count.** A `key → items` index replaces
  the linear scan of all list items in the render callback, so rendering a whole document
  is O(n) overall instead of O(n²).

### Internal
- 2 new unit tests (`LowerBound` binary search, `VisibleRange` visible‑window computation).

## [1.13.5] — 2026-07-20

### Fixed
- **PDF thumbnail memory no longer grows across documents.** The page‑bitmap/tile cache
  is now pruned to the pages currently shown: switching the document in *Split* or
  removing pages in *Merge* frees the bitmaps and image‑list tiles of pages that are no
  longer displayed, while reordering (same page set) keeps everything cached — no
  re‑render. Late render results for pages that were removed meanwhile are discarded
  instead of being cached.
- **Open PDF documents (and their file handles) are now bounded.** `PdfThumbnailRenderer`
  keeps at most a few WinRT `PdfDocument`s in a least‑recently‑used cache instead of one
  per file opened for the window’s lifetime, so paging through many files no longer
  accumulates native buffers or keeps every source file locked.

### Internal
- New reusable, unit‑tested `LruCache<T>` (bounded least‑recently‑used cache); the
  renderer’s document eviction reuses `ComSafe.Release`, removing the duplicated WinRT
  COM‑release code. 8 new unit tests (LRU eviction/touch/replace/case/clear/guard,
  grid key‑set and stale‑key computation).
- Assembly version attributes trimmed to `1.13.5` (was `1.13.5.0`): the exe’s File and
  Product version now read exactly `1.13.5`, matching the in‑app title/About and the
  installer/tag (`ToString(3)` unchanged, so the update check is unaffected).

## [1.13.4] — 2026-07-20

### Fixed
- **PDF thumbnail render thread now shuts down cleanly.** When a PDF tool window closes,
  the background render thread is joined (with a timeout) and its `ManualResetEventSlim`
  is disposed instead of being left to the finalizer — no leaked wait‑handle or lingering
  thread (correct `IDisposable` teardown, CA2213). The signal is released only once the
  thread has provably exited, so a slow in‑flight render can never fault on a disposed handle.

### Internal
- `.gitignore` now excludes the local `screenshots/` scratch folder (reference images),
  so it can’t be committed by an accidental `git add .`.

## [1.13.3] — 2026-07-20

### Fixed
- **Start‑screen window title showed only `1.13`** (two components) — now shows the full
  `1.13.3` (`Version.ToString(3)`).
- **Excel Digest header was hard‑coded to “first visible sheet”** even in the all‑sheets
  mode — reworded to the neutral “Листы Excel‑файлов из папки — в один итоговый файл”.
- **Compression dropdown truncated** “Нормально — минимальный размер” — the combo width
  is now computed from the widest item, so every level fits.

### Added
- **About dialog: donation details** (account number + bank) as selectable, copyable
  text, with a one‑click “копировать” for the account number and a copied‑confirmation.

## [1.13.2] — 2026-07-20

### Fixed
- **Installer no longer skipped the install-mode and folder pages on re-install.**
  Inno Setup hides them on upgrade by default (`UsePreviousPrivileges=yes`,
  `DisableDirPage=auto`). Set `UsePreviousPrivileges=no` (always ask all-users vs
  current-user) and `DisableDirPage=no` (always show the destination folder, pre-filled
  with the previous path via `UsePreviousAppDir=yes`). Verified against Inno Setup docs.

### Changed
- **Author is now credited** as **Dodonov Andrey (DedovMosol)** with the GitHub link:
  in the installer license page and publisher/URL fields, the About dialog, and the
  MIT `LICENSE`.

## [1.13.1] — 2026-07-20

### Added
- **PDF compression (Acrobat-level)** on both PDF tools — a “Compression” dropdown
  applied as a post-processing step to the produced file: **Отлично** (no compression,
  default — fidelity and signatures preserved), **Хорошо** (`/ebook`, ~150 DPI),
  **Нормально** (`/screen`, ~72 DPI). It **downsamples images while keeping text and
  vectors** (not rasterization), matching Adobe Acrobat / Foxit “Reduce File Size”.
  Powered by **Ghostscript** invoked as a separate process; the compressed file is
  written to `<pdf>.gstmp`, validated (exit code + `%PDF-` header + strictly smaller)
  and only then replaces the original — an already-optimized PDF is left untouched.
  Output uses PDF 1.4 (classic xref) so a compressed file can still be re-merged/split
  by the app. A shared `CompressionPicker` control is used by both tools (DRY); the
  work runs on the background thread **before** the file is opened, so the replace
  never hits a viewer lock. Compression is a no-op if Ghostscript is absent.
  Pure functions (`Preset`, `BuildArguments`, `ShouldReplace`, `PickFirstExisting`)
  and a live end-to-end compression test (real size reduction, pages preserved) are
  covered by unit tests; `--gscheck` is a CI smoke check.
- **Installer (Inno Setup)** alongside the portable exe: `iwoHelperDesktop-setup-*.exe`
  installs the app **and bundles Ghostscript**, so compression works out of the box.
  Default install is **per-user without administrator rights** (`%LOCALAPPDATA%`),
  with an option to install for all users (Program Files, requires admin). The
  installer's welcome page **explicitly states the per-user default**. Built and
  signed locally via `tools\make_installer.ps1` (`tools\stage_gs.ps1` prepares the
  Ghostscript subset). Ghostscript is bundled under its own AGPL license (invoked as
  a separate process — mere aggregation; the app stays MIT).
- **“About” button on the start screen** (opens the About dialog). It was moved out
  of every tool's Help menu (which now keeps “How to use” and “Statistics”).

### Notes
- Signatures: any real compression changes the bytes, so a signed PDF's signature
  becomes invalid (the same happens in Acrobat). Compress unsigned documents, or
  before signing. The default level does not touch the file.

### Changed
- The application now builds explicitly as **x64** (matches the bundled 64-bit
  Ghostscript engine).
- **Start screen bottom row reworked**: “Check for updates” moved to where the version
  number used to be (left); the right button is now “About” (was “Check for updates”).
  The buttons were enlarged and raised slightly. The version is still shown in the title
  bar and the About dialog.
- **“How to use”** (Help) in both PDF tools now documents the compression dropdown and
  the signature caveat.
- **Branded installer wizard image** (blue gradient + logo + “iwo”) replaces the default
  Inno graphic on the welcome/finish pages; generated by `tools\make_wizard_images.ps1`.
- Compression's in-place replace now uses a **rename-aside** strategy (original → `.gsbak`
  → compressed in place → backup removed; restored on any failure). This works on network
  drives where `File.Replace` can fail, and never leaves the file missing.

### Fixed
- **Compression dropdown rendered as a grey box** on the white form — the shared
  `CompressionPicker` set `BackColor = Transparent` without the
  `SupportsTransparentBackColor` style, so a `UserControl` fell back to the default
  grey. Verified by sampling the rendered background (now white).
- **About dialog text overflowed the window** — the description is now width-constrained
  (wraps) and following lines are positioned relative to it. Verified by rendering (no
  control extends past the client area).

### Verified
- Compression works with **Cyrillic paths and spaces** (e.g. `…\Рабочий стол\Мой
  документ №1.pdf`) — Ghostscript 10.x handles Unicode command-line paths.

## [1.12.1] — 2026-07-20

### Changed
- **Branded message dialogs** replace the native ones everywhere (info, error,
  confirm): a coloured icon by severity, the app's rounded buttons — a single button
  is centred, two are placed at opposite sides (e.g. the “clear statistics” confirm).
  All calls still go through the `Dialogs` facade (`MessageForm`); button placement
  is unit-tested (`ButtonX`).
- **“Check for updates” moved to the start screen** as a dedicated button (with the
  current version shown), instead of repeating in every tool's Help menu.

## [1.12.0] — 2026-07-20

### Added
- **PDF Split → “Combine into one file”** checkbox in the ranges mode: pages from
  all ranges are written into a single PDF (in the given order, duplicates kept),
  instead of one file per range. Reuses the tested extract core (`PageRanges.ToIndices`).
- **“Check for updates”** in the Help menu: reads the latest version from GitHub
  Releases (HTTPS) and, if newer, offers to open the download page in the browser.
  No self-download or self-replacement — the safest fit for a portable, self-signed,
  offline-friendly app (self-updating exes are widely flagged by antivirus). Tag
  parsing and version comparison are unit-tested; the network call runs off the UI thread.
- **“Statistics”** in the Help menu: local counters (no telemetry) of operations —
  Excel digests, PDF merges, page extractions, and splits by mode. Manual **Clear**
  and optional **auto-clear** (daily / every 7 / every 30 days). Counters use
  read-modify-write so concurrent windows can't lose increments; the auto-clear
  period logic (`ShouldAutoClear`) is unit-tested.

## [1.11.2] — 2026-07-20

### Changed
- **PDF Split — you can now choose the output name in every mode.** The
  split-into-many modes (ranges, every N pages, bookmarks) previously only let you
  pick the folder and reused the source file name; now a save dialog lets you set
  both the folder and the base name, to which the numbers/labels are appended
  (`base_1-3.pdf`, `base_часть_1.pdf`, `base_Глава.pdf`). Extract, PDF Merge and
  the Excel digest already allowed choosing the name.

## [1.11.1] — 2026-07-20

### Changed
- **Start-screen header is centred** (title and subtitle), since the hub has no
  window buttons (`HeaderBand.Centered`).
- **“PDF Split” now has its own icon** — scissors — so it is clearly distinct from
  “PDF Merge” at a glance.
- **Lazy thumbnail rendering** in the PDF page grid: only the visible pages (plus a
  small buffer) are rendered in the background instead of every page up front —
  markedly less CPU and memory for large documents (hundreds of pages), so the UI
  stays responsive. Visible-range windowing (`ClampWindow`) is unit-tested; the
  no-crash + lazy behaviour was verified in a real message loop (22/60 pages
  rendered for a 60-page file).
- Drag-and-drop path extraction for the PDF tools was de-duplicated into a shared
  `PdfDrop` helper (DRY).

## [1.11.0] — 2026-07-20

### Added
- **New tool: “PDF Split”** (third start-screen card), complementing “PDF Merge”.
  Open one PDF, see its pages as thumbnails, and either extract or split — following
  the modes of leading offline tools (PDFsam, Acrobat):
  - **Extract selected** — pick pages in the grid (Ctrl+A = all) → one new PDF;
  - **By ranges** — “1-3, 5, 8-”: each range → its own file;
  - **Every N pages** — equal chunks (N=1 → one file per page);
  - **By bookmarks** — one file per top-level bookmark, named from the titles.
  Pages are copied as-is (no re-conversion); the source is never modified; output
  names are never overwritten (a number is appended). The engine (`PdfSplitService`,
  `PageRanges`) is unit-tested and validated live on real PDFs, including bookmarks.
- The PDF page-thumbnail grid was extracted into a reusable `PdfPageGrid` control
  and is now shared by both PDF tools (DRY); the Merge tool was refactored onto it
  with no behaviour change.

## [1.10.7] — 2026-07-20

### Added
- **Keyboard shortcuts in both tools' lists**, handled in `ProcessCmdKey` so they
  are reliable (before dialog-key/menu handling), unit-tested via pure classifiers:
  - **PDF Merge**: `Delete` removes the selected pages, `Alt+←/→` reorder,
    `Ctrl+A` selects all, `Enter` no longer triggers a save from the list.
  - **Excel Digest**: added `Ctrl+A` (select all) and `Delete` (exclude/uncheck
    selected); existing `Alt+↑/↓` (reorder), `Ctrl+C` (copy) and `Enter`-suppression
    consolidated into `ProcessCmdKey` (copy/select-all now also work during a run).
  - Shortcut hints added to button tooltips and the “How to use” help.

### Fixed
- PDF reorder/remove could run from the keyboard **during a save** (the buttons
  were disabled but the methods weren't guarded); `MoveSelected`/`OnRemoveClick`
  now no-op while busy.

## [1.10.6] — 2026-07-20

### Added
- **“Back to contents” button on every sheet** of the digest (when the “Table of
  contents” option is on): a floating, designer-style rounded button — blue
  gradient fill, white bold text — that links to the contents sheet. It is a
  floating shape, so it never shifts or covers the transferred data, and it is
  idempotent (re-generated cleanly when retrying skipped files). Colour packing
  and the sheet-reference helper are unit-tested (`Theme.ToBgr`, `SheetRef`).

## [1.10.5] — 2026-07-19

### Fixed
- **Excel window title.** The Excel tool window was titled like the hub
  (“iwo Helper Desktop 1.10”); it is now “Свод Excel”, so it is distinct in the
  title bar and Task Manager (the PDF tool was already correct).
- **Keyboard handling in the file list now actually works.** `Enter` (the form's
  default button) is a dialog key intercepted before `KeyDown`, so the previous
  suppression never fired; and `Alt+↑/↓` were unreliable next to the menu. Both are
  now handled in `ProcessCmdKey` (which runs first): `Enter` in the list no longer
  starts the merge, `Alt+↑/↓` reorder reliably. Routing is unit-tested
  (`ClassifyListKey`).
- **Self-healing restart no longer double-counts results in the UI.** When a wedged
  Excel instance is restarted, the previous pass is replayed; the merge service now
  raises a `Restarting` event and the window clears the per-file rows so results
  aren't accumulated twice.

### Changed
- **Tool windows are now independent of the hub.** Closing the start screen no
  longer closes (or abruptly kills) the open Excel/PDF tools — they keep running,
  and the process exits only when the last window is closed. Window lifetime is
  owned by a new `ShellContext` (`ApplicationContext`).
- The tool button “◀ Назад в меню” became **“⌂ Главная”** and now re-opens the tool
  chooser (re-creating it if the hub was closed).
- The **“About” window is now blue** (bar and, on Windows 11, the title bar) to
  match the start screen.

## [1.10.4] — 2026-07-18

### Fixed
- **Header text no longer runs under the “Back to menu” button** on narrow
  windows: the title and subtitle are clipped to the leftmost child control with
  an ellipsis (`HeaderBand.TextRightBound`, unit-tested).
- **Bottom links (“Word note”) no longer overlap the “Retry skipped” button**: the
  Excel window's minimum width was widened so the two action areas can't collide.
- **PDF Merge window now has the app icon** (previously the default WinForms icon)
  and appears in the taskbar, consistent with the Excel window.

### Changed
- **Accessibility**: the start-screen tool cards report as buttons with a name and
  description to screen readers (`AccessibleRole`/`AccessibleName`).
- **Keyboard**: in the Excel “Files to merge” list, `Alt+↑`/`Alt+↓` reorder the
  selected file; `Enter` in the list no longer triggers the merge.
- The source-folder field is rescanned with a short debounce instead of on every
  keystroke.
- Tab order: the “Back to menu” button is visited last instead of early.
- App-icon loading was de-duplicated into a single `Ui.AppIcon()` helper.
- CLI usage line updated (correct exe name, `--allsheets`).

## [1.10.3] — 2026-07-18

### Changed
- **Per-tool header colours**: the window header band is now colour-coded by tool —
  green for “Excel Digest”, red for “PDF Merge” (matching its icon), blue for the
  start screen. On Windows 11 the system title bar is tinted to match. The subtitle
  is drawn in a neutral off-white so it stays legible on every background.

## [1.10.2] — 2026-07-17

### Added
- **Fault tolerance for the Excel merge**, in layers, following best practice:
  - **Signature pre-check** (`FileSignature`): each source file's container is
    detected by magic bytes before Excel touches it. A file that is neither a ZIP
    (OOXML) nor an OLE2/CFB document — e.g. text renamed to `.xlsx` — is skipped
    as corrupt; a `.xlsx`/`.xlsm`/`.xlsb` whose container is OLE2 is an encrypted
    (password-protected) workbook and is skipped as such. This matters because
    `Workbooks.Open` on a broken or encrypted file can wedge Excel so that every
    following file fails to open too.
  - **Self-healing restart**: if a file still wedges Excel (`Workbooks` stop
    responding), the Excel instance is torn down and restarted without that file,
    and the merge continues — no machine reboot, no loss of the other files
    (bounded to a few restarts, then a clear error).
  - **Pre-flight free-space check**: if the system, temp or output drive is nearly
    full, the merge stops up front with a clear message (“almost no free space on
    drive C: … Excel can't open files — free up space and retry”) instead of a
    dozen cryptic “unable to get the Open property” failures.
  - Unit tests: `FileSignature.Detect` (ZIP/OLE2/text/empty), `LowSpaceMessage`.

## [1.10.1] — 2026-07-17

### Added
- **Pre-merge file list in “Excel Digest”**: the “Files to merge” list now shows
  the source files before merging. You can set their order (drag rows or the
  “▲ Up” / “▼ Down” buttons), exclude any file by clearing its checkbox, restore
  the natural name order with “By name”, and select the whole set with
  “Check all” / “Uncheck all”. After the merge the per-file result fills the same
  rows. The reorder/exclusion logic is a pure, unit-tested model (`SourceFileList`,
  `ListReorder` — shared with the PDF page order, DRY); the merge service now
  takes an explicit file list (`Merge(files, …)`, `PrepareSourceList`).
- **Branded window header**: the top of every window (the start screen and both
  tools) carries an accent-green gradient header band (`HeaderBand`) with the
  title and subtitle; the “◀ Back to menu” button sits on it. On Windows 11 the
  system title bar is tinted to match via DWM (`WindowChrome`); on Windows 10 the
  title bar stays default and the header band provides the branding. Unit tests:
  `WindowChrome` COLORREF packing, `HeaderBand` construction.

### Changed
- **README and CHANGELOG are now in English**; the changelog moved to
  `docs/CHANGELOG.md`.

## [1.9.0] — 2026-07-17

### Added
- **Sheet selection in “Excel Digest”**: a “Sheets” drop-down — “First sheet
  only” (default, as before) or “All sheets”. In “all sheets” mode every visible
  sheet of each file is transferred with names “file · sheet”; the table of
  contents and the report get a row per sheet. CLI flag `--allsheets`. The result
  model is now one record per sheet; a retry of skipped files correctly expands a
  file into several sheets. Tests: `SheetBaseName`, `FileCount`, multi-sheet retry
  (unit) and `verify_allsheets.ps1` (integration).

## [1.8.3] — 2026-07-17

### Added
- **Several tools at once**: from the start screen you can open both “Excel
  Digest” and “PDF Merge” as separate windows. Opening the same tool again shows
  a notice and brings the already-open window to the front (`ToolRegistry`, unit
  test).
- **“◀ Back to menu” button** in every tool — brings the chooser window back to
  the front (shared `Ui.BackButton`).

### Changed
- Start screen: the “Choose a tool” title is centred, the “What do you need?”
  caption removed.
- In the “About” window only the links themselves are clickable (t.me/…,
  DedovMosol/…); the “Telegram:”, “GitHub:” labels are plain text.

### Fixed
- A chooser card fired twice on a single click (the base control raised Click and
  the handler raised it again): because of this the very first open showed “tool
  already open”. The duplicate call was removed; verified with window messages
  (exactly one Click).

## [1.8.2] — 2026-07-17

### Added
- A **“Help”** menu in the “PDF Merge” tool (as in “Excel Digest”): “How to use”
  (F1) and “About”. The menu was factored into a shared `HelpMenu` (DRY, a unit
  test for the structure).

### Changed
- The PDF icon on the chooser card is a red document with a vector “PDF”
  (from file-pdf.svg), matching the green Excel document.

### Fixed
- In the “About” window the GitHub link overlapped the “OK” button: the window is
  taller, the button dropped below the links, the link shortened.

## [1.8.1] — 2026-07-17

### Added
- **PDF thumbnail zoom**: a slider and Ctrl+mouse wheel. A page is rendered once;
  on zoom the tiles are rebuilt from cache (GDI, no repeated WinRT), and the
  rebuild is throttled — no stutter. Unit tests `ThumbZoom`.

### Changed
- **New chooser-card icons**: a document with a folded corner in the file-excel
  style (a green sheet with a table for Excel, a red one with “PDF”) instead of
  the previous abstract grid.

### Fixed
- A thumbnail tile no longer exceeds the `ImageList` limit (256×256): at maximum
  zoom WinForms threw an exception. The zoom bounds were adjusted, a protective
  clamp and a regression test added.

## [1.8.0] — 2026-07-17

### Changed
- **Rebrand: iwo Helper Desktop** — new name and icon (logo). The name was
  updated in window titles, the “About” window, reports, build metadata and the
  data folder (`%APPDATA%\iwo Helper Desktop`). The internal tools are “Excel
  Digest” and “PDF Merge”.
- **Build moved to the dotnet SDK** (SDK project `iwoHelperDesktop.csproj`,
  net48): a single exe `dist/iwoHelperDesktop.exe`, PdfSharp still embedded as a
  resource. This opened access to WinRT (Windows.Data.Pdf) for thumbnails via the
  NuGet package `Microsoft.Windows.SDK.Contracts` — compile time only, not
  shipped; nothing is installed on the target machine.

### Added
- **Tool-chooser start screen**: “Excel Digest” and “PDF Merge” cards with
  descriptions; after a tool is closed the chooser is shown again.
- **PDF page thumbnails**: the “PDF Merge” tool shows a grid of previews of the
  real pages (the system Windows.Data.Pdf engine), reordered by dragging
  thumbnails and with buttons. Rendering runs in the background (a separate
  thread); if the engine is unavailable (e.g. on Windows Server) it falls back to
  placeholders as designed. Tests: `verify_thumb.ps1` (rendering and aspect ratio)
  and `--thumbcheck` (clean process exit after WinRT rendering).

### Fixed
- Forced process exit (`FastExit`/`ExitProcess`) after working with WinRT: the
  normal finalization of the Windows.Data.Pdf COM wrappers crashed the process on
  unload; the critical cleanup (settings, COM Quit for Excel/Word) runs
  deterministically before exit.

## [1.7.0] — 2026-07-17

### Added
- **PDF Merge** (the “Tools” menu): pick PDF files, a single list of pages,
  reorder with ▲▼ buttons and by dragging, delete, save to a single document.
  Pages are copied without re-conversion — scans, stamps and signatures are not
  distorted (PDFsharp, MIT, embedded into the exe as a resource — still one file
  shipped). Broken/protected PDFs are skipped with a reason.
- Tests: a unit test for the page-order model (reorder/move/delete), the
  integration `verify_pdf.ps1` (order and a duplicated page verified by A4/A5/
  landscape dimensions), `verify_embedded.ps1` (resolving the embedded PdfSharp
  from the exe resource in a clean folder).

## [1.6.0] — 2026-07-16

### Added
- **Word cover note**: a “Word note” link after the merge — a `.docx` next to the
  digest (period, counters, a table of skipped files with reasons), formatted per
  GOST R 7.0.97-2016; generated through the COM of an installed Word, the pure
  text model is covered by unit tests, the document by an integration test
  (`tests/verify_note.ps1`).
- Sorting the log by clicking a column header (natural comparison, a second click
  reverses direction, the system arrow in the header).
- A “file contains macros (not executed)” note in the log and table of contents
  for sources with VBA — when saving the digest to `.xlsm`/`.xls` the sheet code
  is transferred together with the sheet, in `.xlsx` it is dropped.
- A “Processing log” heading above the results list.
- An integration test for the retry of skipped files (`tests/verify_retry.ps1`).

## [1.5.0] — 2026-07-16

### Added
- A **“Retry skipped”** button: fixed files are appended to an existing digest
  without a full rebuild; the table of contents is regenerated from the overall
  result, the order and the successful sheets are preserved.
- **Copying log rows** — Ctrl+C or the context menu: a “file → sheet → reason”
  row in the report format, handy to forward to the owner of a broken file.
- **CHANGELOG.md** (this file), linked from the README.

### Changed
- The “Replace formulas with values” option is no longer **remembered** between
  runs: the mode changes the digest content and is enabled deliberately each time.

## [1.4.0] — 2026-07-16

### Added
- **Output format selection**: `.xlsx`, `.xlsm`, `.xlsb`, `.xls` (a drop-down; in
  the CLI the format is derived from the path extension).
- An integration run to `.xlsb` in the common test set.

### Changed
- Branded checkboxes: a white check on a green background, hover, a focus ring
  (`AccentCheckBox`).
- The log columns share the window width proportionally.
- Per the Windows guidelines, the ellipses were removed from the “How to use” and
  “About” items; the punctuation in the help was fixed.

## [1.3.0] — 2026-07-16

### Added
- A **“Help”** menu: “How to use” (F1), “Reports folder”, “About” (version,
  author, license, clickable Telegram and GitHub links).
- The application version in the window title.

### Changed
- “Merge” and “Cancel” were moved to opposite sides of the window.
- The progress indicator is hidden when idle (an empty grey bar was confusing).

## [1.2.0] — 2026-07-16

### Added
- **Taskbar-button progress** (ITaskbarList3) and a window flash on completion
  when the user is working in another application.
- An **early lock check** for the output file: a busy file is detected before
  Excel starts, not after all sources have been processed.
- **Report history** in `%APPDATA%\ExcelMerger\reports` (at most three), an “Open
  report” link after the merge.
- CI (GitHub Actions): build, unit tests, GUI smoke; the exe is published to
  Releases on a `v*` tag.
- `tools/sign.ps1` — signing the exe with a self-signed certificate (SHA256).
- `tests/run_all` — the whole test pyramid in one command.

## [1.1.1] — 2026-07-16

### Fixed
- **Escaping strings when writing to cells**: a file name or a formula's string
  result that started with “=” turned into a formula (injection); a leading
  apostrophe of a string was lost. Verified experimentally, covered by unit and
  integration tests.

## [1.1.0] — 2026-07-16

### Added
- A **“Table of contents” sheet**: a digest table of contents with hyperlinks to
  the sheets and the status of every file, including skipped ones (an option, on
  by default).
- **Natural file order** as in Explorer: “Report 2” before “Report 10”
  (StrCmpLogicalW).
- A **“Replace formulas with values”** option — a digest without external
  references; merged cells are handled by a per-cell fallback.
- An OLE message filter: automatic retry of COM calls rejected by a busy Excel.
- A manual recalculation mode during the merge (faster with formulas).
- Unit tests without external frameworks (`tests/build_tests.cmd`).

## [1.0.0] — 2026-07-16

First release.

- Merges the first visible sheet of every Excel file in a folder into a single
  `.xlsx` through the COM of an installed Excel — without losing formatting,
  formulas, merged cells, charts and pivot tables.
- Source formats: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`; broken and password-protected
  files are skipped with a reason; hidden sheets are not transferred; sheet names
  come from file names with deduplication and a 31-character limit.
- WinForms GUI: live validation, processing progress, a colour-coded log,
  folder drag-and-drop, path memory; an icon and branded styling.
- A `--cli` mode for scripts and automated tests; integration tests on a corpus
  of 13 files; a single exe ~65 KB with no dependencies (.NET Framework 4.8,
  the compiler bundled with Windows).
