# Changelog

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
versions follow [SemVer](https://semver.org/).

## [1.17.9] — 2026-07-27

### Added
- **A start screen with two sections.** The hub now asks for a section first — **PDF** or
  **Other tools** — and shows the tools inside it. The window keeps one size on every screen so
  it never jumps, `Esc` and a “Back” button in the header return to the top, and “Home” inside a
  tool comes back to **its own section** rather than the top screen, so hopping between PDF tools
  still costs one click. Dropping PDFs onto the “PDF” card carries them inside: the header says
  how many files are waiting and the next tool you pick opens with them.
- **“More operations” — a window of its own** (PDF section). Six actions over a single document
  that until now hid in a “More actions” menu inside Split, where nobody found them: compress,
  convert to grayscale, repair a damaged file, save pages as images, extract the text to a
  `.txt`, edit the document properties. Split keeps a one-click bridge that hands the open
  document over, so nothing has to be opened twice.
- **Merge hands its result on.** “☰ Menu → More operations” continues the work on the file you
  have just saved — compress it, drop the colour, fix its properties — without hunting for it on
  disk. Split has the same bridge for the document it has open, and both open the window even
  when there is nothing to hand over yet: an entry that answers a click with an explanation of
  why it does nothing reads as broken. PDF → Word deliberately has no bridge: it cannot name a
  single PDF, its sources are many and its result is a `.docx`.
- **Compressing a single file** is a real operation at last. Until now compression only existed
  as an afterthought of merging and splitting. It writes a copy — sources are never modified —
  and if the file is already optimized it says so instead of silently returning the same bytes.
- **The user guide ships inside the program.** “About” has a new line, “User guide: open”, which
  unpacks the document next to the settings and opens it. It works in the portable build too and
  needs no internet, which is the whole point of an offline app.
- **The full-size preview can be minimized.** The minimize box and a taskbar button of its own
  come as a pair: the preview is modal, so a minimized window without a taskbar button could not
  have been brought back at all.

### Changed
- **“◀ Earlier” / “Later ▶” became “◀ Move left” / “Move right ▶”.** A bare “Earlier” on a button
  reads as “in the past”, not as an action. The grid is horizontal, the arrows are on the buttons
  and the shortcuts are `Alt+←/→`, so the axis is what the eye follows. The tooltip stays ordinal
  — “move the page one position earlier” — because at a row wrap “to the left” stops being true.
- **“Excel Digest” is now “Merge Excel”**, in a pair with “Merge PDF”. The *result* is still
  called a digest: a digest is what merging workbooks produces.
- **The language globe sits where the eye looks for it** — centred against the title and the
  subtitle instead of hanging above them. The position is computed from the real font heights,
  so it stays centred at any display scale.
- **“Document properties” was laid out again.** The buttons overlapped the last field by a pixel
  and the fields had three pixels between groups: the layout was computed from literals, while
  WinForms sizes a single-line text box by its font and ignores the height you set. Everything is
  measured now, the window says which file it edits, and it states that an empty field clears the
  property.

- **A card's contents stand in the middle of it.** The icon, the title and the description were
  pinned to the top of the card, so a short description left an empty strip at the bottom and the
  card looked unfinished. The block is measured and centred now, with both texts laid out by the
  same rules they are drawn with.

- **Buttons say when they cannot be pressed.** An unavailable ordinary button used to differ
  from a working one by the colour of its caption alone, and in panels where half the buttons
  wait for an open document that had to be found out by clicking. It is filled now, so the state
  reads at a glance. The corner radius follows the height by one formula instead of a step at
  36 px, the focus ring is solid rather than translucent so keyboard users can see where they
  are, and a long caption is now cut with an ellipsis before it reaches the rounded corner. The
  zoom bar of the full-size preview was the only place with square buttons and colours of its
  own — it uses the same button as the rest of the app, in a variant for a dark background.

### Fixed
- **The page in the full-size preview no longer hides in the grey.** Its position was computed in
  the window's coordinates while the viewport was scrolled — WinForms keeps a child control in
  *content* coordinates, so the origin drifted by the scroll on every zoom step and accumulated.
  Enlarging or maximizing the window did not re-centre the page at all, and a stale scrollable
  area left phantom scrollbars behind. All three are fixed, and the centring maths moved into a
  pure function under unit tests.
- **The language chosen in the installer reaches the program on a reinstall.** Setup used to
  write the language straight into `settings.txt` and only when that file did not exist yet, so
  choosing English over an existing Russian installation changed nothing. It now hands the choice
  over in a separate one-line file that the program applies once at startup and removes. Setup
  never touches `settings.txt`, which is what protected the non-ASCII paths inside it.
- **Ghostscript reporting success on a file it could not read.** On a password-protected document
  the engine exits with zero and still leaves a valid-looking two-kilobyte stub, which passed both
  replacement policies: “Repair a damaged PDF” and “Convert to grayscale” answered with a green
  “Done” and an empty document. The engine's own error marker in its error stream is what counts
  now, not the exit code.
- **The preview's zoom bar at display scales above 100 %.** The window is deliberately not
  auto-scaled (its size is measured in physical pixels), so the bar stayed in fixed pixels while
  the font grew with the screen: at 150 % every caption on it was clipped. The bar is measured
  from the current scale now, and the window's minimum size comes from the bar's real width — the
  “Print” button used to sit outside a window squeezed to its minimum.
- **“Check for updates” no longer swallows a browser that will not start.** Answering “Yes” to
  the new-version question opened the release page and, if the system had no browser to open it
  with, did nothing at all — the failure was caught and dropped. The address is shown now, so the
  page can be reached by hand.
- Gaps between the buttons of the right-hand panels differed between Merge and PDF → Word. They
  follow one rhythm now.

## [1.17.8] — 2026-07-27

### Added
- **Printing, on every PDF screen.** A “Print” button on Merge, Split and PDF → Word prints the
  pages you selected, or all of them if you selected none, and the full-size preview has one too
  alongside its right-click entry. Pages are rendered at 200 dpi and fitted into the sheet whole
  and centred — stretching a page to the edges would crop the margins, which is where signatures
  and page numbers live — and a rotation you set in the grid reaches the paper as well.
- **Zoom and pan in the full-size preview.** The page you open by double-clicking a tile now
  has a magnifier: − / + buttons, `Ctrl`+wheel, a “fit to window” button and `Ctrl+0` for actual
  size. Zooming with the wheel keeps the point under the cursor in place — you magnify what you
  are looking at, not the middle of the page — and once the page is larger than the window you
  drag it with the hand cursor. `Esc` and the close button close it.
- **Assemble a double-sided scan** (Merge → menu). For a scanner that takes one side at a time:
  scan the stack face up, flip it, scan again, add both files and use this item. The app asks
  whether the second file came out in reverse and lays the pages out as 1, 2, 3, 4… A single
  `Ctrl+Z` puts the order back.
- **Save pages as images** (Split → “More with this document”): PNG or JPEG at 96, 150, 300 or
  600 dpi, either the selected pages or all of them. The resolution is applied to each page's
  real paper size, so a landscape insert is not stretched. JPEG is written at quality 90 rather
  than the default 75, which shows up as rings around letters on scans.
- **Extract the text layer to a .txt** (Split → “More with this document”). Tables are kept: the
  analysis moves table words out of the paragraph stream, so a naive dump would lose them
  entirely — cells come out separated by tabs and paste into a spreadsheet as a table. Pages are
  separated by a form feed, as `pdftotext` does, and the file is UTF-8 with a byte-order mark so
  Notepad shows Cyrillic correctly.
- **Name the split parts yourself** (Split → “How to name the parts”). Right-clicking the field
  inserts the available keywords: `[BASENAME]`, `[FILENUMBER]`, `[CURRENTPAGE]`, `[BOOKMARK]`,
  `[TIMESTAMP]`. Hashes pad with zeros (`[FILENUMBER###]` → `001`) and a number offsets the
  count (`[FILENUMBER10]` → `11`). The field is optional on purpose: left empty, the parts get
  exactly the names they got before, so a new capability does not rename anyone's files.
- **Print on both sides without a document starting on a back** (Merge → “Add a blank page”): a
  blank page is added after a document with an odd page count so the next one starts on the
  front of a sheet. The last document is deliberately not padded — nothing is printed after
  it. The blanks are added when the file is written, not kept in the page grid: putting them
  in the model would mean drawing them as tiles, carrying them through the clipboard and undo
  and forbidding them in PDF → Word, which is a lot of risk for a print-time detail.
- **Edit document properties** — title, author, subject, keywords (Split → “More with this
  document”). Clearing a field removes the property, which is how an author's name is taken
  out of a file before sending it. The result is written to a new file.
- **Select odd or even pages** from the page grid's context menu — the two sides of a
  double-sided stack, so rotating or deleting one side no longer means ticking fifty tiles.
- **Convert to grayscale** and **repair a damaged PDF** (Split → “More with this document”).
  Repair rewrites the file through the PDF engine, which is what fixes a broken cross-reference
  table — the usual “this file is damaged”. It picks the file with its own dialog, because a
  damaged document cannot be opened into the grid in the first place, and that is exactly when
  it is needed. Both write a **new** file: the app never modifies a source.

### Changed
- **The extra tools are reachable from a button** called “More actions…”, sitting with the other
  inputs in Split, and it opens downwards instead of over the top of the window. Five
  capabilities behind a hamburger submenu were, in practice, invisible; that submenu is gone,
  so there is one list and one way to reach it.
- **The in-app instructions match the app again** — they now describe printing, the extra
  actions, the part-name template, interleaving and double-sided padding.
- **Changing the language no longer throws every window at you.** Rebuilding the windows in the
  new language used to bring each of them to the front and restore the minimized ones. Only the
  window you were working in comes forward now; the rest update where they stand.
- Everything that writes a result beside a source now refuses to write **over** it. The app
  does not modify sources, and writing a file into itself would have damaged it as well — the
  source is open for reading at that moment.
- **The compression levels name their resolution** — “Good — smaller size (150 dpi)” instead of
  “Good — smaller size”. The number comes from the same place the engine arguments do, so the
  label cannot promise one thing while the engine does another.
- The shared Ghostscript pipeline (run, validate, replace only on success, restore the original
  on any failure) is now one piece of code used by compression and by both new conversions. They
  differ only in arguments and in when a replacement is allowed: compression replaces the file
  only if it got **smaller**, while grayscale and repair apply whenever the result is sound.
- One file-name sanitiser instead of two that had drifted apart. The surviving one also trims
  trailing dots and caps the length, which protects split-by-bookmark names from the path-length
  limit.
- **The description in “About” is justified**, the way a paragraph of body text is set, and it is
  still selectable so the text can be copied.

### Fixed
- **A minimized start screen was lost when the language changed.** Windows are rebuilt in the new
  language on the spot they occupied, but a minimized window reports a position far outside every
  screen (-32000, -32000) and that position was copied onto the new start screen. It came back as
  an ordinary window parked past the edge of the desktop: listed in the taskbar, absent from the
  screen, and “Home” appeared to do nothing at all. This happened from any of the four tools, not
  just one. Placement is now read and applied exactly the way the app already stored it between
  runs — restored size and position together with the window state — so a minimized window stays
  minimized and comes back where it was. That rule lives in one place instead of two that had
  drifted apart, and a test drives a real rebuild with the start screen minimized.

## [1.17.7] — 2026-07-26

### Fixed
- **The MIT licence text was damaged.** One line of the permission grant read “to permit
  persons a side the Software is furnished to do so” instead of “to permit persons **to
  whom** the Software is furnished to do so”, in both the repository licence and the copy
  shown by the installer. GitHub could not recognise the file as MIT because of it. The
  canonical wording is restored in both files, verbatim.
- **Closing a PDF window could hang for two seconds and leak memory.** The thumbnail
  renderer is asked to stop through a signal, and the render thread could clear that signal
  in the same instant it arrived — after which it waited forever. The window then sat out
  its two-second timeout on close, and the renderer, which holds a full in-memory copy of
  every PDF it has open, was never released for the rest of the session.
- **A window busy reading a PDF lost its work when the language changed.** Only a running
  operation counted as “busy”, so switching language while a large or network PDF was still
  being read closed the window mid-read. The freshly opened one came up empty, with nothing
  said about the file that had been loading.
- **The Excel Digest window broke when dragged to its smallest size.** “Uncheck all” ended
  up below the bottom edge of the window and “Check all” disappeared underneath “Retry
  skipped” — and that broken size was then remembered between runs. The minimum height is
  now derived from the actual layout instead of a constant that had drifted away from it.
- **Every tool window opened with the focus on “Home”**, so pressing Enter right after
  opening one returned to the start screen without doing anything. The header really is last
  in the tab order now, and the focus lands on the first useful control.
- **The window header was clipped from 125% display scaling upwards.** Title and subtitle
  were drawn into fixed-height boxes while their fonts grow with the scale, slicing off the
  bottom of the letters. Both rows are measured from the fonts now.
- **Merging with compression looked frozen.** The bar reached 100%, Cancel disappeared, and
  the window then sat still for the whole compression run. Compression now names itself in
  the status line and shows a moving bar, since its progress cannot be measured.
- **Long messages ran off the edge of the Excel Digest window**, cut mid-word with no
  ellipsis. Both lines now shorten with “…” and show the full text on hover, as the PDF
  tools already did.
- **A list item numbered with a non-ASCII digit silently lost its number** when written to
  Word.
- **Enter did nothing in PDF Split** while it started the action in the other two PDF tools,
  and “Combine into one file” stayed clickable during a split.
- **Checking for updates gave no sign of life:** the button stayed lit through a request
  that can take ten seconds, and repeated clicks stacked up identical error dialogs.
- **“About” and the page-number prompts showed a generic window icon**, unlike every other
  dialog.
- The language globe on the start screen could not be reached from the keyboard, and would
  have been invisible on Windows 8.1, whose icon font does not have it — it now falls back to
  the flag of the current language and answers Enter and Space.

### Changed
- **Window position memory is wired in one place.** Windows used to restore and save their
  bounds through a pair of overrides each, which is easy to half-remove without noticing.
  They now attach it with a single call, and a test opens each of those windows for real to
  prove the memory still works — a silent loss of it would otherwise go unnoticed.
- **The shared parts really are shared now.** The dialog frame, the rounded-rectangle
  outline, “select all rows”, the PDF file picker and the “operation finished” epilogue each
  lived in three or four copies that had begun to drift apart — one was missing the window
  icon, another the guard against a degenerate corner radius.
- The taskbar progress object is released deterministically, like every other COM object in
  the app, and process handles are closed at the six places that open a file or a folder.

### Testing
- Windows are opened for real and squeezed to their minimum size, with every control checked
  for staying inside the window and no two buttons overlapping; the tab order of each tool
  window is checked the same way. Both layout defects above were caught by these checks
  before they were fixed.
- Twelve new checks cover logic that had none: the median behind every PDF layout threshold,
  the X-Y cut tree the header tables are built from, the missing-key fallback of the string
  catalog, the compression labels, cancellation, the screen chosen when restoring a window,
  the UI-thread delivery guard, and the rule that keeps a chosen language from being
  overwritten by a stale window.
- Two checks that could never fail were repaired: one compared two nulls, the other kept its
  only assertion inside a loop that an empty result skipped entirely.
- The runner refuses to pass if the number of checks drops, so a deleted one cannot slip by.
- Continuous integration installs Ghostscript **before** running the checks, so the
  compression round-trip is actually exercised instead of silently skipped; it builds the
  checks for the architecture under test (the 32-bit branches had never been run); and the
  thumbnail check now reports “nothing was rendered” as a failure instead of success.

## [1.17.6] — 2026-07-26

### Added
- **Rotate a page while previewing it.** Right‑click the full‑size preview (double‑click a
  tile to open it) for “rotate right / left 90°”, the same Ctrl+Shift+“+” / “−” as in the
  grid. The thumbnail follows immediately and the rotation joins the undo history, so
  Ctrl+Z takes it back like any other change.
- **The preview remembers its size.** The full‑size preview reopens at the size and
  position you left it, maximised included, restored safely onto a visible screen.

### Fixed
- **The installer stayed English after Russian was chosen.** Setup remembered the language
  of the previous installation and let it override the system language, and the restart
  that applies a newly chosen language never happened because Setup cannot launch its own
  file while it is running. Setup now follows the system language, hands the restart to the
  command processor, and says so instead of staying silent if that restart fails. The
  choice on the flag screen becomes the app language even in that case, and the picker's
  buttons no longer share a value with “dialog closed”, which used to read as a choice.
- **A language switch left a folder behind in `%TEMP%`.** The instance being replaced ends
  abruptly, which skips Setup's own cleanup, so its temporary folder now gets removed
  right after it exits.
- **Right‑clicking the preview closed it.** The click that opened the context menu also
  reached the close handler, because in Windows Forms a plain click event fires for the
  right button too.
- **Launching the app twice started a second process.** A second launch now wakes the
  running one and brings its window back instead of adding another entry to Task Manager.
  Tool windows stay independent of the hub exactly as before.

### Changed
- **Result messages name the compression resolution.** “Pages saved: 12 · compressed,
  images to 150 dpi” instead of a bare “compressed”, so it is clear what changed. Extracted
  pages are counted too, and the punctuation of these lines is assembled in one place.
- **Clearer drag‑and‑drop wording.** Every hint uses the imperative “drag it onto the
  program window” and names the drop target, following the Microsoft terminology.
- **The About window is readable and copyable.** The description can be selected and
  copied, it states that only born‑digital PDFs convert to Word (scans are not supported
  yet), and the copyright with the licence moved to the bottom‑left corner.

## [1.17.5] — 2026-07-26

### Added
- **Type an exact zoom percentage.** A “%” box next to the zoom slider lets you enter a
  precise scale (or nudge it with the spin arrows), on top of the slider and Ctrl+wheel.
  Ctrl+0, or a double‑click on the “%”, resets the zoom to 100%.
- **The installer opens with a flag language chooser.** Setup starts with a Russian/English
  picker showing each flag (instead of a plain drop‑down), runs the wizard in the chosen
  language, and makes that language the app’s default, so an English user is not shown a
  Russian interface. The portable build follows the system language on first run.
- **Windows remember their size and position.** Each tool window (and the Excel Digest
  window) reopens where and at the size you left it, restored safely onto a visible screen
  (never off‑screen or on a disconnected monitor).
- **Clearer Split hints.** The “ranges” and “every N pages” fields show tooltips that
  explain the format, and the shortcuts cheat sheet and per‑tool help now cover the new
  zoom entry and the reset.

### Changed
- **Consistent zoom control across PDF tools.** The zoom slider and “%” box now line up at
  the same size and position in PDF → Word as they do in Merge and Split.

## [1.17.4] — 2026-07-25

### Fixed
- **Double-clicking a rotate button rotates twice.** A quick double-click on a tile’s
  ↺/↻ hover button now turns the page a full 180° instead of turning it once and then
  opening the full-size preview on top.
- **The last zoom is always remembered.** Closing a tool right after moving the zoom
  slider now saves the slider’s position (previously a value set in the last fraction of
  a second could be lost).
- **A thumbnail that fails to render once retries.** A page whose thumbnail could not be
  produced on the first attempt (a file briefly locked by antivirus or on a flaky network
  share) no longer stays blank until the document is reopened.
- **PDF → Word withdraws Cancel at the point of no return**, the way Merge and Split
  already did, so the button no longer stays active — and then stuck on “Canceling…” —
  while Word is writing the final `.docx`.

### Changed
- **Opening or adding a PDF no longer freezes the window.** The file is parsed in the
  background, so a large document or one on a network drive keeps the window responsive
  instead of showing “Not Responding”.
- **Smoother grid.** Ctrl+wheel zoom over the page grid is throttled like the slider, and
  the grid’s drawing no longer allocates temporary brushes and fonts on every frame — less
  stutter while scrolling and zooming large documents.

## [1.17.3] — 2026-07-25

### Added
- **Zoom readout and Ctrl+wheel on the slider.** The zoom slider now shows the current
  scale as a percentage next to it, and holding Ctrl while turning the wheel over the
  slider steps the zoom (matching Ctrl+wheel over the grid).
- **The last zoom and compression level are remembered** between runs, so a tool opens at
  the scale and quality you left it.
- **Double-click a page for a full-size preview.** Double-clicking a tile opens the page
  at full size on a dark backdrop (rendered in the background, so the window stays
  responsive). Esc or a click closes it. Double-clicking the page NUMBER under a tile
  instead opens the “move after page…” prompt — a fast way to reorder without the menu.
- **Cancel long operations.** Merge, Split and PDF → Word show a Cancel button while
  working on documents of five pages or more. Canceling stops cleanly and leaves no
  half-written file (multi-file splits remove any parts already produced).

### Fixed
- The “Go to page” / “Move after page…” dialog no longer clips its action button (the
  longer **Move** label was cut off), and its two buttons are spread to opposite sides
  (Cancel on the left, the action on the right) to match the app’s other dialogs.

## [1.17.2] — 2026-07-25

### Added
- **Empty page grids explain themselves.** An empty grid now shows a centered hint
  (“Drop PDFs here or click Add PDF…”), and dragging a file over it highlights the drop
  target with an accent frame and a “Drop to add” prompt.
- **Selection counter.** Selecting pages shows “N of M selected” in the status line,
  reverting to the page count when nothing is selected.
- **Per-page progress.** While saving a merged PDF or converting to Word the status
  shows “page N of M”, next to the percentage.
- **Keyboard-shortcuts cheat sheet.** A ☰ Menu → Keyboard shortcuts entry lists the grid
  keys (zoom, select, cut/copy/paste, undo/redo, move, rotate, go to page), the list
  adapts to what the tool supports.
- **Drop PDFs onto the start-screen cards.** Dropping PDF files on the Merge, Split or
  PDF → Word card opens that tool with the files already loaded (an open tool receives
  them without a prompt).

### Changed
- **Release file names carry the version** (`iwoHelperDesktop-1.17.2.exe`,
  `iwoHelperDesktop-setup-1.17.2.exe`, and the `-x86` variants), so a downloaded file is
  self-identifying. The README buttons download the current release directly.

## [1.17.1] — 2026-07-24

### Added
- **Rotation in PDF → Word.** A sideways page can be rotated right in the grid
  (right-click or Ctrl+Shift+«+»/«−») and is straightened BEFORE layout analysis: its
  sideways text becomes normal lines and paragraphs, tables and margins are rebuilt in
  the upright space, images and a text-drawn stamp turn together with the page,
  and the produced `.docx` page swaps orientation accordingly. Without rotation the old
  behaviour is intact (sideways text is filtered out).
- **Mass rotation.** The rotate submenu now also turns ALL pages of the document at
  once, and a selection can still be rotated as before. The tile shows a badge with the
  current angle (90°/180°/270°).
- **Thumbnails zoom up to 400 px.** The grid draws tiles itself (the previous 256 px
  ImageList ceiling is gone) and re-renders visible pages at a larger width when zoomed
  in, so big tiles are sharp. Page numbers, selection and the dimming of cut pages are
  drawn with the tiles.
- **Ctrl+Z undoes order edits** in PDF Merge and PDF → Word: reorder, delete, paste,
  adding files and «Move after page» can be rolled back step by step (up to 50 steps).
  Rotation is a page property and is not part of the order history.
- **Move after page N…** in the context menu sends the selected pages right after the
  page you name (0 — to the very start) — no dragging across 300 thumbnails.
- **The insertion gap hints on hover.** Before clicking, the gap under the cursor shows
  a light insertion bar, so the caret lands exactly where expected.
- **Redo (Ctrl+Y or Ctrl+Shift+Z)** returns an undone gesture, and **rotations are now
  part of the Ctrl+Z history** in PDF Merge and PDF → Word — every snapshot carries both
  the page order and the angles, and a new gesture clears the redo branch, as in any
  editor.
- **Rotate buttons on the tile.** Hovering a thumbnail shows ↺ / ↻ chips on it (as in
  Acrobat) — one click turns that page without touching the selection.

### Changed
- **The README download buttons now download directly.** Release assets use stable
  names (`iwoHelperDesktop-setup.exe`, `iwoHelperDesktop-setup-x86.exe`), so the four
  buttons always fetch the latest build instead of opening the releases page.

### Fixed
- **The rotate chips are no longer clipped at large zoom.** Tiles bigger than the
  native 256 px item bounds were repainted only inside those bounds, so the ↺/↻
  buttons at the bottom of the tile (and the selection ring) could vanish. Repaints
  now invalidate the full grid cell with the same geometry the painter uses.

## [1.17.0] — 2026-07-24

### Added
- **Page numbers under thumbnails.** In PDF Merge and PDF → Word the tile caption is the
  page's position in the future document, in PDF Split it is the original page number.
  The file name moved to the tooltip, so the number is finally readable at a glance.
- **Cut, copy and paste pages.** Ctrl+X / Ctrl+C / Ctrl+V (and the context menu) move or
  duplicate the selected pages inside the window's page buffer — the precise way to move
  a long range: select 300–350 with Shift+Click, Ctrl+X, click the gap after page 2 (an
  insertion caret appears), Ctrl+V. Cut pages stay dimmed in place until pasted, Esc
  cancels. Pasting lands at the caret, after the selection, or at the end.
- **Drop PDF files straight onto the page grid.** The grid accepts dropped files
  everywhere, and in Merge and PDF → Word the pages are inserted AT the drop position
  (the insertion bar shows where) instead of always appending to the end.
- **Rotate pages in Merge and Split.** Right-click → rotate right/left 90°, or
  Ctrl+Shift+«+» / Ctrl+Shift+«−» — per selected page, as in Acrobat. The rotation is
  written into the produced PDF (all Split modes included) and composes with the page's
  own rotation. The source file is never modified.
- **Go to page (Ctrl+G).** Jumps the grid to a page by number — no scrolling through
  hundreds of thumbnails.
- **Auto-scroll while dragging.** Dragging pages (or files) near the top or bottom edge
  of the grid scrolls it — long upward drags no longer require dropping midway.

### Changed
- **Thumbnails are sharper on high-DPI displays.** Pages render in physical pixels
  (scaled by the monitor's DPI) instead of a fixed width, so tiles are no longer blurry
  at 125–150% scaling. At 100% nothing changes.

### Fixed
- **Large PDFs no longer accumulate memory while browsing thumbnails.** Rendered pages
  now live in a bounded LRU cache sized from a byte budget (halved in the 32-bit build),
  and an evicted page re-renders when shown again. Previously a 300-page document could
  pin hundreds of megabytes until the window closed.

## [1.16.9] — 2026-07-24

### Added
- **32-bit packages (x86).** Every release now also ships a portable
  `iwoHelperDesktop-x86.exe` and an `iwoHelperDesktop-setup-<version>-x86.exe` installer
  with a bundled 32-bit Ghostscript — for 32-bit editions of Windows. The tool set is
  identical to x64: Office automation and Ghostscript run as separate processes, so they
  do not depend on the app's bitness. In a 32-bit process the thumbnail document cache is
  halved (the address space is ~2 GB and every shown PDF is held in memory). CI builds
  and validates both architectures.
- **Windows 8.1 support.** The minimum OS is now Windows 8.1 — the oldest Windows with
  the `Windows.Data.Pdf` engine that renders page thumbnails. The installer checks for
  .NET Framework 4.8 (built into Windows 10 1903+ and Windows 11) and opens the download
  page when it is missing on 8.1.

## [1.16.8] — 2026-07-24

### Added
- **A date/number stamp over a form transfers cleanly.** A low image placed over a
  `______ № ______` blank-field row is carried as the image it is: the underlying placeholders
  are not duplicated as text, and invisible white marks around such images are dropped. White
  or invisible text is kept when it sits on a real backdrop (a scan image, a dark filled
  panel), so OCR layers and light-on-dark headers are unaffected.
- **Single-row side-by-side zones are laid out next to each other.** Pieces of one line
  separated by a huge gap (a label on the left, a note on the right, a caption and a date)
  become columns of a borderless row band instead of gluing into one string or stacking under
  each other. Inside table cells the line stays whole.
- **A column that starts lower keeps its height.** In a side-by-side band, a column whose
  content begins below the top of the band (a name opposite the last line of a multi-line
  block) is offset by the source distance instead of floating up to the first row.
- **A block at the bottom of a sparse page stays at the bottom.** The between-block
  spacing cap is raised (the pagination guard still prevents extra pages).

### Fixed
- **Switching the UI language keeps you in your window.** Windows are rebuilt with the active
  one last, so the tool chooser no longer pops up on top of the tool you switched the
  language from. A busy window (operation running) is simply raised back on top.
- **A numbered list that starts above one keeps its numbering.** For a list beginning at
  «5.» the start value is set on the applied list template (the document copy, not the
  user's gallery), and consecutive items no longer re-apply the template — Word used to
  renumber such lists from 1.
- **The installer's file properties now show the real version.** `VersionInfo*` resources
  of the setup executable carried 0.0.0.0 and an empty description.
- **First-line indents are now per paragraph, from the source.** A document-wide indent is
  applied only to paragraphs whose first line was actually indented, footnotes and other
  flush-left lines no longer inherit a false indent. In documents without a common indent,
  an actually indented paragraph keeps its own.
- **Narrow right-hand notes keep their horizontal position.** A left-aligned paragraph that
  starts deep inside the text area (a name, a corner mark) is anchored at its source
  position instead of jumping to the left margin.
- **A list item whose Word list template failed to apply keeps its marker.** The stripped
  `1.` / `•` is put back as text instead of silently losing the number.
- **Mixed-alphabet compound words keep their hyphen** at a line break (Cyrillic on either
  side of the break).
- **Footnote marks in short lines are recognised.** The dominant font size of a line is now
  taken from its widest word, so a two-word line («№ 250» + a small mark) resolves correctly.
- **Ruling connectivity is found via a spatial grid.** Densely dashed borders (thousands of
  strokes) no longer cost a quadratic pass during extraction.
- Small ones: the two-column split tolerates a narrower channel, the folder
  picker releases its COM objects deterministically, the justified label measures each word
  once per paint, and the tool cards reuse a cached title font.

## [1.16.7] — 2026-07-24

### Added
- **Intentional line breaks survive in short-line blocks.** A new line whose first word would
  still have fit on the previous one is a deliberate break, not a soft wrap — multi-line
  signatures, blank-field lines and contact footers now stay as separate paragraphs
  instead of being glued together and re-flowed by Word. The check uses the width actually
  available up to the neighbouring column, not just the block's own frame.
- **Vertical rhythm of the page is reproduced.** Extra space between blocks beyond the page's
  typical gap becomes spacing before the block (inside side-by-side header columns too). A
  pagination damper trims the added spacing if it would push the document past the source page
  count, so no document gains pages compared to previous versions.
- **Font-size based footnote markers.** A digits-only word set noticeably smaller than its line
  is recognised as a superscript footnote mark even when its glyph box sits on the baseline —
  markers of one document no longer come out half raised, half inline.
- **Crash safety net.** Unhandled UI exceptions show a branded error dialog and the app keeps
  running. All unhandled exceptions are appended to `crash.log` in the app profile for support.

### Fixed
- **Compound words keep their hyphen.** A hyphen at a line break after a Cyrillic letter is kept
  when the lines are joined: office suites do not auto-hyphenate, so such a hyphen is part of
  the word. Latin soft hyphenation is still removed.
- **A first-line-indented line is not mistaken for a centred one.** A single line starting
  exactly at the document's first-line indent stays a regular paragraph even when its right gap
  happens to be nearly symmetric. Real centred lines (detected via a tight indent cluster) are
  unaffected.
- **A table Word failed to build is no longer lost.** The partially built table is removed and
  the cell contents are emitted as plain paragraphs, so the text survives the rare COM failure.
- **English Office error messages are classified too.** Permanently unreadable workbooks
  (password, damaged, wrong format) are reported at once instead of pointless retries when
  Office speaks English.
- **Language switch no longer interrupts a busy window.** Switching the UI language while an
  operation is running skips that window (it keeps the old language) instead of popping a
  "cancel the operation?" prompt in the middle of the switch.
- Small ones: the header-band fallback centres images against the real page width, usage
  counters are guarded by a cross-process mutex, the About dialog releases its icon bitmap
  deterministically, and out-of-memory errors are no longer masked as "the file is damaged".

## [1.16.6] — 2026-07-23

### Added
- **Two-column headers are laid out side by side.** A header area that the reading-order
  analysis splits into columns (one block on the left, another on the right) is now emitted as a
  borderless Word table, so the columns sit next to each other exactly as in the original — a
  logo is centred above its column, and the right-hand block lines up with the top of the left
  one — instead of being stacked one under another. Single-column pages and page-wide content
  are unaffected.
- **Label/value forms keep their vertical grouping.** In a borderless grid (a receipt), extra
  space between groups of fields is measured from the source and reproduced as spacing after the
  row, instead of collapsing every row to a uniform tight pitch.

### Fixed
- **Images with a soft transparency mask no longer come out on a black background.** The raw
  decoder ignores such a mask and fills the transparent areas of a logo or a stamp with black.
  The image is now rendered from the page instead, compositing the mask onto white as it looks
  on the page.

## [1.16.5] — 2026-07-23

### Added
- **PDF → Word rebuilds label/value forms without any ruling as borderless tables.** A block of
  rows where each row splits by a wide inner gap into segments whose left edges line up into
  columns (a receipt-style “label … value” layout with no drawn borders) becomes a real Word
  table with borders off, so the pairs stay aligned on their own rows instead of being read as
  one flat stream. Thresholds are strict — ordinary paragraphs, justified text and single
  “signature … date” rows are left as plain text.
- **Blank fill-in lines survive as underscore placeholders.** A standalone horizontal rule
  (a fill-in line such as `______ №  ______`), including one drawn in several collinear pieces
  with gaps under the labels, is carried across as an underscore run sized to the line, pieces
  under a word stay underlines, double-drawn lines are de-duplicated, and a rule inside a table
  frame is left to the table.

### Fixed
- **Two-column headers keep each block in its own column.** A centred paragraph that belongs
  to a narrow column (a block on the right, a block on the left) is now centred
  **within that column** via left/right indents, instead of on the whole page where the two
  columns overlapped into one confused centred stack. Full-width content and a page-wide centred
  title are unaffected.
- **Installed-but-non-native fonts no longer letter-space cyrillic.** Word tags cyrillic runs set
  in some installed families with an East-Asian hint and spaces the letters out under CJK
  justification. Cyrillic text now stays only in Word-native families and otherwise falls back to
  Times New Roman (a metric twin of the common clones), latin text keeps its original font.
- **Private-use-area glyphs are dropped.** Field-placeholder and symbol-font glyphs that render as
  empty boxes outside their font are no longer emitted as garbage.

## [1.16.4] — 2026-07-23

### Added
- **PDF → Word reads multi‑column layouts in the correct order.** Pages are decomposed by a
  recursive XY‑cut over the empty gutters (horizontal bands into “floors”, a full‑height gap
  inside a floor into columns), and each block is laid out with its own column geometry. A
  two‑column header now comes out as coherent blocks — the left column in full, then the
  right one — instead of lines from both columns interleaved one after another. The same cut
  orders whole page blocks (paragraphs, tables, images), so a picture that belongs to the left
  column stays with it. Single‑column pages are a single block and read exactly as before.
- **Pseudo‑bold text (drawn twice with a ~0.3 pt offset) is de‑duplicated and set in real
  bold.** Some generators imitate bold by painting every glyph twice, extraction used to
  produce doubled letters (“74” became “7744”), and the doubling also hid the intended weight.
  Duplicate glyphs are now dropped (the tolerance is far below the advance of genuinely
  repeated characters, so “77” is never collapsed) and the affected words are written bold.
- **Rotated edge text is filtered out.** Vertical strings along a page edge used to shred into
  dozens of single‑character paragraphs wedged between normal
  lines, text that is not horizontal is now skipped entirely.

### Fixed
- **First‑line indent survives a header block.** The indent share used to be measured against
  all paragraphs of a page, so a dozen short header lines diluted it below the threshold
  and the body lost its indent. The share is now measured among justified paragraphs (with the
  old page‑wide rule as a fallback for ragged‑right documents).
- **Page margins account for images.** Margins were computed from words only: a logo above the
  first line was pushed out of the top margin and shifted the whole page down by its height, and
  an image below the last line inflated the bottom margin up to half a page — which, combined
  with inline image insertion, spilled a one‑page document onto a second page. Margins
  now cover words and images, and the bottom margin is capped (it only limits page fill, so
  the cap cannot hurt fidelity).
- **Inside a table cell, “label … value” rows are no longer split into separate columns** by
  the new column detection: cell content is laid out without the vertical cut.
- **Reading order of adjacent text lines no longer flips.** Blocks used to be grouped into a
  “row band” by the closeness of their tops (±12 pt), which could reorder two neighbouring
  lines by their left edges, a band now requires a real vertical overlap of at least half the
  smaller height (as with words within a line), so stacked lines always read top to bottom
  while genuinely side‑by‑side tables and images still read left to right.

### Changed
- **A scanned PDF is rejected instantly.** The scan check now runs before any page image is
  decoded: a 15 MB scan is turned down in ~0.1 s without ~18 MB of raster data ever being
  built, instead of ~3 s of wasted decoding.

## [1.16.3] — 2026-07-23

### Fixed
- **PDF merge/split: “could not save … a file with a user‑mapped section open.”** Saving the
  result under the **same name** as one of the PDFs shown in the thumbnail grid failed, because
  the thumbnail engine (`Windows.Data.Pdf`) loaded each displayed file straight from disk and
  kept it memory‑mapped, so Windows blocked overwriting it. Thumbnails are now rendered from an
  in‑memory copy, so a file shown in the grid is never locked on disk and can be overwritten by
  the merge/split output. A regression guard covers this in the thumbnail self‑check.
- **PDF → Word: “could not save” on some Word builds.** On a few Word versions/states the
  late‑bound `SaveAs2` call did not resolve and the conversion ended with an error
  (“…does not contain a definition for SaveAs2”). The save now falls back to the classic
  `SaveAs`, which is present in every Word version, so the `.docx` is written normally. A
  genuine write failure (a file open in another program) is still reported as before.
- **PDF → Word no longer spills onto extra pages, and pagination is now deterministic.** The
  produced document set explicit single line spacing with no space before/after paragraphs,
  instead of inheriting the machine’s `Normal` template (whose “Office” default adds 8 pt
  after every paragraph and 1.08 line spacing). A dense source that separates paragraphs by a
  first‑line indent — not by blank space — now keeps its original page count and looks the
  same on any machine.

### Changed
- **PDF → Word keeps a centered, multi‑line title centered, line for line.** A heading whose
  lines are centered about a common axis — even when some lines are wide enough to reach the
  margins and look justified — is recognised as centered and each source line is kept on its own
  centred line, matching the original, instead of being split into a left/justified fragment
  with a stray last word (or collapsed into one re‑wrapped line). Detection is by the shared
  centre axis plus at least one clearly floating line, so ordinary justified body text (whose
  short last line hugs the left margin, and whose first line’s indent shifts the axis) is never
  mistaken for centered. Centered blocks also no longer skew the first‑line‑indent measurement.
- **PDF → Word centers a horizontally centered image.** An image that sits centered on the page
  (a logo or emblem with near‑equal left/right margins) is placed centered in Word, as in the
  source, images anchored to a margin, or stamps and signatures off to one side, stay left as
  before. The centering test is a pure, unit‑tested method.

## [1.16.2] — 2026-07-23

### Added
- **English interface, with a language switch.** The whole UI — windows, menus, buttons,
  tooltips and messages — is now available in English as well as Russian. Pick the language
  from a globe icon in the top‑right corner of the start screen, or from **☰ Menu →
  Язык / Language** in any tool window, each option shows a small country flag. The choice is
  saved and applied instantly — open windows rebuild in the new language on the spot. The
  tool menu was renamed from “Справка” to **☰ Menu**. Content of generated documents (the
  cover note, reports, the digest’s table of contents) intentionally stays in Russian.
- The **About** box text was refreshed to list all current tools (Excel digest, merge, split
  and compress PDFs, PDF → Word).

## [1.16.1] — 2026-07-23

### Added
- **PDF → Word rebuilds numbered and bulleted lists as native Word lists.** A paragraph that
  starts with a list marker (`1.`, `12)`, `•`, `—`, …) becomes a real Word list item: the
  marker is dropped and Word draws its own, with a proper hanging indent. Numbering continues
  across paragraphs nested inside an item (an item can hold ordinary text and the next item
  keeps counting 2, 3, 4), and a fresh list restarts at 1 — so a second list after the first
  starts over correctly. Detection is deliberately strict (marker punctuation + space +
  content) so “2025 г.” or “12.5 %” are never mistaken for a list. The marker classifier is a
  pure, unit‑tested method.
- **PDF → Word turns a text‑drawn stamp into an image.** When a stamp is drawn as text rather
  than a picture, that region is rendered with the bundled Ghostscript and placed as an image,
  and its text is removed so it isn’t duplicated. Detection needs several anchor words in a
  compact box, so ordinary prose is never affected, if the region can’t be rendered (no
  Ghostscript) the text is kept unchanged — no regression. Picture stamps keep transferring as
  images as before. The region detector is a pure, unit‑tested method.

## [1.16.0] — 2026-07-23

### Added
- **PDF → Word converts several PDFs into one document.** Add several PDFs (the button now
  takes a multi‑selection, or drop several at once) and every file’s pages appear in one
  thumbnail grid, reorder or drop pages across all of them and convert once to a single
  `.docx` in the shown order. Pages from different files (and different page sizes) sit in
  the same document. The default output name is the file’s own name for a single source, or
  “Объединённый.docx” for several. The page assembler is a pure, unit‑tested method.

### Fixed
- **An image the PDF decoder can’t handle is now recovered instead of dropped.** When the
  embedded image decoder fails or returns a single solid colour (which previously left the
  image skipped, e.g. a monochrome barcode coming out as a black box), the page region is
  rendered with the bundled Ghostscript and the image is cropped out by its bounding box —
  so it transfers faithfully, exactly as drawn. Normal images still take the fast decode
  path untouched, the fallback simply skips when Ghostscript is unavailable.

## [1.15.0] — 2026-07-22

### Added
- **PDF → Word now reconstructs bordered tables as real Word tables.** Digital PDFs (Word
  exports, “Microsoft Print to PDF”, browser exports) draw table grids as ruling lines,
  those lines are now read from the page vector graphics and, together with the words,
  turned back into a Word table — column widths from the ruling geometry, per‑cell text
  (each cell laid out by the same reading‑order engine as the body), and **merged cells**
  (colspan/rowspan) inferred from missing internal borders. Before, a table came out as
  garbled text read straight across the cells, now the structure and every cell are
  correct. Detection is conservative: only clearly bordered ≥2×2 grids become tables, and
  on any doubt the words stay in the ordinary text flow, so output is never worse than
  before. Borderless tables, multi‑column layouts and lists are still flattened.
- **Per‑page size and orientation.** Each source page becomes its own Word section with its
  own page size, so a document that mixes portrait and landscape pages is preserved and a
  wide landscape table is no longer clipped by a portrait page.
- **Underline is carried over.** In a digital PDF an underline is a drawn line under the
  text, not a text attribute, a horizontal rule sitting on a word’s baseline across its
  width now marks that word underlined in Word. A full‑width rule (a section divider) is
  not mistaken for an underline — it is far wider than the word above it.
- **Images that PdfPig can’t turn into PNG (typically JPEG/DCTDecode) are now recovered.**
  The raw image stream is decoded through GDI and re‑saved as PNG, so a JPEG‑embedded photo
  transfers instead of silently disappearing.
- **Left‑sidebar (two‑column) pages read in the right order.** A narrow left column of
  labels or dates used to interleave word‑by‑word into the body text and corrupt its
  indentation. Such a sidebar is now detected — the body and the sidebar are laid out
  separately (each with its own margins) and merged top‑to‑bottom — so the body reads
  cleanly and each label sits by its section. Detection is conservative (it needs a clear
  left column separated by wide in‑line gaps) and no‑ops on ordinary single‑column pages.
  Dense multi‑column body text is still out of scope.
- **Side‑by‑side blocks read left‑to‑right.** Tables (and paragraphs) that sit on the same
  row band — e.g. two small tables placed next to each other — are now ordered left before
  right instead of by vertical position alone, so they no longer come out swapped.

### Changed
- **Tool subtitles reworded** to state each tool’s purpose (no semicolons), and the header
  now lays the subtitle across the **full width** below the “⌂ Главная” button, so a longer
  subtitle is not clipped with an ellipsis at the default or minimum window size.
- **PDF → Word Help lists what is now supported** (underline, bordered tables with merges,
  per‑page orientation) and narrows the remaining limitations to borderless tables,
  multi‑column/list layouts and stamp graphics (whose inner text
  is still extracted as text).

### Fixed
- **A broken image no longer prints as a solid black box.** Some monochrome images that the
  PDF image decoder mangles into a single colour are now detected and skipped, so any text
  they accompanied is kept while the bogus rectangle is dropped.

## [1.14.1] — 2026-07-22

### Fixed
- **Crash on exit after using PDF → Word (access violation in the hidden
  .NET‑BroadcastEventWindow).** Root cause: the WinRT runtime (`Windows.Data.Pdf`, used for
  page thumbnails) crashes inside the native `DLL_PROCESS_DETACH` performed by `ExitProcess`
  during an orderly shutdown. The forced exit now uses **`TerminateProcess`**, which ends the
  process without the DLL detach phase (and without finalizers). This is safe by design: all
  critical cleanup — saving settings, `Quit` on Excel/Word COM — runs deterministically
  before the exit call.
- **The progress bar no longer collides with the zoom slider.** The zoom `TrackBar` is 45 px
  tall (WinForms enforces this AutoSize height, the layout assumed 30), so on all three PDF
  screens it bled over the progress bar underneath. The bottom strip is now laid out in three
  non‑overlapping rows — zoom + compression, the progress bar, status + action button — and
  the PDF tool windows grew 40 px taller (default and minimum size) so the thumbnail grid
  keeps its height. Anchors are unchanged, so resizing keeps the rows apart at any window
  size.
- **A long status line no longer runs under the action button** on the PDF screens. The
  status label used to auto‑size without a width limit and shares its row with the button,
  so a long result (e.g. the “file is large — enable Compression” hint) could disappear
  under it. The label is now clipped at the button with an ellipsis, and the built‑in
  tooltip shows the full text.
- **Two small `ToolTip` leaks.** The Excel window and the Statistics window created a
  `ToolTip` but never disposed it (a `ToolTip` is a component, not a child control, so it
  is not released automatically). Both now dispose it, matching the PDF windows.

### Changed
- **Page‑grid hotkeys unified in the base form.** The identical Delete / Alt+←→ / Ctrl+A /
  Enter handling was duplicated in PDF Merge and PDF → Word (the latter reaching into the
  former’s classifier), and PDF Split carried its own Ctrl+A copy. The classifier and the
  dispatch now live once in `PdfToolFormBase`. Editable grids override two small hooks, and
  the behaviour of all three screens is bit‑for‑bit unchanged (the unit test moved along).
- **PDF Split: picking a mode from the opened drop‑down focuses its input field** (ranges
  or page count), so you can type right away. Arrow keys on the closed list still cycle
  modes without stealing focus.
- **Embedded‑resource loading fails fast on a short read.** If an embedded assembly
  resource cannot be read completely (a corrupted exe), the loader now throws a clear
  end‑of‑stream error instead of handing truncated bytes to `Assembly.Load`.

## [1.14.0] — 2026-07-22

### Added
- **PDF → Word: reorder and drop pages before converting.** The page‑thumbnail grid on the
  PDF → Word screen is now interactive — **drag** a thumbnail to a new position, or select one
  and use **◀ Раньше / Позже ▶** (Alt+←/→). Remove pages you don’t need with **Удалить**
  (Delete), and Ctrl+A selects all. Word receives the pages in exactly the order shown, with the
  dropped pages excluded. The order model (`PdfPageOrder`) and the reorder grid are the ones
  already used by PDF Merge (reused, not re‑implemented). The conversion picks and reorders the
  extracted pages through a pure, unit‑tested `SelectPages`, and the progress bar counts the
  selected pages. Converting the whole document unchanged still works exactly as before.

## [1.13.11] — 2026-07-22

### Added
- **Progress bar on all three PDF screens** (PDF Merge, PDF Split, PDF → Word) plus the
  Windows taskbar‑button progress. It is a **real** determinate bar driven by the actual work,
  not a timer: PDF → Word reports each page of both passes (text extraction, then writing to
  Word), Merge reports each page added, and Split reports each part written and then each part
  compressed. Shared once in `PdfToolFormBase` (DRY) and shown by every PDF tool. Updates are
  marshalled to the UI thread, throttled by whole percent, and the bar is drawn **exactly** at
  each value (bypassing the Vista progress‑bar catch‑up animation that would otherwise leave a
  visible gap in the fill during rapid updates), so there is no flooding, no flicker and no
  broken fill. The bar sits in the free strip band **above** the status/action row, so it cannot
  overlap the buttons, the zoom slider, the compression picker or the page grid. The
  “done/total → percent” calculation is a pure method covered by unit tests (division‑by‑zero
  and clamping included).

### Changed
- **PDF → Word Help → “How to use” now lists the real limitations.** Scanned image‑only PDFs
  are not supported. If the source font is not installed the text is set in Times New Roman.
  Tables, side‑boxes, multiple columns and lists are flattened to single‑column paragraphs.
  Underline is not carried over (in PDF it is a drawn line, not a text attribute). A PDF saved
  with a broken text encoding (no valid ToUnicode) extracts as unreadable text — a defect of
  the file itself, checkable by copying the text inside the PDF.

## [1.13.10] — 2026-07-22

### Fixed
- **PDF → Word: Cyrillic no longer comes out letter‑spaced (“р а з р я д к а”).** This was the
  visible bug — every Cyrillic word rendered with gaps between the letters while Latin stayed
  solid. Root cause: the source font (e.g. **PT Astra Serif**) is often **not installed** on
  the target machine. When Word is handed an uninstalled font it routes Cyrillic to the East
  Asian fallback slot (`rFonts w:hint="eastAsia"`), and a justified paragraph then gets
  **CJK‑style character distribution** — the letters are spread to fill the line. The extracted
  text was always correct (single spaces between words). Only Word’s rendering spread it, which
  is why a text‑only check missed it. Fix: each run’s font is resolved against the installed
  fonts — an installed family is kept, an unknown one falls back to Times New Roman — so
  Cyrillic stays in the normal (hAnsi) slot and justifies like ordinary text. Verified by
  rendering the output back to PDF: the page now matches the reference. `w:hint="eastAsia"`
  drops from 184 to 0 on the sample page.
- **Font‑family normalisation splits an all‑caps prefix from the following word** —
  “PTAstraSerif” → **“PT Astra Serif”** (and “MSGothic” → “MS Gothic”), so where the font *is*
  installed it is matched and kept instead of being replaced.
- **PDF → Word: words no longer glue together.** A separate bug in the line word‑join: it used
  a gap threshold of **0.2 × font size**, but in narrow fonts (e.g. Calibri Light) a real
  inter‑word space is only ≈ 0.18 × size — so the space was dropped and neighbouring words
  merged (two adjacent words glued into one). The threshold is now **0.08 × size**, safely
  below the smallest real word‑space measured across the sample documents (0.179). Only truly touching fragments (gap < 0.08) are
  glued. Verified across the sample set: extracted text is character‑for‑character identical
  except that dropped spaces are restored (16 of 29 documents gained spaces, one of them
  +213), with no document losing a space.
- **Correction to the 1.13.9 note on letter‑spacing.** PdfPig’s default word extractor does
  **not** over‑split PT Astra Serif into letter fragments (verified): that document comes out
  with solid words. The earlier gap‑based join was unnecessary and, at 0.2, actively harmful
  (see above). The real letter‑spacing was the East Asian rendering issue, now fixed.

## [1.13.9] — 2026-07-21

### Added
- **New tool: “PDF → Word”** (fourth start‑screen card) — extracts the text layer of a
  **born‑digital** PDF (saved from Word, “Microsoft Print to PDF”, exported from a browser)
  into an editable `.docx`. Text is read with **PdfPig** (Apache 2.0, embedded), the `.docx`
  is written through Word COM. Scanned documents (image pages with no text layer) are not
  supported yet — a clear message is shown and the file is untouched. `PdfToWordService`,
  `PdfTextExtract`, `OcrLayout`, `FontNames` and `WordDocxWriter` are unit‑tested, and
  `verify_pdfword.ps1` is an end‑to‑end round‑trip through Word.
- **Reading‑order layout** for PDF → Word — words with their boxes become lines (by vertical
  overlap, so thin punctuation such as an em‑dash stays on its line), lines become paragraphs
  split by any of three signals: a larger vertical gap, a first‑line indent, or a short last
  line in justified text. Line wraps are joined and hyphen‑wraps de‑hyphenated. Words on a
  line are joined by their horizontal gap, so a font whose glyphs PdfPig over‑splits (e.g.
  PT Astra Serif) does not come out letter‑spaced.
- **Formatting inherited from the source**, per run: font family (normalised from the PDF
  font name, no longer hard‑coded Times New Roman), size, bold, italic, colour, and
  super/subscript. Per paragraph: alignment (left / justify / centre — a centred line such as
  a page number is centred, not stretched) and the first‑line indent (красная строка), applied
  only when most paragraphs use one so a flush‑left document is left alone.
- **Page geometry inherited** — the document page size and margins are taken from the source
  (page media box and the text bounding box), clamped to sane limits.
- **Images** — each page’s raster images are extracted (PdfPig `GetImages` → PNG) and placed
  inline in reading order, sized to their PDF bounds. Formats that do not decode to PNG are
  skipped, and a broken image never derails the document.
- **Hyperlinks** — link annotations (`GetHyperlinks`) are carried through to real Word
  hyperlinks over the matching text.
- **Usage statistics** — a “PDF → Word” counter (with a row in the Statistics window).

### Changed
- **Action buttons moved to the bottom‑right** in “PDF → Word” (“Convert to Word…”) and
  “PDF Split” (“Extract…/Split…”), matching “Save PDF…” in “PDF Merge”.

### Notes
- **Underline is not inherited** — in PDF an underline is a drawn line, not a text attribute,
  so it cannot be read from the text layer. Tables, multi‑column layouts and bullet/number
  lists are linearised as plain paragraphs.

## [1.13.8] — 2026-07-21

### Added
- **Split tool: a gentle "enable Compression" hint.** When an extract/split is done
  *without* compression and the result comes out almost as large as the source (≥ 90% and
  over 1 MB — which happens when pages share heavy resources that are copied along with
  them), the status line appends an unobtrusive note suggesting the Compression option to
  reduce the size. Purely advisory, with no change to the produced files.

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
- New reusable, unit‑tested `LruCache<T>` (bounded least‑recently‑used cache). The
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
  Powered by **Ghostscript** invoked as a separate process. The compressed file is
  written to `<pdf>.gstmp`, validated (exit code + `%PDF-` header + strictly smaller)
  and only then replaces the original — an already-optimized PDF is left untouched.
  Output uses PDF 1.4 (classic xref) so a compressed file can still be re-merged/split
  by the app. A shared `CompressionPicker` control is used by both tools (DRY). The
  work runs on the background thread **before** the file is opened, so the replace
  never hits a viewer lock. Compression is a no-op if Ghostscript is absent.
  Pure functions (`Preset`, `BuildArguments`, `ShouldReplace`, `PickFirstExisting`)
  and a live end-to-end compression test (real size reduction, pages preserved) are
  covered by unit tests, and `--gscheck` is a CI smoke check.
- **Installer (Inno Setup)** alongside the portable exe: `iwoHelperDesktop-setup-*.exe`
  installs the app **and bundles Ghostscript**, so compression works out of the box.
  Default install is **per-user without administrator rights** (`%LOCALAPPDATA%`),
  with an option to install for all users (Program Files, requires admin). The
  installer's welcome page **explicitly states the per-user default**. Built and
  signed locally via `tools\make_installer.ps1` (`tools\stage_gs.ps1` prepares the
  Ghostscript subset). Ghostscript is bundled under its own AGPL license (invoked as
  a separate process — mere aggregation, and the app stays MIT).
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
  number used to be (left), and the right button is now “About” (was “Check for updates”).
  The buttons were enlarged and raised slightly. The version is still shown in the title
  bar and the About dialog.
- **“How to use”** (Help) in both PDF tools now documents the compression dropdown and
  the signature caveat.
- **Branded installer wizard image** (blue gradient + logo + “iwo”) replaces the default
  Inno graphic on the welcome/finish pages, generated by `tools\make_wizard_images.ps1`.
- Compression's in-place replace now uses a **rename-aside** strategy (original → `.gsbak`
  → compressed in place → backup removed, restored on any failure). This works on network
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
  All calls still go through the `Dialogs` facade (`MessageForm`), and button placement
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
  parsing and version comparison are unit-tested, and the network call runs off the UI thread.
- **“Statistics”** in the Help menu: local counters (no telemetry) of operations —
  Excel digests, PDF merges, page extractions, and splits by mode. Manual **Clear**
  and optional **auto-clear** (daily / every 7 / every 30 days). Counters use
  read-modify-write so concurrent windows can't lose increments, and the auto-clear
  period logic (`ShouldAutoClear`) is unit-tested.

## [1.11.2] — 2026-07-20

### Changed
- **PDF Split — you can now choose the output name in every mode.** The
  split-into-many modes (ranges, every N pages, bookmarks) previously only let you
  pick the folder and reused the source file name. Now a save dialog lets you set
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
  stays responsive. Visible-range windowing (`ClampWindow`) is unit-tested, and the
  no-crash + lazy behaviour was verified in a real message loop (22/60 pages
  rendered for a 60-page file).
- Drag-and-drop path extraction for the PDF tools was de-duplicated into a shared
  `PdfDrop` helper (DRY).

## [1.11.0] — 2026-07-20

### Added
- **New tool: “PDF Split”** (third start-screen card), complementing “PDF Merge”.
  Open one PDF, see its pages as thumbnails, and either extract or split — following
  the modes of leading offline tools (PDFsam, Acrobat):
  - **Extract selected** — pick pages in the grid (Ctrl+A = all) → one new PDF.
  - **By ranges** — “1-3, 5, 8-”: each range → its own file.
  - **Every N pages** — equal chunks (N=1 → one file per page).
  - **By bookmarks** — one file per top-level bookmark, named from the titles.
  Pages are copied as-is (no re-conversion). The source is never modified, and output
  names are never overwritten (a number is appended). The engine (`PdfSplitService`,
  `PageRanges`) is unit-tested and validated live on real PDFs, including bookmarks.
- The PDF page-thumbnail grid was extracted into a reusable `PdfPageGrid` control
  and is now shared by both PDF tools (DRY), and the Merge tool was refactored onto it
  with no behaviour change.

## [1.10.7] — 2026-07-20

### Added
- **Keyboard shortcuts in both tools' lists**, handled in `ProcessCmdKey` so they
  are reliable (before dialog-key/menu handling), unit-tested via pure classifiers:
  - **PDF Merge**: `Delete` removes the selected pages, `Alt+←/→` reorder,
    `Ctrl+A` selects all, `Enter` no longer triggers a save from the list.
  - **Excel Digest**: added `Ctrl+A` (select all) and `Delete` (exclude/uncheck
    selected). Existing `Alt+↑/↓` (reorder), `Ctrl+C` (copy) and `Enter`-suppression
    are consolidated into `ProcessCmdKey` (copy/select-all now also work during a run).
  - Shortcut hints added to button tooltips and the “How to use” help.

### Fixed
- PDF reorder/remove could run from the keyboard **during a save** (the buttons
  were disabled but the methods weren't guarded). `MoveSelected`/`OnRemoveClick`
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
  (“iwo Helper Desktop 1.10”). It is now “Свод Excel”, so it is distinct in the
  title bar and Task Manager (the PDF tool was already correct).
- **Keyboard handling in the file list now actually works.** `Enter` (the form's
  default button) is a dialog key intercepted before `KeyDown`, so the previous
  suppression never fired, and `Alt+↑/↓` were unreliable next to the menu. Both are
  now handled in `ProcessCmdKey` (which runs first): `Enter` in the list no longer
  starts the merge, `Alt+↑/↓` reorder reliably. Routing is unit-tested
  (`ClassifyListKey`).
- **Self-healing restart no longer double-counts results in the UI.** When a wedged
  Excel instance is restarted, the previous pass is replayed. The merge service now
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
  selected file, and `Enter` in the list no longer triggers the merge.
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
    as corrupt. A `.xlsx`/`.xlsm`/`.xlsb` whose container is OLE2 is an encrypted
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
  `ListReorder` — shared with the PDF page order, DRY), and the merge service now
  takes an explicit file list (`Merge(files, …)`, `PrepareSourceList`).
- **Branded window header**: the top of every window (the start screen and both
  tools) carries an accent-green gradient header band (`HeaderBand`) with the
  title and subtitle, and the “◀ Back to menu” button sits on it. On Windows 11 the
  system title bar is tinted to match via DWM (`WindowChrome`), while on Windows 10
  the title bar stays default and the header band provides the branding. Unit tests:
  `WindowChrome` COLORREF packing, `HeaderBand` construction.

### Changed
- **README and CHANGELOG are now in English**, and the changelog moved to
  `docs/CHANGELOG.md`.

## [1.9.0] — 2026-07-17

### Added
- **Sheet selection in “Excel Digest”**: a “Sheets” drop-down — “First sheet
  only” (default, as before) or “All sheets”. In “all sheets” mode every visible
  sheet of each file is transferred with names “file · sheet”, and the table of
  contents and the report get a row per sheet. CLI flag `--allsheets`. The result
  model is now one record per sheet, and a retry of skipped files correctly expands
  a file into several sheets. Tests: `SheetBaseName`, `FileCount`, multi-sheet retry
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
  DedovMosol/…), and the “Telegram:”, “GitHub:” labels are plain text.

### Fixed
- A chooser card fired twice on a single click (the base control raised Click and
  the handler raised it again): because of this the very first open showed “tool
  already open”. The duplicate call was removed and verified with window messages
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
- **PDF thumbnail zoom**: a slider and Ctrl+mouse wheel. A page is rendered once,
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
  shipped, so nothing is installed on the target machine.

### Added
- **Tool-chooser start screen**: “Excel Digest” and “PDF Merge” cards with
  descriptions. After a tool is closed the chooser is shown again.
- **PDF page thumbnails**: the “PDF Merge” tool shows a grid of previews of the
  real pages (the system Windows.Data.Pdf engine), reordered by dragging
  thumbnails and with buttons. Rendering runs in the background (a separate
  thread), and if the engine is unavailable (e.g. on Windows Server) it falls back
  to placeholders as designed. Tests: `verify_thumb.ps1` (rendering and aspect ratio)
  and `--thumbcheck` (clean process exit after WinRT rendering).

### Fixed
- Forced process exit (`FastExit`/`ExitProcess`) after working with WinRT: the
  normal finalization of the Windows.Data.Pdf COM wrappers crashed the process on
  unload, so the critical cleanup (settings, COM Quit for Excel/Word) runs
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
  digest (period, counters, a table of skipped files with reasons), generated through
  the COM of an installed Word. The pure
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
  without a full rebuild, and the table of contents is regenerated from the overall
  result, the order and the successful sheets are preserved.
- **Copying log rows** — Ctrl+C or the context menu: a “file → sheet → reason”
  row in the report format, handy to forward to the owner of a broken file.
- **CHANGELOG.md** (this file), linked from the README.

### Changed
- The “Replace formulas with values” option is no longer **remembered** between
  runs: the mode changes the digest content and is enabled deliberately each time.

## [1.4.0] — 2026-07-16

### Added
- **Output format selection**: `.xlsx`, `.xlsm`, `.xlsb`, `.xls` (a drop-down — in
  the CLI the format is derived from the path extension).
- An integration run to `.xlsb` in the common test set.

### Changed
- Branded checkboxes: a white check on a green background, hover, a focus ring
  (`AccentCheckBox`).
- The log columns share the window width proportionally.
- Per the Windows guidelines, the ellipses were removed from the “How to use” and
  “About” items, and the punctuation in the help was fixed.

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
- CI (GitHub Actions): build, unit tests, GUI smoke. The exe is published to
  Releases on a `v*` tag.
- `tools/sign.ps1` — signing the exe with a self-signed certificate (SHA256).
- `tests/run_all` — the whole test pyramid in one command.

## [1.1.1] — 2026-07-16

### Fixed
- **Escaping strings when writing to cells**: a file name or a formula's string
  result that started with “=” turned into a formula (injection), and a leading
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
  references. Merged cells are handled by a per-cell fallback.
- An OLE message filter: automatic retry of COM calls rejected by a busy Excel.
- A manual recalculation mode during the merge (faster with formulas).
- Unit tests without external frameworks (`tests/build_tests.cmd`).

## [1.0.0] — 2026-07-16

First release.

- Merges the first visible sheet of every Excel file in a folder into a single
  `.xlsx` through the COM of an installed Excel — without losing formatting,
  formulas, merged cells, charts and pivot tables.
- Source formats: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`. Broken and password-protected
  files are skipped with a reason, hidden sheets are not transferred, and sheet
  names come from file names with deduplication and a 31-character limit.
- WinForms GUI: live validation, processing progress, a colour-coded log,
  folder drag-and-drop, path memory, an icon and branded styling.
- A `--cli` mode for scripts and automated tests, integration tests on a corpus
  of 13 files, and a single exe ~65 KB with no dependencies (.NET Framework 4.8,
  the compiler bundled with Windows).
