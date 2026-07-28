# Privacy Policy

**Application:** iwo Helper Desktop
**Last updated:** 2026-07-27

## Summary

iwo Helper Desktop is an **offline desktop application**. It does **not** collect,
transmit, sell, or share any personal data, and it does **not** contain any telemetry,
analytics, advertising, or user accounts. Everything you do with your files happens
**locally on your own computer**. The only time the app talks to the network is the
update check — when **you** click “Check for updates,” and once at startup — and even
then it only reads a version number from GitHub. No file contents or personal data are
ever sent, and the startup check can be turned off.

## What the app does with your files

The tools read the files you choose and write results to the folders you choose:

- **PDF Merge** — build one PDF out of several,
- **PDF Split** — extract or split pages,
- **PDF → Word** — turn the text layer of a born‑digital PDF into a `.docx`,
- **More operations** — compress, save pages as images, extract text to `.txt`, convert
  to grayscale, repair a damaged file, view and edit document properties,
- **Excel Digest** — combine sheets from several workbooks, with an optional Word cover note.

This processing runs entirely on your machine using local components:

- Microsoft Excel / Word automation (COM) for the Excel Digest and its Word cover note,
  and for writing the `.docx` in PDF → Word,
- the embedded PdfSharp library for reading and writing PDFs (merge, split, blank pages,
  document properties),
- the embedded PdfPig library for extracting the text layer of born‑digital PDFs
  (PDF → Word and “extract text to `.txt`”),
- Ghostscript, run as a separate local process, for compression, grayscale conversion,
  repairing damaged files, and for rendering the picture regions that PDF → Word puts
  into the `.docx`,
- the built‑in Windows PDF engine (`Windows.Data.Pdf`) for rendering page thumbnails,
  the full‑size preview, and the images produced by “save pages as images.”

**Your documents are never uploaded, copied off your device, or transmitted anywhere.**
The app does not modify your source files except where you explicitly ask it to write
output. Split, PDF → Word, image export, and text export never change the source.

**Document properties.** The “document properties” screen shows what is already stored
inside the PDF you opened — title, author, subject, keywords — and can be used to change
or clear them. That is often the only place a file still carries a personal name, so the
app lets you see and remove it. The edited copy is written where you choose, on your
machine only, and nothing is read or sent beyond that.

**Printing** hands the pages to the printer you pick in the standard Windows dialog. The
app has no print service of its own and sends nothing over the network by itself.

Some operations write short‑lived working files, and all of them are removed as soon as
the operation ends — successfully or not:

- **PDF → Word** puts any images found in the PDF into a folder under your system temp
  directory, and deletes that folder as soon as the `.docx` is saved.
- **Compression, grayscale, and repair** write the new copy beside the output file
  (`<name>.pdf.gstmp`) and keep the original as `<name>.pdf.gsbak` while they swap them,
  so a failure part‑way through cannot lose the file. Both are deleted afterwards.
- **Page rendering through Ghostscript** writes one temporary `.png` in the system temp
  directory and deletes it immediately after reading it.

## Data the app stores on your computer

The app keeps a small amount of data **locally**, under
`%APPDATA%\iwo Helper Desktop\`, and never sends it anywhere:

- `settings.txt` — your last‑used folders and options (output format, thumbnail zoom,
  compression level, interface language, and the size and position of each tool window
  and of the full‑size page preview),
- `stats.txt` — local counters of how many operations you have run (no file names,
  no content), these exist only for your own reference and can be cleared manually or
  automatically from the app’s **Statistics** window,
- `reports\` — text reports of the three most recent Excel Digest runs, saved next to
  where you chose to save the digest and mirrored here,
- `crash.log` — a local, size‑rotated error log written only if the app hits an
  unexpected error (exception type, message and stack trace, which may include the path
  of the file being processed — never its contents). It is never transmitted anywhere,
  it exists so **you** can attach it to a bug report if you choose to,
- `setup-language.txt` — one word, `ru` or `en`: the language you picked in the installer.
  The app applies it at the first start and deletes the file,
- `Инструкция пользователя.docx` — the user guide, unpacked from the program itself when you
  open it from **About**. It is a copy of what already sits inside the `.exe`, nothing is
  downloaded.

You can delete this folder at any time, the app recreates only what it needs.

## Network use

The single network feature is the **update check**. It runs when you click “Check for
updates”, and once when the app starts:

- It sends an HTTPS request to the GitHub Releases API
  (`https://api.github.com/repos/DedovMosol/iwoHelperDesktop/releases/latest`) with a
  generic `User-Agent` header and reads the latest published version tag.
- No file contents, file names, identifiers, or personal data are included in the request.
- The app never downloads or installs updates automatically, if a newer version exists,
  it asks first and only then opens the release page in **your** browser.
- The **startup** check says nothing unless a newer version exists: no network, no answer
  and “you are up to date” are all silent, because you did not ask the question. When it
  does have news, the notice carries a “Don't remind me about this version” box, and the
  check itself can be switched off — the setting is stored in `settings.txt`
  (`updateCheckOnStart`), together with the version you asked to skip (`skippedVersion`).

Apart from this check the app makes **no background network calls**.

Links in the app (download page, project page, Telegram, this policy) simply open in your
default browser when clicked, the app itself does not track those clicks.

## Third parties

- **GitHub** — contacted only for the update check and when you open a project/download
  link. See GitHub’s own privacy policy for how they handle web requests.
- **Ghostscript** — used locally as a separate process for compression, grayscale, repair
  and page rendering, it does not make network calls in this app.
- **Microsoft Excel / Word** — used locally via automation for Excel/Word features and
  are governed by your own Microsoft Office configuration.

The app is not packed or obfuscated and requests no special permissions, the installer
installs per‑user by default and needs no administrator rights.

## Children

The app is a general‑purpose office utility and is not directed at children. It collects
no personal data from anyone.

## Changes to this policy

If this policy changes, the updated version will be published in this repository with a
new “Last updated” date.

## Contact

Questions about this policy: open an issue at
<https://github.com/DedovMosol/iwoHelperDesktop> or reach the author on Telegram
([@i_wantout](https://t.me/i_wantout)).
