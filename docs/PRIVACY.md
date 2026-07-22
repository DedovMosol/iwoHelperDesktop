# Privacy Policy

**Application:** iwo Helper Desktop
**Last updated:** 2026-07-22

## Summary

iwo Helper Desktop is an **offline desktop application**. It does **not** collect,
transmit, sell, or share any personal data, and it does **not** contain any telemetry,
analytics, advertising, or user accounts. Everything you do with your files happens
**locally on your own computer**. The only time the app talks to the network is when
**you** click “Check for updates,” and even then it only reads a version number from
GitHub — no file contents or personal data are ever sent.

## What the app does with your files

The tools (Excel Digest, PDF Merge, PDF Split, PDF Compression, PDF → Word) read the
files you choose and write results to the folders you choose. This processing runs
entirely on your machine using local components:

- Microsoft Excel / Word automation (COM) for the Excel Digest and its Word cover note,
  and for writing the `.docx` in PDF → Word,
- the embedded PdfSharp library for reading and writing PDFs,
- the embedded PdfPig library for extracting the text layer of born‑digital PDFs (PDF → Word),
- Ghostscript, run as a separate local process, for optional PDF compression,
- the built‑in Windows PDF engine (`Windows.Data.Pdf`) for rendering page thumbnails.

**Your documents are never uploaded, copied off your device, or transmitted anywhere.**
The app does not modify your source files except where you explicitly ask it to write
output; PDF split and PDF → Word never change the source. When converting a PDF to Word,
any images found in the PDF are written to a temporary folder under your system temp
directory and deleted as soon as the `.docx` is saved.

## Data the app stores on your computer

The app keeps a small amount of data **locally**, under
`%APPDATA%\iwo Helper Desktop\`, and never sends it anywhere:

- `settings.txt` — your last‑used folders and options (e.g. output format),
- `stats.txt` — local counters of how many operations you have run (no file names,
  no content), these exist only for your own reference and can be cleared manually or
  automatically from the app’s **Statistics** window,
- `reports\` — text reports of the three most recent Excel Digest runs, saved next to
  where you chose to save the digest and mirrored here.

You can delete this folder at any time, the app recreates only what it needs.

## Network use

The app makes **no background network calls**. The single network feature is the
**update check**, which runs **only when you start it** (“Check for updates”):

- It sends an HTTPS request to the GitHub Releases API
  (`https://api.github.com/repos/DedovMosol/iwoHelperDesktop/releases/latest`) with a
  generic `User-Agent` header and reads the latest published version tag.
- No file contents, file names, identifiers, or personal data are included in the request.
- The app never downloads or installs updates automatically, if a newer version exists,
  it offers to open the release page in **your** browser.

Links in the app (download page, project page, Telegram, this policy) simply open in your
default browser when clicked, the app itself does not track those clicks.

## Third parties

- **GitHub** — contacted only for the update check and when you open a project/download
  link. See GitHub’s own privacy policy for how they handle web requests.
- **Ghostscript** — used locally as a separate process for compression, it does not make
  network calls in this app.
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
