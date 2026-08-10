# -*- coding: utf-8 -*-
"""The text of the user manual — English.

Layout and building blocks live in engine.py; only the words and their order are here.
The Russian twin is body_ru.py and must have THE SAME structure — build_manual.py checks it.
"""
from engine import (CHART, bullet, footnote, front_matter, h1, h2, listed, note,
                    p, picture, ref, rich, step, table, table_ref)


def build():

    front_matter()

    # ================================================================= 1

    h1("1. About the program")

    h2("1.1. What it is for")
    par = rich("iwo Helper Desktop is a desktop program for everyday work with documents: it "
               "gathers the sheets of several Excel workbooks into one digest, merges and "
               "splits PDF files, makes them smaller, and turns the text of a PDF into an "
               "editable Word document or a PowerPoint presentation. Everything happens on "
               "your own computer; no file is sent anywhere.")
    footnote(par, "The program does not go to the internet for any operation on a document. "
                  "Its only network request is the check for a newer version, which can be "
                  "switched off in Settings.")

    p("The program is six tools, laid out on the start screen in two sections: five work with "
      "PDF, the sixth with Excel workbooks. Each does its own job and opens in its own window, "
      "so several can run at once: while one builds a digest, another can prepare a PDF.")

    table("The tools of the program",
          ["Tool", "What it does", "What for"],
          [["Merge PDF",
            "Builds one PDF out of several files, letting you pick the pages and their order",
            "To hand over one set instead of a scattering of files, without distorting the originals"],
           ["Split PDF",
            "Extracts the pages you need, or cuts a document into parts: by ranges, "
            "every N pages, or by bookmarks",
            "To take out of a thick document only what the recipient asked for"],
           ["PDF → Word",
            "Extracts the text and tables of a born-digital PDF into an editable .docx",
            "To edit text rather than type it again"],
           ["PDF → PowerPoint",
            "Turns the pages of a born-digital PDF into a .pptx presentation: the text stays "
            "editable, everything else arrives as a background",
            "To get back a presentation of which only the PDF is left"],
           ["More operations",
            "Six actions on a single document: compression, pages as images, text into .txt, "
            "grayscale, repair, document properties",
            "To do to a finished file what would otherwise send you looking for a separate program"],
           ["Merge Excel",
            "Moves the sheets of several workbooks into one — a digest, with contents and a "
            "Word cover note",
            "Not to open dozens of files by hand and copy the sheets one at a time"]],
          widths=[4.0, 6.5, 5.5])

    h2("1.2. What the program gives you")
    p("The program assumes documents are working material, not something to experiment on. "
      "That leads to a few properties worth knowing in advance.")
    bullet("No operation changes the source files. The result is always written to a new file "
           "and the original is left alone. Even if you try to save the result over the source, "
           "the program refuses and says so.",
           bold_head="Source files are never changed. ")
    bullet("A broken or password-protected file does not bring the work down: it is skipped, "
           "and the reason is shown in the list and in the report. The rest are processed.",
           bold_head="One bad file does not ruin the run. ")
    bullet("A long operation can be stopped with the Cancel button, and no half-finished file "
           "is left on disk.", bold_head="Long operations can be undone. ")
    bullet("Every window remembers its size and position, the thumbnail zoom and the chosen "
           "compression level, so there is no need to set up your workplace again at each start.",
           bold_head="The program remembers your habits. ")
    bullet("What the program does to the pages — order, rotation, removal — can be undone with "
           "Ctrl + Z, and redone with Ctrl + Y.",
           bold_head="Mistakes are cheap. ")

    h2("1.3. What the workplace needs")
    p("The program runs on Windows 8.1, 10 and 11, both 64-bit and 32-bit. No libraries need "
      "to be installed separately: everything needed is inside the program file.")
    par = rich("Two capabilities need other products. The digest itself and its cover note are "
               "produced through the Microsoft Office installed on the computer, and the "
               "conversion of PDF into Word through Microsoft Word. Compressing a PDF, turning "
               "it to grayscale and repairing a damaged file are done by the free Ghostscript "
               "engine.")
    footnote(par, "Ghostscript is part of the installer and is installed together with the "
                  "program. For the portable version it has to be installed separately — the "
                  "program offers the link the first time compression is needed.")

    h2("1.4. Installing, starting and updating")
    p("The program comes in two forms, and both are equal.")
    bullet("An ordinary installation for the current user, without administrator rights. The "
           "compression engine is included and a Start-menu shortcut is created. The wizard "
           "begins by asking for a language, showing flags, and the program takes that language "
           "at its first start — including when installed over a previous version.",
           bold_head="Installer. ")
    bullet("A single program file that only has to be copied and run. Suitable where "
           "installing anything is not possible.", bold_head="Portable version. ")
    p("The bitness follows the bitness of Windows: in the overwhelming majority of cases that "
      "is the 64-bit build; the 32-bit one is only for older systems.")
    par = p("The program looks for a newer version itself at start — quietly, in the "
            "background, and it says nothing about anything except news: no connection, no "
            "answer or “you have the latest” all pass without a word. When a version really is "
            "newer, the program names it and asks whether to open the download page in a "
            "browser. Nothing is downloaded or installed by itself — the decision stays with "
            "you. The check can be switched off in Settings, where there is also a “Check now” "
            "button.")
    footnote(par, "Settings, merge reports and the statistics counters are kept in the user "
                  "profile and carry over between versions — reinstalling does not erase them.")
    note("The program runs as a single instance. Starting the shortcut again does not open a "
         "second application but brings the running one to the front — which rules out two "
         "windows editing the same document.")

    # ================================================================= 2

    h1("2. How working with the program is arranged")

    h2("2.1. The start screen")
    p("The start screen opens when the program starts. It asks first about the section %s: "
      "“PDF” — everything done to *.pdf files, and “Other tools” — the rest. Under the name of "
      "each section is a list of what is inside, so you can choose without recalling the names "
      "of the tools." % ref("hub"))
    picture("hub", "The start screen: choosing a section")
    p("Inside a section are the same cards, one per tool %s. “Back” on the left of the header, "
      "or the Esc key, returns to the sections." % ref("hub-pdf"))
    picture("hub-pdf", "The “PDF” section: five tools")
    p("A card only has to be clicked. A quicker way is to drag PDF files straight onto a card: "
      "the tool opens and takes those files at once, skipping the file dialog. Files can also "
      "be dropped on the “PDF” section card — the program goes inside and holds them until you "
      "pick a tool. The “Merge Excel” card does not accept dropped files, because it works "
      "with a folder rather than with separate files.")
    p("Two icons sit at the bottom of the window: the gear on the left opens Settings, the "
      "question mark on the right opens About. The globe button in the top right corner "
      "switches the language. Holding the pointer over an icon shows its name in a tooltip.")
    note("The start screen can be closed without closing the tools — they keep working. “Home” "
         "on the left of a tool's coloured header brings the screen back, opening the section "
         "the tool was started from, so moving to a neighbouring tool is one click, not two.")

    h2("2.2. The language of the interface")
    p("The program speaks Russian and English. The language is chosen with the globe button on "
      "the start screen %s or with “Язык / Language” in the Menu of any tool." % ref("hub-lang"))
    picture("hub-lang", "Choosing the language of the interface")
    p("The change takes effect at once: open windows are rebuilt in the new language, staying "
      "where they were and keeping the documents they hold. The window you were working in "
      "stays active, and minimised windows stay minimised.")
    note("The language affects the interface only. The documents the program itself produces — "
         "the cover note and the merge report — are always in Russian.")

    h2("2.3. The built-in help")
    p("Every tool has “How to use” in its Menu (the F1 key). It opens a short description of "
      "the order of work for that particular window %s — the quickest way to recall the "
      "sequence without leaving what you are doing." % ref("help-split"))
    picture("help-split", "The built-in help of the “Split PDF” tool")
    p("The PDF tools have a second item as well — “Keyboard shortcuts”. It shows a crib sheet "
      "for working with the page grid %s. The set of keys depends on the tool: where the order "
      "of pages cannot be changed, those lines are absent." % ref("shortcuts"))
    picture("shortcuts", "The crib sheet for working with pages")

    table("Keys for working with the page grid",
          ["Keys", "Action"],
          [["Ctrl + mouse wheel", "Change the thumbnail zoom"],
           ["Ctrl + “+” / “−”", "Zoom the thumbnails in or out"],
           ["Ctrl + 0", "Return the zoom to 100 %"],
           ["Home / End", "Go to the first or the last page"],
           ["Ctrl + A", "Select all pages"],
           ["Ctrl + G", "Go to a page by its number"],
           ["Alt + ← / Alt + →", "Move a page left or right"],
           ["Ctrl + X / Ctrl + C", "Cut or copy the selected pages"],
           ["Ctrl + V", "Paste pages from the clipboard"],
           ["Delete", "Remove the selected pages from the result"],
           ["Esc", "Cancel a cut"],
           ["Ctrl + Z / Ctrl + Y", "Undo or redo the last action"],
           ["Ctrl + Shift + “+” / “−”", "Rotate the selected pages right or left"]],
          widths=[5.5, 10.5])

    h2("2.4. The page grid")
    p("All the PDF tools are built around the same grid of thumbnails. Learn it once and you "
      "can work with every one of them. The only difference is what the grid lets you do: in "
      "merging and in the conversions the pages are also reordered, while in splitting and in "
      "the other operations they are only selected, because the order does not change there.")
    p("Pages are shown as thumbnails in the order in which they will reach the result, with a "
      "number under each. They are selected with the mouse, with Ctrl (one at a time) and with "
      "Shift (a run). The thumbnail zoom is changed with the slider at the bottom of the "
      "window, the “%” box beside it, or Ctrl + mouse wheel. A double click on the “%” box, or "
      "Ctrl + 0, returns the natural size.")
    p("A right click on a page opens a context menu with every action on the selected pages "
      "%s: the clipboard, moving, rotating, removing and going to a page. The separate items "
      "“Select even pages” and “Select odd pages” save time with double-sided scans: one side "
      "is selected whole, rather than with fifty clicks." % ref("merge-ctx"))
    picture("merge-ctx", "The context menu of a page in the grid")
    p("“Go to page…” (Ctrl + G) is useful in thick documents: instead of scrolling, name the "
      "number and the grid scrolls to that page and selects it %s." % ref("goto"))
    picture("goto", "Going to a page by its number")
    p("Any action on the pages is undone with Ctrl + Z and redone with Ctrl + Y — that covers "
      "dragging, removing and rotating alike. It is safe to make a mistake.")
    note("The order of pages can also be changed by dragging the thumbnails. The insertion "
         "point is shown by a vertical bar, so it is hard to miss.")

    h2("2.5. Viewing a page full size")
    p("A double click on a thumbnail opens the page whole %s. That is how you make sure it is "
      "the right page: small type is illegible on a thumbnail." % ref("preview"))
    picture("preview", "Viewing a page full size")
    p("The viewer has a magnifier: the “−” and “+” buttons, a box with the current zoom, a "
      "“Fit” button and Ctrl + mouse wheel. When zooming with the wheel the point under the "
      "pointer stays put — what grows is what you are looking at. Once the page no longer fits "
      "the window it can be moved with the mouse, “taken by hand”.")
    p("A right click offers rotation right and left and a “Print” button. A rotation made here "
      "applies to the page in the grid and is undone by the same Ctrl + Z. The viewer "
      "minimises to the taskbar like any other window and closes with Esc or the cross.")

    h2("2.6. Printing")
    p("A “Print” button is in merging, splitting, both conversions and the viewer. The rule is "
      "one: the selected pages are printed, and if nothing is selected, the whole document. "
      "That spares you the usual fiddling with page numbers in the printer dialog — you have "
      "already chosen the pages with your eyes in the grid.")
    p("The page is fitted to the sheet whole, centred and in proportion. Stretching it to the "
      "edges would cut off the margins, which is where signatures and page numbers usually "
      "are. A rotation set in the grid reaches the paper too.")

    h2("2.7. Making the file smaller (compression)")
    p("Compression is offered wherever the output is a PDF — in merging, splitting and the "
      "other operations: the “Compression” list is on the bottom line of the window %s. It "
      "applies to the resulting file; the source is untouched." % ref("merge-compress"))
    picture("merge-compress", "Choosing the level of compression")
    p("The size goes down because the resolution of the pictures inside the document goes "
      "down — the same way Adobe Acrobat works. Text and vector graphics are left alone and "
      "lose no sharpness, which is why a document typed in Word gains almost nothing from "
      "compression while a scanned one gains many times over.")

    table("Levels of compression",
          ["Level", "What it does", "When to choose it"],
          [["Excellent — no compression",
            "The file is saved as it is",
            "The default. When every detail of the image matters"],
           ["Very good — smaller without losing sharpness",
            "The document is rebuilt without recompressing images",
            "When you need to reduce size while preserving picture quality"],
           ["Good — smaller size (150 dpi)",
            "Pictures are brought down to 150 dots per inch",
            "Sending by mail, publishing. Text on a scan stays legible"],
           ["Normal — minimal size (72 dpi)",
            "Pictures are brought down to 72 dots per inch",
            "When only the content has to arrive and the size is critical"]],
          widths=[5.0, 5.0, 6.0])

    p("How much that is worth is shown by a measurement on a scanned document of four pages "
      "%s. The source file took %s MB, after “Very good” — %s MB, after "
      "“Good” — %s MB, after “Normal” — %s MB. The difference between "
      "“Good” and “Normal” is barely visible to the eye, while “Very good” saves "
      "the least but leaves the images themselves untouched."
      % (ref("chart"), "%.2f" % CHART[0], "%.2f" % CHART[1], "%.2f" % CHART[2],
         "%.2f" % CHART[3]))
    picture("chart", "The size of a scanned document at different levels of compression")
    par = p("Compression changes the contents of the file, so a digital signature on the "
            "document becomes invalid. Compress before signing, not after.")
    footnote(par, "This is a property of any PDF compression, not a quirk of this program: a "
                  "signature certifies particular bytes, and rebuilding the file changes them.")

    h2("2.8. Settings")
    p("The gear at the bottom of the start screen opens Settings — what belongs to the whole "
      "program rather than to an open document %s. The same item is in the Menu of every "
      "tool." % ref("settings"))
    picture("settings", "The settings of the program")
    p("“Check for updates at startup” is the only request the program makes to the network. "
      "Only a request for the latest version number is sent; no files and no data of yours are "
      "ever sent. The check can be switched off entirely, and the “Check now” button asks on "
      "demand.")
    p("“Keep a history of operations” switches on a log: what exactly the program did, when, "
      "and which file came out. It exists so that yesterday's result can be found without "
      "recalling where it was saved: the “History” button opens the list, and from there the "
      "file itself or its folder. The “Remove entries older than” list sets how long entries "
      "live; “Clear history” erases them at once.")
    note("The history keeps only file names, paths and the time of the operation — not the "
         "contents — and it lives on your computer. A history switched off is not written at all.")
    p("The compression level and the thumbnail zoom are deliberately left out of Settings: "
      "they are changed while working, and walking to a separate window for each file would "
      "take longer than the work itself.")

    h2("2.9. Statistics")
    p("“Statistics” in the Menu shows how many operations of each kind have been done %s. It "
      "is a simple way to see how much manual work the program has already taken away."
      % ref("stats"))
    picture("stats", "The counters of completed operations")
    p("The counters can be reset with “Clear”, or set to reset themselves once a day, a week "
      "or a month. The data is kept on your computer only.")

    h2("2.10. What happens to your files")
    p("The program does not send documents anywhere, does not keep copies of them and does not "
      "log their contents. All it writes is window settings, operation counters and reports on "
      "merging Excel workbooks. A report holds the names of files and sheets, not what is in them.")
    p("The result of any operation is always a new file. If you name as the result the same "
      "file that is open in the program, the operation is refused with a plain message: "
      "writing a file into itself would destroy it.")

    # ================================================================= 3

    h1("3. Merge PDF")

    h2("3.1. What it is for")
    p("The tool builds one PDF out of several files and individual pages. Unlike plain "
      "“gluing”, here every page is visible and can be moved, rotated or thrown out before the "
      "file is written.")
    p("Its main property is that pages are copied as they are, without being converted again. "
      "A scan stays the same scan; stamps and signatures do not drift. That is why merging is "
      "safe for documents that will travel further up the chain.")

    h2("3.2. The window")
    p("On the left is the grid of pages of all the added files, on the right the actions on "
      "them, at the bottom the zoom, the compression and the save button %s." % ref("merge"))
    picture("merge", "The “Merge PDF” window")

    h2("3.3. The order of work")
    step(1, "or drag them into the window. There may be any number of files; the pages of all "
            "of them line up in one grid.",
         bold_head="Add the files with “Add PDF…” ")
    step(2, "by dragging the thumbnails or with the “◀ Move left” and “Move right ▶” buttons. "
            "Remove pages you do not need with “Remove” or the Delete key.",
         bold_head="Set the order of the pages ")
    step(3, "if the resulting file has to be smaller (see 2.7).",
         bold_head="Choose the level of compression ")
    step(4, "and give the file a name. The program offers a default name and the folder of the "
            "first added document.",
         bold_head="Press “Save PDF…” ")

    h2("3.4. Double-sided printing: “Add a blank page”")
    p("The “Add a blank page” tick box on the right is for a merged document that will be "
      "printed on both sides of the sheet. If a document has an odd number of pages, the next "
      "one starts on the back of the previous one — the program puts a blank page between them "
      "so that each document starts on the front of a sheet.")
    note("No blank page is added after the last document: nothing is printed after it anyway. "
         "Blank pages appear only in the written file and do not clutter the grid.")

    h2("3.5. Assembling a double-sided scan")
    p("“Assemble a double-sided scan (front + back)” in the Menu solves a common trouble with "
      "single-sided scanners %s. The stack is scanned face up, turned over and scanned again — "
      "which gives two files, the second of them usually in reverse order." % ref("merge-menu"))
    picture("merge-menu", "The Menu of the “Merge PDF” tool")
    p("Add both files and choose that item. The program asks whether the second document came "
      "out back to front and lays the pages out in turn: 1, 2, 3, 4… One Ctrl + Z brings the "
      "previous order back if something went wrong.")

    # ================================================================= 4

    h1("4. Split PDF")

    h2("4.1. What it is for")
    p("The tool solves two related tasks: taking the pages you need out of a document, and "
      "cutting a document into parts. Beyond that, the extra actions on an open file live "
      "here — saving pages as images, extracting text, converting to grayscale, repairing a "
      "damaged file and editing the document properties.")
    p("Its value is that the recipient gets exactly what is addressed to them. Not “a volume "
      "of three hundred pages, see section four”, but section four as a file of its own.")

    h2("4.2. The window")
    p("On the left is the grid of pages of the open document, on the right the choice of mode "
      "and its settings %s. The set of fields on the right changes with the mode, so nothing "
      "superfluous is on screen." % ref("split"))
    picture("split", "The “Split PDF” window")

    h2("4.3. The modes of splitting")
    p("The mode is chosen in the “Mode” list %s. Under the list the program says briefly what "
      "the chosen mode will do, and the caption on the big button changes from “Extract…” to "
      "“Split…” — so that what will happen is clear before it does." % ref("split-modes"))
    picture("split-modes", "Choosing the mode of splitting")

    table("The modes of splitting",
          ["Mode", "What comes out", "When it suits"],
          [["Extract selected",
            "One file from the pages selected in the grid",
            "Particular pages are needed and are easier to pick by eye"],
           ["By ranges",
            "A file per range, or one common file",
            "The page numbers are known in advance"],
           ["Every N pages",
            "Files of equal size, N pages each",
            "The document has to be broken into equal parts"],
           ["By bookmarks",
            "A file per top-level bookmark",
            "The document has bookmarks: the sections separate themselves"]],
          widths=[4.5, 5.5, 6.0])

    p("“By ranges” takes numbers and ranges — “1-3, 5, 8-” %s. For the common cases the "
      "“Ranges” field offers ready-made choices: “All pages”, “Odd pages”, “Even pages”, "
      "“Every 2nd” and “Every 3rd”. Pick one from the list and the range is filled in for you, "
      "ready to edit or to use as it is. By default every range becomes a separate file, and "
      "the “Combine into one file” tick box gathers all the listed pages into a single "
      "document." % ref("split-ranges"))
    picture("split-ranges", "The “By ranges” mode")
    p("“Every N pages” cuts the document into parts of the given size %s. A value of 1 means "
      "“every page as its own file”." % ref("split-everyn"))
    picture("split-everyn", "The “Every N pages” mode")
    p("“By bookmarks” relies on the table of contents inside the PDF: every top-level bookmark "
      "gives a separate file, and its name goes into the file name %s. If the document has no "
      "bookmarks, the program says so and writes nothing." % ref("split-bookmarks"))
    picture("split-bookmarks", "The “By bookmarks” mode")

    h2("4.4. Naming the parts")
    p("The “How to name the parts” field sets how the names of the resulting files are built "
      "from the base name. The field is optional: left empty, the parts get the same names as "
      "before — the base name plus a number or a range.")
    p("A right click on the field inserts the available placeholders %s. The folder and the "
      "base name you will give later, in the ordinary save dialog." % ref("split-template"))
    picture("split-template", "Inserting placeholders into the name template")

    table("Placeholders in the name template",
          ["Placeholder", "What is substituted", "Example"],
          [["[BASENAME]", "The base name from the save dialog", "Contract"],
           ["[FILENUMBER]", "The number of the part in order", "1, 2, 3"],
           ["[FILENUMBER###]", "The same number with leading zeros", "001, 002, 003"],
           ["[FILENUMBER10]", "Numbering with an offset", "11, 12, 13"],
           ["[CURRENTPAGE]", "The first page of the part in the source document", "5"],
           ["[TOTAL_FILES]", "How many parts came out in all", "20"],
           ["[BOOKMARK]", "The name of the bookmark (in “By bookmarks”)", "Section 1"],
           ["[TIMESTAMP]", "The date and time of the operation", "2026-07-27_10-45-03"]],
          widths=[5.0, 7.0, 4.0])

    p("Placeholders can be mixed with ordinary text. The template "
      "“[BASENAME]_part_[FILENUMBER###]”, for instance, gives “Contract_part_001”, "
      "“Contract_part_002” and so on — which sort correctly in the file manager even at a "
      "hundred parts.")
    note("Characters that a file name may not contain are replaced automatically, and long "
         "bookmark names are shortened. It is not possible to write a file with a name Windows "
         "would refuse.")

    h2("4.5. Other operations on the open document")
    p("The “More operations” button on the right opens the tool of that name (section 5) and "
      "hands it the document open here. There is no need to open the file again: compressing "
      "it, converting it to grayscale, saving its pages as images, extracting its text or "
      "editing its properties can be done at once, without leaving the work you started.")

    # ================================================================= 5

    h1("5. More operations")

    h2("5.1. What it is for")
    p("The tool is a page workshop over ONE document. First you assemble the document the way "
      "you need it: open a PDF and add images if you have any, reorder the pages, rotate them, "
      "remove the ones you do not want. Then you pick an action — save it as a PDF, make it "
      "smaller, take the colour out, repair a damaged file, pull pages out as images or as "
      "text, print, edit the properties — and it applies to the ASSEMBLED document rather "
      "than to the file on disk.")
    par = p("Every action writes a new file and leaves the source alone, so they are safe to "
            "try: if you do not like the result, you still have the original. While the pages "
            "are untouched the program simply copies the source, so ordinary work did not "
            "become any slower.")
    footnote(par, "The program will not let the result be written over the open document: such "
                  "a write would destroy the file the operation itself is reading.")

    h2("5.2. The window")
    p("On the left is the grid of assembled pages, on the right are first the two ways to "
      "collect them (“Open PDF…” and “Add images…”) and below them the action buttons, "
      "grouped by meaning: “Convert the document”, “Extract from the document” and "
      "“Edit” %s." % ref("ops"))
    picture("ops", "The “More operations” window")
    p("A document reaches this window three ways: with the “Open PDF…” button, by being "
      "dragged into the window, and by coming across from a neighbouring tool — “Split PDF” "
      "hands over its open document with its own “More operations” button, and “Merge PDF” "
      "hands over the file it has just built through “More operations” in its Menu. Images can "
      "simply be dragged in as well — onto the window, onto the grid, or onto the tool’s card "
      "on the start screen.")
    note("While the grid holds no pages the action buttons are greyed out: it is immediately "
         "clear that a document has to be assembled first, and there is no need to find that "
         "out by clicking.")

    h2("5.3. Assembling the document")
    p("The grid here is edited just as it is in merging: pages are reordered by dragging, "
      "rotated with the buttons on the tile, removed with Delete, moved through the clipboard "
      "(Ctrl+X, Ctrl+C, Ctrl+V), and Ctrl+Z takes any of those edits back.")
    p("What is assembled is what every action in the window works on. While the grid differs "
      "from the file, the status line says how many pages are assembled in it — so that the "
      "result does not come as a surprise.")
    bullet("Photos and scans (JPEG, PNG, BMP, GIF, TIFF, multi-page TIFF included) are appended "
           "to the set as pages: each image becomes an A4 sheet with margins, fitted whole and "
           "without distorting its proportions, and the sheet takes the orientation of the "
           "image %s. From there it is an ordinary page: rotate it, reorder it, print it, "
           "compress it." % ref("ops-images"),
           bold_head="Add images. ")
    picture("ops-images", "An image added as a page next to the pages of a document")
    note("A phone photo arrives the way it was taken: the program reads the orientation tag out "
         "of the picture itself. A JPEG is carried into the PDF as it is, without re-encoding, "
         "so no quality is lost and the file does not swell. Transparent areas of a PNG become "
         "white.")

    h2("5.4. The actions")
    p("The “Convert” button opens a menu of operations on the assembled document:")
    bullet("Writes what is assembled into a new PDF: that is how images become a document and "
           "reordered or rotated pages become a finished file. The compression level comes from "
           "the same list at the bottom of the window.", bold_head="Save PDF. ")
    bullet("Makes a finished file smaller — with the same levels as merging (see 2.7). If the "
           "file is already optimised, the program says so rather than silently handing back "
           "the same file.", bold_head="Compress. ")
    bullet("A colour document becomes black and white. That makes the file noticeably smaller "
           "and printing on a monochrome printer predictable.",
           bold_head="To grayscale. ")
    bullet("The file is rebuilt by the PDF engine, which repairs a broken cross-reference "
           "table — the commonest cause of the message “the file is damaged”. The file is "
           "chosen in a separate dialog, because a damaged document cannot be opened into the "
           "grid — and that is exactly when repair is needed.",
           bold_head="Repair. ")
    p("The “Extract” button opens a menu for extracting content:")
    bullet("PNG or JPEG with a choice of resolution %s. Needed when a page has to go into a "
           "presentation, a letter or a memo. Available resolutions are 96, 150, 300 and 600 dots "
           "per inch. The resolution applies to the real size of the sheet, so a landscape "
           "insertion is not stretched. The selected pages are saved, or all of them if nothing "
           "is selected." % ref("ops-dpi"),
           bold_head="Pages as images. ")
    picture("ops-dpi", "Choosing the resolution when saving pages as images")
    bullet("The text layer of the document is saved into a .txt file. Tables are not lost: "
           "cells are separated by tabs and paste into Excel as a table. Pages are separated by "
           "a page-break character, and the file is written in UTF-8.",
           bold_head="Text into .txt. ")
    p("Separate buttons:")
    bullet("The selected pages (or all of them if nothing is selected) go to the printer "
           "together with the rotations assigned to them.", bold_head="Print. ")
    bullet("Title, author, subject and keywords — what any viewer shows in the properties of a "
           "file %s. An empty field clears the property: that is how an author's name is taken "
           "out of a document before it is sent." % ref("metadata"),
           bold_head="Document properties. ")
    picture("metadata", "Editing the properties of a document")
    note("Compression, conversion to grayscale and repair are done by the Ghostscript engine. "
         "In the installed version it is already there; for the portable one the program offers "
         "a link to download it.")

    # ================================================================= 6

    h1("6. PDF → Word")

    h2("6.1. What it is for")
    p("The tool extracts the text and tables of a born-digital PDF into an editable Word "
      "document (.docx). Born-digital means saved from Word, from a browser or through "
      "“Microsoft Print to PDF”: in such a file the text is held as text, not as a picture.")
    p("The point is simple: text somebody has already typed does not have to be typed again. "
      "From the resulting .docx you can take a paragraph, a table or the whole document and "
      "work with it as usual.")

    h2("6.2. The window")
    p("The window is arranged like “Merge PDF”: on the left the grid of pages of all the added "
      "files, on the right the actions, at the bottom the zoom and the “Convert to Word…” "
      "button %s." % ref("ocr"))
    picture("ocr", "The “PDF → Word” window")

    h2("6.3. The order of work")
    step(1, "or drag them into the window. The pages of all the files are shown in one grid, "
            "so one Word document can be assembled from several PDFs.",
         bold_head="Add one or more PDFs with “Add PDF…” ")
    step(2, "by dragging the thumbnails or with the “◀ Move left” and “Move right ▶” buttons. "
            "Remove pages you do not need with “Remove”. The pages reach Word in exactly the "
            "order shown.",
         bold_head="Change the order of the pages if you need to ")
    step(3, "If a page lies on its side, turn it before converting: the layout is analysed in "
            "the position you see, and sideways text will not become lines otherwise.",
         bold_head="Straighten rotated pages. ")
    step(4, "and give the file a name. When it is done the document opens by itself.",
         bold_head="Press “Convert to Word…” ")

    h2("6.4. What comes across and what does not")
    p("What comes across: text in paragraphs in reading order — with its font, size, weight, "
      "colour, underline, alignment and first-line indent. Tables with visible lines are "
      "rebuilt as cells, merged ones included. Portrait and landscape orientation is kept page "
      "by page. Images and hyperlinks come across too.")
    p("The limits are worth knowing in advance, so as not to waste time.")
    bullet("pages that are images with no text layer. The program says so, and the file is not "
           "harmed.", bold_head="Scanned documents are not supported — ")
    bullet("the text is set in Times New Roman. The shape of the letters may differ a little "
           "from the original.", bold_head="If the font of the PDF is not in the system, ")
    bullet("boxed insets and text in several columns come across as plain paragraphs in one "
           "column — they may have to be tidied by hand.",
           bold_head="Tables without lines, ")
    par = p("One rare case is worth mentioning separately: if the PDF was saved with a broken "
            "text encoding, what is extracted will be unreadable. That is a defect of the file "
            "itself, not of the conversion.")
    footnote(par, "It is easy to check: open the PDF in any viewer, select the text and copy "
                  "it (Ctrl + C). If nonsense lands in the clipboard, the text layer of the "
                  "document was broken to begin with.")

    # ================================================================= 7

    h1("7. PDF → PowerPoint")

    h2("7.1. What it is for")
    p("The tool turns the pages of a born-digital PDF into a PowerPoint presentation (.pptx). "
      "Every page becomes a slide, and the text on it stays TEXT: it can be selected, "
      "corrected, retyped — it is not a picture of letters.")
    p("The situation this exists for comes up constantly: the presentation went to its "
      "recipient as a PDF, the source file is lost, and it has to be edited. The usual way is "
      "to retype the slides; this tool brings them back in one action.")
    par = p("PowerPoint itself is not needed for it: the program assembles the .pptx on its "
            "own. The tool is used precisely when there is no office suite at hand, and "
            "requiring one to be installed in order to put text into slides would be odd.")
    footnote(par, "A .pptx file is a set of XML files in an archive, described by an open "
                  "standard. The program writes them directly.")

    h2("7.2. How a slide is built")
    p("A slide is made of two layers. The lower one is the page of the source PDF, rendered "
      "WITHOUT its text layer: background, frames, charts, logos, stamps. The upper one is the "
      "text boxes, placed exactly where the text stood in the original.")
    p("That division was not chosen out of laziness. On an ordinary slide most of the image is "
      "not text, and a text-only conversion would leave a dozen labels on a white sheet. A "
      "whole-page-as-a-picture conversion would give a handsome but dead slide in which not a "
      "letter could be corrected. Two layers give both the look and the editing.")
    note("The background is rendered by Ghostscript. Without it the presentation is assembled "
         "from the text and the images alone — silently, without an error: it is still a "
         "working file, only without the background.")

    h2("7.3. The window")
    p("The window is the same as “PDF → Word”: on the left the grid of pages, on the right the "
      "actions, at the bottom the zoom and the “Convert to PowerPoint…” button %s. Every habit "
      "of the grid — selecting, ordering, rotating, previewing, undoing — works the same."
      % ref("pptx"))
    picture("pptx", "The “PDF → PowerPoint” window")

    h2("7.4. The order of work")
    step(1, "or drag them into the window. The pages of several files are gathered into one "
            "presentation.",
         bold_head="Add PDFs with “Add PDF…” ")
    step(2, "by dragging or with the “◀ Move left” and “Move right ▶” buttons; remove the ones "
            "you do not need with “Remove”. The slides follow exactly the order shown.",
         bold_head="Arrange the pages ")
    step(3, "Straighten a rotated page before converting — the layout is read in the position "
            "you see.",
         bold_head="Check the orientation. ")
    step(4, "and give the file a name. The finished presentation opens by itself.",
         bold_head="Press “Convert to PowerPoint…” ")

    h2("7.5. What comes across")
    p("Text comes across in paragraphs, with its font, size, weight, colour, underline, "
      "super- and subscripts and hyperlinks. Tables become real slide tables, merged cells "
      "included. The slide size follows the page: a 16:9 deck stays 16:9, and an A4 document "
      "gives A4 slides.")
    p("The boundaries of the paragraphs are taken from the file itself where it carries them: "
      "a PDF saved from Word, PowerPoint or Acrobat holds the structure of the document, and "
      "in it the author has already marked where one paragraph ends and the next begins. Then "
      "a paragraph arrives as one text box and is edited as a whole. Where there is no such "
      "structure, the boundaries are recovered from the distances between the lines.")

    h2("7.6. What not to expect")
    bullet("pages saved as a picture in their entirety. Such slides arrive as a background "
           "without a single letter, and the program says honestly how many there were: only "
           "recognition can bring the text back from those.",
           bold_head="There will be no text where the source has none — ")
    bullet("animation, transitions, speaker notes and the links from charts to their source "
           "data are not kept at all — there is nowhere to take them from.",
           bold_head="What is not in the PDF does not come across: ")
    bullet("lines justified by stretching the spaces arrive with ordinary spaces: reproducing "
           "the stretch would mean laying the line out in fragments, and then the paragraph "
           "would stop being editable as a whole.",
           bold_head="Small differences in setting are possible: ")

    # ================================================================= 8

    h1("8. Merge Excel")

    h2("8.1. What it is for")
    p("The tool gathers the sheets of several Excel workbooks into one workbook — a digest. "
      "The typical task: the divisions have each sent a file, and what is wanted is one "
      "document where each file is a separate sheet, there are contents with links, and it is "
      "clear which files could not be processed and why.")
    p("Assembling such a digest by hand means opening every workbook, copying a sheet, "
      "pasting it, renaming it, and round again. The program does it in one pass, keeping the "
      "formatting, the formulas and the charts.")

    h2("8.2. The window")
    p("The window is divided into blocks by meaning: where to take from, where to save, with "
      "what options and what exactly to merge %s." % ref("excel"))
    picture("excel", "The “Merge Excel” window")

    h2("8.3. The order of work")
    step(1, "or drag it straight into the window. The program counts the files found at once "
            "and shows them in the list below.",
         bold_head="Give the folder of the source files with “Browse…” ")
    step(2, "The extension is added automatically, so there is no need to type it. The “.xls” "
            "format is left for those who need Excel 97–2003. If the “Folder” field is left "
            "empty, the digest is put beside the source files.",
         bold_head="Give the resulting file a name. ")
    step(3, "The order is set by dragging the rows or with the “▲ Up” and “▼ Down” buttons; "
            "“By name” restores the natural order. A file is excluded by clearing its tick, "
            "and “Check all” and “Uncheck all” do it wholesale.",
         bold_head="Check the contents and the order in “Files to merge”. ")
    step(4, "Progress is shown in the status line and the progress bar, and “Cancel” stops the "
            "run after the current file.",
         bold_head="Press “Merge”. ")

    h2("8.4. The options")
    bullet("“First sheet only” takes the first visible sheet of each workbook — the usual case "
           "when a file holds a single form. “All sheets” moves every visible sheet.",
           bold_head="Sheets. ")
    bullet("The first sheet of the digest becomes its contents: the names of the files, "
           "hyperlinks to the corresponding sheets and the status of each file. The digest "
           "becomes something to move around in, rather than fifty tabs to search through.",
           bold_head="The “Contents” sheet. ")
    bullet("Calculated values reach the digest instead of formulas. The digest stops depending "
           "on the source workbooks: it can be sent on, and its numbers will not drift or turn "
           "into errors for want of a reference.",
           bold_head="Replace formulas with values. ")

    h2("8.5. The result, the report and the cover note")
    p("After the run the outcome for each file appears in the list itself — in the “Result” "
      "and “Note” columns. Broken and password-protected workbooks are skipped with the reason "
      "given, and the rest are moved across. Links appear at the bottom of the window: open "
      "the finished file, open its folder, open the merge report.")
    p("“Retry skipped” adds the corrected files into the digest already assembled, without "
      "rebuilding it whole. That saves time when two files out of thirty failed.")
    par = p("“Word note” produces a cover note for the digest as a .docx: the totals, the list "
            "of skipped files and standard formatting. The note only has to be signed, not "
            "written from nothing.")
    footnote(par, "Merge reports are kept in the user profile; the three most recent remain. "
                  "Their folder can be opened through Menu → “Reports folder”.")

    h2("8.6. The Menu of the tool")
    p("The Menu holds the built-in help, the statistics, the choice of language and “Reports "
      "folder”, which opens the folder with the most recent merge reports %s." % ref("excel-menu"))
    picture("excel-menu", "The Menu of the “Merge Excel” tool")
    note("The usual keys work in the file list: Alt + ↑ and Alt + ↓ change the order, Delete "
         "excludes a file, Ctrl + A selects everything, Ctrl + C copies the rows to the clipboard.")

    # ================================================================= 9

    h1("9. Reference")

    h2("9.1. About the program")
    p("The About window is opened by the button on the start screen %s. It gives the version, "
      "the author, links to the project page and to the privacy policy, and the licence."
      % ref("about"))
    picture("about", "The About window")
    p("The version number is useful when asking for help: it says at once which build is meant. "
      "The details for supporting the project voluntarily can be selected and copied as "
      "ordinary text.")
    par = p("The line “User manual: open” opens this document. It lives inside the program "
            "itself, so it is available both in the portable version and without the internet: "
            "at the first request the program unpacks it beside its settings and opens it.")
    footnote(par, "The file is unpacked into the user profile, beside the settings of the "
                  "program. It may be deleted — it will appear again at the next request.")

    h2("9.2. If something went wrong")
    p("The program tries to explain failures in words rather than in error codes. Below are "
      "the situations that come up more often than others.")

    table("Common difficulties and what to do",
          ["What you see", "What it means and what to do"],
          [["The message “Ghostscript is needed” when choosing compression",
            "The compression engine is not installed. The window has a link to download it, "
            "and in the installed version of the program it is put in place at once"],
           ["A file was skipped while building the digest",
            "The workbook is damaged or password-protected. The reason is in the “Note” column "
            "and in the report. Fix the file and press “Retry skipped”"],
           ["A message that the document is scanned",
            "The PDF has no text layer, there is nothing to convert. The file is not damaged"],
           ["The text after conversion to Word is unreadable",
            "The source PDF has a broken encoding. Check whether the text copies out of the "
            "PDF itself: if it does not, the defect is in the source file"],
           ["The program refuses to save the result",
            "You named the same file that is open in the program, or a folder there is no "
            "right to write to. Choose another name or another folder"],
           ["A signature in the PDF became invalid",
            "The document was compressed or rebuilt after it was signed. Sign the finished "
            "file as the last action"],
           ["The window does not appear when the program is started again",
            "The program is already running: it brings the existing window to the front "
            "instead of opening a second one"],
           ["“Repair” or “To grayscale” answer with a refusal",
            "The document is password-protected: the engine cannot read it. Remove the "
            "protection in the program that applied it and try again"],
           ["The download page did not open after the message about a new version",
            "There is no default browser in the system. The program shows the address itself — "
            "it can be typed by hand"]],
          widths=[6.0, 10.0])

    h2("9.3. A short reminder")
    p("Three rules that answer most of the questions.")
    bullet("The program never changes the source files. The result is always a new file.")
    bullet("Anything done to the pages is undone with Ctrl + Z.")
    bullet("If it is not clear what a button does, hold the pointer over it: the tooltip "
           "explains what it is for and names the shortcut.")
