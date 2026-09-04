# metaclean

Find/strip invisible Unicode characters from text (e.g. text copied out of a
chat/browser), and find/strip metadata from PDFs and Office files
(docx/xlsx/pptx). Usable as a CLI, or as a drag-and-drop app on Windows —
see [Drag-and-drop use](#drag-and-drop-use) below.

## What it does

- **Text — invisible/format Unicode.** Scans a string (clipboard, file, or
  stdin) for zero-width spaces/joiners, BOMs, bidi marks, variation
  selectors, Unicode "tag" characters, and stray control characters, and can
  strip them. This is the kind of thing that rides along silently when text
  is copied out of a rendered web page (a chat UI, a browser) — invisible
  on screen, still present in what gets pasted.
- **PDF — document metadata.** Reads/strips the Document Info dictionary and
  XMP metadata (`pikepdf`) — author, creator app, producer, creation/mod
  timestamps, document/instance IDs, etc.
- **Office files (docx/xlsx/pptx) — document metadata.** Same idea, for the
  `docProps/core.xml` / `app.xml` / `custom.xml` parts inside the OOXML zip —
  author, last-modified-by, title/subject/keywords, revision number,
  company/manager, timestamps.
- **Drag-and-drop app.** `Clean Files.cmd` dispatches by file extension to
  whichever of the above applies, so there's nothing to type for routine use.

In every mode, "find" only reports (never modifies), and "clean" writes a
separate cleaned copy by default — originals are left alone unless
`--in-place` is passed explicitly.

**Scope note:** this strips *metadata*, not content. It won't remove a
watermark stamped into a page/slide as visible text or an image overlay, and
it doesn't touch tracked changes or comments in Office files — those need to
be handled in the source application first.

## Use cases

1. **Cleaning up a paste from an AI chat before it goes into a document.**
   Copy a response out of a chat interface, run `metaclean text --clean`
   (or just hit "Clean Files.cmd" after saving it to a `.txt`), and paste the
   result into Word/email/a doc without carrying along invisible characters
   picked up from the page's HTML.
2. **Scrubbing identity/authorship metadata before sharing or publishing a
   file.** A Word doc or exported PDF carries your name, network username,
   organization, and timestamps in its metadata even if none of that is
   visible on the page — Word/Adobe write it automatically. Before uploading
   a document somewhere public (a repo, a shared drive, a submission
   portal), drag it onto `Clean Files.cmd` and check the "before" with
   `metaclean pdf find` / `metaclean ooxml find` to see exactly what would
   have gone out with it.

## Setup

```
py -m pip install -r requirements.txt
```

(Text mode has no dependencies beyond the standard library — `tkinter` handles
clipboard access. `pdf` mode needs `pikepdf`. `ooxml` mode has no extra
dependencies — it edits the docx/xlsx/pptx zip directly.)

## Drag-and-drop use

**"Clean Files.cmd"** in this folder is the app-like entry point:

- **Drag one or more files straight onto its icon** (or a shortcut to it —
  right-click the file → *Send to* → *Desktop (create shortcut)* to put one
  somewhere handier). Each gets cleaned and written to a `cleaned\` folder
  next to the original; originals are left untouched.
- **Or drop files into the `inbox\` folder** next to this script, then
  double-click "Clean Files.cmd" with nothing selected. Cleaned copies land
  in `outbox\`, and the originals move into `inbox\_processed\` so re-running
  it doesn't redo them.

A console window stays open afterward (press any key to close it) so you can
read what happened to each file.

Supported by extension: `.pdf`, `.docx`/`.xlsx`/`.pptx` (and `.dotx`/`.docm`/
etc.), `.txt`/`.md`/`.csv`/`.log`. Anything else is reported as unsupported
rather than silently skipped or guessed at.

## Command-line use

Optional: add this folder to your `PATH` so `metaclean` works from anywhere:

```powershell
setx PATH "$env:PATH;$(Resolve-Path .)"   # run from this folder
```

(open a new terminal afterward for it to take effect)

## Text — invisible/format characters

Scans for zero-width spaces/joiners, BOM, bidi/variation-selector marks,
Unicode "tag" characters, and other non-printable format/control characters —
the kind of thing that can ride along invisibly when you copy text out of a
web page and paste it elsewhere.

```powershell
# Report what's on the clipboard right now, don't change it
metaclean text

# Clean the clipboard in place
metaclean text --clean

# Clean a file
metaclean text --clean --file notes.txt --out notes.clean.txt

# Pipe through it
type notes.txt | metaclean text --clean --file -
```

Exit code is `0` if nothing suspicious was found, `1` if something was (so
you can use it in a check without `--clean`).

## PDF — document metadata

```powershell
# See what's stored (Document Info dict + XMP)
metaclean pdf find "C:\path\to\file.pdf"

# Strip it (writes file.clean.pdf by default)
metaclean pdf clean "C:\path\to\file.pdf"

# Or overwrite in place
metaclean pdf clean "C:\path\to\file.pdf" --in-place
```

**Limitation:** this strips document *metadata* (author, producer, creation
date, XMP fields, etc.) only. It does not remove a watermark that's stamped
into the page content itself (visible text/image overlays) — that requires
different tooling and depends entirely on how the watermark was applied.

## Office files (docx/xlsx/pptx) — document metadata

```powershell
# See what's stored (docProps/core.xml + app.xml + custom.xml)
metaclean ooxml find "C:\path\to\file.docx"

# Strip it (writes file.clean.docx by default)
metaclean ooxml clean "C:\path\to\file.docx"

# Or overwrite in place
metaclean ooxml clean "C:\path\to\file.docx" --in-place
```

Clears author/creator, last-modified-by, title/subject/keywords, revision
number, created/modified timestamps, company/manager, etc. Document content
(the actual text, formatting, images) is untouched — only the `docProps/*.xml`
parts inside the zip are rewritten. This does not touch tracked
changes/comments if the file has any; if that matters for a given file,
accept all changes and delete comments in Word first.
