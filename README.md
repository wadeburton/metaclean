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
- **PDF — watermarks (best-effort).** Finds/strips `/Watermark` annotations,
  watermark-named optional-content layers, and the same image/vector overlay
  drawn on every page — all structurally distinct objects, so removal is
  exact with no visual side effects. Also has a heuristic scan for text drawn
  directly into the page (repeated across pages, rotated and/or translucent)
  — opt-in to remove (`--aggressive`) since it's pattern-matching, not exact.
- **docx header watermarks (best-effort).** Detects and removes the VML
  shape Word's own "Insert Watermark" feature places in the header — reliable
  because Word's output is consistent (`PowerPlusWaterMarkObject` shape id,
  or a WordArt text-path / absolutely-positioned image).
- **Drag-and-drop app.** `Clean Files.cmd` dispatches by file extension to
  whichever of the above applies, so there's nothing to type for routine use.

In every mode, "find" only reports (never modifies), and "clean" writes a
separate cleaned copy by default — originals are left alone unless
`--in-place` is passed explicitly.

**Scope note:** watermark removal is fundamentally best-effort — see
[Watermarks: what this can and can't do](#watermarks-what-this-can-and-cant-do)
below before relying on it. This tool also doesn't touch tracked
changes/comments in Office files — handle those in the source application
first.

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

Two entry points, both work the same way (drag files onto the icon, or drop
into `inbox\` and double-click with nothing selected):

- **"Clean Files.cmd"** — metadata only. Always safe, no false-positive risk.
- **"Clean Files (+ Watermarks).cmd"** — metadata, plus best-effort watermark
  removal (PDF and docx). See the watermarks section below before trusting
  this one blindly.

Right-click either → *Send to* → *Desktop (create shortcut)* to put one
somewhere handier. Each dropped file is cleaned into a `cleaned\` folder next
to the original (originals untouched); inbox-mode output lands in `outbox\`
and processed originals move to `inbox\_processed\`.

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

This strips document *metadata* (author, producer, creation date, XMP
fields, etc.) only — not the page content. For watermarks, see below.

## PDF — watermarks

```powershell
# Report what's found, remove nothing
metaclean pdf watermark-find "C:\path\to\file.pdf"

# Strip the high-confidence cases (annotations, layers, repeated overlays)
metaclean pdf watermark-clean "C:\path\to\file.pdf"

# Also attempt the heuristic repeated-rotated/translucent-text case
metaclean pdf watermark-clean "C:\path\to\file.pdf" --aggressive
```

See [Watermarks: what this can and can't do](#watermarks-what-this-can-and-cant-do).

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

## docx — header watermarks

```powershell
# Report what's found, remove nothing
metaclean ooxml watermark-find "C:\path\to\file.docx"

# Remove it
metaclean ooxml watermark-clean "C:\path\to\file.docx"
```

Covers Word's built-in "Insert Watermark" feature (text or picture), which
always lands in a header part with a recognizable shape signature. Doesn't
cover: xlsx/pptx watermarks (not standardized the same way), or a watermark
someone built by hand out of an unusual shape/text box.

## Watermarks: what this can and can't do

Watermarks are applied in structurally different ways, and only some of them
can be undone cleanly:

| How the watermark was applied | Removable? | Notes |
|---|---|---|
| PDF `/Watermark` annotation (Acrobat's "Add Watermark" tool) | Yes, exactly | Distinct object, deleted outright |
| PDF optional-content layer named like "Watermark" | Yes, exactly | Layer + the content region referencing it |
| PDF: same image/vector overlay drawn on every page | Yes, exactly | Draw instruction removed; page content otherwise untouched |
| PDF: text drawn directly into each page (rotated/translucent, repeated) | Best-effort | Heuristic pattern match; opt-in via `--aggressive` since a repeated running header could theoretically match too |
| docx header watermark (Word's Insert Watermark feature) | Yes, exactly | Recognized by Word's own consistent shape signature |
| Watermark baked into a flattened/scanned page image | **No** | It's pixels, not a separate object — this tool can't detect or remove it, and nothing short of image inpainting could |

`watermark-find` always reports what it found *before* you strip anything,
so you can see which category you're dealing with. If `watermark-find`
comes back empty, that does not prove the file has no watermark — it means
none of the above patterns matched, which is exactly the "baked into a scan"
case this tool is upfront about not handling.
