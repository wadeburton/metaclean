#!/usr/bin/env python3
"""Find/remove invisible Unicode characters (text), and metadata from PDF and
Office (docx/xlsx/pptx) files.

Usage:
    metaclean text  [--clean] [--file PATH | -] [--to-clipboard] [--out PATH]
    metaclean pdf   find|clean   PATH [--out PATH | --in-place]
    metaclean ooxml find|clean   PATH [--out PATH | --in-place]
    metaclean auto  [PATH ...]

`auto` is the drag-and-drop entry point (see Clean Files.cmd): dispatches
each path to the right cleaner by extension. With no paths given, it
processes everything sitting in inbox/ next to this script and writes
results to outbox/.

With no --file/stdin, `text` mode reads from and (with --clean --to-clipboard)
writes back to the Windows clipboard.
"""
from __future__ import annotations

import argparse
import shutil
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SCRIPT_DIR = Path(__file__).resolve().parent

# --- text mode -------------------------------------------------------------

# Characters worth flagging even though some (soft hyphen, word joiner) have
# legitimate typographic uses. Flag everything; --clean strips all of it.
_NAMED_SUSPECTS = {
    0x00AD: "SOFT HYPHEN",
    0x180E: "MONGOLIAN VOWEL SEPARATOR",
    0x200B: "ZERO WIDTH SPACE",
    0x200C: "ZERO WIDTH NON-JOINER",
    0x200D: "ZERO WIDTH JOINER",
    0x200E: "LEFT-TO-RIGHT MARK",
    0x200F: "RIGHT-TO-LEFT MARK",
    0x2060: "WORD JOINER",
    0x2061: "FUNCTION APPLICATION",
    0x2062: "INVISIBLE TIMES",
    0x2063: "INVISIBLE SEPARATOR",
    0x2064: "INVISIBLE PLUS",
    0xFEFF: "ZERO WIDTH NO-BREAK SPACE (BOM)",
    0xFFF9: "INTERLINEAR ANNOTATION ANCHOR",
    0xFFFA: "INTERLINEAR ANNOTATION SEPARATOR",
    0xFFFB: "INTERLINEAR ANNOTATION TERMINATOR",
}


def _is_suspect(cp: int) -> bool:
    if cp in _NAMED_SUSPECTS:
        return True
    # Variation selectors (incl. supplementary block) and Unicode "tag"
    # characters — the block used by several invisible-watermarking schemes.
    if 0xFE00 <= cp <= 0xFE0F or 0xE0100 <= cp <= 0xE01EF:
        return True
    if 0xE0000 <= cp <= 0xE007F:
        return True
    # General category Cf (format) or Cc (control, excluding common
    # whitespace) anywhere in the string.
    ch = chr(cp)
    cat = unicodedata.category(ch)
    if cat == "Cf":
        return True
    if cat == "Cc" and ch not in ("\n", "\r", "\t"):
        return True
    return False


def _char_label(cp: int) -> str:
    if cp in _NAMED_SUSPECTS:
        return _NAMED_SUSPECTS[cp]
    try:
        return unicodedata.name(chr(cp))
    except ValueError:
        return "UNNAMED"


def find_suspects(text: str) -> list[tuple[int, int]]:
    return [(i, ord(ch)) for i, ch in enumerate(text) if _is_suspect(ord(ch))]


def strip_suspects(text: str) -> str:
    return "".join(ch for ch in text if not _is_suspect(ord(ch)))


def _report_hits(hits: list[tuple[int, int]], source_desc: str) -> None:
    if not hits:
        print(f"No suspicious invisible/format characters found in {source_desc}.", file=sys.stderr)
        return
    print(f"Found {len(hits)} suspicious character(s) in {source_desc}:", file=sys.stderr)
    seen: dict[int, int] = {}
    for _, cp in hits:
        seen[cp] = seen.get(cp, 0) + 1
    for cp, count in sorted(seen.items(), key=lambda kv: -kv[1]):
        print(f"  U+{cp:04X}  {_char_label(cp):<40}  x{count}", file=sys.stderr)


def _get_clipboard() -> str:
    import tkinter

    root = tkinter.Tk()
    root.withdraw()
    try:
        return root.clipboard_get()
    except tkinter.TclError:
        return ""
    finally:
        root.destroy()


def _set_clipboard(text: str) -> None:
    import tkinter

    root = tkinter.Tk()
    root.withdraw()
    try:
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()  # keep the clipboard populated after the process exits
    finally:
        root.destroy()


def cmd_text(args: argparse.Namespace) -> int:
    if args.file == "-":
        source = sys.stdin.read()
        source_desc = "stdin"
    elif args.file:
        with open(args.file, "r", encoding="utf-8") as fh:
            source = fh.read()
        source_desc = args.file
    else:
        source = _get_clipboard()
        source_desc = "clipboard"

    hits = find_suspects(source)
    _report_hits(hits, source_desc)

    if not args.clean:
        return 1 if hits else 0

    cleaned = strip_suspects(source)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(cleaned)
        print(f"Cleaned text written to {args.out}", file=sys.stderr)
    if args.to_clipboard or (not args.out and not args.file):
        _set_clipboard(cleaned)
        print("Cleaned text copied back to clipboard.", file=sys.stderr)
    if not args.out and not args.to_clipboard and args.file:
        sys.stdout.write(cleaned)

    return 0


# --- pdf mode ----------------------------------------------------------------

def get_pdf_metadata(path: str) -> tuple[dict, dict]:
    import pikepdf

    with pikepdf.open(path) as pdf:
        docinfo = {str(k): str(v) for k, v in (pdf.docinfo or {}).items()}
        with pdf.open_metadata() as meta:
            xmp = {str(k): str(meta[k]) for k in meta}
    return docinfo, xmp


def strip_pdf_metadata(in_path: str, out_path: str) -> None:
    import pikepdf

    with pikepdf.open(in_path) as pdf:
        with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
            meta.clear()
        for key in list(pdf.docinfo.keys()):
            del pdf.docinfo[key]
        pdf.save(out_path)


def cmd_pdf_find(args: argparse.Namespace) -> int:
    docinfo, xmp = get_pdf_metadata(args.path)
    print(f"Document Info dictionary ({args.path}):")
    _print_kv(docinfo)
    print("\nXMP metadata:")
    _print_kv(xmp)
    return 0


def cmd_pdf_clean(args: argparse.Namespace) -> int:
    out_path = args.path if args.in_place else (args.out or _default_clean_name(args.path, ".pdf"))
    strip_pdf_metadata(args.path, out_path)
    print(f"Metadata stripped -> {out_path}", file=sys.stderr)
    print(
        "Note: this removes document metadata only. A watermark stamped into the "
        "page content/images is not touched by this tool.",
        file=sys.stderr,
    )
    return 0


def _print_kv(d: dict) -> None:
    if not d:
        print("  (none)")
    else:
        for key, value in d.items():
            print(f"  {key}: {value}")


def _default_clean_name(path: str, ext: str) -> str:
    p = Path(path)
    if p.suffix.lower() == ext:
        return str(p.with_name(p.stem + ".clean" + ext))
    return path + ".clean" + ext


# --- ooxml (docx/xlsx/pptx) mode --------------------------------------------

_OOXML_METADATA_PARTS = ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml")

OOXML_EXTENSIONS = {
    ".docx", ".dotx", ".docm",
    ".xlsx", ".xltx", ".xlsm",
    ".pptx", ".potx", ".pptm",
}


def _is_ooxml(path: str) -> bool:
    try:
        with zipfile.ZipFile(path) as zf:
            return "[Content_Types].xml" in zf.namelist()
    except zipfile.BadZipFile:
        return False


def get_ooxml_metadata(path: str) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        for part in _OOXML_METADATA_PARTS:
            if part not in names:
                continue
            root = ET.fromstring(zf.read(part))
            fields = {}
            for child in root:
                tag = child.tag.split("}", 1)[-1]  # strip namespace
                fields[tag] = (child.text or "").strip()
            result[part] = fields
    return result


def strip_ooxml_metadata(in_path: str, out_path: str) -> None:
    with zipfile.ZipFile(in_path) as zin:
        names = set(zin.namelist())
        tmp_out = out_path + ".tmp"
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in _OOXML_METADATA_PARTS and item.filename in names:
                    root = ET.fromstring(data)
                    for child in list(root):
                        root.remove(child)
                    data = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
                zout.writestr(item, data)
    Path(tmp_out).replace(out_path)


def cmd_ooxml_find(args: argparse.Namespace) -> int:
    parts = get_ooxml_metadata(args.path)
    if not parts:
        print("No docProps metadata parts found.")
        return 0
    for part, fields in parts.items():
        print(f"{part}:")
        _print_kv(fields)
        print()
    return 0


def cmd_ooxml_clean(args: argparse.Namespace) -> int:
    p = Path(args.path)
    out_path = args.path if args.in_place else (args.out or _default_clean_name(args.path, p.suffix.lower()))
    strip_ooxml_metadata(args.path, out_path)
    print(f"Metadata stripped -> {out_path}", file=sys.stderr)
    return 0


# --- auto (drag-and-drop) mode ---------------------------------------------

TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".log"}


def _clean_one_file(src: Path, dest: Path) -> str:
    """Clean src into dest (dest's parent is created if needed). Returns a
    short status string describing what happened."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    ext = src.suffix.lower()

    if ext == ".pdf":
        strip_pdf_metadata(str(src), str(dest))
        return "PDF metadata stripped"

    if ext in OOXML_EXTENSIONS or (ext not in TEXT_EXTENSIONS and _is_ooxml(str(src))):
        strip_ooxml_metadata(str(src), str(dest))
        return "Office document metadata stripped"

    if ext in TEXT_EXTENSIONS:
        content = src.read_text(encoding="utf-8", errors="replace")
        hits = find_suspects(content)
        dest.write_text(strip_suspects(content), encoding="utf-8")
        return f"text cleaned ({len(hits)} suspicious char(s) removed)" if hits else "text OK, no suspicious chars"

    raise ValueError(f"no cleaner for {ext or '(no extension)'} files")


def cmd_auto(args: argparse.Namespace) -> int:
    if args.paths:
        jobs = [(Path(p), Path(p).parent / "cleaned" / Path(p).name) for p in args.paths]
        mode = "direct"
    else:
        inbox = SCRIPT_DIR / "inbox"
        outbox = SCRIPT_DIR / "outbox"
        processed = inbox / "_processed"
        inbox.mkdir(exist_ok=True)
        outbox.mkdir(exist_ok=True)
        files = [f for f in inbox.iterdir() if f.is_file() and not f.name.startswith(".")]
        if not files:
            print(f"No files were dropped, and {inbox} is empty.")
            print("Drag files directly onto Clean Files.cmd, or drop them into inbox\\ and run it again.")
            return 0
        jobs = [(f, outbox / f.name) for f in files]
        mode = "inbox"

    ok, failed = 0, 0
    for src, dest in jobs:
        try:
            status = _clean_one_file(src, dest)
            print(f"[ok]   {src.name}: {status} -> {dest}")
            ok += 1
            if mode == "inbox":
                processed_path = SCRIPT_DIR / "inbox" / "_processed" / src.name
                processed_path.parent.mkdir(exist_ok=True)
                shutil.move(str(src), str(processed_path))
        except Exception as exc:  # noqa: BLE001 - report and keep going
            print(f"[FAIL] {src.name}: {exc}")
            failed += 1

    print(f"\n{ok} cleaned, {failed} failed.")
    return 1 if failed else 0


# --- cli -----------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="metaclean", description=__doc__)
    sub = parser.add_subparsers(dest="mode", required=True)

    p_text = sub.add_parser("text", help="find/strip invisible Unicode characters")
    p_text.add_argument("--clean", action="store_true", help="strip the suspicious characters")
    p_text.add_argument("--file", help="read from this file instead of the clipboard ('-' for stdin)")
    p_text.add_argument("--out", help="write cleaned text to this file")
    p_text.add_argument("--to-clipboard", action="store_true", help="copy cleaned text back to the clipboard")
    p_text.set_defaults(func=cmd_text)

    p_pdf = sub.add_parser("pdf", help="find/strip PDF document metadata")
    pdf_sub = p_pdf.add_subparsers(dest="pdf_mode", required=True)
    p_pdf_find = pdf_sub.add_parser("find", help="print docinfo + XMP metadata")
    p_pdf_find.add_argument("path")
    p_pdf_find.set_defaults(func=cmd_pdf_find)
    p_pdf_clean = pdf_sub.add_parser("clean", help="strip docinfo + XMP metadata")
    p_pdf_clean.add_argument("path")
    group = p_pdf_clean.add_mutually_exclusive_group()
    group.add_argument("--out", help="output path (default: <name>.clean.pdf)")
    group.add_argument("--in-place", action="store_true", help="overwrite the input file")
    p_pdf_clean.set_defaults(func=cmd_pdf_clean)

    p_ooxml = sub.add_parser("ooxml", help="find/strip docx/xlsx/pptx metadata")
    ooxml_sub = p_ooxml.add_subparsers(dest="ooxml_mode", required=True)
    p_ooxml_find = ooxml_sub.add_parser("find", help="print docProps metadata")
    p_ooxml_find.add_argument("path")
    p_ooxml_find.set_defaults(func=cmd_ooxml_find)
    p_ooxml_clean = ooxml_sub.add_parser("clean", help="strip docProps metadata")
    p_ooxml_clean.add_argument("path")
    group2 = p_ooxml_clean.add_mutually_exclusive_group()
    group2.add_argument("--out", help="output path (default: <name>.clean.<ext>)")
    group2.add_argument("--in-place", action="store_true", help="overwrite the input file")
    p_ooxml_clean.set_defaults(func=cmd_ooxml_clean)

    p_auto = sub.add_parser("auto", help="drag-and-drop dispatcher: clean by extension")
    p_auto.add_argument("paths", nargs="*", help="files to clean; if omitted, process inbox/")
    p_auto.set_defaults(func=cmd_auto)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
