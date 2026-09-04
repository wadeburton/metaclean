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
import math
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


# --- pdf watermark detection/removal ----------------------------------------
#
# Watermarks are applied to PDFs in several structurally different ways, and
# only some of them can be undone cleanly:
#
#   1. A /Watermark annotation (Acrobat's native "Add Watermark" tool) —
#      a distinct object; deleting it is clean and exact.
#   2. An optional-content (layer) group named like "Watermark" — deleting
#      the layer and the marked-content region that references it is clean.
#   3. The same image/vector overlay (XObject) drawn on every page — deleting
#      the draw instruction is clean; the XObject itself is left orphaned
#      (harmless — pikepdf.save() won't emit unreferenced objects that matter
#      here, and we don't bother reclaiming the space).
#   4. Text drawn directly into each page's content stream (no reusable
#      object at all) — detectable only by heuristic (repeated across pages,
#      rotated and/or translucent) and removable only by deleting the exact
#      matching text-show instructions. False positives are possible (e.g. a
#      repeated running header), so this is opt-in via --aggressive.
#   5. A watermark baked into a flattened page image (a scan) — not
#      detectable or removable by this tool at all.
#
# find_pdf_watermarks() reports all of 1-4 it can find. strip_pdf_watermarks()
# always removes 1-3 (high confidence, no visual side effects); 4 only runs
# under aggressive=True.

def _pdf_string_operand(operand) -> str:
    try:
        return str(operand)
    except Exception:
        return ""


def _text_matrix_rotation_degrees(tm: list) -> float:
    try:
        a, b = float(tm[0]), float(tm[1])
    except (IndexError, TypeError, ValueError):
        return 0.0
    return math.degrees(math.atan2(b, a))


def _mat_multiply(m1: list, m2: list) -> list:
    """Combine two PDF affine matrices: apply m1, then m2 (PDF's `cm`
    semantics — the operand is applied before whatever the CTM already was).
    Real-world watermarking tools very often express rotation this way
    (canvas.rotate() -> a `cm` operator) rather than folding it into the
    text matrix (`Tm`) directly, so both have to be tracked and combined to
    get the actually-rendered rotation."""
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return [
        a1 * a2 + b1 * c2,
        a1 * b2 + b1 * d2,
        c1 * a2 + d1 * c2,
        c1 * b2 + d1 * d2,
        e1 * a2 + f1 * c2 + e2,
        e1 * b2 + f1 * d2 + f2,
    ]


def _scan_pdf_watermarks(pdf, n_pages: int) -> dict:
    import pikepdf

    report: dict[str, list] = {
        "annotations": [], "ocg_layers": [], "repeated_xobjects": [], "heuristic_text": [],
    }

    for i, page in enumerate(pdf.pages):
        annots = page.obj.get("/Annots")
        if not annots:
            continue
        for a in annots:
            if str(a.get("/Subtype", "")) == "/Watermark":
                report["annotations"].append({"page": i + 1})

    ocp = pdf.Root.get("/OCProperties")
    if ocp is not None:
        for ocg in ocp.get("/OCGs", []) or []:
            name = str(ocg.get("/Name", ""))
            if "watermark" in name.lower():
                report["ocg_layers"].append({"name": name})

    xobj_pages: dict[tuple, int] = {}
    xobj_name: dict[tuple, str] = {}
    for page in pdf.pages:
        resources = page.obj.get("/Resources")
        xobjects = resources.get("/XObject") if resources else None
        if not xobjects:
            continue
        seen = set()
        for name, ref in xobjects.items():
            try:
                key = ref.objgen
            except AttributeError:
                continue
            if key in seen:
                continue
            seen.add(key)
            xobj_pages[key] = xobj_pages.get(key, 0) + 1
            xobj_name.setdefault(key, str(name))
    if n_pages > 1:
        threshold = max(2, n_pages - 1)
        for key, count in xobj_pages.items():
            if count >= threshold:
                report["repeated_xobjects"].append({"name": xobj_name[key], "pages": count, "of": n_pages})

    text_occurrences: dict[str, list[dict]] = {}
    for page_index, page in enumerate(pdf.pages):
        resources = page.obj.get("/Resources")
        extgstates = resources.get("/ExtGState") if resources else None
        current_tm = [1, 0, 0, 1, 0, 0]
        current_alpha = 1.0
        ctm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
        ctm_stack: list[list] = []
        try:
            instructions = pikepdf.parse_content_stream(page)
        except Exception:
            continue
        for ins in instructions:
            op = str(ins.operator)
            if op == "q":
                ctm_stack.append(ctm)
            elif op == "Q":
                if ctm_stack:
                    ctm = ctm_stack.pop()
            elif op == "cm" and len(ins.operands) == 6:
                ctm = _mat_multiply([float(x) for x in ins.operands], ctm)
            elif op == "Tm" and len(ins.operands) == 6:
                current_tm = [float(x) for x in ins.operands]
            elif op == "gs" and extgstates is not None and ins.operands:
                gs = extgstates.get(str(ins.operands[0]))
                if gs is not None and "/ca" in gs:
                    try:
                        current_alpha = float(gs["/ca"])
                    except Exception:
                        pass
            elif op in ("Tj", "'", '"') and ins.operands:
                text = _pdf_string_operand(ins.operands[-1])
                if text.strip():
                    rotation = _text_matrix_rotation_degrees(_mat_multiply(current_tm, ctm))
                    text_occurrences.setdefault(text, []).append(
                        {"page": page_index + 1, "rotation": rotation, "alpha": current_alpha}
                    )
            elif op == "TJ" and ins.operands:
                text = "".join(_pdf_string_operand(el) for el in ins.operands[0])
                if text.strip():
                    rotation = _text_matrix_rotation_degrees(_mat_multiply(current_tm, ctm))
                    text_occurrences.setdefault(text, []).append(
                        {"page": page_index + 1, "rotation": rotation, "alpha": current_alpha}
                    )

    threshold = max(1, n_pages - 1) if n_pages > 1 else 1
    for text, occurrences in text_occurrences.items():
        if not (1 <= len(text) <= 80):
            continue
        pages_hit = sorted({o["page"] for o in occurrences})
        rotated = any(abs(o["rotation"]) > 5 for o in occurrences)
        transparent = any(o["alpha"] < 0.7 for o in occurrences)
        if len(pages_hit) >= threshold and (rotated or transparent):
            report["heuristic_text"].append(
                {"text": text, "pages": pages_hit, "of": n_pages, "rotated": rotated, "transparent": transparent}
            )

    return report


def find_pdf_watermarks(path: str) -> dict:
    import pikepdf

    with pikepdf.open(path) as pdf:
        return _scan_pdf_watermarks(pdf, len(pdf.pages))


def _strip_marked_content(pdf, page, marked_names: set[str]) -> bool:
    import pikepdf

    instructions = pikepdf.parse_content_stream(page)
    out = []
    depth = 0
    changed = False
    for ins in instructions:
        op = str(ins.operator)
        if depth == 0 and op == "BDC" and len(ins.operands) >= 2 and str(ins.operands[0]) == "/OC" and str(ins.operands[1]) in marked_names:
            depth = 1
            changed = True
            continue
        if depth > 0:
            if op in ("BDC", "BMC"):
                depth += 1
            elif op == "EMC":
                depth -= 1
            continue
        out.append(ins)
    if changed:
        page.contents_coalesce()
        page.obj.Contents = pdf.make_stream(pikepdf.unparse_content_stream(out))
    return changed


def _strip_xobject_draws(pdf, page, names_to_strip: set[str]) -> bool:
    import pikepdf

    instructions = pikepdf.parse_content_stream(page)
    out = []
    changed = False
    for ins in instructions:
        if str(ins.operator) == "Do" and ins.operands and str(ins.operands[0]) in names_to_strip:
            changed = True
            continue
        out.append(ins)
    if changed:
        page.contents_coalesce()
        page.obj.Contents = pdf.make_stream(pikepdf.unparse_content_stream(out))
    return changed


def _strip_text_runs(pdf, page, texts: set[str]) -> bool:
    import pikepdf

    instructions = pikepdf.parse_content_stream(page)
    out = []
    changed = False
    for ins in instructions:
        op = str(ins.operator)
        if op in ("Tj", "'", '"') and ins.operands and _pdf_string_operand(ins.operands[-1]) in texts:
            changed = True
            continue
        if op == "TJ" and ins.operands:
            text = "".join(_pdf_string_operand(el) for el in ins.operands[0])
            if text in texts:
                changed = True
                continue
        out.append(ins)
    if changed:
        page.contents_coalesce()
        page.obj.Contents = pdf.make_stream(pikepdf.unparse_content_stream(out))
    return changed


def strip_pdf_watermarks(in_path: str, out_path: str, aggressive: bool = False) -> dict:
    import pikepdf

    removed = {"annotations": 0, "ocg_layers": 0, "repeated_xobjects": 0, "heuristic_text": 0}

    with pikepdf.open(in_path) as pdf:
        n_pages = len(pdf.pages)
        report = _scan_pdf_watermarks(pdf, n_pages)

        for page in pdf.pages:
            annots = page.obj.get("/Annots")
            if not annots:
                continue
            keep = [a for a in annots if str(a.get("/Subtype", "")) != "/Watermark"]
            if len(keep) != len(annots):
                removed["annotations"] += len(annots) - len(keep)
                page.obj.Annots = keep

        watermark_ocg_keys = set()
        ocp = pdf.Root.get("/OCProperties")
        if ocp is not None and report["ocg_layers"]:
            kept_ocgs = []
            for ocg in ocp.get("/OCGs", []) or []:
                if "watermark" in str(ocg.get("/Name", "")).lower():
                    watermark_ocg_keys.add(ocg.objgen)
                    removed["ocg_layers"] += 1
                else:
                    kept_ocgs.append(ocg)
            if watermark_ocg_keys:
                ocp.OCGs = kept_ocgs
                d = ocp.get("/D")
                if d is not None:
                    for key in ("/ON", "/OFF"):
                        arr = d.get(key)
                        if arr:
                            d[key] = [x for x in arr if x.objgen not in watermark_ocg_keys]

        if watermark_ocg_keys:
            for page in pdf.pages:
                resources = page.obj.get("/Resources")
                properties = resources.get("/Properties") if resources else None
                if not properties:
                    continue
                marked_names = {
                    str(name) for name, ref in properties.items()
                    if getattr(ref, "objgen", None) in watermark_ocg_keys
                }
                if marked_names:
                    _strip_marked_content(pdf, page, marked_names)

        if report["repeated_xobjects"]:
            xobj_pages: dict[tuple, int] = {}
            for page in pdf.pages:
                resources = page.obj.get("/Resources")
                xobjects = resources.get("/XObject") if resources else None
                if not xobjects:
                    continue
                seen = set()
                for _, ref in xobjects.items():
                    try:
                        key = ref.objgen
                    except AttributeError:
                        continue
                    if key not in seen:
                        seen.add(key)
                        xobj_pages[key] = xobj_pages.get(key, 0) + 1
            threshold = max(2, n_pages - 1)
            watermark_xobj_keys = {k for k, c in xobj_pages.items() if c >= threshold}
            for page in pdf.pages:
                resources = page.obj.get("/Resources")
                xobjects = resources.get("/XObject") if resources else None
                if not xobjects:
                    continue
                names_to_strip = {
                    str(name) for name, ref in xobjects.items()
                    if getattr(ref, "objgen", None) in watermark_xobj_keys
                }
                if names_to_strip and _strip_xobject_draws(pdf, page, names_to_strip):
                    removed["repeated_xobjects"] += 1

        if aggressive and report["heuristic_text"]:
            texts = {h["text"] for h in report["heuristic_text"]}
            for page in pdf.pages:
                if _strip_text_runs(pdf, page, texts):
                    removed["heuristic_text"] += 1

        pdf.save(out_path)

    return removed


def cmd_pdf_watermark_find(args: argparse.Namespace) -> int:
    report = find_pdf_watermarks(args.path)
    _print_watermark_report(report)
    return 0


def cmd_pdf_watermark_clean(args: argparse.Namespace) -> int:
    out_path = args.path if args.in_place else (args.out or _default_clean_name(args.path, ".pdf"))
    removed = strip_pdf_watermarks(args.path, out_path, aggressive=args.aggressive)
    print(f"Watermarks stripped -> {out_path}", file=sys.stderr)
    print(f"  annotations removed:       {removed['annotations']}", file=sys.stderr)
    print(f"  OCG watermark layers:      {removed['ocg_layers']}", file=sys.stderr)
    print(f"  repeated overlays removed: {removed['repeated_xobjects']}", file=sys.stderr)
    print(f"  heuristic text runs:       {removed['heuristic_text']}" + ("" if args.aggressive else "  (pass --aggressive to attempt these)"), file=sys.stderr)
    print(
        "Note: a watermark baked into a flattened/scanned page image cannot be removed by this tool.",
        file=sys.stderr,
    )
    return 0


def _print_watermark_report(report: dict) -> None:
    if not any(report.values()):
        print("No watermark-like structures found (annotations, layers, repeated overlays, or repeated rotated/translucent text).")
        print("This does NOT rule out a watermark baked directly into a flattened/scanned page image.")
        return

    if report["annotations"]:
        pages = ", ".join(str(a["page"]) for a in report["annotations"])
        print(f"Watermark annotations - pages: {pages}  (safe to remove)")
    if report["ocg_layers"]:
        names = ", ".join(o["name"] for o in report["ocg_layers"])
        print(f"Watermark-named layers (OCGs): {names}  (safe to remove)")
    if report["repeated_xobjects"]:
        for x in report["repeated_xobjects"]:
            print(f"Repeated overlay object '{x['name']}' on {x['pages']}/{x['of']} pages  (safe to remove)")
    if report["heuristic_text"]:
        for h in report["heuristic_text"]:
            traits = []
            if h["rotated"]:
                traits.append("rotated")
            if h["transparent"]:
                traits.append("translucent")
            traits_s = "/".join(traits) or "plain"
            print(f"Possible text watermark {h['text']!r} - {traits_s}, on {len(h['pages'])}/{h['of']} pages  (heuristic - needs --aggressive to remove, may false-positive on repeated headers/footers)")


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


# --- docx header watermark detection/removal --------------------------------
#
# Word's built-in "Insert Watermark" feature always places a legacy VML
# <w:pict> shape into the document's header part(s). This is a very
# consistent, well-known pattern: text watermarks use <v:textpath> (WordArt
# on a path), picture watermarks use an absolutely-positioned <v:imagedata>,
# and Word itself names the shape with an id containing
# "PowerPlusWaterMarkObject". Any one of those is a reliable, low-false-
# positive signal — real header content (page numbers, running titles, a
# logo inserted the normal DrawingML way) doesn't look like this.
#
# Scope: docx/dotx/docm headers only. pptx/xlsx watermarks aren't
# standardized the same way and aren't covered here.

_WATERMARK_ID_MARKER = "PowerPlusWaterMarkObject"


def _local_name(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _find_watermark_picts(root) -> list[dict]:
    parent_map = {child: parent for parent in root.iter() for child in parent}
    hits = []
    for elem in root.iter():
        if _local_name(elem.tag) != "pict":
            continue
        shape = None
        shape_id = ""
        has_textpath = False
        has_absolute_image = False
        for sub in elem.iter():
            name = _local_name(sub.tag)
            if name == "shape" and shape is None:
                shape = sub
                shape_id = sub.get("id", "")
            elif name == "textpath":
                has_textpath = True
            elif name == "imagedata":
                style = shape.get("style", "") if shape is not None else ""
                if "position:absolute" in style:
                    has_absolute_image = True
        if _WATERMARK_ID_MARKER in shape_id or has_textpath or has_absolute_image:
            hits.append({
                "element": elem,
                "parent": parent_map.get(elem),
                "shape_id": shape_id,
                "kind": "text" if has_textpath else ("image" if has_absolute_image else "unknown"),
            })
    return hits


def find_ooxml_watermarks(path: str) -> list[dict]:
    hits = []
    with zipfile.ZipFile(path) as zf:
        header_parts = sorted(
            n for n in zf.namelist() if n.startswith("word/header") and n.endswith(".xml")
        )
        for part in header_parts:
            root = ET.fromstring(zf.read(part))
            for hit in _find_watermark_picts(root):
                hits.append({"part": part, "shape_id": hit["shape_id"], "kind": hit["kind"]})
    return hits


def strip_ooxml_watermarks(in_path: str, out_path: str) -> int:
    removed = 0
    with zipfile.ZipFile(in_path) as zin:
        tmp_out = out_path + ".tmp"
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/header") and item.filename.endswith(".xml"):
                    root = ET.fromstring(data)
                    hits = _find_watermark_picts(root)
                    for hit in hits:
                        parent = hit["parent"]
                        if parent is not None:
                            parent.remove(hit["element"])
                            removed += 1
                    if hits:
                        data = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
                zout.writestr(item, data)
    Path(tmp_out).replace(out_path)
    return removed


def cmd_ooxml_watermark_find(args: argparse.Namespace) -> int:
    hits = find_ooxml_watermarks(args.path)
    if not hits:
        print(
            "No Word-style header watermark shapes found. (pptx/xlsx watermarks and "
            "non-standard docx watermarks aren't covered by this scan.)"
        )
        return 0
    for h in hits:
        marker = f"  id={h['shape_id']!r}" if h["shape_id"] else ""
        print(f"{h['part']}: {h['kind']} watermark shape{marker}  (safe to remove)")
    return 0


def cmd_ooxml_watermark_clean(args: argparse.Namespace) -> int:
    p = Path(args.path)
    out_path = args.path if args.in_place else (args.out or _default_clean_name(args.path, p.suffix.lower()))
    removed = strip_ooxml_watermarks(args.path, out_path)
    print(f"Removed {removed} watermark shape(s) -> {out_path}", file=sys.stderr)
    if removed == 0:
        print(
            "Note: no Word-style header watermark was found. A picture/text box added "
            "by hand, or a pptx/xlsx watermark, isn't covered by this scan.",
            file=sys.stderr,
        )
    return 0


# --- auto (drag-and-drop) mode ---------------------------------------------

TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".log"}
DOCX_WATERMARK_EXTENSIONS = {".docx", ".dotx", ".docm"}


def _clean_one_file(src: Path, dest: Path, strip_watermarks: bool = False) -> str:
    """Clean src into dest (dest's parent is created if needed). Returns a
    short status string describing what happened."""
    import os
    import tempfile

    dest.parent.mkdir(parents=True, exist_ok=True)
    ext = src.suffix.lower()

    if ext == ".pdf":
        if not strip_watermarks:
            strip_pdf_metadata(str(src), str(dest))
            return "PDF metadata stripped"
        fd, tmp_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        try:
            strip_pdf_metadata(str(src), tmp_path)
            removed = strip_pdf_watermarks(tmp_path, str(dest))
            total = sum(removed.values())
            status = "PDF metadata stripped"
            status += f", {total} watermark element(s) removed" if total else ", no removable watermark structures found"
            return status
        finally:
            Path(tmp_path).unlink(missing_ok=True)

    if ext in OOXML_EXTENSIONS or (ext not in TEXT_EXTENSIONS and _is_ooxml(str(src))):
        if not (strip_watermarks and ext in DOCX_WATERMARK_EXTENSIONS):
            strip_ooxml_metadata(str(src), str(dest))
            return "Office document metadata stripped"
        fd, tmp_path = tempfile.mkstemp(suffix=ext)
        os.close(fd)
        try:
            strip_ooxml_metadata(str(src), tmp_path)
            removed = strip_ooxml_watermarks(tmp_path, str(dest))
            status = "Office document metadata stripped"
            status += f", {removed} watermark shape(s) removed" if removed else ", no header watermark found"
            return status
        finally:
            Path(tmp_path).unlink(missing_ok=True)

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
            status = _clean_one_file(src, dest, strip_watermarks=args.watermarks)
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

    p_pdf_wm_find = pdf_sub.add_parser("watermark-find", help="scan for removable watermark structures")
    p_pdf_wm_find.add_argument("path")
    p_pdf_wm_find.set_defaults(func=cmd_pdf_watermark_find)
    p_pdf_wm_clean = pdf_sub.add_parser("watermark-clean", help="strip watermark structures")
    p_pdf_wm_clean.add_argument("path")
    p_pdf_wm_clean.add_argument(
        "--aggressive", action="store_true",
        help="also strip heuristically-detected repeated rotated/translucent text (may false-positive)",
    )
    group_wm = p_pdf_wm_clean.add_mutually_exclusive_group()
    group_wm.add_argument("--out", help="output path (default: <name>.clean.pdf)")
    group_wm.add_argument("--in-place", action="store_true", help="overwrite the input file")
    p_pdf_wm_clean.set_defaults(func=cmd_pdf_watermark_clean)

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

    p_ooxml_wm_find = ooxml_sub.add_parser("watermark-find", help="scan docx headers for watermark shapes")
    p_ooxml_wm_find.add_argument("path")
    p_ooxml_wm_find.set_defaults(func=cmd_ooxml_watermark_find)
    p_ooxml_wm_clean = ooxml_sub.add_parser("watermark-clean", help="strip docx header watermark shapes")
    p_ooxml_wm_clean.add_argument("path")
    group3 = p_ooxml_wm_clean.add_mutually_exclusive_group()
    group3.add_argument("--out", help="output path (default: <name>.clean.<ext>)")
    group3.add_argument("--in-place", action="store_true", help="overwrite the input file")
    p_ooxml_wm_clean.set_defaults(func=cmd_ooxml_watermark_clean)

    p_auto = sub.add_parser("auto", help="drag-and-drop dispatcher: clean by extension")
    p_auto.add_argument("paths", nargs="*", help="files to clean; if omitted, process inbox/")
    p_auto.add_argument(
        "--watermarks", action="store_true",
        help="also attempt watermark removal (PDF: annotations/layers/repeated overlays; docx: header watermark shapes)",
    )
    p_auto.set_defaults(func=cmd_auto)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
