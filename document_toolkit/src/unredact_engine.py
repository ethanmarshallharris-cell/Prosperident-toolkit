"""
Generic un-redaction: given a redacted PDF (produced by any document-type
plugin in document_types/, using the `<EntityLabel N>` placeholder
convention documented in document_types/base.py) and the CSV key file
written alongside it by that plugin's write_key(), produces a PDF with
every placeholder swapped back for the real value it originally replaced.

This is deliberately independent of any single document-type plugin: the
placeholder convention and the key-file format (see e.g.
document_types/dentrix_audit_trail.py's write_key()) are the same no
matter which plugin produced a given redacted file, so there is only ever
one way this needs to work -- it doesn't need to know or care what kind of
document it's un-redacting.

Uses the same whole-page "capture every character, wipe the page, redraw
everything except what's being replaced" strategy as redaction, for the
same reason it's used there: it's the only approach shown safe on
tightly-spaced real-world reports, where a partial/rectangle-only
text-replacement risks touching a neighboring line (see any plugin's
apply_redactions() docstring for the full history of why).

A couple of the small helper functions below (_map_fontname,
_collect_kept_runs, _fit_fontsize) are intentionally duplicated from
document_types/dentrix_audit_trail.py rather than imported from it -- this
module has to work regardless of which plugin produced the file it's
un-redacting, so it shouldn't depend on any one of them.

Limitation worth knowing: the restored name is drawn back in a plain
sans-serif font at an auto-fit size, not necessarily whatever font the
ORIGINAL pre-redaction document used at that spot -- once a page has been
redacted, that original styling information is gone. The restored text is
correct; its typography is an approximation.
"""
import csv
import re

import fitz  # PyMuPDF

PLACEHOLDER_RE = re.compile(r"^<[^<>]+ \d+>$")


def read_key_csv(key_path: str) -> dict:
    """Returns {placeholder_text: real_value}, e.g. {'<Patient 1>': 'Smith, John'}.
    Skips the explanatory/confidentiality-warning rows at the top of the
    file and reads only the "Placeholder,Real Name" table beneath them.
    Raises ValueError if the file doesn't look like a key file at all
    (e.g. the wrong file was picked), rather than silently returning an
    empty/useless mapping.
    """
    with open(key_path, "r", newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))

    mapping = {}
    started = False
    for row in rows:
        if not row:
            continue
        if not started:
            if len(row) >= 2 and row[0].strip() == "Placeholder" and row[1].strip() == "Real Name":
                started = True
            continue
        if len(row) >= 2 and PLACEHOLDER_RE.match(row[0].strip()):
            mapping[row[0].strip()] = row[1]

    if not mapping:
        raise ValueError(
            "This doesn't look like a name key CSV -- no 'Placeholder, Real Name' "
            "table was found in it. Make sure this is the *_NAME_KEY.csv file "
            "saved alongside the redacted PDF, not some other file."
        )
    return mapping


def _map_fontname(original_font: str, flags: int) -> str:
    # Identical logic to document_types/dentrix_audit_trail.py's
    # _map_fontname() -- duplicated rather than imported; see module
    # docstring for why.
    bold = bool(flags & 16)
    italic = bool(flags & 2)
    serifed = bool(flags & 4)
    monospaced = bool(flags & 8)
    name_lower = (original_font or "").lower()
    if monospaced or "courier" in name_lower or "mono" in name_lower:
        base = "co"
    elif serifed or "times" in name_lower or "georgia" in name_lower or "serif" in name_lower:
        base = "ti"
    else:
        base = "he"
    if base == "he":
        suffix = {(False, False): "lv", (True, False): "bo", (False, True): "it", (True, True): "bi"}
        return "he" + suffix[(bold, italic)]
    if base == "ti":
        suffix = {(False, False): "ro", (True, False): "bo", (False, True): "it", (True, True): "bi"}
        return "ti" + suffix[(bold, italic)]
    suffix = {(False, False): "u", (True, False): "ub", (False, True): "ui", (True, True): "ubi"}
    return "co" + suffix[(bold, italic)]


def _fit_fontsize(label: str, available_width: float, base: float = 9.0) -> float:
    # Identical logic to document_types/dentrix_audit_trail.py's
    # _fit_fontsize().
    approx_width = len(label) * base * 0.5
    if approx_width <= available_width or available_width <= 0:
        return base
    scaled = base * (available_width / approx_width)
    return max(scaled, 5.0)


def _collect_kept_runs(raw: dict, redact_rects):
    # Identical logic to document_types/dentrix_audit_trail.py's
    # _collect_kept_runs() -- center-point containment (not bbox overlap)
    # against plain float tuples, for the same correctness and performance
    # reasons documented in detail there.
    redact_tuples = [(r.x0, r.y0, r.x1, r.y1) for r in redact_rects]
    runs = []
    for block in raw.get("blocks", []):
        if block.get("type", 0) != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                font = span.get("font", "")
                size = span.get("size", 9.0)
                color = span.get("color", 0)
                flags = span.get("flags", 0)
                cur_run = None
                for ch in span.get("chars", []):
                    cx0, cy0, cx1, cy1 = ch["bbox"]
                    cx = (cx0 + cx1) / 2.0
                    cy = (cy0 + cy1) / 2.0
                    redacted = False
                    for rx0, ry0, rx1, ry1 in redact_tuples:
                        if rx0 <= cx <= rx1 and ry0 <= cy <= ry1:
                            redacted = True
                            break
                    if redacted:
                        cur_run = None
                        continue
                    if cur_run is None:
                        cur_run = {
                            "origin": ch["origin"], "text": ch["c"], "font": font,
                            "size": size, "color": color, "flags": flags,
                        }
                        runs.append(cur_run)
                    else:
                        cur_run["text"] += ch["c"]
    return runs


def find_placeholders(pdf_path: str, placeholder_set) -> dict:
    """Returns {page_index: [(rect, placeholder_text), ...]} for every
    occurrence of a known placeholder found in the PDF's text layer. Uses
    'dict'-mode span text: each placeholder was originally drawn as one
    contiguous insert_text() call, so it comes back as one span with that
    exact text, rather than needing word-level reassembly.
    """
    doc = fitz.open(pdf_path)
    hits_by_page = {}
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            d = page.get_text("dict")
            page_hits = []
            for block in d.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "")
                        if text in placeholder_set:
                            page_hits.append((fitz.Rect(span["bbox"]), text))
            if page_hits:
                hits_by_page[page_index] = page_hits
    finally:
        doc.close()
    return hits_by_page


def apply_unredaction(
    input_path: str,
    output_path: str,
    key_mapping: dict,
    progress_callback=None,
    phase_callback=None,
) -> dict:
    """Writes a copy of input_path with every placeholder in key_mapping
    swapped back for its real value. progress_callback(current, total) and
    phase_callback(phase_name) behave exactly as documented in
    document_types.base for a plugin's apply_redactions().

    Returns a small summary dict: {'pages_changed': int, 'replacements': int,
    'placeholders_not_found': list[str]} -- the last of these is placeholders
    present in the key file but never found in the PDF, which usually means
    this key file doesn't actually belong to this PDF (wrong pairing) and is
    worth surfacing to the user rather than silently ignoring.
    """
    doc = fitz.open(input_path)
    try:
        placeholder_set = set(key_mapping)
        hits_by_page = {}
        found_placeholders = set()
        for page_index in range(len(doc)):
            page = doc[page_index]
            d = page.get_text("dict")
            page_hits = []
            for block in d.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "")
                        if text in placeholder_set:
                            page_hits.append((fitz.Rect(span["bbox"]), text))
                            found_placeholders.add(text)
            if page_hits:
                hits_by_page[page_index] = page_hits

        total = len(hits_by_page)
        replacements = 0
        for done_count, (page_index, page_hits) in enumerate(hits_by_page.items(), start=1):
            page = doc[page_index]
            raw = page.get_text("rawdict")
            rects = [r for r, _ in page_hits]
            runs = _collect_kept_runs(raw, rects)

            page.add_redact_annot(page.rect)
            page.apply_redactions(images=0, graphics=0, text=0)

            shape = page.new_shape()
            for run in runs:
                fontname = _map_fontname(run["font"], run["flags"])
                color = fitz.sRGB_to_pdf(run["color"])
                try:
                    shape.insert_text(
                        run["origin"], run["text"], fontname=fontname, fontsize=run["size"], color=color
                    )
                except Exception:
                    shape.insert_text(
                        run["origin"], run["text"], fontname="helv", fontsize=run["size"], color=color
                    )

            for rect, placeholder_text in page_hits:
                real_value = key_mapping[placeholder_text]
                fontsize = _fit_fontsize(real_value, max(rect.width, 60.0), base=9.0)
                shape.insert_text(
                    (rect.x0, rect.y1 - 1.5), real_value, fontname="helv", fontsize=fontsize, color=(0, 0, 0)
                )
                replacements += 1
            shape.commit()

            if progress_callback is not None:
                try:
                    progress_callback(done_count, total)
                except Exception:
                    pass

        if phase_callback is not None:
            try:
                phase_callback("saving")
            except Exception:
                pass

        doc.set_metadata({})
        try:
            doc.del_xml_metadata()
        except Exception:
            pass
        doc.save(output_path, garbage=1, deflate=True, clean=True)

        return {
            "pages_changed": total,
            "replacements": replacements,
            "placeholders_not_found": sorted(placeholder_set - found_placeholders),
        }
    finally:
        doc.close()


def verify_unredaction(output_path: str, key_mapping: dict) -> dict:
    """Independent post-check, same spirit as a redaction plugin's
    verify_redaction(): none of the placeholders should still be present
    in the output's extracted text -- every one that was found should have
    been fully swapped for its real value, nothing left half-done.
    Returns {'ok': bool, 'issues': list[str]}.
    """
    issues = []
    doc = fitz.open(output_path)
    try:
        text_all = "\n".join(p.get_text() for p in doc)
        leftover = sorted(ph for ph in key_mapping if ph in text_all)
        if leftover:
            shown = ", ".join(leftover[:10])
            more = f" (+{len(leftover) - 10} more)" if len(leftover) > 10 else ""
            issues.append(f"Placeholder(s) still present in output, not replaced: {shown}{more}")
    finally:
        doc.close()
    return {"ok": len(issues) == 0, "issues": issues}
