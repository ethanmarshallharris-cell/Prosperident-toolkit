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

Restoring the original font: newer key files (see
document_types/dentrix_audit_trail.py's write_name_key()) record the
redacted text's own font/size/color/flags, sampled at redaction time
before that text was wiped, in extra "Font", "Size", "Color", "Flags"
columns. When present, the restored name is drawn back using that
captured styling instead of a plain default -- as close a match to the
original as redaction (which by design keeps nothing else about that
exact spot) allows. Key files from before this existed only have
"Placeholder" and "Real Name" columns; those are still read the same as
ever, just with no style to restore, so un-redaction still works, just
back in a plain default font as before.
"""
import csv
import re

import fitz  # PyMuPDF

PLACEHOLDER_RE = re.compile(r"^<[^<>]+ \d+>$")
# Same shape as PLACEHOLDER_RE but not anchored to a whole string -- used to
# scan free-flowing extracted PDF text for ANY placeholder-shaped substring,
# not just ones already known from the loaded key file(s). See
# verify_unredaction()'s docstring for why this matters: a placeholder that
# never matched during apply_unredaction() (e.g. a literal-text mismatch
# against what the key file has, even by a single character) is invisible
# to a check that only looks for KNOWN placeholders, since it was never
# "known" to begin with from that mismatched entry's point of view.
_PLACEHOLDER_SCAN_RE = re.compile(r"<[^<>]+ \d+>")


def read_key_csv(key_path: str) -> dict:
    """Returns {placeholder_text: entry}, where entry is a dict with at
    least 'value' (the real value that placeholder replaced), and
    optionally 'font' (str), 'size' (float), 'color' (fitz-ready (r,g,b)
    0-1 tuple), 'flags' (int) if this key file recorded them (see module
    docstring) -- any of those not present in the file are None.

    Skips the explanatory/confidentiality-warning rows at the top of the
    file and reads only the table beneath them. Raises ValueError if the
    file doesn't look like a key file at all (e.g. the wrong file was
    picked), rather than silently returning an empty/useless mapping.
    """
    with open(key_path, "r", newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))

    mapping = {}
    header = None
    for row in rows:
        if not row:
            continue
        if header is None:
            if len(row) >= 2 and row[0].strip() == "Placeholder" and row[1].strip() == "Real Name":
                header = [c.strip() for c in row]
            continue
        if len(row) >= 2 and PLACEHOLDER_RE.match(row[0].strip()):
            by_col = dict(zip(header, row))
            font = by_col.get("Font") or None
            size_raw = by_col.get("Size") or ""
            color_raw = by_col.get("Color") or ""
            flags_raw = by_col.get("Flags") or ""
            size = float(size_raw) if size_raw.strip() else None
            color = _hex_to_unit_rgb(color_raw.strip()) if color_raw.strip() else None
            flags = int(flags_raw) if flags_raw.strip() else None
            mapping[row[0].strip()] = {
                "value": row[1], "font": font, "size": size, "color": color, "flags": flags,
            }

    if not mapping:
        raise ValueError(
            "This doesn't look like a name key CSV -- no 'Placeholder, Real Name' "
            "table was found in it. Make sure this is the *_NAME_KEY.csv file "
            "saved alongside the redacted PDF, not some other file."
        )
    return mapping


def merge_key_mappings(mappings: "list"):
    """Combines multiple {placeholder: entry} key mappings (one per key
    CSV the user picked in the Un-redact tab -- a PDF may have been
    redacted in more than one pass, each producing its own *_NAME_KEY.csv)
    into a single mapping usable by apply_unredaction(). The common case
    is several key files with no overlapping placeholders at all -- those
    combine with no fuss.

    Returns (merged, conflicts). conflicts is a list of {'placeholder':
    str, 'values': list[str]} entries, one per placeholder whose 'value'
    (the real name it stands for) differs across the given files --
    styling differences (font/size/color) alone are NOT a conflict, only
    the actual value is compared, since two genuine key files for the
    same underlying redaction should never disagree about what a
    placeholder means. A conflict almost always means one of the selected
    files doesn't actually belong with the others (wrong file added by
    mistake, or a key file from an entirely different PDF).
    main_gui.py treats any conflict as blocking and shows it to the user
    rather than silently guessing; 'merged' still contains one (the
    first-seen) of the conflicting entries for each conflicting
    placeholder, in case a caller ever needs to proceed anyway, but the
    normal path is to stop and let the user fix their file selection.
    """
    values_by_ph = {}   # placeholder -> [distinct 'value's, first-seen order]
    entry_by_ph = {}     # placeholder -> first-seen full entry (kept for its styling)
    for mapping in mappings:
        for ph, entry in mapping.items():
            val = entry["value"]
            seen = values_by_ph.setdefault(ph, [])
            if val not in seen:
                seen.append(val)
            entry_by_ph.setdefault(ph, entry)

    merged = dict(entry_by_ph)
    conflicts = [
        {"placeholder": ph, "values": values} for ph, values in values_by_ph.items() if len(values) > 1
    ]
    conflicts.sort(key=lambda c: c["placeholder"])
    return merged, conflicts


def _hex_to_unit_rgb(hex_str: str):
    """'#RRGGBB' -> (r, g, b) with each component 0-1, as fitz's insert_text
    color= expects. Returns None for anything that doesn't parse cleanly
    -- a malformed color is a reason to fall back to a default, not to
    fail the whole un-redaction."""
    h = hex_str.lstrip("#")
    if len(h) != 6:
        return None
    try:
        r = int(h[0:2], 16) / 255.0
        g = int(h[2:4], 16) / 255.0
        b = int(h[4:6], 16) / 255.0
        return (r, g, b)
    except ValueError:
        return None


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


def _find_placeholder_hits(raw: dict, placeholder_set) -> list:
    """Scans one page's 'rawdict' text for every occurrence of a
    placeholder from placeholder_set, returning [(rect, placeholder_text),
    ...] -- one entry per occurrence, with rect computed from just the
    matched characters.

    Looks for each placeholder as a SUBSTRING within a span's characters,
    rather than requiring the whole span's text to equal the placeholder
    exactly. This matters because MuPDF's text extraction groups
    characters into a "span" by style continuity (same font/size/color/
    flags) and line position, not by which insert_text() call originally
    drew them -- so a placeholder run can end up merged into the same span
    as neighboring kept text when both happen to share styling and sit
    close together on the line (common on real-world reports where a name
    column directly abuts the next field for some row lengths but not
    others). Requiring an exact whole-span match silently missed exactly
    those occurrences: nothing was found to redact there, so the original
    placeholder text got redrawn verbatim as "kept" text, appearing to
    have "not un-redacted" a placeholder the key file actually did cover.
    Matching by substring and slicing the exact matched characters' bboxes
    (rather than the whole span's bbox) finds these occurrences too,
    without needing to know in advance whether a given placeholder shares
    its span with other text.
    """
    hits = []
    for block in raw.get("blocks", []):
        if block.get("type", 0) != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                chars = span.get("chars", [])
                if not chars:
                    continue
                text = "".join(ch["c"] for ch in chars)
                for m in _PLACEHOLDER_SCAN_RE.finditer(text):
                    ph_text = m.group(0)
                    if ph_text not in placeholder_set:
                        continue
                    matched_chars = chars[m.start():m.end()]
                    x0 = min(c["bbox"][0] for c in matched_chars)
                    y0 = min(c["bbox"][1] for c in matched_chars)
                    x1 = max(c["bbox"][2] for c in matched_chars)
                    y1 = max(c["bbox"][3] for c in matched_chars)
                    hits.append((fitz.Rect(x0, y0, x1, y1), ph_text))
    return hits


def find_placeholders(pdf_path: str, placeholder_set) -> dict:
    """Returns {page_index: [(rect, placeholder_text), ...]} for every
    occurrence of a known placeholder found in the PDF's text layer. See
    _find_placeholder_hits()'s docstring for why this matches placeholders
    as substrings of a 'rawdict' span's characters rather than requiring
    the whole span to equal the placeholder exactly.
    """
    doc = fitz.open(pdf_path)
    hits_by_page = {}
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            raw = page.get_text("rawdict")
            page_hits = _find_placeholder_hits(raw, placeholder_set)
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
        raws_by_page = {}
        found_placeholders = set()
        for page_index in range(len(doc)):
            page = doc[page_index]
            raw = page.get_text("rawdict")
            page_hits = _find_placeholder_hits(raw, placeholder_set)
            if page_hits:
                hits_by_page[page_index] = page_hits
                raws_by_page[page_index] = raw
                found_placeholders.update(text for _, text in page_hits)
        # (Same scan as find_placeholders() above, done inline here because
        # we also need found_placeholders for the placeholders_not_found
        # summary below, and we keep each hit page's rawdict around so the
        # redraw pass below doesn't have to re-extract it -- find_placeholders()
        # is kept as a standalone entry point for callers that just want the
        # raw hit locations.)

        total = len(hits_by_page)
        replacements = 0
        for done_count, (page_index, page_hits) in enumerate(hits_by_page.items(), start=1):
            page = doc[page_index]
            raw = raws_by_page[page_index]
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
                entry = key_mapping[placeholder_text]
                real_value = entry["value"]
                if entry.get("size") is not None:
                    fontsize = entry["size"]
                else:
                    fontsize = _fit_fontsize(real_value, max(rect.width, 60.0), base=9.0)
                if entry.get("font") is not None:
                    fontname = _map_fontname(entry["font"], entry.get("flags") or 0)
                else:
                    fontname = "helv"
                color = entry.get("color") if entry.get("color") is not None else (0, 0, 0)
                try:
                    shape.insert_text(
                        (rect.x0, rect.y1 - 1.5), real_value, fontname=fontname, fontsize=fontsize, color=color
                    )
                except Exception:
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

    Checks two different things, because they catch two different
    failures:
      1. Every placeholder KNOWN from the loaded key file(s) -- same as
         before. This is the common case (key file simply doesn't cover
         this PDF, wrong pairing, etc).
      2. ANY placeholder-SHAPED text still in the output at all, known or
         not. This exists because #1 alone has a blind spot: if a
         placeholder's literal text in the PDF doesn't EXACTLY match the
         key file's entry for it -- even a mismatch invisible to the eye,
         or a genuinely different number assigned to it than what the key
         file has for that name -- apply_unredaction() has nothing to
         match it against and silently leaves it untouched, and a check
         that only looks for KNOWN placeholders would report success
         anyway, hiding the very problem it exists to catch. Since a
         redacted PDF should never legitimately contain free text that
         happens to look like "<Word 123>", this generic scan is safe to
         treat as always worth flagging.
    Returns {'ok': bool, 'issues': list[str], 'leftover': list[dict]}.
    'leftover' is the full, UNTRUNCATED list behind 'issues' -- one
    {'placeholder': str, 'known': bool} entry per distinct placeholder-shaped
    string still in the output ('known' True if it was recognized from the
    loaded key file(s), False if it merely looks like a placeholder but
    matched nothing at all) -- main_gui.py uses this to write a complete
    diagnostic log even when there are more than the 10 shown on-screen.
    """
    issues = []
    leftover = []
    doc = fitz.open(output_path)
    try:
        text_all = "\n".join(p.get_text() for p in doc)
        known_leftover = sorted(ph for ph in key_mapping if ph in text_all)
        if known_leftover:
            shown = ", ".join(known_leftover[:10])
            more = f" (+{len(known_leftover) - 10} more)" if len(known_leftover) > 10 else ""
            issues.append(f"Placeholder(s) still present in output, not replaced: {shown}{more}")
            leftover.extend({"placeholder": ph, "known": True} for ph in known_leftover)

        all_shaped = sorted(set(_PLACEHOLDER_SCAN_RE.findall(text_all)))
        unknown_shaped = [ph for ph in all_shaped if ph not in key_mapping]
        if unknown_shaped:
            shown = ", ".join(unknown_shaped[:10])
            more = f" (+{len(unknown_shaped) - 10} more)" if len(unknown_shaped) > 10 else ""
            issues.append(
                "Placeholder-shaped text still present in output that isn't in any loaded key "
                f"file at all: {shown}{more}. This usually means either the key file you loaded "
                "doesn't actually cover every placeholder in this PDF (a follow-up redaction pass "
                "produced a second key file you haven't loaded, for instance), or this exact "
                "placeholder's number doesn't match what any loaded key file recorded for it -- "
                "check the key file directly for the placeholder text shown above."
            )
            leftover.extend({"placeholder": ph, "known": False} for ph in unknown_shaped)
    finally:
        doc.close()
    return {"ok": len(issues) == 0, "issues": issues, "leftover": leftover}
