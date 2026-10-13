"""
Un-redaction for Word documents (.docx) -- added October 2026.

The use case: a report or memo drafted in Word FROM redacted material (a
redacted Dentrix audit trail, a redacted Excel export, an analysis
someone wrote using only the de-identified files) quotes placeholders such
as <Patient 7> or <Name 3> in its sentences and tables. Once the work is
done, every placeholder needs its real value back. Nothing in this app ever
redacts a Word document itself -- this module only reverses placeholders a
Word document happens to contain, using the key CSV(s) saved by the
redactions they came from.

--------------------------------------------------------------------------
KEY FILES ACCEPTED
--------------------------------------------------------------------------
Both key formats this app writes, in any combination:
  - a PDF name key ("Placeholder, Real Name[, Font, Size, Color, Flags]"),
    e.g. *_NAME_KEY.csv from a Dentrix Audit Trail Report; and
  - a spreadsheet redaction key ("Column, Placeholder, Real Value"),
    *_REDACTION_KEY.csv, including the "Combined: First Name + Last Name"
    summary rows a combined-columns redaction writes (see
    excel_spreadsheet.py's "COMBINED COLUMNS" section). A combination is
    put back as one name: "John Smith", or "Smith, John" when the first
    combined column is a last-name column (its header contains "last" or
    "surname") -- the same way Dentrix writes names, so it agrees with a
    PDF key for the same linked patient.

A Word document has no columns, so a placeholder that means different
things in different places of the key(s) -- the same <Name 3> label typed
for two different spreadsheet columns, or one patient spelled "Smith,
John" in a PDF key and "John Smith" in a spreadsheet key -- cannot be put
back automatically. placeholders_needing_choice() lists those (only the
ones actually present in the document), main_gui.py asks the user to pick
one value for each, and resolve_mapping() turns the keys plus those picks
into one final {placeholder: real value} table. A placeholder the user
chooses to leave, or one no key file knows, stays as it is and is listed
in the completion message.

--------------------------------------------------------------------------
HOW THE DOCUMENT IS EDITED
--------------------------------------------------------------------------
Standard library only (zipfile + re) -- no python-docx, because that needs
lxml, which the build deliberately leaves out (see build_windows.bat).
And deliberately NOT by parsing and re-writing the XML with ElementTree:
re-serializing WordprocessingML that way renames namespace prefixes and
drops declarations Word relies on (mc:Ignorable), which Word then reports
as a corrupt file. Instead, the text of every <w:t> element is edited in
place inside the original XML string; every other byte of the part is left
exactly as it was, so formatting, tables, images, tracked changes, comments
and layout all survive untouched.

Word often splits one visible word across several <w:t> runs (a spelling
mark, a change in formatting, an edit made later), so "<Patient 7>" can be
stored as "<Pat" + "ient 7>". Runs are therefore joined into SEGMENTS --
consecutive text that sits in the same paragraph with nothing between it
but formatting -- and placeholders are found in the joined text. The
replacement is written into the run where the placeholder starts (keeping
that run's formatting) and the rest of the placeholder is removed from the
following runs. Anything that breaks a line or paragraph (a paragraph end,
tab, line break, text box, field code, drawing, footnote mark...) ends a
segment, so a replacement can never reach across it.

Parts processed: the body, headers, footers, footnotes, endnotes and
comments (<w:t>), text in charts, SmartArt and drawing shapes (<a:t>), chart
category/value labels (<c:v>), and the document's title/subject/keywords
properties. Deleted text in tracked changes (<w:delText>) is left alone.
NOT changed: the spreadsheet data embedded behind a chart (opening "Edit
Data" in Word shows that copy) and pictures of text.

--------------------------------------------------------------------------
SELF-CHECK
--------------------------------------------------------------------------
verify_unredaction() re-opens the saved file and confirms (1) it is a
readable .docx whose edited XML parts are all still well-formed, and (2)
none of the placeholders that were meant to be replaced is still present.
Either failure deletes the output (main_gui.py's existing hard-block path).
Placeholders that were deliberately left (unknown to every key file, or
left by the user's choice) are NOT failures -- they're reported.

It also lists "bracketless" mentions -- e.g. "Patient 7" written without
the < > brackets, which an AI-drafted or retyped text often does. Those are
never changed automatically (the bare words could be legitimate text); the
completion message points them out to check by hand.
"""
import csv
import json
import os
import re
import zipfile
from xml.etree import ElementTree

# Same placeholder shape the other engines use (<Label N>), unanchored --
# in a Word document a placeholder is always inside running text.
_PLACEHOLDER_RE = re.compile(r"<[^<>]+ \d+>")
_PLACEHOLDER_SHAPE_RE = re.compile(r"^<[^<>]+ \d+>$")

# Kept in step with excel_spreadsheet.COMBINED_COLUMN_PREFIX (this module
# does not import it, matching the other engines' rule of not depending
# across that boundary).
_COMBINED_PREFIX = "Combined: "
_COMBINED_SEPARATOR = " + "

# Which parts of the .docx package hold document text, and how.
_WORD_TEXT_PART_RE = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$"
)
_DRAWING_TEXT_PART_RE = re.compile(r"^word/(charts/chart\d*|diagrams/data\d*|diagrams/drawing\d*)\.xml$")
_PROPERTIES_PART = "docProps/core.xml"

# <w:t>text</w:t> (with or without attributes). The lookahead stops it
# matching <w:tab/>, <w:tbl>, <w:tc>, <w:tr> etc.
_W_T_RE = re.compile(r"<w:t(?=[\s>])([^>]*)>([^<]*)</w:t>")
_A_T_RE = re.compile(r"<a:t(?=[\s>])([^>]*)>([^<]*)</a:t>")
_C_V_RE = re.compile(r"<c:v(?=[\s>])([^>]*)>([^<]*)</c:v>")

# Anything between two text runs that means they are NOT one continuous
# piece of text (see module docstring). Pure formatting (<w:rPr>, </w:r>,
# <w:r>, bookmarks, proofing marks) is fine to join across.
_W_BOUNDARY_RE = re.compile(
    r"<(?:/?w:p[\s>/]|w:tab[\s>/]|w:ptab[\s>/]|w:br[\s>/]|w:cr[\s>/]|/?w:txbxContent[\s>]|"
    r"w:drawing[\s>]|w:pict[\s>]|w:object[\s>]|w:sym[\s>/]|w:fldChar[\s>/]|w:instrText[\s>]|"
    r"w:delText[\s>]|w:footnoteReference[\s>/]|w:endnoteReference[\s>/]|"
    r"w:commentReference[\s>/]|w:noBreakHyphen[\s>/]|w:softHyphen[\s>/]|/?w:tc[\s>]|/?w:tr[\s>])"
)
_A_BOUNDARY_RE = re.compile(r"<(?:/?a:p[\s>/]|a:br[\s>/]|/?a:tc[\s>]|/?a:tr[\s>])")

_ENTITY_RE = re.compile(r"&(#x[0-9A-Fa-f]+|#\d+|lt|gt|amp|quot|apos);")
_NAMED = {"lt": "<", "gt": ">", "amp": "&", "quot": '"', "apos": "'"}


def _unescape(text: str) -> str:
    def sub(m):
        name = m.group(1)
        if name in _NAMED:
            return _NAMED[name]
        if name.startswith("#x"):
            return chr(int(name[2:], 16))
        return chr(int(name[1:]))
    return _ENTITY_RE.sub(sub, text)


def _escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# --------------------------------------------------------------------------
# Key files
# --------------------------------------------------------------------------
def _render_combined(column: str, encoded: str) -> str:
    """'Combined: Last Name + First Name', '["Smith", "John"]' -> 'Smith, John'."""
    parts = column[len(_COMBINED_PREFIX):].split(_COMBINED_SEPARATOR)
    try:
        values = [str(v) for v in json.loads(encoded)]
    except (ValueError, TypeError):
        return encoded
    present = [v.strip() for v in values if str(v).strip()]
    if not present:
        return ""
    first_header = parts[0].lower() if parts else ""
    if len(present) >= 2 and values and values[0].strip() and ("last" in first_header or "surname" in first_header):
        return f"{present[0]}, {' '.join(present[1:])}"
    return " ".join(present)


def read_key_csv(key_path: str) -> dict:
    """Reads one key CSV of either format (see module docstring) and
    returns {placeholder: [{"value": str, "source": str}, ...]} -- usually
    one candidate per placeholder; more than one when the same placeholder
    text was used for different columns. Raises ValueError if the file is
    neither format."""
    with open(key_path, "r", newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    name = os.path.basename(key_path)

    mapping = {}

    def add(placeholder, value, source):
        cands = mapping.setdefault(placeholder, [])
        if not any(c["value"] == value for c in cands):
            cands.append({"value": value, "source": source})

    # Spreadsheet format ("Column, Placeholder, Real Value")
    table = None
    for row in rows:
        if not row:
            continue
        if table is None:
            if len(row) >= 3 and [c.strip() for c in row[:3]] == ["Column", "Placeholder", "Real Value"]:
                table = []
            continue
        if len(row) >= 3 and _PLACEHOLDER_SHAPE_RE.match(row[1].strip()):
            table.append((row[0], row[1].strip(), row[2]))
    if table is not None:
        # A combination's per-part rows are covered by its summary row --
        # put back as one name, not as "John" or "Smith" alone.
        combined_parts = set()
        for column, placeholder, value in table:
            if column.startswith(_COMBINED_PREFIX):
                for part in column[len(_COMBINED_PREFIX):].split(_COMBINED_SEPARATOR):
                    combined_parts.add((part, placeholder))
        for column, placeholder, value in table:
            if column.startswith(_COMBINED_PREFIX):
                add(placeholder, _render_combined(column, value), f"{column} ({name})")
            elif (column, placeholder) not in combined_parts:
                add(placeholder, value, f"column '{column}' ({name})")
        if mapping:
            return mapping

    # PDF name key format ("Placeholder, Real Name, ...")
    header_seen = False
    for row in rows:
        if not row:
            continue
        if not header_seen:
            if len(row) >= 2 and row[0].strip() == "Placeholder" and row[1].strip() == "Real Name":
                header_seen = True
            continue
        if len(row) >= 2 and _PLACEHOLDER_SHAPE_RE.match(row[0].strip()):
            add(row[0].strip(), row[1], name)
    if mapping:
        return mapping

    raise ValueError(
        "This doesn't look like a redaction key CSV this app recognizes -- neither a "
        "'Column, Placeholder, Real Value' table (a *_REDACTION_KEY.csv from a redacted "
        "workbook) nor a 'Placeholder, Real Name' table (a *_NAME_KEY.csv from a redacted "
        "PDF) was found in it."
    )


def merge_key_mappings(mappings: list):
    """Combines several read_key_csv() results. Returns (merged, conflicts)
    with conflicts always empty: for a Word document a disagreement between
    key files is not an error to stop on (a PDF key and a spreadsheet key
    can legitimately spell the same linked patient differently) -- each
    disagreement becomes a choice the user makes before anything is
    written (see placeholders_needing_choice())."""
    merged = {}
    for mapping in mappings:
        for placeholder, cands in mapping.items():
            dest = merged.setdefault(placeholder, [])
            for c in cands:
                if not any(d["value"] == c["value"] for d in dest):
                    dest.append(dict(c))
    return merged, []


def resolve_mapping(mapping: dict, choices: dict = None) -> dict:
    """{placeholder: real value} for every placeholder that can be put
    back: one with a single candidate value, or one the user picked a value
    for in `choices` ({placeholder: value, or None to leave it}). A
    placeholder with several candidates and no pick is left out (stays as
    a placeholder in the document)."""
    choices = choices or {}
    resolved = {}
    for placeholder, cands in mapping.items():
        if placeholder in choices:
            if choices[placeholder] is not None:
                resolved[placeholder] = choices[placeholder]
        elif len(cands) == 1:
            resolved[placeholder] = cands[0]["value"]
    return resolved


# --------------------------------------------------------------------------
# Reading the document
# --------------------------------------------------------------------------
def _text_parts(names):
    for name in names:
        if _WORD_TEXT_PART_RE.match(name):
            yield name, "w"
        elif _DRAWING_TEXT_PART_RE.match(name):
            yield name, "a"
        elif name == _PROPERTIES_PART:
            yield name, "props"


def _segments(xml: str, kind: str):
    """Groups a part's text runs into segments of continuous text (see
    module docstring). Yields lists of (match, unescaped_text) per segment.
    For charts, <c:v> values are each their own segment."""
    if kind == "props":
        return
    text_re, boundary_re = (_W_T_RE, _W_BOUNDARY_RE) if kind == "w" else (_A_T_RE, _A_BOUNDARY_RE)
    segment = []
    last_end = None
    for m in text_re.finditer(xml):
        if segment and boundary_re.search(xml, last_end, m.start()):
            yield segment
            segment = []
        segment.append((m, _unescape(m.group(2))))
        last_end = m.end()
    if segment:
        yield segment
    if kind == "a":
        for m in _C_V_RE.finditer(xml):
            yield [(m, _unescape(m.group(2)))]


def _document_texts(path: str):
    """(part name, kind, xml, [segment joined texts]) for every text part."""
    with zipfile.ZipFile(path) as z:
        for name, kind in _text_parts(z.namelist()):
            xml = z.read(name).decode("utf-8")
            if kind == "props":
                texts = [_unescape(m.group(0)) for m in re.finditer(r">[^<]+<", xml)]
            else:
                texts = ["".join(t for _m, t in seg) for seg in _segments(xml, kind)]
            yield name, kind, xml, texts


def find_placeholders(path: str) -> dict:
    """{placeholder: number of occurrences} for every placeholder-shaped
    text in the document. Read-only."""
    counts = {}
    for _name, _kind, _xml, texts in _document_texts(path):
        for text in texts:
            for m in _PLACEHOLDER_RE.finditer(text):
                counts[m.group(0)] = counts.get(m.group(0), 0) + 1
    return counts


def placeholders_needing_choice(path: str, mapping: dict) -> list:
    """The placeholders PRESENT in this document that the key file(s)
    give more than one value for -- [{"placeholder", "count",
    "candidates": [{"value", "source"}, ...]}, ...], in document-text
    order of first appearance. Empty when nothing needs a decision."""
    counts = find_placeholders(path)
    out = []
    for placeholder, count in counts.items():
        cands = mapping.get(placeholder) or []
        if len(cands) > 1:
            out.append({"placeholder": placeholder, "count": count, "candidates": cands})
    return out


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------
def _rewrite_segment(segment, resolved, counters):
    """Returns [(match, new_text or None)] -- None where a run is unchanged."""
    joined = "".join(t for _m, t in segment)
    hits = []
    for m in _PLACEHOLDER_RE.finditer(joined):
        value = resolved.get(m.group(0))
        if value is None:
            continue
        hits.append((m.start(), m.end(), value))
        counters[m.group(0)] = counters.get(m.group(0), 0) + 1
    if not hits:
        return None
    edits = []
    pos_in_joined = 0
    for m, text in segment:
        a, b = pos_in_joined, pos_in_joined + len(text)
        pieces = []
        pos = a
        for hs, he, value in hits:
            if he <= a or hs >= b:
                continue
            s, e = max(hs, a), min(he, b)
            if s > pos:
                pieces.append(joined[pos:s])
            if hs >= a:
                pieces.append(value)
            pos = e
        if pos < b:
            pieces.append(joined[pos:b])
        new = "".join(pieces)
        edits.append((m, new if new != text else None))
        pos_in_joined = b
    return edits


def _apply_edits(xml: str, edits):
    """Rewrites the text of the given runs inside the original XML string,
    leaving every other byte as it was."""
    out = []
    last = 0
    for m, new_text in sorted(edits, key=lambda e: e[0].start()):
        if new_text is None:
            continue
        attrs = m.group(1)
        if (new_text != new_text.strip() or not new_text) and "xml:space" not in attrs and m.group(0).startswith("<w:t"):
            attrs = attrs + ' xml:space="preserve"'
        tag = m.group(0)[1:4]  # "w:t", "a:t" or "c:v"
        out.append(xml[last:m.start()])
        out.append(f"<{tag}{attrs}>{_escape(new_text)}</{tag}>")
        last = m.end()
    out.append(xml[last:])
    return "".join(out)


def _rewrite_part(xml: str, kind: str, resolved: dict, counters: dict) -> str:
    if kind == "props":
        def sub(m):
            text = _unescape(m.group(1))
            new = _PLACEHOLDER_RE.sub(
                lambda p: (counters.__setitem__(p.group(0), counters.get(p.group(0), 0) + 1)
                           or resolved[p.group(0)]) if p.group(0) in resolved else p.group(0),
                text,
            )
            return ">" + (_escape(new) if new != text else m.group(1)) + "<"
        return re.sub(r">([^<]+)<", sub, xml)
    edits = []
    for segment in _segments(xml, kind):
        seg_edits = _rewrite_segment(segment, resolved, counters)
        if seg_edits:
            edits.extend(seg_edits)
    return _apply_edits(xml, edits) if edits else xml


def apply_unredaction(input_path, output_path, resolved, progress_callback=None, phase_callback=None):
    """Writes a copy of the .docx at input_path to output_path with every
    placeholder in `resolved` ({placeholder: real value}, from
    resolve_mapping()) replaced. Everything else in the package is copied
    byte for byte.

    Returns {'parts_changed': int, 'replacements': int,
    'placeholders_not_found': [placeholders in `resolved` the document
    never mentions], 'left_as_placeholders': {placeholder: count} for
    placeholder-shaped text still in the document (unknown to every key,
    or left by choice), 'bracketless': [{'text', 'count'}] (see module
    docstring), 'ambiguous': []}.
    """
    counters = {}
    parts_changed = 0
    with zipfile.ZipFile(input_path) as zin:
        infos = zin.infolist()
        text_parts = dict(_text_parts([i.filename for i in infos]))
        total = max(len(text_parts), 1)
        done = 0
        tmp_path = output_path + ".tmp"
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in infos:
                data = zin.read(info.filename)
                kind = text_parts.get(info.filename)
                if kind is not None:
                    xml = data.decode("utf-8")
                    new_xml = _rewrite_part(xml, kind, resolved, counters)
                    if new_xml != xml:
                        parts_changed += 1
                        data = new_xml.encode("utf-8")
                    done += 1
                    if progress_callback is not None:
                        try:
                            progress_callback(done, total)
                        except Exception:  # noqa: BLE001
                            pass
                zout.writestr(info, data)
    if phase_callback is not None:
        try:
            phase_callback("saving")
        except Exception:  # noqa: BLE001
            pass
    os.replace(tmp_path, output_path)

    remaining = find_placeholders(output_path)
    labels = sorted({p[1:p.rindex(" ")] for p in resolved} | {p[1:p.rindex(" ")] for p in remaining})
    return {
        "parts_changed": parts_changed,
        "replacements": sum(counters.values()),
        "placeholders_not_found": sorted(p for p in resolved if p not in counters),
        "left_as_placeholders": remaining,
        "bracketless": find_bracketless(output_path, labels),
        "ambiguous": [],
    }


def find_bracketless(path: str, labels) -> list:
    """Mentions like "Patient 7" with no < > around them, for the given
    placeholder labels -- [{'text': 'Patient 7', 'count': n}], most frequent
    first. Read-only; never changed automatically."""
    if not labels:
        return []
    pattern = re.compile(
        r"(?<![<\w])(" + "|".join(re.escape(l) for l in sorted(labels, key=len, reverse=True)) + r") (\d+)(?![\d>])"
    )
    counts = {}
    for _name, _kind, _xml, texts in _document_texts(path):
        for text in texts:
            for m in pattern.finditer(text):
                key = f"{m.group(1)} {m.group(2)}"
                counts[key] = counts.get(key, 0) + 1
    return [{"text": k, "count": v} for k, v in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))]


def verify_unredaction(output_path: str, resolved: dict) -> dict:
    """Independent post-check (see module docstring). Returns {'ok',
    'issues', 'leftover'} in the same shape the other engines use, so
    main_gui.py's existing failure handling applies unchanged."""
    issues = []
    leftover = []
    try:
        with zipfile.ZipFile(output_path) as z:
            bad = z.testzip()
            if bad:
                issues.append(f"The saved file is damaged (part {bad} could not be read back).")
            for name, kind in _text_parts(z.namelist()):
                try:
                    ElementTree.fromstring(z.read(name))
                except ElementTree.ParseError as exc:
                    issues.append(f"Part {name} is no longer valid XML after editing ({exc}).")
    except zipfile.BadZipFile as exc:
        return {"ok": False, "issues": [f"The saved file is not a readable .docx ({exc})."], "leftover": []}

    if not issues:
        remaining = find_placeholders(output_path)
        for placeholder, count in sorted(remaining.items()):
            if placeholder in resolved:
                leftover.append({"placeholder": placeholder, "known": True})
        if leftover:
            shown = ", ".join(e["placeholder"] for e in leftover[:10])
            more = f" (+{len(leftover) - 10} more)" if len(leftover) > 10 else ""
            issues.append(
                f"Placeholder(s) that should have been replaced are still in the document: {shown}{more}"
            )
    return {"ok": not issues, "issues": issues, "leftover": leftover}
