"""
Generic un-redaction for spreadsheet-type document plugins (currently just
document_types/excel_spreadsheet.py, but written to work for any future
plugin using the same "Column,Placeholder,Real Value" key-file convention
-- see excel_spreadsheet.py's write_key()).

This is the spreadsheet-workbook sibling of unredact_engine.py (which
handles PDF-derived redactions). They are deliberately separate modules,
not one module branching internally, for the same reason excel_spreadsheet
is a different FORMAT_KIND rather than being squeezed into the pdf_report
plugin contract: the underlying operation is genuinely different (swap a
whole cell's value back, vs. redraw text into a PDF content stream), so
sharing a module would mean one or the other growing an awkward branch for
no real benefit. main_gui.py picks between this module and
unredact_engine.py by the redacted file's extension (.xlsx/.csv here, .pdf
there).

Key mapping shape: {column_name: {placeholder_text: real_value}}. Nested
by column (rather than one flat {placeholder: value} dict, as the PDF
engine uses) because column redaction numbering restarts at 1 in every
column (see excel_spreadsheet.py's module docstring) -- if two different
columns ever ended up with the same placeholder text (e.g. the user typed
the same custom label for two columns), a flat mapping could not tell
those two placeholders apart, but a per-column mapping always can PROVIDED
we still know which column a given cell is in.

--------------------------------------------------------------------------
HANDLING A FILE THAT'S DRIFTED SINCE IT WAS REDACTED
--------------------------------------------------------------------------
In practice, un-redaction is often attempted well after redaction, on a
file that's been through other hands and other software in between: it
may have been re-saved in a different format (.xlsx -> .csv is the common
one; this module handles both), or the SAME .xlsx may have had columns
renamed, reordered, added, or removed, or a sheet renamed, since it was
originally redacted. Row order/count changes were never a problem (values
are matched by content, not position). Column drift used to be: the
per-sheet header-row match in apply_unredaction() would simply fail to
find a renamed column, silently leaving that column's placeholders
un-replaced -- caught by verify_unredaction()'s workbook-wide scan at the
end (so nothing was ever silently WRONG), but with no way to actually fix
it short of re-running with a hand-edited file.

apply_unredaction() now has a two-tier strategy:
  1. Fast path: match columns by their header text, exactly as before.
     This is what disambiguates a placeholder that was reused across two
     different columns (see key mapping shape, above) -- if the header
     still matches, we know for certain which column's value belongs
     here.
  2. Fallback: for any cell whose value is still an exact, known
     placeholder text but which the fast path didn't already resolve
     (its column's header didn't match anything -- renamed, reordered
     into a spot with no recognized header, etc.), look up that
     placeholder text across the WHOLE key file. If it's unique --
     belongs to only one (column, value) pair anywhere in the key --
     it's safe to resolve directly, regardless of which column or sheet
     the cell now sits in. If it's ambiguous (the same placeholder text
     was used for more than one column), it can't be safely guessed;
     that occurrence is reported in the 'ambiguous' field for a human to
     resolve (see apply_ambiguous_decisions()) rather than silently
     picked.
This means a renamed/reordered/moved column now resolves itself
automatically in the overwhelmingly common case (placeholder text unique
across the file, e.g. the default "<Patient 3>" labels always are), and
only asks a human when it's genuinely unable to tell.

--------------------------------------------------------------------------
MULTIPLE KEY FILES
--------------------------------------------------------------------------
A single redacted workbook may have been redacted in more than one pass
over time (a follow-up redaction covering additional columns, say), each
producing its own separate key CSV. main_gui.py lets the user pick more
than one key file in the Un-redact tab and combines them with
merge_key_mappings() below before doing anything else -- see that
function's docstring for how a disagreement between files is handled.

--------------------------------------------------------------------------
A GENERIC/FLAT KEY, FOR A FILE THAT ISN'T ONE OF THIS APP'S OWN REDACTED
WORKBOOKS AT ALL
--------------------------------------------------------------------------
David hit this directly: he had a *_NAME_KEY.csv from redacting a Dentrix
Audit Trail PDF, and wanted to un-redact an entirely different Excel
workbook -- a financial-analysis summary an assistant had produced FROM
that redacted PDF, quoting its <Patient N> placeholders in narrative text
and tables. That file was never redacted by excel_spreadsheet.py, so it
has no "Column, Placeholder, Real Value" key of its own -- the only key
that exists for it is the flat "Placeholder, Real Name[, Font, Size,
Color, Flags]" table any PDF-based document-type plugin writes (see
unredact_engine.py's module docstring: that format is already deliberately
plugin-agnostic, "the same no matter which plugin produced it"). Loading
that flat key here used to fail outright with "no 'Column, Placeholder,
Real Value' table was found in it" -- a real key file, just the wrong
shape for this module's own key format.

read_key_csv() below now also accepts that flat format. Internally it's
stored under the reserved sentinel column name _GENERIC_KEY_COLUMN, kept
completely separate from real, named columns everywhere it matters:

  - apply_unredaction() resolves real (non-generic) columns exactly as
    before (see the two-tier strategy above, exact whole-cell match).
  - A generic entry has no column to be ambiguous about (a flat key means
    exactly one thing, globally, everywhere <Patient 7> appears in the
    whole file) and no guarantee it sits ALONE in a cell rather than
    embedded in a longer sentence a summary is likely to contain -- so
    _apply_generic_unredaction() runs as a separate pass doing a plain
    text substitution of every placeholder-shaped substring, in every
    cell, of every row (there's no header row to skip -- a generic file
    has no guaranteed header semantics at all).
  - verify_unredaction()'s leftover scan looks for a placeholder shape
    ANYWHERE inside a cell's text (not just as the cell's entire value),
    so a stray un-replaced generic placeholder embedded in a sentence is
    never silently missed.
"""
import csv
import os
import re

import openpyxl
from openpyxl.utils import get_column_letter

# Used for the exact-whole-cell-match path (structured, column-keyed
# entries) -- placeholders are matched by exact-value lookup against the
# key file's placeholder set there, not by regex, since (unlike the PDF
# engine, which finds placeholders anywhere in free text) a REDACTED
# spreadsheet/CSV cell's value either exactly equals a known placeholder or
# it doesn't; that file's own redaction always put a placeholder alone in
# its own cell.
_PLACEHOLDER_SHAPE_RE = re.compile(r"^<[^<>]+ \d+>$")

# Unanchored counterpart of _PLACEHOLDER_SHAPE_RE, used wherever a
# placeholder might be sitting INSIDE a longer string rather than being the
# cell's entire value -- the generic/flat-key pass (see module docstring's
# "A GENERIC/FLAT KEY" section) and verify_unredaction()'s leftover scan,
# both of which have to work on a file this app never redacted itself and
# so can make no assumption about a placeholder ever being alone in a cell.
_PLACEHOLDER_SCAN_RE = re.compile(r"<[^<>]+ \d+>")

# Reserved key_mapping "column name" for a flat/generic key's entries (see
# module docstring) -- guaranteed to never collide with a real column
# header, which is always a non-empty string pulled straight from a
# worksheet cell.
_GENERIC_KEY_COLUMN = None

_COORD_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


def _open_workbook(path, **kwargs):
    # Same rationale and duplicated-not-imported reasoning as
    # document_types/excel_spreadsheet.py's identical helper: a workbook
    # openpyxl can't parse is almost always one not saved by genuine Excel
    # -- including a genuine legacy .xls file, which openpyxl doesn't
    # support at all and will fail to open with exactly this kind of
    # error -- and the fix (re-save a fresh copy from Excel, in .xlsx
    # format) is the same regardless of which module hit the failure or
    # why.
    try:
        return openpyxl.load_workbook(path, **kwargs)
    except Exception as exc:
        raise RuntimeError(
            f"openpyxl could not open this workbook ({type(exc).__name__}: {exc}). "
            "This usually means the file wasn't saved by genuine Microsoft Excel in "
            ".xlsx format -- either it has some structural quirk in its internal XML "
            "that Excel tolerates but openpyxl doesn't, or it's actually a legacy .xls "
            "file (openpyxl only reads .xlsx). The most reliable fix either way: open "
            "the file in Excel and use File > Save As to save a fresh copy in .xlsx "
            "format, then use that copy instead."
        ) from exc


def _csv_col_label(idx0: int) -> str:
    """0-based column index -> Excel-style column letters: 0->'A',
    25->'Z', 26->'AA', ... Used so a CSV location can be reported/addressed
    the same "letter+row number" way an .xlsx cell coordinate already is
    (e.g. 'C5'), rather than inventing a different convention for CSV."""
    n = idx0 + 1
    label = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        label = chr(65 + rem) + label
    return label


def _csv_label_to_idx(label: str) -> int:
    """Inverse of _csv_col_label(): 'A' -> 0, 'Z' -> 25, 'AA' -> 26, ..."""
    n = 0
    for ch in label.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _is_csv(path: str) -> bool:
    return os.path.splitext(path)[1].lower() == ".csv"


def _try_parse_flat_key_rows(rows: list) -> dict:
    """Parses the "Placeholder, Real Name[, Font, Size, Color, Flags]"
    table any PDF-based document-type plugin writes (see
    unredact_engine.py's read_key_csv(), which this deliberately mirrors
    the table-shape of) -- the flat, plugin-agnostic key format described
    in this module's "A GENERIC/FLAT KEY" docstring section. Only the
    Placeholder/Real Name columns matter here (no font/color to restore
    in a spreadsheet cell); returns a plain {placeholder: real_value}
    dict, or {} if `rows` doesn't contain that table either -- callers
    treat an empty result as "not this format" and fall through to their
    own error, not a hard failure here.
    """
    mapping = {}
    header = None
    for row in rows:
        if not row:
            continue
        if header is None:
            if len(row) >= 2 and row[0].strip() == "Placeholder" and row[1].strip() == "Real Name":
                header = True
            continue
        if len(row) >= 2 and _PLACEHOLDER_SHAPE_RE.match(row[0].strip()):
            mapping[row[0].strip()] = row[1]
    return mapping


def read_key_csv(key_path: str) -> dict:
    """Returns {column_name: {placeholder_text: real_value}}, where
    column_name is a real header string for an ordinary spreadsheet key,
    OR the reserved _GENERIC_KEY_COLUMN sentinel for entries loaded from a
    flat, plugin-agnostic key file (see module docstring's "A GENERIC/FLAT
    KEY" section) -- the same "Placeholder, Real Name" table format any
    PDF-based document-type plugin writes, for un-redacting a workbook
    that was never itself redacted by excel_spreadsheet.py (e.g. a
    downstream analysis file quoting placeholders copied from a redacted
    PDF).

    Skips the explanatory/confidentiality-warning rows at the top of the
    file. Tries this module's own "Column, Placeholder, Real Value" table
    shape first; if that table isn't present, falls back to the flat
    shape before giving up. Raises ValueError if the file matches NEITHER
    shape (e.g. the wrong file was picked entirely). This is always a CSV
    regardless of whether the REDACTED file being un-redacted is an .xlsx
    or a .csv -- don't confuse the two.
    """
    with open(key_path, "r", newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))

    mapping = {}
    header_seen = False
    for row in rows:
        if not row:
            continue
        if not header_seen:
            if len(row) >= 3 and [c.strip() for c in row[:3]] == ["Column", "Placeholder", "Real Value"]:
                header_seen = True
            continue
        if len(row) >= 3:
            column, placeholder, value = row[0], row[1], row[2]
            mapping.setdefault(column, {})[placeholder] = value

    if not mapping:
        flat = _try_parse_flat_key_rows(rows)
        if flat:
            return {_GENERIC_KEY_COLUMN: flat}
        raise ValueError(
            "This doesn't look like a redaction key CSV this app recognizes -- "
            "neither a 'Column, Placeholder, Real Value' table (a "
            "*_REDACTION_KEY.csv saved alongside a redacted workbook) nor a "
            "'Placeholder, Real Name' table (a *_NAME_KEY.csv saved alongside "
            "a redacted PDF, also usable here to un-redact any OTHER file -- "
            "spreadsheet or otherwise -- that quotes the same placeholders) "
            "was found in it. Make sure this is the actual key file saved "
            "alongside the redaction, not some other file."
        )
    return mapping


def merge_key_mappings(mappings: "list"):
    """Combines multiple {column: {placeholder: value}} key mappings (one
    per key CSV the user picked in the Un-redact tab -- see module
    docstring's "MULTIPLE KEY FILES" section) into a single mapping usable
    by apply_unredaction(). The common case is several key files that
    simply cover different columns (or the same column across separate
    redaction passes with no overlap) -- those combine with no fuss.

    Returns (merged, conflicts). conflicts is a list of {'column':
    str, 'placeholder': str, 'values': list[str]} entries, one per
    (column, placeholder) pair that maps to more than one DISTINCT real
    value across the given files. Two genuine key files describing the
    same underlying redaction should never disagree about what a
    placeholder means -- a conflict almost always means one of the
    selected files doesn't actually belong with the others (e.g. the
    wrong file was added by mistake, or a key file from an entirely
    different workbook). main_gui.py treats any conflict as blocking and
    shows it to the user rather than silently guessing; 'merged' still
    contains one (the first-seen) of the conflicting values for each
    conflicting entry, in case a caller ever needs to proceed anyway, but
    the normal path is to stop and let the user fix their file selection.
    """
    values_by_key = {}  # (column, placeholder) -> [distinct values, first-seen order]
    for mapping in mappings:
        for column, phmap in mapping.items():
            for ph, val in phmap.items():
                key = (column, ph)
                seen = values_by_key.setdefault(key, [])
                if val not in seen:
                    seen.append(val)

    merged = {}
    conflicts = []
    for (column, ph), values in values_by_key.items():
        merged.setdefault(column, {})[ph] = values[0]
        if len(values) > 1:
            conflicts.append({"column": column, "placeholder": ph, "values": values})
    # `column` can be the _GENERIC_KEY_COLUMN sentinel (None) mixed in
    # alongside ordinary string column names (see module docstring's "A
    # GENERIC/FLAT KEY" section) -- None and str aren't orderable against
    # each other in Python 3, so the sort key has to normalize it first.
    conflicts.sort(key=lambda c: (c["column"] or "", c["placeholder"]))
    return merged, conflicts


def _build_placeholder_index(key_mapping: dict) -> dict:
    """placeholder_text -> [(column, real_value), ...] -- every (column,
    value) pair anywhere in the key file that used this exact placeholder
    text. Length 1 means the text is unique across the whole file (safe to
    resolve automatically wherever it's found); length > 1 means it was
    reused across more than one column and needs the header-match fast
    path (or a human) to say which one applies to a given cell.

    Deliberately EXCLUDES _GENERIC_KEY_COLUMN entries (see module
    docstring's "A GENERIC/FLAT KEY" section) -- a generic entry has no
    column to ever be ambiguous about and no guarantee of sitting alone in
    a cell, so it's resolved entirely separately by
    _apply_generic_unredaction(), not folded into this column-disambiguation
    strategy at all.
    """
    index = {}
    for column, phmap in key_mapping.items():
        if column == _GENERIC_KEY_COLUMN:
            continue
        for ph, val in phmap.items():
            index.setdefault(ph, []).append((column, val))
    return index


def _substitute_generic_placeholders(text: str, generic_map: dict):
    """Shared by the .xlsx and .csv code paths: replaces every
    placeholder-shaped substring found in `text` -- not just the case
    where the ENTIRE string is one -- with its real value from the flat
    {placeholder: real_value} generic_map (the _GENERIC_KEY_COLUMN entry
    from key_mapping, if any; see module docstring's "A GENERIC/FLAT KEY"
    section), wherever that placeholder's value is known.

    A placeholder-shaped substring with NO entry in generic_map is left
    untouched rather than guessed at or errored on -- it's either a
    genuinely unrelated bit of text that happens to look like one (rare,
    but not this function's business to assume), or it belongs to some
    OTHER key file the user didn't select; either way verify_unredaction()
    below will still flag it in the output rather than silently dropping
    it.

    Returns (new_text, replaced_count, found_texts) -- found_texts is
    every distinct placeholder-shaped text actually seen in `text`
    (matched or not), used by the caller to compute
    'placeholders_not_found' the same way the structured path already
    does. Cheap no-op (returns text unchanged) when generic_map is empty
    or `text` contains no '<' at all.
    """
    found_texts = set()
    if not generic_map or "<" not in text:
        return text, 0, found_texts

    replaced = 0

    def _sub(match):
        nonlocal replaced
        found = match.group(0)
        found_texts.add(found)
        if found in generic_map:
            replaced += 1
            return generic_map[found]
        return found

    new_text = _PLACEHOLDER_SCAN_RE.sub(_sub, text)
    return new_text, replaced, found_texts


def apply_unredaction(
    input_path: str,
    output_path: str,
    key_mapping: dict,
    progress_callback=None,
    phase_callback=None,
) -> dict:
    """Writes a copy of input_path with every placeholder cell swapped
    back for its real value -- input_path may be an .xlsx workbook or a
    .csv file (same key file format either way; see read_key_csv() and,
    for combining more than one key file, merge_key_mappings()).
    progress_callback(current, total) and phase_callback(phase_name)
    behave the same as documented in unredact_engine.py's
    apply_unredaction() -- current/total count rows processed, not pages,
    but the shape is the same so the GUI can drive either engine
    identically.

    See the module docstring's "HANDLING A FILE THAT'S DRIFTED SINCE IT
    WAS REDACTED" section for the two-tier column-matching strategy.

    Returns {'sheets_changed': int, 'replacements': int,
    'placeholders_not_found': list[str], 'ambiguous': list[dict]}.
    'placeholders_not_found' is "<column>: <placeholder>" entries present
    in the key file but never found as an actual cell value ANYWHERE in
    the file, which usually means this key file doesn't actually belong
    to this file (wrong pairing). 'ambiguous' is occurrences that WERE
    found but couldn't be safely auto-resolved (see module docstring) --
    each entry is {'sheet': str, 'coordinate': str, 'row': int, 'value':
    str, 'candidates': [{'column': str, 'value': str}, ...]} -- pass these
    to apply_ambiguous_decisions() once a human has picked (or skipped)
    each one.
    """
    if _is_csv(input_path):
        return _apply_unredaction_csv(input_path, output_path, key_mapping, progress_callback, phase_callback)
    return _apply_unredaction_xlsx(input_path, output_path, key_mapping, progress_callback, phase_callback)


def _apply_unredaction_xlsx(input_path, output_path, key_mapping, progress_callback, phase_callback):
    placeholder_index = _build_placeholder_index(key_mapping)
    wb = _open_workbook(input_path)
    try:
        sheet_col_map = {}
        # Every column's CURRENT header text, regardless of whether it's
        # one this key file recognizes -- unlike col_idx_by_name (below),
        # which only keeps the ones that matter for the fast path. This is
        # what lets an ambiguous occurrence's location be described by
        # its column's actual (if renamed/unrecognized) current label,
        # e.g. "sitting in column 'Comments'", instead of just a bare
        # cell reference.
        sheet_raw_headers = {}
        for ws in wb.worksheets:
            col_idx_by_name = {}
            raw_headers = {}
            try:
                header_row = next(ws.iter_rows(min_row=1, max_row=1))
            except StopIteration:
                continue
            for cell in header_row:
                if cell.value is None:
                    continue
                text = str(cell.value).strip()
                if text:
                    raw_headers[cell.column] = text
                if text in key_mapping:
                    col_idx_by_name[cell.column] = text
            if col_idx_by_name:
                sheet_col_map[ws.title] = col_idx_by_name
            sheet_raw_headers[ws.title] = raw_headers

        # Every sheet is scanned now (not just ones with a matched
        # header), since the fallback pass needs to see cells sitting
        # under an unrecognized/renamed header too.
        total_rows = sum(max(ws.max_row - 1, 0) for ws in wb.worksheets)
        current = 0
        replacements = 0
        sheets_changed = set()
        seen_placeholder_texts = set()
        ambiguous = []

        for ws in wb.worksheets:
            col_idx_by_name = sheet_col_map.get(ws.title, {})
            raw_headers = sheet_raw_headers.get(ws.title, {})
            changed_this_sheet = False
            for row in ws.iter_rows(min_row=2):
                current += 1
                for cell in row:
                    if cell.value is None:
                        continue
                    val_str = str(cell.value)
                    candidates = placeholder_index.get(val_str)
                    if candidates is None:
                        continue  # not a known placeholder at all -- ordinary data
                    seen_placeholder_texts.add(val_str)

                    col_name = col_idx_by_name.get(cell.column)
                    if col_name is not None and val_str in key_mapping.get(col_name, {}):
                        # Fast path: this cell's column still has a
                        # recognized header, so we know for certain which
                        # (column, value) pair applies here even if the
                        # placeholder text is reused elsewhere.
                        cell.value = key_mapping[col_name][val_str]
                        replacements += 1
                        changed_this_sheet = True
                        continue

                    if len(candidates) == 1:
                        # Fallback: header didn't match (column renamed,
                        # reordered under an unrecognized label, etc.),
                        # but this exact placeholder text is unique across
                        # the whole key file -- safe to resolve regardless
                        # of which column/sheet it's actually sitting in.
                        _, real_value = candidates[0]
                        cell.value = real_value
                        replacements += 1
                        changed_this_sheet = True
                    else:
                        # Genuinely ambiguous: reused placeholder text,
                        # AND we've lost the header match that would have
                        # told us which column's value belongs here.
                        row_context = []
                        for other in row:
                            if other.column == cell.column or other.value in (None, ""):
                                continue
                            label = raw_headers.get(other.column) or f"column {get_column_letter(other.column)}"
                            row_context.append({"label": label, "value": str(other.value)})
                            if len(row_context) >= 4:
                                break
                        ambiguous.append({
                            "sheet": ws.title, "coordinate": cell.coordinate, "row": cell.row,
                            "value": val_str,
                            "column_header": raw_headers.get(cell.column),
                            "row_context": row_context,
                            "candidates": [{"column": c, "value": v} for c, v in candidates],
                        })
                if progress_callback is not None:
                    try:
                        progress_callback(current, max(total_rows, 1))
                    except Exception:
                        pass
            if changed_this_sheet:
                sheets_changed.add(ws.title)

        # Generic/flat-key pass (see module docstring's "A GENERIC/FLAT
        # KEY" section) -- runs over EVERY row of EVERY sheet, including
        # row 1, since a file this app never redacted itself has no
        # guaranteed header row to skip, and matches placeholder-shaped
        # SUBSTRINGS rather than requiring a cell's entire value to be
        # one, since a downstream summary is likely to embed a placeholder
        # in a longer sentence rather than isolate it in its own cell.
        generic_map = key_mapping.get(_GENERIC_KEY_COLUMN, {})
        if generic_map:
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue
                        new_val, count, found = _substitute_generic_placeholders(str(cell.value), generic_map)
                        seen_placeholder_texts |= found
                        if count:
                            cell.value = new_val
                            replacements += count
                            sheets_changed.add(ws.title)

        if phase_callback is not None:
            try:
                phase_callback("saving")
            except Exception:
                pass
        wb.save(output_path)

        not_found = sorted(
            f"{col if col is not None else '(no column)'}: {ph}"
            for col, phmap in key_mapping.items()
            for ph in phmap
            if ph not in seen_placeholder_texts
        )
        return {
            "sheets_changed": len(sheets_changed),
            "replacements": replacements,
            "placeholders_not_found": not_found,
            "ambiguous": ambiguous,
        }
    finally:
        wb.close()


def _apply_unredaction_csv(input_path, output_path, key_mapping, progress_callback, phase_callback):
    placeholder_index = _build_placeholder_index(key_mapping)
    try:
        with open(input_path, "r", newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
    except OSError as exc:
        raise RuntimeError(f"Could not read this CSV file: {exc}") from exc
    if not rows:
        raise RuntimeError("This CSV file is empty -- nothing to un-redact.")

    pseudo_sheet = os.path.basename(input_path)  # only "sheet" a CSV has, for display/addressing
    header = rows[0]
    col_idx_by_name = {}
    # Every column's CURRENT header text, regardless of whether it's one
    # this key file recognizes -- same purpose as sheet_raw_headers in
    # _apply_unredaction_xlsx: lets an ambiguous occurrence's location be
    # described by its column's actual (if renamed/unrecognized) label.
    raw_header_by_col_idx = {}
    for idx, text in enumerate(header):
        text = text.strip()
        if text:
            raw_header_by_col_idx[idx] = text
        if text in key_mapping:
            col_idx_by_name[idx] = text

    total = max(len(rows) - 1, 0)
    replacements = 0
    sheets_changed = False
    seen_placeholder_texts = set()
    ambiguous = []

    for row_i, row in enumerate(rows[1:], start=2):  # 1-based row numbers; row 1 is the header
        for col_i, val in enumerate(row):
            if val == "":
                continue
            candidates = placeholder_index.get(val)
            if candidates is None:
                continue
            seen_placeholder_texts.add(val)

            col_name = col_idx_by_name.get(col_i)
            if col_name is not None and val in key_mapping.get(col_name, {}):
                row[col_i] = key_mapping[col_name][val]
                replacements += 1
                sheets_changed = True
                continue

            if len(candidates) == 1:
                _, real_value = candidates[0]
                row[col_i] = real_value
                replacements += 1
                sheets_changed = True
            else:
                row_context = []
                for other_i, other_val in enumerate(row):
                    if other_i == col_i or other_val in (None, ""):
                        continue
                    label = raw_header_by_col_idx.get(other_i) or f"column {_csv_col_label(other_i)}"
                    row_context.append({"label": label, "value": str(other_val)})
                    if len(row_context) >= 4:
                        break
                ambiguous.append({
                    "sheet": pseudo_sheet, "coordinate": f"{_csv_col_label(col_i)}{row_i}", "row": row_i,
                    "value": val,
                    "column_header": raw_header_by_col_idx.get(col_i),
                    "row_context": row_context,
                    "candidates": [{"column": c, "value": v} for c, v in candidates],
                })
        if progress_callback is not None:
            try:
                progress_callback(row_i - 1, max(total, 1))
            except Exception:
                pass

    # Generic/flat-key pass (see module docstring's "A GENERIC/FLAT KEY"
    # section and _apply_unredaction_xlsx's identical pass) -- runs over
    # EVERY row INCLUDING row 0 (the header row above treats as off-limits
    # to the structured path), since a generic file has no guaranteed
    # header semantics at all, and matches placeholder-shaped substrings
    # rather than requiring a cell's entire value to be one.
    generic_map = key_mapping.get(_GENERIC_KEY_COLUMN, {})
    if generic_map:
        for row in rows:
            for col_i, val in enumerate(row):
                if val == "":
                    continue
                new_val, count, found = _substitute_generic_placeholders(val, generic_map)
                seen_placeholder_texts |= found
                if count:
                    row[col_i] = new_val
                    replacements += count
                    sheets_changed = True

    if phase_callback is not None:
        try:
            phase_callback("saving")
        except Exception:
            pass
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)

    not_found = sorted(
        f"{col if col is not None else '(no column)'}: {ph}"
        for col, phmap in key_mapping.items()
        for ph in phmap
        if ph not in seen_placeholder_texts
    )
    return {
        "sheets_changed": 1 if sheets_changed else 0,
        "replacements": replacements,
        "placeholders_not_found": not_found,
        "ambiguous": ambiguous,
    }


def apply_ambiguous_decisions(output_path: str, decisions: "list") -> int:
    """Applies manual-review decisions for AMBIGUOUS placeholders (see
    apply_unredaction()'s 'ambiguous' field) directly to output_path.

    decisions: [{'sheet': str, 'coordinate': str, 'action': 'replace' or
    'ignore', 'value': str (required for 'replace' -- the real value the
    user picked from that occurrence's candidate list)}, ...]

    Works for both an .xlsx output ('sheet' is the real worksheet name,
    'coordinate' an openpyxl cell coordinate like 'C5') and a .csv output
    ('sheet' is the pseudo-name used throughout this module for CSV -- the
    file's own base filename -- and 'coordinate' is the same
    letter+row-number convention _csv_col_label() produces). Returns the
    number of cells actually changed.
    """
    to_replace = {(d["sheet"], d["coordinate"]): d["value"] for d in decisions if d.get("action") == "replace"}
    if not to_replace:
        return 0

    if _is_csv(output_path):
        with open(output_path, "r", newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        pseudo_sheet = os.path.basename(output_path)
        changed = 0
        for (sheet, coordinate), value in to_replace.items():
            if sheet != pseudo_sheet:
                continue
            m = _COORD_RE.match(coordinate)
            if not m:
                continue
            col_idx = _csv_label_to_idx(m.group(1))
            row_idx = int(m.group(2)) - 1  # coordinate row numbers are 1-based
            if 0 <= row_idx < len(rows) and col_idx < len(rows[row_idx]):
                rows[row_idx][col_idx] = value
                changed += 1
        with open(output_path, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(rows)
        return changed

    changed = 0
    wb = _open_workbook(output_path)
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    key = (ws.title, cell.coordinate)
                    if key in to_replace:
                        cell.value = to_replace[key]
                        changed += 1
        wb.save(output_path)
    finally:
        wb.close()
    return changed


def verify_unredaction(output_path: str, key_mapping: dict) -> dict:
    """Independent post-check, same spirit as unredact_engine.py's
    verify_unredaction(): none of the placeholders should still be present
    as a cell value anywhere in the output file -- every one that was
    found should have been fully swapped for its real value, nothing left
    half-done. Checked file-wide (not just within the column it started
    in), same reasoning as excel_spreadsheet.py's verify_column_redaction()
    leak check. Works for either an .xlsx or a .csv output file.

    Flags a cell for TWO different reasons, because they point to two
    different problems:
      1. Its value exactly equals a placeholder KNOWN from the loaded key
         file(s) -- the original check. Usually means the key file simply
         doesn't cover this file (wrong pairing, or a follow-up redaction
         pass produced a second key file that wasn't loaded too).
      2. Its value merely LOOKS like a placeholder (matches the same
         "<Label N>" shape every plugin uses) even though it doesn't match
         anything in any loaded key file. This exists to close a real
         blind spot: apply_unredaction() only ever replaces a cell it can
         match against the key, so if a placeholder's exact text doesn't
         match its own key file's entry for it -- for instance the wrong
         number got recorded somewhere for that identity, even by a
         single digit -- the cell is left untouched AND, without this
         check, verify_unredaction() would have no way to notice either,
         since it was never "known" to begin with. A genuine business
         value practically never happens to look like "<Word 123>", so
         this generic shape match is safe to always flag.

    Both checks scan for a placeholder shape ANYWHERE inside a cell's
    text, not just as its entire value -- necessary for a file un-redacted
    via a generic/flat key (see module docstring's "A GENERIC/FLAT KEY"
    section), where a placeholder is as likely to be embedded in a longer
    sentence as to sit alone in its own cell, so a stray, un-replaced one
    is never silently missed just because the rest of the cell's text
    isn't placeholder-shaped.

    Returns {'ok': bool, 'issues': list[str], 'leftover': list[dict]}.
    'leftover' is the full, structured, untruncated list behind 'issues'
    ({'sheet', 'coordinate', 'value'} per entry, plus 'known': bool marking
    which of the two reasons above applied) -- main_gui.py's manual review
    flow cross-checks this against apply_unredaction()'s 'ambiguous'
    entries to tell "a human already flagged this and hasn't resolved it
    yet" apart from "something else, unexplained, is still wrong" (which
    still gets the hard block).
    """
    all_placeholders = {ph for phmap in key_mapping.values() for ph in phmap}
    issues = []
    leftover = []

    def _flag(sheet, coordinate, val):
        if "<" not in val:
            return  # cheap pre-filter -- every placeholder shape needs at least one '<'
        for match in _PLACEHOLDER_SCAN_RE.finditer(val):
            ph_text = match.group(0)
            leftover.append({
                "sheet": sheet, "coordinate": coordinate, "value": ph_text,
                "known": ph_text in all_placeholders,
            })

    if _is_csv(output_path):
        with open(output_path, "r", newline="", encoding="utf-8-sig") as f:
            pseudo_sheet = os.path.basename(output_path)
            for row_i, row in enumerate(csv.reader(f), start=1):
                for col_i, val in enumerate(row):
                    if val:
                        _flag(pseudo_sheet, f"{_csv_col_label(col_i)}{row_i}", val)
    else:
        wb = openpyxl.load_workbook(output_path, read_only=True, data_only=True)
        try:
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            _flag(ws.title, cell.coordinate, str(cell.value))
        finally:
            wb.close()

    known_leftover = [e for e in leftover if e["known"]]
    unknown_leftover = [e for e in leftover if not e["known"]]
    if known_leftover:
        shown = ", ".join(f"{e['sheet']}!{e['coordinate']}: {e['value']}" for e in known_leftover[:10])
        more = f" (+{len(known_leftover) - 10} more)" if len(known_leftover) > 10 else ""
        issues.append(f"Placeholder(s) still present in output, not replaced: {shown}{more}")
    if unknown_leftover:
        shown = ", ".join(f"{e['sheet']}!{e['coordinate']}: {e['value']}" for e in unknown_leftover[:10])
        more = f" (+{len(unknown_leftover) - 10} more)" if len(unknown_leftover) > 10 else ""
        issues.append(
            "Placeholder-shaped text still present in output that isn't in any loaded key file "
            f"at all: {shown}{more}. This usually means either the key file(s) you loaded don't "
            "cover every placeholder in this file (a follow-up redaction pass produced a second "
            "key file you haven't loaded, for instance), or this exact placeholder doesn't match "
            "what any loaded key file recorded for that value -- check the key file directly for "
            "the placeholder text shown above."
        )

    return {"ok": len(issues) == 0, "issues": issues, "leftover": leftover}
