"""
Excel Spreadsheet (choose columns) -- a document-type plugin, but a
different SHAPE of one than dentrix_audit_trail.py. See base.py's
"TWO SHAPES OF PLUGIN" section for why: there's no free-text pattern to
scan for here, so this plugin sets FORMAT_KIND = "spreadsheet" and
implements a different, smaller contract -- the user picks which
column(s) to redact by name, rather than the plugin auto-detecting
matches the way the PDF-report plugins do.

--------------------------------------------------------------------------
WHAT THIS PLUGIN DOES
--------------------------------------------------------------------------
Given an .xlsx workbook:
  1. list_columns() reads every worksheet's header row (row 1) and returns
     the distinct column names found, each with a sample value, so the GUI
     can show a checkbox per column.
  2. The user checks the column(s) to redact and, if they want, edits the
     placeholder label for each (default: the column's own name).
  3. apply_column_redaction() replaces every non-blank cell in each chosen
     column with a placeholder -- <Label 1>, <Label 2>, etc -- assigning
     the SAME placeholder to the SAME underlying value every time it
     recurs in that column (so, e.g., every row for the same SSN gets the
     same placeholder), independently per column (numbering restarts at 1
     in each column; a placeholder is only ever compared for reuse within
     the column it came from, never across columns).
  4. write_key() saves a CSV recording every (column, placeholder, real
     value) triple, so the Un-redact tab can reverse it later -- read by
     spreadsheet_unredact_engine.py, a sibling to unredact_engine.py (the
     PDF one). main_gui.py picks between them by the redacted file's
     extension.

A column is matched by its header text (row 1), independently per
worksheet -- a workbook with several sheets that all have (for instance)
a "Patient Name" column will have that column redacted on every sheet
that has it, not just the first. A sheet with none of the chosen column
names is left completely untouched.

--------------------------------------------------------------------------
WHY THIS COUNTS AS TRUE REDACTION
--------------------------------------------------------------------------
Cell values are set directly (cell.value = placeholder_text) and the
workbook is then written out fresh via openpyxl -- the original value
is not hidden behind formatting or left in the file's shared-string table
for someone to recover; it is simply not present anywhere in the output
file's XML once saved. This is the same "genuinely removed, not covered
up" standard the PDF plugins hold themselves to, just achieved differently
because a spreadsheet cell (unlike a PDF content stream run) has no
in-place partial-overwrite ambiguity to begin with -- replacing a cell's
value replaces all of it, there's no analog to the tightly-spaced-PDF-line
problem that motivated the whole-page-rebuild strategy used elsewhere.

--------------------------------------------------------------------------
KNOWN LIMITATIONS
--------------------------------------------------------------------------
- Only modern Excel (.xlsx) workbooks are supported -- not the legacy
  .xls binary format, and not .csv (a .csv has no concept of multiple
  named sheets or the header-row detection this plugin relies on; it
  could be supported as a separate, simpler plugin later if needed).
- Only row 1 of each sheet is checked for headers. A workbook whose
  tables don't start at row 1 (title rows, merged banner cells, etc.
  above the real header) won't have its columns detected -- there is no
  "Needs review" concept here the way the PDF plugins have, since this
  plugin doesn't scan free text for candidate matches; either a column
  shows up in the checklist or it doesn't. This is a case where a
  redaction can be missed entirely, not just mis-parsed, precisely
  because it isn't found in the first place -- if a workbook's layout is
  unusual, check the Excel file's own header row before trusting the
  "Load Columns" list.
- Cell formatting (number format, fill color, conditional formatting,
  etc.) on a redacted cell is not specifically preserved or restored --
  openpyxl generally keeps a cell's existing style when only .value is
  changed, but formulas, charts, or other cells that referenced the
  original value are not rewritten and may show errors or stale results.
- Merged cells: only the top-left cell of a merged range carries a value
  in openpyxl's model, so a selected column that runs through merged
  cells will only have that anchor cell's value (and hence be reachable
  for redaction) -- the other cells in the merge have no value to find.
"""
import csv
import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List

import openpyxl
from openpyxl.utils import get_column_letter

ID = "excel_spreadsheet"
DISPLAY_NAME = "Excel Spreadsheet (choose columns)"
FORMAT_KIND = "spreadsheet"
FILE_TYPES = [("Excel files", "*.xlsx"), ("All files", "*.*")]
SUPPORTS_KEY_FILE = True
PROGRESS_UNIT_LABEL = "row"
# See base.py's "BATCH REDACTION" section and apply_column_redaction()'s
# `state` parameter below -- main_gui.py's Batch Redact tab checks this
# before offering batch mode for a document type.
SUPPORTS_BATCH = True

# The next three exist only so main_gui.py can read them unconditionally
# the same way it does for pdf_report-shaped plugins, without special
# casing every constant access by FORMAT_KIND. This plugin doesn't scan
# free text for candidate matches, so it has no real "entity" or "needs
# review" concept -- these are placeholders, not used to build any UI text
# that's actually shown for this document type.
ENTITY_LABEL = "Value"
CONTEXT_COLUMN_LABEL = ""
REVIEW_HINT = ""
NO_MATCHES_WARNING = (
    "No column headers were found in row 1 of any worksheet in this file. "
    "This plugin only looks at row 1 -- if this workbook's table starts "
    "somewhere else (a title row, merged banner cells, etc. above the real "
    "header), columns can't be auto-detected. Check the file's layout, or "
    "move the header row to row 1 and try again."
)


@dataclass
class ColumnInfo:
    """One distinct column header found across the workbook's sheet(s)."""
    name: str               # exact header text as found in row 1
    sheets: List[str]       # worksheet name(s) this header appears in
    sample: str              # first non-blank cell value seen under this header, or ""
    default_label: str       # sanitized default placeholder label, e.g. "Patient_Name"


def _sanitize_label(name: str) -> str:
    """Turns a column header into a safe placeholder label: whitespace runs
    collapse to a single underscore, anything that isn't a letter, digit,
    underscore or hyphen is dropped. Falls back to "Column" if that leaves
    nothing usable (e.g. a header that was purely punctuation/emoji)."""
    collapsed = re.sub(r"\s+", "_", name.strip())
    cleaned = re.sub(r"[^A-Za-z0-9_\-]", "", collapsed)
    return cleaned or "Column"


def _open_workbook(path: str, **kwargs):
    """openpyxl.load_workbook(), but with a much more actionable error if
    it fails. In practice, a workbook that openpyxl can't parse is almost
    always one that wasn't saved by genuine Microsoft Excel -- exported
    from a practice-management system, general ledger, or other reporting
    tool -- and has some structural quirk in its internal XML (a missing
    relationship id, an out-of-range style reference, etc.) that Excel
    itself silently tolerates but openpyxl's stricter parser doesn't. The
    single most reliable fix for this, confirmed to work across a wide
    range of these quirks, is to open the file in genuine Excel and use
    File > Save As to write a fresh copy -- Excel rewrites the internal
    XML cleanly when it does this, which clears out whatever the original
    exporting program left broken. Without the actual failing file in
    hand, guessing which specific quirk this run hit isn't reliable, so
    this re-raises with the original exception type/message preserved
    (via `raise ... from exc`) rather than trying to interpret it.
    """
    try:
        return openpyxl.load_workbook(path, **kwargs)
    except Exception as exc:
        raise RuntimeError(
            f"openpyxl could not open this workbook ({type(exc).__name__}: {exc}). "
            "This usually means the file wasn't saved by genuine Microsoft Excel (e.g. "
            "it was exported from another program) and has some structural quirk in its "
            "internal XML that Excel tolerates but openpyxl doesn't. The most reliable "
            "fix: open the file in Excel and use File > Save As to save a fresh copy in "
            ".xlsx format, then use that copy instead."
        ) from exc


def list_columns(path: str) -> List[ColumnInfo]:
    """Opens the workbook read-only and returns every distinct column
    header found in row 1 of any worksheet, in first-seen order. Does not
    modify anything.
    """
    wb = _open_workbook(path, read_only=True, data_only=True)
    try:
        columns: "Dict[str, ColumnInfo]" = {}
        for ws in wb.worksheets:
            rows = ws.iter_rows()
            try:
                header_row = next(rows)
            except StopIteration:
                continue
            header_by_idx = {}
            for cell in header_row:
                if cell.value is None:
                    continue
                text = str(cell.value).strip()
                if not text:
                    continue
                header_by_idx[cell.column] = text
                if text not in columns:
                    columns[text] = ColumnInfo(
                        name=text, sheets=[], sample="", default_label=_sanitize_label(text)
                    )
                if ws.title not in columns[text].sheets:
                    columns[text].sheets.append(ws.title)
            if not header_by_idx:
                continue
            # Grab one sample value per header, from the first data row that has one.
            still_needed = {idx for idx in header_by_idx if not columns[header_by_idx[idx]].sample}
            if not still_needed:
                continue
            for row in rows:
                if not still_needed:
                    break
                for cell in row:
                    # Check .value BEFORE .column, always, in this order --
                    # read_only mode pads a sparse row (any blank cell
                    # sitting between two filled columns, extremely common
                    # in real spreadsheets) with openpyxl's EmptyCell
                    # placeholder, which has .value (always None) but has
                    # NO .column attribute at all. Reading .column first,
                    # as an earlier version of this loop did, raised
                    # AttributeError the moment a sparse row was hit.
                    if cell.value in (None, ""):
                        continue
                    if cell.column in still_needed:
                        col = columns[header_by_idx[cell.column]]
                        col.sample = str(cell.value)
                        still_needed.discard(cell.column)
        return list(columns.values())
    finally:
        wb.close()


def distinct_values_in_column(path: str, column_name: str) -> "set[str]":
    """Opens the workbook read-only and returns every distinct non-blank
    cell value (as str) found under the given header, across every sheet
    that has it in row 1. Read-only, single-purpose scan -- does not
    modify anything and does not build the full placeholder maps that
    apply_column_redaction() does.

    Used by the Batch Redact tab's cross-type identity linking (see
    base.py's "BATCH REDACTION" section): before redacting a file whose
    column is manually linked to another document type's identity (e.g. a
    "Patient Name" column linked to the Dentrix plugin's Patient
    numbering), the caller needs every distinct value that will appear in
    that column so each one can be resolved against the shared identity
    pool ahead of time -- this is what supplies that list.
    """
    wb = _open_workbook(path, read_only=True, data_only=True)
    try:
        values = set()
        for ws in wb.worksheets:
            rows = ws.iter_rows()
            try:
                header_row = next(rows)
            except StopIteration:
                continue
            target_idx = None
            for cell in header_row:
                if cell.value is not None and str(cell.value).strip() == column_name:
                    target_idx = cell.column
                    break
            if target_idx is None:
                continue
            for row in rows:
                for cell in row:
                    # See list_columns()'s comment on this same pattern --
                    # check .value before .column, a sparse row's padding
                    # cells have no .column attribute at all.
                    if cell.value in (None, ""):
                        continue
                    if cell.column == target_idx:
                        values.add(str(cell.value))
        return values
    finally:
        wb.close()


def apply_column_redaction(
    input_path: str,
    output_path: str,
    selected_columns: List[dict],
    progress_callback=None,
    phase_callback=None,
    state=None,
) -> "tuple[dict, list]":
    """Writes a copy of input_path with every non-blank cell in each chosen
    column replaced by a placeholder. `selected_columns` is
    [{"name": <exact header text>, "label": <placeholder label to use>}, ...]
    -- `label` is what goes inside the angle brackets; the caller (the GUI)
    is responsible for defaulting it to the column's own sanitized name and
    letting the user override it.

    The same real value always gets the same placeholder within its own
    column (independent numbering per column -- see module docstring).

    progress_callback(current, total) is called with 1-based row progress
    across every sheet that has at least one chosen column, if given; any
    exception it raises is swallowed. phase_callback(phase_name) is called
    once, with "saving", right before the final write to disk.

    `state`, if given, is a running {column_name: {value: placeholder_text}}
    mapping to seed this call's numbering from, instead of starting every
    column's counter fresh at 1 -- see base.py's "BATCH REDACTION" section.
    Passing the SAME dict through a series of calls (one per file in a
    batch) makes a value that recurs across those files keep the same
    placeholder text everywhere, rather than each file numbering
    independently and risking the same placeholder text meaning two
    different values in two different files. This function does not
    mutate `state` itself (a fresh dict is built from it internally) --
    the caller is expected to fold each call's returned key_rows back into
    its own running state before the next call; see main_gui.py's batch
    redact worker for the actual accumulation loop. Omit (or pass None)
    for the original single-file behavior, unchanged.

    Returns (summary, key_rows):
      summary: {'columns_selected': int, 'columns_found': int,
                'columns_not_found': list[str], 'cells_redacted': int,
                'unique_values_redacted': int}
      key_rows: [{'column': str, 'placeholder': str, 'value': str}, ...],
                one row per unique (column, value) pair KNOWN SO FAR for a
                column found in this file -- including any carried forward
                via `state` even if this particular file didn't happen to
                contain that exact value -- in the order placeholders were
                assigned. This is what write_key() saves for a single file;
                a batch caller should accumulate across every call's
                key_rows rather than assume the last file's alone covers
                the whole batch (a column absent from the last file
                wouldn't otherwise be represented).
    """
    label_by_name = {c["name"]: c["label"] for c in selected_columns}
    wb = _open_workbook(input_path)
    try:
        # Pass 1: find which sheets have which chosen columns, and the
        # total row count across them, so progress can be reported.
        sheet_col_map = {}   # ws.title -> {col_idx: header_name}
        found_names = set()
        for ws in wb.worksheets:
            col_idx_by_name = {}
            try:
                header_row = next(ws.iter_rows(min_row=1, max_row=1))
            except StopIteration:
                continue
            for cell in header_row:
                if cell.value is None:
                    continue
                text = str(cell.value).strip()
                if text in label_by_name:
                    col_idx_by_name[cell.column] = text
                    found_names.add(text)
            if col_idx_by_name:
                sheet_col_map[ws.title] = col_idx_by_name

        total_rows = sum(
            max(wb[title].max_row - 1, 0) for title in sheet_col_map
        )

        state = state or {}
        # Seed each found column's map from any carried-forward state (see
        # this function's docstring) -- for a fresh single-file call,
        # state is {} so this is identical to the old `{name: {} ...}`.
        placeholder_maps: "Dict[str, Dict[str, str]]" = {
            name: dict(state.get(name, {})) for name in found_names
        }
        # A column's counter is just how many placeholders it already has --
        # correct whether those came from this file or were carried forward,
        # since placeholders are always assigned as a gapless 1..N sequence.
        counters = {name: len(placeholder_maps[name]) for name in found_names}
        cells_redacted = 0
        current = 0

        for title, col_idx_by_name in sheet_col_map.items():
            ws = wb[title]
            for row in ws.iter_rows(min_row=2):
                current += 1
                for cell in row:
                    col_name = col_idx_by_name.get(cell.column)
                    if col_name is None:
                        continue
                    if cell.value is None or str(cell.value).strip() == "":
                        continue
                    value_str = str(cell.value)
                    col_map = placeholder_maps[col_name]
                    placeholder_text = col_map.get(value_str)
                    if placeholder_text is None:
                        counters[col_name] += 1
                        placeholder_text = f"<{label_by_name[col_name]} {counters[col_name]}>"
                        col_map[value_str] = placeholder_text
                    cell.value = placeholder_text
                    cells_redacted += 1
                if progress_callback is not None:
                    try:
                        progress_callback(current, max(total_rows, 1))
                    except Exception:
                        pass

        if phase_callback is not None:
            try:
                phase_callback("saving")
            except Exception:
                pass
        wb.save(output_path)

        key_rows = []
        for col_name, col_map in placeholder_maps.items():
            for value_str, placeholder_text in col_map.items():
                key_rows.append({"column": col_name, "placeholder": placeholder_text, "value": value_str})

        not_found = sorted(set(label_by_name) - found_names)
        summary = {
            "columns_selected": len(selected_columns),
            "columns_found": len(found_names),
            "columns_not_found": not_found,
            "cells_redacted": cells_redacted,
            "unique_values_redacted": sum(len(m) for m in placeholder_maps.values()),
        }
        return summary, key_rows
    finally:
        wb.close()


def _scan_output_for_leaks_and_shape(output_path: str, placeholders_by_value: "Dict[str, List[dict]]"):
    """One read-only pass over output_path collecting both structured leak
    entries and each sheet's (max_row, max_column) shape -- both need a
    full walk of every cell, and there's no reason to make two passes over
    what can be a very large workbook (some real-world exports run to
    tens of thousands of rows) just to get them separately.

    Returns (leaks, out_shape). Each leak entry:
      {'sheet': str, 'coordinate': str, 'row': int, 'value': str,
       'is_header_row': bool, 'column_header': str or None,
       'row_context': [{'label': str, 'value': str}, ...],
       'placeholders': [{'column': str, 'placeholder': str}, ...]}
    'placeholders' lists every (column, placeholder) pair already on
    record for this exact value -- there can be more than one if the same
    value was redacted out of more than one selected column.

    'column_header' is the header text actually sitting in row 1 of the
    leak's OWN column (not necessarily one of the redacted columns listed
    in 'placeholders' -- a leak can just as easily turn up in an unrelated
    column, e.g. a free-text "Notes" field that happens to repeat a
    patient's name). 'row_context' is up to 4 other non-blank cell values
    from the SAME row (excluding the leak cell itself), each paired with
    its own column's header when known -- this is what lets a reviewer
    tell which actual record a leak belongs to (e.g. "Patient ID: 4521,
    Visit Date: 07/01/2016") without having to go open the workbook
    itself. Left empty for a header-row leak, where "other headers in
    this row" wouldn't mean anything as record context.
    """
    real_values = set(placeholders_by_value)
    leaks = []
    out_shape = {}
    wb_out = openpyxl.load_workbook(output_path, read_only=True, data_only=True)
    try:
        for ws in wb_out.worksheets:
            out_shape[ws.title] = (ws.max_row, ws.max_column)
            if not real_values:
                continue
            headers_by_col = {}
            for row_idx, row in enumerate(ws.iter_rows(), start=1):
                if row_idx == 1:
                    # Populated before any leak in this row is processed
                    # below (same row, same iteration) -- and stays
                    # populated for every later row, since row 1 is always
                    # this loop's first iteration.
                    for cell in row:
                        if cell.value is not None:
                            text = str(cell.value).strip()
                            if text:
                                headers_by_col[cell.column] = text
                for cell in row:
                    if cell.value is None:
                        continue
                    val_str = str(cell.value)
                    if val_str not in real_values:
                        continue
                    is_header_row = row_idx == 1
                    row_context = []
                    if not is_header_row:
                        for other in row:
                            if other.column == cell.column or other.value in (None, ""):
                                continue
                            label = headers_by_col.get(other.column) or f"column {get_column_letter(other.column)}"
                            row_context.append({"label": label, "value": str(other.value)})
                            if len(row_context) >= 4:
                                break
                    leaks.append({
                        "sheet": ws.title,
                        "coordinate": cell.coordinate,
                        "row": cell.row,
                        "value": val_str,
                        # Row 1 is always the header row this plugin reads
                        # column names from, and redaction never touches it
                        # (apply_column_redaction() starts at min_row=2) --
                        # flagging this explicitly is what lets a
                        # header-label coincidence be told apart from a
                        # real leak at a glance, without needing to open
                        # the file and check.
                        "is_header_row": is_header_row,
                        "column_header": headers_by_col.get(cell.column),
                        "row_context": row_context,
                        "placeholders": placeholders_by_value[val_str],
                    })
    finally:
        wb_out.close()
    return leaks, out_shape


def verify_column_redaction(input_path: str, output_path: str, selected_columns: List[dict], key_rows: List[dict]) -> dict:
    """Independent post-check, same spirit as a pdf_report plugin's
    verify_redaction(): none of the real values that were redacted should
    still be found anywhere in the output workbook (not just in their own
    column -- a value that leaked into an unrelated cell is just as much a
    problem), and the workbook's overall shape (sheet names, sheet
    dimensions) should be unchanged, since only cell VALUES should have
    been touched.

    Returns {'ok': bool, 'issues': list[str], 'leaks': list[dict],
    'shape_issue': str or None, 'other_issue': str or None}.

    'issues' is the original human-readable summary (what the message-box
    error dialogs show). 'leaks' is the full, structured, UNtruncated list
    behind it (see _scan_output_for_leaks_and_shape()'s docstring for the
    shape of each entry) -- main_gui.py's manual-review screen acts on
    this directly, one leak at a time, rather than parsing 'issues'
    strings.

    'shape_issue' and 'other_issue', when set (else None), flag problems
    that are NOT a per-cell value leak a manual review screen can
    meaningfully resolve -- the workbook's own shape changed unexpectedly,
    or the source file couldn't even be re-opened to compare against.
    main_gui.py only offers manual review when 'leaks' is non-empty AND
    both of these are None; otherwise it falls back to the original hard
    block (delete the output, show an error, nothing saved).
    """
    issues = []
    other_issue = None
    shape_issue = None

    # Map each real value string to every (column, placeholder) pair
    # already on record for it, so a leaked cell can say WHICH selected
    # column's data this is -- e.g. "matches a value redacted from column
    # 'SSN'" -- and, for the manual-review screen, exactly which
    # placeholder text is available to replace it with. This adds no new
    # sensitive information to a message (a column header name and a
    # placeholder tag, not the value itself).
    placeholders_by_value: "Dict[str, List[dict]]" = {}
    for row in key_rows:
        placeholders_by_value.setdefault(row["value"], []).append(
            {"column": row["column"], "placeholder": row["placeholder"]}
        )

    try:
        wb_in = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
        in_shape = {ws.title: (ws.max_row, ws.max_column) for ws in wb_in.worksheets}
        wb_in.close()
    except Exception as exc:  # noqa: BLE001
        other_issue = f"Could not re-open the source file to compare structure: {exc}"
        issues.append(other_issue)
        in_shape = None

    leaks, out_shape = _scan_output_for_leaks_and_shape(output_path, placeholders_by_value)

    if leaks:
        shown = []
        for leak in leaks[:10]:
            cols = ", ".join(sorted({p["column"] for p in leak["placeholders"]}))
            row_note = (
                " [row 1 -- this workbook's header row, not a data row]"
                if leak["is_header_row"] else ""
            )
            found_in = f" -- sitting in column '{leak['column_header']}'" if leak["column_header"] else ""
            shown.append(
                f"{leak['sheet']}!{leak['coordinate']}{row_note}{found_in} "
                f"(matches a value redacted from column '{cols}')"
            )
        more = f" (+{len(leaks) - 10} more)" if len(leaks) > 10 else ""
        issues.append(f"Original value(s) still present in output at: {'; '.join(shown)}{more}")

    if in_shape is not None and in_shape != out_shape:
        shape_issue = (
            "The workbook's sheet names or dimensions changed between input and output "
            "-- something beyond the chosen columns' values may have been affected "
            f"(before: {in_shape}, after: {out_shape})."
        )
        issues.append(shape_issue)

    return {
        "ok": len(issues) == 0,
        "issues": issues,
        "leaks": leaks,
        "shape_issue": shape_issue,
        "other_issue": other_issue,
    }


def apply_leak_decisions(output_path: str, decisions: List[dict]) -> int:
    """Applies manual-review decisions (see verify_column_redaction()'s
    'leaks' field) directly to output_path and re-saves it.

    decisions: [{'sheet': str, 'coordinate': str, 'action': 'replace' or
    'ignore', 'placeholder': str (required when action == 'replace')},
    ...] -- one entry per leak the review screen showed.

    A 'replace' decision overwrites that exact cell with the chosen
    placeholder text, the same as if the original redaction pass had
    caught it. An 'ignore' decision is a no-op here -- this function has
    no way to tell "reviewed and deliberately kept" apart from "never
    looked at", so the caller (main_gui.py) is responsible for recording
    every decision, including ignores, somewhere durable (the manual
    review log) before treating the file as finished.

    Returns the number of cells actually changed. The caller should
    re-run verify_column_redaction() afterward to confirm every leak was
    actually accounted for (replaced, or explicitly, recordedly ignored)
    rather than silently missed because of a coordinate mismatch upstream
    -- this function does not re-verify on its own.
    """
    to_replace = {
        (d["sheet"], d["coordinate"]): d["placeholder"]
        for d in decisions
        if d.get("action") == "replace"
    }
    if not to_replace:
        return 0
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


def _write_key_body(writer, key_rows: List[dict]) -> None:
    writer.writerow(["Column", "Placeholder", "Real Value"])
    for row in key_rows:
        writer.writerow([row["column"], row["placeholder"], row["value"]])


def write_key(key_path: str, input_path: str, output_path: str, key_rows: List[dict]) -> None:
    """Writes a CSV recording every (column, placeholder, real value)
    triple, so the redaction can be reversed later via the Un-redact tab
    (spreadsheet_unredact_engine.py). One row per unique value per column.

    This file is exactly as sensitive as the ORIGINAL, unredacted
    workbook -- anyone holding both it and the redacted file can fully
    de-redact it. It must never be sent alongside the redacted file to
    whatever party the redaction exists to protect against. That warning
    is written into the first few rows of the file itself, since the CSV
    is what actually travels with the case file.
    """
    with open(key_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["CONFIDENTIAL -- Spreadsheet redaction key"])
        writer.writerow(["Same sensitivity as the ORIGINAL, unredacted workbook -- do NOT send or"])
        writer.writerow(["store this file alongside the redacted workbook. Keep it with your own"])
        writer.writerow(["case file."])
        writer.writerow([f"Source workbook: {os.path.basename(input_path)}"])
        writer.writerow([f"Redacted workbook: {os.path.basename(output_path)}"])
        writer.writerow([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
        writer.writerow([])
        _write_key_body(writer, key_rows)


def write_key_batch(key_path: str, file_pairs: "list", key_rows: List[dict]) -> None:
    """Same as write_key(), but for a BATCH of workbooks redacted together
    with shared placeholder numbering (see apply_column_redaction()'s
    `state` parameter and base.py's "BATCH REDACTION" section) -- one
    combined key file covering every file in the batch, instead of one
    key file per file.

    `file_pairs` is [(source_basename, redacted_basename), ...] for every
    file in the batch, in the order they were processed; `key_rows` is the
    caller's full accumulated set across every file's apply_column_redaction()
    call (deduplicated by the caller so each (column, value) pair appears
    once, with the placeholder text it was first assigned).

    Same confidentiality as write_key(): this file is exactly as sensitive
    as the ORIGINAL, unredacted workbooks -- anyone holding it plus the
    redacted files can fully de-redact the whole batch.
    """
    with open(key_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["CONFIDENTIAL -- Spreadsheet redaction key (batch)"])
        writer.writerow(["Same sensitivity as the ORIGINAL, unredacted workbooks -- do NOT send or"])
        writer.writerow(["store this file alongside the redacted workbooks. Keep it with your own"])
        writer.writerow(["case file."])
        writer.writerow([f"Files in this batch: {len(file_pairs)}"])
        for source_name, redacted_name in file_pairs:
            writer.writerow([f"  Source: {source_name}  ->  Redacted: {redacted_name}"])
        writer.writerow([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
        writer.writerow([])
        _write_key_body(writer, key_rows)


def existing_numbering_from_key(key_path: str) -> dict:
    """Reads a spreadsheet redaction key CSV saved by an EARLIER, SEPARATE
    redaction run (write_key()'s own output, or a combined batch key from
    write_key_batch()) and returns a state mapping usable by
    apply_column_redaction()'s `state` parameter --
    {column_name: {real_value: placeholder_text}}. This is what lets the
    Redact tab's "Use existing key file(s) for consistent numbering"
    picker carry a value's placeholder forward from an earlier,
    separately-redacted workbook into this one, without needing both
    workbooks in the same Batch Redact run.

    Unlike the PDF side, the placeholder TEXT itself (not just a number)
    is carried forward verbatim -- apply_column_redaction() only ever
    assigns a NEW placeholder when a value isn't already present in its
    seeded state, so a value that recurs keeps the exact placeholder
    (including whatever label was chosen when the key file was written);
    only a genuinely new value in this document gets a fresh placeholder
    built from the label chosen THIS time. Use the same placeholder label
    each time you redact a given column if you want every document to
    show identical label text for it.

    Also worth knowing: apply_column_redaction()'s key_rows already
    include every value carried forward via `state`, even ones not
    present in the CURRENT file (see its own docstring) -- so the key CSV
    saved after redacting with an earlier key loaded is itself a superset
    covering both documents, and can be handed forward again next time
    instead of needing to keep collecting every individual key file (the
    PDF side does not have this shortcut -- see
    dentrix_audit_trail.existing_numbering_from_key()'s docstring).

    Raises ValueError if this doesn't look like a spreadsheet key CSV at
    all -- the same "Column, Placeholder, Real Value" header check this
    module's own write_key() output always has. This function is
    intentionally self-contained (a small, local CSV parse) rather than
    importing spreadsheet_unredact_engine.py, matching the sibling
    engine's own preference for not depending across that boundary.
    """
    with open(key_path, "r", newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))

    header_seen = False
    state = {}
    for row in rows:
        if not row:
            continue
        if not header_seen:
            if len(row) >= 3 and [c.strip() for c in row[:3]] == ["Column", "Placeholder", "Real Value"]:
                header_seen = True
            continue
        if len(row) >= 3:
            column, placeholder, value = row[0], row[1], row[2]
            state.setdefault(column, {})[value] = placeholder

    if not header_seen:
        raise ValueError(
            "This doesn't look like a spreadsheet redaction key CSV -- no 'Column, "
            "Placeholder, Real Value' table was found in it. Make sure this is a "
            "*_REDACTION_KEY.csv file saved by an earlier redaction, not some other file "
            "(a PDF name key, for instance, has a different format)."
        )
    return state


def merge_existing_numbering(mappings: List[dict]) -> "tuple[dict, list]":
    """Combines several existing_numbering_from_key() results (one per key
    file the user picked in the Redact tab) into one state mapping, the
    same idea as spreadsheet_unredact_engine.merge_key_mappings()
    combining key files for un-redaction. Returns (merged, conflicts);
    conflicts is [{'column': str, 'value': str, 'placeholders': [p1, p2, ...]}]
    for any (column, value) pair assigned a DIFFERENT placeholder across
    the files given -- this should never happen for key files that
    genuinely belong together, so the caller (main_gui.py) treats any
    conflict as blocking rather than silently picking one placeholder
    over another.
    """
    placeholders_by_key = {}
    for mapping in mappings:
        for column, value_map in mapping.items():
            for value, placeholder in value_map.items():
                seen = placeholders_by_key.setdefault((column, value), [])
                if placeholder not in seen:
                    seen.append(placeholder)
    merged = {}
    conflicts = []
    for (column, value), placeholders in placeholders_by_key.items():
        merged.setdefault(column, {})[value] = placeholders[0]
        if len(placeholders) > 1:
            conflicts.append({"column": column, "value": value, "placeholders": placeholders})
    return merged, conflicts
