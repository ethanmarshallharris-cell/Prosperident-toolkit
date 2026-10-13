"""
Which documents a key file covers -- shared by every document type's key
writer (added October 2026).

When a document is redacted with "Use existing key file(s)" loaded (the
Redact tab's consistent-numbering option), the key saved for it is written
as a ONE-STOP key: it carries every entry of the earlier key(s) as well as
this document's own, so it alone can un-redact this document AND every
document the earlier keys covered (and, through them, everything before
that -- the chain just keeps growing). The earlier keys are no longer
needed for un-redaction.

To make that visible, a key's header lists each covered document on its
own "Covers: <file name>" row. Keys written before this existed have no
such rows; covered_files() falls back to their "Source workbook:", "Source
PDF:" and batch "Source: X -> Redacted: Y" header rows, so older keys still
report what they cover and still pass their coverage on.

Header rows only -- none of this touches the key's table, which every
reader (redaction carry-forward and all three un-redaction engines) finds
by its column-heading row and reads exactly as before.
"""
import csv
import os

COVERS_PREFIX = "Covers: "


def _read_rows(key_path):
    with open(key_path, "r", newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def _is_table_header(row):
    cells = [c.strip() for c in row[:3]]
    return (len(cells) >= 2 and cells[0] == "Placeholder" and cells[1] == "Real Name") or (
        len(cells) >= 3 and cells == ["Column", "Placeholder", "Real Value"]
    )


def covered_files(key_path):
    """The document names (source files) a key covers, in the order listed.
    Empty if the key names none (or can't be read)."""
    try:
        rows = _read_rows(key_path)
    except OSError:
        return []
    covers, legacy = [], []
    for row in rows:
        if not row:
            continue
        if _is_table_header(row):
            break
        cell = row[0].strip()
        if cell.startswith(COVERS_PREFIX):
            covers.append(cell[len(COVERS_PREFIX):].strip())
        elif cell.startswith("Source workbook: ") or cell.startswith("Source PDF: "):
            legacy.append(cell.split(": ", 1)[1].strip())
        elif cell.startswith("Source: ") and "->" in cell:
            legacy.append(cell[len("Source: "):].split("->")[0].strip())
    return covers or legacy


def coverage(own_sources, carried_from=None):
    """Own documents first, then everything the carried key(s) cover,
    without repeats."""
    out = []
    for name in list(own_sources) + [c for p in (carried_from or []) for c in covered_files(p)]:
        if name and name not in out:
            out.append(name)
    return out


def write_coverage_rows(writer, own_sources, carried_from=None):
    """Writes the "Covers:" header rows (see module docstring)."""
    covers = coverage(own_sources, carried_from)
    if carried_from:
        writer.writerow(["ONE-STOP KEY: this file alone can un-redact every document listed below."])
        writer.writerow([
            "It includes everything from: " + ", ".join(os.path.basename(p) for p in carried_from)
        ])
    for name in covers:
        writer.writerow([COVERS_PREFIX + name])


def table_rows(key_path):
    """The rows of a key's table (below its column-heading row), or [] if
    the file has no recognizable table."""
    try:
        rows = _read_rows(key_path)
    except OSError:
        return []
    out, seen_header = [], False
    for row in rows:
        if not row:
            continue
        if not seen_header:
            seen_header = _is_table_header(row)
            continue
        out.append(row)
    return out


def is_spreadsheet_key(key_path):
    for row in _read_rows(key_path):
        if row and _is_table_header(row):
            return row[0].strip() == "Column"
    return False
