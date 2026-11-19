"""
Reading .xls files for the Document Redactor (added November 2026).

openpyxl reads only modern .xlsx workbooks. A file named .xls is one of
several different things, and this module opens each of them as an
openpyxl Workbook, so the spreadsheet plugin (document_types/
excel_spreadsheet.py) redacts it exactly like an .xlsx. The redacted copy
is always SAVED AS .xlsx: nothing reliable writes the old binary format
any more, and Excel opens the .xlsx the same way.

What a ".xls" can really be -- told apart by the file's first bytes, not
its name:
  * A genuine Excel 97-2003 workbook (binary, starts D0 CF 11 E0): read
    with xlrd, every sheet. Numbers stay numbers (whole numbers as whole
    numbers, so 123456789 is not "123456789.0"), dates become dates,
    TRUE/FALSE stay true/false, text stays text exactly.
  * An .xlsx that was simply named .xls (starts "PK"): read as .xlsx.
  * An "Excel" export that is really a web page (HTML table) -- common from
    practice-management and banking systems: each table becomes a sheet;
    every value stays text exactly as shown.
  * An XML Spreadsheet 2003 file (<?xml ... office:spreadsheet): every
    worksheet, with its numbers, dates and text.
  * A tab- or comma-delimited text file named .xls: read like a .csv (see
    csv_table.py), values as text.
Formatting (fonts, colors, column widths) is not carried over -- only the
values, which is what redaction works on.
"""
import io
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from html.parser import HTMLParser

import openpyxl
from openpyxl.cell.cell import TYPE_STRING

_SS = "{urn:schemas-microsoft-com:office:spreadsheet}"


def is_xls(path) -> bool:
    return str(path).lower().endswith(".xls")


def kind_of(path) -> str:
    """'binary', 'xlsx', 'html', 'xml2003' or 'text' -- what the file is."""
    with open(path, "rb") as f:
        head = f.read(4096)
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "binary"
    if head.startswith(b"PK"):
        return "xlsx"
    text = head.lstrip(b"\xef\xbb\xbf\xff\xfe\xfe\xff \t\r\n").lower()
    if text.startswith(b"<"):
        if b"urn:schemas-microsoft-com:office:spreadsheet" in head.lower() and b"<html" not in text[:200]:
            return "xml2003"
        return "html"
    if b"<table" in head.lower():
        return "html"
    return "text"


def _safe_title(name, used):
    t = re.sub(r"[\[\]:*?/\\]", "_", str(name or "Sheet")).strip() or "Sheet"
    t = t[:31]
    base, n = t, 2
    while t in used:
        suffix = f" ({n})"
        t = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(t)
    return t


def _put(ws, r, c, value, as_text=False):
    if value is None or value == "":
        return
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except Exception:  # characters openpyxl refuses (control codes)
        cell._value = str(value)
        cell.data_type = TYPE_STRING
    if as_text or (isinstance(value, str) and value.startswith("=")):
        cell.data_type = TYPE_STRING   # never turned into a formula


def _load_binary(path):
    try:
        import xlrd
    except ImportError as exc:  # pragma: no cover - bundled with the app
        raise RuntimeError("Reading .xls files needs the xlrd package (pip install xlrd).") from exc
    try:
        book = xlrd.open_workbook(path)
    except Exception as exc:
        msg = str(exc)
        if "encrypt" in msg.lower() or "password" in msg.lower():
            raise RuntimeError("This .xls file is password-protected. Open it in Excel, remove the "
                               "password and save a copy, then use that.") from exc
        raise RuntimeError(f"Could not read this .xls file ({msg}). Open it in Excel and use File > "
                           "Save As to save an .xlsx copy, then use that.") from exc
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    for sh in book.sheets():
        ws = wb.create_sheet(_safe_title(sh.name, used))
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                t, v = cell.ctype, cell.value
                if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                if t == xlrd.XL_CELL_NUMBER:
                    v = int(v) if float(v).is_integer() and abs(v) < 1e15 else v
                elif t == xlrd.XL_CELL_DATE:
                    try:
                        v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                    except Exception:  # an out-of-range date: keep the number
                        pass
                elif t == xlrd.XL_CELL_BOOLEAN:
                    v = bool(v)
                elif t == xlrd.XL_CELL_ERROR:
                    v = xlrd.error_text_from_code.get(v, "#ERROR")
                _put(ws, r + 1, c + 1, v, as_text=(t == xlrd.XL_CELL_TEXT))
    if not wb.worksheets:
        wb.create_sheet("Sheet1")
    return wb


class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []      # [[ [cell, ...], ... ], ...]
        self._stack = []      # open tables (nested tables are read as their own)
        self._cell = None
        self._span = 1

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "table":
            self._stack.append([])
        elif not self._stack:
            return
        elif tag == "tr":
            self._stack[-1].append([])
        elif tag in ("td", "th"):
            if not self._stack[-1]:
                self._stack[-1].append([])
            self._cell = []
            try:
                self._span = max(1, int(dict(attrs).get("colspan") or 1))
            except ValueError:
                self._span = 1
        elif tag == "br" and self._cell is not None:
            self._cell.append("\n")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("td", "th") and self._cell is not None and self._stack:
            text = " ".join("".join(self._cell).replace("\xa0", " ").split())
            row = self._stack[-1][-1]
            row.append(text)
            row.extend([""] * (self._span - 1))
            self._cell = None
        elif tag == "table" and self._stack:
            t = self._stack.pop()
            if any(any(c for c in r) for r in t):
                self.tables.append(t)

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _decode(raw: bytes) -> str:
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", errors="replace")
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16", errors="replace")
    m = re.search(rb'charset=["\']?([A-Za-z0-9_\-]+)', raw[:4096])
    for enc in ([m.group(1).decode()] if m else []) + ["utf-8", "cp1252"]:
        try:
            return raw.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("cp1252", errors="replace")


def _load_html(path):
    with open(path, "rb") as f:
        text = _decode(f.read())
    p = _TableParser()
    p.feed(text)
    p.close()
    if not p.tables:
        raise RuntimeError("This .xls file is a web page with no table in it -- nothing to redact.")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    for n, table in enumerate(p.tables, 1):
        ws = wb.create_sheet(_safe_title("Sheet1" if len(p.tables) == 1 else f"Table {n}", used))
        for r, row in enumerate(table, start=1):
            for c, v in enumerate(row, start=1):
                _put(ws, r, c, v, as_text=True)
    return wb


def _xml_value(data):
    t = data.get(_SS + "Type", "String")
    text = data.text or "".join(data.itertext())
    if t == "Number":
        try:
            f = float(text)
            return int(f) if f.is_integer() and abs(f) < 1e15 else f
        except ValueError:
            return text
    if t == "DateTime":
        try:
            return datetime.fromisoformat(text.rstrip("Z")[:26])
        except ValueError:
            return text
    if t == "Boolean":
        return text.strip() in ("1", "true", "TRUE")
    return text


def _load_xml2003(path):
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError as exc:
        raise RuntimeError(f"Could not read this .xls file (it is an XML spreadsheet, but damaged: {exc}).") from exc
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    for sheet in root.iter(_SS + "Worksheet"):
        ws = wb.create_sheet(_safe_title(sheet.get(_SS + "Name"), used))
        table = sheet.find(_SS + "Table")
        if table is None:
            continue
        r = 0
        for row in table.iter(_SS + "Row"):
            r = int(row.get(_SS + "Index", r + 1))
            c = 0
            for cell in row.iter(_SS + "Cell"):
                c = int(cell.get(_SS + "Index", c + 1))
                data = cell.find(_SS + "Data")
                if data is not None:
                    v = _xml_value(data)
                    _put(ws, r, c, v, as_text=isinstance(v, str))
                c += int(cell.get(_SS + "MergeAcross", 0))
    if not wb.worksheets:
        wb.create_sheet("Sheet1")
    return wb


def _load_text(path):
    import csv_table
    rows, _fmt = csv_table.read_csv(path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            _put(ws, r, c, v, as_text=True)
    return wb


def load_as_workbook(path):
    """An openpyxl Workbook with the .xls file's sheets and values."""
    kind = kind_of(path)
    if kind == "binary":
        return _load_binary(path)
    if kind == "xlsx":
        with open(path, "rb") as f:
            return openpyxl.load_workbook(io.BytesIO(f.read()))
    if kind == "html":
        return _load_html(path)
    if kind == "xml2003":
        return _load_xml2003(path)
    return _load_text(path)
