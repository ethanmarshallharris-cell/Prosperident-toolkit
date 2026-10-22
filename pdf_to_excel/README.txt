PDF TO EXCEL  -  Prosperident Tool Suite
========================================

Converts a PDF laid out in columns (a bank or credit-card statement, a
practice-management report, a deposit listing) into an Excel workbook.
You decide where the columns break by dragging bars on the page.

USING IT
--------
1. Open PDF...  The first page is shown, and column breaks are suggested
   from the white space between columns.

2. Set the TABLE AREA with the two blue dashed bars. Drag the top bar down
   to just above the column headings and the bottom bar up to just below
   the last transaction line. Everything in the shaded parts (bank name,
   address, page footer) is left out. "Reset table area" puts them back at
   the page edges.

3. Set the COLUMN BREAKS with the red bars:
     - drag a bar left or right to move it;
     - double-click anywhere on the page to add one (or "Add break");
     - right-click a bar to remove it (or click it and press Delete);
     - Left/Right arrow keys nudge the selected bar one point
       (hold Shift for five);
     - "Suggest breaks" re-suggests from the current page and table area
       (it works best after step 2).
   The letters above the page show which Excel column each strip becomes.

4. Watch the PREVIEW on the right. It shows exactly what the current page
   will become and updates as bars move. Use the arrow buttons (or
   Page Up/Page Down) to check other pages: the same breaks and table
   area apply to every page.

   If one page is laid out differently (the first page of a statement
   usually has a longer header), tick "This page has its own table
   top/bottom" on that page and move its blue bars; other pages keep the
   shared area.

5. Convert to Excel...  Choose the pages (blank = all, or e.g. 1-3, 5),
   one sheet for everything (with a Page column) or one sheet per page.

OPTIONS
-------
Join wrapped lines to the row above
    A line whose first column is blank is added to the row above it, so a
    description that wraps onto a second line stays one transaction.
    Turn off if blank first cells are legitimate rows of their own.

Show column headings only once   (Heading lines: 1)
    When the PDF reprints the column headings at the top of every page,
    they appear once, at the top of the workbook, instead of on every
    page. The headings are taken from the first page converted: its first
    row, or its first 2 or 3 rows if you set "Heading lines" to 2 or 3
    (headings printed on more than one line). On later pages those rows
    are removed wherever they appear; capitals, spacing and numbers that
    change ("Page 1", "Page 2") don't stop the match. The preview shows
    the headings in green on that first page and, on later pages, how
    many heading rows were hidden. With "One sheet per page", each sheet
    still starts with the headings. Untick to keep every repeat.
    Tip: with "Join wrapped lines" on, a second heading line whose first
    column is blank is joined onto the first, so Heading lines = 1 is
    usually right; with it off, count each printed heading line.

Convert amounts to numbers
    1,234.56  $1,234.56  (45.00)  45.00-  45.00 DR  become numbers
    (negative where shown so). Left as text on purpose: dates, account or
    check numbers with leading zeros, and anything longer than 15 digits.

LAYOUTS
-------
"Save layout..." stores the breaks, table area and options in a small
.pdflayout.json file. Next month, open the new statement and
"Load layout..." to reuse them. Breaks you have already placed are also
kept when you open another PDF in the same session.

LIMITS
------
* The PDF must contain real text. A scanned PDF is a picture of text; the
  program says so when it opens one. Run it through OCR first (for
  example Adobe Acrobat's Scan & OCR), then open the OCR'd copy.
* Password-protected PDFs: save an unprotected copy first.
* Always reconcile totals to the PDF before relying on the conversion.
  The workbook's "About" sheet records the source file and pages.

Errors are logged to:
  Windows: %LOCALAPPDATA%\Prosperident\PDFtoExcel\pdf_to_excel.log
  Mac:     ~/Library/Logs/Prosperident PDF to Excel/pdf_to_excel.log

FILES
-----
pdf_to_excel_app.py   the window
pdf_table_engine.py   the conversion (rows, columns, numbers, Excel)
branding.py, update_check.py, prosperident_logo.png   shared suite files
PDF to Excel.bat      run from source (needs Python)
build_windows.bat / build_mac.sh   build the standalone app by hand
Version: APP_VERSION in pdf_to_excel_app.py (Ship-Update sets it).
