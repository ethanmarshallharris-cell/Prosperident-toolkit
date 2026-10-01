# Prosperident Document Redactor

A Windows desktop tool for redacting identifying information out of
files -- report PDFs and Excel workbooks today -- and for reversing that
redaction later. It's built as a central app with a **document type**
dropdown: each supported file layout is a separate, self-contained
plugin, so new types can be added over time without changing the app
itself. Two document types exist today:

- **Dentrix Audit Trail Report** -- scans a PDF for patient names inside
  audit trail entries and redacts them.
- **Excel Spreadsheet (choose columns)** -- lets you pick which column(s)
  of an .xlsx workbook to redact, rather than auto-detecting matches (see
  below -- this one works differently from the PDF type, by design).

## What it does, and how

**Redact tab, PDF-style document types (e.g. Dentrix Audit Trail Report).**
Pick the document type from the dropdown, browse to the source PDF, and
click **Scan**. The chosen document type's own logic finds every
occurrence of whatever it's built to redact (for the Dentrix plugin,
patient names) and assigns each distinct one a placeholder --
`<Patient 1>`, `<Patient 2>`, and so on -- the same number every time that
same one appears in the file. Review what was found (and what didn't
clearly match, in "Needs review"), then **Redact & Save**.

Redaction is a genuine removal of the underlying text from the PDF's
content stream, not a black box drawn on top -- the original text cannot
be recovered by copying text out of the output file or by editing it. Before
anything is written to disk, the tool re-opens its own output and
independently re-checks it; if that check doesn't pass, the tool deletes
the output and refuses to report success.

**Redact tab, Excel Spreadsheet (choose columns).** Pick that document
type, browse to the source .xlsx workbook, and click **Load Columns**
(this replaces "Scan" for this document type -- there's nothing to
auto-detect, since you choose the columns yourself). Every column header
found in row 1 of any sheet is listed with a checkbox, a sample value, and
an editable placeholder label (defaulting to the column's own name). Check
the column(s) you want redacted, adjust any label you'd rather use
instead, then **Redact & Save**.

Every non-blank cell in a chosen column is replaced by a placeholder --
`<label 1>`, `<label 2>`, etc. The same real value always gets the same
placeholder everywhere it recurs **within that column** (numbering is
independent per column, so a value repeating in two different chosen
columns does not have to share a placeholder). This is a genuine
replacement of the cell's value, written out to a fresh workbook -- the
original value isn't left recoverable in the file. Before anything is
written to disk, the tool re-opens its own output and checks that none of
the original values it redacted are still present anywhere in the
workbook and that the workbook's sheet structure is unchanged; if that
check doesn't pass, the output is deleted and the tool refuses to report
success -- same standard as the PDF flow.

**Un-redact tab.** Pick a previously redacted file (PDF, Excel workbook,
or a workbook that's since been re-saved as a CSV) and the key CSV(s)
that were saved alongside it, and get back a copy with every placeholder
swapped for the real value it replaced. Which kind of file it is gets
detected automatically from its extension -- the same tab and button
work for any of them, and this works regardless of which document type
originally produced it. See "Un-redacting a file that's changed since it
was redacted" below for how much drift this tolerates, using more than
one key file at once, and what happens when it can't safely guess.

## Consistent numbering across documents redacted separately (Redact tab)

Batch Redact (below) is one way to make sure the same real name or value
gets the same placeholder across several documents -- but it needs every
file together, up front, in one run. If instead you're redacting documents
one at a time, often on different days, the Redact tab has a lighter-weight
option for the same result: **"Consistent numbering"**, next to Scan/Load
Columns.

Click **Use existing key file(s)...** and pick the key CSV(s) an earlier,
separate redaction already saved (a multi-select dialog, same as the
Un-redact tab). Scan (or Load Columns) this document as usual, and any
name or value that also appears in one of those key files gets the exact
same placeholder it already had there; anything genuinely new in this
document gets the next number in the sequence. **Pick the key file(s)
before clicking Scan/Load Columns** -- for a PDF-style document type,
numbering is assigned right at Scan time, so changing this selection
afterward needs a re-scan to take effect (the tool tells you so). For
Excel, numbering isn't resolved until Redact & Save itself, so there's
nothing to re-trigger.

A typical flow: redact document 1 with nothing loaded, keep its key file.
Before scanning document 2, click **Use existing key file(s)...** and pick
document 1's key. Redact document 2 -- its key file now covers both
documents (for Excel, document 2's key is a strict superset of document
1's, since every value carried forward is included even if document 2
didn't happen to contain it; a PDF-style document's key only ever lists
names actually present in that one file, so for a chain of more than two
PDFs, pick *every* earlier key file each time, not just the most recent
one). **Clear** removes every loaded key file and goes back to fresh
numbering.

If two or more of the key files you pick disagree with each other (the
same name or value maps to a different placeholder in different files),
the tool stops before scanning or redacting anything and shows exactly
which one -- same "don't guess through it" handling the Un-redact tab
already uses when its own selected key files conflict. This almost always
means one of the files doesn't actually belong with the others.

This control only appears for a document type whose plugin supports it
(both Dentrix Audit Trail Report and Excel Spreadsheet do). Switching the
document type dropdown clears whatever key file(s) were loaded, since
they were validated against the previous type's key format.

### Linking a spreadsheet column to a PDF's patient names, one file at a time

You don't need Batch Redact to make a name match across a PDF audit trail
and an Excel workbook -- the same linking is available right here on the
single-file Redact tab, for someone redacting the PDF today and the
spreadsheet on a different day. Redact the Dentrix Audit Trail Report PDF
first as usual (getting its `..._NAME_KEY.csv`). Then, redacting the Excel
workbook, click **Use existing key file(s)...** and pick that PDF's key
file -- the tool recognizes it as coming from a *different* document type
and offers it for identity linking instead of same-type numbering.

Click **Load Columns**, and a **"Same identity as:"** dropdown now appears
next to each column, offering "Patient" (or whichever PDF type's key file
you loaded). Pick it for the column that holds patient names, and that
column's placeholder label locks to "Patient" so the text matches too, not
just the number. When you redact, any name that also appears in the PDF's
key file gets the exact same `<Patient N>` it already has there; a name
the PDF never saw gets a new number continuing that same sequence. You can
mix this with same-type key files too (e.g. also loading an earlier
spreadsheet's own key for its non-linked columns) -- pick every relevant
key file before clicking Load Columns.

As with the batch version of this feature, the link is always a
deliberate, manual choice -- never guessed from a column's name -- and two
loaded key files that disagree about the same person's number are
reported and blocked rather than silently resolved.

## Batch Redact tab

Redacts several files together, in one run, so that the same real-world
value (a patient's name, a value in a chosen Excel column) gets the
**same placeholder everywhere it appears** -- instead of each file's
numbering starting over from 1 independently, which is what you'd get
running them one at a time through the regular Redact tab. The file list
can mix document types (e.g. Dentrix Audit Trail Report PDFs and Excel
Spreadsheet workbooks together in one run) and produces one combined key
file **per document type** in the batch.

Add files with **Add Files...** -- pick any mix of supported files in one
dialog; each is matched to its document type automatically by its file
extension. Click **Load Columns** to read the header row of every
selected Excel file and list the union of column names found across all
of them (same idea as the single-file Redact tab's column checklist),
check the columns to redact, choose a folder to save the redacted copies
to, and click **Redact All**. Each output file is named
`<original name>_REDACTED.<ext>` in the folder you chose.

### Making the same person's name match across a PDF and a spreadsheet

By default, a PDF's patient numbering and a spreadsheet column's value
numbering are independent pools, even within the same batch -- "Patient
7" in a redacted PDF has no particular relationship to "\<Name 7\>" in a
redacted spreadsheet. If you want a name to get the exact same
placeholder wherever it shows up, whether that's in a PDF audit trail or
an Excel column, link that column manually: next to the column in the
checklist, a **"Same identity as:"** dropdown appears whenever the batch
also contains a PDF-type file (Dentrix Audit Trail Report today) --
choose that PDF type's entity ("Patient") from the dropdown. The
column's placeholder label locks to match (so the placeholder *text*
matches too, not just the number), and from then on that column's values
are compared against the linked PDF type's own patient names using the
exact same name-matching logic the PDF plugin already uses for its own
internal consistency.

This link is always something **you choose explicitly, one column at a
time** -- it is never guessed from a column's name or a file's name.
Guessing carries a real risk: if two unrelated things were ever matched
up automatically and turned out not to actually be the same person, two
different people's information could end up merged under one placeholder
without anyone noticing. Requiring a deliberate choice avoids that.

It also depends on the spreadsheet values already looking close to how
the PDF plugin writes names (Dentrix uses "Last, First") -- matching is
case- and whitespace-insensitive, but it does **not** try to guess that
"John Smith" and "Smith, John" are the same person by reordering words;
that kind of guess is exactly the wrong-merge risk described above. If a
linked column's values don't line up with the PDF format, those specific
values just fall back to getting their own independent placeholder --
never an incorrect merge.

A few things work differently on this tab than on the single-file Redact
tab, by design:

- **No per-match preview/exclude step.** Batch mode redacts everything
  each file's scan (or chosen columns) finds automatically -- unlike the
  single-file Redact tab, there's no screen to preview and uncheck
  individual matches before redacting. If a particular file needs that
  level of curation, redact it by itself on the regular Redact tab first,
  then include the clean result in a batch with others.
- **Manual review, per file, when a file's self-check finds something.**
  If a workbook's self-check turns up a value that still looks present
  after redaction, or a PDF's self-check turns up a name that's still
  visible somewhere, the batch **pauses** and opens the same review
  screen the single-file Redact tab uses for that one file -- for Excel,
  showing where each value was found (its column, plus a few other values
  from the same row, to help identify the record); for a PDF, showing
  which page each leftover name occurrence is on, plus a snippet of
  nearby text -- so you can decide Replace/Redact or Ignore for each one.
  Click **Apply Decisions** and that file is fixed and the batch carries
  on to the next file. Click **Cancel** and only *that one file* is
  skipped (nothing saved for it) -- the rest of the batch still finishes;
  skipped files are listed in the final summary so you know to redact
  them individually afterward.
- **Still all-or-nothing for anything a review can't fix.** A structural
  problem -- not a single cell value or name occurrence, but something
  like the workbook's own sheet dimensions changing unexpectedly, a PDF's
  page count changing, or an expected placeholder going missing entirely
  -- still takes down the whole batch, the same as before: every file in
  a batch shares its placeholder numbering with the files processed
  before it, so if one of THOSE happens, none of the files produced so
  far in that run can be trusted as a self-consistent set, and the tool
  deletes every output written during that run and reports which file was
  the problem.
- **One combined key file per document type**, saved to the output folder
  as `BATCH_NAME_KEY.csv` (Dentrix Audit Trail Report) or
  `BATCH_REDACTION_KEY.csv` (Excel Spreadsheet) -- same format and same
  confidentiality as the single-file key described below, just covering
  every file of that type in the batch instead of one. When you later
  un-redact one of this batch's files, pick whichever key file matches
  *that file's* format -- a redacted PDF needs the PDF-shaped key, a
  redacted workbook needs the spreadsheet-shaped key. The placeholder
  numbers still agree across both files for anything you linked, so
  either file un-redacts correctly with its own key.
- **PDF needs-review lines are reported, not blocking.** If a Dentrix scan
  finds lines that didn't clearly match the expected pattern, they don't
  stop the batch -- they're written out to a `BATCH_NEEDS_REVIEW.csv` in
  the output folder (source file, page, context, and the raw text) so you
  can check them afterward, the same way the single-file tab's "Needs
  review" list works, just collected across the whole batch into one
  file.

## The key file

With "Also save a name key CSV" checked (on by default, where the document
type supports it), every successful redaction also produces a CSV file
alongside the output -- listing every placeholder next to the real value
it replaced. Opens directly in Excel as a table. This is what the
Un-redact tab reads to reverse a redaction later. The Dentrix plugin names
this `..._NAME_KEY.csv`; the Excel plugin names it
`..._REDACTION_KEY.csv` (columns: Column, Placeholder, Real Value) -- both
are read automatically, matched to the right un-redaction engine by the
redacted file you pick, not by the key file's name.

**This file is exactly as sensitive as the original, unredacted
document.** Anyone holding both it and the redacted file can fully
de-redact it. Keep it with your own case file, not anywhere it could
travel alongside the redacted file -- do not attach both to the same
email, save both to a shared/client-facing folder, and so on. If a given
redaction doesn't need to be reversible, uncheck the box before clicking
**Redact & Save...** and no key file will be written at all.

## Un-redacting a file that's changed since it was redacted

Un-redaction is often attempted well after redaction, on a file that's
been through other hands and other software in between -- what actually
tolerates that varies by document type.

**PDF.** Un-redaction re-scans the whole document for placeholder text on
every page independently -- it doesn't depend on any fixed page number or
position. Page reordering, merging the redacted PDF into a larger
document, or other page-level changes since redaction already work with
no extra steps, as long as the file is still a genuine text-based PDF. The
one thing that breaks it: the PDF getting flattened, printed-to-PDF, or
scanned into an image, which removes the selectable text layer the
placeholders live in entirely -- there is currently no way to recover from
that (it would need OCR, which this tool doesn't do).

**Excel / CSV.** Row order, row count, and which sheet a value lives on
never mattered -- placeholders are matched by their exact text, not by
position. Column drift is also handled automatically in the common case:
if a column was renamed, reordered, or moved since redaction, un-redaction
still finds and replaces that column's placeholders as long as each
placeholder's text is unique across the whole key file (the default
`<Label N>` placeholders always are). The only time this needs a decision
from you is the rarer case where the *same* placeholder text was used for
two different columns (e.g. you typed the same custom label for both) AND
the column headers that would normally tell them apart no longer match --
when that happens, a **Manual Review** window opens listing each such
occurrence with its candidate real values by column, so you can pick the
right one (or leave it as a placeholder). Each occurrence is shown with
more than just a bare cell reference -- the current column header it's
actually sitting under (useful when that column has been renamed since
redaction) and a few other non-blank values from the same row, so you can
tell which record it belongs to at a glance rather than having to open the
spreadsheet yourself. Same "nothing changes until you click Apply,
decisions are logged" behavior as the redact-side Manual Review described
below.

A redacted workbook that's since been re-saved as a **.csv** is also
handled directly -- just pick the .csv file in the Un-redact tab the same
way you would a .xlsx. The output stays a .csv, so you get back a plain
CSV with the real values restored. A workbook saved down to the legacy
**.xls** format isn't read natively (openpyxl only reads .xlsx) -- open it
in Excel and use File > Save As to save a fresh .xlsx copy first, the same
fix as for any other file openpyxl can't parse.

**Using more than one key file at once.** If a file was redacted in more
than one pass over time -- a follow-up redaction that covered additional
columns, say, each producing its own key CSV -- select all of the
relevant key files together in the Un-redact tab's file picker (it's a
multi-select dialog) rather than un-redacting one at a time. They're
combined automatically before anything is un-redacted. If two selected
key files disagree with each other -- the same placeholder mapping to a
different real value in each -- the tool stops before touching anything
and shows you exactly which placeholder and which two values conflict.
That's not something to guess through: it almost always means one of the
files you picked doesn't actually belong with the others (the wrong file
added by mistake, or a key file from a different document entirely), so
fix the file selection and try again rather than expecting the tool to
pick one for you.

## If the tool ever refuses to save a file

The self-check exists so a questionable file never reaches you looking
fine.

For a PDF redaction, what happens next depends on what kind of problem was
found:

- **A name still visible somewhere in the output (most common)** -- the
  same name you redacted is still present on some page, either because it
  genuinely recurs somewhere the scan didn't catch (a line format it
  didn't recognize) or in a spot the scan wasn't looking -- opens a
  **Manual Review** window instead of just refusing to save. See the next
  section.
- **A structural problem** -- the page count changed unexpectedly, or an
  expected `<Patient N>` placeholder is missing entirely -- this is not
  something a per-occurrence review can fix, so the tool falls back to the
  original behavior: no file is written, and the message describes what
  it found. Do not send out the file.

  1. For the structural case, check the **Needs review** tab for the same
     item first -- if it shows up there, the report had a second
     occurrence in a line format the tool didn't recognize.
  2. If it doesn't appear in "Needs review" at all, that points to a
     different underlying issue -- please report it back along with the
     page number and snippet from the message.

For an Excel redaction, what happens next depends on what kind of problem
was found:

- **A value still present in the output (most common)** -- e.g.
  `Sheet1!C14` still contains something that was supposed to be
  redacted -- opens a **Manual Review** window instead of just refusing
  to save. See the next section.
- **The workbook's sheet structure changed unexpectedly** (a sheet was
  added/removed, or a sheet's dimensions changed) -- this is not something
  a cell-by-cell review can fix, so the tool falls back to the original
  behavior: no file is written, and the message tells you before/after
  dimensions so you can judge what happened. Do not send out the file.

### Manual Review (Excel and PDF)

A value the self-check found still present in the output almost always
means one of two things: the value genuinely recurs somewhere you didn't
select for redaction (a real duplicate that should probably also be
scrubbed), or -- Excel only -- a column *header* elsewhere in the
workbook happens to be textually identical to a value you redacted (row 1
is never touched by redaction, so this is a coincidence, not a real
disclosure). The Manual Review window lists every such occurrence and
lets you decide per occurrence; what it shows alongside each one differs
slightly by document type, since a spreadsheet cell and a PDF page
position aren't the same kind of location:

- **Excel** -- tells you which selected column's data each value matches,
  flags a header-row match explicitly (`row 1 -- header row, not a data
  row`), and for a data-row match shows the current column the value is
  actually sitting in (which may differ from the column it was redacted
  out of) plus a few other non-blank values from that same row, so you
  can identify which record it belongs to without opening the spreadsheet
  yourself.
- **PDF** -- shows which page each leftover name occurrence is on, plus a
  short snippet of the surrounding text when one could be captured, so
  you can judge whether it's the same person you already redacted
  elsewhere in the file.

The choice offered is the same idea either way, just named to fit each
format:

- **Replace** (Excel) / **Redact** (PDF) -- for Excel, overwrite that
  exact cell with the placeholder already assigned to that value
  elsewhere (chosen for you if there's only one candidate; pick from a
  dropdown if the value was redacted out of more than one column). For a
  PDF, genuinely remove that exact occurrence from the file -- the same
  whole-page rebuild every other redaction in the file goes through, not
  a visual cover -- reusing the same `<Patient N>` placeholder already
  used for that person elsewhere in the file.
- **Ignore** -- leave the value or name exactly as it is. This is usually
  the right choice for an Excel header-row match.

Nothing is changed until you click **Apply Decisions**; clicking
**Cancel** discards the run entirely, same as the old hard-block
behavior. Every decision -- both replaced/redacted and ignored -- is
written to a `..._manual_review_log.txt` file next to the output,
including the real value or name for anything you chose to ignore, so
there is always a durable record of exactly what was knowingly left
un-redacted and why. **That log file carries the same sensitivity as the
original document whenever it has an "ignored" entry in it** -- treat it
like the key CSV file: keep it with your own case file, never alongside
the redacted output. If you ignore anything, the success message says so
plainly and will not claim the file is fully de-identified.

The same detail shown in that message box is also always saved to a
`..._selfcheck_log.txt` file right next to the file you were saving --
for example, redacting `case.xlsx` to `case_REDACTED.xlsx` writes
`case_REDACTED_selfcheck_log.txt` in that same folder. The dialog itself
always tells you the exact path it used. Since the message box is easy to
close, scroll past, or just find hard to read/copy from in the moment,
this file is the place to go back to later, or to attach if you're
reporting a problem back to us -- it has the full issues list in plain
text, not just whatever fit on screen. (If that folder can't be written
to for some reason, it falls back to `%TEMP%\DocumentRedactor_..._selfcheck_log.txt`
instead -- again, the dialog tells you which.)

If a save or un-redact ever seems to be taking a very long time or not
completing, check `%TEMP%\DocumentRedactor_redact_log.txt` (or
`..._unredact_log.txt`) -- both are written to live, one line per phase as
it actually happens, so that file shows exactly how far it got even if the
run hasn't finished. These two always live in the OS temp folder, since
they're written continuously while the run is still in progress rather
than once at the end. If the run fails outright (an unexpected error, not
a self-check catching something), a companion `..._error_log.txt` has the
full technical detail behind whatever short error message you saw --
saved next to the file you were working with, the same way the
self-check log is, with the same temp-folder fallback if needed. For a
scan or "Load Columns" failure specifically (which happens before you've
chosen an output file at all), the error log is saved next to the source
file you were trying to open instead.

## One-time setup: building the program

This only needs to be done once (or again if the source files ever
change). You'll need Python installed on a Windows PC -- the installer
from [python.org](https://www.python.org/downloads/) works fine and
already includes everything needed (tkinter).

1. Copy this whole folder to the Windows PC.
2. Double-click `build_windows.bat`.
3. When it finishes, your program is `dist\DocumentRedactor.exe`.

That one `.exe` file is everything staff need -- copy it wherever's
convenient (desktop, shared drive, etc.). No installation, no Python
needed on the machine that runs it.

## Using it

**To redact a PDF (e.g. Dentrix Audit Trail Report):**

1. Open `DocumentRedactor.exe`, and stay on the **Redact** tab.
2. Choose the **document type** from the dropdown.
3. **Browse...** and select the source PDF.
4. Click **Scan**. Every item it found is listed in "Detected items",
   already assigned its placeholder.
5. **Review before trusting it:**
   - Click any row to see the exact snippet from the file on the right,
     with a red box around exactly what will be redacted.
   - If a row is a false positive, double-click it to exclude it from
     redaction.
   - Check the **Needs review** tab. This lists anything that looked like
     it might need redacting but didn't match the expected pattern -- it
     will **not** be redacted unless you catch it here. This tab should
     normally be empty or short.
6. Leave "Also save a name key CSV" checked if you may need to reverse the
   redaction later; uncheck it if this particular output never should be.
   Click **Redact & Save...**, choose where to save, and wait for the
   confirmation.

**To redact an Excel workbook:**

1. Open `DocumentRedactor.exe`, and stay on the **Redact** tab.
2. Choose **"Excel Spreadsheet (choose columns)"** from the dropdown.
3. **Browse...** and select the source .xlsx workbook.
4. Click **Load Columns**. Every column header found (across every sheet)
   is listed with a checkbox, a sample value, and a placeholder label you
   can edit.
5. Check the column(s) to redact. Leave the label as-is to use the
   column's own name, or type something else -- whatever you enter is
   what appears inside the placeholder (e.g. label "SSN" produces
   `<SSN 1>`, `<SSN 2>`, ...).
6. Leave "Also save a name key CSV" checked if you may need to reverse the
   redaction later. Click **Redact & Save...**, choose where to save, and
   wait for the confirmation.

**To un-redact (either kind of file):**

1. Switch to the **Un-redact** tab.
2. Browse to the redacted file (PDF or .xlsx), and to the key CSV that
   was saved alongside it.
3. Click **Un-redact & Save...**, choose where to save, and confirm --
   you'll be reminded that the output is exactly as sensitive as the
   original document.

## Adding a new document type

Each document type lives as one module in `src/document_types/`. See
`src/document_types/base.py` for the full contract. There are two shapes
a plugin can take (base.py's "TWO SHAPES OF PLUGIN" section explains
when to use which):

- **pdf_report** (the default) -- auto-detects matches in free text, the
  way `dentrix_audit_trail.py` does. Use this for another paginated report
  format.
- **spreadsheet** -- lets the user choose columns directly instead of
  auto-detecting, the way `excel_spreadsheet.py` does. Use this for
  another tabular/column-oriented format.

Once a new module is written, it's registered with a single line in
`src/document_types/__init__.py` -- nothing in `main_gui.py` needs to
change for a new type to show up in the dropdown (the Redact tab's UI
switches automatically based on the plugin's FORMAT_KIND).

The Un-redact tab needs no changes at all when a new document type is
added, in either shape -- it works from the placeholder/key-file
convention alone (`src/unredact_engine.py` for pdf_report-shaped
plugins, `src/spreadsheet_unredact_engine.py` for spreadsheet-shaped
ones, chosen automatically by the redacted file's extension), independent
of which plugin produced a given file.

To make a new document type available on the **Batch Redact** tab too,
set `SUPPORTS_BATCH = True` and accept the optional carry-forward
parameter on its numbering function (`existing_map` for a pdf_report
plugin, `state` for a spreadsheet plugin), plus a `write_key_batch()`
function alongside its `write_key()`. See base.py's "BATCH REDACTION"
section for the full contract -- this is entirely optional; a plugin
that leaves `SUPPORTS_BATCH` unset simply isn't offered in that tab, and
nothing about its single-file Redact tab behavior changes either way. A
pdf_report plugin can additionally define `normalize_identity()` so a
spreadsheet column can be manually linked to its entity in a mixed batch
(see base.py's "BATCH REDACTION: CROSS-TYPE IDENTITY LINKING" section) --
also optional, and a plugin that skips it still works fine on its own,
it just can't be the target of a cross-type link.

## Known limitations

- **Format assumption:** each document type expects a specific layout. For
  a pdf_report-shaped plugin, if a source system ever changes its report
  format, entries may land in "Needs review" instead of being
  auto-redacted (safe), or in rare cases could be missed entirely (why the
  review step matters). For the Excel plugin, only row 1 of each sheet is
  checked for column headers -- a workbook whose table doesn't start at
  row 1 won't have its columns detected at all (there's no "Needs review"
  equivalent for a column that was never found in the first place).
- **Numbering is per file on the Redact tab.** Running the tool
  separately on two files will number placeholders independently in
  each -- the same person/entity will not necessarily get the same
  number in both. For the Excel plugin, numbering also restarts
  independently in each column of the same file. Use the **Batch Redact**
  tab instead when the same value needs the same placeholder across a
  set of files processed together.
- **Un-redaction restores styling too (PDF only), but it's still an
  approximation.** At redaction time, the Dentrix plugin samples the
  font, size, color, and bold/italic style of the text it's about to
  remove and saves that in the name key CSV. When you later un-redact,
  that captured styling is used to draw the real value back, instead of a
  generic default -- so the result is normally very close to the
  original. It's not a perfect match: the original document's exact
  embedded font program is gone (redaction removes it, by design), so the
  tool substitutes the closest standard font in the same style
  (serif/sans-serif/monospace, bold/italic). Key files saved before this
  feature existed only have "Placeholder" and "Real Name" columns;
  un-redacting from one of those still works, just back in a plain
  default font as before.
- **Excel formatting and formulas are not specially handled.** Redacting a
  cell replaces its value only; number formats/fills generally survive
  (openpyxl keeps a cell's existing style), but formulas or charts that
  referenced the original value are not rewritten and may show stale
  results or errors after redaction.
- **Only modern Excel workbooks (.xlsx) are supported** -- not the legacy
  .xls format, and not .csv.
- See each document type's own module docstring (e.g.
  `dentrix_audit_trail.py`, `excel_spreadsheet.py`) for limitations
  specific to that type.

## Files in this folder

- `src/main_gui.py` -- the desktop application (entry point). Generic
  across all document types -- drives the dropdown, all three tabs
  (Redact, Batch Redact, Un-redact), and the whole workflow entirely
  through the plugin contract in `document_types/base.py`, branching its
  Redact-tab UI on each plugin's FORMAT_KIND and offering a document type
  on the Batch Redact tab only when it sets SUPPORTS_BATCH.
- `src/document_types/` -- one module per supported file layout. `base.py`
  documents the plugin contract (both shapes); `dentrix_audit_trail.py`
  (PDF, pdf_report-shaped) and `excel_spreadsheet.py` (Excel,
  spreadsheet-shaped) are the two document types today.
- `src/unredact_engine.py` -- generic un-redaction for pdf_report-shaped
  plugins' output, used by the Un-redact tab.
- `src/spreadsheet_unredact_engine.py` -- generic un-redaction for
  spreadsheet-shaped plugins' output, used by the Un-redact tab.
- `requirements.txt`, `build_windows.bat` -- build tooling.
- `test/` -- a synthetic sample PDF and generator script used during
  development (not needed to run the program; safe to delete).

## Migrating from the old single-purpose version

This replaces the earlier `dentrix_redactor` folder (which only did
Dentrix redaction, with no document-type dropdown or un-redact feature).
Build and use this folder going forward; the old one can be deleted once
you've confirmed this one builds and runs correctly.
