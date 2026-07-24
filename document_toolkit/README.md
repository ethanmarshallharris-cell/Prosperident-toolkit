# Prosperident Document Toolkit

A Windows desktop tool for redacting identifying information out of
report PDFs -- and for reversing that redaction later. It's built as a
central app with a **document type** dropdown: each supported report
layout (Dentrix's Audit Trail Report is the first one) is a separate,
self-contained plugin, so new report types can be added over time without
changing the app itself.

## What it does, and how

**Redact tab.** Pick a document type from the dropdown, browse to the
source PDF, and click **Scan**. The chosen document type's own logic finds
every occurrence of whatever it's built to redact (for the Dentrix plugin,
patient names inside audit trail entries) and assigns each distinct one a
placeholder -- `<Patient 1>`, `<Patient 2>`, and so on -- the same number
every time that same one appears in the file. Review what was found (and
what didn't clearly match, in "Needs review"), then **Redact & Save**.

Redaction is a genuine removal of the underlying text from the PDF's
content stream, not a black box drawn on top -- the original text cannot
be recovered by copying text out of the output file or by editing it. Before
anything is written to disk, the tool re-opens its own output and
independently re-checks it; if that check doesn't pass, the tool deletes
the output and refuses to report success.

**Un-redact tab.** Pick a previously redacted PDF and the `_NAME_KEY.csv`
file that was saved alongside it, and get back a copy with every
placeholder swapped for the real value it replaced. This works the same
way regardless of which document type originally produced the file --
it doesn't need to know or care.

## The name key file

With "Also save a name key CSV" checked (on by default, where the document
type supports it), every successful redaction also produces a CSV file
alongside the PDF -- named the same, with `_NAME_KEY.csv` added -- listing
every placeholder next to the real value it replaced. Opens directly in
Excel as a two-column table. This is what the Un-redact tab reads to
reverse a redaction later.

**This file is exactly as sensitive as the original, unredacted PDF.**
Anyone holding both it and the redacted PDF can fully de-redact the file.
Keep it with your own case file, not anywhere it could travel alongside
the redacted PDF -- do not attach both to the same email, save both to a
shared/client-facing folder, and so on. If a given redaction doesn't need
to be reversible, uncheck the box before clicking **Redact & Save...**
and no key file will be written at all.

## If the tool ever refuses to save a file

The self-check exists so a questionable file never reaches you looking
fine. If it reports a problem, no file is written -- there's nothing to
accidentally send out. For a redaction, the message tells you which page
the leftover text is on and shows a short snippet of the surrounding text.
If this happens:

1. Check the **Needs review** tab for the same item. If it shows up there,
   the report had a second occurrence in a line format the tool didn't
   recognize -- one occurrence redacted cleanly, another was never caught
   in the first place.
2. If it doesn't appear in "Needs review" at all, that points to a
   different underlying issue -- please report it back along with the
   page number and snippet from the message.

Either way, do not send out the file -- the tool already refused to save
it for exactly that reason.

If a save or un-redact ever seems to be taking a very long time or not
completing, check `%TEMP%\DocumentToolkit_redact_log.txt` (or
`..._unredact_log.txt`) -- both are written to live, one line per phase as
it actually happens, so that file shows exactly how far it got even if the
run hasn't finished. If it fails outright, a companion `..._error_log.txt`
has the full technical detail behind whatever short error message you saw.

## One-time setup: building the program

This only needs to be done once (or again if the source files ever
change). You'll need Python installed on a Windows PC -- the installer
from [python.org](https://www.python.org/downloads/) works fine and
already includes everything needed (tkinter).

1. Copy this whole folder to the Windows PC.
2. Double-click `build_windows.bat`.
3. When it finishes, your program is `dist\DocumentToolkit.exe`.

That one `.exe` file is everything staff need -- copy it wherever's
convenient (desktop, shared drive, etc.). No installation, no Python
needed on the machine that runs it.

## Using it

**To redact:**

1. Open `DocumentToolkit.exe`, and stay on the **Redact** tab.
2. Choose the **document type** from the dropdown (only "Dentrix Audit
   Trail Report" exists today; more will appear here as they're added).
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

**To un-redact:**

1. Switch to the **Un-redact** tab.
2. Browse to the redacted PDF, and to the `_NAME_KEY.csv` file that was
   saved alongside it.
3. Click **Un-redact & Save...**, choose where to save, and confirm --
   you'll be reminded that the output is exactly as sensitive as the
   original document.

## Adding a new document type

Each document type lives as one module in `src/document_types/`. See
`src/document_types/base.py` for the full contract (what constants and
functions a module needs to define) and
`src/document_types/dentrix_audit_trail.py` for a complete, working
reference implementation. Once a new module is written, it's registered
with a single line in `src/document_types/__init__.py` -- nothing in
`main_gui.py` needs to change for a new type to show up in the dropdown.

The Un-redact tab needs no changes at all when a new document type is
added -- it works from the placeholder/key-file convention alone
(`src/unredact_engine.py`), independent of which plugin produced a given
file.

## Known limitations

- **Format assumption:** each document type expects a specific layout. If
  a source system ever changes its report format, entries may land in
  "Needs review" instead of being auto-redacted (safe), or in rare cases
  could be missed entirely (why the review step matters).
- **Numbering is per file.** Running the tool separately on two files will
  number placeholders independently in each -- the same person/entity
  will not necessarily get the same number in both.
- **Un-redaction restores the value, not the original styling.** Once a
  page has been redacted, the exact original font at that spot is gone;
  the restored text is drawn back in a plain sans-serif font, auto-sized
  to fit. The text itself is correct -- only its typography is an
  approximation of the original.
- See each document type's own module docstring (e.g.
  `dentrix_audit_trail.py`) for limitations specific to that type, such as
  Dentrix's page-break and inconsistent-name-entry behavior.

## Files in this folder

- `src/main_gui.py` -- the desktop application (entry point). Generic
  across all document types -- drives the dropdown, both tabs, and the
  whole workflow entirely through the plugin contract in
  `document_types/base.py`.
- `src/document_types/` -- one module per supported report layout.
  `base.py` documents the plugin contract; `dentrix_audit_trail.py` is the
  first (and currently only) document type.
- `src/unredact_engine.py` -- the generic un-redaction logic used by the
  Un-redact tab, independent of document type.
- `requirements.txt`, `build_windows.bat` -- build tooling.
- `test/` -- a synthetic sample PDF and generator script used during
  development (not needed to run the program; safe to delete).

## Migrating from the old single-purpose version

This replaces the earlier `dentrix_redactor` folder (which only did
Dentrix redaction, with no document-type dropdown or un-redact feature).
Build and use this folder going forward; the old one can be deleted once
you've confirmed this one builds and runs correctly.
