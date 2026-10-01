"""
The contract every document-type plugin implements, and the generic data
shapes the central app (main_gui.py) uses to work with any of them without
knowing anything document-type-specific.

--------------------------------------------------------------------------
HOW TO ADD A NEW DOCUMENT TYPE
--------------------------------------------------------------------------
1. Create a new module in this package, document_types/, e.g.
   eaglesoft_audit_trail.py. dentrix_audit_trail.py is a complete
   reference implementation -- copying it and adjusting the parsing
   logic is the fastest way to start a close variant (a different
   practice-management system's report, for instance).

2. In that module, define the module-level constants and functions listed
   below under "Required module contents". main_gui.py never imports a
   specific document type directly -- it only ever goes through this
   contract, so nothing about the GUI needs to change when a new type is
   added.

3. Register it by adding one line to DOCUMENT_TYPES in
   document_types/__init__.py.

4. Run the regression check for the new type (see dentrix_audit_trail.py's
   module docstring for how that one's tested) before shipping it --
   redaction correctness bugs are exactly the kind of thing that's cheap
   to catch with a test and expensive to find out about from a client
   file.

--------------------------------------------------------------------------
TWO SHAPES OF PLUGIN: "pdf_report" (default) AND "spreadsheet"
--------------------------------------------------------------------------
Everything below this note describes the original contract, used by any
plugin that redacts free-text matches out of a paginated document (PDFs,
so far). Set FORMAT_KIND = "spreadsheet" instead to use the other shape:
column-based redaction of a workbook, where the user picks which columns
to redact by name rather than the plugin auto-detecting matches. See
excel_spreadsheet.py for a complete reference implementation and its own
module docstring for the full spreadsheet contract (list_columns,
apply_column_redaction, verify_column_redaction, write_key -- deliberately
different function names/signatures from the pdf_report contract below,
since the workflows aren't actually the same shape). main_gui.py branches
its Redact-tab UI on FORMAT_KIND; everything else (the document-type
dropdown, the Un-redact tab's format detection by file extension) already
works generically. If FORMAT_KIND isn't set at all, "pdf_report" is
assumed, so existing plugins don't need to change.

--------------------------------------------------------------------------
REQUIRED MODULE CONTENTS (FORMAT_KIND = "pdf_report", the default)
--------------------------------------------------------------------------
Constants:
  ID                 str  -- stable internal key, e.g. "dentrix_audit_trail".
                             Used in log filenames and the registry; never
                             shown to the user, so it can be a plain
                             lowercase/underscore slug.
  DISPLAY_NAME        str -- shown in the document-type dropdown, e.g.
                             "Dentrix Audit Trail Report".
  ENTITY_LABEL        str -- singular noun for what's being redacted, e.g.
                             "Patient". Used to build placeholder text
                             (f"<{ENTITY_LABEL} {n}>") and summary wording.
  ENTITY_LABEL_PLURAL  str, optional -- defaults to ENTITY_LABEL + "s" if
                             not set. Override this if that default reads
                             wrong (e.g. an ENTITY_LABEL that already ends
                             in an "s", or an irregular plural).
  FILE_TYPES          list[tuple[str, str]] -- tkinter filedialog filetypes
                             for the browse dialog, e.g.
                             [("PDF files", "*.pdf")].
  CONTEXT_COLUMN_LABEL str -- header for the "context" column in both
                             result tables, e.g. "Entry (Date / Time / User)".
  REVIEW_HINT         str  -- explanatory text shown above the "Needs
                             review" tab, describing what lands there and
                             why for this document type.
  NO_MATCHES_WARNING  str  -- shown if a scan finds nothing at all, to
                             suggest this file's layout may not match what
                             this document type expects.
  SUPPORTS_KEY_FILE   bool -- whether write_key() is implemented. If
                             False, the "save a name key" checkbox is
                             hidden for this document type.
  SUPPORTS_BATCH       bool, optional -- defaults to False (via getattr,
                             same pattern as FORMAT_KIND) if not set.
                             Whether this plugin supports the Batch Redact
                             tab's cross-file consistent-numbering mode --
                             see "BATCH REDACTION" below. Existing plugins
                             that don't set it simply don't offer batch
                             mode; nothing else about them needs to change.

Functions:
  scan(path) -> (matches: list[Match], unmatched: list[UnmatchedLine], page_count: int)
      Reads the source file and finds every redaction candidate.

  assign_placeholder_numbers(matches, existing_map=None) -> dict
      Assigns .placeholder_no to each match (same number for the same
      underlying entity throughout the file). Mutates the matches in
      place; the returned dict (normalized identity -> number) is for the
      caller's convenience (e.g. counting unique entities), not required
      by the GUI itself.

      `existing_map`, if given, seeds the numbering instead of starting
      fresh at 1 -- see "BATCH REDACTION" below. Only meaningful if
      SUPPORTS_BATCH is True; a plugin that doesn't support batch mode can
      leave this parameter out of its own signature entirely, since the
      GUI only ever passes it when batch mode is active.

  apply_redactions(input_path, output_path, matches, progress_callback=None, phase_callback=None) -> None
      Writes the redacted file. progress_callback(current, total), if
      given, is called repeatedly with 1-based page progress; any
      exception it raises must be swallowed, never allowed to interrupt
      the redaction. phase_callback(phase_name), if given, is called at
      major phase transitions that can't be given per-page progress (e.g.
      the final disk write) -- also with exceptions swallowed. Called once
      per file even in batch mode -- nothing about this function changes
      for batch; only the matches it's given already carry batch-wide
      placeholder numbers.

  verify_redaction(input_path, output_path, matches, progress_callback=None) -> dict
      Independently re-checks the output before it's trusted. Returns
      {'ok': bool, 'issues': list[str], 'leaks': list[dict], 'structural_issue':
      str | None}. A non-ok result means the GUI will refuse to hand back
      the file UNLESS it can be resolved through manual review instead --
      see apply_pdf_leak_decisions() below. progress_callback behaves as
      above. Also called once per file in batch mode, unchanged.

      'leaks' and 'structural_issue' are what make a non-ok result
      reviewable rather than an automatic hard block: 'leaks' is
      [{'name': str, 'occurrences': [{'page_index': int, 'rect':
      [x0, y0, x1, y1], 'context': str | None}, ...]}, ...] -- every
      leaked name, each with the exact on-page location(s) it was still
      found at. 'structural_issue' is a single string (or None) covering
      anything that ISN'T a single leaked-name occurrence -- a page-count
      mismatch, a missing expected placeholder, anything a per-occurrence
      decision has no way to fix. The GUI only offers manual review when
      'leaks' is non-empty AND 'structural_issue' is None; a plugin that
      doesn't distinguish these can simply always return
      'structural_issue': None with an empty 'leaks' list (or omit both
      keys entirely, since the GUI treats them as absent via .get()) and
      every non-ok result falls back to the original hard block, same as
      before this fields existed.

  write_key(key_path, input_path, output_path, matches) -> None
      Only required if SUPPORTS_KEY_FILE is True. Writes a file that lets
      every placeholder be reversed back to the real value it replaced.

  write_key_batch(key_path, file_pairs, matches) -> None
      Only required if SUPPORTS_KEY_FILE and SUPPORTS_BATCH are both True.
      Same as write_key(), but writes ONE combined key file for a whole
      batch of files redacted together, instead of one key file per file.
      `file_pairs` is [(source_basename, redacted_basename), ...] for
      every file in the batch, in processing order; `matches` is the
      concatenation of every file's own matches (already numbered via the
      shared existing_map/state thread), in that same order.

  apply_pdf_leak_decisions(output_path, decisions, name_to_patient_no) -> int
      OPTIONAL, pdf_report shape only -- the PDF-side analog of the
      spreadsheet shape's apply_leak_decisions() (see
      excel_spreadsheet.py). Only meaningful (and only called) when
      verify_redaction() returns a non-empty 'leaks' list with
      'structural_issue': None -- see above. Applies the user's manual
      review decisions for leaked occurrences directly to output_path, in
      place: `decisions` is [{'name': str, 'page_index': int, 'rect':
      [x0, y0, x1, y1], 'action': 'redact' | 'ignore'}, ...] (one entry
      per occurrence from verify_redaction()'s 'leaks', built by the
      GUI's PDF leak review screen); `name_to_patient_no` is
      {raw_name: patient_no} for every name already assigned a number in
      this file, so a 'redact' decision reuses that name's EXISTING
      placeholder rather than allocating a new one -- see
      dentrix_audit_trail.py's own docstring on this function for the
      full reasoning (genuine content-stream removal via a second
      apply_redactions() pass, not a visual cover). Returns the number of
      occurrences actually redacted. A plugin that doesn't define this
      attribute simply never offers PDF manual review -- the GUI checks
      hasattr() before ever calling it, and every non-ok verify_redaction()
      result falls back to the original hard block for that plugin.

--------------------------------------------------------------------------
BATCH REDACTION
--------------------------------------------------------------------------
The Batch Redact mode in main_gui.py lets the user select several files
of the SAME document type at once and redact them together so that the
same real-world value (e.g. one patient's name) gets the SAME placeholder
across every file in the batch, instead of each file's numbering
restarting at 1 independently -- and produces one combined key file for
the whole batch rather than one per file.

A plugin opts in by setting SUPPORTS_BATCH = True and accepting the
optional carry-forward parameter on its numbering function:
  - pdf_report shape: assign_placeholder_numbers(matches, existing_map)
  - spreadsheet shape: apply_column_redaction(..., state=...) (see
    excel_spreadsheet.py's own docstring for the spreadsheet contract)

The GUI's batch worker calls each file through the plugin's ordinary
single-file functions (scan/apply_redactions/verify_redaction, or the
spreadsheet equivalents) exactly as it does outside batch mode -- nothing
about per-file processing changes. The only difference is that it keeps
the running existing_map/state dict across the loop's iterations (seeding
each new file's call with what came out of the previous one) and, at the
end, hands the accumulated matches/key rows to write_key_batch() /
write_key_batch() instead of calling write_key() once per file. This
keeps the per-file redaction logic itself completely unaware of batching;
only the numbering seed and the final key-file write know about it.

If a plugin doesn't implement the optional existing_map/state parameter
(SUPPORTS_BATCH left unset/False), the GUI simply doesn't offer batch
mode for that document type -- single-file redaction is unaffected either
way.

A batch can mix files of DIFFERENT document types together (e.g. some PDF
audit trails and some Excel workbooks in one run) as long as every type
involved supports batch mode -- each file is still processed through its
own plugin's ordinary functions; mixing types only changes which files
land in the same run and, optionally, whether some of them share an
identity pool (next).

--------------------------------------------------------------------------
REDACT TAB: CONSISTENT NUMBERING FROM AN EXISTING KEY FILE (optional)
--------------------------------------------------------------------------
Batch Redact (above) is one way to keep the same real-world value getting
the same placeholder across several documents -- but it requires having
every file to redact together, up front, in one run. The single-file
Redact tab offers a second, complementary way to get the same result for
someone redacting documents one at a time, often on different days: load
the key file(s) an EARLIER, separate redaction already wrote, and this
document's own numbering picks up where that one left off, using the
exact same existing_map/state parameter Batch Redact's cross-file
numbering already relies on above.

A plugin opts into this by implementing two more functions (independent
of SUPPORTS_BATCH/write_key_batch, though in practice every plugin that
supports one will want to support both -- both need the same
existing_map/state carry-forward parameter on the numbering function):

  existing_numbering_from_key(key_path: str) -> dict
      Reads a key CSV this plugin's own write_key() (or write_key_batch())
      wrote in some EARLIER, unrelated run and returns an existing_map (or
      state, for the spreadsheet shape) usable by the numbering function's
      carry-forward parameter -- the identical shape existing_map/state
      already has in "BATCH REDACTION" above. Raises ValueError if
      key_path doesn't look like this plugin's own key-file format (e.g.
      the wrong file was picked). See dentrix_audit_trail.py's and
      excel_spreadsheet.py's own docstrings on this function for the
      per-shape details -- notably, the PDF shape only carries a NUMBER
      forward (the placeholder LABEL is always rebuilt fresh from this
      plugin's own ENTITY_LABEL), while the spreadsheet shape carries the
      placeholder TEXT forward verbatim.

  merge_existing_numbering(mappings: list[dict]) -> tuple[dict, list]
      Combines several existing_numbering_from_key() results (the user
      can load more than one prior key file at once) into one
      existing_map/state, the same way each plugin's sibling *_unredact_
      engine.py already has a merge_key_mappings() for combining key
      files on the Un-redact side. Returns (merged, conflicts); a
      non-empty conflicts list means two of the given files disagree
      about what number/placeholder the same identity should get --
      main_gui.py treats that as blocking (nothing is scanned or
      redacted) rather than silently picking one file's answer over the
      other's, the same "don't guess through it" principle the Un-redact
      tab's own key-conflict handling already uses.

A plugin that doesn't define these two functions (or doesn't set
SUPPORTS_BATCH) simply never offers this control on the Redact tab --
the GUI checks hasattr() for both before ever using them, and ordinary
single-file redaction is unaffected either way.

--------------------------------------------------------------------------
BATCH REDACTION: CROSS-TYPE IDENTITY LINKING (optional)
--------------------------------------------------------------------------
Within one document type, "the same value gets the same placeholder"
above is enough on its own -- assign_placeholder_numbers()'s existing_map
(or apply_column_redaction()'s state) already IS that mechanism. Linking
identities ACROSS two different document types in the same batch (e.g. a
patient named in both a Dentrix audit trail PDF and an Excel export, who
should get the same "<Patient N>" placeholder in both) needs one more
piece: a way to compare a raw value from one plugin's world (a PDF name
match) against a raw value from another's (an Excel cell) and decide
they're the same identity.

That comparison is always a DELIBERATE, MANUAL choice made by the user in
the Batch Redact tab -- one spreadsheet column, explicitly linked by the
user to one pdf_report-shaped document type's identity (by ENTITY_LABEL),
regardless of what that column happens to be named. There is no
automatic inference from column names or file names; the risk of quietly
merging two different real people under one placeholder because two
unrelated columns happened to share a word is exactly the kind of
mistake this tool exists to prevent, not create.

  normalize_identity(text) -> str      -- OPTIONAL, pdf_report shape only
      Reduces a raw display value (a PDF name match's display_text, or a
      linked spreadsheet cell's value) to a comparable identity key --
      e.g. dentrix_audit_trail.py's is trim/collapse-whitespace/lowercase
      (the same normalization assign_placeholder_numbers() already uses
      for its own within/across-file matching, exposed under this
      generic name specifically so cross-type linking uses the IDENTICAL
      function, not a separate lookalike that could disagree with it on
      an edge case). If a plugin doesn't define this, the Batch Redact
      tab falls back to a generic trim/collapse-whitespace/lowercase
      normalizer when that plugin's identity is linked to a column.

      This does NOT attempt to reconcile different NAME ORDERINGS (e.g.
      "Smith, John" vs "John Smith") -- doing that automatically risks
      the exact wrong-merge mistake described above (a heuristic reorder
      guess can misfire on a real name and silently combine two
      different people). If a linked column's values aren't already
      close to the PDF plugin's own format, cross-type matching for that
      column will simply not find shared identities, safely falling back
      to independent numbering for those specific values -- never an
      incorrect merge.

  When the user links a spreadsheet column to a pdf_report type's
  identity, the Batch Redact tab also forces that column's placeholder
  LABEL to match the linked type's ENTITY_LABEL (and disables editing it)
  -- otherwise the placeholder NUMBER could match while the placeholder
  TEXT still visibly differs (e.g. "<Patient 7>" vs "<Name 7>"), which
  would defeat the point.

  A mixed batch with any linked columns still produces one key file PER
  FORMAT (one combined PDF key, one combined spreadsheet key -- see
  write_key_batch() above), not one unified file -- the two formats'
  key-file shapes differ (the PDF key carries font/size/color/flags,
  which has no spreadsheet equivalent) and the Un-redact tab already
  supports selecting more than one key file at once for exactly this
  situation. The placeholder NUMBERS still agree across both files for
  any linked identity, so un-redacting either file (or both, together)
  with its own key restores the right value either way.

--------------------------------------------------------------------------
REDACT TAB: CROSS-TYPE IDENTITY LINKING FROM A KEY FILE (optional)
--------------------------------------------------------------------------
This is the single-file Redact tab's analog of "BATCH REDACTION:
CROSS-TYPE IDENTITY LINKING" above, for someone who doesn't have every
related file in hand at once -- e.g. redact a Dentrix Audit Trail Report
PDF today, and next week redact an unrelated Excel export that happens to
name some of the same patients, wanting the SAME "<Patient N>" they
already got in the PDF. No new plugin-contract functions are needed for
this: it's a GUI-side reuse, in a new combination, of functions both
sections above already define --
existing_numbering_from_key()/merge_existing_numbering() (from "REDACT
TAB: CONSISTENT NUMBERING FROM AN EXISTING KEY FILE") and
normalize_identity() (from "BATCH REDACTION: CROSS-TYPE IDENTITY
LINKING").

The Redact tab's "Use existing key file(s)..." picker (same control used
for same-type carry-forward) accepts a key file from a DIFFERENT,
pdf_report-shaped document type when the CURRENT type being redacted is
the spreadsheet shape. main_gui.py tells the two kinds of key file apart
automatically, per file: it first tries the CURRENT type's own
existing_numbering_from_key(); if that fails and the current type is a
spreadsheet, it then tries every other, non-spreadsheet type's
existing_numbering_from_key() in turn. A file that no type recognizes is
rejected with an error; a file that more than one type's parser accepts
(not possible today, with only one pdf_report type, but handled
defensively for whenever a second one exists) is also rejected rather than
guessed at.

Once a cross-type key file is loaded, clicking Load Columns adds a "Same
identity as:" dropdown next to each spreadsheet column -- the exact same
manual, per-column link described in "BATCH REDACTION: CROSS-TYPE
IDENTITY LINKING" above, just offered outside of a batch run and seeded
from a loaded key file's already-assigned numbers instead of starting an
identity pool empty. Linking a column still forces its placeholder LABEL
to match the linked type's ENTITY_LABEL (and locks it), and a value never
seen in any loaded key file still gets a genuinely new number continuing
that same pool -- both exactly as in the batch case. Two loaded cross-type
key files that disagree about the same identity's number are blocked with
an error, the same "don't guess through it" handling used everywhere else
key-file conflicts can occur.

--------------------------------------------------------------------------
WHY DUCK TYPING, NOT AN ABSTRACT BASE CLASS
--------------------------------------------------------------------------
A plugin here is a plain module, not a class instance -- there's no
DocumentType base class to subclass. That's deliberate: every existing
document type so far is a fairly direct port of already-working,
already-tested standalone code (see dentrix_audit_trail.py's history), and
forcing that into a class hierarchy would mean touching tested internals
just to satisfy an interface, for no real benefit at this scale. If a
second document type ever needs to genuinely SHARE logic with the first
(not just follow the same shape), that's the point to introduce a real
base class or shared helper module -- not before.

--------------------------------------------------------------------------
GENERIC DATA SHAPES
--------------------------------------------------------------------------
"""
from dataclasses import dataclass
from typing import Optional


@dataclass
class Match:
    """One detected, redactable item found in the source document (e.g. a
    patient name on one line of a Dentrix audit trail entry).

    A document-type module's own scan() may use its own dataclass
    internally (dentrix_audit_trail.py does, for historical reasons --
    see that module's docstring) as long as it exposes these same
    attributes, whether as real fields or read-only properties. The GUI
    only ever accesses matches through these names.
    """
    page_index: int          # 0-based
    rect: object              # fitz.Rect (or equivalent) -- opaque to the GUI, used only
                               # for preview rendering and re-passed into apply_redactions
    right_boundary: float     # rightmost x-coordinate the redaction may safely extend to
    display_text: str         # the actual detected text, e.g. "Smith, John"
    context: str               # short label for what section/entry this came from
    placeholder_no: Optional[int] = None   # set by assign_placeholder_numbers()


@dataclass
class UnmatchedLine:
    """A line that sat somewhere a match was expected but didn't cleanly
    fit the expected pattern -- surfaced to the user rather than silently
    skipped (which could hide a real, unredacted identifier) or silently
    redacted (which could redact something that wasn't actually sensitive).
    """
    page_index: int           # 0-based
    context: Optional[str]
    raw_text: str
