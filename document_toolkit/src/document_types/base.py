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
REQUIRED MODULE CONTENTS
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

Functions:
  scan(path) -> (matches: list[Match], unmatched: list[UnmatchedLine], page_count: int)
      Reads the source file and finds every redaction candidate.

  assign_placeholder_numbers(matches) -> dict
      Assigns .placeholder_no to each match (same number for the same
      underlying entity throughout the file). Mutates the matches in
      place; the returned dict (normalized identity -> number) is for the
      caller's convenience (e.g. counting unique entities), not required
      by the GUI itself.

  apply_redactions(input_path, output_path, matches, progress_callback=None, phase_callback=None) -> None
      Writes the redacted file. progress_callback(current, total), if
      given, is called repeatedly with 1-based page progress; any
      exception it raises must be swallowed, never allowed to interrupt
      the redaction. phase_callback(phase_name), if given, is called at
      major phase transitions that can't be given per-page progress (e.g.
      the final disk write) -- also with exceptions swallowed.

  verify_redaction(input_path, output_path, matches, progress_callback=None) -> dict
      Independently re-checks the output before it's trusted. Returns
      {'ok': bool, 'issues': list[str]}. A non-ok result means the GUI
      will refuse to hand back the file. progress_callback behaves as
      above.

  write_key(key_path, input_path, output_path, matches) -> None
      Only required if SUPPORTS_KEY_FILE is True. Writes a file that lets
      every placeholder be reversed back to the real value it replaced.

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
