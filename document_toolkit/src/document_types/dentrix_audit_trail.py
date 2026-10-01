"""
Document-type plugin: Dentrix "Audit Trail Report" PDFs.

This is the original, standalone engine this whole app started as (it
predates the multi-document-type framework in document_types/__init__.py
and base.py) -- ported into plugin form with the smallest possible diff,
specifically to avoid touching any of the parsing/redaction logic itself.
That logic has been through many rounds of real-world bug fixes (page
breaks, false-positive page numbers, retry-margin corruption, performance
work at 1000+ page scale, and more), each validated against the regression
fixture in test/synthetic_audit_trail.pdf -- so this port keeps every
internal name, field, and function exactly as tested, and only ADDS the
plugin-contract metadata and thin generic-name aliases documented in
base.py. See base.py for what any of that means if you're adding a new
document type and using this as a reference.

Report structure this targets (confirmed against a real, sanitized sample):

    DATE: MM/DD/YYYY    TIME: HH:MM:SS    USER: nnnn    TYPE: <type>
    MM/DD/YYYY LastName, FirstName    <description>    <amount>    <code>
    MM/DD/YYYY LastName, FirstName    <description>    <amount>    <code>   (repeats if split)

A "DATE:" header line is followed by one or more detail lines. Each detail
line repeats the transaction date and then shows the patient/guarantor name
as "Last, First" before further tab-separated columns. This module finds
every such name occurrence (by its exact position on the page, using the
PDF's real text layer -- not OCR), assigns each distinct name a sequential
<Patient N> placeholder (first-seen order, consistent for the whole file),
and can apply a TRUE redaction: the underlying name text is removed from
the PDF's content stream (not just visually covered), and the placeholder
is drawn in its place.

Nothing here modifies a file until apply_redactions() is called -- scanning
only ever reports what it found, so the caller (the GUI) can show the user
a review list first.

Regression check: run test/make_test_pdf.py (if test/synthetic_audit_trail.pdf
ever needs regenerating), then scan_pdf() on it should return exactly 11
matches and 1 unmatched line (an "Audit #: 50" footer line), and
verify_redaction() on the result should report ok=True.
"""
import csv
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

import fitz  # PyMuPDF

# --------------------------------------------------------------------
# Plugin contract metadata -- see document_types/base.py for what each
# of these means and how the GUI uses them.
# --------------------------------------------------------------------
ID = "dentrix_audit_trail"
DISPLAY_NAME = "Dentrix Audit Trail Report"
ENTITY_LABEL = "Patient"
FILE_TYPES = [("PDF files", "*.pdf"), ("All files", "*.*")]
CONTEXT_COLUMN_LABEL = "Entry (Date / Time / User)"
REVIEW_HINT = (
    "These lines sat directly under a 'DATE:' entry but didn't match the "
    "expected 'Last, First' pattern, so they will NOT be redacted. "
    "Check each one -- it may contain a name in an unexpected format."
)
NO_MATCHES_WARNING = (
    "No 'DATE:' audit entries were found in this PDF. This document type "
    "may not match this report's layout -- nothing will be redacted."
)
SUPPORTS_KEY_FILE = True
# See base.py's "BATCH REDACTION" section and assign_patient_numbers()'s
# `existing_map` parameter above -- main_gui.py's Batch Redact tab checks
# this before offering batch mode for a document type.
SUPPORTS_BATCH = True

DATE_TOKEN_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")
HEADER_RE = re.compile(
    r"^DATE:\s*\d{1,2}/\d{1,2}/\d{2,4}\s+TIME:\s*\d{1,2}:\d{2}:\d{2}\s+USER:\s*\S+\s+TYPE:",
    re.IGNORECASE,
)
# The body character class below (after the required leading letter) covers
# the normal name punctuation (apostrophe, hyphen, period) PLUS a handful of
# stray characters real-world Dentrix data has been seen gluing into an
# otherwise ordinary name: a slash (e.g. a joint/split-file name typo like
# "Smith/Jones, John"), a backslash, an ampersand, and an underscore. This
# class of "erroneous character embedded in an otherwise name-shaped word"
# matters a great deal here: _find_name_span() below validates EVERY word it
# already collected as part of the name, and if even ONE of them fails this
# check, the WHOLE line is rejected as unmatched -- meaning the name doesn't
# get redacted at all, silently, rather than just that one odd character
# being left alone. Tolerating a stray character here so the token still
# reads as "recognizably a name" is far safer than that all-or-nothing
# failure. This only affects which CLOSE, name-column-adjacent words are
# accepted as part of the name (see _find_name_span()'s gap-threshold logic,
# which decides which words are even candidates in the first place) -- it
# doesn't change what counts as a name-column word to begin with.
NAME_TOKEN_RE = re.compile(r"^[0-9]*[A-Za-z][A-Za-z'\-./\\&_]*,?$")
# A token that's a parenthetical aside attached to a name, e.g. a preferred
# name/nickname: "(Marshall)". Optional leading digits mirrors NAME_TOKEN_RE
# in case a chart number is glued onto it too.
PAREN_TOKEN_RE = re.compile(r"^\(?[0-9]*[A-Za-z][A-Za-z'\-./\\&_]*\)?,?$")
# A hyphen/dash standing alone as its own word -- happens when a hyphenated
# double-barreled last name has spaces around the hyphen, e.g.
# "Marshall - Harris, Ethan" tokenizes as ["Marshall", "-", "Harris,", ...].
HYPHEN_TOKENS = {"-", "‐", "‑", "‒", "–", "—"}
# A middle initial -- a single letter, optionally followed by a period
# and/or a comma, e.g. "B", "B.", "B,". Used by
# _trim_to_plausible_name_words() as a structural continuation marker,
# the same role HYPHEN_TOKENS and a parenthetical nickname play there:
# unlike a real word (which could just as easily be the start of the
# report's description column, e.g. "Payment"), a single bare letter is
# essentially never how an English description word begins, so it's a
# safe, distinct signal that this is "Last, First MI", not the start of
# unrelated column content.
MIDDLE_INITIAL_RE = re.compile(r"^[A-Za-z]\.?,?$")
# A bare numeric token (e.g. a chart/ID number glued on as a separate word
# rather than fused onto the name itself) is only plausible as the very
# first word of a name -- allowing it anywhere else risks swallowing an
# unrelated number from later in the line.
LEADING_DIGITS_RE = re.compile(r"^[0-9]+$")
# A dollar-amount-shaped token (the report's own "amount" column, e.g.
# "-20.00", "1,234.56", "(50.00)", "$99.99") can never legitimately be part
# of a person's name -- unlike a bare chart-number token, this doesn't need
# a "first word only" carve-out, since a name is never followed by a
# standalone dollar figure either. This exists because the gap-based
# column-break detection just below (see _extract_position_words()) was
# found, on a real report, to NOT reliably separate the name column from
# the description/amount columns that follow it -- some report layouts
# space every column boundary about as evenly as the words within a field,
# so "this gap is much wider than a normal word gap" never fires, and the
# sweep runs straight through into real transaction content. An amount
# token is an unambiguous, layout-independent stop signal regardless of
# whether the gap heuristic caught the boundary or not.
AMOUNT_TOKEN_RE = re.compile(r"^\(?-?\$?[0-9][0-9,]*\.[0-9]{2}\)?$")

Y_TOLERANCE = 2.0  # points; words within this vertical distance are treated as one line
MIN_GAP_FOR_COLUMN_BREAK = 8.0  # points; absolute floor for "this is a new column"
GAP_MULTIPLIER = 3.0  # a gap this many times the line's typical word-spacing = column break
NAME_RECT_PADDING = 0.0  # vertical padding must stay 0: PyMuPDF's redaction can corrupt
                          # unrelated text elsewhere in the same content "block" if the
                          # redaction rectangle's Y-range touches even a sliver of a
                          # neighboring line's bounding box (confirmed via testing -- a
                          # 1pt vertical pad was enough to blank out part of the DATE
                          # header line sitting directly above a redacted name). The
                          # tight word bbox from get_text('words') is already precise
                          # enough with no padding. Horizontal-only extension (to make
                          # room for the placeholder label) is safe since it stays within
                          # the same line's Y-range.


@dataclass
class NameMatch:
    page_index: int
    rect: "fitz.Rect"          # tight box around just the name text
    right_boundary: float      # x-position where the next column starts (or page edge)
    raw_name: str              # text as extracted, e.g. "Smith, John"
    header_context: Optional[str] = None  # nearest preceding "DATE: ... TYPE: ..." line, for review display
    patient_no: Optional[int] = field(default=None)
    # Sampled from the redacted text's own rawdict span during
    # apply_redactions() (None until then) -- captured purely so the name
    # key CSV can record it, letting the Un-redact tab draw a restored name
    # back in something closer to the original font/size/color instead of
    # always falling back to plain black Helvetica. See
    # _sample_original_style() and write_name_key().
    original_font: Optional[str] = field(default=None)
    original_size: Optional[float] = field(default=None)
    original_color: Optional[int] = field(default=None)
    original_flags: Optional[int] = field(default=None)

    # Generic aliases so the shared GUI (main_gui.py) can read any plugin's
    # matches through the same names (see document_types.base.Match) without
    # this module's own, already-tested field names having to change.
    @property
    def display_text(self):
        return self.raw_name

    @property
    def context(self):
        return self.header_context

    @property
    def placeholder_no(self):
        return self.patient_no


def _group_words_into_lines(page):
    """Return list of lines; each line is a list of word tuples sorted left-to-right.
    A word tuple is (x0, y0, x1, y1, text, block_no, line_no, word_no) as returned
    by page.get_text('words'). Lines are grouped purely by vertical position, which
    is robust regardless of how the source PDF chose to break text into blocks.
    """
    words = page.get_text("words")
    if not words:
        return []
    words = sorted(words, key=lambda w: (w[1], w[0]))
    lines = []
    cur = [words[0]]
    ref_y = words[0][1]
    for w in words[1:]:
        if abs(w[1] - ref_y) <= Y_TOLERANCE:
            cur.append(w)
        else:
            lines.append(sorted(cur, key=lambda ww: ww[0]))
            cur = [w]
            ref_y = w[1]
    lines.append(sorted(cur, key=lambda ww: ww[0]))
    return lines


def _line_text(line_words):
    return " ".join(w[4] for w in line_words)


def _is_name_token(tok: str, position: int) -> bool:
    """Whether a word looks like it could be (part of) a person's name,
    covering the format variations seen in real Dentrix data: hyphenated
    or multi-word last names, a parenthetical nickname/preferred name
    (e.g. "Harris, Ethan (Marshall)"), a hyphen floating as its own word
    when spaced ("Marshall - Harris, Ethan"), and a chart/ID number glued
    onto the start of the name ("33Marshall-Harris, Ethan").
    """
    if NAME_TOKEN_RE.match(tok):
        return True
    if PAREN_TOKEN_RE.match(tok):
        return True
    if tok in HYPHEN_TOKENS:
        return True
    if position == 0 and LEADING_DIGITS_RE.match(tok):
        return True
    return False


def _trim_to_plausible_name_words(words):
    """Trims a raw, gap-swept word list (see the loop in
    _extract_position_words() just below) down to whatever plausibly
    belongs to a name, using the report's OWN structural convention --
    "Last, First", optionally continued by a spaced hyphen (a
    double-barreled surname, e.g. "Marshall - Harris") or a parenthetical
    nickname -- rather than trusting the gap-based loop to have found the
    true column boundary on its own.

    This exists because that gap-based detection was found, on a real
    report, to not reliably separate the name column from the
    description column that follows it: some report layouts space every
    column boundary about as evenly as the words within a single field --
    sometimes literally a single space everywhere, name-internal or
    column-to-column alike -- so "this gap is much wider than a normal
    word gap" never fires, and the sweep runs on into ordinary
    description words. Those words pass _is_name_token()'s shape rules
    just as easily as a real name does (a word made of plain letters is a
    word made of plain letters, whether it's someone's surname or the
    word "Payment"), so shape validation downstream in _find_name_span()
    doesn't catch this either -- the whole swept span, name and
    description and sometimes even the amount, would otherwise be
    accepted as "the name" and redacted (or, from
    find_unresolved_name_positions(), flagged for review) as one piece.

    A name, though, reliably SIGNALS its own continuation: a trailing
    comma inviting the next word ("Last," -> First), a bare hyphen
    standing as its own word, a parenthetical nickname, or a middle
    initial (a single letter, optionally with a trailing period/comma --
    see MIDDLE_INITIAL_RE; "Smith, John B" is common enough in a
    guarantor/patient name field that this needed its own carve-out,
    same as the others). An ordinary word with NONE of those markers,
    once the name already has its first two words in hand, is exactly
    what an unrelated word from the next column looks like -- so it's
    excluded here rather than assumed to be more of the name.

    This ONLY EVER SHORTENS `words` (or returns it unchanged) -- it never
    adds words the gap-based loop didn't already include -- so it can
    never turn a name the gap logic already found correctly (2 or fewer
    words, the overwhelmingly common case) into something narrower than
    it should be, and it never causes a match that gap detection alone
    already got right to be missed.
    """
    if len(words) <= 2:
        return words
    kept = words[:2]
    idx = 2
    while idx < len(words):
        prev_tok = words[idx - 1][4]
        cur_tok = words[idx][4]
        prev_ends_comma = prev_tok.rstrip().endswith(",")
        prev_is_hyphen = prev_tok in HYPHEN_TOKENS
        # NOTE: deliberately NOT PAREN_TOKEN_RE here -- its "(" and ")"
        # are themselves optional (so it also matches a plain word with
        # no parenthesis at all, by design, for use inside
        # _is_name_token()), which would make this check treat every
        # ordinary word as a "marker" and defeat the whole trim. An
        # actual parenthesis character is what signals a parenthetical
        # nickname continuation here.
        cur_is_marker = (
            cur_tok in HYPHEN_TOKENS
            or "(" in cur_tok
            or ")" in cur_tok
            or bool(MIDDLE_INITIAL_RE.match(cur_tok))
        )
        if prev_ends_comma or prev_is_hyphen or cur_is_marker:
            kept.append(words[idx])
            idx += 1
            continue
        break
    return kept


def _extract_position_words(line_words):
    """Phase 1 of _find_name_span(): figures out which words occupy the
    "name column" on a detail line -- purely by horizontal position (a
    leading date token, then a run of words separated by less than this
    line's column-break gap, additionally trimmed by
    _trim_to_plausible_name_words() -- see that function for why) --
    WITHOUT checking whether those words look like a person's name.
    Returns (words, right_boundary) or None if this doesn't even look
    like a detail line (no leading date token, or nothing follows it).

    This is split out from _find_name_span() so the SAME positional logic
    -- which column the report format guarantees the name sits in -- can
    be reused by find_unresolved_name_positions() to check what's sitting
    there AFTER redaction, regardless of whether it happens to be
    name-shaped. _find_name_span() adds the shape validation on top of
    this; find_unresolved_name_positions() deliberately doesn't, since its
    whole point is to catch whatever the shape rules didn't anticipate.
    """
    if len(line_words) < 2:
        return None
    if not DATE_TOKEN_RE.match(line_words[0][4]):
        return None

    gaps = [line_words[i][0] - line_words[i - 1][2] for i in range(1, len(line_words))]
    if not gaps:
        return None
    sorted_gaps = sorted(gaps)
    median_gap = sorted_gaps[len(sorted_gaps) // 2]
    threshold = max(median_gap * GAP_MULTIPLIER, MIN_GAP_FOR_COLUMN_BREAK)

    idx = 1
    words = []
    right_boundary = None
    while idx < len(line_words):
        tok = line_words[idx][4]
        gap_before = line_words[idx][0] - line_words[idx - 1][2]
        if words and gap_before > threshold:
            right_boundary = line_words[idx][0]
            break
        # A pure-digit word after we already have at least one name word
        # looks like a chart/ID number appended AFTER the name (e.g.
        # "Marshall-Harris, Ethan 839") rather than part of it -- stop the
        # name here so that number is left alone and not redacted. (A
        # leading digit -- the very first word -- is handled separately
        # below, since that's the "chart number glued onto the front"
        # case and should stay part of the name.)
        if words and LEADING_DIGITS_RE.match(tok):
            right_boundary = line_words[idx][0]
            break
        # A dollar-amount token can never be part of a name, regardless of
        # how the gap-based check above scored the space before it -- see
        # AMOUNT_TOKEN_RE's comment for why this is needed as its own,
        # layout-independent stop condition.
        if words and AMOUNT_TOKEN_RE.match(tok):
            right_boundary = line_words[idx][0]
            break
        words.append(line_words[idx])
        idx += 1
    if right_boundary is None:
        # name ran to the end of the line with no further column after it
        right_boundary = line_words[-1][2] + 200  # generous room; caller clamps to page width

    if not words:
        return None

    trimmed_words = _trim_to_plausible_name_words(words)
    if len(trimmed_words) < len(words):
        # The gap-based loop above swept in content past the name (see
        # _trim_to_plausible_name_words()) -- pull the right boundary
        # back to right where that excluded content starts, so the
        # caller's redaction rectangle (_redact_rect_for_match()) can
        # never reach into and erase that content too.
        right_boundary = words[len(trimmed_words)][0]
        words = trimmed_words

    return words, right_boundary


def _find_name_span(line_words):
    """If this line looks like 'MM/DD/YYYY Last, First ...<column break>...',
    return (name_text, rect, right_boundary_x). Otherwise return None.
    """
    extracted = _extract_position_words(line_words)
    if extracted is None:
        return None
    name_words, right_boundary = extracted

    # A comma ("Last, First") is the normal, expected shape and the
    # strongest signal this is really a name -- but Dentrix data has also
    # been seen storing the name as "Last First" with no comma at all. Since
    # this position in the line is structurally guaranteed by the report
    # format to be the guarantor/patient name (nothing else appears there),
    # word-shape validation below is relied on instead of requiring a comma.
    for i, w in enumerate(name_words):
        tok = w[4]
        if not _is_name_token(tok, i):
            return None
    # A real name has a letter in it somewhere. Without this, a line whose
    # only "name" word is a bare number -- a page number, an invoice/chart
    # count, anything of that shape sitting right after something
    # date-token-like -- passes every check above on its own (the leading-
    # digit allowance exists for a chart number FUSED onto the front of an
    # actual name, e.g. "33Marshall-Harris", not for a number standing
    # completely alone with no name word following it) and gets accepted
    # as a "name" consisting of nothing but digits.
    if not any(re.search(r"[A-Za-z]", w[4]) for w in name_words):
        return None

    joined = _line_text(name_words)
    x0 = name_words[0][0]
    y0 = min(w[1] for w in name_words)
    x1 = name_words[-1][2]
    y1 = max(w[3] for w in name_words)
    return joined, fitz.Rect(x0, y0, x1, y1), right_boundary


@dataclass
class UnmatchedLine:
    """A line that appeared directly under a DATE: header (so the report
    format says it should be a detail line with a name on it) but didn't
    match the expected 'MM/DD/YYYY Last, First ...' shape. Surfaced so a
    human can check whether a name was simply missed.
    """
    page_index: int
    header_context: Optional[str]
    raw_text: str

    @property
    def context(self):
        return self.header_context


CARRYOVER_HEADER_LABEL = "(entry split across a page break -- header was on the previous page)"


def scan_pdf(path: str):
    """Open the PDF read-only and find every name occurrence. Does not modify anything.

    Returns (matches, unmatched_lines, page_count). unmatched_lines are ANY
    line that sat within a DATE: header's block (i.e. after that header and
    before the next one) but didn't parse as a name line -- review these
    before trusting the output, since they may indicate a name the automatic
    scan missed (e.g. an unexpected formatting variant, including a detail
    line that doesn't repeat the leading date the way most do). Every such
    line is surfaced, not just ones that happen to start with something
    date-shaped -- a narrower filter here previously let some lines (and the
    name on them) disappear with no trace in either list.

    Whether an entry's block is "open" carries over across a page boundary.
    On a large report, a single entry's DATE:/TIME:/USER:/TYPE: header can
    land at the very bottom of one page with its detail line (the one
    actually carrying the name) pushed onto the next -- and that next page
    starts with nothing but the report's own repeating title block (run
    date, "AUDIT TRAIL REPORT", practice name, "Page: N") before any content
    of its own. Previously, every page unconditionally started as "not yet
    in a block" until ITS OWN first header appeared, which meant an orphaned
    detail line pushed onto a new page like this was invisible -- not
    redacted, not listed in Needs Review, nothing, because the line before
    the page's own header was silently skipped outright. Carrying the "still
    in a block" state across the page break fixes that: this page stays
    receptive to a name line right from its first line. To avoid flooding
    Needs Review with the report's own boilerplate title lines on every
    single page as a side effect, the narrower "only flag something
    date-shaped" rule is used for this carried-over leading stretch, same as
    the very first page's title block already effectively gets (since it
    never matches a header and is skipped) -- the wider "flag anything
    unmatched" rule only applies from that page's own first real header
    onward, same as before.
    """
    doc = fitz.open(path)
    matches: List[NameMatch] = []
    unmatched: List[UnmatchedLine] = []
    try:
        page_count = len(doc)
        block_carried_over = False
        for page_index in range(page_count):
            page = doc[page_index]
            lines = _group_words_into_lines(page)
            current_header = None
            in_block = block_carried_over
            pending_carryover = block_carried_over
            for line_words in lines:
                text = _line_text(line_words)
                if HEADER_RE.match(text):
                    current_header = text
                    in_block = True
                    pending_carryover = False
                    continue
                if not in_block:
                    # Lines before the first header (report title/date/page
                    # footer) are never detail lines -- skip them entirely.
                    continue
                found = _find_name_span(line_words)
                if found:
                    name_text, rect, right_boundary = found
                    page_width = page.rect.width
                    right_boundary = min(right_boundary, page_width - 20)
                    matches.append(
                        NameMatch(
                            page_index=page_index,
                            rect=rect,
                            right_boundary=right_boundary,
                            raw_name=name_text,
                            header_context=current_header or CARRYOVER_HEADER_LABEL,
                        )
                    )
                    pending_carryover = False
                elif pending_carryover:
                    # This page's own header hasn't appeared yet, and a
                    # block was carried over from the previous page -- only
                    # flag something that at least starts with a repeated
                    # date AND has something after it (a lone date token by
                    # itself, like the report's own run-date line at the top
                    # of every page, can never be a name line -- same
                    # 2-word minimum _find_name_span itself requires), so
                    # ordinary report-title boilerplate at the top of the
                    # page doesn't get listed on every single page. A line
                    # with no letters anywhere is also skipped here: a real
                    # detail line always has a description/amount/code
                    # after the name, so "date-shaped token, then nothing
                    # but more digits" can only be the report's own
                    # date-stamp/page-number line, never a genuine entry.
                    if (
                        len(line_words) >= 2
                        and DATE_TOKEN_RE.match(line_words[0][4])
                        and re.search(r"[A-Za-z]", text)
                    ):
                        unmatched.append(
                            UnmatchedLine(
                                page_index=page_index,
                                header_context=CARRYOVER_HEADER_LABEL,
                                raw_text=text,
                            )
                        )
                elif line_words:
                    # Sat inside this header's block but didn't parse as a
                    # name line -- flag it regardless of whether it happens
                    # to start with a date, since requiring that was exactly
                    # what let some real-world lines vanish silently.
                    unmatched.append(
                        UnmatchedLine(
                            page_index=page_index,
                            header_context=current_header,
                            raw_text=text,
                        )
                    )
            block_carried_over = in_block
    finally:
        doc.close()
    return matches, unmatched, page_count


_POSITION_PLACEHOLDER_RE = re.compile(r"^<[^<>]+ \d+>$")  # same shape as _KEY_PLACEHOLDER_RE, defined below

# --------------------------------------------------------------------
# Manual override list for find_unresolved_name_positions(): report
# vocabulary that structurally lands in the guaranteed "name position" on
# some layouts (a payment-method column that starts close enough to the
# name column that a report-specific quirk puts it there, a report
# variant this module hasn't seen yet, etc.) but is never actually a
# person's name -- "Credit Card" being the first confirmed real-world
# case. No fixed word list can anticipate every practice-management
# system's report vocabulary, so rather than keep chasing individual
# values here one bug report at a time, this lets a reviewer permanently
# silence a specific recurring false positive themselves, the moment they
# see it, without needing a code change.
#
# Stored as a small JSON file NEXT TO THIS MODULE when running from
# source (not a user-profile/app-data folder) so it's easy to find, back
# up, or copy to a colleague's install, and survives an application
# update that replaces this .py file -- see _overrides_path() for where
# "next to this module" actually resolves to in the shipped .exe, which
# is NOT the same folder a developer would find this .py file in.
# Entries are matched by EXACT text (case-insensitively,
# whitespace-collapsed) against what find_unresolved_name_positions()
# would otherwise flag -- not a substring/prefix match -- so adding
# "Credit Card" only silences that literal recurring value and can never
# accidentally swallow a real name that merely contains those words.
# --------------------------------------------------------------------
_OVERRIDES_FILENAME = "dentrix_name_position_overrides.json"


def _overrides_path() -> str:
    """Where the override JSON file lives. Deliberately NOT simply "next
    to this .py file" via __file__ -- that breaks under the app's real
    deployed form.

    DocumentRedactor.exe is built by PyInstaller in --onefile mode (see
    DocumentRedactor.spec: a.binaries/a.datas go straight into EXE(...)
    with no COLLECT step). A onefile build extracts every bundled module
    -- including this one -- into a FRESH TEMPORARY DIRECTORY
    (sys._MEIPASS) each time the .exe launches, and deletes that
    directory again on exit. os.path.dirname(os.path.abspath(__file__))
    inside a frozen build resolves to that throwaway extraction folder:
    a file "saved" there vanishes the moment the app closes (so an
    override a reviewer just added would never survive a restart), and a
    file placed there ahead of time is never even seen in the first
    place (nothing outside the bundled .py/.pyc files gets extracted
    there at all). Confirmed the hard way: a manually pre-seeded
    "Credit Card" override kept getting flagged anyway because the
    running .exe was never looking at the same file this module's source
    tree has.

    sys.executable, by contrast, is the real, permanent path to
    DocumentRedactor.exe wherever the user actually put it on disk -- so
    when frozen (sys.frozen is set by PyInstaller), the override file
    lives next to THAT instead, which is the closest persistent
    equivalent "next to the module" has once there's no separate .py
    file on disk to sit beside. Running from source (sys.frozen unset,
    e.g. every test in this project) keeps the original, simpler
    behavior of sitting next to this .py file.
    """
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, _OVERRIDES_FILENAME)


def _normalize_override_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip()).lower()


def load_name_position_overrides() -> set:
    """Returns the set of normalized (lowercased, whitespace-collapsed)
    values a reviewer has manually marked as "not a name" -- see
    find_unresolved_name_positions()'s use of this. Returns an empty set
    if the file doesn't exist yet or can't be read/parsed (a broken or
    missing override file must never block the self-check itself -- it
    just means nothing is overridden yet).
    """
    path = _overrides_path()
    if not os.path.exists(path):
        return set()
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return {_normalize_override_text(str(v)) for v in raw if str(v).strip()}
    except Exception:  # noqa: BLE001
        return set()


def _load_raw_overrides_list() -> list:
    path = _overrides_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return [str(v) for v in raw if str(v).strip()]
    except Exception:  # noqa: BLE001
        return []


def add_name_position_override(value: str) -> None:
    """Permanently adds one more "this is not a name" exception (see
    load_name_position_overrides()), persisted for every future run --
    used by main_gui.py's leak-review screen when a reviewer checks "also
    add this to the ignore list" on an occurrence. A no-op for
    blank/whitespace-only input. Safe to call with a value already
    present (case-insensitively) -- it won't be duplicated.
    """
    value = value.strip()
    if not value:
        return
    existing_raw = _load_raw_overrides_list()
    existing_normalized = {_normalize_override_text(v) for v in existing_raw}
    if _normalize_override_text(value) not in existing_normalized:
        existing_raw.append(value)
        with open(_overrides_path(), "w", encoding="utf-8") as f:
            json.dump(existing_raw, f, indent=2)


def remove_name_position_override(value: str) -> None:
    """Removes a previously-added override (case-insensitive, exact
    match) -- e.g. if a reviewer added one by mistake. A no-op if it
    isn't present.
    """
    norm = _normalize_override_text(value)
    existing_raw = _load_raw_overrides_list()
    kept = [v for v in existing_raw if _normalize_override_text(v) != norm]
    if len(kept) != len(existing_raw):
        with open(_overrides_path(), "w", encoding="utf-8") as f:
            json.dump(kept, f, indent=2)


# A real name -- even a hyphenated one with a parenthetical nickname, e.g.
# "Marshall - Harris, Ethan (Marshall)" -- essentially never runs past this
# many words. AMOUNT_TOKEN_RE's stop condition in _extract_position_words()
# closes the most dangerous gap (a dollar figure ending up "in the name"),
# but it does nothing about a report/description column's ORDINARY WORDS
# (e.g. "Credit Card Payment - Thank You") getting swept in ahead of it --
# those pass _is_name_token()'s own shape rules just fine on their own, so
# shape-checking this candidate wouldn't catch it either. A word-count
# ceiling this generous can't be crossed by any real name, so it's a safe,
# layout-independent signal that the gap-based column-break detection
# above didn't find the true boundary on THIS report's specific column
# spacing and swept real, non-PHI report content in beside (or instead of)
# the name. See the 'suspicious' field below for what this changes.
MAX_PLAUSIBLE_NAME_WORDS = 5


def find_unresolved_name_positions(output_path: str, ignore_values: set = None) -> list:
    """Independent, STRUCTURAL safety net over an ALREADY-REDACTED file,
    meant to be run alongside (and merged into) verify_redaction()'s
    name-based leak check.

    The report format guarantees that the "name column" on every detail
    line (see _extract_position_words()) holds exactly one thing: the
    patient/guarantor's name before redaction, or its <Patient N>
    placeholder after a successful one. Nothing else is ever supposed to
    be there. This re-scans the redacted output's own detail lines using
    that SAME positional logic -- but, unlike scan_pdf(), without
    requiring the text found there to look name-shaped -- and flags any
    position that isn't the placeholder.

    This matters because scan_pdf()'s _is_name_token() shape rules, no
    matter how many real-world formatting quirks get added to them one at
    a time (stray punctuation, unusual capitalization, and so on), can
    never anticipate every way a name might be entered wrong in the
    source system. A name garbled badly enough to fail every shape check
    was never a candidate for redaction in the first place, so it's
    invisible to both the "known name still present" leak check (it was
    never in `matches` to look for) AND a purely shape-based scan of the
    output. Checking the one thing the report format itself guarantees --
    there is ALWAYS something in this exact position -- closes that gap:
    whatever is sitting there that isn't the placeholder is worth a
    human's attention, regardless of why the normal scan missed it.

    Returns the same shape as verify_redaction()'s 'leaks' field --
    [{'name': str, 'occurrences': [{'page_index': int, 'rect': [x0, y0,
    x1, y1], 'context': str or None, 'suspicious': bool}, ...]}, ...], one
    entry per distinct leftover text (occurrences of the exact same text
    grouped together) -- so it can be merged directly into that field and
    reuses the existing manual-review screen with no UI changes needed
    beyond reading 'suspicious'. 'name' here just means "whatever text was
    found in the name position", not necessarily a validated person's name
    -- that's the point.

    'suspicious' (see MAX_PLAUSIBLE_NAME_WORDS) is set when the text found
    runs longer than any real name plausibly would -- the report format
    only guarantees SOMETHING sits in this position, not that the gap-based
    logic in _extract_position_words() correctly found where the name
    itself ends and the next column begins on THIS report's specific
    layout. A layout where every column boundary is spaced about as evenly
    as the words within a field (seen in the wild) defeats that "gap much
    wider than normal" heuristic entirely, and the sweep runs on into real
    description/amount text instead of stopping at the name. Reporting the
    occurrence either way is still correct -- there is genuinely something
    other than a placeholder sitting in a position the format guarantees
    should hold one -- but main_gui.py's review screen treats 'suspicious'
    occurrences differently: it defaults them to Ignore instead of Redact
    and warns the reviewer to check the text isn't more than just the
    name, since blindly redacting an over-wide catch like this would
    destroy real, non-PHI report content, not just fix a leak.

    Unlike scan_pdf()'s 'unmatched' list (which reports the WHOLE line
    text for anything under a header that didn't parse as a name), this
    only reports the ISOLATED name-position text, and only when that
    position is neither blank nor placeholder-shaped -- so an ordinary,
    correctly-redacted line (position holds "<Patient 3>") produces
    nothing, and this never floods the result with every successfully
    redacted line in the file.

    'ignore_values', if given, is a set of already-normalized (see
    _normalize_override_text()) strings to silently skip -- text that a
    reviewer has previously confirmed is report vocabulary, not a name
    (e.g. "Credit Card", a payment-method value that can legitimately
    land in the name position on some report layouts -- see
    load_name_position_overrides()). Defaults to whatever's currently
    persisted via load_name_position_overrides() so every caller gets
    the manual-override list automatically without having to know it
    exists; pass an explicit set (or set()) only to override that, e.g.
    in tests that want to see what WOULD be flagged before any
    overrides are applied.
    """
    if ignore_values is None:
        ignore_values = load_name_position_overrides()
    doc = fitz.open(output_path)
    found = {}
    try:
        block_carried_over = False
        for page_index in range(len(doc)):
            page = doc[page_index]
            lines = _group_words_into_lines(page)
            current_header = None
            in_block = block_carried_over
            for line_words in lines:
                text = _line_text(line_words)
                if HEADER_RE.match(text):
                    current_header = text
                    in_block = True
                    continue
                if not in_block:
                    continue
                extracted = _extract_position_words(line_words)
                if extracted is None:
                    continue
                words, _right_boundary = extracted
                position_text = _line_text(words)
                if _POSITION_PLACEHOLDER_RE.match(position_text):
                    continue  # correctly redacted -- nothing to flag
                if _normalize_override_text(position_text) in ignore_values:
                    continue  # a reviewer already confirmed this exact text isn't a name
                x0 = words[0][0]
                y0 = min(w[1] for w in words)
                x1 = words[-1][2]
                y1 = max(w[3] for w in words)
                found.setdefault(position_text, []).append({
                    "page_index": page_index,
                    "rect": [x0, y0, x1, y1],
                    "context": current_header or CARRYOVER_HEADER_LABEL,
                    "suspicious": len(words) > MAX_PLAUSIBLE_NAME_WORDS,
                })
            block_carried_over = in_block
    finally:
        doc.close()
    return [{"name": text, "occurrences": occs} for text, occs in found.items()]


def normalize_name(name: str) -> str:
    return re.sub(r"\s+", " ", name.strip()).lower().rstrip(",")


def assign_patient_numbers(matches: List[NameMatch], existing_map: dict = None) -> dict:
    """Mutates each match's .patient_no in place (first-seen order within
    this call's matches, continuing any earlier numbering already present
    in existing_map). Returns the updated name->number mapping used, for
    display/audit purposes -- pass this back in as existing_map on a
    later call (e.g. the next file in a batch) to keep the same real name
    assigned the same number across every file processed together,
    instead of each file starting its numbering over from 1. See
    base.py's "BATCH REDACTION" section. Omit (or pass None) for the
    original single-file behavior, unchanged.
    """
    mapping = dict(existing_map) if existing_map else {}
    for m in matches:
        norm = normalize_name(m.raw_name)
        if norm not in mapping:
            mapping[norm] = len(mapping) + 1
        m.patient_no = mapping[norm]
    return mapping


def _redact_rect_for_match(m: "NameMatch", margin: float = 0.0) -> "fitz.Rect":
    return fitz.Rect(
        m.rect.x0 - NAME_RECT_PADDING - margin,
        m.rect.y0 - NAME_RECT_PADDING - margin,
        max(m.rect.x1, m.right_boundary - 3) + NAME_RECT_PADDING + margin,
        m.rect.y1 + NAME_RECT_PADDING + margin,
    )


def _collect_kept_runs(raw: dict, redact_rects: List["fitz.Rect"]):
    """Walk a page's rawdict character data and split it into (a) 'runs' of
    consecutive characters to redraw (batched by contiguous stretches within
    the same span) and (b) a single lowercased, whitespace-collapsed string
    of all KEPT text, used to sanity-check that nothing targeted for
    redaction survived. A character is dropped only if its CENTER POINT
    falls inside a redact rect -- not merely if its bbox overlaps one.
    Adjacent report lines routinely have bounding boxes that overlap each
    other by a fraction of a point (normal font-metrics slop), so an
    overlap-based test drops characters from neighboring, unrelated lines.
    Center-point containment avoids that false-positive while still reliably
    catching every character that's actually part of the redacted name --
    the retry-with-growing-margin loop in apply_redactions() is what
    guards against any remaining edge-of-name characters surviving, by
    re-running this same center-point test against a slightly larger rect
    if a leak is detected.

    Performance note: this inner loop runs once per character per redact
    rect on the page (potentially millions of times on a large report), so
    the containment test is done with plain float comparisons on plain
    tuples rather than fitz.Rect.contains(fitz.Point(...)). Profiling a
    150-page/2,700-match stress file showed the fitz.Rect/fitz.Point path
    spending the vast majority of its time in Python/SWIG call overhead
    crossing into the C library on every single check, not in the
    geometry itself -- it was the direct cause of multi-minute save times
    on large real-world files. The plain-tuple version below does the
    identical containment math with none of that per-call overhead.
    """
    redact_tuples = [(r.x0, r.y0, r.x1, r.y1) for r in redact_rects]
    runs = []
    kept_text_parts = []
    for block in raw.get("blocks", []):
        if block.get("type", 0) != 0:
            continue  # non-text (image) block -- left alone entirely
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                font = span.get("font", "")
                size = span.get("size", 9.0)
                color = span.get("color", 0)
                flags = span.get("flags", 0)
                cur_run = None
                span_kept = []
                for ch in span.get("chars", []):
                    cx0, cy0, cx1, cy1 = ch["bbox"]
                    cx = (cx0 + cx1) / 2.0
                    cy = (cy0 + cy1) / 2.0
                    redacted = False
                    for rx0, ry0, rx1, ry1 in redact_tuples:
                        if rx0 <= cx <= rx1 and ry0 <= cy <= ry1:
                            redacted = True
                            break
                    if redacted:
                        cur_run = None  # break the run; drop this char
                        continue
                    span_kept.append(ch["c"])
                    if cur_run is None:
                        cur_run = {
                            "origin": ch["origin"],
                            "text": ch["c"],
                            "font": font,
                            "size": size,
                            "color": color,
                            "flags": flags,
                        }
                        runs.append(cur_run)
                    else:
                        cur_run["text"] += ch["c"]
                if span_kept:
                    kept_text_parts.append("".join(span_kept))
    kept_text = re.sub(r"\s+", " ", " ".join(kept_text_parts)).strip().lower()
    return runs, kept_text


def _sample_original_style(raw: dict, rect: "fitz.Rect"):
    """Finds the font/size/color of whatever character(s) sit inside
    `rect` in this page's rawdict data -- used to capture a redacted
    name's own original styling (before it's wiped) purely so the name
    key CSV can record it for the Un-redact tab. Only needs one
    representative character, since a single name is drawn in one
    consistent style; returns None if nothing is found there (e.g. the
    rect is slightly off due to a font-metrics edge case -- the caller
    falls back to a default rather than failing the redaction over this,
    since it's a cosmetic nicety for un-redaction, not a safety property).
    """
    for block in raw.get("blocks", []):
        if block.get("type", 0) != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                for ch in span.get("chars", []):
                    cx0, cy0, cx1, cy1 = ch["bbox"]
                    cx, cy = (cx0 + cx1) / 2.0, (cy0 + cy1) / 2.0
                    if rect.x0 <= cx <= rect.x1 and rect.y0 <= cy <= rect.y1:
                        return (
                            span.get("font", ""), span.get("size", 9.0),
                            span.get("color", 0), span.get("flags", 0),
                        )
    return None


def _names_leaked(kept_text: str, target_names: List[str]) -> List[str]:
    """Which of the target names are still findable (as a normalized
    substring) in text that's about to be kept/redrawn."""
    leaked = []
    for name in target_names:
        norm = re.sub(r"\s+", " ", name.strip()).lower()
        if norm and norm in kept_text:
            leaked.append(name)
    return leaked


LOCAL_LEAK_Y_WINDOW = 4.0  # points; must stay well under a real line's height


def _locally_leaked_matches(runs, page_matches: List["NameMatch"]) -> List["NameMatch"]:
    """Which matches still have their OWN name text reconstructable from kept
    runs sitting close to that match's own line (a tight Y window around its
    rectangle) -- i.e. a boundary character of THIS specific name survived
    its own redaction.

    This is deliberately scoped to nearby text only, unlike a page-wide
    substring search. A page-wide check is fooled by a second, unrelated
    occurrence of the same name text elsewhere on the page (a different
    entry, a note, a different section) that was never a redaction target
    in the first place -- growing that unrelated match's own rectangle can
    never make that other occurrence disappear, so a page-wide check would
    escalate the margin all the way to its ceiling for no benefit, and a
    large margin is exactly what risks reaching into a neighboring line and
    corrupting unrelated content (the original failure mode this whole
    whole-page-rebuild strategy exists to avoid). Scoping the check to the
    match's own line means the retry loop only grows a rectangle when doing
    so could plausibly help, and never over-expands chasing a leak it has
    no way to fix.
    """
    still_leaked = []
    for m in page_matches:
        y_lo = m.rect.y0 - LOCAL_LEAK_Y_WINDOW
        y_hi = m.rect.y1 + LOCAL_LEAK_Y_WINDOW
        nearby = [r for r in runs if y_lo <= r["origin"][1] <= y_hi]
        nearby.sort(key=lambda r: r["origin"][0])
        local_text = re.sub(r"\s+", " ", "".join(r["text"] for r in nearby)).strip().lower()
        norm = re.sub(r"\s+", " ", m.raw_name.strip()).lower()
        if norm and norm in local_text:
            still_leaked.append(m)
    return still_leaked


def _map_fontname(original_font: str, flags: int) -> str:
    """Map an extracted font name to one of PyMuPDF's built-in base-14 fonts
    for redrawing. We don't have the original embedded font program, so we
    approximate using the PDF text-flags bold/italic/serif/monospace bits
    (see PyMuPDF docs: 2=italic, 4=serifed, 8=monospaced, 16=bold).
    """
    bold = bool(flags & 16)
    italic = bool(flags & 2)
    serifed = bool(flags & 4)
    monospaced = bool(flags & 8)
    name_lower = (original_font or "").lower()
    if monospaced or "courier" in name_lower or "mono" in name_lower:
        base = "co"
    elif serifed or "times" in name_lower or "georgia" in name_lower or "serif" in name_lower:
        base = "ti"
    else:
        base = "he"
    if base == "he":
        suffix = {(False, False): "lv", (True, False): "bo", (False, True): "it", (True, True): "bi"}
        return "he" + suffix[(bold, italic)]
    if base == "ti":
        suffix = {(False, False): "ro", (True, False): "bo", (False, True): "it", (True, True): "bi"}
        return "ti" + suffix[(bold, italic)]
    suffix = {(False, False): "u", (True, False): "ub", (False, True): "ui", (True, True): "ubi"}
    return "co" + suffix[(bold, italic)]


def apply_redactions(
    input_path: str,
    output_path: str,
    matches: List[NameMatch],
    progress_callback=None,
    phase_callback=None,
) -> None:
    """Applies TRUE redaction: removes the underlying name text from the PDF
    (not just a black box drawn over it) and writes <Patient N> in its place.
    `matches` must already have .patient_no set (see assign_patient_numbers).
    Only matches with .patient_no set and not flagged as excluded should be passed in.

    Implementation note: this does NOT use PyMuPDF's per-rectangle
    add_redact_annot()/apply_redactions() to selectively strip just the name
    from a page's existing content. Testing found that approach unsafe: when
    a report's line spacing is tighter than the font's own ascent+descent
    (common in compact printed reports), two adjacent text lines' bounding
    boxes genuinely overlap by a point or two regardless of how tightly the
    redaction rectangle is drawn -- and PyMuPDF's redaction groups text into
    multi-line "blocks" internally, so touching one line can silently corrupt
    or fail to remove text on the neighboring line sharing that block. This
    was reproduced reliably in testing and is the leading explanation for a
    name surviving redaction.

    Instead, for any page that has at least one match, every character on
    that page is captured (exact position, font, size, color) first, then
    ALL text on the page is removed via a single whole-page redaction (no
    partial-block ambiguity is possible when nothing is left behind), and
    finally every captured character that did NOT fall inside a redacted
    name's rectangle is redrawn at its original position. This guarantees
    the target names cannot survive, since nothing survives the whole-page
    pass except what we deliberately redraw.

    If given, `progress_callback(current_page_number, total_pages)` is
    invoked once per page that actually has a redaction on it (1-based
    current_page_number), so a caller (the GUI) can show real progress
    instead of a spinner on large files. `phase_callback(phase_name)` is
    invoked once, right before the final write to disk begins, with the
    string "saving" -- that write is a single call into the PDF library
    and can't be subdivided into per-page progress, so this at least lets
    a caller distinguish "still drawing pages" from "now writing the
    finished file," rather than the UI looking stuck at 100% for however
    long that final write takes. Either callback's exceptions are
    swallowed -- a broken progress display must never interrupt or fail
    a redaction.
    """
    doc = fitz.open(input_path)
    try:
        by_page = {}
        for m in matches:
            by_page.setdefault(m.page_index, []).append(m)

        total_to_process = len(by_page)
        for done_count, (page_index, page_matches) in enumerate(by_page.items(), start=1):
            page = doc[page_index]

            # 1. Capture every character on the page before touching anything,
            # collapsing consecutive KEPT characters within the same span into
            # single runs (one insert_text call per run instead of per
            # character -- this is a large speedup on multi-hundred-page
            # files with thousands of characters, since each low-level
            # insert_text call has real overhead).
            #
            # The redaction rectangle for each match comes from get_text
            # ("words") during scanning, but this pass reads character
            # positions from get_text("rawdict") -- a different extraction
            # call. On most PDFs these agree exactly, but a real-world file
            # surfaced cases where they didn't quite line up, letting an
            # edge character survive with zero margin between the two. To
            # guard against that, this retries with a small, capped safety
            # margin around each rectangle, re-checking each time whether
            # that SAME match's own name is still reconstructable from kept
            # text near its own line -- using _locally_leaked_matches, not a
            # page-wide search. A page-wide "is this name still anywhere on
            # the page" check was tried first and found unsafe: a second,
            # unrelated occurrence of the same name elsewhere on the page
            # (one that was never a redaction target) can never be fixed by
            # growing a different match's rectangle, so that check kept
            # escalating to its ceiling for no benefit -- and a large
            # margin is exactly what risks reaching into a neighboring line
            # and corrupting unrelated content. The margin ceiling here
            # stays small (2pt) for the same reason: real reports can have
            # single-digit-point line spacing, so even the retry margin
            # must stay comfortably under a full line's height.
            raw = page.get_text("rawdict")
            runs, kept_text = [], ""
            for margin in (0.0, 0.5, 1.0, 2.0):
                redact_rects = [_redact_rect_for_match(m, margin) for m in page_matches]
                runs, kept_text = _collect_kept_runs(raw, redact_rects)
                if not _locally_leaked_matches(runs, page_matches):
                    break

            # 1b. Before wiping anything, sample each match's OWN original
            # font/size/color from this same rawdict capture -- purely so
            # write_name_key() can record it for the Un-redact tab, which
            # would otherwise have no way to know what the redacted text
            # originally looked like. This has no bearing on redaction
            # safety (the sampled rect is the tight, zero-margin match
            # rect, not involved in the retry loop above at all).
            for m in page_matches:
                style = _sample_original_style(raw, _redact_rect_for_match(m))
                if style is not None:
                    m.original_font, m.original_size, m.original_color, m.original_flags = style

            # 2. Remove ALL text on the page in one pass (no partial-block ambiguity).
            page.add_redact_annot(page.rect)
            page.apply_redactions(images=0, graphics=0, text=0)

            # 3. Redraw every surviving run, then the <Patient N> placeholders,
            # batched into one Shape so the whole page commits in a single
            # content-stream write instead of one write per insert_text call.
            #
            # Note on a dead end already tried here: Shape.insert_text()
            # internally re-resolves its fontname on every call by scanning
            # the page's font resources from scratch (Shape.insert_text ->
            # Page.insert_font -> CheckFont -> Page.get_fonts(), which
            # re-parses the page's resource dictionary each time). Explicitly
            # pre-registering each font once via page.insert_font() before
            # this loop was tried as a fix, on the theory that a warm
            # resource table would make each call's internal re-scan cheap.
            # Clean, non-profiled A/B timing on a large realistic file showed
            # it makes no measurable difference (the per-call scan happens
            # regardless of pre-registration), so that step was removed --
            # this per-call resource scan is a cost inherent to Shape's
            # public API, not something reachable from here.
            shape = page.new_shape()
            for run in runs:
                fontname = _map_fontname(run["font"], run["flags"])
                color = fitz.sRGB_to_pdf(run["color"])
                try:
                    shape.insert_text(
                        run["origin"], run["text"], fontname=fontname, fontsize=run["size"], color=color
                    )
                except Exception:
                    shape.insert_text(
                        run["origin"], run["text"], fontname="helv", fontsize=run["size"], color=color
                    )

            # 4. Draw the <Patient N> placeholders in place of each redacted name.
            for m in page_matches:
                rect = _redact_rect_for_match(m)
                label = f"<{ENTITY_LABEL} {m.patient_no}>"
                fontsize = _fit_fontsize(label, rect.width, base=9.0)
                shape.insert_text(
                    (rect.x0, rect.y1 - 1.5), label, fontname="helv", fontsize=fontsize, color=(0, 0, 0)
                )
            shape.commit()

            if progress_callback is not None:
                try:
                    progress_callback(done_count, total_to_process)
                except Exception:
                    pass

        if phase_callback is not None:
            try:
                phase_callback("saving")
            except Exception:
                pass

        # Scrub document metadata/XML metadata too -- redaction tools shouldn't
        # leave the original filename, author, or other identifying info behind
        # in the file's Info dictionary or XMP packet.
        doc.set_metadata({})
        try:
            doc.del_xml_metadata()
        except Exception:
            pass

        # garbage=1 here, not the more thorough garbage=4 used previously.
        # This isn't a redaction-safety tradeoff: apply_redactions() above
        # rewrites each redacted page's content stream in place, so the
        # removed name text is gone from the moment that call returns --
        # confirmed directly by decompressing and byte-scanning the raw,
        # un-garbage-collected output file for every redacted name and
        # finding none, at garbage=0 through garbage=4 alike. The garbage
        # level only controls how aggressively leftover, already-harmless
        # PDF objects (old font/resource entries, that kind of thing) get
        # deduplicated and dropped to shrink the file -- garbage=4's extra
        # passes (duplicate-stream detection across the whole document)
        # scale worse than the other levels on a file with many pages, and
        # are believed responsible for a save that looked hung on a large
        # real-world file. garbage=1 still drops unreferenced objects so
        # the file doesn't balloon in size, just without that expensive
        # extra pass.
        doc.save(output_path, garbage=1, deflate=True, clean=True)
    finally:
        doc.close()


def _fit_fontsize(label: str, available_width: float, base: float = 9.0) -> float:
    """Rough estimate so the placeholder text doesn't overflow the available
    column width; PyMuPDF's helv metric is ~0.5 * fontsize per character on average.
    """
    approx_width = len(label) * base * 0.5
    if approx_width <= available_width or available_width <= 0:
        return base
    scaled = base * (available_width / approx_width)
    return max(scaled, 5.0)


STRUCTURAL_MARKERS = ("DATE:", "TIME:", "USER:", "TYPE:")


def verify_redaction(
    input_path: str,
    output_path: str,
    matches: List[NameMatch],
    progress_callback=None,
) -> dict:
    """Automatic QA gate, run after every redaction before the file is handed back.

    Checks:
      1. Every redacted name (in its exact extracted form) is completely gone
         from the output's extracted text -- catches a name that silently
         failed to redact. When a name IS still found, the issue reports
         exactly which page(s) and a short snippet of surrounding text, so
         the occurrence can actually be located and diagnosed instead of
         just being reported as present somewhere in an unknown location.
         A common cause: the same patient's name appears a second time on
         the page in a spot the scan didn't recognize as a name line (so it
         was never a candidate for redaction in the first place) -- this
         will show up here as a leak even though nothing "failed."
      2. Every expected <Patient N> placeholder actually appears in the output.
      3. The count of structural report markers (DATE:, TIME:, USER:, TYPE:)
         is unchanged between input and output -- catches unrelated content
         getting silently corrupted/erased as a side effect of redacting a
         nearby line (a real failure mode seen during testing with PyMuPDF's
         redaction when a rectangle's Y-range brushes a neighboring line).
      4. Page count is unchanged.
      5. The report's own guarantee about WHERE a name lives: every detail
         line's name column (see find_unresolved_name_positions()) holds
         either the placeholder or nothing else recognizable -- catches a
         name garbled badly enough that it never matched scan_pdf()'s
         name-shape rules in the first place, so it was never even a
         candidate for redaction, and check #1 above never knew to look
         for it (it only checks names ALREADY known from `matches`).

    Returns {'ok': bool, 'issues': list[str], 'leaks': list[dict],
    'structural_issue': str or None}. A non-ok result means the file
    should NOT be trusted/delivered as-is.

    'issues' is the original human-readable summary. 'leaks' is a
    structured, per-occurrence breakdown of checks #1 and #5 (something
    still visible where a name shouldn't be, or where one is missing) --
    [{'name': str, 'occurrences': [{'page_index': int, 'rect': [x0, y0,
    x1, y1], 'context': str or None}, ...]}, ...] -- used by main_gui.py's
    manual-review screen (see apply_pdf_leak_decisions()) to actually
    locate and, if the reviewer chooses, redact each occurrence, the same
    way excel_spreadsheet.py's verify_column_redaction() 'leaks' field
    drives its own review screen. A leak from check #5 has 'name' set to
    whatever text was actually found there -- not necessarily a
    recognized person's name shape, since that's exactly what check #5
    exists to catch regardless of shape -- and apply_pdf_leak_decisions()
    assigns it a brand-new placeholder number on a 'redact' decision,
    since check #1's leaks always reuse an EXISTING number but a check #5
    leak, by definition, never had one. 'structural_issue', when set,
    flags a problem from checks #2-4 -- NOT a single leaked/unresolved
    occurrence a per-occurrence decision can fix -- so main_gui.py only
    offers manual review when 'leaks' is non-empty AND this is None;
    otherwise it falls back to the original hard block (delete the
    output, show an error, nothing saved).

    If given, `progress_callback(current_page_number, total_pages)` is
    invoked once per page (1-based) while re-extracting text for this
    check, so a caller (the GUI) can show real progress on large files.
    Any exception it raises is swallowed -- a broken progress display must
    never interrupt or fail verification.
    """
    issues = []
    doc_in = fitz.open(input_path)
    # The output file was just written a moment ago by this same process.
    # On Windows, it's common for something else -- antivirus real-time
    # scanning, or a cloud-sync client like OneDrive/Dropbox watching the
    # folder -- to grab a brief exclusive lock on a freshly-written file
    # before this can reopen it, which surfaces as an opaque OS-level
    # error that has nothing to do with the redaction itself. A few short
    # retries covers that without weakening the check in any way: it's
    # still the same independent re-open-and-verify, just not required to
    # succeed on the very first attempt.
    doc_out = None
    last_exc = None
    for attempt in range(6):
        try:
            doc_out = fitz.open(output_path)
            break
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            if attempt < 5:
                time.sleep(0.5)
    if doc_out is None:
        raise last_exc
    structural_issues = []
    try:
        if len(doc_in) != len(doc_out):
            msg = f"Page count changed: {len(doc_in)} -> {len(doc_out)}"
            issues.append(msg)
            structural_issues.append(msg)

        page_total = len(doc_out)
        in_page_texts = [p.get_text() for p in doc_in]
        out_page_texts = []
        for idx, p in enumerate(doc_out, start=1):
            out_page_texts.append(p.get_text())
            if progress_callback is not None:
                try:
                    progress_callback(idx, page_total)
                except Exception:
                    pass
        text_in_all = "\n".join(in_page_texts)
        text_out_all = "\n".join(out_page_texts)

        # Cheap pass first: one substring check per unique name against the
        # whole-document text (same cost this check always had). Only a
        # name that's ACTUALLY still present -- which should be zero in a
        # healthy run -- pays for the more expensive per-page scan needed to
        # report exactly where it is. An earlier version ran the per-page
        # scan for every name unconditionally, which on a large, real,
        # high-cardinality file (hundreds of pages x thousands of distinct
        # names) turned into millions of string searches and made this
        # check itself a significant slice of the total save time.
        # 'leaks' mirrors excel_spreadsheet.py's verify_column_redaction()
        # 'leaks' field: a STRUCTURED record of every leaked occurrence
        # (name + page + on-page rectangle + surrounding text), alongside
        # the plain-English 'issues' strings above -- this is what lets a
        # manual review screen actually locate and, if the reviewer
        # chooses, redact each occurrence, rather than just reporting that
        # something is wrong. 'structural_issue' (built from
        # structural_issues, accumulated throughout this function) flags a
        # problem that is NOT a single leaked name a per-occurrence
        # decision can fix (page count or structural-marker count changed,
        # or an expected placeholder is altogether missing) --
        # main_gui.py only offers manual review when 'leaks' is non-empty
        # AND this is None.
        leaks = []

        seen_names = sorted(set(m.raw_name for m in matches))
        leaked_names = [name for name in seen_names if name in text_out_all]
        for name in leaked_names:
            text_locations = []
            occurrences = []
            for page_idx, page_text in enumerate(out_page_texts):
                snippets_this_page = []
                start_search = 0
                while True:
                    found_at = page_text.find(name, start_search)
                    if found_at == -1:
                        break
                    ctx_start = max(0, found_at - 30)
                    ctx_end = min(len(page_text), found_at + len(name) + 30)
                    snippet = " ".join(page_text[ctx_start:ctx_end].split())
                    snippets_this_page.append(snippet)
                    text_locations.append(f"page {page_idx + 1}: \"...{snippet}...\"")
                    start_search = found_at + len(name)
                if not snippets_this_page:
                    continue
                # Locate the actual on-page rectangle(s) for this name, so
                # a manual review decision can be applied precisely (see
                # apply_pdf_leak_decisions()). This reuses PyMuPDF's own
                # text search rather than re-deriving positions from the
                # word-grouping logic scan_pdf() uses (which is tuned for
                # this report's specific "Last, First" line shape, not for
                # finding an arbitrary leaked occurrence anywhere on the
                # page).
                try:
                    rects = doc_out[page_idx].search_for(name)
                except Exception:  # noqa: BLE001
                    rects = []
                # Pair rects with snippets positionally -- both are the
                # same left-to-right, top-to-bottom occurrences of the
                # same exact string on the same page, so the counts agree
                # in the overwhelming common case. If they don't (a rare
                # extraction-method mismatch), fall back to no snippet per
                # rect rather than risk pairing the wrong context with the
                # wrong occurrence.
                paired_snippets = (
                    snippets_this_page if len(snippets_this_page) == len(rects) else [None] * len(rects)
                )
                for rect, snippet in zip(rects, paired_snippets):
                    occurrences.append({
                        "page_index": page_idx,
                        "rect": [rect.x0, rect.y0, rect.x1, rect.y1],
                        "context": snippet,
                    })
            if text_locations:
                shown = text_locations[:3]
                more = f" (+{len(text_locations) - 3} more occurrence(s))" if len(text_locations) > 3 else ""
                issues.append(
                    f"Name still present in output text: {name!r} -- " + "; ".join(shown) + more
                )
            if occurrences:
                leaks.append({"name": name, "occurrences": occurrences})

        for marker in STRUCTURAL_MARKERS:
            in_count = text_in_all.count(marker)
            out_count = text_out_all.count(marker)
            if in_count != out_count:
                msg = (
                    f"{marker!r} count changed ({in_count} -> {out_count}); "
                    f"unrelated report content may have been lost"
                )
                issues.append(msg)
                structural_issues.append(msg)

        expected_labels = sorted(set(f"<{ENTITY_LABEL} {m.patient_no}>" for m in matches))
        for label in expected_labels:
            if label not in text_out_all:
                msg = f"Expected placeholder missing from output: {label}"
                issues.append(msg)
                structural_issues.append(msg)
    finally:
        doc_in.close()
        doc_out.close()

    # Check #5 -- the structural, position-based safety net (see
    # find_unresolved_name_positions()'s docstring): independent of
    # everything above, since it never relies on `matches` at all. This
    # is what catches a name garbled badly enough that scan_pdf() never
    # recognized it as a name in the first place, so it was never a
    # candidate for redaction and check #1 above has no way to know to
    # look for it.
    #
    # Dedupe against check #1's own findings first: a name ALREADY
    # flagged as a leak on a given page (same name, same page) would
    # otherwise show up twice in 'leaks' -- once from the whole-document
    # text search above, once from this position scan -- for what is
    # really the same underlying occurrence.
    already_flagged = {(leak["name"], occ["page_index"]) for leak in leaks for occ in leak["occurrences"]}
    for position_leak in find_unresolved_name_positions(output_path):
        name = position_leak["name"]
        new_occurrences = [
            occ for occ in position_leak["occurrences"] if (name, occ["page_index"]) not in already_flagged
        ]
        if not new_occurrences:
            continue
        existing_entry = next((leak for leak in leaks if leak["name"] == name), None)
        if existing_entry is not None:
            existing_entry["occurrences"].extend(new_occurrences)
        else:
            leaks.append({"name": name, "occurrences": new_occurrences})
        issues.append(
            f"Unredacted text found in the expected name position (not a recognized name "
            f"shape, but the report format guarantees a name or placeholder belongs here): "
            f"{name!r} ({len(new_occurrences)} occurrence(s))"
        )

    return {
        "ok": len(issues) == 0,
        "issues": issues,
        "leaks": leaks,
        "structural_issue": "; ".join(structural_issues) if structural_issues else None,
    }


def apply_pdf_leak_decisions(
    output_path: str, decisions: List[dict], name_to_patient_no: dict
) -> List[NameMatch]:
    """Applies manual-review decisions for LEAKED/unresolved occurrences
    (see verify_redaction()'s 'leaks' field, which now covers both a name
    still visible somewhere AND unrecognized leftover text sitting in the
    report's known name position -- see find_unresolved_name_positions())
    directly to output_path, in place. main_gui.py's PDF leak review
    screen builds `decisions` from that same 'leaks' field: [{'name':
    str, 'page_index': int, 'rect': [x0, y0, x1, y1], 'action': 'redact'
    or 'ignore'}, ...].

    A 'redact' decision is turned into one more NameMatch and run through
    apply_redactions() itself -- the exact same whole-page-rebuild pass
    every other redaction in this file goes through, so the additional
    occurrence is GENUINELY removed from the PDF's content stream, not
    just covered with a box. This is why output_path is read back in as
    the INPUT for that call (it's already been redacted once; this is a
    second pass over the same file, targeting only the newly-found
    rectangles) and written to a temp file first, then swapped into place
    with os.replace() -- PyMuPDF refuses to save a document over the same
    path it was opened from.

    `name_to_patient_no` is {raw_name: patient_no} for every name already
    assigned a number in this file (the same mapping assign_patient_numbers()
    returned) -- a 'redact' decision reuses that name's EXISTING number if
    it has one, so the newly-redacted occurrence gets the identical
    <Patient N> label already used everywhere else in this file for that
    same person. A name with NO existing number -- which only happens for
    a check-#5 "unresolved name position" leak, since scan_pdf() never
    recognized it as a name in the first place and so never assigned it
    one -- gets the next number in sequence instead, exactly as if it had
    been found during the original scan. This dict is MUTATED in place
    with any such new assignment, so a caller that built it from its own
    match list sees the new number too (e.g. to log it, or to build the
    'placeholder' text for a review log entry) without needing a second
    return value for it. Every occurrence of the SAME exact leftover text
    reviewed together gets the SAME new number, whichever occurrence is
    processed first.

    An 'ignore' decision touches nothing -- the real name (or unresolved
    text) stays visible at that exact location, the same "leave it, but
    the reviewer explicitly chose to" principle as excel_spreadsheet.py's
    apply_leak_decisions(). main_gui.py is responsible for logging that
    choice durably (see write_manual_review_log()), same as the
    spreadsheet side.

    Returns the list of NameMatch objects actually redacted -- empty if
    nothing was written to output_path at all (every decision was
    'ignore'). apply_redactions() mutates each of these in place with the
    original font/size/color it sampled before wiping that spot, the same
    as any other match it processes -- so a caller that also maintains a
    name key file should fold these into whatever match list it next
    passes to write_key()/write_key_batch(), the same way it would any
    other match, rather than treating them separately. Without that, a
    newly-discovered name from a check-#5 leak would end up correctly
    redacted in the PDF but ABSENT from the key file, making it
    impossible to un-redact later.
    """
    to_redact = [d for d in decisions if d.get("action") == "redact"]
    if not to_redact:
        return []

    next_no = (max(name_to_patient_no.values()) + 1) if name_to_patient_no else 1
    synthetic_matches = []
    for d in to_redact:
        patient_no = name_to_patient_no.get(d["name"])
        if patient_no is None:
            patient_no = next_no
            name_to_patient_no[d["name"]] = patient_no
            next_no += 1
        x0, y0, x1, y1 = d["rect"]
        rect = fitz.Rect(x0, y0, x1, y1)
        synthetic_matches.append(
            NameMatch(
                page_index=d["page_index"],
                rect=rect,
                # No known "next column" for an arbitrary leaked
                # occurrence (unlike a scanned match, which sits in a
                # known report layout) -- right_boundary = rect.x1 means
                # _redact_rect_for_match() adds no extension beyond the
                # name's own tight box (plus the same padding every
                # redaction rect gets), which is the safe, conservative
                # choice: it can never reach into unrelated content.
                right_boundary=rect.x1,
                raw_name=d["name"],
                patient_no=patient_no,
            )
        )
    if not synthetic_matches:
        return []

    tmp_path = output_path + ".leak_review_tmp"
    apply_redactions(output_path, tmp_path, synthetic_matches)
    os.replace(tmp_path, output_path)
    return synthetic_matches


def write_name_key(key_path: str, input_path: str, output_path: str, matches: List[NameMatch]) -> None:
    """Writes a CSV key mapping each <Patient N> placeholder back to the
    real name it replaced, so the redaction can be reversed later if needed
    (e.g. to re-identify a specific patient's entries during the
    engagement). One row per unique patient number, in placeholder order.
    CSV rather than plain text so it opens directly as a sortable/filterable
    table in Excel.

    This file is exactly as sensitive as the ORIGINAL, unredacted PDF --
    anyone holding both it and the redacted PDF can fully de-redact the
    file. It must never be sent alongside the redacted PDF to whatever
    party the redaction exists to protect against. That warning is written
    into the first few rows of the file itself (not just this docstring),
    since the CSV is what actually travels with the case file.

    Also records each placeholder's original font/size/color (sampled
    during apply_redactions(), before the text was wiped -- see
    _sample_original_style()) so the Un-redact tab can restore the name
    closer to how it originally looked, rather than always falling back to
    plain black Helvetica. This is a cosmetic detail, not a safety one --
    font/size/color carry no identifying information on their own, so
    including them adds nothing to what this file already needs to be
    kept confidential for.
    """
    with open(key_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["CONFIDENTIAL -- Patient name key"])
        writer.writerow(["Same sensitivity as the ORIGINAL, unredacted PDF -- do NOT send or store"])
        writer.writerow(["this file alongside the redacted PDF. Keep it with your own case file."])
        writer.writerow([f"Source PDF: {os.path.basename(input_path)}"])
        writer.writerow([f"Redacted PDF: {os.path.basename(output_path)}"])
        writer.writerow([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
        writer.writerow([])
        _write_name_key_body(writer, matches)


def _write_name_key_body(writer, matches: List[NameMatch]) -> None:
    by_patient = {}
    for m in matches:
        # first-seen wins, consistent with the name itself -- for a batch
        # call this means the file the name FIRST appeared in (matches are
        # passed in file-processing order) supplies the style sample.
        by_patient.setdefault(m.patient_no, m)

    writer.writerow(["Placeholder", "Real Name", "Font", "Size", "Color", "Flags"])
    for patient_no in sorted(by_patient):
        m = by_patient[patient_no]
        color_hex = f"#{m.original_color:06x}" if m.original_color is not None else ""
        writer.writerow([
            f"<{ENTITY_LABEL} {patient_no}>",
            m.raw_name,
            m.original_font or "",
            m.original_size if m.original_size is not None else "",
            color_hex,
            m.original_flags if m.original_flags is not None else "",
        ])


def write_name_key_batch(key_path: str, file_pairs: "list", matches: List[NameMatch]) -> None:
    """Same as write_name_key(), but for a BATCH of PDFs redacted together
    with shared placeholder numbering (see assign_patient_numbers()'s
    `existing_map` parameter and base.py's "BATCH REDACTION" section) --
    one combined key file covering every file in the batch, instead of one
    key file per file.

    `file_pairs` is [(source_basename, redacted_basename), ...] for every
    file in the batch, in the order they were processed; `matches` is the
    concatenation of every file's own matches list, in that same order,
    after assign_patient_numbers() has assigned .patient_no to all of them
    (so the same real name across files shares one row here, same as
    write_name_key() dedupes repeats within a single file).

    Same confidentiality as write_name_key(): this file is exactly as
    sensitive as the ORIGINAL, unredacted PDFs -- anyone holding it plus
    the redacted files can fully de-redact the whole batch.
    """
    with open(key_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["CONFIDENTIAL -- Patient name key (batch)"])
        writer.writerow(["Same sensitivity as the ORIGINAL, unredacted PDFs -- do NOT send or store"])
        writer.writerow(["this file alongside the redacted PDFs. Keep it with your own case file."])
        writer.writerow([f"Files in this batch: {len(file_pairs)}"])
        for source_name, redacted_name in file_pairs:
            writer.writerow([f"  Source: {source_name}  ->  Redacted: {redacted_name}"])
        writer.writerow([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
        writer.writerow([])
        _write_name_key_body(writer, matches)


_KEY_PLACEHOLDER_RE = re.compile(r"^<[^<>]+ (\d+)>$")


def existing_numbering_from_key(key_path: str) -> dict:
    """Reads a name key CSV saved by an EARLIER, SEPARATE redaction run
    (write_name_key()'s own output, or a combined batch key from
    write_name_key_batch() -- either shape works, both share the same
    "Placeholder, Real Name, ..." table) and returns an existing_map
    usable by assign_patient_numbers()'s existing_map parameter --
    {normalize_name(real_name): patient_no}. This is what lets the
    Redact tab's "Use existing key file(s) for consistent numbering"
    picker carry a patient's number forward from an earlier, separately
    redacted document into this one, without needing both documents in
    the same Batch Redact run.

    Only the placeholder's trailing NUMBER is used -- its label text is
    ignored, since apply_redactions() always rebuilds the placeholder
    text fresh from THIS module's own ENTITY_LABEL regardless of what
    label an earlier key file happened to use for that number. That's
    what makes it safe to carry a number forward even from a key file
    whose placeholders read e.g. "<Name N>" instead of "<Patient N>" (a
    linked cross-type identity from a Batch Redact run, say).

    Raises ValueError if this doesn't look like a name key CSV at all
    (e.g. a spreadsheet key or some unrelated file was picked) -- the
    same "Placeholder, Real Name" header check unredact_engine.py's own
    read_key_csv() uses, since this is the identical file format. This
    function is intentionally self-contained (a small, local CSV parse)
    rather than importing unredact_engine.py, matching this module's
    existing preference for not depending on the un-redaction side.
    """
    with open(key_path, "r", newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))

    header_seen = False
    existing_map = {}
    for row in rows:
        if not row:
            continue
        if not header_seen:
            if len(row) >= 2 and row[0].strip() == "Placeholder" and row[1].strip() == "Real Name":
                header_seen = True
            continue
        if len(row) < 2:
            continue
        m = _KEY_PLACEHOLDER_RE.match(row[0].strip())
        if not m:
            continue
        number = int(m.group(1))
        norm = normalize_name(row[1])
        # The same normalized identity showing up more than once
        # (shouldn't normally happen -- write_name_key() already dedupes
        # by patient_no) keeps the SMALLEST number, deterministically,
        # rather than whichever row happened to be read last.
        if norm not in existing_map or number < existing_map[norm]:
            existing_map[norm] = number

    if not header_seen:
        raise ValueError(
            "This doesn't look like a name key CSV -- no 'Placeholder, Real Name' table "
            "was found in it. Make sure this is a *_NAME_KEY.csv file saved by an earlier "
            "redaction, not some other file (a spreadsheet redaction key, for instance, has "
            "a different format)."
        )
    return existing_map


def merge_existing_numbering(mappings: List[dict]) -> "tuple[dict, list]":
    """Combines several existing_numbering_from_key() results (one per key
    file the user picked in the Redact tab) into one existing_map, the
    same idea as unredact_engine.merge_key_mappings() combining key files
    for un-redaction. Returns (merged, conflicts); conflicts is
    [{'name': normalized_name, 'numbers': [n1, n2, ...]}] for any
    identity assigned a DIFFERENT number across the files given -- this
    should never happen for key files that genuinely belong together, so
    the caller (main_gui.py) treats any conflict as blocking rather than
    silently picking one number over another.
    """
    numbers_by_name = {}
    for mapping in mappings:
        for norm, number in mapping.items():
            seen = numbers_by_name.setdefault(norm, [])
            if number not in seen:
                seen.append(number)
    merged = {}
    conflicts = []
    for norm, numbers in numbers_by_name.items():
        merged[norm] = numbers[0]
        if len(numbers) > 1:
            conflicts.append({"name": norm, "numbers": numbers})
    return merged, conflicts


# --------------------------------------------------------------------
# Generic-name aliases for the plugin contract (document_types/base.py).
# apply_redactions and verify_redaction already match the contract names
# exactly, so only these need an alias.
# --------------------------------------------------------------------
scan = scan_pdf
assign_placeholder_numbers = assign_patient_numbers
write_key = write_name_key
write_key_batch = write_name_key_batch
# Optional contract function (see base.py's "BATCH REDACTION" section,
# "cross-type identity linking") -- the exact same normalization
# assign_patient_numbers() already uses internally for its own
# same-file/same-batch matching, exposed under the generic contract name
# so the Batch Redact tab can apply it to an Excel column's cell values
# too, when the user manually links that column to this document type's
# Patient identity. Using this SAME function (not a separate lookalike)
# is what guarantees "Smith, John" is recognized as the same identity
# whether it came from this plugin's own PDF parsing or from a linked
# spreadsheet cell.
normalize_identity = normalize_name
