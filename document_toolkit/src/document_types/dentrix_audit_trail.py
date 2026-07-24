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
import os
import re
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

DATE_TOKEN_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")
HEADER_RE = re.compile(
    r"^DATE:\s*\d{1,2}/\d{1,2}/\d{2,4}\s+TIME:\s*\d{1,2}:\d{2}:\d{2}\s+USER:\s*\S+\s+TYPE:",
    re.IGNORECASE,
)
NAME_TOKEN_RE = re.compile(r"^[0-9]*[A-Za-z][A-Za-z'\-.]*,?$")
# A token that's a parenthetical aside attached to a name, e.g. a preferred
# name/nickname: "(Marshall)". Optional leading digits mirrors NAME_TOKEN_RE
# in case a chart number is glued onto it too.
PAREN_TOKEN_RE = re.compile(r"^\(?[0-9]*[A-Za-z][A-Za-z'\-.]*\)?,?$")
# A hyphen/dash standing alone as its own word -- happens when a hyphenated
# double-barreled last name has spaces around the hyphen, e.g.
# "Marshall - Harris, Ethan" tokenizes as ["Marshall", "-", "Harris,", ...].
HYPHEN_TOKENS = {"-", "‐", "‑", "‒", "–", "—"}
# A bare numeric token (e.g. a chart/ID number glued on as a separate word
# rather than fused onto the name itself) is only plausible as the very
# first word of a name -- allowing it anywhere else risks swallowing an
# unrelated number from later in the line.
LEADING_DIGITS_RE = re.compile(r"^[0-9]+$")

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


def _find_name_span(line_words):
    """If this line looks like 'MM/DD/YYYY Last, First ...<column break>...',
    return (name_text, rect, right_boundary_x). Otherwise return None.
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
    name_words = []
    right_boundary = None
    while idx < len(line_words):
        tok = line_words[idx][4]
        gap_before = line_words[idx][0] - line_words[idx - 1][2]
        if name_words and gap_before > threshold:
            right_boundary = line_words[idx][0]
            break
        # A pure-digit word after we already have at least one name word
        # looks like a chart/ID number appended AFTER the name (e.g.
        # "Marshall-Harris, Ethan 839") rather than part of it -- stop the
        # name here so that number is left alone and not redacted. (A
        # leading digit -- the very first word -- is handled separately
        # below, since that's the "chart number glued onto the front"
        # case and should stay part of the name.)
        if name_words and LEADING_DIGITS_RE.match(tok):
            right_boundary = line_words[idx][0]
            break
        name_words.append(line_words[idx])
        idx += 1
    if right_boundary is None:
        # name ran to the end of the line with no further column after it
        right_boundary = line_words[-1][2] + 200  # generous room; caller clamps to page width

    if not name_words:
        return None
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


def normalize_name(name: str) -> str:
    return re.sub(r"\s+", " ", name.strip()).lower().rstrip(",")


def assign_patient_numbers(matches: List[NameMatch]) -> dict:
    """Mutates each match's .patient_no in place (first-seen order).
    Returns the name->number mapping used, for display/audit purposes.
    """
    mapping = {}
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

    Returns {'ok': bool, 'issues': list[str]}. A non-ok result means the file
    should NOT be trusted/delivered as-is.

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
    try:
        if len(doc_in) != len(doc_out):
            issues.append(f"Page count changed: {len(doc_in)} -> {len(doc_out)}")

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
        seen_names = sorted(set(m.raw_name for m in matches))
        leaked_names = [name for name in seen_names if name in text_out_all]
        for name in leaked_names:
            locations = []
            for page_idx, page_text in enumerate(out_page_texts):
                start_search = 0
                while True:
                    found_at = page_text.find(name, start_search)
                    if found_at == -1:
                        break
                    ctx_start = max(0, found_at - 30)
                    ctx_end = min(len(page_text), found_at + len(name) + 30)
                    snippet = " ".join(page_text[ctx_start:ctx_end].split())
                    locations.append(f"page {page_idx + 1}: \"...{snippet}...\"")
                    start_search = found_at + len(name)
            if locations:
                shown = locations[:3]
                more = f" (+{len(locations) - 3} more occurrence(s))" if len(locations) > 3 else ""
                issues.append(
                    f"Name still present in output text: {name!r} -- " + "; ".join(shown) + more
                )

        for marker in STRUCTURAL_MARKERS:
            in_count = text_in_all.count(marker)
            out_count = text_out_all.count(marker)
            if in_count != out_count:
                issues.append(
                    f"{marker!r} count changed ({in_count} -> {out_count}); "
                    f"unrelated report content may have been lost"
                )

        expected_labels = sorted(set(f"<{ENTITY_LABEL} {m.patient_no}>" for m in matches))
        for label in expected_labels:
            if label not in text_out_all:
                issues.append(f"Expected placeholder missing from output: {label}")
    finally:
        doc_in.close()
        doc_out.close()
    return {"ok": len(issues) == 0, "issues": issues}


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
    """
    by_patient = {}
    for m in matches:
        by_patient.setdefault(m.patient_no, m.raw_name)  # first-seen wins, consistent with numbering

    with open(key_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["CONFIDENTIAL -- Patient name key"])
        writer.writerow(["Same sensitivity as the ORIGINAL, unredacted PDF -- do NOT send or store"])
        writer.writerow(["this file alongside the redacted PDF. Keep it with your own case file."])
        writer.writerow([f"Source PDF: {os.path.basename(input_path)}"])
        writer.writerow([f"Redacted PDF: {os.path.basename(output_path)}"])
        writer.writerow([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
        writer.writerow([])
        writer.writerow(["Placeholder", "Real Name"])
        for patient_no in sorted(by_patient):
            writer.writerow([f"<{ENTITY_LABEL} {patient_no}>", by_patient[patient_no]])


# --------------------------------------------------------------------
# Generic-name aliases for the plugin contract (document_types/base.py).
# apply_redactions and verify_redaction already match the contract names
# exactly, so only these three need an alias.
# --------------------------------------------------------------------
scan = scan_pdf
assign_placeholder_numbers = assign_patient_numbers
write_key = write_name_key
