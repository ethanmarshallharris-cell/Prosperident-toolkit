"""
Prosperident Document Redactor -- desktop GUI.

A central app for redacting identifying information out of different kinds
of report PDFs, and for reversing that redaction later given the name key
file saved alongside it. Each supported report layout is a separate
"document type" plugin (see document_types/base.py for the plugin
contract, document_types/dentrix_audit_trail.py for the first one) --
choosing a document type from the dropdown is what tells the app which
parsing/redaction logic to use; adding a new report layout later never
requires changing this file, only adding a new plugin module.

Two tabs:
  Redact      Pick a document type and source file, then either scan it
              (document types that auto-detect matches in free text, e.g.
              PDF reports) or load its columns (document types that
              redact by column selection, e.g. Excel workbooks -- see
              document_types/base.py's "TWO SHAPES OF PLUGIN" section).
              Review what was found/chosen, then redact & save. The tool
              re-verifies its own output before reporting success -- if
              anything looks off, it refuses to hand back a file that
              might be wrong.
  Un-redact   Pick a previously redacted file (PDF or Excel workbook) and
              the *_KEY.csv saved alongside it, and get back a copy with
              every placeholder swapped for the real value it replaced.
              Which engine handles it (unredact_engine.py for PDFs,
              spreadsheet_unredact_engine.py for workbooks) is chosen
              automatically from the redacted file's extension -- the
              same tab and buttons work for either.

Built with only the Python standard library (tkinter) plus PyMuPDF and
openpyxl, so it packages cleanly with PyInstaller.
"""
import csv
import os
import re
import sys
import tempfile
import threading
import time
import traceback
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import document_types
import unredact_engine
import spreadsheet_unredact_engine

# update_check.py and branding.py ship alongside this file in the GitHub
# repository (document_toolkit/src/). Imported defensively so a local
# working copy that lacks them still runs -- just without the update
# banner or the Prosperident header -- instead of failing to start.
try:
    import update_check
except ImportError:
    update_check = None
try:
    import branding
except ImportError:
    branding = None

import fitz  # PyMuPDF

APP_TITLE = "Prosperident Document Redactor"
# Bumped whenever the build changes in a way that's worth being able to
# confirm at a glance -- e.g. "is this machine actually running the
# version with the per-page progress display, or an older build?" Shown
# in the window title bar so that question never requires guessing.
APP_BUILD = "2026-10-01.1"
PREVIEW_PAD_X = 40   # extra horizontal context (points) around a previewed line
PREVIEW_PAD_Y = 3    # extra vertical context (points)
PREVIEW_ZOOM = 3.0


def entity_plural(doc_type):
    return getattr(doc_type, "ENTITY_LABEL_PLURAL", doc_type.ENTITY_LABEL + "s")


def check_file_not_locked(path):
    """Returns None if `path` doesn't exist yet or can be opened for
    writing right now; otherwise returns the OSError that opening it
    raised. Used to catch a locked destination (open in another program,
    mid-sync with a cloud client, held by antivirus) immediately, rather
    than after however many minutes of work it takes to reach the final
    write step.
    """
    if not os.path.exists(path):
        return None
    try:
        with open(path, "ab"):
            pass
        return None
    except OSError as exc:
        return exc


# Batch Redact's cross-type identity linking (see document_types/base.py's
# "BATCH REDACTION: CROSS-TYPE IDENTITY LINKING" section). A batch's
# running `state` dict is normally {column_name: {value: placeholder}} or
# {"_map"-like key: {...}} -- shared identity pools are stashed in that
# SAME dict under a tuple key built from this prefix + an ENTITY_LABEL, so
# they can never collide with an actual (string) column name.
_BATCH_IDENTITY_PREFIX = "__identity__"
# Shown in the "Same identity as:" column dropdown, and used as the
# sentinel meaning "don't link this column to anything" -- deliberately
# not a blank string, so it reads clearly in the dropdown itself.
_BATCH_NO_IDENTITY_LINK = "(none -- independent numbering)"


def _default_normalize_identity(text):
    """Fallback identity-comparison key when a linked pdf_report document
    type doesn't define its own normalize_identity() -- trim, collapse
    internal whitespace, lowercase. See base.py's cross-type identity
    linking section: this deliberately does NOT attempt to reconcile
    different name orderings (e.g. "Smith, John" vs "John Smith") -- that
    kind of guess risks merging two different real people under one
    placeholder, which is a worse mistake than two placeholders for one
    person.
    """
    return re.sub(r"\s+", " ", str(text).strip()).lower()


class _BatchFileFailure(Exception):
    """Raised internally by DocumentRedactorApp._batch_redact_worker when
    one file in a Batch Redact run fails its self-check. Carries enough
    detail for the all-or-nothing cleanup and the error message shown to
    the user; never escapes the worker thread."""

    def __init__(self, source_path, output_path, issues):
        super().__init__(f"self-check failed for {source_path}")
        self.source_path = source_path
        self.output_path = output_path
        self.issues = issues


class RunLogger:
    """Writes an incremental, append-only log for one run of a long
    operation (redact-and-save, or un-redact-and-save) -- one line the
    instant each phase actually happens, not assembled and written only
    at the end. That distinction matters: if a single call genuinely never
    returns, nothing after it would ever execute, so a "write it all out
    at the end" log would stay completely empty for exactly the case
    where a log is needed most (distinguishing a genuine hang from
    just-very-slow). log_name should be a short, stable, human-meaningful
    filename fragment, e.g. "redact" or "unredact" -- the actual log lives
    in the OS temp folder so it's always somewhere writable.
    """

    def __init__(self, log_name, header_line):
        self.log_path = os.path.join(tempfile.gettempdir(), f"DocumentRedactor_{log_name}_log.txt")
        self.run_start = time.time()
        self._append(f"\n--- {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} -- {header_line} ---")

    def _append(self, text):
        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(text + "\n")
        except OSError:
            pass

    def line(self, text):
        elapsed = time.time() - self.run_start
        self._append(f"[{datetime.now().strftime('%H:%M:%S')} +{elapsed:6.1f}s] {text}")


def _candidate_log_paths(near_path, log_name, suffix):
    """Where to try saving a diagnostic log file, in order of preference.

    First choice: right next to `near_path` -- the destination file the
    user chose (for a redact/un-redact failure), or the source file they
    were working with (for a scan/load failure, which happens before any
    destination is chosen). That's a location the user already has open
    in Explorer and knows how to find, rather than the OS temp folder,
    which most people have never navigated to and wouldn't think to look
    in without being told the exact path every time.

    Second choice (fallback): the OS temp folder, same as before this
    existed. This can't be the ONLY place tried, because the destination
    folder not being writable is itself a plausible reason a run just
    failed -- if that's what happened, trying to write the log there too
    would just fail a second time, silently, leaving no log at all.
    """
    candidates = []
    if near_path:
        base = os.path.splitext(os.path.basename(near_path))[0]
        near_dir = os.path.dirname(near_path) or "."
        candidates.append(os.path.join(near_dir, f"{base}_{suffix}.txt"))
    candidates.append(os.path.join(tempfile.gettempdir(), f"DocumentRedactor_{log_name}_{suffix}.txt"))
    return candidates


def _write_log_file(near_path, log_name, suffix, body_writer):
    """Tries each candidate location from _candidate_log_paths() in turn,
    writing via body_writer(file_handle) -- returns the path that
    actually worked, or None if every candidate failed."""
    for path in _candidate_log_paths(near_path, log_name, suffix):
        try:
            with open(path, "w", encoding="utf-8") as f:
                body_writer(f)
            return path
        except OSError:
            continue
    return None


def write_error_log(log_name, input_path, output_path, tb, near_path=None):
    """near_path: where to prefer saving this log -- normally the
    destination the user chose (output_path), passed explicitly here only
    for the scan/load case, where there is no destination yet and the
    source file (input_path) is used instead. Defaults to output_path.
    """
    def _body(f):
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Input file: {input_path}\n")
        f.write(f"Intended output: {output_path}\n\n")
        f.write(tb)
    return _write_log_file(near_path if near_path is not None else output_path, log_name, "error_log", _body)


def write_selfcheck_log(log_name, input_path, output_path, issues, near_path=None, extra_detail_lines=None):
    """Like write_error_log(), but for a *self-check* failure specifically
    (verify_redaction/verify_column_redaction/verify_unredaction returning
    ok=False) rather than an exception -- there's no traceback for this
    case, just the issues list, but it deserves the exact same durable,
    findable log file. Before this existed, a self-check failure only ever
    showed its detail in one on-screen message box: fine as far as it
    went, but nothing was left afterwards to point at, forward, or refer
    back to -- if that box got closed, scrolled past, or was just hard to
    read/copy from on the day, whatever it said was gone. Every
    self-check failure now leaves this file behind regardless, the same
    way every exception already does. See write_error_log() for near_path.

    extra_detail_lines: optional list[str] of additional, UNTRUNCATED lines
    appended after the (possibly "+N more"-truncated) issues section --
    used by the Un-redact tab to record every single leftover placeholder
    (not just the first 10 shown in the on-screen dialog) so a failure
    affecting many cells can still be fully diagnosed from this one file.
    Every caller of this is expected to only ever pass placeholder TEXT and
    cell locations here, never a real value -- this log is not treated as
    sensitive, unlike write_manual_review_log()'s.
    """
    def _body(f):
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Input file: {input_path}\n")
        f.write(f"Intended output: {output_path}\n\n")
        f.write("Self-check found the following problem(s):\n\n")
        for issue in issues:
            f.write(f"  - {issue}\n")
        if extra_detail_lines:
            f.write(f"\nFull, untruncated detail ({len(extra_detail_lines)} item(s)):\n\n")
            for line in extra_detail_lines:
                f.write(f"  - {line}\n")
    return _write_log_file(near_path if near_path is not None else output_path, log_name, "selfcheck_log", _body)


def write_manual_review_log(log_name, input_path, output_path, replaced, ignored, near_path=None):
    """Records the outcome of a manual-review pass over a self-check
    finding -- used by BOTH directions: redaction's per-cell leak review
    (excel_spreadsheet.py's verify_column_redaction()'s 'leaks' field /
    apply_leak_decisions()), where a 'replaced' cell was overwritten with
    its placeholder; and un-redaction's ambiguous-placeholder review
    (spreadsheet_unredact_engine.py's 'ambiguous' field /
    apply_ambiguous_decisions()), where a 'replaced' cell was overwritten
    with the real value the user picked. Either way, the actual real
    value is included for every 'ignored' entry, since that is exactly
    what's now sitting in the delivered file, and a forensic case file
    needs a durable record of what was knowingly left un-redacted and
    why, not just a bare cell reference.

    This log file carries the SAME sensitivity as the original,
    unredacted document whenever it has any 'ignored' entries in it (it
    contains real values) -- treat it exactly like the *_REDACTION_KEY.csv
    file: keep it with your own case file, never alongside the redacted
    output. See write_error_log() for the near_path/fallback behavior
    this shares.

    replaced / ignored: lists of {'sheet', 'coordinate', 'action', ...}
    dicts, as built by main_gui.py's manual-review screens. A 'replaced'
    entry's resulting cell content is read from whichever of 'placeholder'
    or 'value' is present (redaction's leak review sets 'placeholder';
    un-redaction's ambiguous review sets 'value') -- this function doesn't
    need to care which flow it came from beyond that.
    """
    def _resulting_text(d):
        return d.get("placeholder", d.get("value", "?"))

    def _body(f):
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Input file: {input_path}\n")
        f.write(f"Output file: {output_path}\n\n")
        if ignored:
            f.write(
                "CONFIDENTIAL -- this section lists real, un-redacted values left in the "
                "output file by deliberate choice. Same sensitivity as the original "
                "document -- keep with your case file, do not send alongside the output.\n\n"
            )
        f.write("Manual review of self-check findings:\n\n")
        if replaced:
            f.write(f"Replaced ({len(replaced)}):\n")
            for d in replaced:
                f.write(f"  - {d['sheet']}!{d['coordinate']} -> {_resulting_text(d)}\n")
            f.write("\n")
        if ignored:
            f.write(f"Left as-is by deliberate choice, value still present ({len(ignored)}):\n")
            for d in ignored:
                f.write(f"  - {d['sheet']}!{d['coordinate']}: {d['value']}\n")
            f.write("\n")
        if not replaced and not ignored:
            f.write("(no decisions recorded)\n")
    return _write_log_file(near_path if near_path is not None else output_path, log_name, "manual_review_log", _body)


def write_pdf_manual_review_log(log_name, input_path, output_path, replaced, ignored, near_path=None):
    """Same purpose as write_manual_review_log(), but for the PDF leak
    review screen (_apply_pdf_leak_review_decisions / the Batch Redact
    tab's equivalent) -- a PDF occurrence doesn't have a spreadsheet cell
    reference, so this records a page number and the real name instead.
    Same confidentiality note applies whenever 'ignored' is non-empty.

    replaced: [{'name': str, 'page_index': int, 'placeholder': str}, ...]
    ignored: [{'name': str, 'page_index': int}, ...]
    """
    def _body(f):
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Input file: {input_path}\n")
        f.write(f"Output file: {output_path}\n\n")
        if ignored:
            f.write(
                "CONFIDENTIAL -- this section lists real, un-redacted names left in the "
                "output file by deliberate choice. Same sensitivity as the original "
                "document -- keep with your case file, do not send alongside the output.\n\n"
            )
        f.write("Manual review of self-check findings:\n\n")
        if replaced:
            f.write(f"Redacted ({len(replaced)}):\n")
            for d in replaced:
                f.write(f"  - page {d['page_index'] + 1}: {d['name']!r} -> {d['placeholder']}\n")
            f.write("\n")
        if ignored:
            f.write(f"Left as-is by deliberate choice, name still present ({len(ignored)}):\n")
            for d in ignored:
                f.write(f"  - page {d['page_index'] + 1}: {d['name']!r}\n")
            f.write("\n")
        if not replaced and not ignored:
            f.write("(no decisions recorded)\n")
    return _write_log_file(near_path if near_path is not None else output_path, log_name, "manual_review_log", _body)


def write_batch_needs_review_log(output_dir, file_unmatched_pairs):
    """Writes one combined CSV listing every 'needs review' line found
    across an entire Batch Redact run (see DocumentRedactorApp._batch_redact_worker)
    -- the batch-mode counterpart to the single-file Redact tab's "Needs
    review" panel, which isn't shown at all in batch mode (batch mode has
    no per-match curation UI -- see base.py's "BATCH REDACTION" section).

    `file_unmatched_pairs` is [(source_basename, [UnmatchedLine, ...]), ...]
    for every file that had at least one unmatched line; only called when
    that list is non-empty. Not confidential in the way a key file is --
    every line in it is text that's still sitting, un-redacted, in that
    file's own output (this is just an index of where to go look), so it
    carries no MORE sensitivity than the redacted output files themselves.
    """
    path = os.path.join(output_dir, "BATCH_NEEDS_REVIEW.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Source file", "Page", "Context", "Text (not redacted -- check manually)"])
        for source_name, unmatched in file_unmatched_pairs:
            for u in unmatched:
                writer.writerow([source_name, u.page_index + 1, (u.context or "").strip(), u.raw_text])
    return path


LOCKED_FILE_MESSAGE = (
    "This file appears to be open or locked by something else right now:\n\n"
    "{path}\n\n{exc}\n\n"
    "Close whatever has it open (a PDF viewer, File Explorer preview pane, "
    "antivirus scan, or a cloud-sync client like OneDrive/Dropbox if this "
    "file is in a synced folder), then try again -- this can take a while "
    "on a large file, and would only fail again at the very end if the "
    "lock is still there.\n\nContinue anyway?"
)

FAILURE_TROUBLESHOOTING = (
    "This kind of error is usually the operating system refusing the write, "
    "not a problem with the file itself -- worth trying:\n"
    "  - Save to a plain local folder (e.g. Documents) instead of Desktop "
    "or a OneDrive/Dropbox/Google Drive folder, at least as a test\n"
    "  - Close any program that might have that file or folder open (a "
    "PDF viewer, File Explorer preview pane, antivirus scan)\n"
    "  - Check that the drive isn't full"
)


class DocumentRedactorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  (build {APP_BUILD})")
        self.geometry("1080x720")
        self.minsize(880, 600)

        if branding is not None:
            branding.apply_theme(self)
            branding.add_header(
                self, "Document Redactor",
                "Redact & un-redact identifying information in report PDFs and Excel workbooks",
            )

        self.doc_types = document_types.DOCUMENT_TYPES
        self.current_doc_type = self.doc_types[0]

        self.input_path = None
        self.matches = []
        self.unmatched = []
        self.page_count = 0
        self.mapping = {}
        self.excluded_ids = set()
        self._preview_photo = None
        self.spreadsheet_columns = []
        self.column_vars = {}

        # Redact tab: optional key file(s) from an EARLIER, separate
        # redaction, loaded so a name/value that recurs in THIS document
        # gets the same placeholder it already got there -- an
        # alternative to Batch Redact's cross-file numbering for someone
        # redacting documents one at a time rather than all at once. See
        # on_redact_add_existing_keys() / _resolve_redact_existing_numbering().
        # redact_existing_key_maps mirrors redact_existing_key_paths 1:1,
        # caching each file's already-parsed/validated mapping so it
        # isn't re-read from disk on every Scan.
        self.redact_existing_key_paths = []
        self.redact_existing_key_maps = {}
        self._redact_existing_map = None

        # Redact tab, spreadsheet types only: a key file loaded from a
        # DIFFERENT (PDF-shaped) document type, so a spreadsheet column
        # can be manually linked to that type's identity pool -- the
        # single-file analog of Batch Redact's cross-type identity
        # linking (see base.py). redact_cross_type_key_maps[path] =
        # {"doc_type": <module>, "entity_label": str, "mapping": dict}.
        # Kept separate from redact_existing_key_maps above since it's a
        # different shape (per-file entity_label + doc_type, not just a
        # numbering dict) and only ever applies to a spreadsheet's linked
        # columns, never to a plain "same document type" carry-forward.
        self.redact_cross_type_key_paths = []
        self.redact_cross_type_key_maps = {}

        self.unredact_pdf_path = None
        self.unredact_key_paths = []

        # Batch Redact tab state -- deliberately separate from the
        # single-file Redact tab's self.input_path/self.matches/etc. above
        # so switching between the two tabs never clobbers either one's
        # in-progress work. A batch can mix document types in one file
        # list (see base.py's "BATCH REDACTION" section) -- batch_files
        # holds one {"path", "doc_type"} entry per added file, each
        # resolved to its plugin automatically by extension.
        self.batch_doc_types = [dt for dt in self.doc_types if getattr(dt, "SUPPORTS_BATCH", False)]
        self.batch_files = []
        self.batch_column_vars = {}
        self.batch_output_dir = None
        # entity_label -> normalize_identity function, recomputed each
        # time columns are (re)loaded from the current file list -- see
        # _batch_entity_label_options().
        self.batch_identity_normalize_fns = {}

        self._build_widgets()
        self._start_update_check()

    # ================================================================
    # Shared window title / progress helpers
    # ================================================================
    def _set_title(self, suffix=""):
        base = f"{APP_TITLE}  (build {APP_BUILD})"
        self.title(f"{base}  --  {suffix}" if suffix else base)

    # ================================================================
    # Update check (background, silent unless there's something to show)
    # ================================================================
    def _start_update_check(self):
        if update_check is None:
            return
        threading.Thread(target=self._update_check_worker, daemon=True).start()

    def _update_check_worker(self):
        try:
            result = update_check.check_for_update("document_toolkit", APP_BUILD)
        except Exception:
            return  # never let an update-check problem affect the tool itself
        if result:
            self.after(0, lambda: self._show_update_banner(result))

    def _show_update_banner(self, entry):
        notes = entry.get("notes") or ""
        text = f"A newer version is available (build {entry.get('version')}). Click here to get it."
        if notes:
            text += f"  \u2014 {notes}"
        link = ttk.Label(
            self, text=text, foreground="#0645AD",
            cursor="hand2", font=("", 9, "underline"),
        )
        # Packed before the main notebook so the expanding notebook can
        # never squeeze the banner out of view.
        nb = getattr(self, "_main_notebook", None)
        if nb is not None:
            link.pack(side="bottom", fill="x", padx=10, pady=4, before=nb)
        else:
            link.pack(side="bottom", fill="x", padx=10, pady=4)
        download_url = entry.get("download_url")
        if download_url:
            link.bind("<Button-1>", lambda e: webbrowser.open(download_url))

    @staticmethod
    def _build_scrollable_frame(parent):
        """A vertically-scrollable area (canvas + inner frame + scrollbar)
        for content whose length isn't known ahead of time -- used for the
        spreadsheet document types' column checklist, which can run to
        many rows on a wide workbook. Returns (outer_frame_to_pack, inner_frame_to_put_content_in).
        """
        outer = ttk.Frame(parent)
        canvas = tk.Canvas(outer, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)

        def _on_canvas_resize(event):
            canvas.itemconfigure(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_resize)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        return outer, inner

    # ================================================================
    # Top-level layout: two tabs sharing one window
    # ================================================================
    def _build_widgets(self):
        outer = ttk.Notebook(self)
        outer.pack(fill="both", expand=True)
        self._main_notebook = outer

        redact_tab = ttk.Frame(outer)
        outer.add(redact_tab, text="Redact")
        self._build_redact_tab(redact_tab)

        batch_tab = ttk.Frame(outer)
        outer.add(batch_tab, text="Batch Redact")
        self._build_batch_tab(batch_tab)

        unredact_tab = ttk.Frame(outer)
        outer.add(unredact_tab, text="Un-redact")
        self._build_unredact_tab(unredact_tab)

    # ================================================================
    # REDACT TAB
    # ================================================================
    def _build_redact_tab(self, root):
        pad = {"padx": 10, "pady": 6}

        top = ttk.Frame(root)
        top.pack(fill="x", **pad)

        ttk.Label(top, text="Document type:").pack(side="left")
        self.doc_type_var = tk.StringVar(value=self.current_doc_type.DISPLAY_NAME)
        doc_type_combo = ttk.Combobox(
            top, textvariable=self.doc_type_var, state="readonly", width=28,
            values=[dt.DISPLAY_NAME for dt in self.doc_types],
        )
        doc_type_combo.pack(side="left", padx=(4, 16))
        doc_type_combo.bind("<<ComboboxSelected>>", self.on_doc_type_changed)

        self.file_label_var = tk.StringVar(value="(no file selected)")
        ttk.Label(top, textvariable=self.file_label_var, foreground="#333").pack(
            side="left", padx=8, fill="x", expand=True
        )
        ttk.Button(top, text="Browse...", command=self.on_browse).pack(side="left")
        self.scan_btn = ttk.Button(top, text="Scan", command=self.on_scan, state="disabled")
        self.scan_btn.pack(side="left", padx=(8, 0))

        # Optional: carry numbering forward from an earlier, SEPARATE
        # redaction's key file, so redacting a series of documents one at
        # a time (rather than all together in Batch Redact) still gives
        # the same real name/value the same placeholder across all of
        # them. Packed/unpacked by on_doc_type_changed() -- only shown
        # for a document type whose plugin actually supports it (see
        # base.py's optional existing_numbering_from_key()/
        # merge_existing_numbering() hooks).
        self.redact_key_row = ttk.Frame(root)
        self.redact_key_row.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Label(self.redact_key_row, text="Consistent numbering:").pack(side="left")
        self.redact_existing_key_var = tk.StringVar(value="(none -- numbering starts fresh)")
        ttk.Label(
            self.redact_key_row, textvariable=self.redact_existing_key_var, foreground="#333"
        ).pack(side="left", padx=8, fill="x", expand=True)
        ttk.Button(
            self.redact_key_row, text="Use existing key file(s)...",
            command=self.on_redact_add_existing_keys,
        ).pack(side="left")
        self.redact_existing_key_clear_btn = ttk.Button(
            self.redact_key_row, text="Clear", command=self.on_redact_clear_existing_keys, state="disabled"
        )
        self.redact_existing_key_clear_btn.pack(side="left", padx=(6, 0))
        self._redact_key_hint_text_same_type = (
            "Optional: pick the key CSV file(s) saved from an earlier, SEPARATE "
            "redaction (e.g. yesterday's document) so a name or value that recurs here "
            "gets the exact same placeholder it got there -- redact document 1, then "
            "load document 1's key file before scanning document 2, and so on. Pick "
            "this before clicking Scan/Load Columns; if you change it after, click "
            "Scan/Load Columns again so the change is actually used."
        )
        self._redact_key_hint_text_spreadsheet = (
            self._redact_key_hint_text_same_type +
            " You can also pick the key CSV from a DIFFERENT document type here (e.g. a "
            "Dentrix Audit Trail Report PDF's key) -- once loaded, click Load Columns and "
            "a \"Same identity as:\" option appears next to each column so you can link it "
            "to that type's identity pool (e.g. \"Patient\"), giving matching names the "
            "same placeholder number across both files."
        )
        self.redact_key_hint = ttk.Label(
            root,
            text=self._redact_key_hint_text_same_type,
            foreground="#666", wraplength=900, justify="left",
        )
        self.redact_key_hint.pack(anchor="w", padx=10, pady=(0, 6))

        self.summary_var = tk.StringVar(value="Select a document type and file, then click Scan.")
        self.summary_label = ttk.Label(root, textvariable=self.summary_var, font=("", 10, "bold"))
        self.summary_label.pack(anchor="w", padx=10)

        self.progress = ttk.Progressbar(root, mode="indeterminate")

        # Holds whichever of the two result panels below is relevant to
        # the current document type's FORMAT_KIND -- only one is packed
        # (visible) at a time; on_doc_type_changed() swaps between them.
        self.result_container = ttk.Frame(root)
        self.result_container.pack(fill="both", expand=True, padx=10, pady=6)

        main = ttk.PanedWindow(self.result_container, orient="horizontal")
        self.pdf_panel = main
        main.pack(fill="both", expand=True)

        # --- left: results notebook (matches / needs-review) ---
        left = ttk.Frame(main)
        main.add(left, weight=3)

        notebook = ttk.Notebook(left)
        notebook.pack(fill="both", expand=True)

        matches_tab = ttk.Frame(notebook)
        notebook.add(matches_tab, text="Detected items (0)")
        self.matches_tab = matches_tab
        self.notebook = notebook

        columns = ("page", "placeholder", "text", "context", "status")
        self.tree = ttk.Treeview(matches_tab, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("page", text="Page")
        self.tree.heading("placeholder", text="Placeholder")
        self.tree.heading("text", text="Detected Text")
        self.tree.heading("context", text="Context")
        self.tree.heading("status", text="Status")
        self.tree.column("page", width=50, anchor="center")
        self.tree.column("placeholder", width=90, anchor="center")
        self.tree.column("text", width=160)
        self.tree.column("context", width=330)
        self.tree.column("status", width=110, anchor="center")
        self.tree.pack(fill="both", expand=True, side="left")
        self.tree.tag_configure("excluded", foreground="#999999")
        vsb = ttk.Scrollbar(matches_tab, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_row_selected)
        self.tree.bind("<Double-1>", self.on_row_double_click)

        hint = ttk.Label(
            left,
            text="Double-click a row to include/exclude it from redaction "
                 "(use this only for a false positive).",
            foreground="#666",
        )
        hint.pack(anchor="w", pady=(4, 0))

        review_tab = ttk.Frame(notebook)
        notebook.add(review_tab, text="Needs review (0)")
        self.review_tab = review_tab

        rcolumns = ("page", "context", "text")
        self.review_tree = ttk.Treeview(review_tab, columns=rcolumns, show="headings")
        self.review_tree.heading("page", text="Page")
        self.review_tree.heading("context", text="Context")
        self.review_tree.heading("text", text="Line text that wasn't recognized")
        self.review_tree.column("page", width=50, anchor="center")
        self.review_tree.column("context", width=300)
        self.review_tree.column("text", width=400)
        self.review_tree.pack(fill="both", expand=True, side="left")
        rvsb = ttk.Scrollbar(review_tab, orient="vertical", command=self.review_tree.yview)
        self.review_tree.configure(yscrollcommand=rvsb.set)
        rvsb.pack(side="right", fill="y")

        self.review_hint_var = tk.StringVar(value=self.current_doc_type.REVIEW_HINT)
        review_hint = ttk.Label(
            review_tab, textvariable=self.review_hint_var, foreground="#a15c00",
            wraplength=600, justify="left",
        )
        review_hint.pack(anchor="w", pady=(4, 0), fill="x")

        # --- right: preview pane ---
        right = ttk.Frame(main)
        main.add(right, weight=2)
        ttk.Label(right, text="Preview (exact text from the file for the selected row)",
                  font=("", 9, "bold")).pack(anchor="w")
        self.preview_canvas = tk.Canvas(right, background="white", highlightthickness=1,
                                         highlightbackground="#ccc")
        self.preview_canvas.pack(fill="both", expand=True, pady=(4, 0))

        def _set_initial_sash():
            main.update_idletasks()
            total = main.winfo_width()
            if total > 200:
                main.sashpos(0, max(total - 320, 480))
        self.after(50, _set_initial_sash)

        # --- spreadsheet-style result panel (column checklist) ---
        # Not packed here -- on_doc_type_changed() packs whichever of
        # pdf_panel / spreadsheet_panel matches the current document
        # type's FORMAT_KIND, and forgets the other one.
        self.spreadsheet_panel = ttk.Frame(self.result_container)
        ttk.Label(
            self.spreadsheet_panel,
            text="Check the column(s) to redact. The same value in a column always gets "
                 "the same placeholder wherever it recurs in that column (numbering is "
                 "independent per column). Edit the label if you don't want the column's "
                 "own name -- the placeholder written will be <label 1>, <label 2>, etc.",
            wraplength=900, justify="left", foreground="#444",
        ).pack(anchor="w", pady=(0, 8))
        col_scroll_outer, self.column_list_inner = self._build_scrollable_frame(self.spreadsheet_panel)
        col_scroll_outer.pack(fill="both", expand=True)
        self.column_vars = {}

        bottom = ttk.Frame(root)
        bottom.pack(fill="x", padx=10, pady=10)
        self.redact_btn = ttk.Button(
            bottom, text="Redact && Save...", command=self.on_redact_save, state="disabled"
        )
        self.redact_btn.pack(side="right")
        self.save_key_var = tk.BooleanVar(value=True)
        self.save_key_check = ttk.Checkbutton(
            bottom,
            text="Also save a name key CSV (needed to un-redact later -- keep it separate "
                 "from the redacted file)",
            variable=self.save_key_var,
        )
        self.save_key_check.pack(side="right", padx=(0, 16))

    def on_doc_type_changed(self, _event=None):
        selected_name = self.doc_type_var.get()
        for dt in self.doc_types:
            if dt.DISPLAY_NAME == selected_name:
                self.current_doc_type = dt
                break
        # Switching document types mid-workflow starts over -- a scan
        # result from one document type's parser has no meaning under a
        # different one's.
        self.input_path = None
        self.file_label_var.set("(no file selected)")
        self.scan_btn.configure(state="disabled")
        self.redact_btn.configure(state="disabled")
        self._clear_results()

        # A loaded existing-key-file mapping was validated (parsed and
        # checked for the "Placeholder, Real Name" / "Column, Placeholder,
        # Real Value" table) against the PREVIOUS document type's key
        # format -- meaningless (and potentially the wrong shape entirely)
        # once the document type changes, so it's cleared the same way
        # input_path/matches are above.
        self.redact_existing_key_paths = []
        self.redact_existing_key_maps = {}
        self.redact_cross_type_key_paths = []
        self.redact_cross_type_key_maps = {}
        self._update_existing_key_display()

        is_spreadsheet = self._is_spreadsheet_type(self.current_doc_type)
        if is_spreadsheet:
            self.scan_btn.configure(text="Load Columns")
            self.pdf_panel.pack_forget()
            self.spreadsheet_panel.pack(fill="both", expand=True)
            self.summary_var.set(
                f"Select an {self.current_doc_type.DISPLAY_NAME} file, then click Load Columns."
            )
        else:
            self.scan_btn.configure(text="Scan")
            self.spreadsheet_panel.pack_forget()
            self.pdf_panel.pack(fill="both", expand=True)
            self.summary_var.set(f"Select a {self.current_doc_type.DISPLAY_NAME} file, then click Scan.")

        self.review_hint_var.set(self.current_doc_type.REVIEW_HINT)
        if self.current_doc_type.SUPPORTS_KEY_FILE:
            self.save_key_check.pack(side="right", padx=(0, 16))
        else:
            self.save_key_check.pack_forget()

        if self._doc_type_supports_existing_key(self.current_doc_type) or is_spreadsheet:
            self.redact_key_hint.configure(
                text=self._redact_key_hint_text_spreadsheet if is_spreadsheet
                else self._redact_key_hint_text_same_type
            )
            self.redact_key_row.pack(fill="x", padx=10, pady=(0, 4), before=self.summary_label)
            self.redact_key_hint.pack(anchor="w", padx=10, pady=(0, 6), before=self.summary_label)
        else:
            self.redact_key_row.pack_forget()
            self.redact_key_hint.pack_forget()

    @staticmethod
    def _is_spreadsheet_type(doc_type):
        return getattr(doc_type, "FORMAT_KIND", "pdf_report") == "spreadsheet"

    @staticmethod
    def _doc_type_supports_existing_key(doc_type):
        """Whether this plugin implements the optional pair of hooks that
        make the Redact tab's "Use existing key file(s) for consistent
        numbering" control meaningful -- see
        existing_numbering_from_key()/merge_existing_numbering() in
        dentrix_audit_trail.py or excel_spreadsheet.py for what each does.
        Gated on SUPPORTS_BATCH too, since seeding numbering from a key
        file uses the exact same existing_map/state parameter Batch
        Redact's cross-file numbering already relies on (see base.py's
        "BATCH REDACTION" section) -- a plugin whose numbering function
        doesn't accept that parameter at all can't support this either.
        """
        return (
            getattr(doc_type, "SUPPORTS_BATCH", False)
            and hasattr(doc_type, "existing_numbering_from_key")
            and hasattr(doc_type, "merge_existing_numbering")
        )

    def on_browse(self):
        path = filedialog.askopenfilename(
            title=f"Select {self.current_doc_type.DISPLAY_NAME}",
            filetypes=self.current_doc_type.FILE_TYPES,
        )
        if not path:
            return
        self.input_path = path
        self.file_label_var.set(path)
        self.scan_btn.configure(state="normal")
        self.redact_btn.configure(state="disabled")
        self._clear_results()
        self.summary_var.set("Ready to scan.")

    def _clear_results(self):
        self.matches = []
        self.unmatched = []
        self.excluded_ids = set()
        for row in self.tree.get_children():
            self.tree.delete(row)
        for row in self.review_tree.get_children():
            self.review_tree.delete(row)
        self.preview_canvas.delete("all")
        self.spreadsheet_columns = []
        self._clear_column_checklist()

    def _clear_column_checklist(self):
        for child in self.column_list_inner.winfo_children():
            child.destroy()
        self.column_vars = {}

    def _populate_column_checklist(self, columns):
        self._clear_column_checklist()
        # Only offered for a spreadsheet redaction with at least one
        # cross-type key file loaded (see on_redact_add_existing_keys())
        # -- the single-file analog of Batch Redact's "Same identity as:"
        # column link (base.py's cross-type identity linking section).
        entity_labels = sorted({v["entity_label"] for v in self.redact_cross_type_key_maps.values()})
        for col in columns:
            row = ttk.Frame(self.column_list_inner)
            row.pack(fill="x", pady=2, padx=4)

            sel_var = tk.BooleanVar(value=False)
            label_var = tk.StringVar(value=col.default_label)
            ttk.Checkbutton(row, variable=sel_var).pack(side="left")

            sheets_note = f"  (sheet(s): {', '.join(col.sheets)})" if len(col.sheets) > 1 else ""
            sample_note = f'  --  e.g. "{col.sample}"' if col.sample else "  --  (no sample value found)"
            ttk.Label(
                row, text=f"{col.name}{sheets_note}{sample_note}", width=42, anchor="w"
            ).pack(side="left", padx=(4, 10))

            ttk.Label(row, text="Placeholder label:").pack(side="left")
            label_entry = ttk.Entry(row, textvariable=label_var, width=16)
            label_entry.pack(side="left", padx=(4, 0))
            ttk.Label(row, text="→ e.g. <label 1>", foreground="#888").pack(side="left", padx=(6, 0))

            # "Same identity as:" -- a MANUAL, per-column link to the
            # identity pool loaded from a cross-type (PDF-shaped) key
            # file -- e.g. redact a Dentrix Audit Trail Report PDF first,
            # load ITS key file here, then link this workbook's "Patient
            # Name" column to "Patient" so the same person gets the exact
            # same <Patient N> placeholder in both files. Never inferred
            # from the column's own name -- see base.py for why an
            # automatic guess would be dangerous. Picking a link forces
            # this column's placeholder label to match the linked type's
            # ENTITY_LABEL (and locks it), since a matching NUMBER with a
            # different placeholder LABEL wouldn't actually read as the
            # same placeholder.
            link_var = None
            if entity_labels:
                ttk.Label(row, text="  Same identity as:").pack(side="left", padx=(10, 2))
                link_var = tk.StringVar(value=_BATCH_NO_IDENTITY_LINK)

                def on_link_change(
                    _event=None, label_var=label_var, label_entry=label_entry,
                    link_var=link_var, default_label=col.default_label,
                ):
                    chosen = link_var.get()
                    if chosen == _BATCH_NO_IDENTITY_LINK:
                        label_var.set(default_label)
                        label_entry.configure(state="normal")
                    else:
                        label_var.set(chosen)
                        label_entry.configure(state="disabled")

                link_combo = ttk.Combobox(
                    row, textvariable=link_var, state="readonly", width=22,
                    values=[_BATCH_NO_IDENTITY_LINK] + entity_labels,
                )
                link_combo.bind("<<ComboboxSelected>>", on_link_change)
                link_combo.pack(side="left")

            self.column_vars[col.name] = {"selected": sel_var, "label": label_var, "shared_identity": link_var}

    def on_redact_add_existing_keys(self):
        doc_type = self.current_doc_type
        supports_own = self._doc_type_supports_existing_key(doc_type)
        is_spreadsheet = self._is_spreadsheet_type(doc_type)
        if not supports_own and not is_spreadsheet:
            return
        filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
        paths = filedialog.askopenfilenames(
            title="Select key file(s) from an earlier, separate redaction", filetypes=filetypes
        )
        if not paths:
            return
        added_own = []
        added_cross = []
        skipped = []
        for p in paths:
            if p in self.redact_existing_key_maps or p in self.redact_cross_type_key_maps:
                continue
            own_error = None
            if supports_own:
                try:
                    mapping = doc_type.existing_numbering_from_key(p)
                    self.redact_existing_key_maps[p] = mapping
                    self.redact_existing_key_paths.append(p)
                    added_own.append(p)
                    continue
                except Exception as exc:  # noqa: BLE001
                    own_error = str(exc)
            # A spreadsheet redaction also accepts a key file from a
            # DIFFERENT, PDF-shaped document type -- that's what makes a
            # column's "Same identity as:" link possible (see
            # _populate_column_checklist() above). Cross-type linking
            # only ever runs "spreadsheet column <-> PDF identity", never
            # the reverse, so this is only attempted when the CURRENT
            # type is the spreadsheet shape.
            if is_spreadsheet:
                candidates = []
                for other in self.doc_types:
                    if self._is_spreadsheet_type(other) or not hasattr(other, "existing_numbering_from_key"):
                        continue
                    try:
                        cross_mapping = other.existing_numbering_from_key(p)
                    except Exception:  # noqa: BLE001
                        continue
                    candidates.append((other, cross_mapping))
                if len(candidates) == 1:
                    other, cross_mapping = candidates[0]
                    self.redact_cross_type_key_maps[p] = {
                        "doc_type": other, "entity_label": other.ENTITY_LABEL, "mapping": cross_mapping,
                    }
                    self.redact_cross_type_key_paths.append(p)
                    added_cross.append(p)
                    continue
                if len(candidates) > 1:
                    skipped.append((
                        p,
                        "This key file's format matches more than one document type -- "
                        "please report this to the developer rather than guessing which one "
                        "it belongs to.",
                    ))
                    continue
            skipped.append((p, own_error or "This doesn't look like a key CSV this app recognizes."))
        if skipped:
            lines = "\n\n".join(f"{os.path.basename(p)}:\n{err}" for p, err in skipped)
            messagebox.showerror(
                APP_TITLE, f"Could not use {len(skipped)} of the selected file(s):\n\n{lines}"
            )
        if added_own or added_cross:
            self._update_existing_key_display()
            self._warn_existing_key_rescan_needed(cross_type_changed=bool(added_cross))

    def on_redact_clear_existing_keys(self):
        if not self.redact_existing_key_paths and not self.redact_cross_type_key_paths:
            return
        had_cross = bool(self.redact_cross_type_key_paths)
        self.redact_existing_key_paths = []
        self.redact_existing_key_maps = {}
        self.redact_cross_type_key_paths = []
        self.redact_cross_type_key_maps = {}
        self._update_existing_key_display()
        self._warn_existing_key_rescan_needed(cross_type_changed=had_cross)

    def _update_existing_key_display(self):
        n = len(self.redact_existing_key_paths) + len(self.redact_cross_type_key_paths)
        if n == 0:
            self.redact_existing_key_var.set("(none -- numbering starts fresh)")
            self.redact_existing_key_clear_btn.configure(state="disabled")
        else:
            all_paths = self.redact_existing_key_paths + self.redact_cross_type_key_paths
            names = ", ".join(os.path.basename(p) for p in all_paths)
            cross_note = ""
            if self.redact_cross_type_key_maps:
                labels = ", ".join(sorted({v["entity_label"] for v in self.redact_cross_type_key_maps.values()}))
                cross_note = f"  (identity link available: {labels})"
            self.redact_existing_key_var.set(f"{n} file(s) loaded: {names}{cross_note}")
            self.redact_existing_key_clear_btn.configure(state="normal")

    def _warn_existing_key_rescan_needed(self, cross_type_changed=False):
        # Only the PDF side bakes numbering into a scan result up front
        # (assign_placeholder_numbers() at Scan time, shown right away in
        # the Detected items tree) -- the spreadsheet side doesn't
        # resolve numbering until Redact & Save itself (see
        # _on_redact_save_spreadsheet()), so there's nothing stale to
        # warn about there EXCEPT the column checklist's "Same identity
        # as:" dropdown, which is only (re)built by Load Columns.
        if not self._is_spreadsheet_type(self.current_doc_type):
            if self.matches:
                messagebox.showinfo(
                    APP_TITLE,
                    "This changes the numbering source -- click Scan again so the change is "
                    "picked up; the results already on screen won't update on their own.",
                )
            return
        if cross_type_changed and self.spreadsheet_columns:
            messagebox.showinfo(
                APP_TITLE,
                "This changes which document type(s) are available under \"Same identity "
                "as:\" -- click Load Columns again so the column checklist picks up the "
                "change; columns already listed won't update on their own.",
            )

    def _resolve_redact_existing_numbering(self):
        """Merges every currently-loaded SAME-document-type existing-key
        mapping (see on_redact_add_existing_keys()) into one
        existing_map/state usable by the current document type's
        numbering function. Returns (merged_or_None, error_message_or_None):
        merged is None with no error when nothing is loaded (the
        original, unseeded behavior). error_message is set (merged is
        None) when two or more loaded key files disagree with each other
        -- the same "don't guess through it" handling the Un-redact tab
        already uses for its own multi-key-file conflicts (see
        on_unredact_save()); the caller should treat that as blocking
        rather than silently picking one file's numbering over another's.

        This does NOT include cross-type (PDF-identity-linked) key files
        -- see _resolve_redact_cross_type_pools() for those.
        """
        doc_type = self.current_doc_type
        if not self.redact_existing_key_paths or not self._doc_type_supports_existing_key(doc_type):
            return None, None
        mappings = [self.redact_existing_key_maps[p] for p in self.redact_existing_key_paths]
        merged, conflicts = doc_type.merge_existing_numbering(mappings)
        if not conflicts:
            return merged, None
        lines = []
        for c in conflicts[:10]:
            if "column" in c:
                versus = " vs. ".join(c["placeholders"])
                lines.append(f"  - Column '{c['column']}', value {c['value']!r}: {versus}")
            else:
                versus = " vs. ".join(str(n) for n in c["numbers"])
                lines.append(f"  - {c['name']!r}: {versus}")
        more = f"\n  ... and {len(conflicts) - 10} more" if len(conflicts) > 10 else ""
        msg = (
            "The key file(s) you selected for consistent numbering disagree with each "
            "other -- the same name/value maps to a different placeholder in different "
            "files:\n\n" + "\n".join(lines) + more +
            "\n\nThis almost always means one of these key files doesn't actually belong "
            "with the others (e.g. the wrong file was picked, or it's a key file from an "
            "unrelated document). Fix the selection above (Use existing key file(s)... / "
            "Clear) and try again -- nothing was scanned or redacted."
        )
        return None, msg

    def _resolve_redact_cross_type_pools(self):
        """Merges every loaded cross-type key file (see
        on_redact_add_existing_keys()) into one identity pool PER LINKED
        ENTITY LABEL -- {entity_label: {normalized_name: number}} -- used
        to seed a linked spreadsheet column's placeholders at Redact &
        Save time (see _on_redact_save_spreadsheet()). Also returns each
        label's normalize_identity function (falling back to the generic
        trim/collapse-whitespace/lowercase one, same convention as Batch
        Redact's _batch_entity_label_options()).

        Returns (pools_by_label, normalize_fns_by_label, error_message).
        error_message is set (the first two are None) when two or more
        loaded key files for the SAME entity label disagree with each
        other -- the same "don't guess through it" handling
        _resolve_redact_existing_numbering() uses for same-type key
        files, just per label here since each label's files are merged
        (and could conflict) independently of any other label's.
        """
        by_label = {}
        for entry in self.redact_cross_type_key_maps.values():
            info = by_label.setdefault(entry["entity_label"], {"doc_type": entry["doc_type"], "mappings": []})
            info["mappings"].append(entry["mapping"])

        pools = {}
        normalize_fns = {}
        for label, info in by_label.items():
            dt = info["doc_type"]
            merged, conflicts = dt.merge_existing_numbering(info["mappings"])
            if conflicts:
                lines = []
                for c in conflicts[:10]:
                    if "column" in c:
                        versus = " vs. ".join(c["placeholders"])
                        lines.append(f"  - Column '{c['column']}', value {c['value']!r}: {versus}")
                    else:
                        versus = " vs. ".join(str(n) for n in c["numbers"])
                        lines.append(f"  - {c['name']!r}: {versus}")
                more = f"\n  ... and {len(conflicts) - 10} more" if len(conflicts) > 10 else ""
                msg = (
                    f"The key file(s) loaded for the '{label}' identity link disagree with "
                    "each other:\n\n" + "\n".join(lines) + more +
                    "\n\nThis almost always means one of these key files doesn't actually "
                    "belong with the others. Fix the selection above (Use existing key "
                    "file(s)... / Clear) and try again -- nothing was scanned or redacted."
                )
                return None, None, msg
            pools[label] = merged
            normalize_fns[label] = getattr(dt, "normalize_identity", _default_normalize_identity)
        return pools, normalize_fns, None

    def on_scan(self):
        if not self.input_path:
            return
        existing_map, key_error = self._resolve_redact_existing_numbering()
        if key_error:
            messagebox.showerror(APP_TITLE, key_error)
            return
        self._redact_existing_map = existing_map
        self.scan_btn.configure(state="disabled")
        self.redact_btn.configure(state="disabled")
        self._clear_results()
        self.progress.configure(mode="indeterminate")
        self.progress.pack(fill="x", padx=10, pady=(0, 4))
        self.progress.start(12)
        if self._is_spreadsheet_type(self.current_doc_type):
            self.summary_var.set("Reading column headers...")
            threading.Thread(target=self._list_columns_worker, daemon=True).start()
        else:
            self.summary_var.set("Scanning...")
            threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        doc_type = self.current_doc_type
        existing_map = self._redact_existing_map
        try:
            matches, unmatched, page_count = doc_type.scan(self.input_path)
            if existing_map is not None and getattr(doc_type, "SUPPORTS_BATCH", False):
                mapping = doc_type.assign_placeholder_numbers(matches, existing_map=existing_map)
            else:
                mapping = doc_type.assign_placeholder_numbers(matches)
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            # Default-argument capture (exc=exc), NOT a bare closure over
            # `exc` -- Python implicitly deletes an `except ... as exc`
            # binding the instant this except block exits, which happens
            # (via an implicit finally) before self.after()'s callback
            # actually runs. A plain `lambda: self._scan_failed(exc, tb)`
            # would raise NameError: cannot access free variable 'exc'
            # the moment Tk invokes it -- silently, since that error surfaces
            # inside Tk's callback machinery, not as anything the user sees.
            # Binding exc as a default argument copies its value into the
            # lambda's own scope right now, while it's still valid.
            self.after(0, lambda exc=exc, tb=tb: self._scan_failed(exc, tb))
            return
        self.after(0, lambda: self._scan_done(matches, unmatched, page_count, mapping))

    def _scan_failed(self, exc, tb=None):
        # Always show the exception TYPE, not just str(exc) -- some
        # exceptions (e.g. a bare KeyError) stringify to something
        # unhelpful like "None" on their own, and the type name alone is
        # often enough to point at the real cause. If a traceback was
        # captured, it's also saved to a log file (same convention as the
        # Redact/Un-redact operations' error logs) so the exact internal
        # line that failed can be found, not just guessed at.
        self.progress.stop()
        self.progress.pack_forget()
        self.scan_btn.configure(state="normal")
        log_note = ""
        if tb:
            log_path = write_error_log(
                "scan", self.input_path, "(no output -- scan/load only)", tb, near_path=self.input_path
            )
            if log_path:
                log_note = f"\n\nFull technical detail was saved to:\n{log_path}"
        action = "load this file's columns" if self._is_spreadsheet_type(self.current_doc_type) else "scan this file"
        messagebox.showerror(
            APP_TITLE,
            f"Could not {action}:\n\n{type(exc).__name__}: {exc}\n\n"
            "If this keeps happening, the file's layout may differ from what "
            "this document type expects -- please check before relying on it "
            f"for this file.{log_note}",
        )

    def _list_columns_worker(self):
        doc_type = self.current_doc_type
        try:
            columns = doc_type.list_columns(self.input_path)
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            # See the comment in _scan_worker() -- default-argument capture
            # is required here, a bare closure over `exc` raises NameError
            # (silently, inside Tk's callback machinery) once this except
            # block's implicit cleanup deletes the `exc` binding.
            self.after(0, lambda exc=exc, tb=tb: self._scan_failed(exc, tb))
            return
        self.after(0, lambda: self._columns_loaded(columns))

    def _columns_loaded(self, columns):
        doc_type = self.current_doc_type
        self.progress.stop()
        self.progress.pack_forget()
        self.scan_btn.configure(state="normal")
        self.spreadsheet_columns = columns
        self._populate_column_checklist(columns)
        if columns:
            self.summary_var.set(
                f"{len(columns)} column(s) found. Check the ones to redact below, then Redact & Save."
            )
            self.redact_btn.configure(state="normal")
        else:
            self.summary_var.set(
                "No columns found -- check that this workbook's header row is row 1 of a sheet."
            )
            messagebox.showwarning(APP_TITLE, doc_type.NO_MATCHES_WARNING)

    def _scan_done(self, matches, unmatched, page_count, mapping):
        doc_type = self.current_doc_type
        self.progress.stop()
        self.progress.pack_forget()
        self.matches = matches
        self.unmatched = unmatched
        self.page_count = page_count
        self.mapping = mapping
        self.scan_btn.configure(state="normal")

        for m in matches:
            self.tree.insert(
                "", "end", iid=str(id(m)),
                values=(
                    m.page_index + 1,
                    f"<{doc_type.ENTITY_LABEL} {m.placeholder_no}>",
                    m.display_text,
                    (m.context or "").strip(),
                    "Will redact",
                ),
            )
        for u in unmatched:
            self.review_tree.insert(
                "", "end",
                values=(u.page_index + 1, (u.context or "").strip(), u.raw_text),
            )

        n_unique = len(mapping)
        n_occurrences = len(matches)
        n_review = len(unmatched)
        self.notebook.tab(self.matches_tab, text=f"Detected items ({n_occurrences})")
        self.notebook.tab(self.review_tab, text=f"Needs review ({n_review})")

        entity_lower = doc_type.ENTITY_LABEL.lower()
        summary = (
            f"{page_count} page(s) scanned  •  {n_occurrences} {entity_lower} occurrence(s) found  •  "
            f"{n_unique} unique {entity_lower}(s)"
        )
        if n_review:
            summary += f"  •  ⚠ {n_review} line(s) need manual review before you trust this file"
        self.summary_var.set(summary)

        if n_occurrences > 0 or n_review > 0:
            self.redact_btn.configure(state="normal")
        if n_occurrences == 0 and n_review == 0:
            messagebox.showwarning(APP_TITLE, doc_type.NO_MATCHES_WARNING)

    def on_row_selected(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        match = self._match_by_iid(sel[0])
        if match:
            self._render_preview(match.page_index, match.rect)

    def on_row_double_click(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        match = self._match_by_iid(iid)
        if not match:
            return
        mid = id(match)
        if mid in self.excluded_ids:
            self.excluded_ids.discard(mid)
            self.tree.item(iid, tags=())
            self.tree.set(iid, "status", "Will redact")
        else:
            self.excluded_ids.add(mid)
            self.tree.item(iid, tags=("excluded",))
            self.tree.set(iid, "status", "Excluded")

    def _match_by_iid(self, iid):
        for m in self.matches:
            if str(id(m)) == iid:
                return m
        return None

    def _render_preview(self, page_index, rect):
        try:
            doc = fitz.open(self.input_path)
            page = doc[page_index]
            clip = fitz.Rect(
                max(rect.x0 - PREVIEW_PAD_X, 0),
                max(rect.y0 - PREVIEW_PAD_Y, 0),
                min(rect.x1 + PREVIEW_PAD_X * 6, page.rect.width),
                min(rect.y1 + PREVIEW_PAD_Y, page.rect.height),
            )
            pix = page.get_pixmap(clip=clip, matrix=fitz.Matrix(PREVIEW_ZOOM, PREVIEW_ZOOM))
            ppm = pix.tobytes("ppm")
            doc.close()
        except Exception:
            return
        self.preview_canvas.delete("all")
        photo = tk.PhotoImage(data=ppm)
        self._preview_photo = photo
        self.preview_canvas.create_image(10, 10, anchor="nw", image=photo)
        rx0 = (rect.x0 - clip.x0) * PREVIEW_ZOOM + 10
        ry0 = (rect.y0 - clip.y0) * PREVIEW_ZOOM + 10
        rx1 = (rect.x1 - clip.x0) * PREVIEW_ZOOM + 10
        ry1 = (rect.y1 - clip.y0) * PREVIEW_ZOOM + 10
        self.preview_canvas.create_rectangle(rx0, ry0, rx1, ry1, outline="#d00000", width=2)

    def on_redact_save(self):
        if self._is_spreadsheet_type(self.current_doc_type):
            self._on_redact_save_spreadsheet()
            return
        doc_type = self.current_doc_type
        included = [m for m in self.matches if id(m) not in self.excluded_ids]
        if self.unmatched:
            proceed = messagebox.askyesno(
                APP_TITLE,
                f"{len(self.unmatched)} line(s) in the 'Needs review' tab did not match "
                "the expected pattern and will NOT be redacted. If one of those lines "
                "actually contains something that should be redacted, it will still be "
                "visible in the output file.\n\nContinue anyway?",
                icon="warning",
            )
            if not proceed:
                return
        if not included:
            messagebox.showinfo(APP_TITLE, "There is nothing selected to redact.")
            return

        base = os.path.splitext(os.path.basename(self.input_path))[0]
        default_name = f"{base}_REDACTED.pdf"
        default_dir = os.path.dirname(self.input_path)
        output_path = filedialog.asksaveasfilename(
            title="Save redacted file as...",
            initialdir=default_dir,
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output_path:
            return

        lock_exc = check_file_not_locked(output_path)
        if lock_exc is not None:
            proceed = messagebox.askyesno(
                APP_TITLE, LOCKED_FILE_MESSAGE.format(path=output_path, exc=lock_exc), icon="warning"
            )
            if not proceed:
                return

        key_path = None
        if doc_type.SUPPORTS_KEY_FILE and self.save_key_var.get():
            key_path = os.path.splitext(output_path)[0] + "_NAME_KEY.csv"

        self.redact_btn.configure(state="disabled")
        self.scan_btn.configure(state="disabled")
        self.summary_var.set("Redacting...")
        self.progress.configure(mode="determinate", maximum=100, value=0)
        self.progress.pack(fill="x", padx=10, pady=(0, 4))
        threading.Thread(
            target=self._redact_worker, args=(doc_type, included, output_path, key_path), daemon=True
        ).start()

    def _on_redact_progress(self, current, total):
        pct = int((current / total) * 100) if total else 0
        unit = getattr(self.current_doc_type, "PROGRESS_UNIT_LABEL", "page")
        self.summary_var.set(f"Redacting -- {unit} {current} of {total}...")
        self.progress.configure(value=pct)
        self._set_title(f"{pct}% redacting")

    def _on_redact_saving_phase(self):
        self.summary_var.set(
            "Saving to disk -- finalizing the file (can take a little while on a very large file)..."
        )
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self._set_title("saving...")

    def _on_verify_progress(self, current, total):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        pct = int((current / total) * 100) if total else 0
        unit = getattr(self.current_doc_type, "PROGRESS_UNIT_LABEL", "page")
        self.summary_var.set(f"Verifying -- {unit} {current} of {total}...")
        self.progress.configure(value=pct)
        self._set_title(f"{pct}% verifying")

    @staticmethod
    def _make_instrumented_callback(gui_after, handler, label, logger, unit_label="page"):
        step = None
        seen_first = False

        def callback(current, total):
            nonlocal step, seen_first
            if not seen_first:
                seen_first = True
                logger.line(f"{label}: started ({total} {unit_label}(s) to process)")
            if step is None:
                step = max(1, total // 100)
            if current == total:
                logger.line(f"{label}: finished ({current} of {total} {unit_label}s)")
            if current == total or current % step == 0:
                gui_after(0, lambda: handler(current, total))

        return callback

    def _redact_worker(self, doc_type, included, output_path, key_path):
        logger = RunLogger(
            "redact",
            f"starting, document type: {doc_type.DISPLAY_NAME}, "
            f"{len(included)} occurrence(s), output: {output_path}",
        )
        try:
            doc_type.apply_redactions(
                self.input_path, output_path, included,
                progress_callback=self._make_instrumented_callback(
                    self.after, self._on_redact_progress, "Redacting", logger
                ),
                phase_callback=lambda phase: (
                    logger.line("Redacting: all pages drawn -- now writing final file to disk"),
                    self.after(0, self._on_redact_saving_phase),
                ),
            )
            logger.line("Saving to disk: finished -- file fully written")
            result = doc_type.verify_redaction(
                self.input_path, output_path, included,
                progress_callback=self._make_instrumented_callback(
                    self.after, self._on_verify_progress, "Verifying", logger
                ),
            )
            logger.line(f"Verifying: finished -- ok={result['ok']}")
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            logger.line(f"FAILED: {type(exc).__name__}: {exc}")
            # Default-argument capture required -- see _scan_worker()'s
            # comment: a bare `lambda: self._redact_failed(exc, tb, ...)`
            # would raise NameError once this except block's implicit
            # cleanup deletes the `exc` binding, silently, before Tk ever
            # runs the callback.
            self.after(0, lambda exc=exc, tb=tb: self._redact_failed(exc, tb, output_path))
            return

        key_error = None
        if result["ok"] and key_path:
            try:
                doc_type.write_key(key_path, self.input_path, output_path, included)
                logger.line("Name key CSV: written")
            except Exception as exc:  # noqa: BLE001
                key_error = str(exc)
                logger.line(f"Name key CSV: FAILED -- {key_error}")
        logger.line(f"Done -- total {time.time() - logger.run_start:.1f}s")

        self.after(0, lambda: self._redact_done(result, output_path, included, key_path, key_error))

    def _redact_failed(self, exc, tb, output_path):
        self._set_title()
        self.progress.stop()
        self.progress.pack_forget()
        self.scan_btn.configure(state="normal")
        self.redact_btn.configure(state="normal")
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
        log_path = write_error_log("redact", self.input_path, output_path, tb)
        log_note = f"\n\nFull error details were saved to:\n{log_path}" if log_path else ""
        messagebox.showerror(
            APP_TITLE,
            f"Redaction failed and no file was saved:\n\n{type(exc).__name__}: {exc}\n\n"
            "Please do not use any partially-written output file.\n\n"
            f"{FAILURE_TROUBLESHOOTING}{log_note}",
        )

    def _redact_done(self, result, output_path, included, key_path, key_error):
        doc_type = self.current_doc_type
        self._set_title()
        self.progress.stop()
        self.progress.pack_forget()
        self.scan_btn.configure(state="normal")
        self.redact_btn.configure(state="normal")

        if not result["ok"]:
            can_review = (
                hasattr(doc_type, "apply_pdf_leak_decisions")
                and bool(result.get("leaks"))
                and not result.get("structural_issue")
            )
            if can_review:
                self._open_pdf_leak_review(result, output_path, key_path, doc_type, included)
                return
            # Fall back to the original hard block -- either this plugin
            # doesn't support leak review at all, there were no structured
            # leaks to review, or something STRUCTURAL (not a single
            # leaked name) went wrong that a per-occurrence decision has
            # no way to fix.
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            issues = "\n".join(f"  - {i}" for i in result["issues"])
            log_path = write_selfcheck_log("redact", self.input_path, output_path, result["issues"])
            log_note = f"\n\nThis was also saved to:\n{log_path}" if log_path else ""
            messagebox.showerror(
                APP_TITLE,
                "The self-check on the redacted file found problems, so the file was "
                "NOT saved -- do not deliver anything from this run:\n\n"
                f"{issues}\n\n"
                f"Please report this to the developer with the source file.{log_note}",
            )
            return

        self._finish_redact_success_pdf(output_path, included, key_path, key_error)

    def _finish_redact_success_pdf(
        self, output_path, included, key_path, key_error, replaced=None, ignored=None, review_log_path=None
    ):
        """Shared tail for a successful PDF redaction -- reached either
        directly (self-check passed the first time, from _redact_done)
        or after a manual leak review resolved every occurrence (see
        _apply_pdf_leak_review_decisions()). replaced/ignored (lists of
        leak-review decision dicts) and review_log_path are only set on
        the manual-review path; their absence is what tells this apart
        from the plain, nothing-to-review case -- same convention as
        excel_spreadsheet.py's _finish_redact_success_spreadsheet().
        """
        doc_type = self.current_doc_type
        replaced = replaced or []
        ignored = ignored or []

        if ignored:
            selfcheck_line = (
                f"The self-check found {len(replaced) + len(ignored)} occurrence(s) that "
                "needed a decision; you reviewed each one manually before this file was saved."
            )
        else:
            selfcheck_line = (
                "The self-check confirmed everything targeted is fully removed and no other "
                "content was affected."
            )

        n_unique = len({m.placeholder_no for m in included})
        entity_lower = doc_type.ENTITY_LABEL.lower()
        msg = (
            f"Done. Redacted {len(included)} {entity_lower} occurrence(s) across {n_unique} "
            f"unique {entity_lower}(s).\n\nSaved to:\n{output_path}\n\n{selfcheck_line}"
        )
        if replaced:
            msg += (
                f"\n\n{len(replaced)} occurrence(s) flagged by the self-check were manually "
                "redacted during review."
            )
        if ignored:
            msg += (
                f"\n\n{len(ignored)} name occurrence(s) were deliberately left un-redacted in "
                "this output after manual review -- do NOT treat this file as fully "
                "de-identified without checking why. Exact page and name for each one is in "
                f"the review log:\n{review_log_path or '(could not be saved)'}"
            )
        if key_path and not key_error:
            msg += (
                f"\n\nA name key file was also saved to:\n{key_path}\n\n"
                "That file can reverse every placeholder back to the real value -- keep it "
                "with your own case file and do not send it along with the redacted file. "
                "Use the Un-redact tab to reverse it later."
            )
        elif key_error:
            msg += (
                f"\n\nThe redacted file above is complete and verified, but the name key "
                f"file could not be saved:\n{key_error}"
            )
        timing_log_path = os.path.join(tempfile.gettempdir(), "DocumentRedactor_redact_log.txt")
        if os.path.exists(timing_log_path):
            msg += (
                f"\n\nA timing breakdown for this run was recorded to:\n{timing_log_path}\n"
                "If this took longer than expected, that file is the fastest way to show "
                "exactly where the time went."
            )
        messagebox.showinfo(APP_TITLE, msg)
        self.summary_var.set(f"Saved: {output_path}")

    # ----------------------------------------------------------------
    # Redact tab: spreadsheet-shaped document types (column selection)
    # ----------------------------------------------------------------
    def _on_redact_save_spreadsheet(self):
        doc_type = self.current_doc_type
        selected = []
        for name, v in self.column_vars.items():
            if not v["selected"].get():
                continue
            label = v["label"].get().strip()
            if not label:
                messagebox.showerror(
                    APP_TITLE,
                    f'The placeholder label for column "{name}" is empty -- enter a label '
                    "or uncheck that column.",
                )
                return
            shared_label = None
            link_var = v.get("shared_identity")
            if link_var is not None:
                chosen = link_var.get()
                if chosen and chosen != _BATCH_NO_IDENTITY_LINK:
                    shared_label = chosen
            selected.append({"name": name, "label": label, "shared_identity_label": shared_label})
        if not selected:
            messagebox.showinfo(APP_TITLE, "Check at least one column to redact.")
            return

        existing_state, key_error = self._resolve_redact_existing_numbering()
        if key_error:
            messagebox.showerror(APP_TITLE, key_error)
            return

        cross_pools, cross_normalize_fns, cross_error = self._resolve_redact_cross_type_pools()
        if cross_error:
            messagebox.showerror(APP_TITLE, cross_error)
            return

        base = os.path.splitext(os.path.basename(self.input_path))[0]
        default_name = f"{base}_REDACTED.xlsx"
        default_dir = os.path.dirname(self.input_path)
        output_path = filedialog.asksaveasfilename(
            title="Save redacted workbook as...",
            initialdir=default_dir,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return

        lock_exc = check_file_not_locked(output_path)
        if lock_exc is not None:
            proceed = messagebox.askyesno(
                APP_TITLE, LOCKED_FILE_MESSAGE.format(path=output_path, exc=lock_exc), icon="warning"
            )
            if not proceed:
                return

        key_path = None
        if doc_type.SUPPORTS_KEY_FILE and self.save_key_var.get():
            key_path = os.path.splitext(output_path)[0] + "_REDACTION_KEY.csv"

        self.redact_btn.configure(state="disabled")
        self.scan_btn.configure(state="disabled")
        self.summary_var.set("Redacting...")
        self.progress.configure(mode="determinate", maximum=100, value=0)
        self.progress.pack(fill="x", padx=10, pady=(0, 4))
        threading.Thread(
            target=self._redact_worker_spreadsheet,
            args=(doc_type, selected, output_path, key_path, existing_state, cross_pools, cross_normalize_fns),
            daemon=True,
        ).start()

    def _redact_worker_spreadsheet(
        self, doc_type, selected, output_path, key_path, existing_state=None,
        cross_pools=None, cross_normalize_fns=None,
    ):
        logger = RunLogger(
            "redact",
            f"starting, document type: {doc_type.DISPLAY_NAME}, "
            f"{len(selected)} column(s) selected, output: {output_path}",
        )
        try:
            extra_kwargs = {}
            state = None
            if existing_state is not None and getattr(doc_type, "SUPPORTS_BATCH", False):
                state = dict(existing_state)

            linked_columns = [c for c in selected if c.get("shared_identity_label")]
            if linked_columns:
                # Cross-type identity linking, single-file analog of Batch
                # Redact's own version (see the near-identical loop in
                # _batch_redact_worker() above) -- for every column
                # manually linked to a loaded PDF-shaped key file's
                # identity pool (see _populate_column_checklist() and
                # _resolve_redact_cross_type_pools()), resolve every
                # distinct value in THIS column against that pool: reuse
                # the number the key file already assigned an identity,
                # or allocate the next one for a name the key file never
                # saw. Seeded straight into this column's own `state`
                # entry, in the {literal_value: placeholder_text} shape
                # apply_column_redaction() already expects.
                if state is None:
                    state = {}
                working_pools = {
                    label: dict(pool) for label, pool in (cross_pools or {}).items()
                }
                for col in linked_columns:
                    shared_label = col["shared_identity_label"]
                    pool = working_pools.setdefault(shared_label, {})
                    normalize_fn = (cross_normalize_fns or {}).get(shared_label, _default_normalize_identity)
                    distinct_values = doc_type.distinct_values_in_column(self.input_path, col["name"])
                    seed = state.setdefault(col["name"], {})
                    # Sorted, not raw set-iteration order: distinct_values is
                    # a plain set, and Python randomizes a set's internal
                    # ordering separately for every process launch (string
                    # hash randomization) -- looping over it as-is would
                    # assign a brand-new identity a DIFFERENT number every
                    # time this file is redacted, even with byte-identical
                    # input. That's invisible within a single Redact & Save
                    # click (the same numbers go into both the output file
                    # and its key CSV together), but it silently breaks
                    # un-redaction the moment a workbook and key CSV end up
                    # paired across two SEPARATE redaction runs -- the
                    # PDF-shared identities still match (those numbers come
                    # from the PDF's own fixed key), but newly-discovered
                    # identities can land on a different number in each run.
                    # Sorting first makes the assignment a pure function of
                    # the input data, identical on every run.
                    for literal_value in sorted(distinct_values, key=normalize_fn):
                        norm_key = normalize_fn(literal_value)
                        if norm_key not in pool:
                            pool[norm_key] = len(pool) + 1
                        seed[literal_value] = f"<{shared_label} {pool[norm_key]}>"

            if state is not None:
                extra_kwargs["state"] = state
            summary, key_rows = doc_type.apply_column_redaction(
                self.input_path, output_path, selected,
                progress_callback=self._make_instrumented_callback(
                    self.after, self._on_redact_progress, "Redacting", logger, unit_label="row"
                ),
                phase_callback=lambda phase: (
                    logger.line("Redacting: all rows processed -- now writing final file to disk"),
                    self.after(0, self._on_redact_saving_phase),
                ),
                **extra_kwargs,
            )
            logger.line("Saving to disk: finished -- file fully written")
            result = doc_type.verify_column_redaction(self.input_path, output_path, selected, key_rows)
            logger.line(f"Verifying: finished -- ok={result['ok']}")
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            logger.line(f"FAILED: {type(exc).__name__}: {exc}")
            # Default-argument capture required -- see _scan_worker()'s
            # comment.
            self.after(0, lambda exc=exc, tb=tb: self._redact_failed(exc, tb, output_path))
            return

        key_error = None
        if result["ok"] and key_path:
            try:
                doc_type.write_key(key_path, self.input_path, output_path, key_rows)
                logger.line("Redaction key CSV: written")
            except Exception as exc:  # noqa: BLE001
                key_error = str(exc)
                logger.line(f"Redaction key CSV: FAILED -- {key_error}")
        logger.line(f"Done -- total {time.time() - logger.run_start:.1f}s")

        self.after(
            0,
            lambda: self._redact_done_spreadsheet(
                result, summary, output_path, key_path, key_error, doc_type, selected, key_rows
            ),
        )

    def _redact_done_spreadsheet(self, result, summary, output_path, key_path, key_error, doc_type, selected, key_rows):
        self._set_title()
        self.progress.stop()
        self.progress.pack_forget()
        self.scan_btn.configure(state="normal")
        self.redact_btn.configure(state="normal")

        if not result["ok"]:
            can_review = bool(result.get("leaks")) and not result.get("shape_issue") and not result.get("other_issue")
            if can_review:
                self._open_leak_review(result, summary, output_path, key_path, doc_type, selected, key_rows)
                return
            # Fall back to the original hard block -- either there were no
            # structured leaks to review (shouldn't normally happen when
            # 'ok' is False, but keep this as a safety net) or something
            # STRUCTURAL (not a per-cell value) went wrong that a manual,
            # cell-by-cell review has no way to fix -- e.g. a sheet's
            # dimensions changed unexpectedly, which isn't something a
            # "replace this cell" / "ignore this cell" decision addresses.
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            issues = "\n".join(f"  - {i}" for i in result["issues"])
            log_path = write_selfcheck_log("redact", self.input_path, output_path, result["issues"])
            log_note = f"\n\nThis was also saved to:\n{log_path}" if log_path else ""
            messagebox.showerror(
                APP_TITLE,
                "The self-check on the redacted file found problems, so the file was "
                "NOT saved -- do not deliver anything from this run:\n\n"
                f"{issues}\n\n"
                f"Please report this to the developer with the source file.{log_note}",
            )
            return

        self._finish_redact_success_spreadsheet(summary, output_path, key_path, key_error)

    def _finish_redact_success_spreadsheet(
        self, summary, output_path, key_path, key_error, replaced=None, ignored=None, review_log_path=None
    ):
        """Shared tail for a successful spreadsheet redaction -- reached
        either directly (self-check passed the first time) or after a
        manual review resolved every leak (see _apply_leak_review_decisions()).
        replaced/ignored (lists of decision dicts) and review_log_path are
        only set on the manual-review path; their absence is what tells
        this apart from the plain, nothing-to-review case.
        """
        replaced = replaced or []
        ignored = ignored or []

        if ignored:
            # Something IS still un-redacted in this file, by deliberate
            # choice -- the plain "self-check confirmed none of the
            # original values remain" line would be false here, so it's
            # replaced rather than just appended to.
            selfcheck_line = (
                f"The self-check found {len(replaced) + len(ignored)} value(s) that needed a "
                "decision; you reviewed each one manually before this file was saved."
            )
        else:
            selfcheck_line = (
                "The self-check confirmed none of the original values remain anywhere in the "
                "output workbook, and the workbook's sheet structure is unchanged."
            )

        msg = (
            f"Done. Redacted {summary['cells_redacted']} cell(s) across "
            f"{summary['columns_found']} column(s) ({summary['unique_values_redacted']} unique "
            f"value(s) total).\n\nSaved to:\n{output_path}\n\n{selfcheck_line}"
        )
        if replaced:
            msg += (
                f"\n\n{len(replaced)} value(s) flagged by the self-check were manually "
                "replaced with their placeholder during review."
            )
        if ignored:
            msg += (
                f"\n\n{len(ignored)} original value(s) were deliberately left un-redacted in "
                "this output after manual review -- do NOT treat this file as fully "
                "de-identified without checking why. Exact location and value for each one "
                f"is in the review log:\n{review_log_path or '(could not be saved)'}"
            )
        if summary["columns_not_found"]:
            shown = ", ".join(summary["columns_not_found"])
            msg += (
                f"\n\nNote: {len(summary['columns_not_found'])} selected column(s) were "
                f"never found in this workbook's header row ({shown}) -- nothing was "
                "redacted for those; double check the column name(s)."
            )
        if key_path and not key_error:
            msg += (
                f"\n\nA redaction key file was also saved to:\n{key_path}\n\n"
                "That file can reverse every placeholder back to the real value -- keep it "
                "with your own case file and do not send it along with the redacted file. "
                "Use the Un-redact tab to reverse it later."
            )
        elif key_error:
            msg += (
                f"\n\nThe redacted file above is complete and verified, but the redaction "
                f"key file could not be saved:\n{key_error}"
            )
        timing_log_path = os.path.join(tempfile.gettempdir(), "DocumentRedactor_redact_log.txt")
        if os.path.exists(timing_log_path):
            msg += (
                f"\n\nA timing breakdown for this run was recorded to:\n{timing_log_path}\n"
                "If this took longer than expected, that file is the fastest way to show "
                "exactly where the time went."
            )
        messagebox.showinfo(APP_TITLE, msg)
        self.summary_var.set(f"Saved: {output_path}")

    # ----------------------------------------------------------------
    # Manual review: shown instead of the hard block when a self-check
    # failure is made up entirely of per-cell value leaks (see
    # verify_column_redaction()'s 'leaks' field) -- lets the user decide,
    # cell by cell, to either replace the leftover value with the
    # placeholder already assigned to it elsewhere, or knowingly leave it
    # (both choices are written to a durable review log before the file
    # is treated as finished; see write_manual_review_log()).
    # ----------------------------------------------------------------
    def _build_leak_review_rows(self, parent, leaks):
        """Builds one review row (a bordered Frame with location, row
        context, and Replace/Ignore controls) per leak inside `parent`.
        Shared by the single-file Redact tab's review window
        (_open_leak_review) and the Batch Redact tab's per-file review
        window (_open_batch_leak_review) so the two never drift apart --
        same enriched location description (column header + other values
        in the same row) either way. Returns row_vars, one dict per leak:
        {'leak', 'action_var', 'placeholder_var', 'placeholder_by_choice'}.
        """
        row_vars = []
        for leak in leaks:
            frame = tk.Frame(parent, relief="groove", borderwidth=1)
            frame.pack(fill="x", pady=4, padx=2)

            loc_text = f"{leak['sheet']}!{leak['coordinate']}"
            if leak["is_header_row"]:
                loc_text += "   (row 1 -- header row, not a data row)"
            elif leak.get("column_header"):
                loc_text += f"   -- in column '{leak['column_header']}'"
            tk.Label(frame, text=loc_text, font=("TkDefaultFont", 9, "bold"), anchor="w").pack(
                fill="x", padx=8, pady=(6, 0)
            )

            row_context = leak.get("row_context") or []
            if row_context and not leak["is_header_row"]:
                context_text = "Same row also has: " + "; ".join(
                    f"{c['label']}: {c['value']}" for c in row_context
                )
                tk.Label(
                    frame, text=context_text, wraplength=760, justify="left", anchor="w",
                    fg="#444444",
                ).pack(fill="x", padx=8, pady=(0, 0))

            value_row = tk.Frame(frame)
            value_row.pack(fill="x", padx=8, pady=(2, 2))
            tk.Label(value_row, text="Value still present:").pack(side="left")
            value_entry = tk.Entry(value_row)
            value_entry.insert(0, leak["value"])
            value_entry.configure(state="readonly")
            value_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

            choices = [
                f"{p['placeholder']}  (from column '{p['column']}')" for p in leak["placeholders"]
            ]
            placeholder_by_choice = {c: p["placeholder"] for c, p in zip(choices, leak["placeholders"])}
            placeholder_var = tk.StringVar(value=choices[0])

            # Default to Ignore for a header-row match (very likely just a
            # label coincidence, not real data) and Replace for a data-row
            # match (very likely a genuine leak worth scrubbing) -- either
            # way, the user sees and can override the default before
            # anything is actually applied.
            action_var = tk.StringVar(value="ignore" if leak["is_header_row"] else "replace")

            action_row = tk.Frame(frame)
            action_row.pack(fill="x", padx=8, pady=(0, 6))
            tk.Radiobutton(action_row, text="Replace with:", variable=action_var, value="replace").pack(side="left")
            if len(choices) > 1:
                combo = ttk.Combobox(
                    action_row, textvariable=placeholder_var, values=choices, state="readonly", width=44
                )
                combo.pack(side="left", padx=(4, 12))
            else:
                tk.Label(action_row, text=choices[0]).pack(side="left", padx=(4, 12))
            tk.Radiobutton(
                action_row, text="Ignore (leave value as-is)", variable=action_var, value="ignore"
            ).pack(side="left")

            row_vars.append({
                "leak": leak,
                "action_var": action_var,
                "placeholder_var": placeholder_var,
                "placeholder_by_choice": placeholder_by_choice,
            })
        return row_vars

    @staticmethod
    def _collect_leak_decisions(row_vars):
        """row_vars (from _build_leak_review_rows) -> the 'decisions' list
        shape both apply_leak_decisions() and the review log writer
        expect: [{'sheet', 'coordinate', 'action', 'value', 'placeholder'
        (only when action == 'replace')}, ...]."""
        decisions = []
        for rv in row_vars:
            leak = rv["leak"]
            action = rv["action_var"].get()
            entry = {
                "sheet": leak["sheet"],
                "coordinate": leak["coordinate"],
                "action": action,
                "value": leak["value"],
            }
            if action == "replace":
                choice = rv["placeholder_var"].get()
                entry["placeholder"] = rv["placeholder_by_choice"][choice]
            decisions.append(entry)
        return decisions

    def _open_leak_review(self, result, summary, output_path, key_path, doc_type, selected, key_rows):
        win = tk.Toplevel(self)
        win.title(f"{APP_TITLE} -- Manual Review")
        win.geometry("820x560")
        win.transient(self)
        win.grab_set()

        intro = (
            "The self-check found value(s) that still match something you redacted -- "
            "either in the header row (row 1 is never touched by redaction, so this can "
            "just mean a column header happens to say the same thing as a redacted value) "
            "or in a data cell outside the column(s) you chose. Nothing has been changed "
            "yet. Decide for each one below, then click Apply.\n\n"
            "Replace overwrites that exact cell with the placeholder shown. Ignore leaves "
            "the cell exactly as it is. Either choice is written to a review log saved "
            "alongside this file, so there is a durable record of what was left in and why."
        )
        tk.Label(win, text=intro, wraplength=780, justify="left", anchor="w").pack(
            fill="x", padx=12, pady=(12, 6)
        )

        outer, inner = self._build_scrollable_frame(win)
        outer.pack(fill="both", expand=True, padx=12, pady=6)

        row_vars = self._build_leak_review_rows(inner, result["leaks"])

        btn_row = tk.Frame(win)
        btn_row.pack(fill="x", padx=12, pady=(6, 12))

        def on_cancel():
            win.destroy()
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            self.scan_btn.configure(state="normal")
            self.redact_btn.configure(state="normal")
            messagebox.showinfo(
                APP_TITLE, "Manual review cancelled -- no file was saved. Nothing was changed."
            )

        def on_apply():
            decisions = self._collect_leak_decisions(row_vars)
            win.destroy()
            self._apply_leak_review_decisions(
                decisions, doc_type, selected, key_rows, output_path, key_path, summary
            )

        tk.Button(btn_row, text="Cancel (discard this run)", command=on_cancel).pack(side="right")
        tk.Button(btn_row, text="Apply Decisions", command=on_apply).pack(side="right", padx=(0, 8))

    def _apply_leak_review_decisions(self, decisions, doc_type, selected, key_rows, output_path, key_path, summary):
        try:
            doc_type.apply_leak_decisions(output_path, decisions)
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            log_path = write_error_log("redact", self.input_path, output_path, tb)
            log_note = f"\n\nFull error details were saved to:\n{log_path}" if log_path else ""
            messagebox.showerror(
                APP_TITLE,
                f"Applying your manual review decisions failed:\n\n{type(exc).__name__}: {exc}\n\n"
                f"The file at\n{output_path}\nmay be partially modified -- do not send it out; "
                f"re-run the redaction from scratch instead.{log_note}",
            )
            self.scan_btn.configure(state="normal")
            self.redact_btn.configure(state="normal")
            return

        replaced = [d for d in decisions if d["action"] == "replace"]
        ignored = [d for d in decisions if d["action"] == "ignore"]
        review_log_path = write_manual_review_log("redact", self.input_path, output_path, replaced, ignored)

        # Re-run the exact same self-check used the first time, rather
        # than trusting that applying the replacements worked -- a
        # coordinate mismatch, a save that silently failed, or anything
        # else unexpected should still be caught here rather than waved
        # through just because the user made decisions about it.
        final_result = doc_type.verify_column_redaction(self.input_path, output_path, selected, key_rows)
        ignored_keys = {(d["sheet"], d["coordinate"]) for d in ignored}
        unresolved = [
            leak for leak in final_result.get("leaks", [])
            if (leak["sheet"], leak["coordinate"]) not in ignored_keys
        ]
        still_blocked = bool(unresolved) or final_result.get("shape_issue") or final_result.get("other_issue")

        self.scan_btn.configure(state="normal")
        self.redact_btn.configure(state="normal")

        if still_blocked:
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            remaining = "\n".join(f"  - {leak['sheet']}!{leak['coordinate']}" for leak in unresolved)
            other = "\n".join(
                f"  - {i}" for i in (final_result.get("shape_issue"), final_result.get("other_issue")) if i
            )
            remaining_text = "\n".join(t for t in (remaining, other) if t)
            messagebox.showerror(
                APP_TITLE,
                "After applying your decisions, the self-check still found unresolved "
                f"problem(s), so the file was NOT saved:\n\n{remaining_text}\n\n"
                "This shouldn't normally happen -- please report it to the developer with "
                f"the source file.\n\nReview log:\n{review_log_path or '(could not be saved)'}",
            )
            return

        key_error = None
        if key_path:
            try:
                doc_type.write_key(key_path, self.input_path, output_path, key_rows)
            except Exception as exc:  # noqa: BLE001
                key_error = str(exc)

        self._finish_redact_success_spreadsheet(
            summary, output_path, key_path, key_error, replaced=replaced, ignored=ignored,
            review_log_path=review_log_path,
        )

    # ----------------------------------------------------------------
    # PDF manual review: the same "review instead of hard-block" idea as
    # the spreadsheet side above, but for a pdf_report type's leaked NAME
    # occurrences (see dentrix_audit_trail.py's verify_redaction()
    # 'leaks' field: one entry per distinct leaked name, each with a list
    # of exact on-page occurrences). Unlike a spreadsheet cell, a PDF
    # occurrence has no natural single-cell identity, so review happens
    # per OCCURRENCE (page + exact rectangle), not per name -- the same
    # name can leak in more than one place on a page or across pages, and
    # each needs its own decision. A 'redact' decision is run back through
    # apply_redactions() itself (see apply_pdf_leak_decisions()'s
    # docstring) so it's genuinely removed from the content stream, not
    # just covered -- and it reuses the SAME <Patient N> label already
    # assigned to that person elsewhere in the file, via name_to_patient_no
    # (built from `included`, the exact match list this run already
    # redacted). A structural_issue (not a per-occurrence name leak) is
    # NEVER reviewable here -- see _redact_done()'s can_review check.
    # ----------------------------------------------------------------
    def _build_pdf_leak_review_rows(self, parent, leaks, doc_type=None):
        """Builds one review row per LEAKED OCCURRENCE (not per name --
        see this section's header comment) inside `parent`. Shared by the
        single-file Redact tab's review window (_open_pdf_leak_review)
        and the Batch Redact tab's per-file review window
        (_open_batch_pdf_leak_review). Returns row_vars, one dict per
        occurrence: {'name', 'page_index', 'rect', 'action_var',
        'add_to_ignore_var' (or None)}.

        `doc_type`, when given, is checked for add_name_position_override
        -- a plugin-specific manual-override mechanism (currently only
        dentrix_audit_trail.py has one; see its docstring) that lets a
        reviewer permanently silence a specific recurring false positive
        (e.g. "Credit Card", a payment-method value that isn't a name at
        all but can land in the name position on some report layouts)
        rather than being asked about the exact same non-name text on
        every future file. When the plugin doesn't support it, this is
        omitted entirely rather than showing a control that would do
        nothing.
        """
        row_vars = []
        for leak in leaks:
            name = leak["name"]
            for occ in leak.get("occurrences", []):
                frame = tk.Frame(parent, relief="groove", borderwidth=1)
                frame.pack(fill="x", pady=4, padx=2)

                loc_text = f"Page {occ['page_index'] + 1}"
                tk.Label(frame, text=loc_text, font=("TkDefaultFont", 9, "bold"), anchor="w").pack(
                    fill="x", padx=8, pady=(6, 0)
                )

                if occ.get("context"):
                    tk.Label(
                        frame, text=f"Nearby text: {occ['context']}", wraplength=760, justify="left",
                        anchor="w", fg="#444444",
                    ).pack(fill="x", padx=8, pady=(0, 0))

                if occ.get("suspicious"):
                    tk.Label(
                        frame,
                        text=(
                            "⚠ This is longer than a typical name -- it may include "
                            "description/amount text swept in alongside it, not just the "
                            "name. Check exactly what's here before choosing Redact, since "
                            "Redact permanently removes everything shown below, not just a name."
                        ),
                        wraplength=760, justify="left", anchor="w", fg="#8a5a00",
                    ).pack(fill="x", padx=8, pady=(0, 0))

                value_row = tk.Frame(frame)
                value_row.pack(fill="x", padx=8, pady=(2, 2))
                tk.Label(value_row, text="Name still present:").pack(side="left")
                value_entry = tk.Entry(value_row)
                value_entry.insert(0, name)
                value_entry.configure(state="readonly")
                value_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

                # Default to Redact -- a leaked name occurrence found by
                # the self-check is very likely a genuine miss worth
                # scrubbing (unlike the spreadsheet side's header-row
                # case, there's no equivalent "probably just a coincidence"
                # situation here). The user sees and can override this
                # before anything is actually applied.
                #
                # A 'suspicious' occurrence (see dentrix_audit_trail.py's
                # find_unresolved_name_positions()) is the one exception:
                # it's flagged because it runs longer than any real name
                # plausibly would, which usually means the position-based
                # scan swept real report content (a description, an
                # amount) in alongside or instead of the name, on a report
                # layout whose column spacing doesn't give the normal
                # gap-based detection enough of a signal to find the true
                # boundary. Redact would PERMANENTLY DELETE whatever
                # that swept text actually is -- defaulting to Ignore
                # here means nothing is destroyed until a human has
                # actually looked at it and confirmed Redact is safe.
                action_var = tk.StringVar(value="ignore" if occ.get("suspicious") else "redact")

                action_row = tk.Frame(frame)
                action_row.pack(fill="x", padx=8, pady=(0, 6))
                tk.Radiobutton(
                    action_row, text="Redact this occurrence", variable=action_var, value="redact"
                ).pack(side="left")
                tk.Radiobutton(
                    action_row, text="Ignore (leave name as-is)", variable=action_var, value="ignore"
                ).pack(side="left", padx=(12, 0))

                # Optional, plugin-specific: let the reviewer durably
                # silence this EXACT text (case/whitespace-insensitive,
                # not a substring match -- see add_name_position_override()'s
                # docstring) so it's never flagged again on any future
                # file, instead of re-litigating the same known-not-a-name
                # report vocabulary every time it shows up. Independent of
                # the Redact/Ignore choice above -- it only affects future
                # self-checks, never this occurrence's own disposition.
                add_to_ignore_var = None
                if doc_type is not None and hasattr(doc_type, "add_name_position_override"):
                    add_to_ignore_var = tk.BooleanVar(value=False)
                    ignore_row = tk.Frame(frame)
                    ignore_row.pack(fill="x", padx=8, pady=(0, 6))
                    tk.Checkbutton(
                        ignore_row,
                        text=(
                            f'This text ("{name}") is not actually a name -- '
                            "never flag it again on future files"
                        ),
                        variable=add_to_ignore_var,
                    ).pack(side="left")

                row_vars.append({
                    "name": name,
                    "page_index": occ["page_index"],
                    "rect": occ["rect"],
                    "action_var": action_var,
                    "add_to_ignore_var": add_to_ignore_var,
                })
        return row_vars

    @staticmethod
    def _collect_pdf_leak_decisions(row_vars):
        """row_vars (from _build_pdf_leak_review_rows) -> the 'decisions'
        list shape both apply_pdf_leak_decisions() and the review log
        writer expect: [{'name', 'page_index', 'rect', 'action',
        'add_to_ignore'}, ...]. 'add_to_ignore' is always a plain bool
        (False when the plugin doesn't support the manual-override
        mechanism at all, i.e. add_to_ignore_var was None)."""
        decisions = []
        for rv in row_vars:
            add_to_ignore_var = rv.get("add_to_ignore_var")
            decisions.append({
                "name": rv["name"],
                "page_index": rv["page_index"],
                "rect": rv["rect"],
                "action": rv["action_var"].get(),
                "add_to_ignore": bool(add_to_ignore_var.get()) if add_to_ignore_var is not None else False,
            })
        return decisions

    def _open_pdf_leak_review(self, result, output_path, key_path, doc_type, included):
        win = tk.Toplevel(self)
        win.title(f"{APP_TITLE} -- Manual Review")
        win.geometry("820x560")
        win.transient(self)
        win.grab_set()

        entity_lower = doc_type.ENTITY_LABEL.lower()
        intro = (
            f"The self-check found {entity_lower} name(s) that are still visible somewhere "
            "in the redacted file. Nothing has been changed yet. Decide for each occurrence "
            "below, then click Apply.\n\n"
            "Redact genuinely removes that exact occurrence from the file (the same "
            "whole-page rebuild every other redaction in this file goes through, not a "
            "visual cover) and reuses the same placeholder already used for that person "
            "elsewhere in this file. Ignore leaves the name exactly as it is. Either choice "
            "is written to a review log saved alongside this file, so there is a durable "
            "record of what was left in and why."
        )
        tk.Label(win, text=intro, wraplength=780, justify="left", anchor="w").pack(
            fill="x", padx=12, pady=(12, 6)
        )

        outer, inner = self._build_scrollable_frame(win)
        outer.pack(fill="both", expand=True, padx=12, pady=6)

        row_vars = self._build_pdf_leak_review_rows(inner, result["leaks"], doc_type)

        btn_row = tk.Frame(win)
        btn_row.pack(fill="x", padx=12, pady=(6, 12))

        def on_cancel():
            win.destroy()
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            self.scan_btn.configure(state="normal")
            self.redact_btn.configure(state="normal")
            messagebox.showinfo(
                APP_TITLE, "Manual review cancelled -- no file was saved. Nothing was changed."
            )

        def on_apply():
            decisions = self._collect_pdf_leak_decisions(row_vars)
            win.destroy()
            self._apply_pdf_leak_review_decisions(decisions, doc_type, included, output_path, key_path)

        tk.Button(btn_row, text="Cancel (discard this run)", command=on_cancel).pack(side="right")
        tk.Button(btn_row, text="Apply Decisions", command=on_apply).pack(side="right", padx=(0, 8))

    def _apply_pdf_leak_review_decisions(self, decisions, doc_type, included, output_path, key_path):
        # Persist any "never flag this text again" choices BEFORE the
        # re-verify pass below runs -- that pass calls verify_redaction()
        # -> find_unresolved_name_positions() again, and it should already
        # see today's new overrides take effect, not just future runs.
        if hasattr(doc_type, "add_name_position_override"):
            for d in decisions:
                if d.get("add_to_ignore"):
                    doc_type.add_name_position_override(d["name"])

        name_to_patient_no = {m.raw_name: m.patient_no for m in included}
        try:
            redacted_matches = doc_type.apply_pdf_leak_decisions(output_path, decisions, name_to_patient_no)
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            log_path = write_error_log("redact", self.input_path, output_path, tb)
            log_note = f"\n\nFull error details were saved to:\n{log_path}" if log_path else ""
            messagebox.showerror(
                APP_TITLE,
                f"Applying your manual review decisions failed:\n\n{type(exc).__name__}: {exc}\n\n"
                f"The file at\n{output_path}\nmay be partially modified -- do not send it out; "
                f"re-run the redaction from scratch instead.{log_note}",
            )
            self.scan_btn.configure(state="normal")
            self.redact_btn.configure(state="normal")
            return

        # A 'redact' decision on a leak that scan_pdf() never recognized
        # as a name in the first place (see dentrix_audit_trail.py's
        # find_unresolved_name_positions()) has no existing entry in
        # `included` -- apply_pdf_leak_decisions() just assigned it a
        # brand-new placeholder number instead (mutating name_to_patient_no
        # in place, which is why the `replaced` lookup below already sees
        # it). `included_plus` folds those newly-redacted matches in
        # everywhere `included` would otherwise be used from here on, so
        # the re-verify pass, the "Redacted N occurrence(s)" summary, and
        # the key file all account for them -- without this, the name
        # would be correctly removed from the PDF but absent from the key
        # file, making it impossible to un-redact later.
        included_plus = included + redacted_matches

        replaced = [
            {
                "name": d["name"],
                "page_index": d["page_index"],
                "placeholder": f"<{doc_type.ENTITY_LABEL} {name_to_patient_no.get(d['name'])}>",
            }
            for d in decisions if d["action"] == "redact"
        ]
        ignored = [{"name": d["name"], "page_index": d["page_index"]} for d in decisions if d["action"] == "ignore"]
        review_log_path = write_pdf_manual_review_log("redact", self.input_path, output_path, replaced, ignored)

        # Re-run the exact same self-check used the first time, rather
        # than trusting that applying the redact decisions worked -- same
        # reasoning as the spreadsheet side's _apply_leak_review_decisions.
        final_result = doc_type.verify_redaction(self.input_path, output_path, included_plus)
        ignored_keys = {(d["name"], d["page_index"]) for d in decisions if d["action"] == "ignore"}
        unresolved = [
            {"name": leak["name"], "page_index": occ["page_index"]}
            for leak in final_result.get("leaks", [])
            for occ in leak.get("occurrences", [])
            if (leak["name"], occ["page_index"]) not in ignored_keys
        ]
        still_blocked = bool(unresolved) or bool(final_result.get("structural_issue"))

        self.scan_btn.configure(state="normal")
        self.redact_btn.configure(state="normal")

        if still_blocked:
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            remaining = "\n".join(f"  - page {u['page_index'] + 1}: {u['name']!r}" for u in unresolved)
            other = final_result.get("structural_issue") or ""
            remaining_text = "\n".join(t for t in (remaining, other) if t)
            messagebox.showerror(
                APP_TITLE,
                "After applying your decisions, the self-check still found unresolved "
                f"problem(s), so the file was NOT saved:\n\n{remaining_text}\n\n"
                "This shouldn't normally happen -- please report it to the developer with "
                f"the source file.\n\nReview log:\n{review_log_path or '(could not be saved)'}",
            )
            return

        key_error = None
        if key_path:
            try:
                doc_type.write_key(key_path, self.input_path, output_path, included_plus)
            except Exception as exc:  # noqa: BLE001
                key_error = str(exc)

        self._finish_redact_success_pdf(
            output_path, included_plus, key_path, key_error, replaced=replaced, ignored=ignored,
            review_log_path=review_log_path,
        )

    def _batch_run_leak_review(self, name, index, total, input_path, output_path, dt, selected_columns, result, key_rows):
        """Runs a manual leak review for ONE spreadsheet file inside a
        Batch Redact run, reusing the exact same review UI as the
        single-file Redact tab's _open_leak_review (via
        _build_leak_review_rows) -- same enriched location description,
        same Replace/Ignore choices per leak -- but scoped to just this
        file: Apply fixes this file and the batch continues; Cancel skips
        ONLY this one file (its output is deleted) and the REST of the
        batch still continues. This is what keeps a batch from
        automatically discarding everything just because one file needed
        a human decision.

        This is called from the BACKGROUND worker thread
        (_batch_redact_worker). All Tk widget creation happens on the
        main thread via self.after() (see _open_batch_leak_review); this
        method just blocks the calling worker thread on a
        threading.Event until the user clicks Apply or Cancel, so the
        worker loop can pick back up (or skip to the next file) the
        instant this call returns -- the Event and the plain dict used to
        pass the outcome back are the only things touched from both
        threads, never a Tk widget itself.

        Returns {'ok': bool, 'skipped': bool, 'replaced': list, 'ignored': list}.
        'skipped' True means Cancel was chosen (or applying/re-verifying
        the decisions still failed) -- the caller should exclude this
        file from the batch's saved outputs and move on.
        """
        done = threading.Event()
        outcome = {}
        self.after(
            0,
            lambda: self._open_batch_leak_review(
                name, index, total, input_path, output_path, dt, selected_columns, result, key_rows, done, outcome
            ),
        )
        done.wait()
        return outcome

    def _open_batch_leak_review(
        self, name, index, total, input_path, output_path, dt, selected_columns, result, key_rows, done, outcome
    ):
        win = tk.Toplevel(self)
        win.title(f"{APP_TITLE} -- Manual Review ({index} of {total}: {name})")
        win.geometry("820x560")
        win.transient(self)
        win.grab_set()

        intro = (
            f'While redacting "{name}" ({index} of {total} in this batch), the self-check '
            "found value(s) that still match something redacted -- either in the header row "
            "(row 1 is never touched by redaction, so this can just mean a column header "
            "happens to say the same thing as a redacted value) or in a data cell outside "
            "the column(s) chosen. Nothing has been changed yet. Decide for each one below, "
            "then click Apply.\n\n"
            "Replace overwrites that exact cell with the placeholder shown. Ignore leaves "
            "the cell exactly as it is. Either choice is written to a review log saved "
            "alongside this file.\n\n"
            "Cancel skips ONLY this one file -- nothing is saved for it, but the rest of "
            "the batch keeps going; it does not stop the whole run."
        )
        tk.Label(win, text=intro, wraplength=780, justify="left", anchor="w").pack(
            fill="x", padx=12, pady=(12, 6)
        )

        outer, inner = self._build_scrollable_frame(win)
        outer.pack(fill="both", expand=True, padx=12, pady=6)

        row_vars = self._build_leak_review_rows(inner, result["leaks"])

        btn_row = tk.Frame(win)
        btn_row.pack(fill="x", padx=12, pady=(6, 12))

        def finish(final_outcome):
            outcome.clear()
            outcome.update(final_outcome)
            done.set()

        def on_cancel():
            win.destroy()
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            self._batch_log(f"{name}: manual review cancelled -- this file was skipped, nothing saved for it.")
            finish({"ok": False, "skipped": True, "replaced": [], "ignored": []})

        def on_apply():
            decisions = self._collect_leak_decisions(row_vars)
            win.destroy()
            try:
                dt.apply_leak_decisions(output_path, decisions)
            except Exception as exc:  # noqa: BLE001
                tb = traceback.format_exc()
                log_path = write_error_log("batch_redact", input_path, output_path, tb)
                log_note = f"\n\nFull error details were saved to:\n{log_path}" if log_path else ""
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except OSError:
                        pass
                messagebox.showerror(
                    APP_TITLE,
                    f'Applying your manual review decisions to "{name}" failed:\n\n'
                    f"{type(exc).__name__}: {exc}\n\nThis file was skipped; the rest of the "
                    f"batch will continue.{log_note}",
                )
                self._batch_log(f"{name}: applying review decisions failed -- this file was skipped.")
                finish({"ok": False, "skipped": True, "replaced": [], "ignored": []})
                return

            replaced = [d for d in decisions if d["action"] == "replace"]
            ignored = [d for d in decisions if d["action"] == "ignore"]
            write_manual_review_log("batch_redact", input_path, output_path, replaced, ignored)

            # Re-run the exact same self-check used the first time, same
            # reasoning as the single-file tab's _apply_leak_review_decisions.
            final_result = dt.verify_column_redaction(input_path, output_path, selected_columns, key_rows)
            ignored_keys = {(d["sheet"], d["coordinate"]) for d in ignored}
            unresolved = [
                leak for leak in final_result.get("leaks", [])
                if (leak["sheet"], leak["coordinate"]) not in ignored_keys
            ]
            still_blocked = bool(unresolved) or final_result.get("shape_issue") or final_result.get("other_issue")

            if still_blocked:
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except OSError:
                        pass
                remaining = "\n".join(f"  - {leak['sheet']}!{leak['coordinate']}" for leak in unresolved)
                messagebox.showerror(
                    APP_TITLE,
                    f'After applying your decisions, the self-check on "{name}" still found '
                    f"unresolved problem(s), so this file was NOT saved:\n\n{remaining}\n\n"
                    "This shouldn't normally happen -- please report it to the developer with "
                    "the source file. The rest of the batch will continue.",
                )
                self._batch_log(f"{name}: still unresolved after review -- this file was skipped.")
                finish({"ok": False, "skipped": True, "replaced": [], "ignored": []})
                return

            self._batch_log(
                f"{name}: manual review applied -- {len(replaced)} replaced, {len(ignored)} left as-is."
            )
            finish({"ok": True, "skipped": False, "replaced": replaced, "ignored": ignored})

        tk.Button(btn_row, text="Cancel (skip only this file)", command=on_cancel).pack(side="right")
        tk.Button(btn_row, text="Apply Decisions", command=on_apply).pack(side="right", padx=(0, 8))

    def _batch_run_pdf_leak_review(self, name, index, total, input_path, output_path, dt, matches, result):
        """PDF analog of _batch_run_leak_review() -- see that method's
        docstring for the threading pattern (the worker thread blocks on
        a threading.Event while the review Toplevel runs on the main
        thread). `matches` is this file's own full match list (from
        dt.scan() + dt.assign_placeholder_numbers(), the same list this
        file was already redacted with) -- reused here to build
        name_to_patient_no, so a 'redact' decision reuses the correct
        existing placeholder instead of guessing a new one.

        Returns {'ok': bool, 'skipped': bool, 'replaced': list, 'ignored': list}.
        """
        done = threading.Event()
        outcome = {}
        self.after(
            0,
            lambda: self._open_batch_pdf_leak_review(
                name, index, total, input_path, output_path, dt, matches, result, done, outcome
            ),
        )
        done.wait()
        return outcome

    def _open_batch_pdf_leak_review(
        self, name, index, total, input_path, output_path, dt, matches, result, done, outcome
    ):
        win = tk.Toplevel(self)
        win.title(f"{APP_TITLE} -- Manual Review ({index} of {total}: {name})")
        win.geometry("820x560")
        win.transient(self)
        win.grab_set()

        entity_lower = dt.ENTITY_LABEL.lower()
        intro = (
            f'While redacting "{name}" ({index} of {total} in this batch), the self-check '
            f"found {entity_lower} name(s) still visible in the output. Nothing has been "
            "changed yet. Decide for each occurrence below, then click Apply.\n\n"
            "Redact genuinely removes that exact occurrence from the file (the same "
            "whole-page rebuild every other redaction in this file goes through, not a "
            "visual cover) and reuses the same placeholder already used for that person "
            "elsewhere in this file. Ignore leaves the name exactly as it is. Either choice "
            "is written to a review log saved alongside this file.\n\n"
            "Cancel skips ONLY this one file -- nothing is saved for it, but the rest of "
            "the batch keeps going; it does not stop the whole run."
        )
        tk.Label(win, text=intro, wraplength=780, justify="left", anchor="w").pack(
            fill="x", padx=12, pady=(12, 6)
        )

        outer, inner = self._build_scrollable_frame(win)
        outer.pack(fill="both", expand=True, padx=12, pady=6)

        row_vars = self._build_pdf_leak_review_rows(inner, result["leaks"], dt)

        btn_row = tk.Frame(win)
        btn_row.pack(fill="x", padx=12, pady=(6, 12))

        def finish(final_outcome):
            outcome.clear()
            outcome.update(final_outcome)
            done.set()

        def on_cancel():
            win.destroy()
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            self._batch_log(f"{name}: manual review cancelled -- this file was skipped, nothing saved for it.")
            finish({"ok": False, "skipped": True, "replaced": [], "ignored": []})

        def on_apply():
            decisions = self._collect_pdf_leak_decisions(row_vars)
            win.destroy()

            # See the single-file tab's _apply_pdf_leak_review_decisions
            # for why this happens before apply_pdf_leak_decisions() and
            # the re-verify pass below, not after.
            if hasattr(dt, "add_name_position_override"):
                for d in decisions:
                    if d.get("add_to_ignore"):
                        dt.add_name_position_override(d["name"])

            name_to_patient_no = {m.raw_name: m.patient_no for m in matches}
            try:
                redacted_matches = dt.apply_pdf_leak_decisions(output_path, decisions, name_to_patient_no)
            except Exception as exc:  # noqa: BLE001
                tb = traceback.format_exc()
                log_path = write_error_log("batch_redact", input_path, output_path, tb)
                log_note = f"\n\nFull error details were saved to:\n{log_path}" if log_path else ""
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except OSError:
                        pass
                messagebox.showerror(
                    APP_TITLE,
                    f'Applying your manual review decisions to "{name}" failed:\n\n'
                    f"{type(exc).__name__}: {exc}\n\nThis file was skipped; the rest of the "
                    f"batch will continue.{log_note}",
                )
                self._batch_log(f"{name}: applying review decisions failed -- this file was skipped.")
                finish({"ok": False, "skipped": True, "replaced": [], "ignored": []})
                return

            # See the single-file tab's _apply_pdf_leak_review_decisions
            # for why this is needed: a 'redact' decision on a leak
            # scan_pdf() never recognized as a name (find_unresolved_name_
            # positions()) got a brand-new placeholder number just now,
            # with no entry yet in `matches` -- matches_plus folds it in
            # everywhere `matches` is used from here on (the re-verify
            # pass and the caller's own key-file bookkeeping), so it ends
            # up in the batch's combined key file too, not just the PDF.
            matches_plus = matches + redacted_matches

            replaced = [
                {
                    "name": d["name"],
                    "page_index": d["page_index"],
                    "placeholder": f"<{dt.ENTITY_LABEL} {name_to_patient_no.get(d['name'])}>",
                }
                for d in decisions if d["action"] == "redact"
            ]
            ignored = [
                {"name": d["name"], "page_index": d["page_index"]} for d in decisions if d["action"] == "ignore"
            ]
            write_pdf_manual_review_log("batch_redact", input_path, output_path, replaced, ignored)

            # Re-run the exact same self-check used the first time, same
            # reasoning as the single-file tab's _apply_pdf_leak_review_decisions.
            final_result = dt.verify_redaction(input_path, output_path, matches_plus)
            ignored_keys = {(d["name"], d["page_index"]) for d in decisions if d["action"] == "ignore"}
            unresolved = [
                {"name": leak["name"], "page_index": occ["page_index"]}
                for leak in final_result.get("leaks", [])
                for occ in leak.get("occurrences", [])
                if (leak["name"], occ["page_index"]) not in ignored_keys
            ]
            still_blocked = bool(unresolved) or bool(final_result.get("structural_issue"))

            if still_blocked:
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except OSError:
                        pass
                remaining = "\n".join(f"  - page {u['page_index'] + 1}: {u['name']!r}" for u in unresolved)
                messagebox.showerror(
                    APP_TITLE,
                    f'After applying your decisions, the self-check on "{name}" still found '
                    f"unresolved problem(s), so this file was NOT saved:\n\n{remaining}\n\n"
                    "This shouldn't normally happen -- please report it to the developer with "
                    "the source file. The rest of the batch will continue.",
                )
                self._batch_log(f"{name}: still unresolved after review -- this file was skipped.")
                finish({"ok": False, "skipped": True, "replaced": [], "ignored": []})
                return

            self._batch_log(
                f"{name}: manual review applied -- {len(replaced)} redacted, {len(ignored)} left as-is."
            )
            finish({
                "ok": True, "skipped": False, "replaced": replaced, "ignored": ignored,
                "new_matches": redacted_matches,
            })

        tk.Button(btn_row, text="Cancel (skip only this file)", command=on_cancel).pack(side="right")
        tk.Button(btn_row, text="Apply Decisions", command=on_apply).pack(side="right", padx=(0, 8))

    # ================================================================
    # BATCH REDACT TAB
    #
    # Redacts several files together, in one run, so the same real-world
    # value gets the same placeholder everywhere it appears -- across
    # files of the SAME document type automatically, and across DIFFERENT
    # document types (e.g. a PDF audit trail and an Excel export in the
    # same batch) when the user manually links a spreadsheet column to a
    # PDF type's identity. See document_types/base.py's "BATCH REDACTION"
    # and "BATCH REDACTION: CROSS-TYPE IDENTITY LINKING" sections for the
    # plugin contract this relies on (SUPPORTS_BATCH, the
    # existing_map/state carry-forward parameter, write_key_batch,
    # normalize_identity).
    #
    # Deliberately simpler than the single-file Redact tab in one
    # respect: no per-match preview/exclude UI (everything a
    # scan/column-selection finds gets redacted). But a self-check
    # finding leftover values in a file's output no longer automatically
    # fails the whole batch -- it opens the same manual review screen the
    # single-file tab uses: for a spreadsheet file, see
    # _batch_run_leak_review / _open_batch_leak_review; for a PDF file,
    # see _batch_run_pdf_leak_review / _open_batch_pdf_leak_review (both
    # built on the same _build_pdf_leak_review_rows /
    # apply_pdf_leak_decisions the single-file Redact tab's PDF review
    # uses). Either way it's scoped to just that one file: Apply fixes it
    # and the batch continues, Cancel skips only that file and the rest
    # still runs. Only a NON-reviewable problem still aborts the WHOLE
    # batch, all-or-nothing -- a shape/structure issue (spreadsheet) or a
    # structural_issue (PDF) that no per-cell/per-occurrence decision can
    # fix. That's because every file's numbering depends on the files
    # processed before it, so a batch that stopped partway through on one
    # of THOSE can't be trusted as a self-consistent set.
    #
    # A batch's file list can mix document types -- each file is resolved
    # to its own plugin automatically by extension when added (see
    # _resolve_batch_doc_type()). Cross-type identity linking is always a
    # DELIBERATE, MANUAL choice: a "Same identity as:" dropdown next to
    # each spreadsheet column, populated only with the entity labels of
    # pdf_report-shaped document types actually present in the current
    # file list -- never inferred from column names or file names. See
    # base.py for why an automatic guess here would be dangerous (it
    # could silently merge two different real people under one
    # placeholder).
    # ================================================================
    def _build_batch_tab(self, root):
        pad = {"padx": 10, "pady": 6}

        if not self.batch_doc_types:
            ttk.Label(
                root,
                text="No document type currently supports Batch Redact mode.",
                font=("", 10, "bold"),
            ).pack(anchor="w", **pad)
            return

        # The Redact All button (with its status line and progress bar)
        # lives in its own footer frame, packed side="bottom" BEFORE the
        # scrollable body below -- that claims its space first, so it
        # always stays visible at the bottom of the tab no matter how
        # tall the setup controls above it get (file list, column
        # checklist, output folder, save-key option, progress log easily
        # add up to more than fits in the window, especially with a
        # spreadsheet's column checklist open). Previously everything
        # shared one plain, non-scrolling frame, and on a normal-sized
        # window "Redact All" itself could end up pushed below the
        # visible area with no way to reach it. Everything else still
        # goes in the scrollable body beneath, so it stays reachable too.
        footer = ttk.Frame(root)
        footer.pack(side="bottom", fill="x")

        action_row = ttk.Frame(footer)
        action_row.pack(fill="x", **pad)
        self.batch_redact_btn = ttk.Button(
            action_row, text="Redact All", command=self.on_batch_redact, state="disabled"
        )
        self.batch_redact_btn.pack(side="left")

        self.batch_summary_var = tk.StringVar(value="Add files, choose an output folder, then Redact All.")
        ttk.Label(footer, textvariable=self.batch_summary_var, font=("", 10, "bold")).pack(anchor="w", padx=10)

        # Created but not packed here -- on_batch_redact() packs it when a
        # run starts and pack_forget()s it when done, same as before.
        self.batch_progress = ttk.Progressbar(footer, mode="determinate", maximum=100)

        outer, body = self._build_scrollable_frame(root)
        outer.pack(fill="both", expand=True)

        intro = ttk.Label(
            body,
            text="Redact several files together, in one run, so the same real-world value "
                 "(e.g. one person's name) gets the SAME placeholder everywhere it appears -- "
                 "across files of the same document type automatically, and across DIFFERENT "
                 "document types (e.g. a PDF audit trail and an Excel export together) when "
                 "you manually link a spreadsheet column to a PDF type's identity below. "
                 "Produces one combined key file per document type in the batch. Batch mode "
                 "redacts everything it detects automatically; it does not offer the "
                 "per-match preview/exclude review the single-file Redact tab does -- use "
                 "that tab first for any file that needs that level of curation, then include "
                 "it in a batch once you're satisfied with it.",
            wraplength=1000, justify="left", foreground="#333",
        )
        intro.pack(anchor="w", **pad)

        files_frame = ttk.LabelFrame(body, text="Files in this batch (any supported document type)")
        files_frame.pack(fill="both", expand=False, padx=10, pady=6)

        list_row = ttk.Frame(files_frame)
        list_row.pack(fill="x", padx=6, pady=6)
        self.batch_tree = ttk.Treeview(
            list_row, columns=("path", "type"), show="headings", height=6, selectmode="extended"
        )
        self.batch_tree.heading("path", text="File")
        self.batch_tree.heading("type", text="Document type")
        self.batch_tree.column("path", width=560, anchor="w")
        self.batch_tree.column("type", width=220, anchor="w")
        self.batch_tree.pack(side="left", fill="both", expand=True)
        tree_scroll = ttk.Scrollbar(list_row, orient="vertical", command=self.batch_tree.yview)
        tree_scroll.pack(side="left", fill="y")
        self.batch_tree.configure(yscrollcommand=tree_scroll.set)

        btn_col = ttk.Frame(files_frame)
        btn_col.pack(side="left", padx=(6, 6), pady=6, fill="y")
        ttk.Button(btn_col, text="Add Files...", command=self.on_batch_add_files).pack(fill="x", pady=(0, 4))
        ttk.Button(btn_col, text="Remove Selected", command=self.on_batch_remove_selected).pack(fill="x", pady=(0, 4))
        ttk.Button(btn_col, text="Clear All", command=self.on_batch_clear_files).pack(fill="x")

        # Spreadsheet-only: a column checklist built from the UNION of
        # header names found across every selected spreadsheet file (see
        # on_batch_load_columns / _batch_list_columns_worker). Shown only
        # when the file list contains at least one spreadsheet-type file
        # -- see _refresh_batch_panels(). Each row optionally carries a
        # "Same identity as:" link, shown only when the file list also
        # contains at least one pdf_report-type file (see
        # _batch_columns_loaded()).
        self.batch_column_frame = ttk.LabelFrame(
            body, text="Columns to redact (found across the selected spreadsheet files)"
        )
        load_cols_row = ttk.Frame(self.batch_column_frame)
        load_cols_row.pack(fill="x", padx=6, pady=(6, 0))
        self.batch_load_columns_btn = ttk.Button(
            load_cols_row, text="Load Columns", command=self.on_batch_load_columns
        )
        self.batch_load_columns_btn.pack(side="left")
        ttk.Label(
            load_cols_row,
            text="  (reads every selected spreadsheet file's header row -- do this again if "
                 "you change the file list)",
            foreground="#888",
        ).pack(side="left")
        col_scroll_outer, self.batch_column_list_inner = self._build_scrollable_frame(self.batch_column_frame)
        col_scroll_outer.pack(fill="both", expand=True, padx=6, pady=6)
        self.batch_column_frame.configure(height=240)
        self.batch_column_frame.pack_propagate(False)

        self.batch_out_row = ttk.Frame(body)
        self.batch_out_row.pack(fill="x", **pad)
        ttk.Label(self.batch_out_row, text="Save redacted files to:").pack(side="left")
        self.batch_output_dir_var = tk.StringVar(value="(no folder selected)")
        ttk.Label(self.batch_out_row, textvariable=self.batch_output_dir_var, foreground="#333").pack(
            side="left", padx=8, fill="x", expand=True
        )
        ttk.Button(self.batch_out_row, text="Browse...", command=self.on_batch_browse_output_dir).pack(side="left")

        opt_row = ttk.Frame(body)
        opt_row.pack(fill="x", **pad)
        self.batch_save_key_var = tk.BooleanVar(value=True)
        self.batch_save_key_check = ttk.Checkbutton(
            opt_row,
            text="Save combined key file(s) for the whole batch (one per document type)",
            variable=self.batch_save_key_var,
        )
        self.batch_save_key_check.pack(side="left")

        log_frame = ttk.LabelFrame(body, text="Progress")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.batch_log_text = tk.Text(log_frame, height=10, state="disabled", wrap="word")
        self.batch_log_text.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.batch_log_text.yview)
        log_scroll.pack(side="left", fill="y", pady=4)
        self.batch_log_text.configure(yscrollcommand=log_scroll.set)

        self._refresh_batch_panels()

    @staticmethod
    def _batch_key_filename(doc_type):
        """Combined key filename for one document type's group within a
        batch -- kept stable/predictable (matching the names this tab has
        always used) for the two built-in types, with a generic fallback
        for any future SUPPORTS_BATCH plugin."""
        if doc_type.ID == "dentrix_audit_trail":
            return "BATCH_NAME_KEY.csv"
        if doc_type.ID == "excel_spreadsheet":
            return "BATCH_REDACTION_KEY.csv"
        return f"BATCH_KEY_{doc_type.ID}.csv"

    def _resolve_batch_doc_type(self, path):
        """Which batch-capable document type a file belongs to, purely by
        its extension against each type's FILE_TYPES patterns. Returns
        None if zero or more than one type claims that extension (nothing
        today is ambiguous -- PDF vs Excel don't overlap -- but this stays
        honest if a future document type's extensions ever do)."""
        ext = os.path.splitext(path)[1].lower()
        matches = []
        for dt in self.batch_doc_types:
            exts = set()
            for _, pattern in dt.FILE_TYPES:
                pattern = pattern.lower()
                if pattern == "*.*":
                    continue
                exts.update(p.strip() for p in pattern.split(";"))
            if f"*{ext}" in exts:
                matches.append(dt)
        return matches[0] if len(matches) == 1 else None

    def _batch_entity_label_options(self):
        """The distinct ENTITY_LABELs (and each one's normalize_identity
        function) belonging to pdf_report-shaped files currently in the
        batch's file list -- what the "Same identity as:" dropdown offers
        a spreadsheet column, and what the worker uses to normalize a
        linked column's cell values the same way the linked PDF type
        normalizes its own matches. Only labels for TYPES ACTUALLY PRESENT
        in this batch are offered -- linking to an identity pool no file
        in this run will ever populate would be pointless."""
        labels = []
        seen = set()
        fn_by_label = {}
        for f in self.batch_files:
            dt = f["doc_type"]
            if self._is_spreadsheet_type(dt):
                continue
            if dt.ENTITY_LABEL not in seen:
                seen.add(dt.ENTITY_LABEL)
                labels.append(dt.ENTITY_LABEL)
                fn_by_label[dt.ENTITY_LABEL] = getattr(dt, "normalize_identity", _default_normalize_identity)
        return labels, fn_by_label

    def _refresh_batch_panels(self):
        has_spreadsheet = any(self._is_spreadsheet_type(f["doc_type"]) for f in self.batch_files)
        if has_spreadsheet:
            self.batch_column_frame.pack(fill="both", expand=True, padx=10, pady=6, before=self.batch_out_row)
        else:
            self.batch_column_frame.pack_forget()

    def on_batch_add_files(self):
        if not self.batch_doc_types:
            return
        filetypes = []
        seen_patterns = set()
        for dt in self.batch_doc_types:
            for label, pattern in dt.FILE_TYPES:
                if pattern == "*.*":
                    continue
                if pattern not in seen_patterns:
                    seen_patterns.add(pattern)
                    filetypes.append((label, pattern))
        filetypes.append(("All files", "*.*"))
        paths = filedialog.askopenfilenames(title="Select files for this batch", filetypes=filetypes)
        if not paths:
            return

        existing_paths = {f["path"] for f in self.batch_files}
        unresolved = []
        for p in paths:
            if p in existing_paths:
                continue
            dt = self._resolve_batch_doc_type(p)
            if dt is None:
                unresolved.append(p)
                continue
            self.batch_files.append({"path": p, "doc_type": dt})
            self.batch_tree.insert("", "end", iid=p, values=(p, dt.DISPLAY_NAME))
            existing_paths.add(p)

        if unresolved:
            listed = "\n".join(f"  - {p}" for p in unresolved)
            messagebox.showwarning(
                APP_TITLE,
                f"{len(unresolved)} file(s) were not added because their type couldn't be "
                f"determined automatically from their extension:\n\n{listed}",
            )

        # A changed file list invalidates any previously loaded column
        # list -- a different set of spreadsheet files may have different
        # headers, and the set of pdf_report types available to link to
        # may have changed too. Rather than silently redact against a
        # stale list, force a fresh Load Columns before Redact All
        # becomes available again.
        self._clear_batch_column_checklist()
        self._refresh_batch_panels()
        self._update_batch_redact_btn_state()

    def on_batch_remove_selected(self):
        sel = list(self.batch_tree.selection())
        for iid in sel:
            self.batch_tree.delete(iid)
        removed = set(sel)
        self.batch_files = [f for f in self.batch_files if f["path"] not in removed]
        self._clear_batch_column_checklist()
        self._refresh_batch_panels()
        self._update_batch_redact_btn_state()

    def on_batch_clear_files(self):
        self.batch_tree.delete(*self.batch_tree.get_children())
        self.batch_files = []
        self._clear_batch_column_checklist()
        self._refresh_batch_panels()
        self._update_batch_redact_btn_state()

    def on_batch_browse_output_dir(self):
        path = filedialog.askdirectory(title="Choose a folder to save the redacted files to")
        if not path:
            return
        self.batch_output_dir = path
        self.batch_output_dir_var.set(path)
        self._update_batch_redact_btn_state()

    def _clear_batch_column_checklist(self):
        for child in self.batch_column_list_inner.winfo_children():
            child.destroy()
        self.batch_column_vars = {}

    def _update_batch_redact_btn_state(self):
        ok = bool(self.batch_files) and bool(self.batch_output_dir)
        if ok and any(self._is_spreadsheet_type(f["doc_type"]) for f in self.batch_files):
            ok = any(v["selected"].get() for v in self.batch_column_vars.values())
        self.batch_redact_btn.configure(state="normal" if ok else "disabled")

    def on_batch_load_columns(self):
        if not any(self._is_spreadsheet_type(f["doc_type"]) for f in self.batch_files):
            messagebox.showinfo(APP_TITLE, "Add at least one spreadsheet-type file first.")
            return
        self.batch_redact_btn.configure(state="disabled")
        self.batch_load_columns_btn.configure(state="disabled")
        self.batch_summary_var.set("Reading column headers from every selected spreadsheet file...")
        threading.Thread(target=self._batch_list_columns_worker, daemon=True).start()

    def _batch_list_columns_worker(self):
        merged = {}   # column name -> {"sample", "default_label", "file_count"}
        order = []
        spreadsheet_files = [f for f in self.batch_files if self._is_spreadsheet_type(f["doc_type"])]
        try:
            for f in spreadsheet_files:
                for col in f["doc_type"].list_columns(f["path"]):
                    if col.name not in merged:
                        merged[col.name] = {
                            "sample": col.sample,
                            "default_label": col.default_label,
                            "file_count": 0,
                        }
                        order.append(col.name)
                    merged[col.name]["file_count"] += 1
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            # See _scan_worker()'s comment -- default-argument capture is
            # required, a bare closure over `exc` raises NameError once
            # this except block's implicit cleanup deletes the binding.
            self.after(0, lambda exc=exc, tb=tb: self._batch_scan_failed(exc, tb))
            return
        entity_labels, normalize_fns = self._batch_entity_label_options()
        self.after(
            0,
            lambda: self._batch_columns_loaded(order, merged, len(spreadsheet_files), entity_labels, normalize_fns),
        )

    def _batch_scan_failed(self, exc, tb=None):
        self.batch_load_columns_btn.configure(state="normal")
        log_note = ""
        if tb:
            log_path = write_error_log("batch_load_columns", "(multiple files)", "(no output -- load only)", tb)
            if log_path:
                log_note = f"\n\nFull technical detail was saved to:\n{log_path}"
        messagebox.showerror(
            APP_TITLE,
            f"Could not read columns:\n\n{type(exc).__name__}: {exc}\n\n"
            f"Check that every selected spreadsheet file is a valid workbook of its type.{log_note}",
        )
        self.batch_summary_var.set("Add files, choose an output folder, then Redact All.")
        self._update_batch_redact_btn_state()

    def _batch_columns_loaded(self, order, merged, n_files, entity_labels, normalize_fns):
        self.batch_load_columns_btn.configure(state="normal")
        self._clear_batch_column_checklist()
        self.batch_identity_normalize_fns = normalize_fns
        if not order:
            self.batch_summary_var.set(
                "No columns found across the selected spreadsheet files -- check that each "
                "file's header row is row 1 of a sheet."
            )
            self._update_batch_redact_btn_state()
            return

        for name in order:
            info = merged[name]
            row = ttk.Frame(self.batch_column_list_inner)
            row.pack(fill="x", pady=2, padx=4)

            sel_var = tk.BooleanVar(value=False)
            label_var = tk.StringVar(value=info["default_label"])
            ttk.Checkbutton(row, variable=sel_var, command=self._update_batch_redact_btn_state).pack(side="left")

            found_note = f"  (found in {info['file_count']} of {n_files} file(s))"
            sample_note = f'  --  e.g. "{info["sample"]}"' if info["sample"] else "  --  (no sample value found)"
            ttk.Label(
                row, text=f"{name}{found_note}{sample_note}", width=46, anchor="w"
            ).pack(side="left", padx=(4, 10))

            ttk.Label(row, text="Placeholder label:").pack(side="left")
            label_entry = ttk.Entry(row, textvariable=label_var, width=16)
            label_entry.pack(side="left", padx=(4, 0))

            # "Same identity as:" -- a MANUAL, per-column link to one of
            # this batch's pdf_report-shaped document types, offered only
            # when at least one is present in the file list. Never
            # inferred from the column's own name -- see base.py's
            # cross-type identity linking section for why. Picking a link
            # forces this column's placeholder label to match the linked
            # type's ENTITY_LABEL (and locks it), since a matching NUMBER
            # with a different placeholder LABEL wouldn't actually read as
            # the same placeholder.
            link_var = None
            if entity_labels:
                ttk.Label(row, text="  Same identity as:").pack(side="left", padx=(10, 2))
                link_var = tk.StringVar(value=_BATCH_NO_IDENTITY_LINK)

                def on_link_change(
                    _event=None, label_var=label_var, label_entry=label_entry,
                    link_var=link_var, default_label=info["default_label"],
                ):
                    chosen = link_var.get()
                    if chosen == _BATCH_NO_IDENTITY_LINK:
                        label_var.set(default_label)
                        label_entry.configure(state="normal")
                    else:
                        label_var.set(chosen)
                        label_entry.configure(state="disabled")

                link_combo = ttk.Combobox(
                    row, textvariable=link_var, state="readonly", width=22,
                    values=[_BATCH_NO_IDENTITY_LINK] + entity_labels,
                )
                link_combo.bind("<<ComboboxSelected>>", on_link_change)
                link_combo.pack(side="left")

            self.batch_column_vars[name] = {"selected": sel_var, "label": label_var, "shared_identity": link_var}

        self.batch_summary_var.set(
            f"{len(order)} column(s) found across {n_files} spreadsheet file(s). Check the "
            "ones to redact below, then Redact All."
        )
        self._update_batch_redact_btn_state()

    def _batch_log_clear(self):
        self.batch_log_text.configure(state="normal")
        self.batch_log_text.delete("1.0", "end")
        self.batch_log_text.configure(state="disabled")

    def _batch_log(self, text):
        self.batch_log_text.configure(state="normal")
        self.batch_log_text.insert("end", text + "\n")
        self.batch_log_text.see("end")
        self.batch_log_text.configure(state="disabled")

    def on_batch_redact(self):
        if not self.batch_files:
            messagebox.showinfo(APP_TITLE, "Add at least one file first.")
            return
        if not self.batch_output_dir:
            messagebox.showinfo(APP_TITLE, "Choose an output folder first.")
            return

        has_spreadsheet = any(self._is_spreadsheet_type(f["doc_type"]) for f in self.batch_files)
        selected_columns = []
        if has_spreadsheet:
            for name, v in self.batch_column_vars.items():
                if not v["selected"].get():
                    continue
                label = v["label"].get().strip()
                if not label:
                    messagebox.showerror(
                        APP_TITLE,
                        f'The placeholder label for column "{name}" is empty -- enter a '
                        "label or uncheck that column.",
                    )
                    return
                shared_label = None
                if v.get("shared_identity") is not None:
                    val = v["shared_identity"].get()
                    if val and val != _BATCH_NO_IDENTITY_LINK:
                        shared_label = val
                selected_columns.append({"name": name, "label": label, "shared_identity_label": shared_label})
            if not selected_columns:
                messagebox.showinfo(APP_TITLE, "Check at least one column to redact.")
                return

        # Every output filename is derived from its own source file's name
        # (same extension preserved) -- two different source files (e.g.
        # from two different folders) sharing a basename AND extension
        # would silently overwrite one another partway through the run,
        # so this is checked up front rather than discovered after some
        # files are already written.
        planned_names = {}
        dupes = set()
        for f in self.batch_files:
            p = f["path"]
            out_name = f"{os.path.splitext(os.path.basename(p))[0]}_REDACTED{os.path.splitext(p)[1]}"
            if out_name in planned_names and planned_names[out_name] != p:
                dupes.add(out_name)
            planned_names[out_name] = p
        if dupes:
            listed = "\n".join(f"  - {d}" for d in sorted(dupes))
            messagebox.showerror(
                APP_TITLE,
                "Two or more selected files would produce the same output filename, which "
                f"would overwrite each other:\n\n{listed}\n\n"
                "Rename one of the source files (or remove the duplicate) and try again.",
            )
            return

        distinct_doctypes = []
        seen_ids = set()
        for f in self.batch_files:
            dt = f["doc_type"]
            if dt.ID not in seen_ids:
                seen_ids.add(dt.ID)
                distinct_doctypes.append(dt)

        key_paths = {}   # dt.ID -> path
        if self.batch_save_key_var.get():
            existing = []
            for dt in distinct_doctypes:
                if not getattr(dt, "SUPPORTS_KEY_FILE", False):
                    continue
                kp = os.path.join(self.batch_output_dir, self._batch_key_filename(dt))
                key_paths[dt.ID] = kp
                if os.path.exists(kp):
                    existing.append(kp)
            if existing:
                listed = "\n".join(f"  - {p}" for p in existing)
                proceed = messagebox.askyesno(
                    APP_TITLE,
                    f"Combined key file(s) already exist here and will be overwritten:\n\n"
                    f"{listed}\n\nContinue?",
                    icon="warning",
                )
                if not proceed:
                    return

        normalize_fns = dict(self.batch_identity_normalize_fns)

        self.batch_redact_btn.configure(state="disabled")
        if has_spreadsheet:
            self.batch_load_columns_btn.configure(state="disabled")
        self.batch_progress.configure(mode="determinate", maximum=max(len(self.batch_files), 1), value=0)
        self.batch_progress.pack(fill="x", padx=10, pady=(0, 4))
        self._batch_log_clear()
        self._batch_log(f"Starting batch redaction of {len(self.batch_files)} file(s)...")
        threading.Thread(
            target=self._batch_redact_worker,
            args=(list(self.batch_files), self.batch_output_dir, selected_columns, key_paths, normalize_fns),
            daemon=True,
        ).start()

    def _batch_redact_worker(self, files, output_dir, selected_columns, key_paths, normalize_fns):
        logger = RunLogger(
            "batch_redact",
            f"starting, {len(files)} file(s), output dir: {output_dir}",
        )
        produced_outputs = []     # output paths written so far -- for all-or-nothing cleanup
        per_file_summaries = []   # human-readable per-file lines for the final report
        state = {}                 # shared across the whole run -- see module-level comment
                                    # near _BATCH_IDENTITY_PREFIX for this dict's two kinds of
                                    # entries (spreadsheet column state, and identity pools).
        unmatched_by_file = []     # pdf: [(source_basename, [UnmatchedLine, ...]), ...]
        skipped_files = []         # [(name, issues), ...] -- spreadsheet files skipped after
                                    # the user cancelled (or couldn't resolve) a manual leak
                                    # review; the rest of the batch still completes normally.
        reviewed_files = []        # [(name, replaced_count, ignored_count), ...] -- spreadsheet
                                    # files that went through manual review and were RESOLVED
                                    # (kept in the batch); an ignored_count > 0 here means some
                                    # value was deliberately left un-redacted in that file.
        groups = {}                 # dt.ID -> {"doc_type", "is_spreadsheet", "matches", "key_rows", "file_pairs"}

        def group_for(dt):
            g = groups.get(dt.ID)
            if g is None:
                g = {
                    "doc_type": dt,
                    "is_spreadsheet": self._is_spreadsheet_type(dt),
                    "matches": [],
                    "key_rows": [],
                    "file_pairs": [],
                }
                groups[dt.ID] = g
            return g

        def report(i, total, name, text):
            logger.line(f"[{i}/{total}] {name}: {text}")
            self.after(0, lambda: self._batch_log(f"[{i}/{total}] {name}: {text}"))
            self.after(0, lambda: self.batch_progress.configure(value=i))

        try:
            for i, f in enumerate(files, start=1):
                input_path = f["path"]
                dt = f["doc_type"]
                is_spreadsheet = self._is_spreadsheet_type(dt)
                name = os.path.basename(input_path)
                base, ext = os.path.splitext(name)
                output_path = os.path.join(output_dir, f"{base}_REDACTED{ext}")
                self.after(
                    0,
                    lambda i=i, total=len(files), name=name: self.batch_summary_var.set(
                        f"Redacting {i} of {total}: {name}..."
                    ),
                )
                group = group_for(dt)

                if is_spreadsheet:
                    # Cross-type identity linking: for every column
                    # manually linked to a pdf_report type's identity,
                    # resolve every distinct value THIS FILE will need
                    # against the shared identity pool before redacting
                    # it -- reusing an existing number for an identity
                    # already known (from an earlier PDF or spreadsheet
                    # file in this run), or allocating the next one for a
                    # never-seen identity. The result is seeded straight
                    # into this column's own `state` entry, in the exact
                    # {literal_value: placeholder_text} shape
                    # apply_column_redaction() already expects -- no
                    # engine changes needed for this part.
                    for col in selected_columns:
                        shared_label = col.get("shared_identity_label")
                        if not shared_label:
                            continue
                        identity_key = (_BATCH_IDENTITY_PREFIX, shared_label)
                        pool = state.setdefault(identity_key, {})
                        normalize_fn = normalize_fns.get(shared_label, _default_normalize_identity)
                        distinct_values = dt.distinct_values_in_column(input_path, col["name"])
                        seed = state.setdefault(col["name"], {})
                        # Sorted, not raw set-iteration order -- see the
                        # matching comment in _redact_worker_spreadsheet()
                        # (single-file Redact tab): a plain set's iteration
                        # order is randomized per process launch, so looping
                        # over it as-is would assign a new identity a
                        # different number on separate runs even for
                        # byte-identical input, silently breaking
                        # un-redaction if a file and key CSV from different
                        # runs are ever paired together.
                        for literal_value in sorted(distinct_values, key=normalize_fn):
                            norm_key = normalize_fn(literal_value)
                            if norm_key not in pool:
                                pool[norm_key] = len(pool) + 1
                            seed[literal_value] = f"<{shared_label} {pool[norm_key]}>"

                    summary, key_rows = dt.apply_column_redaction(
                        input_path, output_path, selected_columns, state=state,
                    )
                    result = dt.verify_column_redaction(input_path, output_path, selected_columns, key_rows)
                    review_note = ""
                    if not result["ok"]:
                        can_review = (
                            bool(result.get("leaks"))
                            and not result.get("shape_issue")
                            and not result.get("other_issue")
                        )
                        if not can_review:
                            # A shape/structure problem, not a per-cell
                            # value -- no manual review can fix this, so
                            # it still takes down the whole batch (see
                            # this section's header comment).
                            raise _BatchFileFailure(input_path, output_path, result["issues"])

                        report(i, len(files), name, "needs manual review -- waiting for your decision...")
                        review_outcome = self._batch_run_leak_review(
                            name, i, len(files), input_path, output_path, dt, selected_columns, result, key_rows
                        )
                        if review_outcome.get("skipped"):
                            # Only THIS file is skipped -- its output was
                            # already deleted by the review handler. The
                            # identities apply_column_redaction() already
                            # assigned for it are still folded into the
                            # shared state below, so numbering for the
                            # REST of the batch stays consistent whether
                            # or not this file's own output survived.
                            for row in key_rows:
                                state.setdefault(row["column"], {})[row["value"]] = row["placeholder"]
                            skipped_files.append((name, result["issues"]))
                            report(i, len(files), name, "SKIPPED -- manual review did not resolve this file")
                            continue
                        review_note = (
                            f"; {len(review_outcome['replaced'])} replaced / "
                            f"{len(review_outcome['ignored'])} left as-is via manual review"
                        )
                        reviewed_files.append(
                            (name, len(review_outcome["replaced"]), len(review_outcome["ignored"]))
                        )

                    for row in key_rows:
                        state.setdefault(row["column"], {})[row["value"]] = row["placeholder"]
                    group["key_rows"].extend(key_rows)
                    not_found_note = (
                        f" (not found in this file: {', '.join(summary['columns_not_found'])})"
                        if summary["columns_not_found"] else ""
                    )
                    per_file_summaries.append(
                        f"{name}: {summary['cells_redacted']} cell(s) redacted across "
                        f"{summary['columns_found']} column(s){not_found_note}{review_note}"
                    )
                    report(
                        i, len(files), name,
                        f"OK -- {summary['cells_redacted']} cell(s) redacted"
                        + (" (after manual review)" if review_note else ""),
                    )
                else:
                    identity_key = (_BATCH_IDENTITY_PREFIX, dt.ENTITY_LABEL)
                    existing_map = state.get(identity_key)
                    matches, unmatched, page_count = dt.scan(input_path)
                    mapping = dt.assign_placeholder_numbers(matches, existing_map=existing_map)
                    state[identity_key] = mapping
                    dt.apply_redactions(input_path, output_path, matches)
                    result = dt.verify_redaction(input_path, output_path, matches)
                    leak_review_note = ""
                    if not result["ok"]:
                        can_review = (
                            hasattr(dt, "apply_pdf_leak_decisions")
                            and bool(result.get("leaks"))
                            and not result.get("structural_issue")
                        )
                        if not can_review:
                            # A structural problem, not a per-occurrence
                            # name leak -- no manual review can fix this,
                            # so it still takes down the whole batch (see
                            # this section's header comment).
                            raise _BatchFileFailure(input_path, output_path, result["issues"])

                        report(i, len(files), name, "needs manual review -- waiting for your decision...")
                        review_outcome = self._batch_run_pdf_leak_review(
                            name, i, len(files), input_path, output_path, dt, matches, result
                        )
                        if review_outcome.get("skipped"):
                            # Only THIS file is skipped -- its output was
                            # already deleted by the review handler. The
                            # identity numbering this file was assigned
                            # (state[identity_key], set above) is already
                            # folded into the shared state regardless of
                            # whether this file's own output survived, so
                            # the REST of the batch stays numbered
                            # consistently -- same reasoning as the
                            # spreadsheet branch's key_rows carry-forward,
                            # just already done for PDF by the time we get
                            # here.
                            skipped_files.append((name, result["issues"]))
                            report(i, len(files), name, "SKIPPED -- manual review did not resolve this file")
                            continue
                        leak_review_note = (
                            f"; {len(review_outcome['replaced'])} redacted / "
                            f"{len(review_outcome['ignored'])} left as-is via manual review"
                        )
                        reviewed_files.append(
                            (name, len(review_outcome["replaced"]), len(review_outcome["ignored"]))
                        )
                        # See the single-file tab's _apply_pdf_leak_review_
                        # decisions for why: a 'redact' decision on a leak
                        # scan_pdf() never recognized as a name in the
                        # first place (find_unresolved_name_positions())
                        # got a brand-new placeholder number with no
                        # existing entry in `matches` -- folding it in here
                        # is what gets it into this file's occurrence/
                        # unique counts below AND into group["matches"],
                        # so the batch's combined key file covers it too.
                        matches = matches + review_outcome.get("new_matches", [])

                    group["matches"].extend(matches)
                    if unmatched:
                        unmatched_by_file.append((name, unmatched))
                    n_unique = len({m.placeholder_no for m in matches})
                    review_note = f"; {len(unmatched)} line(s) need manual review" if unmatched else ""
                    per_file_summaries.append(
                        f"{name}: {len(matches)} occurrence(s) redacted across {n_unique} unique "
                        f"{dt.ENTITY_LABEL.lower()}(s){review_note}{leak_review_note}"
                    )
                    review_note2 = f", {len(unmatched)} need review" if unmatched else ""
                    report(
                        i, len(files), name,
                        f"OK -- {len(matches)} occurrence(s) redacted{review_note2}"
                        + (" (after manual review)" if leak_review_note else ""),
                    )

                group["file_pairs"].append((name, os.path.basename(output_path)))
                produced_outputs.append(output_path)

        except _BatchFileFailure as bf:
            logger.line(f"FAILED on {os.path.basename(bf.source_path)}: {'; '.join(bf.issues)}")
            # All-or-nothing across the WHOLE batch, not just the failing
            # file: every output already written this run shares numbering
            # state (and, for linked columns, shared identity pools) with
            # the one that failed, so none of them can be trusted as a
            # self-consistent set on their own. Same reasoning as the
            # single-file tabs' self-check handling, applied at the batch
            # level.
            for p in produced_outputs + [bf.output_path]:
                if p and os.path.exists(p):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
            log_path = write_selfcheck_log("batch_redact", bf.source_path, bf.output_path, bf.issues, near_path=bf.output_path)
            # Default-argument capture required -- see _scan_worker()'s
            # comment: Python deletes an `except ... as bf` binding the
            # instant this except suite exits (via an implicit finally),
            # which happens as soon as the `return` below runs -- well
            # before Tk actually invokes this scheduled callback. A bare
            # `lambda: ...` closing over `bf` would raise NameError the
            # moment Tk calls it, silently, inside Tk's callback
            # machinery.
            self.after(0, lambda bf=bf, log_path=log_path: self._batch_redact_selfcheck_failed(bf, log_path))
            return
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            logger.line(f"FAILED (exception): {type(exc).__name__}: {exc}")
            for p in produced_outputs:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
            self.after(0, lambda exc=exc, tb=tb: self._batch_redact_error(exc, tb))
            return

        key_written = []
        key_errors = {}
        for dt_id, group in groups.items():
            key_path = key_paths.get(dt_id)
            if not key_path:
                continue
            try:
                if group["is_spreadsheet"]:
                    # First-seen wins on a (column, placeholder) collision
                    # -- the same convention the PDF side's write_name_key
                    # already uses (by_patient.setdefault) when the same
                    # normalized identity recurs with a slightly different
                    # literal spelling (e.g. via a cross-type link).
                    deduped = {}
                    for row in group["key_rows"]:
                        k = (row["column"], row["placeholder"])
                        if k not in deduped:
                            deduped[k] = row
                    group["doc_type"].write_key_batch(key_path, group["file_pairs"], list(deduped.values()))
                else:
                    group["doc_type"].write_key_batch(key_path, group["file_pairs"], group["matches"])
                key_written.append(key_path)
                logger.line(f"Combined key CSV ({group['doc_type'].DISPLAY_NAME}): written")
            except Exception as exc:  # noqa: BLE001
                key_errors[group["doc_type"].DISPLAY_NAME] = str(exc)
                logger.line(f"Combined key CSV ({group['doc_type'].DISPLAY_NAME}): FAILED -- {exc}")

        review_log_path = None
        if unmatched_by_file:
            try:
                review_log_path = write_batch_needs_review_log(output_dir, unmatched_by_file)
            except OSError:
                review_log_path = None

        logger.line(f"Done -- total {time.time() - logger.run_start:.1f}s")
        if skipped_files:
            logger.line(
                f"{len(skipped_files)} file(s) skipped (manual review cancelled or unresolved): "
                + ", ".join(n for n, _ in skipped_files)
            )
        n_files = sum(len(g["file_pairs"]) for g in groups.values())
        self.after(
            0,
            lambda: self._batch_redact_done(
                per_file_summaries, n_files, output_dir, key_written, key_errors,
                unmatched_by_file, review_log_path, skipped_files, reviewed_files,
            ),
        )

    def _batch_redact_reset_buttons(self):
        self.batch_progress.stop()
        self.batch_progress.pack_forget()
        if any(self._is_spreadsheet_type(f["doc_type"]) for f in self.batch_files):
            self.batch_load_columns_btn.configure(state="normal")
        self._update_batch_redact_btn_state()

    def _batch_redact_selfcheck_failed(self, bf, log_path):
        self._batch_redact_reset_buttons()
        issues = "\n".join(f"  - {i}" for i in bf.issues)
        log_note = f"\n\nThis was also saved to:\n{log_path}" if log_path else ""
        messagebox.showerror(
            APP_TITLE,
            f'The self-check on "{os.path.basename(bf.source_path)}" found problems, so '
            "NOTHING from this batch run was saved -- every file produced so far this run "
            f"was deleted, since they share numbering (all-or-nothing):\n\n{issues}\n\n"
            "Try redacting that file individually on the Redact tab first -- it offers "
            "full manual review -- and once it's clean, include it in a new batch run.\n\n"
            f"Please also report this to the developer with the source file.{log_note}",
        )
        self.batch_summary_var.set("Batch failed -- nothing was saved. See the message above.")
        self._batch_log("FAILED -- nothing from this run was saved. See the message above.")

    def _batch_redact_error(self, exc, tb):
        self._batch_redact_reset_buttons()
        log_path = write_error_log("batch_redact", "(multiple files)", "(multiple files)", tb)
        log_note = f"\n\nFull error details were saved to:\n{log_path}" if log_path else ""
        messagebox.showerror(
            APP_TITLE,
            f"Batch redaction failed and nothing from this run was saved:\n\n"
            f"{type(exc).__name__}: {exc}\n\n"
            f"Please do not use any partially-written output file.\n\n{FAILURE_TROUBLESHOOTING}{log_note}",
        )
        self.batch_summary_var.set("Batch failed -- nothing was saved. See the message above.")
        self._batch_log("FAILED -- nothing from this run was saved. See the message above.")

    def _batch_redact_done(
        self, per_file_summaries, n_files, output_dir, key_written, key_errors, unmatched_by_file, review_log_path,
        skipped_files=None, reviewed_files=None,
    ):
        skipped_files = skipped_files or []
        reviewed_files = reviewed_files or []
        self._batch_redact_reset_buttons()
        lines = "\n".join(f"  - {s}" for s in per_file_summaries)
        any_ignored = any(ignored_count for _, _, ignored_count in reviewed_files)
        if any_ignored:
            selfcheck_line = (
                "The self-check confirmed every file above is clean EXCEPT where a manual "
                "review deliberately left something un-redacted (see \"left as-is\" counts "
                "above) -- do not treat those specific files as fully de-identified without "
                "checking the review log for exactly what was left and why."
            )
        else:
            selfcheck_line = "The self-check on every file above confirmed everything targeted is fully removed."
        msg = (
            f"Done. Redacted {n_files} file(s):\n\n{lines}\n\n"
            f"Saved to folder:\n{output_dir}\n\n{selfcheck_line}"
        )
        if skipped_files:
            listed = "\n".join(f"  - {n}" for n, _ in skipped_files)
            msg += (
                f"\n\n{len(skipped_files)} file(s) were SKIPPED -- nothing was saved for "
                f"them, because their manual review was cancelled or could not be resolved:"
                f"\n{listed}\n\n"
                "Try redacting each of those individually on the Redact tab first (it offers "
                "the same manual review), and once it's clean, include it in a new batch run."
            )
        if key_written:
            listed = "\n".join(f"  - {p}" for p in key_written)
            plural = "files" if len(key_written) > 1 else "file"
            multi_note = (
                " -- one per document type, since a PDF key and a spreadsheet key are "
                "different shapes; when you un-redact one of this batch's files later, pick "
                "whichever key file matches THAT file's format (the Un-redact tab also "
                "figures this out on its own if you're not sure -- see the key files' own "
                "\"Column, Placeholder, Real Value\" vs \"Placeholder, Real Name\" headers)"
                if len(key_written) > 1 else ""
            )
            msg += (
                f"\n\n{len(key_written)} combined key {plural} covering this batch{multi_note} "
                f"{'were' if len(key_written) > 1 else 'was'} saved:\n{listed}\n\n"
                "Keep these with your own case file and do not send them along with the "
                "redacted files. Use the Un-redact tab to reverse them later."
            )
        for display_name, err in key_errors.items():
            msg += (
                f"\n\nThe redacted files above are complete and verified, but the combined "
                f"key file for {display_name} could not be saved: {err}"
            )
        if unmatched_by_file:
            total_unmatched = sum(len(u) for _, u in unmatched_by_file)
            msg += (
                f"\n\n{total_unmatched} line(s) across {len(unmatched_by_file)} file(s) did "
                "not match the expected pattern and were NOT redacted -- if any of those "
                "actually contain something that should be redacted, it is still visible in "
                f"that file's output. Full list saved to:\n{review_log_path or '(could not be saved)'}"
            )
        messagebox.showinfo(APP_TITLE, msg)
        self.batch_summary_var.set(f"Done -- {n_files} file(s) saved to {output_dir}")
        self._batch_log("Batch complete.")


    # ================================================================
    # UN-REDACT TAB
    # ================================================================
    def _build_unredact_tab(self, root):
        pad = {"padx": 10, "pady": 6}

        intro = ttk.Label(
            root,
            text="Reverse a previous redaction: pick a file and the key CSV(s) that go with "
                 "it, and get back a copy with every placeholder swapped for the real value "
                 "it replaced. Works for a redacted PDF, a redacted Excel workbook, or a "
                 "redacted workbook re-saved as a CSV -- which this is gets detected "
                 "automatically from the file you pick -- regardless of which document type "
                 "originally produced it. The file doesn't have to be the exact one that was "
                 "redacted, either: an Excel/CSV file can be ANY file that quotes the same "
                 "<Label N> placeholders (a downstream analysis or summary workbook, for "
                 "instance) -- pick its key CSV the same way, and every placeholder it "
                 "contains gets swapped back, wherever it appears in a cell. If the file was "
                 "redacted in more than one pass (each producing its own key CSV), pick all "
                 "of those key files together -- they'll be combined automatically.",
            wraplength=760, justify="left",
        )
        intro.pack(anchor="w", **pad)

        pdf_row = ttk.Frame(root)
        pdf_row.pack(fill="x", **pad)
        ttk.Label(pdf_row, text="Redacted file:", width=14, anchor="w").pack(side="left")
        self.unredact_pdf_var = tk.StringVar(value="(no file selected)")
        ttk.Label(pdf_row, textvariable=self.unredact_pdf_var, foreground="#333").pack(
            side="left", padx=8, fill="x", expand=True
        )
        ttk.Button(pdf_row, text="Browse...", command=self.on_browse_unredact_pdf).pack(side="left")

        key_row = ttk.Frame(root)
        key_row.pack(fill="x", **pad)
        ttk.Label(key_row, text="Key CSV(s):", width=14, anchor="w").pack(side="left")
        self.unredact_key_var = tk.StringVar(value="(no files selected)")
        ttk.Label(key_row, textvariable=self.unredact_key_var, foreground="#333").pack(
            side="left", padx=8, fill="x", expand=True
        )
        ttk.Button(key_row, text="Browse...", command=self.on_browse_unredact_key).pack(side="left")

        self.unredact_summary_var = tk.StringVar(value="Select both files to begin.")
        ttk.Label(root, textvariable=self.unredact_summary_var, font=("", 10, "bold")).pack(
            anchor="w", padx=10, pady=(10, 0)
        )

        self.unredact_progress = ttk.Progressbar(root, mode="indeterminate")

        warning = ttk.Label(
            root,
            text="The output file below is exactly as sensitive as the ORIGINAL, "
                 "unredacted document -- handle and store it accordingly.",
            foreground="#a15c00", wraplength=760, justify="left",
        )
        warning.pack(anchor="w", padx=10, pady=(20, 0))

        bottom = ttk.Frame(root)
        bottom.pack(fill="x", padx=10, pady=10, side="bottom")
        self.unredact_btn = ttk.Button(
            bottom, text="Un-redact && Save...", command=self.on_unredact_save, state="disabled"
        )
        self.unredact_btn.pack(side="right")

    def on_browse_unredact_pdf(self):
        # Named for history (this used to only ever be a PDF) -- now
        # accepts either a redacted PDF or a redacted Excel workbook; see
        # _unredact_engine_for().
        path = filedialog.askopenfilename(
            title="Select the redacted file",
            filetypes=[
                ("Redacted files", "*.pdf *.xlsx"),
                ("PDF files", "*.pdf"),
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self.unredact_pdf_path = path
        self.unredact_pdf_var.set(path)
        self._update_unredact_btn_state()

    def on_browse_unredact_key(self):
        # Multi-select: a file may have been redacted in more than one
        # pass, each producing its own key CSV -- see _build_unredact_tab's
        # intro text and on_unredact_save()'s merge step below.
        paths = filedialog.askopenfilenames(
            title="Select the key CSV file(s) -- select more than one if this file was redacted in multiple passes",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not paths:
            return
        self.unredact_key_paths = list(paths)
        if len(self.unredact_key_paths) == 1:
            self.unredact_key_var.set(self.unredact_key_paths[0])
        else:
            names = ", ".join(os.path.basename(p) for p in self.unredact_key_paths)
            self.unredact_key_var.set(f"{len(self.unredact_key_paths)} files selected: {names}")
        self._update_unredact_btn_state()

    def _update_unredact_btn_state(self):
        if self.unredact_pdf_path and self.unredact_key_paths:
            self.unredact_btn.configure(state="normal")
            self.unredact_summary_var.set("Ready. Click Un-redact & Save.")

    @staticmethod
    def _unredact_engine_for(path):
        """Picks the right un-redaction engine from the redacted file's
        extension -- .xlsx or .csv means a workbook/exported-workbook
        (spreadsheet_unredact_engine handles both; a redacted workbook is
        often re-saved as a CSV by other software somewhere between
        redaction and un-redaction), anything else is treated as the
        original PDF case (unredact_engine). Returns (engine_module, kind)
        where kind is "spreadsheet" or "pdf", used only for wording the
        result message and choosing the output file's extension.
        """
        ext = os.path.splitext(path)[1].lower()
        if ext in (".xlsx", ".csv"):
            return spreadsheet_unredact_engine, "spreadsheet"
        return unredact_engine, "pdf"

    def on_unredact_save(self):
        engine, kind = self._unredact_engine_for(self.unredact_pdf_path)

        mappings = []
        for key_path in self.unredact_key_paths:
            try:
                mappings.append(engine.read_key_csv(key_path))
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror(
                    APP_TITLE, f"Could not read this key CSV:\n\n{key_path}\n\n{exc}"
                )
                return

        # merge_key_mappings() lets several key files (from separate
        # redaction passes over the same file) be used together -- see
        # its docstring in the engine module. A conflict there (the same
        # placeholder meaning two different things across the files
        # selected) is not something to guess through: it almost always
        # means one of the files doesn't actually belong with the others,
        # so this stops before touching anything rather than picking one
        # value over the other.
        key_mapping, conflicts = engine.merge_key_mappings(mappings)
        if conflicts:
            lines = []
            for c in conflicts[:10]:
                versus = " vs. ".join(c["values"])
                # A spreadsheet-engine conflict's "column" is either a real
                # header string, or None for an entry that came from a
                # generic/flat key (see spreadsheet_unredact_engine.py's "A
                # GENERIC/FLAT KEY" docstring section) -- worded without
                # "Column" for that case since there isn't one.
                if "column" in c and c["column"] is not None:
                    lines.append(f"  - Column '{c['column']}', placeholder {c['placeholder']}: {versus}")
                else:
                    lines.append(f"  - Placeholder {c['placeholder']}: {versus}")
            more = f"\n  ... and {len(conflicts) - 10} more" if len(conflicts) > 10 else ""
            messagebox.showerror(
                APP_TITLE,
                "The key CSV files you selected disagree with each other -- the same "
                "placeholder maps to a different real value in different files:\n\n"
                + "\n".join(lines) + more +
                "\n\nThis almost always means one of these key files doesn't actually "
                "belong with the others (e.g. the wrong file was added by mistake, or it's "
                "a key file from a different document entirely). Nothing was un-redacted -- "
                "fix the file selection above and try again.",
            )
            return

        base = os.path.splitext(os.path.basename(self.unredact_pdf_path))[0]
        if kind == "spreadsheet":
            # Preserve whatever format the file we were HANDED is in --
            # .xlsx stays .xlsx, .csv stays .csv -- rather than always
            # forcing .xlsx. If the redacted file already drifted to CSV
            # since it was redacted, there's no reason to force it back
            # into a workbook the un-redaction step didn't ask for.
            out_ext = os.path.splitext(self.unredact_pdf_path)[1].lower()
            save_filetypes = [("CSV files", "*.csv")] if out_ext == ".csv" else [("Excel files", "*.xlsx")]
        else:
            out_ext = ".pdf"
            save_filetypes = [("PDF files", "*.pdf")]
        default_name = f"{base}_UNREDACTED{out_ext}"
        default_dir = os.path.dirname(self.unredact_pdf_path)
        output_path = filedialog.asksaveasfilename(
            title="Save un-redacted file as...",
            initialdir=default_dir,
            initialfile=default_name,
            defaultextension=out_ext,
            filetypes=save_filetypes,
        )
        if not output_path:
            return

        lock_exc = check_file_not_locked(output_path)
        if lock_exc is not None:
            proceed = messagebox.askyesno(
                APP_TITLE, LOCKED_FILE_MESSAGE.format(path=output_path, exc=lock_exc), icon="warning"
            )
            if not proceed:
                return

        proceed = messagebox.askyesno(
            APP_TITLE,
            f"This will produce a new file at:\n\n{output_path}\n\n"
            "containing the real, unredacted values -- exactly as sensitive as the "
            "original source document. Continue?",
            icon="warning",
        )
        if not proceed:
            return

        self.unredact_btn.configure(state="disabled")
        self.unredact_summary_var.set("Un-redacting...")
        self.unredact_progress.configure(mode="determinate", maximum=100, value=0)
        self.unredact_progress.pack(fill="x", padx=10, pady=(0, 4))
        threading.Thread(
            target=self._unredact_worker, args=(engine, kind, key_mapping, output_path), daemon=True
        ).start()

    def _on_unredact_progress(self, current, total):
        pct = int((current / total) * 100) if total else 0
        unit = getattr(self, "_unredact_progress_unit", "page")
        self.unredact_summary_var.set(f"Un-redacting -- {unit} {current} of {total}...")
        self.unredact_progress.configure(value=pct)
        self._set_title(f"{pct}% un-redacting")

    def _on_unredact_saving_phase(self):
        self.unredact_summary_var.set("Saving to disk -- finalizing the file...")
        self.unredact_progress.configure(mode="indeterminate")
        self.unredact_progress.start(12)
        self._set_title("saving...")

    def _unredact_worker(self, engine, kind, key_mapping, output_path):
        logger = RunLogger(
            "unredact",
            f"starting, {sum(len(v) for v in key_mapping.values()) if kind == 'spreadsheet' else len(key_mapping)} "
            f"placeholder(s) known, output: {output_path}",
        )

        self._unredact_progress_unit = "row" if kind == "spreadsheet" else "page"
        progress_cb = self._make_instrumented_callback(
            self.after, self._on_unredact_progress, "Un-redacting", logger,
            unit_label=self._unredact_progress_unit,
        )

        def phase_cb(phase):
            logger.line("Un-redacting: all done -- now writing final file to disk")
            self.after(0, self._on_unredact_saving_phase)

        try:
            result = engine.apply_unredaction(
                self.unredact_pdf_path, output_path, key_mapping,
                progress_callback=progress_cb, phase_callback=phase_cb,
            )
            units_changed = result.get("pages_changed", result.get("sheets_changed", 0))
            logger.line(
                f"Saving to disk: finished -- {result['replacements']} replacement(s) "
                f"across {units_changed} {self._unredact_progress_unit.replace('row', 'sheet')}(s)"
            )
            verify_result = engine.verify_unredaction(output_path, key_mapping)
            logger.line(f"Verifying: finished -- ok={verify_result['ok']}")
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            logger.line(f"FAILED: {type(exc).__name__}: {exc}")
            # Default-argument capture required -- see _scan_worker()'s
            # comment.
            self.after(0, lambda exc=exc, tb=tb: self._unredact_failed(exc, tb, output_path))
            return

        logger.line(f"Done -- total {time.time() - logger.run_start:.1f}s")
        self.after(0, lambda: self._unredact_done(verify_result, result, output_path, kind, engine, key_mapping))

    def _unredact_failed(self, exc, tb, output_path):
        self._set_title()
        self.unredact_progress.stop()
        self.unredact_progress.pack_forget()
        self.unredact_btn.configure(state="normal")
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
        log_path = write_error_log("unredact", self.unredact_pdf_path, output_path, tb)
        log_note = f"\n\nFull error details were saved to:\n{log_path}" if log_path else ""
        messagebox.showerror(
            APP_TITLE,
            f"Un-redaction failed and no file was saved:\n\n{type(exc).__name__}: {exc}\n\n"
            f"{FAILURE_TROUBLESHOOTING}{log_note}",
        )

    @staticmethod
    def _leftover_matches_ambiguous(leftover, ambiguous):
        """True when every remaining leftover-placeholder location the
        self-check found is explained by an already-flagged ambiguous
        occurrence (see spreadsheet_unredact_engine.py's apply_unredaction()
        'ambiguous' field) -- i.e. nothing UNEXPECTED is still wrong, just
        the cases that genuinely need a human pick. False (including when
        there are no ambiguous entries at all) means something else is
        going on that a manual pick-the-right-value screen can't explain,
        so main_gui.py falls back to the original hard block instead of
        offering a review that wouldn't actually resolve everything.
        """
        ambiguous_locs = {(a["sheet"], a["coordinate"]) for a in ambiguous}
        if not ambiguous_locs:
            return False
        return all((e["sheet"], e["coordinate"]) in ambiguous_locs for e in leftover)

    def _unredact_done(self, verify_result, apply_result, output_path, kind, engine, key_mapping):
        self._set_title()
        self.unredact_progress.stop()
        self.unredact_progress.pack_forget()
        self.unredact_btn.configure(state="normal")

        if not verify_result["ok"]:
            can_review = kind == "spreadsheet" and self._leftover_matches_ambiguous(
                verify_result.get("leftover", []), apply_result.get("ambiguous") or []
            )
            if can_review:
                self._open_ambiguous_review(apply_result, output_path, kind, engine, key_mapping)
                return
            # Fall back to the original hard block -- either this isn't a
            # spreadsheet (the PDF engine has no per-occurrence structure
            # to offer a review over), there was nothing ambiguous to
            # begin with, or something UNEXPECTED beyond the flagged
            # ambiguous occurrences is still wrong (e.g. a genuinely
            # mismatched key file) -- a manual pick-the-right-value screen
            # has no way to fix that.
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            issues = "\n".join(f"  - {i}" for i in verify_result["issues"])

            # Full, untruncated detail for the log file -- verify_result's
            # 'issues' text above is capped at 10 examples per problem type
            # so the on-screen dialog stays readable, but a failure with
            # many more than that needs the complete list somewhere. Only
            # placeholder text and cell locations go in here, never a real
            # value, so this log is safe to read or forward regardless of
            # how sensitive the document itself is. Handles both engines'
            # slightly different 'leftover' entry shapes (see
            # unredact_engine.py's and spreadsheet_unredact_engine.py's own
            # verify_unredaction() docstrings).
            leftover = verify_result.get("leftover") or []
            detail_lines = []
            for entry in leftover:
                status = "known to a loaded key, but not matched" if entry.get("known") else "not in any loaded key at all"
                if "sheet" in entry:
                    detail_lines.append(f"{entry['sheet']}!{entry['coordinate']}: {entry['value']} ({status})")
                else:
                    detail_lines.append(f"{entry['placeholder']} ({status})")
            known_count = sum(1 for e in leftover if e.get("known"))
            unknown_count = len(leftover) - known_count
            count_note = ""
            if leftover:
                count_note = (
                    f"\n\n({known_count} matched a loaded key file but the value wasn't found where "
                    f"expected; {unknown_count} weren't recognized by any loaded key file at all.)"
                )

            log_path = write_selfcheck_log(
                "unredact", self.unredact_pdf_path, output_path, verify_result["issues"],
                extra_detail_lines=detail_lines,
            )
            log_note = f"\n\nThe complete, un-truncated list was saved to:\n{log_path}" if log_path else ""
            messagebox.showerror(
                APP_TITLE,
                "The self-check on the un-redacted file found problems, so the file was "
                "NOT saved:\n\n" + issues + count_note + log_note,
            )
            return

        self._finish_unredact_success(apply_result, output_path, kind)

    def _finish_unredact_success(self, apply_result, output_path, kind, resolved=None, review_log_path=None):
        resolved = resolved or []
        if kind == "spreadsheet":
            units_changed = apply_result["sheets_changed"]
            unit_word = "sheet" if units_changed == 1 else "sheets"
            file_word = "workbook"
        else:
            units_changed = apply_result["pages_changed"]
            unit_word = "page" if units_changed == 1 else "pages"
            file_word = "PDF"

        msg = (
            f"Done. Replaced {apply_result['replacements']} placeholder occurrence(s) "
            f"across {units_changed} {unit_word}.\n\nSaved to:\n{output_path}\n\n"
            "Remember: this file is exactly as sensitive as the original, unredacted "
            "document."
        )
        if resolved:
            msg += (
                f"\n\n{len(resolved)} of those had reused placeholder text and needed a "
                "manual pick during review to know which real value applied -- see the "
                f"review log for exactly which:\n{review_log_path or '(could not be saved)'}"
            )
        not_found = apply_result.get("placeholders_not_found") or []
        if not_found:
            shown = ", ".join(not_found[:10])
            more = f" (+{len(not_found) - 10} more)" if len(not_found) > 10 else ""
            msg += (
                f"\n\nNote: {len(not_found)} placeholder(s) listed in the key file were "
                f"never found in the {file_word} ({shown}{more}) -- double check this key "
                f"file actually belongs to this {file_word}."
            )
        messagebox.showinfo(APP_TITLE, msg)
        self.unredact_summary_var.set(f"Saved: {output_path}")

    # ----------------------------------------------------------------
    # Manual review for un-redaction: shown instead of the hard block when
    # every remaining self-check issue is a placeholder whose text was
    # reused across more than one column (see spreadsheet_unredact_engine.py's
    # module docstring, "HANDLING A FILE THAT'S DRIFTED SINCE IT WAS
    # REDACTED") and the header match that would normally disambiguate it
    # is gone (renamed/reordered/moved column). The tool genuinely cannot
    # tell which real value belongs at each such cell on its own -- this
    # lets the user pick, occurrence by occurrence.
    # ----------------------------------------------------------------
    def _open_ambiguous_review(self, apply_result, output_path, kind, engine, key_mapping):
        win = tk.Toplevel(self)
        win.title(f"{APP_TITLE} -- Manual Review")
        win.geometry("820x560")
        win.transient(self)
        win.grab_set()

        intro = (
            "The same placeholder text was used for more than one column when this file "
            "was redacted, and the column headers that would normally say which one "
            "applies to each cell below no longer match (renamed, reordered, or the file's "
            "layout otherwise changed since redaction). Nothing has been changed yet -- "
            "pick the correct real value for each occurrence below, or ignore it and leave "
            "the placeholder in place, then click Apply."
        )
        tk.Label(win, text=intro, wraplength=780, justify="left", anchor="w").pack(
            fill="x", padx=12, pady=(12, 6)
        )

        outer, inner = self._build_scrollable_frame(win)
        outer.pack(fill="both", expand=True, padx=12, pady=6)

        row_vars = []
        for occ in apply_result["ambiguous"]:
            frame = tk.Frame(inner, relief="groove", borderwidth=1)
            frame.pack(fill="x", pady=4, padx=2)

            loc_text = f"{occ['sheet']}!{occ['coordinate']}   placeholder: {occ['value']}"
            if occ.get("column_header"):
                loc_text += f"   -- in column '{occ['column_header']}'"
            tk.Label(frame, text=loc_text, font=("TkDefaultFont", 9, "bold"), anchor="w").pack(
                fill="x", padx=8, pady=(6, 0)
            )

            row_context = occ.get("row_context") or []
            if row_context:
                context_text = "Same row also has: " + "; ".join(
                    f"{c['label']}: {c['value']}" for c in row_context
                )
                tk.Label(
                    frame, text=context_text, wraplength=760, justify="left", anchor="w",
                    fg="#444444",
                ).pack(fill="x", padx=8, pady=(0, 0))

            choices = [f"{c['value']}  (from column '{c['column']}')" for c in occ["candidates"]]
            value_by_choice = {c_text: c["value"] for c_text, c in zip(choices, occ["candidates"])}
            choice_var = tk.StringVar(value=choices[0])
            action_var = tk.StringVar(value="replace")

            action_row = tk.Frame(frame)
            action_row.pack(fill="x", padx=8, pady=(2, 6))
            tk.Radiobutton(action_row, text="Replace with:", variable=action_var, value="replace").pack(side="left")
            combo = ttk.Combobox(action_row, textvariable=choice_var, values=choices, state="readonly", width=44)
            combo.pack(side="left", padx=(4, 12))
            tk.Radiobutton(
                action_row, text="Ignore (leave placeholder as-is)", variable=action_var, value="ignore"
            ).pack(side="left")

            row_vars.append({
                "occ": occ, "action_var": action_var, "choice_var": choice_var, "value_by_choice": value_by_choice,
            })

        btn_row = tk.Frame(win)
        btn_row.pack(fill="x", padx=12, pady=(6, 12))

        def on_cancel():
            win.destroy()
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            self.unredact_btn.configure(state="normal")
            messagebox.showinfo(
                APP_TITLE, "Manual review cancelled -- no file was saved. Nothing was changed."
            )

        def on_apply():
            decisions = []
            for rv in row_vars:
                occ = rv["occ"]
                action = rv["action_var"].get()
                entry = {"sheet": occ["sheet"], "coordinate": occ["coordinate"], "action": action}
                if action == "replace":
                    entry["value"] = rv["value_by_choice"][rv["choice_var"].get()]
                else:
                    # For the review log: what's still sitting in this
                    # cell is the placeholder text itself (nothing was
                    # written), not one of the candidate real values.
                    entry["value"] = occ["value"]
                decisions.append(entry)
            win.destroy()
            self._apply_ambiguous_decisions(decisions, apply_result, output_path, kind, engine, key_mapping)

        tk.Button(btn_row, text="Cancel (discard this run)", command=on_cancel).pack(side="right")
        tk.Button(btn_row, text="Apply Decisions", command=on_apply).pack(side="right", padx=(0, 8))

    def _apply_ambiguous_decisions(self, decisions, apply_result, output_path, kind, engine, key_mapping):
        try:
            engine.apply_ambiguous_decisions(output_path, decisions)
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            log_path = write_error_log("unredact", self.unredact_pdf_path, output_path, tb)
            log_note = f"\n\nFull error details were saved to:\n{log_path}" if log_path else ""
            messagebox.showerror(
                APP_TITLE,
                f"Applying your manual review decisions failed:\n\n{type(exc).__name__}: {exc}\n\n"
                f"The file at\n{output_path}\nmay be partially modified -- do not send it out; "
                f"re-run un-redaction from scratch instead.{log_note}",
            )
            self.unredact_btn.configure(state="normal")
            return

        replaced = [d for d in decisions if d["action"] == "replace"]
        ignored = [d for d in decisions if d["action"] == "ignore"]
        review_log_path = write_manual_review_log("unredact", self.unredact_pdf_path, output_path, replaced, ignored)

        # Re-run the exact same self-check used the first time rather than
        # trusting the decisions were all applied cleanly.
        final_result = engine.verify_unredaction(output_path, key_mapping)
        ignored_keys = {(d["sheet"], d["coordinate"]) for d in ignored}
        unresolved = [
            e for e in final_result.get("leftover", [])
            if (e["sheet"], e["coordinate"]) not in ignored_keys
        ]

        self.unredact_btn.configure(state="normal")

        if unresolved:
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            remaining = "\n".join(f"  - {e['sheet']}!{e['coordinate']}: {e['value']}" for e in unresolved)
            messagebox.showerror(
                APP_TITLE,
                "After applying your decisions, the self-check still found unresolved "
                f"placeholder(s), so the file was NOT saved:\n\n{remaining}\n\n"
                "This shouldn't normally happen -- please report it to the developer with "
                f"the source file.\n\nReview log:\n{review_log_path or '(could not be saved)'}",
            )
            return

        # apply_result's counts only reflect the FIRST pass, before this
        # review resolved anything -- correct them so the success message
        # doesn't understate what actually ended up in the file (e.g.
        # claiming "0 replacements" on a file where the only occurrences
        # were exactly the ones just resolved here).
        updated_result = dict(apply_result)
        updated_result["replacements"] = apply_result["replacements"] + len(replaced)
        if replaced:
            # apply_result["sheets_changed"] has no sheet-name set to union
            # against (it's just a count), so this can't be made exact in
            # every case -- but it can't be an undercount either: at least
            # every distinct sheet touched by a manual replace was
            # changed, and the first pass's own count is never wrong about
            # sheets IT touched, so the larger of the two is always a
            # correct floor and usually the exact answer.
            updated_result["sheets_changed"] = max(
                apply_result["sheets_changed"], len({d["sheet"] for d in replaced})
            )

        self._finish_unredact_success(
            updated_result, output_path, kind, resolved=replaced, review_log_path=review_log_path
        )


def main():
    app = DocumentRedactorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
