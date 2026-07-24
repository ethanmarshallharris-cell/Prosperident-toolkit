"""
Prosperident Document Toolkit -- desktop GUI.

A central app for redacting identifying information out of different kinds
of report PDFs, and for reversing that redaction later given the name key
file saved alongside it. Each supported report layout is a separate
"document type" plugin (see document_types/base.py for the plugin
contract, document_types/dentrix_audit_trail.py for the first one) --
choosing a document type from the dropdown is what tells the app which
parsing/redaction logic to use; adding a new report layout later never
requires changing this file, only adding a new plugin module.

Two tabs:
  Redact      Pick a document type and source PDF, scan it, review what
              was found, then redact & save. The tool re-verifies its own
              output before reporting success -- if anything looks off, it
              refuses to hand back a file that might be wrong.
  Un-redact   Pick a previously redacted PDF and the *_NAME_KEY.csv saved
              alongside it, and get back a copy with every placeholder
              swapped for the real value it replaced. Works the same way
              regardless of which document type originally produced the
              file (see unredact_engine.py).

Built with only the Python standard library (tkinter) plus PyMuPDF, so it
packages cleanly with PyInstaller.
"""
import os
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
import update_check
import branding

import fitz  # PyMuPDF

APP_TITLE = "Prosperident Document Toolkit"
# Bumped whenever the build changes in a way that's worth being able to
# confirm at a glance -- e.g. "is this machine actually running the
# version with the per-page progress display, or an older build?" Shown
# in the window title bar so that question never requires guessing.
APP_BUILD = "2026-07-20"
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
        self.log_path = os.path.join(tempfile.gettempdir(), f"DocumentToolkit_{log_name}_log.txt")
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


def write_error_log(log_name, input_path, output_path, tb):
    log_path = os.path.join(tempfile.gettempdir(), f"DocumentToolkit_{log_name}_error_log.txt")
    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Input file: {input_path}\n")
            f.write(f"Intended output: {output_path}\n\n")
            f.write(tb)
        return log_path
    except OSError:
        return None


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


class DocumentToolkitApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  (build {APP_BUILD})")
        self.geometry("1080x720")
        self.minsize(880, 600)

        branding.apply_theme(self)
        branding.add_header(
            self, "Document Toolkit",
            "Redact & un-redact identifying information in report PDFs",
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

        self.unredact_pdf_path = None
        self.unredact_key_path = None

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
        threading.Thread(target=self._update_check_worker, daemon=True).start()

    def _update_check_worker(self):
        result = update_check.check_for_update("document_toolkit", APP_BUILD)
        if result:
            self.after(0, lambda: self._show_update_banner(result))

    def _show_update_banner(self, entry):
        notes = entry.get("notes") or ""
        text = f"A newer version is available (build {entry.get('version')}). Click here to get it."
        if notes:
            text += f"  — {notes}"
        link = ttk.Label(
            self, textvariable=None, text=text, foreground="#0645AD",
            cursor="hand2", font=("", 9, "underline"),
        )
        link.pack(side="bottom", fill="x", padx=10, pady=4)
        download_url = entry.get("download_url")
        if download_url:
            link.bind("<Button-1>", lambda e: webbrowser.open(download_url))

    # ================================================================
    # Top-level layout: two tabs sharing one window
    # ================================================================
    def _build_widgets(self):
        outer = ttk.Notebook(self)
        outer.pack(fill="both", expand=True)

        redact_tab = ttk.Frame(outer)
        outer.add(redact_tab, text="Redact")
        self._build_redact_tab(redact_tab)

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

        self.summary_var = tk.StringVar(value="Select a document type and file, then click Scan.")
        ttk.Label(root, textvariable=self.summary_var, font=("", 10, "bold")).pack(
            anchor="w", padx=10
        )

        self.progress = ttk.Progressbar(root, mode="indeterminate")

        main = ttk.PanedWindow(root, orient="horizontal")
        main.pack(fill="both", expand=True, padx=10, pady=6)

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
        self.summary_var.set(f"Select a {self.current_doc_type.DISPLAY_NAME} file, then click Scan.")
        self.review_hint_var.set(self.current_doc_type.REVIEW_HINT)
        if self.current_doc_type.SUPPORTS_KEY_FILE:
            self.save_key_check.pack(side="right", padx=(0, 16))
        else:
            self.save_key_check.pack_forget()

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

    def on_scan(self):
        if not self.input_path:
            return
        self.scan_btn.configure(state="disabled")
        self.redact_btn.configure(state="disabled")
        self._clear_results()
        self.summary_var.set("Scanning...")
        self.progress.configure(mode="indeterminate")
        self.progress.pack(fill="x", padx=10, pady=(0, 4))
        self.progress.start(12)
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        doc_type = self.current_doc_type
        try:
            matches, unmatched, page_count = doc_type.scan(self.input_path)
            mapping = doc_type.assign_placeholder_numbers(matches)
        except Exception as exc:  # noqa: BLE001
            self.after(0, lambda: self._scan_failed(exc))
            return
        self.after(0, lambda: self._scan_done(matches, unmatched, page_count, mapping))

    def _scan_failed(self, exc):
        self.progress.stop()
        self.progress.pack_forget()
        self.scan_btn.configure(state="normal")
        messagebox.showerror(
            APP_TITLE,
            f"Could not scan this file:\n\n{exc}\n\n"
            "If this keeps happening, the file's layout may differ from what "
            "this document type expects -- please check before relying on it "
            "for this file.",
        )

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
        self.summary_var.set(f"Redacting -- page {current} of {total}...")
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
        self.summary_var.set(f"Verifying -- page {current} of {total}...")
        self.progress.configure(value=pct)
        self._set_title(f"{pct}% verifying")

    @staticmethod
    def _make_instrumented_callback(gui_after, handler, label, logger):
        step = None
        seen_first = False

        def callback(current, total):
            nonlocal step, seen_first
            if not seen_first:
                seen_first = True
                logger.line(f"{label}: started ({total} page(s) to process)")
            if step is None:
                step = max(1, total // 100)
            if current == total:
                logger.line(f"{label}: finished ({current} of {total} pages)")
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
            self.after(0, lambda: self._redact_failed(exc, tb, output_path))
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
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            issues = "\n".join(f"  - {i}" for i in result["issues"])
            messagebox.showerror(
                APP_TITLE,
                "The self-check on the redacted file found problems, so the file was "
                "NOT saved -- do not deliver anything from this run:\n\n"
                f"{issues}\n\n"
                "Please report this to the developer with the source file.",
            )
            return

        n_unique = len({m.placeholder_no for m in included})
        entity_lower = doc_type.ENTITY_LABEL.lower()
        msg = (
            f"Done. Redacted {len(included)} {entity_lower} occurrence(s) across {n_unique} "
            f"unique {entity_lower}(s).\n\nSaved to:\n{output_path}\n\n"
            "The self-check confirmed everything targeted is fully removed and no other "
            "content was affected."
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
        timing_log_path = os.path.join(tempfile.gettempdir(), "DocumentToolkit_redact_log.txt")
        if os.path.exists(timing_log_path):
            msg += (
                f"\n\nA timing breakdown for this run was recorded to:\n{timing_log_path}\n"
                "If this took longer than expected, that file is the fastest way to show "
                "exactly where the time went."
            )
        messagebox.showinfo(APP_TITLE, msg)
        self.summary_var.set(f"Saved: {output_path}")

    # ================================================================
    # UN-REDACT TAB
    # ================================================================
    def _build_unredact_tab(self, root):
        pad = {"padx": 10, "pady": 6}

        intro = ttk.Label(
            root,
            text="Reverse a previous redaction: pick the redacted file and the name key "
                 "CSV that was saved alongside it, and get back a copy with every "
                 "placeholder swapped for the real value it replaced. Works regardless of "
                 "which document type originally produced the file.",
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
        ttk.Label(key_row, text="Name key CSV:", width=14, anchor="w").pack(side="left")
        self.unredact_key_var = tk.StringVar(value="(no file selected)")
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
        path = filedialog.askopenfilename(
            title="Select the redacted PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not path:
            return
        self.unredact_pdf_path = path
        self.unredact_pdf_var.set(path)
        self._update_unredact_btn_state()

    def on_browse_unredact_key(self):
        path = filedialog.askopenfilename(
            title="Select the name key CSV", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        self.unredact_key_path = path
        self.unredact_key_var.set(path)
        self._update_unredact_btn_state()

    def _update_unredact_btn_state(self):
        if self.unredact_pdf_path and self.unredact_key_path:
            self.unredact_btn.configure(state="normal")
            self.unredact_summary_var.set("Ready. Click Un-redact & Save.")

    def on_unredact_save(self):
        try:
            key_mapping = unredact_engine.read_key_csv(self.unredact_key_path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Could not read the name key CSV:\n\n{exc}")
            return

        base = os.path.splitext(os.path.basename(self.unredact_pdf_path))[0]
        default_name = f"{base}_UNREDACTED.pdf"
        default_dir = os.path.dirname(self.unredact_pdf_path)
        output_path = filedialog.asksaveasfilename(
            title="Save un-redacted PDF as...",
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
            target=self._unredact_worker, args=(key_mapping, output_path), daemon=True
        ).start()

    def _on_unredact_progress(self, current, total):
        pct = int((current / total) * 100) if total else 0
        self.unredact_summary_var.set(f"Un-redacting -- page {current} of {total}...")
        self.unredact_progress.configure(value=pct)
        self._set_title(f"{pct}% un-redacting")

    def _on_unredact_saving_phase(self):
        self.unredact_summary_var.set("Saving to disk -- finalizing the file...")
        self.unredact_progress.configure(mode="indeterminate")
        self.unredact_progress.start(12)
        self._set_title("saving...")

    def _unredact_worker(self, key_mapping, output_path):
        logger = RunLogger(
            "unredact",
            f"starting, {len(key_mapping)} placeholder(s) known, output: {output_path}",
        )

        progress_cb = self._make_instrumented_callback(
            self.after, self._on_unredact_progress, "Un-redacting", logger
        )

        def phase_cb(phase):
            logger.line("Un-redacting: all pages drawn -- now writing final file to disk")
            self.after(0, self._on_unredact_saving_phase)

        try:
            result = unredact_engine.apply_unredaction(
                self.unredact_pdf_path, output_path, key_mapping,
                progress_callback=progress_cb, phase_callback=phase_cb,
            )
            logger.line(
                f"Saving to disk: finished -- {result['replacements']} replacement(s) "
                f"across {result['pages_changed']} page(s)"
            )
            verify_result = unredact_engine.verify_unredaction(output_path, key_mapping)
            logger.line(f"Verifying: finished -- ok={verify_result['ok']}")
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            logger.line(f"FAILED: {type(exc).__name__}: {exc}")
            self.after(0, lambda: self._unredact_failed(exc, tb, output_path))
            return

        logger.line(f"Done -- total {time.time() - logger.run_start:.1f}s")
        self.after(0, lambda: self._unredact_done(verify_result, result, output_path))

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

    def _unredact_done(self, verify_result, apply_result, output_path):
        self._set_title()
        self.unredact_progress.stop()
        self.unredact_progress.pack_forget()
        self.unredact_btn.configure(state="normal")

        if not verify_result["ok"]:
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            issues = "\n".join(f"  - {i}" for i in verify_result["issues"])
            messagebox.showerror(
                APP_TITLE,
                "The self-check on the un-redacted file found problems, so the file was "
                "NOT saved:\n\n" + issues,
            )
            return

        msg = (
            f"Done. Replaced {apply_result['replacements']} placeholder occurrence(s) "
            f"across {apply_result['pages_changed']} page(s).\n\nSaved to:\n{output_path}\n\n"
            "Remember: this file is exactly as sensitive as the original, unredacted "
            "document."
        )
        not_found = apply_result.get("placeholders_not_found") or []
        if not_found:
            shown = ", ".join(not_found[:10])
            more = f" (+{len(not_found) - 10} more)" if len(not_found) > 10 else ""
            msg += (
                f"\n\nNote: {len(not_found)} placeholder(s) listed in the key file were "
                f"never found in the PDF ({shown}{more}) -- double check this key file "
                "actually belongs to this PDF."
            )
        messagebox.showinfo(APP_TITLE, msg)
        self.unredact_summary_var.set(f"Saved: {output_path}")


def main():
    app = DocumentToolkitApp()
    app.mainloop()


if __name__ == "__main__":
    main()
