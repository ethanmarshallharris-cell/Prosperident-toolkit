"""
PDF to Excel -- Prosperident Tool Suite.

Shows each page of a PDF with column-break bars you drag into place, then
converts the pages to an Excel workbook using those breaks. The conversion
logic is in pdf_table_engine.py; this file is only the window.

ON THE PAGE
  * Red vertical bars are COLUMN BREAKS. Drag one left or right to move it.
    Double-click anywhere on the page to add one. Right-click a bar (or
    select it and press Delete) to remove it. Left/Right arrow keys nudge
    the selected bar by one point (Shift = five points).
  * The two blue horizontal bars mark the TABLE AREA. Only text between
    them is converted, so a letterhead, page headings or a footer can be
    left out. They start at the very top and bottom of the page.
  * Column letters (A, B, C...) above the page show which Excel column each
    strip of the page becomes.
  * The same breaks apply to every page. The table area does too, unless
    "This page has its own table area" is ticked for a page.

SCANNED PAGES are read with OCR (see pdf_table_engine.py "SCANNED PAGES
(OCR)") in separate worker processes (OcrPool), so the window never
freezes and several pages are read at once; thin teal boxes show every
word OCR found. Conversion runs on its own thread with its own copy of the
PDF, and can be cancelled.

The preview on the right shows exactly what the current page will become,
and updates as bars move. Layouts (breaks + table area + options) can be
saved and loaded again for the next PDF in the same format.
"""
import logging
import multiprocessing
import os
import sys
import threading
import traceback
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pymupdf as fitz  # PyMuPDF
except ImportError:  # older PyMuPDF
    import fitz

import branding
import update_check
import pdf_table_engine as eng

APP_VERSION = "2026-11-12"
TOOL_KEY = "pdf_to_excel"
APP_TITLE = "PDF to Excel"

BREAK_COLOR = "#d0312d"
BREAK_SELECTED = "#74c146"
AREA_COLOR = "#2b6cb0"
SHADE = "#1c2331"
HIT_PX = 6            # how close (pixels) the pointer must be to grab a bar
MARGIN = 26           # pixels above the page for the column letters
ZOOMS = ["50%", "75%", "100%", "125%", "150%", "200%", "Fit width"]
OCR_BOX = "#2a9d8f"
OCR_CHOICES = {
    "auto": "Read scanned pages with OCR",
    "all": "OCR every page (ignore the PDF's own text)",
    "off": "Off (scanned pages convert as empty)",
}


class _Cancelled(Exception):
    pass


# --------------------------------------------------------------------------
# Logging (errors only, so problems can be diagnosed afterwards)
# --------------------------------------------------------------------------
def _log_dir():
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
        return base / "Prosperident" / "PDFtoExcel"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Logs" / "Prosperident PDF to Excel"
    return Path.home() / ".prosperident" / "pdf_to_excel"


def _setup_logging():
    try:
        d = _log_dir()
        d.mkdir(parents=True, exist_ok=True)
        path = d / "pdf_to_excel.log"
        logging.basicConfig(filename=str(path), level=logging.INFO,
                            format="%(asctime)s %(levelname)s %(message)s")
        return path
    except Exception:
        return None


LOG_PATH = _setup_logging()
log = logging.getLogger("pdf_to_excel")


def col_letter(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} - Prosperident  (v{APP_VERSION})")
        self.geometry("1320x860")
        self.minsize(980, 640)
        branding.apply_theme(self)
        branding.add_header(self, APP_TITLE, "Drag the red bars to where the columns break, then convert to Excel")

        self.doc = None
        self.pdf_path = None
        self.page_index = 0
        self.layout = eng.Layout()
        self.scale = 1.0
        self.photo = None
        self.tessdata = eng.find_tessdata()
        self.ocr_var = tk.StringVar(value=OCR_CHOICES["auto"] if self.tessdata else OCR_CHOICES["off"])
        self.source = eng.WordSource(self._ocr_mode(), self.tessdata)
        self._ocr_jobs = {}             # page -> Future, pages being OCR'd in the background
        self.ocr_pool = eng.OcrPool()
        self._autodetect_after_ocr = None
        self._convert_cancel = None     # threading.Event while a conversion runs
        self.drag = None          # ("break", index) or ("top"|"bottom",)
        self.selected = None      # index into self.layout.breaks
        self._preview_job = None
        self._app_update_download_url = None

        self.zoom_var = tk.StringVar(value="Fit width")
        self.page_var = tk.StringVar(value="")
        self.own_area_var = tk.BooleanVar(value=False)
        self.join_var = tk.BooleanVar(value=True)
        self.header_var = tk.BooleanVar(value=True)
        self.heading_rows_var = tk.StringVar(value="1")
        self.numbers_var = tk.BooleanVar(value=True)
        self.sheet_var = tk.StringVar(value="one")
        self.pagecol_var = tk.BooleanVar(value=True)
        self.range_var = tk.StringVar(value="")
        self.range_var.trace_add("write", lambda *a: self.schedule_preview(300) if self.doc else None)
        self.status_var = tk.StringVar(value="Open a PDF to begin.")
        self.app_update_var = tk.StringVar(value="")

        self._build()
        self._bind_keys()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_enabled(False)
        threading.Thread(target=self._update_worker, daemon=True).start()

    # ------------------------------------------------------------------
    # Layout of the window
    # ------------------------------------------------------------------
    def _build(self):
        upd = tk.Label(self, textvariable=self.app_update_var, bg="#fff6d6", fg=branding.NAVY,
                       cursor="hand2", anchor="w", padx=10)
        upd.bind("<Button-1>", self._open_update)
        self._update_label = upd

        bar = ttk.Frame(self, padding=(10, 8))
        bar.pack(side="top", fill="x")
        self._toolbar = bar
        ttk.Button(bar, text="Open PDF...", command=self.open_pdf).pack(side="left")
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)
        self.prev_btn = ttk.Button(bar, text="◀", width=3, command=lambda: self.go_page(-1))
        self.prev_btn.pack(side="left")
        ttk.Label(bar, textvariable=self.page_var, width=16, anchor="center").pack(side="left")
        self.next_btn = ttk.Button(bar, text="▶", width=3, command=lambda: self.go_page(1))
        self.next_btn.pack(side="left")
        ttk.Label(bar, text="  Zoom").pack(side="left")
        zoom = ttk.Combobox(bar, textvariable=self.zoom_var, values=ZOOMS, width=10, state="readonly")
        zoom.pack(side="left", padx=(4, 0))
        zoom.bind("<<ComboboxSelected>>", lambda e: self.render_page())
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)
        self.auto_btn = ttk.Button(bar, text="Suggest breaks", command=self.auto_detect)
        self.auto_btn.pack(side="left")
        self.add_btn = ttk.Button(bar, text="Add break", command=self.add_break_middle)
        self.add_btn.pack(side="left", padx=(6, 0))
        self.clear_btn = ttk.Button(bar, text="Clear breaks", command=self.clear_breaks)
        self.clear_btn.pack(side="left", padx=(6, 0))
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)
        self.load_btn = ttk.Button(bar, text="Load layout...", command=self.load_layout)
        self.load_btn.pack(side="left")
        self.save_btn = ttk.Button(bar, text="Save layout...", command=self.save_layout)
        self.save_btn.pack(side="left", padx=(6, 0))

        status = ttk.Label(self, textvariable=self.status_var, anchor="w", padding=(10, 4),
                           background=branding.GRAY_LIGHT)
        status.pack(side="bottom", fill="x")

        panes = ttk.PanedWindow(self, orient="horizontal")
        panes.pack(side="top", fill="both", expand=True, padx=10, pady=(0, 8))

        # --- left: the page ------------------------------------------------
        left = ttk.Frame(panes)
        self.canvas = tk.Canvas(left, bg="#8a909c", highlightthickness=0)
        vs = ttk.Scrollbar(left, orient="vertical", command=self.canvas.yview)
        hs = ttk.Scrollbar(left, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        ttk.Label(left, foreground=branding.GRAY, padding=(0, 4), justify="left", text=(
            "Red bars = column breaks: drag to move, double-click the page to add, right-click a bar to remove,\n"
            "arrow keys nudge the selected bar.  Blue bars = top and bottom of the table: drag to leave out headers and footers.")
        ).grid(row=2, column=0, columnspan=2, sticky="w")
        panes.add(left, weight=3)

        # --- right: options, preview, convert ----------------------------
        right = ttk.Frame(panes, padding=(10, 0, 0, 0))
        panes.add(right, weight=2)

        convert = ttk.LabelFrame(right, text="Convert", padding=10)
        convert.pack(side="bottom", fill="x", pady=(8, 0))
        r = ttk.Frame(convert); r.pack(fill="x")
        ttk.Label(r, text="Pages (blank = all):").pack(side="left")
        ttk.Entry(r, textvariable=self.range_var, width=16).pack(side="left", padx=6)
        ttk.Label(r, text="e.g. 1-3, 5", foreground=branding.GRAY).pack(side="left")
        r = ttk.Frame(convert); r.pack(fill="x", pady=(6, 0))
        ttk.Radiobutton(r, text="All pages on one sheet", value="one", variable=self.sheet_var).pack(side="left")
        ttk.Radiobutton(r, text="One sheet per page", value="each", variable=self.sheet_var).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(convert, text="Add a Page column (one-sheet only)", variable=self.pagecol_var).pack(anchor="w", pady=(4, 0))
        self.progress = ttk.Progressbar(convert, mode="determinate")
        self.progress.pack(fill="x", pady=(8, 0))
        self.convert_btn = ttk.Button(convert, text="Convert to Excel...", command=self.convert)
        self.convert_btn.pack(fill="x", pady=(8, 0))

        opts = ttk.LabelFrame(right, text="Options", padding=10)
        opts.pack(side="top", fill="x")
        ttk.Checkbutton(opts, text="Join wrapped lines to the row above",
                        variable=self.join_var, command=self._options_changed).pack(anchor="w")
        hr = ttk.Frame(opts); hr.pack(fill="x")
        ttk.Checkbutton(hr, text="Show column headings only once",
                        variable=self.header_var, command=self._options_changed).pack(side="left")
        ttk.Label(hr, text="   Heading lines:").pack(side="left")
        self.heading_spin = ttk.Spinbox(hr, from_=1, to=10, width=3, textvariable=self.heading_rows_var,
                                        command=self._options_changed)
        self.heading_spin.pack(side="left", padx=(4, 0))
        self.heading_spin.bind("<KeyRelease>", lambda e: self._options_changed())
        self.heading_spin.bind("<FocusOut>", lambda e: self._options_changed())
        ttk.Checkbutton(opts, text="Convert amounts to numbers, e.g. (45.00) = -45",
                        variable=self.numbers_var, command=self._options_changed).pack(anchor="w")
        orow = ttk.Frame(opts); orow.pack(fill="x", pady=(6, 0))
        ttk.Label(orow, text="Scanned pages:").pack(side="left")
        self.ocr_combo = ttk.Combobox(orow, textvariable=self.ocr_var, values=list(OCR_CHOICES.values()),
                                      state="readonly" if self.tessdata else "disabled", width=40)
        self.ocr_combo.pack(side="left", padx=(6, 0))
        self.ocr_combo.bind("<<ComboboxSelected>>", lambda e: self._ocr_mode_changed())
        if not self.tessdata:
            ttk.Label(opts, foreground=BREAK_COLOR, wraplength=420, justify="left", text=(
                "OCR isn't available: the OCR language file (tessdata\\eng.traineddata) is missing. "
                "Download PDF to Excel again from the Tool Suite page.")).pack(anchor="w")
        self.own_area_chk = ttk.Checkbutton(opts, text="This page has its own table top/bottom",
                                            variable=self.own_area_var, command=self._own_area_toggled)
        self.own_area_chk.pack(anchor="w", pady=(6, 0))
        r = ttk.Frame(opts); r.pack(fill="x", pady=(6, 0))
        self.reset_area_btn = ttk.Button(r, text="Reset table area", command=self.reset_area)
        self.reset_area_btn.pack(side="left")

        prev = ttk.LabelFrame(right, text="Preview of this page", padding=6)
        prev.pack(side="top", fill="both", expand=True, pady=(8, 0))
        self.tree = ttk.Treeview(prev, show="headings")
        pv = ttk.Scrollbar(prev, orient="vertical", command=self.tree.yview)
        ph = ttk.Scrollbar(prev, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=pv.set, xscrollcommand=ph.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        pv.grid(row=0, column=1, sticky="ns")
        ph.grid(row=1, column=0, sticky="ew")
        prev.rowconfigure(0, weight=1)
        prev.columnconfigure(0, weight=1)
        self.preview_note = ttk.Label(prev, text="", foreground=branding.GRAY, wraplength=420, justify="left")
        self.preview_note.grid(row=2, column=0, columnspan=2, sticky="w")

        c = self.canvas
        c.bind("<Motion>", self._on_motion)
        c.bind("<ButtonPress-1>", self._on_press)
        c.bind("<B1-Motion>", self._on_drag)
        c.bind("<ButtonRelease-1>", self._on_release)
        c.bind("<Double-Button-1>", self._on_double)
        c.bind("<Button-3>", self._on_right)
        c.bind("<Button-2>", self._on_right)   # macOS secondary click on some mice
        c.bind("<Configure>", lambda e: self._maybe_refit())
        c.bind("<MouseWheel>", lambda e: c.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units"))
        c.bind("<Button-4>", lambda e: c.yview_scroll(-3, "units"))
        c.bind("<Button-5>", lambda e: c.yview_scroll(3, "units"))
        self._last_canvas_w = 0

    def _bind_keys(self):
        self.bind("<Delete>", lambda e: self._delete_selected())
        self.bind("<BackSpace>", lambda e: self._delete_selected() if self.focus_get() is self.canvas else None)
        for key, step in (("<Left>", -1), ("<Right>", 1), ("<Shift-Left>", -5), ("<Shift-Right>", 5)):
            self.canvas.bind(key, lambda e, s=step: self._nudge(s))
        self.bind("<Prior>", lambda e: self.go_page(-1))
        self.bind("<Next>", lambda e: self.go_page(1))

    def _set_enabled(self, on):
        st = "normal" if on else "disabled"
        for w in (self.prev_btn, self.next_btn, self.auto_btn, self.add_btn, self.clear_btn,
                  self.save_btn, self.convert_btn, self.reset_area_btn, self.own_area_chk):
            w.configure(state=st)

    # ------------------------------------------------------------------
    # Opening and showing pages
    # ------------------------------------------------------------------
    def open_pdf(self, path=None):
        path = path or filedialog.askopenfilename(title="Choose a PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not path:
            return
        try:
            doc = eng.open_pdf(path)
        except Exception as e:
            log.exception("open failed: %s", path)
            messagebox.showerror(APP_TITLE, f"Couldn't open this PDF.\n\n{e}")
            return
        if doc.page_count == 0:
            messagebox.showerror(APP_TITLE, "This PDF has no pages.")
            return
        if self.doc:
            self.doc.close()
        self.doc = doc
        self.pdf_path = path
        self.page_index = 0
        self.source = eng.WordSource(self._ocr_mode(), self.tessdata)
        self._ocr_jobs = {}
        self.selected = None
        # Keep breaks the user already placed or loaded (same statement
        # format, next month) -- but drop page-specific areas, which belong
        # to the old file's page numbers.
        self.layout.page_areas = {}
        self._set_enabled(True)
        self.title(f"{APP_TITLE} - {Path(path).name}  (v{APP_VERSION})")

        check = range(min(doc.page_count, 5))
        scanned = not any(eng.page_has_text(doc[i]) for i in check)
        if scanned and not self.tessdata:
            messagebox.showwarning(APP_TITLE, (
                "This PDF appears to be scanned: its pages are pictures of text, with no "
                "text the program can read, and OCR isn't available on this computer (the "
                "OCR language file is missing).\n\nDownload PDF to Excel again from the Tool "
                "Suite page, or run the PDF through OCR first (for example Adobe Acrobat's "
                "Scan & OCR tool) and open the OCR'd copy here."))
        self.page_px = None
        self.render_page()
        if not self.layout.breaks:
            self.auto_detect(quiet=True)
        if scanned and self.tessdata and self._ocr_mode() != "off":
            self.status_var.set("Scanned PDF: pages are read with OCR (a few seconds each). "
                                "OCR can misread a character -- check the numbers against the PDF.")

    # ------------------------------------------------------------------
    # Words: the PDF's own text, or OCR in the background for scanned pages
    # ------------------------------------------------------------------
    def _ocr_mode(self):
        label = self.ocr_var.get()
        return next((k for k, v in OCR_CHOICES.items() if v == label), "auto")

    def _ocr_mode_changed(self):
        self.source = eng.WordSource(self._ocr_mode(), self.tessdata)
        self._ocr_jobs = {}
        if self.doc:
            self.render_page()

    def words(self, i):
        """Page i's words, or None while OCR is still reading it (it is
        then started in the background; _ocr_done refreshes the view)."""
        if self.source.ready(self.doc, i):
            try:
                return self.source.words(self.doc, i)
            except Exception:
                log.exception("words page %s", i)
                return []
        self._start_ocr(i)
        return None

    def _start_ocr(self, i):
        if i in self._ocr_jobs or not self.pdf_path or not self.tessdata:
            return
        self.status_var.set(f"Reading page {i + 1} with OCR...")
        fut = self.ocr_pool.submit(self.pdf_path, i, self.tessdata)
        self._ocr_jobs[i] = fut
        source, path = self.source, self.pdf_path
        self.after(150, lambda: self._poll_ocr(i, fut, source, path))

    def _poll_ocr(self, i, fut, source, path):
        if not fut.done():
            self.after(150, lambda: self._poll_ocr(i, fut, source, path))
            return
        err = None
        try:
            source.store_ocr(i, fut.result())
        except Exception as e:  # noqa: BLE001
            log.error("OCR page %s failed: %s", i, e)
            err = e
        self._ocr_done(i, source, path, err)

    def _ocr_done(self, i, source, path, err):
        if source is not self.source or path != self.pdf_path:
            return  # the PDF or the OCR setting changed meanwhile
        self._ocr_jobs.pop(i, None)
        if err is not None:
            self.status_var.set(f"OCR failed on page {i + 1}: {err}")
            if i == self.page_index:
                self.canvas.delete("ocrmsg")
            return
        n = len(source.words(self.doc, i))
        if i == self.page_index:
            self.canvas.delete("ocrmsg")
            self.status_var.set(f"Page {i + 1} read with OCR ({n} words, shown in teal boxes). "
                                "Check the numbers against the PDF.")
            if self._autodetect_after_ocr == i:
                self._autodetect_after_ocr = None
                self.auto_detect(quiet=True)
            else:
                self.draw_overlay()
        self.schedule_preview(0)

    def go_page(self, step):
        if not self.doc:
            return
        n = self.page_index + step
        if 0 <= n < self.doc.page_count:
            self.page_index = n
            self.render_page()

    def _current_zoom(self):
        z = self.zoom_var.get()
        page = self.doc[self.page_index]
        if z == "Fit width":
            w = max(self.canvas.winfo_width() - 30, 300)
            return w / page.rect.width
        return int(z.rstrip("%")) / 100 * 96 / 72

    def _maybe_refit(self):
        w = self.canvas.winfo_width()
        if self.doc and self.zoom_var.get() == "Fit width" and abs(w - self._last_canvas_w) > 20:
            self._last_canvas_w = w
            self.render_page()

    def render_page(self):
        if not self.doc:
            return
        page = self.doc[self.page_index]
        self.scale = self._current_zoom()
        pix = page.get_pixmap(matrix=fitz.Matrix(self.scale, self.scale), alpha=False)
        self.photo = tk.PhotoImage(data=pix.tobytes("ppm"))
        self.page_px = (pix.width, pix.height)
        self.offset = (12, MARGIN)
        c = self.canvas
        c.delete("all")
        c.create_rectangle(self.offset[0] + 3, self.offset[1] + 3, self.offset[0] + pix.width + 3,
                           self.offset[1] + pix.height + 3, fill="#5b6472", outline="")
        c.create_image(self.offset[0], self.offset[1], image=self.photo, anchor="nw", tags="page")
        c.configure(scrollregion=(0, 0, pix.width + 30, pix.height + MARGIN + 20))
        self.page_var.set(f"Page {self.page_index + 1} of {self.doc.page_count}")
        self.prev_btn.configure(state="normal" if self.page_index > 0 else "disabled")
        self.next_btn.configure(state="normal" if self.page_index < self.doc.page_count - 1 else "disabled")
        self.own_area_var.set(self.page_index in self.layout.page_areas)
        if self._ocr_mode() == "off" and not eng.page_has_text(page):
            c.create_text(self.offset[0] + pix.width / 2, self.offset[1] + 40, fill=BREAK_COLOR,
                          font=("Segoe UI", 12, "bold"), tags="overlay",
                          text="No readable text on this page (scanned image) - it will convert as empty.")
        elif not self.source.ready(self.doc, self.page_index):
            c.create_text(self.offset[0] + pix.width / 2, self.offset[1] + 40, fill=OCR_BOX,
                          font=("Segoe UI", 12, "bold"), tags=("overlay", "ocrmsg"),
                          text="Scanned page - reading it with OCR...")
            self._start_ocr(self.page_index)
        self.draw_overlay()
        self.schedule_preview()

    # point <-> canvas pixel
    def px(self, x_pt):
        return self.offset[0] + x_pt * self.scale

    def py(self, y_pt):
        return self.offset[1] + y_pt * self.scale

    def pt_x(self, cx):
        return (cx - self.offset[0]) / self.scale

    def pt_y(self, cy):
        return (cy - self.offset[1]) / self.scale

    def _page_area(self):
        return self.layout.area_for(self.page_index)

    def _area_edges(self):
        """Top and bottom of the table area in points (page edges if unset)."""
        page = self.doc[self.page_index]
        a = self._page_area()
        top = a.top if a.top is not None else 0.0
        bottom = a.bottom if a.bottom is not None else page.rect.height
        return top, bottom

    def draw_overlay(self):
        c = self.canvas
        c.delete("overlay2")
        if not self.doc or not getattr(self, "page_px", None):
            return
        w, h = self.page_px
        x0, y0 = self.offset
        top, bottom = self._area_edges()
        ty, by = self.py(top), self.py(bottom)
        # shade what is outside the table area
        if ty > y0 + 1:
            c.create_rectangle(x0, y0, x0 + w, ty, fill=SHADE, stipple="gray50", outline="", tags="overlay2")
        if by < y0 + h - 1:
            c.create_rectangle(x0, by, x0 + w, y0 + h, fill=SHADE, stipple="gray50", outline="", tags="overlay2")
        for name, y in (("top", ty), ("bottom", by)):
            c.create_line(x0, y, x0 + w, y, fill=AREA_COLOR, width=3, dash=(8, 4), tags=("overlay2", f"area_{name}"))
            label = "Table top" if name == "top" else "Table bottom"
            c.create_text(x0 + w - 6, y + (-8 if name == "top" and y > y0 + 14 else 9), text=label,
                          anchor="e", fill=AREA_COLOR, font=("Segoe UI", 8, "bold"), tags="overlay2")
        # words found by OCR, so the bars can be placed against them
        if self.source.uses_ocr(self.doc, self.page_index) and self.source.ready(self.doc, self.page_index):
            for wd in self.source.words(self.doc, self.page_index):
                c.create_rectangle(self.px(wd.x0), self.py(wd.y0), self.px(wd.x1), self.py(wd.y1),
                                   outline=OCR_BOX, width=1, tags="overlay2")
        # column breaks
        bs = sorted(self.layout.breaks)
        for i, b in enumerate(bs):
            x = self.px(b)
            color = BREAK_SELECTED if self.selected == i else BREAK_COLOR
            c.create_line(x, ty, x, by, fill=color, width=3 if self.selected == i else 2, tags="overlay2")
            c.create_line(x, y0 - 6, x, ty, fill=color, width=1, dash=(2, 3), tags="overlay2")
            c.create_polygon(x - 6, y0 - 12, x + 6, y0 - 12, x, y0 - 3, fill=color, outline="", tags="overlay2")
        # column letters
        edges = [0.0] + bs + [self.doc[self.page_index].rect.width]
        for i in range(len(edges) - 1):
            mid = (self.px(edges[i]) + self.px(edges[i + 1])) / 2
            if self.px(edges[i + 1]) - self.px(edges[i]) > 14:
                c.create_text(mid, y0 - 17, text=col_letter(i), fill=branding.WHITE,
                              font=("Segoe UI", 9, "bold"), tags="overlay2")

    # ------------------------------------------------------------------
    # Mouse
    # ------------------------------------------------------------------
    def _hit(self, cx, cy):
        """What bar (if any) is under the pointer."""
        if not self.doc or not getattr(self, "page_px", None):
            return None
        top, bottom = self._area_edges()
        x0, y0 = self.offset
        w, h = self.page_px
        if not (x0 - HIT_PX <= cx <= x0 + w + HIT_PX and y0 - 14 <= cy <= y0 + h + HIT_PX):
            return None
        for name, y in (("top", self.py(top)), ("bottom", self.py(bottom))):
            if abs(cy - y) <= HIT_PX:
                return (name,)
        bs = sorted(self.layout.breaks)
        best = None
        for i, b in enumerate(bs):
            d = abs(cx - self.px(b))
            if d <= HIT_PX and (best is None or d < best[1]):
                best = (i, d)
        if best:
            return ("break", best[0])
        return None

    def _xy(self, e):
        return self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)

    def _on_motion(self, e):
        hit = self._hit(*self._xy(e))
        cur = ""
        if hit and hit[0] == "break":
            cur = "sb_h_double_arrow"
        elif hit:
            cur = "sb_v_double_arrow"
        self.canvas.configure(cursor=cur)

    def _on_press(self, e):
        self.canvas.focus_set()
        hit = self._hit(*self._xy(e))
        self.drag = hit
        if hit and hit[0] == "break":
            self.selected = hit[1]
        elif hit is None:
            self.selected = None
        self.draw_overlay()

    def _on_drag(self, e):
        if not self.drag or not self.doc:
            return
        cx, cy = self._xy(e)
        page = self.doc[self.page_index]
        if self.drag[0] == "break":
            bs = sorted(self.layout.breaks)
            i = self.drag[1]
            x = min(max(self.pt_x(cx), 1.0), page.rect.width - 1.0)
            bs[i] = x
            # keep the bar's identity while it is dragged past a neighbor
            order = sorted(range(len(bs)), key=lambda k: bs[k])
            self.layout.breaks = [bs[k] for k in order]
            self.drag = ("break", order.index(i))
            self.selected = self.drag[1]
        else:
            y = min(max(self.pt_y(cy), 0.0), page.rect.height)
            top, bottom = self._area_edges()
            area = self._editable_area()
            if self.drag[0] == "top":
                area.top = min(y, bottom - 4)
            else:
                area.bottom = max(y, top + 4)
        self.draw_overlay()
        self.schedule_preview()

    def _on_release(self, e):
        if self.drag:
            self.schedule_preview(0)
        self.drag = None

    def _on_double(self, e):
        if not self.doc:
            return
        cx, cy = self._xy(e)
        x0, y0 = self.offset
        w, h = self.page_px
        if not (x0 <= cx <= x0 + w and y0 - 14 <= cy <= y0 + h):
            return
        if self._hit(cx, cy):
            return
        self._add_break(self.pt_x(cx))

    def _on_right(self, e):
        hit = self._hit(*self._xy(e))
        if hit and hit[0] == "break":
            bs = sorted(self.layout.breaks)
            del bs[hit[1]]
            self.layout.breaks = bs
            self.selected = None
            self.status_var.set(f"Break removed. {len(bs) + 1} column(s).")
            self.draw_overlay()
            self.schedule_preview(0)

    def _delete_selected(self):
        if self.selected is not None and self.selected < len(self.layout.breaks):
            bs = sorted(self.layout.breaks)
            del bs[self.selected]
            self.layout.breaks = bs
            self.selected = None
            self.status_var.set(f"Break removed. {len(bs) + 1} column(s).")
            self.draw_overlay()
            self.schedule_preview(0)

    def _nudge(self, step):
        if self.selected is None or not self.doc:
            return
        bs = sorted(self.layout.breaks)
        w = self.doc[self.page_index].rect.width
        bs[self.selected] = min(max(bs[self.selected] + step, 1.0), w - 1.0)
        order = sorted(range(len(bs)), key=lambda k: bs[k])
        self.layout.breaks = [bs[k] for k in order]
        self.selected = order.index(self.selected)
        self.draw_overlay()
        self.schedule_preview()
        return "break"

    def _add_break(self, x_pt):
        bs = sorted(self.layout.breaks + [x_pt])
        self.layout.breaks = bs
        self.selected = bs.index(x_pt)
        self.draw_overlay()
        self.schedule_preview(0)
        self.status_var.set(f"Break added. {len(bs) + 1} columns.  Drag it into place.")

    # ------------------------------------------------------------------
    # Toolbar actions
    # ------------------------------------------------------------------
    def add_break_middle(self):
        """Adds a break in the middle of the widest column, ready to drag."""
        if not self.doc:
            return
        w = self.doc[self.page_index].rect.width
        edges = [0.0] + sorted(self.layout.breaks) + [w]
        widest = max(range(len(edges) - 1), key=lambda i: edges[i + 1] - edges[i])
        self._add_break((edges[widest] + edges[widest + 1]) / 2)

    def clear_breaks(self):
        if self.layout.breaks and not messagebox.askyesno(APP_TITLE, "Remove every column break?"):
            return
        self.layout.breaks = []
        self.selected = None
        self.draw_overlay()
        self.schedule_preview(0)

    def auto_detect(self, quiet=False):
        if not self.doc:
            return
        if self.layout.breaks and not quiet:
            if not messagebox.askyesno(APP_TITLE, "Replace the current breaks with suggested ones for this page?"):
                return
        words = self.words(self.page_index)
        if words is None:
            self._autodetect_after_ocr = self.page_index
            self.status_var.set("Breaks will be suggested once OCR has read this page...")
            return
        bs = eng.auto_detect_breaks(words, self._page_area())
        self.layout.breaks = bs
        self.selected = None
        self.draw_overlay()
        self.schedule_preview(0)
        if bs:
            self.status_var.set(f"Suggested {len(bs)} break(s) from the white space on this page. "
                                "Drag, add or remove bars until the preview looks right.")
        else:
            self.status_var.set("No clear column gaps found on this page. Double-click the page to add breaks by hand.")

    def _editable_area(self):
        """The Area object dragging the blue bars should change."""
        if self.own_area_var.get():
            if self.page_index not in self.layout.page_areas:
                a = self.layout.area
                self.layout.page_areas[self.page_index] = eng.Area(a.top, a.bottom, a.left, a.right)
            return self.layout.page_areas[self.page_index]
        return self.layout.area

    def _own_area_toggled(self):
        if self.own_area_var.get():
            self._editable_area()
        else:
            self.layout.page_areas.pop(self.page_index, None)
        self.draw_overlay()
        self.schedule_preview(0)

    def reset_area(self):
        a = self._editable_area()
        a.top = a.bottom = None
        self.draw_overlay()
        self.schedule_preview(0)

    def _options_changed(self):
        o = self.layout.options
        o.join_wrapped = self.join_var.get()
        o.skip_repeated_header = self.header_var.get()
        o.convert_numbers = self.numbers_var.get()
        try:
            o.heading_rows = max(1, min(int(self.heading_rows_var.get()), 10))
        except ValueError:
            pass  # half-typed; keep the last good value
        self.heading_spin.configure(state="normal" if o.skip_repeated_header else "disabled")
        self.schedule_preview(0)

    def _sync_option_vars(self):
        o = self.layout.options
        self.join_var.set(o.join_wrapped)
        self.header_var.set(o.skip_repeated_header)
        self.numbers_var.set(o.convert_numbers)
        self.heading_rows_var.set(str(o.heading_rows))
        self.heading_spin.configure(state="normal" if o.skip_repeated_header else "disabled")

    def save_layout(self):
        path = filedialog.asksaveasfilename(title="Save layout", defaultextension=".pdflayout.json",
                                            filetypes=[("PDF to Excel layout", "*.pdflayout.json"), ("JSON", "*.json")],
                                            initialfile=(Path(self.pdf_path).stem + " layout.pdflayout.json") if self.pdf_path else "layout.pdflayout.json")
        if not path:
            return
        try:
            self._options_changed()
            Path(path).write_text(self.layout.to_json(), encoding="utf-8")
            self.status_var.set(f"Layout saved: {path}")
        except Exception as e:
            log.exception("save layout")
            messagebox.showerror(APP_TITLE, f"Couldn't save the layout.\n\n{e}")

    def load_layout(self):
        path = filedialog.askopenfilename(title="Load layout",
                                          filetypes=[("PDF to Excel layout", "*.pdflayout.json *.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            self.layout = eng.Layout.from_json(Path(path).read_text(encoding="utf-8"))
        except Exception as e:
            log.exception("load layout")
            messagebox.showerror(APP_TITLE, f"Couldn't load this layout.\n\n{e}")
            return
        if self.doc:
            n = self.doc.page_count
            self.layout.page_areas = {k: v for k, v in self.layout.page_areas.items() if k < n}
        self._sync_option_vars()
        self.selected = None
        self.status_var.set(f"Layout loaded: {path}  ({len(self.layout.breaks)} break(s))")
        if self.doc:
            self.render_page()

    # ------------------------------------------------------------------
    # Preview
    # ------------------------------------------------------------------
    def schedule_preview(self, delay=120):
        if self._preview_job:
            self.after_cancel(self._preview_job)
        self._preview_job = self.after(delay, self.update_preview)

    def _range_pages(self):
        try:
            return eng.parse_page_range(self.range_var.get(), self.doc.page_count)
        except ValueError:
            return list(range(self.doc.page_count))

    def update_preview(self):
        self._preview_job = None
        if not self.doc:
            return
        hidden = 0
        n_head = 0
        words = self.words(self.page_index)
        if words is None:
            self.tree.delete(*self.tree.get_children())
            self.preview_note.configure(text="Reading this page with OCR... the preview appears in a few seconds.")
            return
        try:
            rows = eng.page_rows(self.doc[self.page_index], self.layout, self.page_index, words)
            if self.layout.options.skip_repeated_header and rows:
                pages = self._range_pages()
                first, headings = None, []
                count = max(1, int(self.layout.options.heading_rows or 1))
                for i in pages:
                    w = words if i == self.page_index else self.words(i)
                    if w is None:
                        break   # that page is still being OCR'd; headings are checked afterwards
                    r = eng.page_rows(self.doc[i], self.layout, i, w)
                    if r:
                        first, headings = i, [list(x) for x in r[:count]]
                        break
                if self.page_index == first:
                    n_head = len(headings)
                elif headings and self.page_index in pages and (first is None or self.page_index > first):
                    kept = eng.strip_headings(rows, headings)
                    hidden = len(rows) - len(kept)
                    rows = kept
        except Exception:
            log.exception("preview")
            rows = []
        ncols = len(self.layout.breaks) + 1
        cols = [col_letter(i) for i in range(ncols)]
        self.tree.delete(*self.tree.get_children())
        self.tree.configure(columns=cols)
        widths = [4] * ncols
        for r in rows[:200]:
            for i, v in enumerate(r):
                widths[i] = max(widths[i], len(v))
        for i, c in enumerate(cols):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=min(max(widths[i] * 8 + 16, 50), 320), minwidth=40, stretch=False)
        for r in rows[:500]:
            vals = []
            for v in r:
                if self.layout.options.convert_numbers:
                    n = eng.to_number(v)
                    vals.append(v if n is None else f"{n:,}" if isinstance(n, int) else f"{n:,.2f}")
                else:
                    vals.append(v)
            k = len(self.tree.get_children())
            self.tree.insert("", "end", values=vals, tags=("heading",) if k < n_head else ())
        self.tree.tag_configure("heading", background="#e3f1d8", font=("Segoe UI", 9, "bold"))
        more = f"  (first 500 shown)" if len(rows) > 500 else ""
        note = f"{len(rows)} row(s) x {ncols} column(s) on this page{more}"
        if n_head:
            note += ".  Green = headings (kept once)"
        if hidden:
            note += f".  {hidden} repeated heading row(s) hidden"
        self.preview_note.configure(text=note)

    # ------------------------------------------------------------------
    # Convert
    # ------------------------------------------------------------------
    def convert(self):
        if self._convert_cancel is not None:      # the button reads "Cancel conversion"
            self._convert_cancel.set()
            self.status_var.set("Cancelling after the current page...")
            return
        if not self.doc:
            return
        self._options_changed()
        try:
            pages = eng.parse_page_range(self.range_var.get(), self.doc.page_count)
        except ValueError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        if not pages:
            messagebox.showerror(APP_TITLE, "No pages selected.")
            return
        out = filedialog.asksaveasfilename(title="Save Excel workbook", defaultextension=".xlsx",
                                           filetypes=[("Excel workbook", "*.xlsx")],
                                           initialdir=str(Path(self.pdf_path).parent),
                                           initialfile=Path(self.pdf_path).stem + ".xlsx")
        if not out:
            return
        self.progress.configure(maximum=len(pages), value=0)
        self._convert_cancel = threading.Event()
        cancel = self._convert_cancel
        self.convert_btn.configure(text="Cancel conversion")
        layout = eng.Layout.from_json(self.layout.to_json())   # a copy: bars may move meanwhile
        source, path = self.source, self.pdf_path
        one_sheet, page_col = self.sheet_var.get() == "one", self.pagecol_var.get()
        ocr_needed = sum(1 for i in pages if not source.ready(self.doc, i))
        self.status_var.set(f"Converting {len(pages)} page(s)"
                            + (f", {ocr_needed} of them by OCR (a few seconds each)..." if ocr_needed else "..."))

        def progress(n, total):
            if cancel.is_set():
                raise _Cancelled()
            self.after(0, lambda: (self.progress.configure(value=n),
                                   self.status_var.set(f"Read {n} of {total} page(s)...")))

        pool, tessdata = self.ocr_pool, self.tessdata

        def work():
            result, err = None, None
            futures = []
            try:
                d = eng.open_pdf(path)
                try:
                    # Scanned pages first, in worker processes, several at a
                    # time (see OcrPool); then the fast part, page by page.
                    todo = [i for i in pages if not source.ready(d, i)]
                    if todo:
                        futures = [(i, pool.submit(path, i, tessdata)) for i in todo]
                        for k, (i, f) in enumerate(futures, 1):
                            while True:
                                if cancel.is_set():
                                    raise _Cancelled()
                                try:
                                    source.store_ocr(i, f.result(timeout=0.25))
                                    break
                                except TimeoutError:
                                    continue
                            self.after(0, lambda k=k: (
                                self.progress.configure(value=k),
                                self.status_var.set(f"Read {k} of {len(todo)} scanned page(s) with OCR...")))
                    rows = eng.extract(d, layout, pages, progress, source=source)
                    headings = (eng.find_headings(d, layout, pages, source)
                                if layout.options.skip_repeated_header else None)
                    ocr_pages = sorted(set(pages) & source.ocr_pages)
                    if cancel.is_set():
                        raise _Cancelled()
                    self.after(0, lambda: self.status_var.set("Writing the workbook..."))
                    n = eng.write_workbook(out, rows, layout.options, one_sheet=one_sheet,
                                           page_column=page_col, source_name=Path(path).name,
                                           headings=headings, ocr_pages=ocr_pages)
                    result = (rows, n, ocr_pages)
                finally:
                    d.close()
            except Exception as e:  # noqa: BLE001
                err = e
                for _i, f in futures:
                    f.cancel()
                if not isinstance(e, (_Cancelled, PermissionError)):
                    log.error("convert failed: %s\n%s", e, traceback.format_exc())
            self.after(0, lambda: self._convert_done(out, pages, result, err))
        threading.Thread(target=work, daemon=True).start()

    def _convert_done(self, out, pages, result, err):
        self._convert_cancel = None
        self.convert_btn.configure(text="Convert to Excel...", state="normal")
        if isinstance(err, _Cancelled):
            self.status_var.set("Conversion cancelled -- nothing was saved.")
            self.progress.configure(value=0)
            return
        if isinstance(err, PermissionError):
            messagebox.showerror(APP_TITLE, f"Couldn't save to\n{out}\n\nIs it open in Excel? Close it and try again.")
            self.status_var.set("Not saved.")
            return
        if err is not None:
            messagebox.showerror(APP_TITLE, f"The conversion failed.\n\n{err}\n\nDetails were written to\n{LOG_PATH}")
            self.status_var.set("Conversion failed.")
            return
        rows, n, ocr_pages = result
        if self.doc:
            self.draw_overlay()
            self.schedule_preview(0)
        empty = [i + 1 for i in pages if not rows.get(i)]
        msg = f"Saved {n:,} row(s) from {len(pages)} page(s) to\n{out}"
        if ocr_pages:
            shown = ", ".join(str(i + 1) for i in ocr_pages[:20]) + ("..." if len(ocr_pages) > 20 else "")
            msg += (f"\n\n{len(ocr_pages)} scanned page(s) were read with OCR ({shown}). Their rows are "
                    "tinted yellow in the workbook: OCR can misread a character (0/O, 1/l, 5/S), so "
                    "check those numbers against the PDF.")
        if empty:
            msg += f"\n\nNo text was found in the table area on page(s): {', '.join(map(str, empty[:20]))}{'...' if len(empty) > 20 else ''}"
        msg += "\n\nCheck the totals against the PDF before relying on the conversion."
        self.status_var.set(f"Saved {out}")
        if messagebox.askyesno(APP_TITLE, msg + "\n\nOpen the workbook now?"):
            _open_file(out)

    def _on_close(self):
        if self._convert_cancel is not None:
            self._convert_cancel.set()
        try:
            self.ocr_pool.shutdown()
        finally:
            self.destroy()

    # ------------------------------------------------------------------
    # Update check
    # ------------------------------------------------------------------
    def _update_worker(self):
        result = update_check.check_for_update(TOOL_KEY, APP_VERSION)
        if result:
            self.after(0, lambda: self._show_update(result))

    def _show_update(self, entry):
        text = f"A newer version of PDF to Excel is available (v{entry.get('version')}) - click to get it."
        if entry.get("notes"):
            text += "  " + entry["notes"]
        self._app_update_download_url = entry.get("download_url")
        self.app_update_var.set(text)
        self._update_label.pack(side="top", fill="x", before=self._toolbar)

    def _open_update(self, event=None):
        if self._app_update_download_url:
            webbrowser.open(self._app_update_download_url)


def _open_file(path):
    try:
        if sys.platform == "win32":
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", path])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", path])
    except Exception:
        log.exception("open file")


def main():
    multiprocessing.freeze_support()   # OCR worker processes in the built app
    app = App()
    if len(sys.argv) > 1 and Path(sys.argv[1]).is_file():
        app.after(100, lambda: app.open_pdf(sys.argv[1]))
    app.mainloop()


if __name__ == "__main__":
    main()
