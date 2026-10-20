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

The preview on the right shows exactly what the current page will become,
and updates as bars move. Layouts (breaks + table area + options) can be
saved and loaded again for the next PDF in the same format.
"""
import logging
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

APP_VERSION = "2026-10-20"
TOOL_KEY = "pdf_to_excel"
APP_TITLE = "PDF to Excel"

BREAK_COLOR = "#d0312d"
BREAK_SELECTED = "#74c146"
AREA_COLOR = "#2b6cb0"
SHADE = "#1c2331"
HIT_PX = 6            # how close (pixels) the pointer must be to grab a bar
MARGIN = 26           # pixels above the page for the column letters
ZOOMS = ["50%", "75%", "100%", "125%", "150%", "200%", "Fit width"]


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
        self.words_cache = {}
        self.drag = None          # ("break", index) or ("top"|"bottom",)
        self.selected = None      # index into self.layout.breaks
        self._preview_job = None
        self._app_update_download_url = None

        self.zoom_var = tk.StringVar(value="Fit width")
        self.page_var = tk.StringVar(value="")
        self.own_area_var = tk.BooleanVar(value=False)
        self.join_var = tk.BooleanVar(value=True)
        self.header_var = tk.BooleanVar(value=True)
        self.numbers_var = tk.BooleanVar(value=True)
        self.sheet_var = tk.StringVar(value="one")
        self.pagecol_var = tk.BooleanVar(value=True)
        self.range_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Open a PDF to begin.")
        self.app_update_var = tk.StringVar(value="")

        self._build()
        self._bind_keys()
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
        ttk.Checkbutton(opts, text="Skip column headings repeated on later pages",
                        variable=self.header_var, command=self._options_changed).pack(anchor="w")
        ttk.Checkbutton(opts, text="Convert amounts to numbers, e.g. (45.00) = -45",
                        variable=self.numbers_var, command=self._options_changed).pack(anchor="w")
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
        self.preview_note = ttk.Label(prev, text="", foreground=branding.GRAY)
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
        self.words_cache = {}
        self.selected = None
        # Keep breaks the user already placed or loaded (same statement
        # format, next month) -- but drop page-specific areas, which belong
        # to the old file's page numbers.
        self.layout.page_areas = {}
        self._set_enabled(True)
        self.title(f"{APP_TITLE} - {Path(path).name}  (v{APP_VERSION})")

        check = range(min(doc.page_count, 5))
        if not any(eng.page_has_text(doc[i]) for i in check):
            messagebox.showwarning(APP_TITLE, (
                "This PDF appears to be scanned: its pages are pictures of text, with no "
                "text the program can read.\n\nPDF to Excel needs a PDF with real text. "
                "Run the PDF through OCR first (for example Adobe Acrobat's Scan & OCR "
                "tool), then open the OCR'd copy here."))
        self.page_px = None
        self.render_page()
        if not self.layout.breaks:
            self.auto_detect(quiet=True)

    def words(self, i):
        if i not in self.words_cache:
            self.words_cache[i] = eng.page_words(self.doc[i])
        return self.words_cache[i]

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
        if not eng.page_has_text(page):
            c.create_text(self.offset[0] + pix.width / 2, self.offset[1] + 40, fill=BREAK_COLOR,
                          font=("Segoe UI", 12, "bold"), tags="overlay",
                          text="No readable text on this page (scanned image) - it will convert as empty.")
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
        bs = eng.auto_detect_breaks(self.words(self.page_index), self._page_area())
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
        self.schedule_preview(0)

    def _sync_option_vars(self):
        o = self.layout.options
        self.join_var.set(o.join_wrapped)
        self.header_var.set(o.skip_repeated_header)
        self.numbers_var.set(o.convert_numbers)

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

    def update_preview(self):
        self._preview_job = None
        if not self.doc:
            return
        try:
            rows = eng.page_rows(self.doc[self.page_index], self.layout, self.page_index)
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
            self.tree.insert("", "end", values=vals)
        more = f"  (first 500 shown)" if len(rows) > 500 else ""
        self.preview_note.configure(text=f"{len(rows)} row(s) x {ncols} column(s) on this page{more}")

    # ------------------------------------------------------------------
    # Convert
    # ------------------------------------------------------------------
    def convert(self):
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
        self.convert_btn.configure(state="disabled")
        self.progress.configure(maximum=len(pages), value=0)

        def progress(n, total):
            self.progress.configure(value=n)
            self.status_var.set(f"Reading page {n} of {total}...")
            self.update_idletasks()

        try:
            rows = eng.extract(self.doc, self.layout, pages, progress)
            self.status_var.set("Writing the workbook...")
            self.update_idletasks()
            n = eng.write_workbook(out, rows, self.layout.options,
                                   one_sheet=self.sheet_var.get() == "one",
                                   page_column=self.pagecol_var.get(),
                                   source_name=Path(self.pdf_path).name)
        except PermissionError:
            messagebox.showerror(APP_TITLE, f"Couldn't save to\n{out}\n\nIs it open in Excel? Close it and try again.")
            self.status_var.set("Not saved.")
            return
        except Exception as e:
            log.error("convert failed: %s\n%s", e, traceback.format_exc())
            messagebox.showerror(APP_TITLE, f"The conversion failed.\n\n{e}\n\nDetails were written to\n{LOG_PATH}")
            self.status_var.set("Conversion failed.")
            return
        finally:
            self.convert_btn.configure(state="normal")
        empty = [i + 1 for i in pages if not rows.get(i)]
        msg = f"Saved {n:,} row(s) from {len(pages)} page(s) to\n{out}"
        if empty:
            msg += f"\n\nNo text was found in the table area on page(s): {', '.join(map(str, empty[:20]))}{'...' if len(empty) > 20 else ''}"
        msg += "\n\nCheck the totals against the PDF before relying on the conversion."
        self.status_var.set(f"Saved {out}")
        if messagebox.askyesno(APP_TITLE, msg + "\n\nOpen the workbook now?"):
            _open_file(out)

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
    app = App()
    if len(sys.argv) > 1 and Path(sys.argv[1]).is_file():
        app.after(100, lambda: app.open_pdf(sys.argv[1]))
    app.mainloop()


if __name__ == "__main__":
    main()
