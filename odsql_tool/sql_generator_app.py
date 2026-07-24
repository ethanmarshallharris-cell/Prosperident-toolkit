"""
Open Dental Security Log SQL Generator
----------------------------------------
Prosperident internal tool.

Lets you pick which PermTypes to include (checkboxes), pick a date
range, and generates a ready-to-run SQL query against Open Dental's
securitylog table. The PermType column in the generated query is
translated from its numeric code to its readable name via a CASE
expression, instead of showing a raw number.

Run: SQL Generator.bat  (double-click)
Or:  python sql_generator_app.py
"""

import json
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date

import refresh_permtypes
import update_check
import resource_paths
import branding

# Bumped whenever this tool's own code changes in a way worth being able
# to confirm at a glance (separate from permtypes.json's own currency,
# which is tracked and refreshed independently). Compared against
# update_check.VERSION_MANIFEST_URL's "odsql_tool" entry on startup.
APP_VERSION = "2026-07-24"

# app_dir()/ensure_seeded() resolve correctly whether this is running as
# a plain script (files live next to this .py, same as always) or as a
# PyInstaller-built standalone .exe/.app (files live next to the built
# executable, seeded from the copy bundled into the build on first run).
APP_DIR = resource_paths.app_dir()
PERMTYPES_FILE = resource_paths.ensure_seeded("permtypes.json")
FAVORITES_FILE = APP_DIR / "favorites.json"
DETAILS_FILE = resource_paths.ensure_seeded("permtype_details.json")

# The PermTypes selected in the original example query -- checked by
# default so the tool is useful immediately on first launch.
DEFAULT_SELECTED = {10, 15, 16, 17, 18, 37, 44, 49, 89, 202}


def load_permtypes():
    if not PERMTYPES_FILE.exists():
        messagebox.showerror(
            "PermTypes file missing",
            f"Could not find {PERMTYPES_FILE.name}.\n\n"
            "Could not reach the Open Dental documentation site on "
            "startup either. Check your internet connection and "
            "restart the app, or run 'Refresh PermTypes.bat'."
        )
        raise SystemExit(1)
    data = json.loads(PERMTYPES_FILE.read_text())
    # returns list of (id:int, name:str) sorted by id
    return sorted(((int(k), v) for k, v in data.items()), key=lambda x: x[0])


def load_details():
    if not DETAILS_FILE.exists():
        return {}
    try:
        return {int(k): v for k, v in json.loads(DETAILS_FILE.read_text()).items()}
    except Exception:
        return {}


def load_favorites():
    if not FAVORITES_FILE.exists():
        return set()
    try:
        return set(int(i) for i in json.loads(FAVORITES_FILE.read_text()))
    except Exception:
        return set()


def save_favorites(favorite_ids):
    FAVORITES_FILE.write_text(json.dumps(sorted(favorite_ids)))


def auto_refresh_permtypes(timeout=6):
    """
    Best-effort re-scan of the Open Dental documentation page, run
    automatically at startup so the PermType list and summaries stay
    current without the user having to remember to click Refresh.
    Any failure (no internet, site unreachable, page format changed,
    or the request just taking too long) is swallowed silently -- the
    app just keeps using whatever is already cached in permtypes.json
    / permtype_details.json. Returns True if the cache files were
    updated, False otherwise.
    """
    try:
        html = refresh_permtypes.fetch_html(refresh_permtypes.DOC_URL, timeout=timeout)
        permtypes = refresh_permtypes.extract_permtypes(html)
    except Exception:
        return False

    sorted_ids = sorted(int(k) for k in permtypes.keys())
    sorted_names = {str(i): permtypes[str(i)]["name"] for i in sorted_ids}
    sorted_details = {
        str(i): {"name": permtypes[str(i)]["name"], "desc": permtypes[str(i)]["desc"]}
        for i in sorted_ids
    }
    try:
        PERMTYPES_FILE.write_text(json.dumps(sorted_names, indent=2))
        DETAILS_FILE.write_text(json.dumps(sorted_details, indent=2))
    except Exception:
        return False
    return True


class PermTypeRow:
    def __init__(self, parent, perm_id, name, checked, favorite, on_toggle,
                 on_favorite_toggle, on_show_summary=None):
        self.perm_id = perm_id
        self.name = name
        self.var = tk.BooleanVar(value=checked)
        self.fav_var = tk.BooleanVar(value=favorite)
        self.frame = ttk.Frame(parent)
        self.chk = ttk.Checkbutton(
            self.frame, variable=self.var, command=on_toggle
        )
        self.chk.pack(side="left")

        self.fav_btn = tk.Checkbutton(
            self.frame,
            variable=self.fav_var,
            onvalue=True,
            offvalue=False,
            indicatoron=False,
            text="★" if favorite else "☆",
            width=2,
            relief="flat",
            font=("Segoe UI", 10),
            fg="#c9a227",
            command=lambda: self._toggle_favorite(on_favorite_toggle),
        )
        self.fav_btn.pack(side="left", padx=(2, 4))

        label_text = f"{perm_id:>4}  {name}"
        self.label = ttk.Label(self.frame, text=label_text, anchor="w")
        self.label.pack(side="left", fill="x", expand=True)

        if on_show_summary is not None:
            # Clicking the name shows its summary/description from the
            # Open Dental documentation site (cached locally).
            self.label.configure(cursor="hand2")
            self.label.bind("<Button-1>", lambda e: on_show_summary(self.perm_id, self.name))

    def _toggle_favorite(self, on_favorite_toggle):
        is_fav = self.fav_var.get()
        self.fav_btn.configure(text="★" if is_fav else "☆")
        on_favorite_toggle(self.perm_id, is_fav)

    def matches_filter(self, text):
        text = text.lower()
        return text in self.name.lower() or text in str(self.perm_id)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Open Dental Security Log SQL Generator - Prosperident  (v{APP_VERSION})")
        self.geometry("900x750")
        self.minsize(760, 600)

        branding.apply_theme(self)
        branding.add_header(
            self, "Open Dental Security Log SQL Generator",
            "Prosperident internal tool",
        )

        self.permtypes = load_permtypes()
        self.details = load_details()
        self.favorites = load_favorites()
        self.rows = []
        self.favorites_only_var = tk.BooleanVar(value=False)
        self.refresh_status_var = tk.StringVar(value="Checking for PermType updates...")
        self.app_update_var = tk.StringVar(value="")

        self._build_ui()
        self._populate_rows()
        self._start_background_refresh()
        self._start_app_update_check()

    # ---------- UI construction ----------

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        self.header_count_var = tk.StringVar(
            value=f"{len(self.permtypes)} PermTypes loaded from permtypes.json. "
                  "Click a PermType's name to see its summary."
        )
        ttk.Label(top, textvariable=self.header_count_var, foreground="#555").pack(anchor="w")
        ttk.Label(top, textvariable=self.refresh_status_var, foreground="#888").pack(anchor="w")
        self.app_update_label = ttk.Label(
            top, textvariable=self.app_update_var, foreground="#0645AD", cursor="hand2",
        )
        self.app_update_label.pack(anchor="w")
        self.app_update_label.bind("<Button-1>", self._open_app_update_link)
        self._app_update_download_url = None

        # --- Date range ---
        date_frame = ttk.LabelFrame(self, text="Date Range", padding=10)
        date_frame.pack(fill="x", padx=10, pady=(5, 5))

        ttk.Label(date_frame, text="Start date (YYYY-MM-DD):").grid(
            row=0, column=0, sticky="w", padx=(0, 5)
        )
        self.start_date_var = tk.StringVar(value=str(date.today().replace(day=1)))
        ttk.Entry(date_frame, textvariable=self.start_date_var, width=15).grid(
            row=0, column=1, sticky="w"
        )

        ttk.Label(date_frame, text="End date (YYYY-MM-DD):").grid(
            row=0, column=2, sticky="w", padx=(20, 5)
        )
        self.end_date_var = tk.StringVar(value=str(date.today()))
        ttk.Entry(date_frame, textvariable=self.end_date_var, width=15).grid(
            row=0, column=3, sticky="w"
        )

        # --- PermType picker ---
        perm_frame = ttk.LabelFrame(
            self, text="PermTypes to include (check the ones you want)", padding=10
        )
        perm_frame.pack(fill="both", expand=True, padx=10, pady=5)

        controls = ttk.Frame(perm_frame)
        controls.pack(fill="x", pady=(0, 5))

        ttk.Label(controls, text="Filter:").pack(side="left")
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *a: self._apply_filter())
        ttk.Entry(controls, textvariable=self.filter_var, width=30).pack(
            side="left", padx=5
        )

        ttk.Button(controls, text="Select All (visible)", command=self._select_all).pack(
            side="left", padx=5
        )
        ttk.Button(controls, text="Clear All (visible)", command=self._clear_all).pack(
            side="left", padx=5
        )
        ttk.Button(
            controls, text="Restore Defaults", command=self._restore_defaults
        ).pack(side="left", padx=5)

        ttk.Checkbutton(
            controls,
            text="★ Favorites only",
            variable=self.favorites_only_var,
            command=self._apply_filter,
        ).pack(side="left", padx=5)

        self.selected_count_var = tk.StringVar()
        ttk.Label(controls, textvariable=self.selected_count_var, foreground="#555").pack(
            side="right"
        )

        # Scrollable checklist
        canvas_frame = ttk.Frame(perm_frame)
        canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(canvas_frame, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.list_frame = ttk.Frame(self.canvas)

        self.list_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # --- Generate button ---
        gen_frame = ttk.Frame(self, padding=10)
        gen_frame.pack(fill="x")
        ttk.Button(
            gen_frame, text="Generate SQL", command=self._generate_sql
        ).pack(side="left")
        ttk.Button(
            gen_frame, text="Copy to Clipboard", command=self._copy_sql
        ).pack(side="left", padx=5)
        ttk.Button(
            gen_frame, text="Save as .sql...", command=self._save_sql
        ).pack(side="left", padx=5)

        # --- Output ---
        out_frame = ttk.LabelFrame(self, text="Generated SQL", padding=10)
        out_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.output_text = tk.Text(out_frame, wrap="none", height=15, font=("Consolas", 10))
        out_scroll_y = ttk.Scrollbar(out_frame, orient="vertical", command=self.output_text.yview)
        out_scroll_x = ttk.Scrollbar(out_frame, orient="horizontal", command=self.output_text.xview)
        self.output_text.configure(yscrollcommand=out_scroll_y.set, xscrollcommand=out_scroll_x.set)

        self.output_text.grid(row=0, column=0, sticky="nsew")
        out_scroll_y.grid(row=0, column=1, sticky="ns")
        out_scroll_x.grid(row=1, column=0, sticky="ew")
        out_frame.rowconfigure(0, weight=1)
        out_frame.columnconfigure(0, weight=1)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _populate_rows(self):
        # Favorites first (alphabetical), then everything else by ID.
        ordered = sorted(
            self.permtypes,
            key=lambda pt: (pt[0] not in self.favorites, pt[0]),
        )
        for perm_id, name in ordered:
            checked = perm_id in DEFAULT_SELECTED
            favorite = perm_id in self.favorites
            row = PermTypeRow(
                self.list_frame, perm_id, name, checked, favorite,
                self._update_count, self._on_favorite_toggle, self._show_summary,
            )
            self.rows.append(row)
        self._layout_rows()
        self._update_count()

    def _on_favorite_toggle(self, perm_id, is_favorite):
        if is_favorite:
            self.favorites.add(perm_id)
        else:
            self.favorites.discard(perm_id)
        save_favorites(self.favorites)
        # Re-sort so favorites float to the top immediately.
        self.rows.sort(key=lambda r: (r.perm_id not in self.favorites, r.perm_id))
        self._layout_rows()

    def _show_summary(self, perm_id, name):
        detail = self.details.get(perm_id)
        desc = (detail or {}).get("desc", "").strip()
        if not desc:
            desc = ("No cached summary for this PermType yet. Click "
                    "'Refresh PermTypes.bat' (or wait for the automatic "
                    "startup refresh to finish) to pull it from the Open "
                    "Dental documentation site.")

        win = tk.Toplevel(self)
        win.title(f"PermType {perm_id} - {name}")
        win.geometry("480x260")
        win.transient(self)

        frame = ttk.Frame(win, padding=15)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame, text=f"{name}  (PermType {perm_id})",
            font=("Segoe UI", 11, "bold"), wraplength=440, justify="left",
        ).pack(anchor="w", pady=(0, 10))

        text = tk.Text(frame, wrap="word", height=9, font=("Segoe UI", 10))
        text.insert("1.0", desc)
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)

        ttk.Button(frame, text="Close", command=win.destroy).pack(anchor="e", pady=(10, 0))

    # ---------- Auto-refresh on startup ----------

    def _start_background_refresh(self):
        thread = threading.Thread(target=self._background_refresh_worker, daemon=True)
        thread.start()

    # ---------- App-version update check (separate from PermTypes) ----------

    def _start_app_update_check(self):
        thread = threading.Thread(target=self._app_update_check_worker, daemon=True)
        thread.start()

    def _app_update_check_worker(self):
        result = update_check.check_for_update("odsql_tool", APP_VERSION)
        if result:
            self.after(0, lambda: self._show_app_update(result))

    def _show_app_update(self, entry):
        notes = entry.get("notes") or ""
        text = f"A newer version of this tool is available (v{entry.get('version')}) — click to get it."
        if notes:
            text += f"  {notes}"
        self._app_update_download_url = entry.get("download_url")
        self.app_update_var.set(text)

    def _open_app_update_link(self, event=None):
        if self._app_update_download_url:
            webbrowser.open(self._app_update_download_url)

    def _background_refresh_worker(self):
        updated = auto_refresh_permtypes()
        self.after(0, lambda: self._on_refresh_complete(updated))

    def _on_refresh_complete(self, updated):
        if not updated:
            self.refresh_status_var.set(
                "Using cached PermType list (couldn't reach the Open Dental "
                "documentation site just now)."
            )
            return

        new_permtypes = load_permtypes()
        new_details = load_details()
        if new_permtypes == self.permtypes:
            self.refresh_status_var.set("PermType list is up to date.")
            self.details = new_details
            return

        # The list changed -- rebuild the checklist in place, preserving
        # the user's current checkbox and favorite selections by ID.
        checked_ids = {r.perm_id for r in self.rows if r.var.get()}
        for r in self.rows:
            r.frame.destroy()
        self.rows = []
        self.permtypes = new_permtypes
        self.details = new_details

        ordered = sorted(
            self.permtypes,
            key=lambda pt: (pt[0] not in self.favorites, pt[0]),
        )
        for perm_id, name in ordered:
            checked = perm_id in checked_ids if checked_ids else perm_id in DEFAULT_SELECTED
            favorite = perm_id in self.favorites
            row = PermTypeRow(
                self.list_frame, perm_id, name, checked, favorite,
                self._update_count, self._on_favorite_toggle, self._show_summary,
            )
            self.rows.append(row)
        self._apply_filter()
        self._update_count()
        self.header_count_var.set(
            f"{len(self.permtypes)} PermTypes loaded from permtypes.json. "
            "Click a PermType's name to see its summary."
        )
        self.refresh_status_var.set("PermType list updated from the Open Dental documentation site.")

    def _layout_rows(self):
        # Grid the (currently) visible rows into two columns for compactness.
        visible = [r for r in self.rows if r.frame.winfo_ismapped() or True]
        for r in self.rows:
            r.frame.grid_forget()
        col_count = 2
        r_i = c_i = 0
        for row in self.rows:
            if not row.frame.winfo_exists():
                continue
            if getattr(row, "_hidden", False):
                continue
            row.frame.grid(row=r_i, column=c_i, sticky="w", padx=5, pady=1)
            c_i += 1
            if c_i >= col_count:
                c_i = 0
                r_i += 1

    def _apply_filter(self):
        text = self.filter_var.get()
        fav_only = self.favorites_only_var.get()
        for row in self.rows:
            match = row.matches_filter(text)
            if fav_only and row.perm_id not in self.favorites:
                match = False
            row._hidden = not match
        self._layout_rows()

    def _select_all(self):
        for row in self.rows:
            if not getattr(row, "_hidden", False):
                row.var.set(True)
        self._update_count()

    def _clear_all(self):
        for row in self.rows:
            if not getattr(row, "_hidden", False):
                row.var.set(False)
        self._update_count()

    def _restore_defaults(self):
        for row in self.rows:
            row.var.set(row.perm_id in DEFAULT_SELECTED)
        self._update_count()

    def _update_count(self):
        n = sum(1 for r in self.rows if r.var.get())
        self.selected_count_var.set(f"{n} selected")

    def _selected_permtypes(self):
        return [(r.perm_id, r.name) for r in self.rows if r.var.get()]

    # ---------- SQL generation ----------

    def _generate_sql(self):
        selected = self._selected_permtypes()
        if not selected:
            messagebox.showwarning("No PermTypes selected", "Check at least one PermType first.")
            return

        start_date = self.start_date_var.get().strip()
        end_date = self.end_date_var.get().strip()
        if not start_date or not end_date:
            messagebox.showwarning("Missing dates", "Enter both a start and end date.")
            return

        selected_sorted = sorted(selected, key=lambda x: x[0])
        ids_csv = ", ".join(str(i) for i, _ in selected_sorted)

        # CASE expression translates PermType number -> readable name
        case_lines = "\n".join(
            f"        WHEN {pid} THEN '{name}'" for pid, name in selected_sorted
        )

        sql = f"""Select
    securitylog.SecurityLogNum As LogNumber,
    CASE securitylog.PermType
{case_lines}
        ELSE CONVERT(VARCHAR(10), securitylog.PermType)
    END As PermType,
    userod.UserName As UserName,
    securitylog.LogDateTime As LogDate,
    securitylog.LogText As LogText,
    securitylog.CompName As Computer,
    patient.LName As PatientLastName,
    patient.FName As PatientFirstName
From
    securitylog Inner Join
    userod On userod.UserNum = securitylog.UserNum Inner Join
    patient On patient.PatNum = securitylog.PatNum
Where securitylog.PermType in ({ids_csv})
    AND securitylog.LogDateTime >= '{start_date}'
    AND securitylog.LogDateTime <= '{end_date}'
Order By
    LogDate
"""
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", sql)

    def _copy_sql(self):
        sql = self.output_text.get("1.0", "end").strip()
        if not sql:
            messagebox.showinfo("Nothing to copy", "Generate the SQL first.")
            return
        self.clipboard_clear()
        self.clipboard_append(sql)
        messagebox.showinfo("Copied", "SQL copied to clipboard.")

    def _save_sql(self):
        sql = self.output_text.get("1.0", "end").strip()
        if not sql:
            messagebox.showinfo("Nothing to save", "Generate the SQL first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".sql",
            filetypes=[("SQL files", "*.sql"), ("All files", "*.*")],
            initialfile="securitylog_query.sql",
        )
        if path:
            Path(path).write_text(sql)
            messagebox.showinfo("Saved", f"Saved to {path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
