"""
Shared Prosperident visual branding for internal desktop tools.

Colors match www.prosperident.com (navy header, green accent/actions,
Segoe UI as the closest system-standard stand-in for the site's
Raleway/Roboto Slab, since bundling web fonts into a Tkinter app isn't
practical). Two entry points:

  apply_theme(root)          Call once, right after creating the root
                              window, before building any other widgets.
                              Re-themes every ttk widget style used by
                              this app's widgets.

  add_header(root, title, subtitle=None)
                              Call once, right after apply_theme(), and
                              before building the rest of the UI -- packs
                              a navy header bar (logo + title) at the very
                              top of the window and sets the window icon.
                              Everything built afterward should also use
                              pack(side="top", ...) so it lands below the
                              header rather than overlapping it.

Same approach as update_check.py: this file is duplicated as-is into
each tool's folder rather than shared via a package, since these tools
are meant to be simple, dependency-free, single-folder PyInstaller
targets.
"""
import sys
import tkinter as tk
from tkinter import ttk
from pathlib import Path

NAVY = "#1c2331"
NAVY_2 = "#2a3550"
GREEN = "#74c146"
GREEN_DARK = "#5da036"
MAROON = "#a3161f"
GRAY = "#5b6472"
GRAY_LIGHT = "#e7e9ee"
BORDER = "#cfd4e0"
BG = "#f4f5f8"
WHITE = "#ffffff"

LOGO_FILENAME = "prosperident_logo.png"


def _bundled_dir():
    """Where read-only assets bundled into the build (like the logo) can
    be found -- PyInstaller's extraction dir when frozen, next to this
    file otherwise. (Read-only, so no need for the seed-a-writable-copy
    dance update_check.py's sibling, resource_paths.py, does for data
    files that get modified at runtime.)"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def apply_theme(root):
    root.configure(bg=BG)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass  # fall back to whatever theme is already active

    base_font = ("Segoe UI", 10)
    style.configure(".", background=BG, foreground=NAVY, font=base_font)
    style.configure("TFrame", background=BG)
    style.configure("TLabelframe", background=BG, foreground=NAVY, bordercolor=BORDER)
    style.configure("TLabelframe.Label", background=BG, foreground=NAVY, font=("Segoe UI", 10, "bold"))
    style.configure("TLabel", background=BG, foreground=NAVY)
    style.configure("TCheckbutton", background=BG, foreground=NAVY)
    style.configure("TRadiobutton", background=BG, foreground=NAVY)
    style.configure("TPanedwindow", background=BG)

    style.configure(
        "TButton", background=GREEN, foreground="#16210c",
        font=("Segoe UI", 10, "bold"), padding=6, borderwidth=0, focuscolor=GREEN,
    )
    style.map(
        "TButton",
        background=[("disabled", GRAY_LIGHT), ("active", GREEN_DARK), ("pressed", GREEN_DARK)],
        foreground=[("disabled", GRAY)],
    )

    style.configure("TEntry", fieldbackground=WHITE, foreground=NAVY, bordercolor=BORDER, padding=4)
    style.configure("TCombobox", fieldbackground=WHITE, foreground=NAVY, bordercolor=BORDER)
    style.map("TCombobox", fieldbackground=[("readonly", WHITE)])

    style.configure("TNotebook", background=BG, bordercolor=BORDER)
    style.configure(
        "TNotebook.Tab", background=GRAY_LIGHT, foreground=NAVY,
        padding=(14, 8), font=("Segoe UI", 10, "bold"),
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", NAVY)],
        foreground=[("selected", WHITE)],
    )

    style.configure(
        "Treeview", background=WHITE, fieldbackground=WHITE, foreground=NAVY,
        bordercolor=BORDER, rowheight=24,
    )
    style.configure("Treeview.Heading", background=NAVY, foreground=WHITE, font=("Segoe UI", 9, "bold"))
    style.map(
        "Treeview",
        background=[("selected", GREEN)],
        foreground=[("selected", "#16210c")],
    )

    style.configure("TProgressbar", background=GREEN, troughcolor=GRAY_LIGHT, bordercolor=BORDER)
    style.configure("TScrollbar", background=GRAY_LIGHT, troughcolor=BG, bordercolor=BORDER)

    style.configure("Header.TFrame", background=NAVY)
    style.configure("Header.TLabel", background=NAVY, foreground=WHITE, font=("Segoe UI", 14, "bold"))
    style.configure("HeaderSub.TLabel", background=NAVY, foreground="#c8cddb", font=("Segoe UI", 9))
    style.configure("LogoChip.TFrame", background=WHITE)


def add_header(root, title, subtitle=None):
    header = ttk.Frame(root, style="Header.TFrame")
    header.pack(side="top", fill="x")

    logo_path = _bundled_dir() / LOGO_FILENAME
    if logo_path.exists():
        try:
            logo_img = tk.PhotoImage(file=str(logo_path))
            chip = tk.Frame(header, bg=WHITE)
            chip.pack(side="left", padx=(14, 12), pady=10)
            logo_label = tk.Label(chip, image=logo_img, bg=WHITE, bd=0, padx=6, pady=3)
            logo_label.image = logo_img  # keep a reference so it isn't garbage-collected
            logo_label.pack()
            try:
                root.iconphoto(True, logo_img)
            except tk.TclError:
                pass
        except tk.TclError:
            pass

    text_frame = ttk.Frame(header, style="Header.TFrame")
    text_frame.pack(side="left", pady=10)
    ttk.Label(text_frame, text=title, style="Header.TLabel").pack(anchor="w")
    if subtitle:
        ttk.Label(text_frame, text=subtitle, style="HeaderSub.TLabel").pack(anchor="w")

    accent = tk.Frame(root, bg=GREEN, height=3)
    accent.pack(side="top", fill="x")

    return header
