"""
Shared helper for finding data files (permtypes.json, favorites.json,
etc.) correctly whether this tool is running as a plain .py script or as
a PyInstaller-built standalone executable.

Why this is needed: a PyInstaller --onefile build unpacks itself into a
temporary folder every time it runs (sys._MEIPASS) -- writing there is
pointless since it's deleted afterward, and reading permtypes.json from
the *source* folder doesn't work anymore either because there is no
source folder on a machine that only has the .exe. So:

  - Read-only defaults that ship with the build (the permtypes.json this
    tool was built with) live bundled inside the executable and get
    copied out ONCE, the first time the exe runs, to sit next to it.
  - Everything after that -- refreshes, favorites -- reads and writes
    that copy next to the executable, exactly like the plain-script
    version already does with files next to sql_generator_app.py.

This means the first launch of a fresh install is seeded from whatever
permtypes.json was current when it was built, and immediately becomes a
normal, independently-updatable file after that.
"""
import shutil
import sys
from pathlib import Path


def is_frozen():
    return getattr(sys, "frozen", False)


def app_dir():
    """Folder the writable/live copies of data files should live in --
    next to the .exe when frozen, next to this script otherwise."""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def bundled_dir():
    """Folder holding the read-only defaults bundled into the build --
    PyInstaller's extraction dir when frozen, same as app_dir() otherwise."""
    if is_frozen() and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def ensure_seeded(filename):
    """Returns the writable path for `filename` in app_dir(), copying the
    bundled default there first if nothing's there yet. Safe to call
    every startup -- a no-op once the file exists."""
    dest = app_dir() / filename
    if not dest.exists():
        src = bundled_dir() / filename
        if src.exists():
            try:
                shutil.copyfile(src, dest)
            except OSError:
                pass
    return dest
