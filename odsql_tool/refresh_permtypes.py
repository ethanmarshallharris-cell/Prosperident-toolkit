"""
refresh_permtypes.py
---------------------
Re-scans the Open Dental online documentation page for the securitylog
table and rebuilds:
  - permtypes.json         PermType ID -> Name (used to populate the
                            checklist and the generated SQL's CASE
                            expression)
  - permtype_details.json  PermType ID -> {name, description} (used by
                            the "click a PermType to see its summary"
                            feature in the app)

Run this whenever Open Dental updates their permission types (they add
new ones with most new releases).

Usage (double-click):
    Refresh PermTypes.bat

Usage (command line):
    python refresh_permtypes.py

Notes:
    - This pulls from the public Open Dental documentation page. The
      page occasionally changes its exact version number in the URL
      (e.g. OpenDentalDocumentation26-1.xml). If a refresh comes back
      with zero results, open DOC_URL below in a browser and update it
      to whatever the current version URL is.
    - The documentation site's robots.txt blocks generic crawlers, so
      this script identifies itself honestly and is meant to be run
      occasionally, by hand, for your own internal use -- not on an
      automated schedule against their servers.
"""

import html as html_lib
import json
import re
import sys
import urllib.request

import resource_paths

DOC_URL = "https://www.opendental.com/OpenDentalDocumentation26-1.xml"
# app_dir(), not Path(__file__).parent -- when this module is running
# inside a PyInstaller-built .exe/.app, __file__ points into the
# temporary extraction folder, which disappears after the run. app_dir()
# resolves to wherever the actual executable lives instead, so a refresh
# persists across launches the same way it does for the plain-script version.
OUTPUT_FILE = resource_paths.app_dir() / "permtypes.json"
DETAILS_FILE = resource_paths.app_dir() / "permtype_details.json"

HEADERS = {
    # Identify honestly as a normal browser-ish fetch; this is a manual,
    # occasional, single-request run by a human, not a crawler.
    "User-Agent": "Mozilla/5.0 (compatible; Prosperident-OD-Tool/1.0)"
}

# Each PermType is documented as its own table cell in the securitylog
# schema table:  <td width="650">Name: Num-Description text</td>
CELL_RE = re.compile(r'<td width="650">(.*?)</td>', re.S)
ENTRY_RE = re.compile(r"^([A-Za-z][A-Za-z0-9_]*)\s*:\s*(\d+)\s*-?\s*(.*)$", re.S)
TAG_RE = re.compile(r"<[^>]+>")


def fetch_html(url: str, timeout: int = 30) -> str:
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", errors="replace")


def clean_text(raw: str) -> str:
    text = TAG_RE.sub(" ", raw)
    text = html_lib.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def extract_permtypes(html: str) -> dict:
    # Narrow down to the PermType enum block: it starts right after
    # the literal marker "Enum:EnumPermType" and runs until the next
    # numbered field definition in the securitylog table.
    start = html.find("Enum:EnumPermType")
    if start == -1:
        raise RuntimeError(
            "Could not find the PermType enum block on the page. "
            "The documentation page structure or URL may have changed - "
            "see the notes at the top of this script."
        )

    # Enum blocks run a few hundred rows; a generous chunk comfortably
    # covers all of them without spilling into the next field.
    chunk = html[start:start + 120000]

    results = {}
    for cell_match in CELL_RE.finditer(chunk):
        cell_text = clean_text(cell_match.group(1))
        entry_match = ENTRY_RE.match(cell_text)
        if not entry_match:
            continue
        name, num, desc = entry_match.group(1), entry_match.group(2), entry_match.group(3)
        # Skip obvious false positives by requiring a plausible PermType
        # number (they stay under 2000).
        if int(num) > 2000:
            continue
        results[num] = {"name": name, "desc": desc.strip()}

    if not results:
        raise RuntimeError("Found the enum marker but parsed zero PermTypes.")

    return results


def main():
    print(f"Fetching {DOC_URL} ...")
    try:
        html = fetch_html(DOC_URL)
    except Exception as e:
        print(f"ERROR: could not download the documentation page: {e}")
        print("Check your internet connection, or update DOC_URL in this "
              "script if Open Dental has moved to a newer version number.")
        sys.exit(1)

    try:
        permtypes = extract_permtypes(html)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    # Sort numerically by ID for clean, readable JSON files.
    sorted_ids = sorted(int(k) for k in permtypes.keys())
    sorted_names = {str(i): permtypes[str(i)]["name"] for i in sorted_ids}
    sorted_details = {
        str(i): {"name": permtypes[str(i)]["name"], "desc": permtypes[str(i)]["desc"]}
        for i in sorted_ids
    }

    old_count = 0
    if OUTPUT_FILE.exists():
        try:
            old_count = len(json.loads(OUTPUT_FILE.read_text()))
        except Exception:
            pass

    OUTPUT_FILE.write_text(json.dumps(sorted_names, indent=2))
    DETAILS_FILE.write_text(json.dumps(sorted_details, indent=2))
    print(f"Saved {len(sorted_names)} PermTypes to {OUTPUT_FILE.name} "
          f"(previously had {old_count}).")
    print(f"Saved descriptions to {DETAILS_FILE.name}.")
    print("Done. You can now run the SQL Generator app.")


if __name__ == "__main__":
    main()
