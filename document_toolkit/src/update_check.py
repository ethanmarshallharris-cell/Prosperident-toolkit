"""
Shared "is a newer version available" check for Prosperident internal
tools.

Standard library only (urllib), so it doesn't add anything to what
build_windows.bat / PyInstaller need to package. Designed to fail
completely silently on anything except success -- offline, the shared
link being briefly down, a firewall blocking it, malformed JSON,
whatever -- because a background version check should never be the
reason a tool hangs, errors, or nags someone who's just trying to get
their work done.

How it works
------------
Prosperident keeps one small JSON file (VERSION_MANIFEST_URL below) in
the same shared location as the tools themselves. It looks like:

    {
      "document_toolkit": {
        "version": "2026-07-24",
        "download_url": "https://.../prosperident_tool_suite.html",
        "notes": "Adds support for a second document type."
      },
      "odsql_tool": {
        "version": "2026-07-24",
        "download_url": "https://.../prosperident_tool_suite.html",
        "notes": ""
      }
    }

Each tool calls check_for_update(tool_key, current_version) on a
background thread at startup. Versions are plain date strings
(YYYY-MM-DD, matching each tool's own build-date versioning) so "is it
newer" is just a string comparison -- no version-parsing library needed.

Updating the manifest is the one step that makes new versions visible:
whenever Ethan ships an updated tool, editing version.json's "version"
field (and re-uploading the updated suite page) is what tells already-
downloaded copies an update exists.
"""

import json
import urllib.request

# Shareable, direct-download link to version.json. This is the ONE
# thing that needs to be filled in once the manifest has a real home
# (e.g. a OneDrive "anyone with the link -> direct download" link).
# Until it's set, every check below fails harmlessly and silently, same
# as being offline.
VERSION_MANIFEST_URL = "https://prosperident27-my.sharepoint.com/:u:/g/personal/ethan_marshallharris_prosperident_com/IQBSbzaRNlEDRYMwIzTqut5PAd0zry4cWn4uGrpgz4XNb_8?download=1"

REQUEST_TIMEOUT_SECONDS = 4


def check_for_update(tool_key, current_version):
    """Returns a dict {"version": ..., "download_url": ..., "notes": ...}
    if the manifest lists a version newer than current_version for
    tool_key, or None if there's no update, the check failed for any
    reason, or the manifest URL hasn't been configured yet.

    Safe to call from a background thread -- does no UI work itself.
    """
    if "REPLACE-WITH-VERSION-MANIFEST-URL" in VERSION_MANIFEST_URL:
        return None
    try:
        with urllib.request.urlopen(VERSION_MANIFEST_URL, timeout=REQUEST_TIMEOUT_SECONDS) as resp:
            manifest = json.loads(resp.read().decode("utf-8"))
        entry = manifest.get(tool_key)
        if not entry:
            return None
        latest_version = entry.get("version", "")
        if latest_version > current_version:
            return entry
        return None
    except Exception:
        # Deliberately broad: any failure here (network, DNS, timeout,
        # bad JSON, missing key) should behave exactly like "no update
        # available right now" -- never an error the user sees.
        return None
