# Setting Up Auto-Update Checks — Ethan's One-Time Setup

Both tools check a small shared `version.json` file each time they launch, and tell
whoever's running them (right in the tool's own window, no popups) if a newer version
exists. This is separate from — and in addition to — the SQL Generator's existing
PermType auto-refresh, and separate from the GitHub Actions build pipeline (see the
repo's `README.md` for that) — this piece is just the "is this the newest build" check.

## One-time setup

1. **Finish the GitHub setup first** (repo README, steps 1–5) so you have a real
   `<owner>/<repo>` to point everything at.

2. **Put `version.json` (repo root) on OneDrive**, in the same shared folder as the
   tool suite page — e.g. `Prosperident Tools/version.json`. Its `download_url` fields
   already point at `https://github.com/<owner>/<repo>/releases/latest` — replace the
   placeholder with your actual repo path.

3. **Get a direct-download share link for `version.json` itself** (not the hub page —
   the small JSON file). A normal OneDrive "Copy link" gives you a *viewer* link (opens
   a preview page), not a raw file link the tools can read automatically:
   - Right-click `version.json` → **Share** → **Copy link** (view-only is fine).
   - Change `?e=` to `?download=1` at the end (or use OneDrive's "Embed"/API link format
     if that trips you up — Microsoft changes this UI occasionally, search "OneDrive
     direct download link" for the current click-path).
   - Test it: paste the modified link into a browser's address bar. If it downloads/shows
     raw JSON text (not a OneDrive webpage), it's correct.

4. **Set that URL in both tools' source code** — open `update_check.py` (it's identical
   in `document_toolkit/src/` and `odsql_tool/`) and replace:

   ```python
   VERSION_MANIFEST_URL = "https://REPLACE-WITH-VERSION-MANIFEST-URL/version.json"
   ```

   with the link from step 3, in both copies. Commit that change.

5. **Set the same repo path in the hub page** — open `prosperident_tool_suite.html`,
   find `GITHUB_OWNER` / `GITHUB_REPO` near the top of the `<script>` block, and fill in
   your actual GitHub username/org and repo name. Re-upload the page to OneDrive.

6. **Ship a build** (tag + push, per the repo README) so the four download links
   actually resolve to something.

Until steps 4–6 are done, both tools run completely normally — the update check just
silently does nothing (same as being offline), so there's no rush or risk in doing this
incrementally.

## Ongoing: shipping an update

Every time you change either tool:

1. Bump its version string — `APP_BUILD` in `document_toolkit/src/main_gui.py`, or
   `APP_VERSION` in `odsql_tool/sql_generator_app.py` — to today's date.
2. Edit `version.json`'s matching entry (`version`, and `notes` with a one-line summary
   of what changed) and re-upload it to OneDrive.
3. Tag and push (`git tag vX.Y.Z && git push origin main --tags`) — GitHub Actions
   builds and publishes the new Windows/Mac apps automatically; the hub page's download
   buttons keep working unchanged since they always point at "latest."

Anyone already running an older copy sees the update notice next time they launch it,
with a link straight to the GitHub release page to grab the new version.
