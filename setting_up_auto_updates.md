# Setting Up Auto-Update Checks — Ethan's One-Time Setup

Both tools check a small shared `version.json` file each time they launch, and tell
whoever's running them (right in the tool's own window, no popups) if a newer version
exists. This is separate from — and in addition to — the SQL Generator's existing
PermType auto-refresh, and separate from the GitHub Actions build pipeline (see the
repo's `README.md` for that) — this piece is just the "is this the newest build" check.

## Where the manifest lives (updated October 2026)

Both tools now read `version.json` straight from the root of this public GitHub
repository (main branch):
`https://raw.githubusercontent.com/ethanmarshallharris-cell/Prosperident-toolkit/main/version.json`

The earlier OneDrive/SharePoint link returned "403 Forbidden" to the tools, because
they fetch it without signing in, so the update check never worked. Steps 2-4 below
(the OneDrive copy and its share link) are no longer needed; they are kept only as a
record of the original setup.

## One-time setup (original)

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
2. Edit `version.json` in the repo root -- the matching entry's `version` (same date as
   step 1) and `notes` (a one-line summary of what changed) -- and commit it. The tools
   read it directly from GitHub; nothing goes to OneDrive. Push this commit only after
   the tagged build in step 3 has published, so nobody is told about a version that
   isn't downloadable yet.
3. Tag and push (`git tag vX.Y.Z && git push origin main --tags`) — GitHub Actions
   builds and publishes the new Windows/Mac apps automatically; the hub page's download
   buttons keep working unchanged since they always point at "latest."

Anyone already running an older copy sees the update notice next time they launch it,
with a link straight to the GitHub release page to grab the new version.
