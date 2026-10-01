# Prosperident Tools

Source for Prosperident's internal desktop tools, built automatically into
standalone Windows and macOS apps by GitHub Actions.

- `document_toolkit/` — redact/un-redact identifying info in report PDFs
- `odsql_tool/` — Open Dental Security Log SQL Generator

## One-time setup (Ethan)

1. **Create a GitHub account** if you don't have one already: https://github.com/signup
   (the free plan is all this needs).
2. **Create a new repository** — private is fine and recommended, since this is
   internal-only code. Name it whatever you like, e.g. `prosperident-tools`.
3. **Push this folder to it.** From inside this folder:

   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```

   (If you're not comfortable with git on the command line, GitHub Desktop
   — https://desktop.github.com — does the same thing with a normal UI:
   "Add local repository" → point it at this folder → "Publish repository".)

4. **Ship the first build** by creating a version tag and pushing it:

   ```
   git tag v1.0.0
   git push origin v1.0.0
   ```

   This is what actually triggers the build — pushing to `main` alone does
   *not* build/release anything, only pushing a tag starting with `v` does.
   Within a few minutes, check the repo's **Actions** tab to watch it build,
   then the **Releases** tab (right sidebar) for the finished .zip files —
   one Windows and one macOS build for each tool, four total.

5. **Get the stable download links.** Each release asset has a permanent
   URL of the form:

   ```
   https://github.com/<your-username>/<your-repo>/releases/latest/download/<asset-name>
   ```

   The four asset names are always:
   - `DocumentToolkit-Windows.zip`
   - `DocumentToolkit-macOS.zip`
   - `OpenDentalSQLGenerator-Windows.zip`
   - `OpenDentalSQLGenerator-macOS.zip`

   These links always resolve to whatever the *latest* tagged release is —
   you don't need to change them again after shipping future updates, only
   the hub page needs updating once, right now, to point at your repo (see
   `prosperident_tool_suite.html` — the four download buttons have
   placeholder URLs marked `REPLACE-WITH-...` near the top of the `TOOLS`
   array in the page's `<script>` section; swap in your actual repo name).

   If the repo is **private**, these links will prompt for a GitHub sign-in
   before downloading — anyone on staff would need their own free GitHub
   account added as a collaborator (Settings → Collaborators) to download
   without hitting a permission error. If that's more friction than you
   want, make the repo **public** instead — the source code isn't sensitive
   (no credentials or client data live in it), only ship it as public if
   you're comfortable with that reasoning.

## Shipping an update from here on

**The easy way:** double-click `Ship-Update.bat` in this folder after changing either tool. It sets the build date, pushes the change, tags the next version, waits for the build, checks all four downloads and updates `version.json`. The manual steps below are what it automates.


1. Make your code changes.
2. Bump the version — `APP_BUILD` in `document_toolkit/src/main_gui.py`,
   or `APP_VERSION` in `odsql_tool/sql_generator_app.py`.
3. Update `version.json` (see `setting_up_auto_updates.md`) so already-
   installed copies show the in-app update notice.
4. Commit, then tag and push a new version:

   ```
   git add .
   git commit -m "Describe what changed"
   git tag v1.1.0
   git push origin main --tags
   ```

That's the whole release process — the Actions workflow does the rest
(builds both platforms for both tools, publishes the release, and the
stable download links above automatically start serving the new build).

## Testing a build without shipping it

Go to the repo's **Actions** tab → **Build desktop apps** → **Run workflow**.
This builds all four apps the same way a tag-push would, but doesn't create
a public release — the results show up as downloadable "Artifacts" at the
bottom of that run's page, good for testing before you're ready to ship.
