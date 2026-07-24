Open Dental Security Log SQL Generator
========================================
Prosperident internal tool.

WHAT'S IN THIS FOLDER
----------------------
- SQL Generator.bat        <- Double-click this to run the app.
- Refresh PermTypes.bat    <- Double-click to re-scan the Open Dental
                               documentation and update the PermType list.
- sql_generator_app.py     Main application (Tkinter GUI).
- refresh_permtypes.py     Scraper that rebuilds permtypes.json and
                            permtype_details.json.
- permtypes.json           Cached PermType ID -> Name list (276 entries,
                            current as of Open Dental documentation
                            version 26.1, pulled July 2026).
- permtype_details.json    Cached PermType ID -> summary/description
                            text (created the first time a refresh
                            succeeds -- see SUMMARIES below).
- favorites.json            Your starred PermTypes (created automatically
                            the first time you favorite something).

REQUIREMENTS
------------
Python 3.8 or later, installed on your Windows machine, with "python"
available on your PATH (Tkinter ships with the standard Windows
installer, so no extra install is needed).
Download: https://www.python.org/downloads/

HOW TO USE
----------
1. Double-click "SQL Generator.bat".
2. Check the PermTypes you want in the query (use the Filter box to
   search by name or number). Ten PermTypes matching your original
   example query are pre-checked.
3. Set the start/end date.
4. Click "Generate SQL". The PermType column in the output uses a
   CASE expression so it shows the PermType's NAME instead of its
   raw number.
5. Copy to clipboard or save as a .sql file.

FAVORITES
---------
Click the star (☆/★) next to any PermType to mark it as a favorite.
Favorites automatically float to the top of the list, and the
"★ Favorites only" checkbox next to the Filter box lets you narrow
the list down to just your favorites. Your favorites are saved to
favorites.json in this folder, so they persist between sessions.

SUMMARIES (click a PermType to see its description)
------------------------------------------------------
Click directly on a PermType's NAME (not the checkbox or star) to
open a small window showing its summary/description exactly as
written on the Open Dental documentation site -- e.g. what triggers
a log entry for that permission, or notes like "DEPRECATED - Uses
date restrictions." This is pulled from permtype_details.json. If
you click a PermType before that file has been populated yet (very
first run, before the startup refresh finishes, or if it's offline),
the popup will say so and suggest refreshing.

KEEPING PERMTYPES CURRENT
--------------------------
Open Dental adds new PermTypes fairly often. The app now handles this
automatically: every time you double-click "SQL Generator.bat", it
launches immediately using the cached list, then quietly checks the
Open Dental documentation site in the background (a few seconds) and
refreshes permtypes.json / permtype_details.json if anything changed
-- you'll see a small status line under the title update to say so.
Your checked boxes and favorites are preserved across that refresh.

If you're offline or the site is briefly unreachable, the app just
keeps using whatever was cached from last time -- no error, no delay.

You can still force a manual refresh any time by double-clicking
"Refresh PermTypes.bat".

If a refresh ever returns an error, Open Dental has likely bumped
their documentation URL to a newer version number (e.g. from 26-1 to
26-2). Open DOC_URL at the top of refresh_permtypes.py and update the
version number in the URL to match the current one shown at
https://www.opendental.com/manual/opendentaldocumentation.html
