OneNote Page Export  (Prosperident Tool Suite)
==============================================

Exports the OneNote page that is open on screen, together with ALL of its
subpages, to a single Word document. Each page title becomes a Word heading
(page = Heading 1, subpage = Heading 2, sub-subpage = Heading 3), with a
table of contents when there is more than one page.

Requirements: Windows, the OneNote desktop app (not the Windows 10 "OneNote
for Windows" store app) and Microsoft Word.

SETUP (once)
  1. Unzip this folder somewhere permanent, for example Documents.
  2. Double-click "Setup Hotkey.bat". It creates a Start menu shortcut with the
     Ctrl+Alt+E hotkey. (Re-run it if you ever move this folder.)

USE
  1. In OneNote, open the top-level page you want (its subpages come along
     automatically).
  2. Press Ctrl+Alt+E (or double-click "Export OneNote Page.bat").
  3. Choose where to save. The document opens in Word when finished.

Settings (page breaks, table of contents, removing date stamps, opening when
done, update check) are at the top of Export-OneNotePage.ps1.

Problems are logged to Export-OneNotePage.log in this folder.

UPDATES
  After an export, the tool tells you once if a newer version is available.
  Download it from www.prosperident.com/tool-suite and unzip it over this
  folder, so the hotkey keeps working.
