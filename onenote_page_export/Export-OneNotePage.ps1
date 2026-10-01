<#
    Export-OneNotePage.ps1
    Exports the OneNote page that is currently open, plus all of its subpages,
    to a single Word document. Each page title becomes a Word heading
    (parent = Heading 1, subpage = Heading 2, sub-subpage = Heading 3).

    Launch with the Ctrl+Alt+E hotkey (created by Setup Hotkey.bat) or by
    double-clicking "Export OneNote Page.bat".

    Everything is logged to Export-OneNotePage.log in this folder.

    Part of the Prosperident Tool Suite (www.prosperident.com/tool-suite).
    Released through the Prosperident-toolkit repository; Ship-Update.bat
    sets $AppVersion below on each release. After a successful export the
    script checks once for a newer version and offers the download page.
#>

# ======================= Settings (edit as desired) =======================
$PageBreakBetweenPages = $true    # start each OneNote page on a new Word page
$AddTableOfContents    = $true    # add a table of contents (only when 2+ pages)
$RemoveDateStamps      = $true    # remove OneNote's date/time lines under each title
$OpenWhenDone          = $true    # open the finished document in Word
$CheckForUpdates       = $true    # after an export, tell me if a newer version exists
# ==========================================================================

$AppVersion  = "2026-10-01"
$ManifestUrl = 'https://raw.githubusercontent.com/ethanmarshallharris-cell/Prosperident-toolkit/main/version.json'

$ErrorActionPreference = 'Stop'
$ScriptDir  = $PSScriptRoot
$LogFile    = Join-Path $ScriptDir 'Export-OneNotePage.log'
$FolderFile = Join-Path $ScriptDir 'LastFolder.txt'
$NoticeFile = Join-Path $ScriptDir 'UpdateNoticeShown.txt'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Write-Log([string]$Message) {
    $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

# An invisible top-most window so dialogs appear in front of OneNote
function New-TopOwner {
    $f = New-Object System.Windows.Forms.Form
    $f.TopMost       = $true
    $f.ShowInTaskbar = $false
    $f.Opacity       = 0
    $f.StartPosition = 'CenterScreen'
    $f.Size          = New-Object System.Drawing.Size(1, 1)
    $f.Show()
    $f.Activate()
    return $f
}

function Show-Message([string]$Text, [string]$Icon = 'Information') {
    $owner = New-TopOwner
    [System.Windows.Forms.MessageBox]::Show($owner, $Text, 'Export OneNote Page', 'OK', $Icon) | Out-Null
    $owner.Close()
}

function Get-SavePath([string]$DefaultName) {
    $initial = [Environment]::GetFolderPath('MyDocuments')
    if (Test-Path $FolderFile) {
        $saved = (Get-Content -Path $FolderFile -Raw).Trim()
        if ($saved -and (Test-Path $saved)) { $initial = $saved }
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title            = 'Save the exported OneNote pages as'
    $dlg.Filter           = 'Word Document (*.docx)|*.docx'
    $dlg.DefaultExt       = 'docx'
    $dlg.AddExtension     = $true
    $dlg.OverwritePrompt  = $true
    $dlg.InitialDirectory = $initial
    $dlg.FileName         = $DefaultName
    $owner  = New-TopOwner
    $result = $dlg.ShowDialog($owner)
    $owner.Close()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    # Remember the folder so the next dialog opens there
    Set-Content -Path $FolderFile -Value (Split-Path -Parent $dlg.FileName) -Encoding UTF8
    return $dlg.FileName
}

# Once per new version: offer the download page if version.json lists a newer
# build. Any problem (offline, firewall, bad data) is logged and ignored.
function Show-UpdateNotice {
    if (-not $CheckForUpdates) { return }
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $m = Invoke-RestMethod -Uri $ManifestUrl -TimeoutSec 4 -UseBasicParsing
        $e = $m.onenote_page_export
        if (-not $e -or -not $e.version) { return }
        $latest = [string]$e.version
        if ([string]::CompareOrdinal($latest, $AppVersion) -le 0) { return }
        if ((Test-Path $NoticeFile) -and ((Get-Content -Path $NoticeFile -Raw).Trim() -eq $latest)) { return }
        Set-Content -Path $NoticeFile -Value $latest -Encoding UTF8
        Write-Log "Newer version available: $latest (this copy: $AppVersion)."
        $text = "A newer version of OneNote Page Export is available (version $latest; you have $AppVersion)."
        if ($e.notes) { $text += "`n`n" + [string]$e.notes }
        $text += "`n`nOpen the download page now? Unzip the new version over this folder so the Ctrl+Alt+E hotkey keeps working."
        $owner  = New-TopOwner
        $answer = [System.Windows.Forms.MessageBox]::Show($owner, $text, 'OneNote Page Export', 'YesNo', 'Information')
        $owner.Close()
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes -and $e.download_url) { Start-Process ([string]$e.download_url) }
    }
    catch { Write-Log ("Update check skipped: {0}" -f $_.Exception.Message) }
}

function Get-SafeFileName([string]$Name) {
    if (-not $Name -or -not $Name.Trim()) { $Name = 'OneNote Export' }
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $clean = -join ($Name.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '_' } else { $_ } })
    $clean = $clean.Trim().TrimEnd('.')
    if ($clean.Length -gt 100) { $clean = $clean.Substring(0, 100).Trim() }
    if (-not $clean) { $clean = 'OneNote Export' }
    return $clean
}

# Turn the page title into a Word heading and optionally strip OneNote's date/time lines
function Format-PageTitle($Doc, [int]$Start, $Item) {
    $level   = [Math]::Min([Math]::Max($Item.Level, 1), 9)
    $styleId = -1 - $level    # wdStyleHeading1 = -2 ... wdStyleHeading9 = -10
    $title   = $Item.Title.Trim()
    if (-not $title) { $title = 'Untitled Page' }

    # Find the first non-empty paragraph of the inserted page
    $titlePara = $null
    $paras = $Doc.Range($Start, $Doc.Content.End).Paragraphs
    $limit = [Math]::Min($paras.Count, 5)
    for ($i = 1; $i -le $limit; $i++) {
        $t = $paras.Item($i).Range.Text.Trim()
        if ($t) { if ($t -eq $title) { $titlePara = $paras.Item($i) }; break }
    }

    # If OneNote's export did not start with the title, insert one
    if (-not $titlePara) {
        $Doc.Range($Start, $Start).InsertBefore($title + "`r")
        $titlePara = $Doc.Range($Start, $Start).Paragraphs.Item(1)
        Write-Log "    Title line not found in export; inserted heading '$title'."
    }

    $titlePara.Range.Style = $styleId
    $titlePara.Range.Font.Reset()
    $titlePara.Range.ParagraphFormat.Reset()

    if ($RemoveDateStamps) {
        try {
            $p = $titlePara.Next()
            $steps = 0
            while ($p -and $steps -lt 5) {
                $steps++
                $t = $p.Range.Text.Trim()
                $following = $p.Next()
                if (-not $t) { $p = $following; continue }
                $dt = [datetime]::MinValue
                if ([datetime]::TryParse($t, [ref]$dt)) {
                    $null = $p.Range.Delete()
                    $p = $following
                    continue
                }
                break
            }
        }
        catch { Write-Log "    Warning: could not remove date stamp ($($_.Exception.Message))." }
    }
}

$tempDir = Join-Path $env:TEMP ('OneNoteExport_' + [guid]::NewGuid().ToString('N'))
$one = $null; $word = $null; $doc = $null; $ownWord = $false; $keepDoc = $false

try {
    Write-Log ('==================== Export started (version {0}) ====================' -f $AppVersion)

    # ---- Find the open page and its subpages ----
    $one = New-Object -ComObject OneNote.Application
    # ---- Diagnostics (written to the log) ----
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    Write-Log ("PowerShell running as administrator: {0}; 64-bit: {1}" -f $isAdmin, [Environment]::Is64BitProcess)
    $procs = @(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue)
    foreach ($pr in $procs) { Write-Log ("  ONENOTE.EXE pid {0}: {1} | window title: '{2}'" -f $pr.Id, $pr.Path, $pr.MainWindowTitle) }
    if ($procs.Count -eq 0) { Write-Log '  No ONENOTE.EXE process was running when the export started.' }

    $pageId = $null; $sectionId = $null

    # Method 1: ask OneNote for its current window
    try {
        $wins = $one.Windows
        Write-Log ("  OneNote windows reported: {0}" -f $wins.Count)
        $cw = $wins.CurrentWindow
        if ($cw -and $cw.CurrentPageId) {
            $pageId = $cw.CurrentPageId; $sectionId = $cw.CurrentSectionId
            Write-Log '  Open page found via CurrentWindow.'
        }
        else {
            foreach ($w in $wins) {
                if ($w.CurrentPageId) {
                    $pageId = $w.CurrentPageId; $sectionId = $w.CurrentSectionId
                    Write-Log '  Open page found via the Windows collection.'
                    break
                }
            }
        }
    }
    catch { Write-Log ("  Window lookup failed: {0}" -f $_.Exception.Message) }

    # Method 2: find the page OneNote marks as currently viewed
    if (-not $pageId) {
        try {
            $allXml = ''
            $one.GetHierarchy('', 4, [ref]$allXml)   # all notebooks, down to pages
            [xml]$all = $allXml
            $ns0 = New-Object System.Xml.XmlNamespaceManager($all.NameTable)
            $ns0.AddNamespace('one', $all.DocumentElement.NamespaceURI)
            $cur = $all.SelectSingleNode("//one:Page[@isCurrentlyViewed='true']", $ns0)
            if ($cur) {
                $pageId = $cur.ID; $sectionId = $cur.ParentNode.ID
                Write-Log ("  Open page found via isCurrentlyViewed: '{0}'" -f $cur.name)
            }
            else { Write-Log '  No page is marked as currently viewed in the notebook hierarchy.' }
        }
        catch { Write-Log ("  Hierarchy lookup failed: {0}" -f $_.Exception.Message) }
    }

    if (-not $pageId) {
        throw 'OneNote did not report an open page. Make sure the OneNote desktop app is open with the parent page showing, then try again.'
    }

    $xmlText = ''
    $one.GetHierarchy($sectionId, 4, [ref]$xmlText)   # 4 = hsPages
    [xml]$hier = $xmlText
    $ns = New-Object System.Xml.XmlNamespaceManager($hier.NameTable)
    $ns.AddNamespace('one', $hier.DocumentElement.NamespaceURI)
    $pages = @($hier.SelectNodes('//one:Page', $ns))

    $idx = -1
    for ($i = 0; $i -lt $pages.Count; $i++) { if ($pages[$i].ID -eq $pageId) { $idx = $i; break } }
    if ($idx -lt 0) { throw 'Could not locate the open page within its section.' }

    $parentLevel = [int]$pages[$idx].pageLevel
    $toExport = @($pages[$idx])
    for ($i = $idx + 1; $i -lt $pages.Count; $i++) {
        if ([int]$pages[$i].pageLevel -gt $parentLevel) { $toExport += $pages[$i] } else { break }
    }
    Write-Log ("Section: {0}" -f $hier.DocumentElement.GetAttribute('name'))
    Write-Log ("Parent page: '{0}' - {1} page(s) in total including subpages" -f $pages[$idx].name, $toExport.Count)

    $parentTitle = [string]$pages[$idx].name
    $out = Get-SavePath (Get-SafeFileName $parentTitle)
    if (-not $out) {
        Write-Log 'Export cancelled at the Save As dialog.'
        return
    }
    Write-Log "  Will save to: $out"

    # ---- Export each page to its own .docx ----
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    $items = @()
    $n = 0
    foreach ($p in $toExport) {
        $n++
        $file = Join-Path $tempDir ('{0:D3}.docx' -f $n)
        $one.Publish($p.ID, $file, 5, '')   # 5 = pfWord
        $wait = 0
        while (-not (Test-Path $file) -and $wait -lt 50) { Start-Sleep -Milliseconds 200; $wait++ }
        if (-not (Test-Path $file)) { throw "OneNote did not produce a Word file for page '$($p.name)'." }
        $lvl = [int]$p.pageLevel - $parentLevel + 1
        $items += [pscustomobject]@{ Title = [string]$p.name; Level = $lvl; File = $file }
        Write-Log ("  Exported page {0}: level {1}, '{2}'" -f $n, $lvl, $p.name)
    }

    # ---- Combine in Word ----
    Write-Log '  Starting Word to combine the pages...'
    $word = New-Object -ComObject Word.Application
    if ($word.Documents.Count -eq 0) { $ownWord = $true; $word.Visible = $false }
    Write-Log ("  Word version {0}; reused an already-open Word: {1}" -f $word.Version, (-not $ownWord))
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()

    $first = $true
    foreach ($it in $items) {
        $end = $doc.Content.End - 1
        $r = $doc.Range($end, $end)
        if (-not $first -and $PageBreakBetweenPages) {
            $r.InsertBreak(7)   # wdPageBreak
            $end = $doc.Content.End - 1
            $r = $doc.Range($end, $end)
        }
        $r.InsertFile($it.File)
        Format-PageTitle -Doc $doc -Start $end -Item $it
        $first = $false
    }

    if ($AddTableOfContents -and $items.Count -gt 1) {
        $maxLevel = [Math]::Min(9, ($items | Measure-Object -Property Level -Maximum).Maximum)
        $r = $doc.Range(0, 0); $r.InsertBreak(7)
        $r = $doc.Range(0, 0)
        $null = $doc.TablesOfContents.Add($r, $true, 1, $maxLevel)
        $doc.Range(0, 0).InsertBefore("Contents`r")
        $h = $doc.Paragraphs.Item(1).Range
        $h.Style = -1           # Normal
        $h.Font.Bold = $true
        $h.Font.Size = 16
        $doc.TablesOfContents.Item(1).Update()
        Write-Log '  Table of contents added.'
    }

    # ---- Save ----
    # Set the document's Title property to the parent page title
    try {
        $props = $doc.BuiltInDocumentProperties
        $tp = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @('Title'))
        [void][System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $tp, @($parentTitle))
    } catch { Write-Log "  Note: could not set the document Title property ($($_.Exception.Message))." }
    # Show Word before saving: if Word needs an answer (for example a required
    # sensitivity label), the prompt must be visible or the save waits forever.
    $word.Visible = $true
    try { $word.WindowState = 1; $doc.Activate(); $word.Activate() } catch {}
    Write-Log '  Word made visible for saving (any Word prompt will appear on screen).'
    Write-Log "  Saving to: $out"
    $fmt = 16   # wdFormatDocumentDefault (.docx)
    $saveErrors = @()
    $saved = $false
    try { $doc.SaveAs2($out, $fmt); $saved = $true }
    catch { $saveErrors += "SaveAs2: $($_.Exception.Message)" }
    if (-not $saved) {
        try { $doc.SaveAs2([ref]$out, [ref]$fmt); $saved = $true }
        catch { $saveErrors += "SaveAs2 [ref]: $($_.Exception.Message)" }
    }
    if (-not $saved) {
        try { $doc.SaveAs([ref]$out, [ref]$fmt); $saved = $true }
        catch { $saveErrors += "SaveAs [ref]: $($_.Exception.Message)" }
    }
    foreach ($e in $saveErrors) { Write-Log "    Save attempt failed - $e" }
    $w = 0
    while ($saved -and -not (Test-Path $out) -and $w -lt 25) { Start-Sleep -Milliseconds 200; $w++ }
    if (-not $saved -or -not (Test-Path $out)) {
        throw "Word could not save the file to $out"
    }
    Write-Log "Saved: $out"
    if ($OpenWhenDone) {
        $keepDoc = $true      # leave the saved document open in Word
    }
    else {
        $doc.Close(0)
        $doc = $null
    }
    Write-Log '==================== Export finished ===================='

    if (-not $OpenWhenDone) { Show-Message ("Exported {0} page(s) to:`n`n{1}" -f $items.Count, $out) }
    Show-UpdateNotice
}
catch {
    Write-Log ("ERROR: {0}" -f $_.Exception.Message)
    Write-Log ("       at line {0}" -f $_.InvocationInfo.ScriptLineNumber)
    $extra = ''
    if ($doc) {
        # Keep the combined document open so the work is not lost
        try {
            $word.Visible = $true
            $doc.Activate()
            $keepDoc = $true
            $extra = "`n`nThe combined document has been left open in Word so you can save it manually (File > Save As)."
            Write-Log '  Combined document left open in Word for manual saving.'
        } catch {}
    }
    Show-Message (("The export did not complete.`n`n{0}`n`nDetails are in:`n{1}" -f $_.Exception.Message, $LogFile) + $extra) 'Error'
}
finally {
    if ($doc -and -not $keepDoc) { try { $doc.Close(0) } catch {} }
    if ($word) {
        try { if (-not $keepDoc -and $ownWord -and $word.Documents.Count -eq 0) { $word.Quit() } } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    if ($one)  { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($one) }
    if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
}
