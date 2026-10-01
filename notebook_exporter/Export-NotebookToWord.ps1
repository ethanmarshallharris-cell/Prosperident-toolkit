<#
    Export-NotebookToWord.ps1
    Prosperident - internal use only

    Exports a whole case notebook to ONE Word document (and, with -Pdf, a PDF).

    WHY (David Harris, 30 September 2026): case notebooks are now used only in
    OneNote on the web, which - unlike the desktop app - cannot export a
    notebook to Word. This does it through Microsoft Graph and Word itself.

    WHAT YOU GET
      - A cover page (notebook, case, lead, export date, who exported it) and a
        table of contents.
      - Every section in notebook order, every page in section order, sub-pages
        indented one level in the contents. Each page starts on a new sheet,
        headed with its title and its created / last-modified dates.
      - Page text, tables, links, checkboxes and pictures. Pictures are
        embedded, at full resolution where OneNote has it, scaled to the page.
      - FILE ATTACHMENTS (PDFs, spreadsheets dropped onto a page) are saved in
        a folder beside the document, and the page shows the file name where
        the attachment sat.
      - Page numbers in the footer.

    WHAT IT CANNOT INCLUDE - the OneNote API does not hand these over:
      handwriting / drawing (ink), audio and video recordings. The cover page
      says so. A page that cannot be read at all (a damaged page) appears with
      a note in its place, so nothing goes missing silently.

    WHERE IT GOES: Documents\Prosperident Notebook Exports\ on this computer,
    unless -OutDir says otherwise. THE DOCUMENT HOLDS CLIENT AND PATIENT
    INFORMATION - file it in the case document library; do not email it.

    Needs Microsoft Word on this computer, and the OneNote sign-in used by the
    other toolkit scripts ("Sign In For OneNote.bat").

    RUN:  "Export Case Notebook to Word.bat"   (asks which case)
          .\Export-NotebookToWord.ps1 -CaseName King
          .\Export-NotebookToWord.ps1 -CaseName King -Pdf
          .\Export-NotebookToWord.ps1 -ListCases
#>

[CmdletBinding()]
param(
    [string]$CaseName,
    [string]$OutDir,
    [switch]$Pdf,
    [switch]$ListCases
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'GraphRetry.ps1')

# Tool Suite release number - Ship-Update.bat sets this on each release.
$AppVersion  = "2026-10-01"
$ManifestUrl = 'https://raw.githubusercontent.com/ethanmarshallharris-cell/Prosperident-toolkit/main/version.json'

# Prints one line at the end of an export when the Tool Suite has a newer
# version. Never stops or delays an export: any problem is ignored.
function Show-UpdateNotice {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $m = Invoke-RestMethod -Uri $ManifestUrl -TimeoutSec 4 -UseBasicParsing
        $e = $m.notebook_exporter
        if ($e -and $e.version -and [string]::CompareOrdinal([string]$e.version, $AppVersion) -gt 0) {
            Write-Host ("  A newer version ({0}) is available at www.prosperident.com/tool-suite - you have {1}." -f $e.version, $AppVersion) -ForegroundColor Yellow
            if ($e.notes) { Write-Host ("  {0}" -f $e.notes) -ForegroundColor Yellow }
            Write-Host '  Extract it over this folder and keep your casecapture-config.json.' -ForegroundColor Yellow
            Write-Host ''
        }
    } catch { }
}

$ConfigFile = Join-Path $PSScriptRoot 'casecapture-config.json'
# The case list names clients, so it is never part of the download - it is
# handed out internally and dropped into this folder.
if (-not (Test-Path $ConfigFile)) {
    Write-Host ''
    Write-Host 'The case list (casecapture-config.json) is not in this folder:' -ForegroundColor Red
    Write-Host ("  {0}" -f $PSScriptRoot) -ForegroundColor Red
    Write-Host 'It is not included in the download because it names clients.' -ForegroundColor Red
    Write-Host 'Ask David for the current copy and save it in this folder, then run this again.' -ForegroundColor Red
    Write-Host ''
    exit 1
}
$cases = @((Get-Content $ConfigFile -Raw | ConvertFrom-Json).cases | Where-Object { $_.active -ne $false })

if ($ListCases -or -not $CaseName) {
    Write-Host ''
    Write-Host 'Cases that can be exported:' -ForegroundColor Cyan
    $cases | ForEach-Object { Write-Host ("  {0,-32} {1}" -f $_.caseName, $_.notebookName) }
    Write-Host ''
    exit 0
}

$case = $cases | Where-Object { $_.caseName -eq $CaseName.Trim() } | Select-Object -First 1
if (-not $case) {
    Write-Host ("No active case named '{0}'. Names are as listed by -ListCases." -f $CaseName) -ForegroundColor Red
    exit 1
}

if (-not $OutDir) { $OutDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Prosperident Notebook Exports' }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$tok = & (Join-Path $PSScriptRoot 'Get-DelegatedToken.ps1') 2>&1 | ForEach-Object { "$_" }
if ($tok -is [array]) { $tok = $tok[-1] }
if ($tok -notmatch '^ey') { Write-Host 'No OneNote credential - run "Sign In For OneNote.bat".' -ForegroundColor Red; exit 1 }
$H    = @{ Authorization = "Bearer $tok" }
$Base = "https://graph.microsoft.com/v1.0/groups/$($case.groupId)/onenote"

function E { param([string]$s) [Net.WebUtility]::HtmlEncode($s) }
function Safe-Name { param([string]$s) (($s -replace '[\\/:*?"<>|]+', ' ') -replace '\s+', ' ').Trim() }
function Fmt-Date { param($d) if ($d) { ([datetime]$d).ToLocalTime().ToString('MM/dd/yyyy h:mm tt') } else { '' } }

function Get-All {
    param([string]$Uri)
    $all = @()
    while ($Uri) {
        $r = Retry { Invoke-RestMethod -Headers $H -Uri $Uri }
        $all += @($r.value)
        $Uri = $r.'@odata.nextLink'
    }
    return ,$all
}

function Read-Html {
    param([string]$Uri)
    Retry {
        $r = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Headers $H -TimeoutSec 180
        [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
    }
}

function Save-Resource {
    # Downloads a OneNote resource (picture or attachment) to $Path.
    # Returns the content type, or $null if it could not be fetched.
    param([string]$Url, [string]$Path)
    try {
        $r = Retry { Invoke-WebRequest -UseBasicParsing -Uri $Url -Headers $H -TimeoutSec 300 -OutFile $Path -PassThru }
        return [string]$r.Headers['Content-Type']
    } catch { return $null }
}

# --------------------------------------------------------------- stamp ---
$stamp    = Get-Date -Format 'yyyy-MM-dd HHmm'
$docName  = Safe-Name ("{0} - exported {1}" -f $case.notebookName, $stamp)
$DocxPath = Join-Path $OutDir "$docName.docx"
$PdfPath  = Join-Path $OutDir "$docName.pdf"
$AttDir   = Join-Path $OutDir "$docName - attachments"
$Work     = Join-Path $env:TEMP ("nbexport-" + [guid]::NewGuid().ToString('N'))
$ImgDir   = Join-Path $Work 'img'
New-Item -ItemType Directory -Path $ImgDir -Force | Out-Null

Write-Host ''
Write-Host ("=== EXPORT CASE NOTEBOOK TO WORD  (version {0}) ===" -f $AppVersion) -ForegroundColor Cyan
Write-Host ("  Case     : {0}" -f $case.caseName)
Write-Host ("  Notebook : {0}" -f $case.notebookName)
Write-Host ("  Output   : {0}" -f $DocxPath)
Write-Host ''

# ------------------------------------------------------- notebook layout ---
$nb = @((Retry { Invoke-RestMethod -Headers $H -Uri "$Base/notebooks" }).value) | Where-Object { $_.displayName -eq $case.notebookName } | Select-Object -First 1
if (-not $nb) { Write-Host ("  Notebook '{0}' not found in the case group." -f $case.notebookName) -ForegroundColor Red; exit 1 }

# Sections, including any inside section groups, in the order OneNote lists them.
$sections = @()
function Add-Sections {
    param([string]$ParentUri, [string]$Prefix)
    foreach ($s in (Get-All "$ParentUri/sections?`$select=id,displayName")) {
        $script:sections += [pscustomobject]@{ Id = $s.id; Name = ($Prefix + $s.displayName) }
    }
    foreach ($g in (Get-All "$ParentUri/sectionGroups?`$select=id,displayName")) {
        Add-Sections "$Base/sectionGroups/$($g.id)" ($Prefix + $g.displayName + ' / ')
    }
}
Add-Sections "$Base/notebooks/$($nb.id)" ''
Write-Host ("  {0} section(s)" -f $sections.Count)

# ------------------------------------------------------------ page html ---
$missingNote = '<p style="color:#a33;"><i>This page could not be read from OneNote (HTTP {0}). Open the notebook to see it.</i></p>'
$sb = New-Object System.Text.StringBuilder
$counts = @{ pages = 0; failed = 0; images = 0; imgFailed = 0; attachments = 0 }
$imgN = 0; $attN = 0

function Convert-PageBody {
    # OneNote's page HTML -> flowing HTML Word imports cleanly.
    param([string]$Html)
    $m = [regex]::Match($Html, '(?is)<body[^>]*>(.*)</body>')
    $b = if ($m.Success) { $m.Groups[1].Value } else { $Html }

    # Absolute positioning: OneNote places outlines at x/y; Word needs flow.
    $b = [regex]::Replace($b, '(?i)position\s*:\s*absolute\s*;?', '')
    $b = [regex]::Replace($b, '(?i)(?<![-\w])(left|top)\s*:\s*-?\d+(\.\d+)?px\s*;?', '')
    $b = [regex]::Replace($b, '(?i)(<div\b[^>]*style="[^"]*?)\bwidth\s*:\s*\d+(\.\d+)?px\s*;?', '$1')

    # Page headings sit below the export's own four TOC levels
    # (levels 1-4 are sections, pages, sub-pages and sub-sub-pages).
    $b = [regex]::Replace($b, '(?i)<(/?)h[1-2]\b', '<$1h5')
    $b = [regex]::Replace($b, '(?i)<(/?)h[3-6]\b', '<$1h6')

    # Checkboxes (OneNote "to-do" tags).
    $b = [regex]::Replace($b, '(?i)(<(p|li|h\d)\b[^>]*data-tag="[^"]*to-do:completed[^"]*"[^>]*>)', '$1&#9745; ')
    $b = [regex]::Replace($b, '(?i)(<(p|li|h\d)\b[^>]*data-tag="[^"]*to-do(?!:completed)[^"]*"[^>]*>)', '$1&#9744; ')

    # Embedded video and the like - a link instead.
    $b = [regex]::Replace($b, '(?is)<iframe\b[^>]*data-original-src="([^"]+)"[^>]*>(</iframe>)?', '<p><i>Embedded media: <a href="$1">$1</a></i></p>')
    $b = [regex]::Replace($b, '(?is)<iframe\b[^>]*>(</iframe>)?', '<p><i>[Embedded media - open the notebook to view]</i></p>')

    # File attachments: save beside the document, name them in the text.
    $b = [regex]::Replace($b, '(?is)<object\b([^>]*)/>|<object\b([^>]*)>(.*?</object>)?', [Text.RegularExpressions.MatchEvaluator]{
        param($x)
        $attrs = $x.Groups[1].Value + $x.Groups[2].Value
        $name  = [regex]::Match($attrs, 'data-attachment="([^"]+)"').Groups[1].Value
        $url   = [regex]::Match($attrs, '\bdata="([^"]+)"').Groups[1].Value
        if (-not $name) { $name = 'attachment' }
        $name = [Net.WebUtility]::HtmlDecode($name)
        if (-not $url) { return "<p><i>[Attachment: $(E $name) - could not be exported]</i></p>" }
        New-Item -ItemType Directory -Path $AttDir -Force | Out-Null
        $script:attN++
        $file = Safe-Name ("{0:000} {1}" -f $script:attN, $name)
        $ok = Save-Resource ([Net.WebUtility]::HtmlDecode($url)) (Join-Path $AttDir $file)
        if ($ok) { $script:counts.attachments++; return "<p style=""background:#F2F2F2;padding:4px;""><b>Attachment:</b> $(E $name) <i>(saved as &quot;$(E $file)&quot; in the attachments folder beside this document)</i></p>" }
        return "<p><i>[Attachment: $(E $name) - could not be downloaded]</i></p>"
    })

    # Pictures: download (full resolution where available) and point at the local file.
    $b = [regex]::Replace($b, '(?is)<img\b[^>]*>', [Text.RegularExpressions.MatchEvaluator]{
        param($x)
        $tag = $x.Value
        $src = [regex]::Match($tag, 'data-fullres-src="([^"]+)"').Groups[1].Value
        if (-not $src) { $src = [regex]::Match($tag, '\bsrc="([^"]+)"').Groups[1].Value }
        if (-not $src) { return '' }
        $src = [Net.WebUtility]::HtmlDecode($src)
        $script:imgN++
        $tmp = Join-Path $ImgDir ("img{0:0000}" -f $script:imgN)
        $ct  = if ($src -match '^https://graph\.microsoft\.com/') { Save-Resource $src $tmp } else {
                   try { Invoke-WebRequest -UseBasicParsing -Uri $src -OutFile $tmp -TimeoutSec 120; 'image/jpeg' } catch { $null } }
        if (-not $ct) { $script:counts.imgFailed++; return '<p><i>[Picture - could not be downloaded]</i></p>' }
        $ext  = switch -regex ($ct) { 'png' { '.png' } 'gif' { '.gif' } 'bmp' { '.bmp' } 'tiff' { '.tif' } default { '.jpg' } }
        Move-Item -LiteralPath $tmp -Destination ($tmp + $ext) -Force
        $script:counts.images++
        $w = [regex]::Match($tag, '\bwidth="([\d.]+)"').Groups[1].Value
        $h = [regex]::Match($tag, '\bheight="([\d.]+)"').Groups[1].Value
        $alt = [regex]::Match($tag, '\balt="([^"]*)"').Groups[1].Value
        $size = ''
        if ($w) { $size += " width=""$([int][double]$w)""" }
        if ($h) { $size += " height=""$([int][double]$h)""" }
        return "<img src=""file:///$(($tmp + $ext) -replace '\\','/')""$size alt=""$alt"" />"
    })

    # Drop OneNote's generated ids; Word does not need them.
    $b = [regex]::Replace($b, '\s(id|data-id|data-fullres-src|data-src-type|data-fullres-src-type)="[^"]*"', '')
    return $b
}

foreach ($sec in $sections) {
    [void]$sb.Append("<h1 style=""page-break-before:always"">$(E $sec.Name)</h1>")
    $pages = Get-All "$Base/sections/$($sec.Id)/pages?`$select=id,title,level,order,createdDateTime,lastModifiedDateTime&`$top=100"
    $pages = @($pages | Sort-Object { [int]$_.order })
    Write-Host ("  {0,-40} {1} page(s)" -f $sec.Name, $pages.Count)
    $first = $true
    foreach ($p in $pages) {
        $title = if ("$($p.title)".Trim()) { $p.title } else { '(untitled page)' }
        # Sub-pages: OneNote reports level 0 (page), 1 (sub-page), 2 (sub-sub-page).
        # They follow their parent in section order and sit one heading level
        # lower, so the contents shows them indented beneath it.
        $lvl   = [math]::Min([math]::Max([int]$p.level, 0), 2)
        $tag   = 'h' + (2 + $lvl)
        $brk   = if ($first) { '' } else { ' style="page-break-before:always"' }
        $first = $false
        [void]$sb.Append("<$tag$brk>$(E $title)</$tag>")
        [void]$sb.Append("<p style=""font-size:8pt;color:#777;"">Created $(Fmt-Date $p.createdDateTime) &nbsp;|&nbsp; Last modified $(Fmt-Date $p.lastModifiedDateTime)</p>")
        try {
            $html = Read-Html "$Base/pages/$($p.id)/content"
            [void]$sb.Append((Convert-PageBody $html))
            $counts.pages++
        }
        catch {
            $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
            [void]$sb.Append(($missingNote -f $code))
            $counts.failed++
            Write-Host ("      ! '{0}' could not be read (HTTP {1}) - noted in the document" -f $title, $code) -ForegroundColor Yellow
        }
        Start-Sleep -Milliseconds 250
    }
}

# --------------------------------------------------------- cover + wrap ---
$me = try { (Retry { Invoke-RestMethod -Headers $H -Uri 'https://graph.microsoft.com/v1.0/me?$select=displayName' }).displayName } catch { $env:USERNAME }
$cover = @"
<p style="font-size:10pt;color:#777;">PROSPERIDENT - CONFIDENTIAL</p>
<p style="font-size:24pt;font-weight:bold;color:#1F4E79;">$(E $case.notebookName)</p>
<p style="font-size:12pt;">Case: $(E $case.caseName)</p>
<p style="font-size:10pt;">Exported from OneNote on $(Get-Date -Format 'MM/dd/yyyy h:mm tt') by $(E $me).<br/>
$($counts.pages) page(s)$(if ($counts.failed) { ", $($counts.failed) unreadable (noted where they fall)" }); $($counts.images) picture(s); $($counts.attachments) attachment(s)$(if ($counts.attachments) { ' saved in the folder beside this document' }).</p>
<p style="font-size:9pt;color:#777;">A snapshot - the notebook remains the working record. Handwriting (ink), audio and video
recordings cannot be exported by the OneNote service and are not included.
This document contains client and patient information: keep it in the case document library and do not email it.</p>
<p>[[TOC]]</p>
"@
$HtmlPath = Join-Path $Work 'export.html'
$full = "<html><head><meta charset=""utf-8""><title>$(E $case.notebookName)</title>" +
        '<style>body{font-family:Calibri,sans-serif;font-size:11pt;} table{border-collapse:collapse;} td,th{border:1px solid #999;padding:3px;vertical-align:top;}</style>' +
        "</head><body>$cover$($sb.ToString())</body></html>"
[IO.File]::WriteAllText($HtmlPath, $full, (New-Object Text.UTF8Encoding $true))

# ------------------------------------------------------------------ Word ---
Write-Host ''
Write-Host ("  Building the Word document ({0:N0} KB of page content)..." -f ((Get-Item $HtmlPath).Length / 1KB)) -ForegroundColor Cyan
function Step { param([string]$t) Write-Host ("    {0:HH:mm:ss}  {1}" -f (Get-Date), $t) }
$word = $null; $doc = $null
try {
    Step 'starting Word'
    $word = New-Object -ComObject Word.Application
    # VISIBLE on purpose (30 Sep 2026): the first live run sat for 15+ minutes
    # at this stage with Word hidden. A hidden Word that raises any prompt waits
    # forever for a click nobody can give; a visible one shows the prompt.
    $word.Visible = $true
    $word.ScreenUpdating = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch { }      # no macros from the import
    Step 'opening the exported pages in Word (the slowest step for a big notebook)'
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ..., Format=wdOpenFormatWebPages)
    $doc = $word.Documents.Open($HtmlPath, $false, $false, $false)
    try { $doc.ActiveWindow.View.Type = 3 } catch { }  # print layout

    $ps = $doc.PageSetup
    $ps.TopMargin = 72; $ps.BottomMargin = 72; $ps.LeftMargin = 72; $ps.RightMargin = 72
    $usable = $ps.PageWidth - $ps.LeftMargin - $ps.RightMargin

    $nShapes = $doc.InlineShapes.Count
    Step ("embedding {0} picture(s)" -f $nShapes)
    for ($i = 1; $i -le $nShapes; $i++) {
        $s = $doc.InlineShapes.Item($i)
        try { if ($s.LinkFormat) { $s.LinkFormat.SavePictureWithDocument = $true; $s.LinkFormat.BreakLink() } } catch { }
        try { if ($s.Width -gt $usable) { $s.LockAspectRatio = -1; $s.Width = $usable } } catch { }
        if ($i % 25 -eq 0) { Step ("  {0} of {1}" -f $i, $nShapes) }
    }
    $nTables = $doc.Tables.Count
    Step ("fitting {0} table(s) to the page" -f $nTables)
    for ($i = 1; $i -le $nTables; $i++) { try { $doc.Tables.Item($i).AutoFitBehavior(2) } catch { } }

    Step 'adding contents and page numbers'
    $rng = $doc.Content
    if ($rng.Find.Execute('[[TOC]]')) {
        $rng.Text = ''
        [void]$doc.TablesOfContents.Add($rng, $true, 1, 4)
    }
    for ($i = 1; $i -le $doc.Sections.Count; $i++) { try { [void]$doc.Sections.Item($i).Footers.Item(1).PageNumbers.Add(1, $true) } catch { } }
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) { $doc.TablesOfContents.Item($i).Update() }

    Step 'saving'
    $doc.SaveAs2($DocxPath, 16)                         # .docx
    Write-Host ("  Saved {0}" -f $DocxPath) -ForegroundColor Green
    $word.ScreenUpdating = $true
    if ($Pdf) { Step 'saving the PDF'; $doc.SaveAs2($PdfPath, 17); Write-Host ("  Saved {0}" -f $PdfPath) -ForegroundColor Green }
}
catch {
    Write-Host ("  ! Word could not build the document: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host ("    The raw export is kept at {0}" -f $HtmlPath) -ForegroundColor Yellow
    if ($doc)  { try { $doc.Close(0) } catch { } }
    if ($word) { try { $word.Quit() } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    exit 2
}
finally {
    if ($doc)  { try { $doc.Close(0) } catch { } }
    if ($word) { try { $word.Quit() } catch { }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { } }
}
Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host ("  Pages {0}   unreadable {1}   pictures {2} (failed {3})   attachments {4}" -f $counts.pages, $counts.failed, $counts.images, $counts.imgFailed, $counts.attachments)
if ($counts.attachments) { Write-Host ("  Attachments: {0}" -f $AttDir) }
Write-Host '  File the document in the case document library - it holds client and patient information.' -ForegroundColor Yellow
Write-Host ''
Show-UpdateNotice
if ($counts.failed -or $counts.imgFailed) { exit 3 }
exit 0
