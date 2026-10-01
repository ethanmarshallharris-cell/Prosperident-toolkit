# Ship-Update.ps1 -- one-click release for Prosperident's desktop tools.
# Lives in the root of the Prosperident-toolkit folder (a full working copy
# of the GitHub repository: document_toolkit\, odsql_tool\, version.json,
# the build workflow). Double-click Ship-Update.bat after changing either tool.
#
# What it does: brings the folder up to date with GitHub -> works out which
# tool(s) changed since the last release -> sets their build date -> shows
# every change and asks YES -> commits and pushes -> asks YES to tag the next
# version -> waits for GitHub to build Windows + Mac versions of both tools ->
# checks all four downloads -> updates version.json so existing copies show
# the "newer version" link. Log: ship_log.txt in this folder.

$ErrorActionPreference = 'Continue'
$Root    = $PSScriptRoot
$Repo    = 'ethanmarshallharris-cell/Prosperident-toolkit'
$Assets  = @('DocumentRedactor-Windows.zip','DocumentRedactor-macOS.zip','DocumentToolkit-Windows.zip','DocumentToolkit-macOS.zip','OpenDentalSQLGenerator-Windows.zip','OpenDentalSQLGenerator-macOS.zip')
$RawUrl  = "https://raw.githubusercontent.com/$Repo/main/version.json"
$LogFile = Join-Path $Root 'ship_log.txt'
$Tools = @(
  @{ Key='document_toolkit'; Name='Document Redactor';         Dir='document_toolkit'; VerFile='document_toolkit\src\main_gui.py';  Var='APP_BUILD' },
  @{ Key='odsql_tool';       Name='OpenDental SQL Generator'; Dir='odsql_tool';       VerFile='odsql_tool\sql_generator_app.py'; Var='APP_VERSION' }
)
Remove-Item $LogFile -ErrorAction SilentlyContinue

function Log($msg) { $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg; Write-Host $line; Add-Content -Path $LogFile -Value $line }
function Fail($msg) { Log "FAILED: $msg"; Write-Host ''; Read-Host 'Press Enter to close'; exit 1 }
function Done { Write-Host ''; Read-Host 'Press Enter to close'; exit 0 }
function Ask($q) { Write-Host ''; return ((Read-Host $q) -eq 'YES') }
function ReadText($p) { [IO.File]::ReadAllText($p) }
function WriteText($p, $t) { [IO.File]::WriteAllText($p, $t, (New-Object Text.UTF8Encoding($false))) }
function JsonEsc($s) { ($s -replace '\\','\\' -replace '"','\"') }
trap { Log ("UNEXPECTED ERROR: " + $_.Exception.Message + " (line " + $_.InvocationInfo.ScriptLineNumber + ")"); Write-Host ''; Read-Host 'Press Enter to close'; exit 1 }

Set-Location $Root
Log "Toolkit folder: $Root"
if (-not (Test-Path (Join-Path $Root '.git'))) { Fail 'This folder is not a working copy of the GitHub repository (no .git folder).' }
foreach ($t in 'git','gh') { if (-not (Get-Command $t -ErrorAction SilentlyContinue)) { Fail "$t is not installed. Install Git (https://git-scm.com) and GitHub CLI (https://cli.github.com), then run this again." } }
& gh auth status 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) { Log 'Signing in to GitHub (a browser window will open).'; & gh auth login --hostname github.com --git-protocol https --web; & gh auth status 2>$null | Out-Null; if ($LASTEXITCODE -ne 0) { Fail 'GitHub sign-in did not complete.' } }
& gh auth setup-git 2>$null | Out-Null
$GhUser = (& gh api user --jq .login) 2>$null
$Push = (& gh api "repos/$Repo" --jq .permissions.push) 2>$null
if ($Push -ne 'true') { Fail "$GhUser does not have Write access to $Repo." }
Log "Signed in as $GhUser with Write access."
if (-not (& git config user.name))  { & git config user.name $GhUser }
if (-not (& git config user.email)) { $id = (& gh api user --jq .id) 2>$null; & git config user.email "$id+$GhUser@users.noreply.github.com" }

# 1. Up to date with GitHub
$Branch = (& git rev-parse --abbrev-ref HEAD).Trim()
& git fetch --quiet --tags origin
$behind = [int]((& git rev-list --count "HEAD..origin/$Branch").Trim())
if ($behind -gt 0) {
    Log "GitHub has $behind newer commit(s); bringing this folder up to date first."
    & git pull --quiet --rebase --autostash origin $Branch
    if ($LASTEXITCODE -ne 0) { Fail 'Could not merge the newer GitHub changes with yours automatically. Nothing was pushed -- send Claude this log.' }
}

# 2. What changed since the last release
$LastTag = (@(& git tag --sort=-v:refname) | Where-Object { $_ -match '^v\d+\.\d+\.\d+$' } | Select-Object -First 1)
if (-not $LastTag) { Fail 'No previous version tag (vX.Y.Z) found.' }
& git add -A
$changed = @(& git diff --cached --name-only $LastTag)
if (-not $changed) { Log "Nothing has changed since $LastTag -- nothing to ship."; Done }
Log "Changed since ${LastTag}:"; $changed | ForEach-Object { Log "   $_" }
$Changed = @($Tools | Where-Object { $d = $_.Dir + '/'; $changed | Where-Object { $_.StartsWith($d) } })
if ($Changed) { Log ("Tools changed: " + (($Changed | ForEach-Object { $_.Name }) -join ', ')) }
else { Log 'Neither tool changed (only shared files) -- changes will be pushed without a new release.' }

# 3. Build dates and user-facing notes
$Manifest = Join-Path $Root 'version.json'
$mtext = ReadText $Manifest
$Today = Get-Date -Format 'yyyy-MM-dd'
foreach ($t in $Changed) {
    $vf = Join-Path $Root $t.VerFile
    $src = ReadText $vf
    $cur = [regex]::Match($src, $t.Var + '\s*=\s*"([^"]+)"').Groups[1].Value
    $pub = [regex]::Match($mtext, '"' + $t.Key + '"\s*:\s*\{[^{}]*?"version"\s*:\s*"([^"]*)"').Groups[1].Value
    if (-not $cur) { Fail "Could not find $($t.Var) in $($t.VerFile)." }
    $new = $cur
    if ([string]::CompareOrdinal($cur, $pub) -le 0) {
        $new = $Today; $n = 1
        while ([string]::CompareOrdinal($new, $pub) -le 0) { $new = "$Today.$n"; $n++ }
        $src = [regex]::Replace($src, '(' + $t.Var + '\s*=\s*")[^"]+(")', ('${1}' + $new + '${2}'))
        WriteText $vf $src
        Log "$($t.Name): $($t.Var) set to $new (was $cur; last released $pub)."
    } else { Log "$($t.Name): $($t.Var) already $new (last released $pub)." }
    $t.NewVer = $new
    $t.Note = Read-Host "One-line note shown to users of $($t.Name) in the update link (press Enter for none)"
}

# 4. Review, commit, push
& git add -A
$pending = @(& git status --short)
$unpushed = @(& git log --oneline "origin/$Branch..HEAD")
if ($pending -or $unpushed) {
    if ($pending)  { Log 'Changes to commit:'; $pending | ForEach-Object { Log "   $_" } }
    if ($unpushed) { Log 'Commits not yet on GitHub:'; $unpushed | ForEach-Object { Log "   $_" } }
    if (-not (Ask "Push these changes to GitHub? Type YES")) { Log 'Cancelled -- nothing was pushed.'; Done }
    if ($pending) {
        $msg = Read-Host 'Short description of this update (press Enter for a default)'
        if (-not $msg) { $msg = 'Update ' + $(if ($Changed) { ($Changed | ForEach-Object { $_.Name }) -join ' and ' } else { 'shared files' }) }
        & git commit -q -m $msg
        if ($LASTEXITCODE -ne 0) { Fail 'git commit failed.' }
    }
    & git push --quiet origin $Branch
    if ($LASTEXITCODE -ne 0) { Fail 'git push failed (see above).' }
    Log ("Pushed: https://github.com/$Repo/commit/" + (& git rev-parse HEAD).Trim())
}
if (-not $Changed) { Log 'SUCCESS -- shared files pushed; no app changed, so no new release.'; Done }

# 5. Tag -> build -> verify release
$m = [regex]::Match($LastTag, '^v(\d+)\.(\d+)\.(\d+)$')
$NewTag = "v{0}.{1}.{2}" -f $m.Groups[1].Value, $m.Groups[2].Value, ([int]$m.Groups[3].Value + 1)
Write-Host "`nTagging $NewTag builds Windows and Mac versions of both tools and publishes them on the Tool Suite downloads (about 2-15 minutes)."
if (-not (Ask "Release $NewTag now? Type YES")) { Log 'Code is pushed, but no release was made. Run this again to release.'; Done }
& git tag -a $NewTag -m ("Release ${NewTag}: " + (($Changed | ForEach-Object { "$($_.Name) $($_.NewVer)" }) -join ', '))
& git push --quiet origin $NewTag
if ($LASTEXITCODE -ne 0) { Fail 'Pushing the tag failed.' }
Log "Tag $NewTag pushed. Waiting for the build to start..."
$RunId = $null
for ($i = 0; $i -lt 24 -and -not $RunId; $i++) { Start-Sleep 5; $RunId = (& gh run list --repo $Repo --workflow build.yml --branch $NewTag --limit 1 --json databaseId --jq '.[0].databaseId') 2>$null }
if (-not $RunId) { Fail "The build did not start within 2 minutes -- check https://github.com/$Repo/actions" }
Log "Building: https://github.com/$Repo/actions/runs/$RunId (this window waits for it)"
& gh run watch $RunId --repo $Repo --exit-status --interval 30 | Out-Null
$run = ((& gh run view $RunId --repo $Repo --json conclusion,jobs) -join "`n") | ConvertFrom-Json
Log "Build result: $($run.conclusion)"
if ($run.conclusion -ne 'success') { $run.jobs | ForEach-Object { Log ("   " + $_.name + ": " + $_.conclusion) }; Fail "Build failed; the previous release stays live. Details: https://github.com/$Repo/actions/runs/$RunId" }
Start-Sleep 5
$rel = ((& gh release view --repo $Repo --json tagName,assets) -join "`n") | ConvertFrom-Json
$names = @($rel.assets | ForEach-Object { $_.name })
Log ("Latest release: " + $rel.tagName + " -- " + ($names -join ', '))
$missing = $Assets | Where-Object { $names -notcontains $_ }
if ($rel.tagName -ne $NewTag -or $missing) { Fail ("Release incomplete. Missing: " + ($missing -join ', ')) }

# 6. Tell existing copies (version.json)
$mtext = ReadText $Manifest
foreach ($t in $Changed) {
    $blkV = '("' + $t.Key + '"\s*:\s*\{[^{}]*?"version"\s*:\s*")[^"]*(")'
    $blkN = '("' + $t.Key + '"\s*:\s*\{[^{}]*?"notes"\s*:\s*")[^"]*(")'
    if ($mtext -notmatch $blkV) { Fail "version.json has no '$($t.Key)' entry in the expected layout." }
    $mtext = [regex]::Replace($mtext, $blkV, ('${1}' + $t.NewVer + '${2}'))
    $mtext = [regex]::Replace($mtext, $blkN, ('${1}' + (JsonEsc $t.Note).Replace('$','$$') + '${2}'))
}
WriteText $Manifest $mtext
& git add version.json
& git commit -q -m ("Version notice: " + (($Changed | ForEach-Object { "$($_.Key) $($_.NewVer)" }) -join ', ') + " ($NewTag)")
& git push --quiet origin $Branch
if ($LASTEXITCODE -ne 0) { Fail 'Release is live, but version.json could not be pushed -- older copies will not be prompted. Send Claude this log.' }
$ok = $false
for ($i = 0; $i -lt 10 -and -not $ok; $i++) {
    try { $mf = Invoke-RestMethod -Uri ($RawUrl + '?x=' + [guid]::NewGuid().ToString('N')) -TimeoutSec 10
          $ok = $true; foreach ($t in $Changed) { if ($mf.($t.Key).version -ne $t.NewVer) { $ok = $false } } } catch {}
    if (-not $ok) { Start-Sleep 30 }
}
Log $(if ($ok) { 'Live version.json confirmed -- older copies will show the update link.' } else { 'version.json pushed; GitHub may take a few minutes to serve it.' })
Log ("SUCCESS -- $NewTag released: " + (($Changed | ForEach-Object { "$($_.Name) build $($_.NewVer)" }) -join ', ') + '. If features changed, ask Claude to update the Tool Suite card.')
Done
