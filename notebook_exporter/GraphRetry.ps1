<#
    GraphRetry.ps1  -  Prosperident, internal use only

    One shared Retry function for every script that calls Microsoft Graph.
    Dot-source it:   . (Join-Path $Root 'GraphRetry.ps1')

    Why this file exists: by 22 Sep 2026, several scripts (Set-CasePolicyLinks,
    Set-CaseCrossLinks, Backfill-QuestionnairePages, Verify-PolicyLinks,
    Check-CrossLinks) each made bare Invoke-RestMethod / Invoke-WebRequest
    calls with no retry logic. A single busy run of "Backfill All
    Notebooks.bat" - which chains all of these, several times each, once per
    case - threw enough traffic at Graph to draw 504s and 429s, and every one
    of those bare calls crashed its script outright and stopped the whole
    pipeline partway through. Backfill-NotebookPages.ps1 already had its own
    copy of this Retry function; rather than keep copy-pasting it (and
    forgetting it in the next new script), it now lives here once and every
    Graph-calling script dot-sources it.

    6 attempts, honouring Retry-After on a 429, 10/20/40/80/160s backoff
    otherwise. A 404 is a real answer, not a transient failure - it is never
    retried; the caller decides what a 404 means.

    Bumped from 4 to 6 attempts on 23 Sep 2026: a full run of "Backfill All
    Notebooks.bat" makes many hundreds of Graph calls across six-plus cases
    back to back, and by the time it reaches the last scripts in the
    pipeline (Verify-PolicyLinks, Check-CrossLinks) Graph's per-minute
    budget can already be exhausted from everything before it, so a 429
    there can outlast 4 attempts (~70s of backoff) even though the same
    throttling clears fine earlier in a run. See also the per-case pacing
    delay added in Backfill All Notebooks.bat for the other half of this fix
    - retrying harder helps less than not hammering it as hard to begin with.
#>
function Retry {
    param([scriptblock]$Call)
    for ($a = 1; $a -le 6; $a++) {
        try { return (& $Call) }
        catch {
            $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
            if ($code -eq 404) { throw }
            if ($a -eq 6 -or -not ($code -ge 500 -or $code -eq 429 -or $code -eq 0)) { throw }
            $wait = 10 * [math]::Pow(2, $a - 1)
            if ($code -eq 429) {
                try { $ra = $_.Exception.Response.Headers['Retry-After']; if ($ra) { $wait = [math]::Max([int]$ra, 5) } } catch { }
            }
            Write-Host ("    Graph {0} - waiting {1}s and retrying" -f $code, $wait) -ForegroundColor DarkYellow
            Start-Sleep -Seconds $wait
        }
    }
}
