<#
    Get-DelegatedToken.ps1
    Prosperident - internal use only

    Returns a Microsoft Graph access token that carries a USER identity, for the
    OneNote calls that can no longer be made app-only.

    WHY THIS EXISTS
    Microsoft retired app-only (application permission) tokens for the OneNote
    Graph API on 31 MARCH 2025. Every Notes endpoint now returns:

        401  code 40001  "this API will no longer support app-only tokens"

    regardless of which permissions are granted. It is a platform decision, not
    a tenant or consent problem, and no permission grant will change it.
    Delegated (app + user) tokens still work.

    So OneNote work runs as a signed-in user. Everything else in this project -
    Planner, Mail.Read, group membership - stays app-only and is unaffected.

    HOW IT WORKS
      One time, per machine:
          .\Get-DelegatedToken.ps1 -SignIn
      Device code flow: it prints a code, you open the URL and sign in AS THE
      SERVICE ACCOUNT. The resulting refresh token is DPAPI-encrypted to this
      Windows account in refresh.dat, exactly like secret.dat.

      Every run after that:
          $token = & .\Get-DelegatedToken.ps1
      Silent. The refresh token is exchanged for an access token and Microsoft
      returns a NEW refresh token each time, which is saved. As long as the
      worker runs at least every 90 days the credential renews itself
      indefinitely.

    WHAT BREAKS IT
      - The service account's password being changed or reset
      - A Conditional Access policy that re-challenges the account
      - The account being disabled, or its licence removed
      - 90+ days without a run
    All produce the same symptom - the refresh fails - and the same fix: run
    -SignIn again. The script says so rather than failing cryptically.

    PREREQUISITES (IT)
      On the 'Investigation Launcher' app registration:
        1. Authentication > Advanced settings >
           "Allow public client flows" = YES     (required for device code)
        2. API permissions > Add > Microsoft Graph > DELEGATED:
               Notes.ReadWrite.All
               Sites.ReadWrite.All
               offline_access
           + grant admin consent
      The existing Application permissions stay exactly as they are.
#>

[CmdletBinding()]
param(
    # Run the one-time interactive sign-in and store the refresh token.
    [switch]$SignIn,

    # Report what is stored and whether it still works. Changes nothing.
    [switch]$Status,

    # Ask for a different delegated scope when redeeming the refresh token.
    # v2 refresh tokens are not locked to the scope they were issued for, so any
    # scope the user has already consented to can be redeemed with the stored
    # credential. Left alone, the default below is unchanged - callers that need
    # something else (Pin-MyCasePlans.ps1 needs Tasks.ReadWrite) pass it here
    # rather than widening the default and putting the notebook flow at risk.
    [string]$Scope,

    # Include Tasks.ReadWrite in the one-time sign-in so the user consents to it
    # at the same time. Only meaningful with -SignIn.
    [switch]$IncludeTasks
)

$ErrorActionPreference = 'Stop'

$TenantId    = '7ac155fb-0e13-4fc2-b4cf-06e90501a737'
$ClientId    = '4c1303b4-ee01-4b20-9014-a659cf23766a'
$RefreshFile = Join-Path $env:LOCALAPPDATA 'Prosperident\CaseToolkit\refresh.dat'

# Whether Entra wants the client secret on these calls depends on how the app
# registration is configured, and BOTH answers have been seen on this app within
# one day:
#   "Allow public client flows" = No  -> AADSTS700218, secret REQUIRED
#   "Allow public client flows" = Yes -> AADSTS700025, secret FORBIDDEN
# Guessing either way breaks the moment someone changes that switch, so don't
# guess: send it or not, read the error, and retry the other way once.
$SecretFile = Join-Path $env:LOCALAPPDATA 'Prosperident\CaseToolkit\secret.dat'
$ClientSecret = $null
if (Test-Path $SecretFile) {
    try {
        $rawSec = [System.IO.File]::ReadAllText($SecretFile).TrimStart([char]0xFEFF).Trim()
        $ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR(($rawSec | ConvertTo-SecureString)))
    } catch { $ClientSecret = $null }
}

# Start by NOT sending it - the app is a public client as of 28 Aug 2026.
$script:UseSecret = $false

function Invoke-TokenRequest {
    param([hashtable]$Fields)
    $attempt = 0
    while ($true) {
        $attempt++
        $body = @{}
        foreach ($k in $Fields.Keys) { $body[$k] = $Fields[$k] }
        if ($script:UseSecret -and $ClientSecret) { $body['client_secret'] = $ClientSecret }
        try {
            return Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                -ContentType 'application/x-www-form-urlencoded' -Body $body
        }
        catch {
            $txt = ''
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $txt = $_.ErrorDetails.Message }
            if (-not $txt -and $_.Exception.Response) {
                try { $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $txt = $sr.ReadToEnd(); $sr.Close() } catch { }
            }
            if ($attempt -eq 1 -and $txt -match 'AADSTS700218') { $script:UseSecret = $true;  continue }
            if ($attempt -eq 1 -and $txt -match 'AADSTS700025') { $script:UseSecret = $false; continue }
            throw
        }
    }
}

# offline_access is what makes Microsoft issue a refresh token at all. Without
# it the sign-in appears to succeed and then nothing works an hour later.
$DefaultScope = 'offline_access https://graph.microsoft.com/Notes.ReadWrite.All https://graph.microsoft.com/Sites.ReadWrite.All'
if (-not $Scope) { $Scope = $DefaultScope }
if ($IncludeTasks -and $Scope -notmatch 'Tasks\.ReadWrite') {
    $Scope = $Scope + ' https://graph.microsoft.com/Tasks.ReadWrite'
}

function Save-Refresh {
    param([string]$Value)
    # ASCII, not UTF8: the BOM that UTF8 prepends corrupts the round-trip
    # through ConvertTo-SecureString. This cost an afternoon once already.
    $enc = ConvertTo-SecureString $Value -AsPlainText -Force | ConvertFrom-SecureString
    [System.IO.File]::WriteAllText($RefreshFile, $enc, [System.Text.Encoding]::ASCII)

    $back = [System.IO.File]::ReadAllText($RefreshFile).TrimStart([char]0xFEFF).Trim()
    $test = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR(($back | ConvertTo-SecureString)))
    if ($test -ne $Value) { throw 'Refresh token did not survive the round-trip to disk.' }
}

function Read-Refresh {
    if (-not (Test-Path $RefreshFile)) { return $null }
    $raw = [System.IO.File]::ReadAllText($RefreshFile).TrimStart([char]0xFEFF).Trim()
    return [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
           [Runtime.InteropServices.Marshal]::SecureStringToBSTR(($raw | ConvertTo-SecureString)))
}

function Show-SignInHelp {
    Write-Host ''
    Write-Host 'To fix this, run:  "Sign In For OneNote.bat"' -ForegroundColor Cyan
    Write-Host 'and sign in as the OneNote service account.' -ForegroundColor Cyan
    Write-Host ''
}

# ---------------------------------------------------------------- sign-in ----
if ($SignIn) {

    Write-Host ''
    Write-Host '=== ONENOTE SERVICE ACCOUNT SIGN-IN ===' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'This is needed once per machine. It stores a renewable credential' -ForegroundColor Gray
    Write-Host 'so the notebook and email-capture jobs can run unattended.' -ForegroundColor Gray
    Write-Host ''

    try {
        $dc = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body @{ client_id = $ClientId; scope = $Scope }
    }
    catch {
        Write-Host 'Could not start the sign-in.' -ForegroundColor Red
        Write-Host ''
        Write-Host 'The usual cause is that device code flow is not enabled on the app.' -ForegroundColor Yellow
        Write-Host 'Ask IT, on the Investigation Launcher app registration:' -ForegroundColor Yellow
        Write-Host '   Authentication > Advanced settings >' -ForegroundColor Yellow
        Write-Host '   "Allow public client flows" = Yes' -ForegroundColor Yellow
        Write-Host ''
        Write-Host ("Detail: {0}" -f $_.Exception.Message) -ForegroundColor DarkGray
        exit 1
    }

    # Open the browser AT the code rather than printing a code to retype. The otc
    # query parameter prefills it; verification_uri_complete is the same thing when
    # Entra bothers to return it. Also put it on the clipboard, because the browser
    # sometimes lands on the plain page and then it has to be typed after all.
    $codeUrl = $dc.verification_uri_complete
    if (-not $codeUrl) { $codeUrl = ("https://microsoft.com/devicelogin?otc={0}" -f $dc.user_code) }

    try { Set-Clipboard -Value $dc.user_code; $onClip = $true } catch { $onClip = $false }

    Write-Host '  Opening your browser...' -ForegroundColor Cyan
    Write-Host ''
    try { Start-Process $codeUrl | Out-Null }
    catch {
        Write-Host '  Could not open a browser. Open this page yourself:' -ForegroundColor Yellow
        Write-Host ("       {0}" -f $dc.verification_uri) -ForegroundColor White
    }

    Write-Host ("  The code is  {0}" -f $dc.user_code) -ForegroundColor White
    if ($onClip) { Write-Host '  (already copied to your clipboard - Ctrl+V if asked)' -ForegroundColor Gray }
    else         { Write-Host ("  If the page asks for it, type it in. Page: {0}" -f $dc.verification_uri) -ForegroundColor Gray }
    Write-Host ''
    Write-Host '  3. Sign in.' -ForegroundColor Yellow
    Write-Host '' -ForegroundColor Gray
    Write-Host '     Prosperident has no dedicated automation account today, so sign in' -ForegroundColor Gray
    Write-Host '     as yourself. That works. What it costs: every notebook edit is' -ForegroundColor Gray
    Write-Host '     attributed to you, and the credential dies when your password' -ForegroundColor Gray
    Write-Host '     changes or a Conditional Access policy challenges you - at which' -ForegroundColor Gray
    Write-Host '     point notebooks silently stop building until this is run again.' -ForegroundColor Gray
    Write-Host '' -ForegroundColor Gray
    Write-Host '     The durable fix is a licensed account of its own (say' -ForegroundColor Gray
    Write-Host '     automation@prosperident.com) with a password that nobody rotates,' -ForegroundColor Gray
    Write-Host '     added as a standing member of every case group. Until that exists,' -ForegroundColor Gray
    Write-Host '     your own sign-in is the right answer, not a workaround to feel bad' -ForegroundColor Gray
    Write-Host '     about.' -ForegroundColor Gray
    Write-Host ''
    Write-Host 'Waiting...' -ForegroundColor Gray

    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    $tok = $null
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds ([int]$dc.interval + 1)
        try {
            $tok = Invoke-TokenRequest @{ client_id   = $ClientId
                                          grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                                          device_code = $dc.device_code }
            break
        }
        catch {
            # Read the body TWO ways. PowerShell 7 raises HttpResponseException and
            # has already consumed the response stream, so GetResponseStream()
            # throws and the old single-path read returned an empty $err. An empty
            # $err then failed the authorization_pending test and killed the loop
            # on the FIRST poll - which looks like "sign-in failed" a second after
            # the code appears, before anyone could possibly have signed in.
            $body = ''
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $body = $_.ErrorDetails.Message }
            if (-not $body) {
                try {
                    $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $body = $sr.ReadToEnd(); $sr.Close()
                } catch { }
            }

            $err = ''; $desc = ''
            if ($body) {
                try { $j = $body | ConvertFrom-Json; $err = $j.error; $desc = $j.error_description } catch { $desc = $body }
            }

            # authorization_pending is the normal "not finished yet" answer. So is
            # an error we could not parse - keep waiting rather than giving up on
            # someone who is still typing their password.
            if ($err -eq 'authorization_pending' -or $err -eq 'slow_down' -or -not $err) { continue }

            Write-Host ''
            Write-Host ("Sign-in failed: {0}" -f $err) -ForegroundColor Red
            if ($desc) { Write-Host ("  {0}" -f $desc) -ForegroundColor DarkGray }
            if ($err -eq 'expired_token')  { Write-Host 'The code timed out. Run it again.' -ForegroundColor Yellow }
            if ($err -eq 'access_denied')  { Write-Host 'Sign-in was declined, or consent has not been granted.' -ForegroundColor Yellow }
            if ($err -eq 'invalid_client') { Write-Host '"Allow public client flows" is not enabled on the app registration.' -ForegroundColor Yellow }
            if ($err -eq 'invalid_grant')  { Write-Host 'The account signed in, but this app cannot use that grant. Check the delegated permissions and consent.' -ForegroundColor Yellow }
            exit 1
        }
    }

    if (-not $tok) { Write-Host 'Timed out waiting for sign-in.' -ForegroundColor Red; exit 1 }
    if (-not $tok.refresh_token) {
        Write-Host 'Signed in, but Microsoft did not return a refresh token.' -ForegroundColor Red
        Write-Host 'The offline_access delegated permission is missing. Ask IT to add it.' -ForegroundColor Yellow
        exit 1
    }

    Save-Refresh $tok.refresh_token

    # Confirm the token actually opens OneNote before declaring success - the
    # whole point of this exercise is the thing that used to return 401.
    try {
        Invoke-RestMethod -Headers @{ Authorization = "Bearer $($tok.access_token)" } `
            -Uri 'https://graph.microsoft.com/v1.0/me/onenote/notebooks' -ErrorAction Stop | Out-Null
        Write-Host ''
        Write-Host 'Signed in, credential stored, and OneNote verified.' -ForegroundColor Green
    }
    catch {
        Write-Host ''
        Write-Host 'Signed in and credential stored, but the OneNote test call failed.' -ForegroundColor Yellow
        Write-Host 'Most likely the DELEGATED Notes.ReadWrite.All permission has not been' -ForegroundColor Yellow
        Write-Host 'granted and consented. Ask IT to add it, then run this again.' -ForegroundColor Yellow
        exit 1
    }

    Write-Host ''
    Write-Host 'Reminder: the service account must be a MEMBER of every case group,' -ForegroundColor Cyan
    Write-Host 'or it cannot see that case notebook.' -ForegroundColor Cyan
    Write-Host ''
    exit 0
}

# ----------------------------------------------------------------- status ----
if ($Status) {
    if (-not (Test-Path $RefreshFile)) {
        Write-Host 'No stored credential (refresh.dat is missing).' -ForegroundColor Yellow
        Show-SignInHelp
        exit 1
    }
    Write-Host ("Credential stored: {0}" -f (Get-Item $RefreshFile).LastWriteTime) -ForegroundColor Gray
}

# ------------------------------------------------------- silent refresh ------
$rt = $null
try { $rt = Read-Refresh }
catch {
    Write-Host 'Stored credential could not be decrypted.' -ForegroundColor Red
    Write-Host 'DPAPI ties it to one Windows account on one machine, so this happens if' -ForegroundColor Yellow
    Write-Host 'the file was copied from elsewhere or the job runs as a different user.' -ForegroundColor Yellow
    Show-SignInHelp
    exit 1
}

if (-not $rt) {
    Write-Host 'No stored credential for OneNote.' -ForegroundColor Red
    Show-SignInHelp
    exit 1
}

try {
    $new = Invoke-TokenRequest @{ client_id     = $ClientId
                                  grant_type    = 'refresh_token'
                                  refresh_token = $rt
                                  scope         = $Scope }
}
catch {
    # Read the body before naming a cause. The four "usual causes" below are
    # guesses; error_description is the fact. The 27 Aug sign-in failure was
    # diagnosed three different wrong ways for exactly this reason.
    $body = ''
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $body = $_.ErrorDetails.Message }
    if (-not $body -and $_.Exception.Response) {
        try { $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $body = $sr.ReadToEnd(); $sr.Close() } catch { }
    }
    if ($body) {
        $ecode = ''; $edesc = ''
        try { $j = $body | ConvertFrom-Json; $ecode = $j.error; $edesc = $j.error_description } catch { $edesc = $body }
        Write-Host ''
        Write-Host ("ENTRA SAID: {0}" -f $ecode) -ForegroundColor Magenta
        Write-Host $edesc -ForegroundColor Magenta
        Write-Host ''
    }
    Write-Host 'The stored credential is no longer valid.' -ForegroundColor Red
    Write-Host 'Usual causes: the service account password was changed, the account was' -ForegroundColor Yellow
    Write-Host 'disabled, a Conditional Access policy now challenges it, or nothing has' -ForegroundColor Yellow
    Write-Host 'run for 90 days.' -ForegroundColor Yellow
    Show-SignInHelp
    exit 1
}

# Microsoft rotates the refresh token on every use. Storing the new one is what
# makes this self-renewing; skip it and the credential dies at the 90-day mark.
if ($new.refresh_token) { Save-Refresh $new.refresh_token }

if ($Status) {
    Write-Host 'Credential is valid and was renewed.' -ForegroundColor Green
    exit 0
}

Write-Output $new.access_token
# Be explicit: callers that test $LASTEXITCODE need it set on the success path too.
exit 0
