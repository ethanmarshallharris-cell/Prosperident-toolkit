<#
    Setup-Hotkey.ps1
    Creates a Start menu shortcut "Export OneNote Page" with the Ctrl+Alt+E hotkey.
    Run once (via Setup Hotkey.bat). Re-run if you move this folder.
    Results are logged to Setup-Hotkey.log in this folder.
#>

$Hotkey = 'CTRL+ALT+E'    # change here if you prefer a different key combination

$ErrorActionPreference = 'Stop'
$ScriptDir = $PSScriptRoot
$LogFile   = Join-Path $ScriptDir 'Setup-Hotkey.log'

function Write-Log([string]$Message) {
    $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    Write-Host $Message
}

try {
    Write-Log '==================== Setup started ===================='
    $exporter = Join-Path $ScriptDir 'Export-OneNotePage.ps1'
    if (-not (Test-Path $exporter)) { throw "Export-OneNotePage.ps1 was not found in $ScriptDir" }

    # Clear the "downloaded from the internet" flag so Windows does not block the scripts
    Get-ChildItem -Path $ScriptDir -File | Unblock-File
    Write-Log 'Files unblocked.'

    $programs = [Environment]::GetFolderPath('Programs')
    $lnkPath  = Join-Path $programs 'Export OneNote Page.lnk'

    $shell = New-Object -ComObject WScript.Shell
    $lnk = $shell.CreateShortcut($lnkPath)
    $lnk.TargetPath       = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $lnk.Arguments        = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$exporter`""
    $lnk.WorkingDirectory = $ScriptDir
    $lnk.WindowStyle      = 7
    $lnk.Hotkey           = $Hotkey
    $lnk.Description      = 'Export the open OneNote page and its subpages to Word'

    $oneNoteExe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ONENOTE.EXE' -ErrorAction SilentlyContinue).'(default)'
    if ($oneNoteExe) { $lnk.IconLocation = ($oneNoteExe.Trim('"') + ',0') }

    $lnk.Save()
    Write-Log "Shortcut created: $lnkPath"
    Write-Log "Hotkey: $Hotkey"
    Write-Log '==================== Setup finished ===================='
    Write-Host ''
    Write-Host 'Done. Open a page in OneNote and press Ctrl+Alt+E.'
}
catch {
    Write-Log ("ERROR: {0}" -f $_.Exception.Message)
}
