#Requires -Version 5.1
<#
    Import-Win11Transfer.ps1

    Run on the NEW Windows 11 machine after the user has signed into OneDrive
    and Desktop\ITWin11Transfer has synced down from the old PC.

    Does three things:
      1. Restores bookmarks into the chosen browser (from TempBookmarks).
      2. Force-installs the Keeper Password Manager browser extension.
      3. Walks the user through logging into Keeper and importing the CSVs
         from TempPasswords (login/2FA cannot be scripted, by design).

    Downloads and OneNote files need no action here - they're already sitting
    on the Desktop under ITWin11Transfer because OneDrive Known Folder Move
    carried them over automatically.

    Run elevated (Run as Administrator) if you want the Keeper extension
    force-installed machine-wide (HKLM). Without elevation it falls back to
    a current-user-only install (HKCU), which still works for a single-user
    machine.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge', 'Opera', 'OperaGX', 'Brave', 'Vivaldi')]
    [string]$Browser,

    # Keeper's published Chrome Web Store extension ID. Verify this against
    # https://chromewebstore.google.com/detail/keeper-password-manager-d/ before
    # relying on it - Web Store IDs are stable but should be confirmed once.
    [string]$KeeperExtensionId = 'bfogiafebfohielmmehodmfbbebbbpei',

    [switch]$RemovePasswordFilesAfterConfirm
)

$ErrorActionPreference = 'Continue'

$Desktop = [Environment]::GetFolderPath('Desktop')
$Root    = Join-Path $Desktop 'ITWin11Transfer'
$Reports = Join-Path $Root 'Reports'
$TempBookmarks = Join-Path $Root 'TempBookmarks'
$TempPasswords = Join-Path $Root 'TempPasswords'

if (-not (Test-Path $Root)) {
    Write-Host "ITWin11Transfer wasn't found on the Desktop yet." -ForegroundColor Red
    Write-Host "Make sure OneDrive has finished syncing before running this script." -ForegroundColor Red
    exit 1
}

$LogPath = Join-Path $Reports 'ImportLog.txt'
function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Write-Log "=== Import-Win11Transfer started for $env:USERNAME on $env:COMPUTERNAME (elevated: $isElevated) ==="

# Chromium browser process names + per-vendor "User Data" root + policy registry path.
$browserInfo = @{
    Chrome  = @{ Process = 'chrome';   UserData = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data';               PolicyPath = 'SOFTWARE\Policies\Google\Chrome' }
    Edge    = @{ Process = 'msedge';   UserData = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data';              PolicyPath = 'SOFTWARE\Policies\Microsoft\Edge' }
    Opera   = @{ Process = 'opera';    UserData = Join-Path $env:APPDATA 'Opera Software\Opera Stable';                PolicyPath = 'SOFTWARE\Policies\Opera Software\Opera' }
    OperaGX = @{ Process = 'opera';    UserData = Join-Path $env:APPDATA 'Opera Software\Opera GX Stable';             PolicyPath = 'SOFTWARE\Policies\Opera Software\Opera' }
    Brave   = @{ Process = 'brave';    UserData = Join-Path $env:LOCALAPPDATA 'BraveSoftware\Brave-Browser\User Data'; PolicyPath = 'SOFTWARE\Policies\BraveSoftware\Brave' }
    Vivaldi = @{ Process = 'vivaldi';  UserData = Join-Path $env:LOCALAPPDATA 'Vivaldi\User Data';                     PolicyPath = 'SOFTWARE\Policies\Vivaldi' }
}
$info = $browserInfo[$Browser]

# ---------------------------------------------------------------------------
# 1. Restore bookmarks
# ---------------------------------------------------------------------------
Write-Log "Restoring bookmarks for $Browser..."

$manifestPath = Join-Path $Reports 'BookmarkSources.csv'
if (-not (Test-Path $manifestPath)) {
    Write-Log "  -> No BookmarkSources.csv manifest found - skipping bookmark restore."
} else {
    $entries = Import-Csv $manifestPath | Where-Object { $_.Browser -eq $Browser }
    if (-not $entries) {
        Write-Log "  -> No exported bookmarks found for $Browser."
    } elseif ($Browser -eq 'Firefox') {
        Write-Log "  -> Firefox bookmarks require a manual restore: in about:profiles locate the profile folder, close Firefox, and replace places.sqlite with the copy in $TempBookmarks (or use 'Restore from backup' in the Bookmarks Library with the .jsonlz4 file)."
    } else {
        Write-Log "  -> Closing $Browser so its profile files aren't locked..."
        Stop-Process -Name $info.Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        foreach ($entry in $entries) {
            $sourceFile = Join-Path $TempBookmarks $entry.ExportedFile
            if (-not (Test-Path $sourceFile)) {
                Write-Log "  -> Missing expected file $sourceFile, skipping."
                continue
            }
            $targetProfileDir = Join-Path $info.UserData $entry.Profile
            if (-not (Test-Path $targetProfileDir)) {
                # Fresh machine may not have created "Profile 1" etc. yet - default to Default.
                $targetProfileDir = Join-Path $info.UserData 'Default'
                New-Item -ItemType Directory -Path $targetProfileDir -Force | Out-Null
            }
            $targetBookmarksFile = Join-Path $targetProfileDir 'Bookmarks'
            if (Test-Path $targetBookmarksFile) {
                Copy-Item $targetBookmarksFile "$targetBookmarksFile.bak" -Force
            }
            Copy-Item $sourceFile $targetBookmarksFile -Force
            Write-Log "  -> Restored $($entry.Profile) bookmarks into $targetBookmarksFile"
        }
    }
}

# ---------------------------------------------------------------------------
# 2. Force-install Keeper Password Manager extension
# ---------------------------------------------------------------------------
Write-Log "Force-installing Keeper Password Manager extension for $Browser..."
Write-Log "  -> Using extension ID $KeeperExtensionId - verify this is still correct on the Web Store if the install doesn't appear."

$hive = if ($isElevated) { 'HKLM:' } else { 'HKCU:' }
$policyKey = Join-Path $hive $info.PolicyPath
$forceListKey = Join-Path $policyKey 'ExtensionInstallForcelist'

New-Item -Path $policyKey -Force | Out-Null
New-Item -Path $forceListKey -Force | Out-Null
Set-ItemProperty -Path $forceListKey -Name '1' -Value "$KeeperExtensionId;https://clients2.google.com/service/update2/crx"

if ($Browser -like 'Opera*') {
    Write-Log "  -> Opera's enterprise policy support is inconsistent. If the extension doesn't appear after a restart, open the Keeper Chrome Web Store page in Opera and click 'Install Extension' manually."
}

Write-Log "  -> Policy written to $forceListKey. Restarting $Browser so it picks up the new policy..."
Stop-Process -Name $info.Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Start-Process $info.Process -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5

# ---------------------------------------------------------------------------
# 3. Guide the user through Keeper login + password import (cannot be scripted)
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms | Out-Null

$passwordFiles = Get-ChildItem -Path $TempPasswords -Filter '*_Passwords.csv' -ErrorAction SilentlyContinue
if (-not $passwordFiles) {
    Write-Log "No password export files found in $TempPasswords - nothing to import."
} else {
    Write-Log ("Found {0} password export file(s) to import into Keeper." -f $passwordFiles.Count)

    Start-Process explorer.exe $TempPasswords
    Start-Process $info.Process -ArgumentList 'https://keepersecurity.com/vault/' -ErrorAction SilentlyContinue

    [System.Windows.Forms.MessageBox]::Show(
        "Keeper is open in $Browser and File Explorer is pointed at:`n$TempPasswords`n`n" +
        "This part can't be automated because it needs your Keeper login (and 2FA):`n`n" +
        "1. Log into (or create) the Keeper vault.`n" +
        "2. Go to Settings > Import.`n" +
        "3. Choose 'Generic CSV' (or the matching browser format) and select each CSV in the folder above, one at a time.`n" +
        "4. Confirm every entry imported correctly.`n`n" +
        "Click OK once you've finished the import.",
        "Import Passwords into Keeper",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    if ($RemovePasswordFilesAfterConfirm) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Delete the plaintext CSV files in TempPasswords now?`n`nOnly do this after confirming every entry imported into Keeper successfully.",
            "Delete Plaintext Password Files?",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-Item -Path (Join-Path $TempPasswords '*') -Force -ErrorAction SilentlyContinue
            Write-Log "  -> TempPasswords contents deleted after confirmed import."
        } else {
            Write-Log "  -> User declined deletion - TempPasswords still contains plaintext credentials. Delete manually once confirmed."
        }
    } else {
        Write-Log "  -> Reminder: TempPasswords still contains plaintext credentials. Delete it (on both this PC and via OneDrive) once the Keeper import is confirmed."
    }
}

Write-Log "=== Import complete ==="
Write-Host "`nDone. Review $LogPath for a full summary." -ForegroundColor Green
Write-Host "Reminder: Downloads and OneNote files already arrived on the Desktop via OneDrive - no action needed there." -ForegroundColor Yellow
