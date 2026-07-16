#Requires -Version 5.1
<#
    Export-Win11Transfer.ps1

    Run on the OLD machine (Windows 10/11) before a hardware refresh.
    Collects everything into Desktop\ITWin11Transfer so OneDrive Known
    Folder Move carries it to the new PC without any manual copying.

    Produces:
      ITWin11Transfer\Reports\InstalledApplications.csv
      ITWin11Transfer\Reports\InstalledApps_UWP.csv
      ITWin11Transfer\Reports\Printers.csv
      ITWin11Transfer\Reports\MappedDrives_FileServerShortcuts.csv
      ITWin11Transfer\Reports\SharePointLibraries.csv
      ITWin11Transfer\Reports\BookmarkSources.csv    (manifest used by the import script)
      ITWin11Transfer\Reports\TransferLog.txt
      ITWin11Transfer\TempDownloads\                 (mirror of Downloads)
      ITWin11Transfer\TempBookmarks\                 (Bookmarks/places.sqlite per browser/profile)
      ITWin11Transfer\TempPasswords\                 (browser-native CSV exports, user-driven)
      ITWin11Transfer\TempOneNote\                   (.one / .onetoc2 files found on disk)

    SECURITY NOTE: TempPasswords will contain PLAINTEXT credentials once the
    user exports them. Treat ITWin11Transfer as sensitive for the duration of
    the migration and delete TempPasswords on both machines once the Keeper
    import is confirmed (Import-Win11Transfer.ps1 can do this for you).
#>

[CmdletBinding()]
param(
    # Which Chromium browsers to look for. Add to this list for anything else in use.
    [string[]]$ChromiumBrowsers = @('Chrome', 'Edge', 'Opera', 'OperaGX', 'Brave', 'Vivaldi')
)

$ErrorActionPreference = 'Continue'

$Desktop = [Environment]::GetFolderPath('Desktop')
$Root    = Join-Path $Desktop 'ITWin11Transfer'
$Reports = Join-Path $Root 'Reports'
$Dirs    = @{
    Downloads = Join-Path $Root 'TempDownloads'
    Bookmarks = Join-Path $Root 'TempBookmarks'
    Passwords = Join-Path $Root 'TempPasswords'
    OneNote   = Join-Path $Root 'TempOneNote'
}
foreach ($d in @($Root, $Reports) + $Dirs.Values) {
    New-Item -ItemType Directory -Path $d -Force | Out-Null
}

$LogPath = Join-Path $Reports 'TransferLog.txt'
function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

Write-Log "=== Export-Win11Transfer started for $env:USERNAME on $env:COMPUTERNAME ==="

# ---------------------------------------------------------------------------
# 1. Installed applications (registry, both bitness + per-user, plus UWP)
# ---------------------------------------------------------------------------
Write-Log "Collecting installed applications..."

$uninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

$apps = foreach ($path in $uninstallPaths) {
    Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName } |
        Select-Object @{n='Name';e={$_.DisplayName}},
                      @{n='Version';e={$_.DisplayVersion}},
                      Publisher,
                      InstallDate,
                      @{n='UninstallString';e={$_.UninstallString}}
}
$apps | Sort-Object Name -Unique | Export-Csv -Path (Join-Path $Reports 'InstalledApplications.csv') -NoTypeInformation -Encoding UTF8
Write-Log ("  -> {0} desktop applications found" -f ($apps | Measure-Object).Count)

try {
    Get-AppxPackage | Select-Object Name, PackageFullName, Version, Publisher |
        Export-Csv -Path (Join-Path $Reports 'InstalledApps_UWP.csv') -NoTypeInformation -Encoding UTF8
    Write-Log "  -> UWP/Store app list exported"
} catch {
    Write-Log "  -> Could not enumerate UWP apps: $($_.Exception.Message)"
}

# ---------------------------------------------------------------------------
# 2. Printers
# ---------------------------------------------------------------------------
Write-Log "Collecting printers..."
try {
    Get-Printer | Select-Object Name, DriverName, PortName, Shared, Type, Comment, Location |
        Export-Csv -Path (Join-Path $Reports 'Printers.csv') -NoTypeInformation -Encoding UTF8
    Write-Log ("  -> {0} printers found" -f (Get-Printer | Measure-Object).Count)
} catch {
    Write-Log "  -> Get-Printer failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------------------------
# 3. Mapped drives + file-server shortcuts
# ---------------------------------------------------------------------------
Write-Log "Collecting mapped drives and file-server shortcuts..."

$mappedRows = @()

try {
    Get-SmbMapping | ForEach-Object {
        $mappedRows += [PSCustomObject]@{
            Source     = 'SMB Mapping'
            Name       = $_.LocalPath
            TargetPath = $_.RemotePath
            Status     = $_.Status
        }
    }
} catch { Write-Log "  -> Get-SmbMapping unavailable: $($_.Exception.Message)" }

# Persistent drive mappings (reconnect-at-logon) live under HKCU:\Network\<Letter>
Get-ChildItem 'HKCU:\Network' -ErrorAction SilentlyContinue | ForEach-Object {
    $remote = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).RemotePath
    if ($remote) {
        $mappedRows += [PSCustomObject]@{
            Source     = 'Persistent Registry Mapping'
            Name       = "$($_.PSChildName):"
            TargetPath = $remote
            Status     = 'Reconnect at logon'
        }
    }
}

# .lnk shortcuts on Desktop/Documents/Links that point at a UNC path
$shellCom = New-Object -ComObject WScript.Shell
$shortcutRoots = @(
    $Desktop,
    [Environment]::GetFolderPath('MyDocuments'),
    (Join-Path $env:USERPROFILE 'Links')
) | Where-Object { Test-Path $_ }

foreach ($root in $shortcutRoots) {
    Get-ChildItem -Path $root -Filter '*.lnk' -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $target = $shellCom.CreateShortcut($_.FullName).TargetPath
            if ($target -like '\\*') {
                $mappedRows += [PSCustomObject]@{
                    Source     = 'Desktop/Documents Shortcut'
                    Name       = $_.Name
                    TargetPath = $target
                    Status     = $_.FullName
                }
            }
        } catch {
            Write-Verbose "Could not resolve shortcut target for $($_.FullName): $($_.Exception.Message)"
        }
    }
}

$mappedRows | Export-Csv -Path (Join-Path $Reports 'MappedDrives_FileServerShortcuts.csv') -NoTypeInformation -Encoding UTF8
Write-Log ("  -> {0} mapped drive / file-server shortcut entries found" -f $mappedRows.Count)

# ---------------------------------------------------------------------------
# 4. SharePoint / OneDrive-synced libraries visible in File Explorer
# ---------------------------------------------------------------------------
Write-Log "Collecting SharePoint / OneDrive library links..."

$spRows = @()

# OneDrive records every synced SharePoint/Teams library here.
Get-ChildItem 'HKCU:\Software\SyncEngines\Providers\OneDrive' -ErrorAction SilentlyContinue | ForEach-Object {
    $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
    if ($p.UrlNamespace -or $p.MountPoint) {
        $spRows += [PSCustomObject]@{
            Source      = 'OneDrive Sync Engine'
            DisplayName = $p.DisplayName
            LocalPath   = $p.MountPoint
            LibraryUrl  = $p.UrlNamespace
        }
    }
}

# Quick Access pins that resolve to a SharePoint/OneDrive-for-Business path.
try {
    $shellApp = New-Object -ComObject Shell.Application
    $quickAccess = $shellApp.Namespace('shell:::{679f85cc-0220-4080-b29b-5540cc05aab6}')
    foreach ($item in $quickAccess.Items()) {
        $path = $item.Path
        if ($path -match 'sharepoint\.com' -or $path -match 'OneDrive.*-.*') {
            $spRows += [PSCustomObject]@{
                Source      = 'Quick Access Pin'
                DisplayName = $item.Name
                LocalPath   = $path
                LibraryUrl  = ''
            }
        }
    }
} catch { Write-Log "  -> Could not enumerate Quick Access: $($_.Exception.Message)" }

$spRows | Export-Csv -Path (Join-Path $Reports 'SharePointLibraries.csv') -NoTypeInformation -Encoding UTF8
Write-Log ("  -> {0} SharePoint/OneDrive library references found" -f $spRows.Count)

# ---------------------------------------------------------------------------
# 5. Downloads -> TempDownloads
# ---------------------------------------------------------------------------
Write-Log "Copying Downloads folder..."
$downloadsSrc = Join-Path $env:USERPROFILE 'Downloads'
if (Test-Path $downloadsSrc) {
    robocopy $downloadsSrc $Dirs.Downloads /E /COPY:DAT /R:1 /W:1 /NFL /NDL /NP | Out-Null
    Write-Log "  -> Downloads copied to $($Dirs.Downloads)"
} else {
    Write-Log "  -> No Downloads folder found at $downloadsSrc"
}

# ---------------------------------------------------------------------------
# 6. Browser bookmarks (all Chromium-based browsers + Firefox)
# ---------------------------------------------------------------------------
Write-Log "Collecting browser bookmarks..."

$chromiumRoots = @{
    Chrome   = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
    Edge     = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data'
    Opera    = Join-Path $env:APPDATA 'Opera Software\Opera Stable'
    OperaGX  = Join-Path $env:APPDATA 'Opera Software\Opera GX Stable'
    Brave    = Join-Path $env:LOCALAPPDATA 'BraveSoftware\Brave-Browser\User Data'
    Vivaldi  = Join-Path $env:LOCALAPPDATA 'Vivaldi\User Data'
}

$bookmarkManifest = @()

foreach ($browser in $ChromiumBrowsers) {
    $userDataRoot = $chromiumRoots[$browser]
    if (-not $userDataRoot -or -not (Test-Path $userDataRoot)) { continue }

    # Opera keeps its single profile directly at the root; Chrome/Edge/Brave/Vivaldi use
    # Default / Profile 1 / Profile 2 ... subfolders.
    $profileDirs = if ($browser -like 'Opera*') {
        @($userDataRoot)
    } else {
        Get-ChildItem -Path $userDataRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq 'Default' -or $_.Name -like 'Profile *' }
    }

    foreach ($profileDir in $profileDirs) {
        $profilePath = if ($profileDir -is [string]) { $profileDir } else { $profileDir.FullName }
        $profileName = if ($profileDir -is [string]) { 'Default' } else { $profileDir.Name }
        $bookmarksFile = Join-Path $profilePath 'Bookmarks'
        if (Test-Path $bookmarksFile) {
            $destName = "${browser}__${profileName}__Bookmarks.json"
            Copy-Item -Path $bookmarksFile -Destination (Join-Path $Dirs.Bookmarks $destName) -Force
            $bookmarkManifest += [PSCustomObject]@{
                Browser      = $browser
                Profile      = $profileName
                OriginalPath = $bookmarksFile
                ExportedFile = $destName
            }
        }
    }
}

# Firefox: copy places.sqlite (bookmarks + history) and any compressed JSON backups.
$ffProfilesIni = Join-Path $env:APPDATA 'Mozilla\Firefox\profiles.ini'
if (Test-Path $ffProfilesIni) {
    $ffRoot = Join-Path $env:APPDATA 'Mozilla\Firefox'
    Get-ChildItem -Path (Join-Path $ffRoot 'Profiles') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $places = Join-Path $_.FullName 'places.sqlite'
        if (Test-Path $places) {
            $destName = "Firefox__$($_.Name)__places.sqlite"
            Copy-Item -Path $places -Destination (Join-Path $Dirs.Bookmarks $destName) -Force
            $bookmarkManifest += [PSCustomObject]@{
                Browser      = 'Firefox'
                Profile      = $_.Name
                OriginalPath = $places
                ExportedFile = $destName
            }
        }
        $backups = Join-Path $_.FullName 'bookmarkbackups'
        if (Test-Path $backups) {
            $latest = Get-ChildItem $backups -Filter '*.jsonlz4' -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($latest) {
                Copy-Item -Path $latest.FullName -Destination (Join-Path $Dirs.Bookmarks "Firefox__$($_.Name)__$($latest.Name)") -Force
            }
        }
    }
}

$bookmarkManifest | Export-Csv -Path (Join-Path $Reports 'BookmarkSources.csv') -NoTypeInformation -Encoding UTF8
Write-Log ("  -> {0} bookmark files exported" -f $bookmarkManifest.Count)

# ---------------------------------------------------------------------------
# 7. Browser passwords - there is no CLI export, so we drive the user to each
#    browser's built-in "Export Passwords" screen and wait for confirmation.
# ---------------------------------------------------------------------------
Write-Log "Starting guided password export..."

Add-Type -AssemblyName System.Windows.Forms | Out-Null

$passwordTargets = @(
    @{ Browser = 'Chrome';  Exe = 'chrome';  Url = 'chrome://password-manager/passwords' }
    @{ Browser = 'Edge';    Exe = 'msedge';  Url = 'edge://settings/passwords' }
    @{ Browser = 'Opera';   Exe = 'opera';   Url = 'opera://settings/passwords' }
    @{ Browser = 'Brave';   Exe = 'brave';   Url = 'brave://settings/passwords' }
    @{ Browser = 'Vivaldi'; Exe = 'vivaldi'; Url = 'vivaldi://settings/passwords' }
    @{ Browser = 'Firefox'; Exe = 'firefox'; Url = 'about:logins' }
)

$passwordManifest = @()

foreach ($target in $passwordTargets) {
    $installed = (Get-Command $target.Exe -ErrorAction SilentlyContinue) -or
                 ($chromiumRoots[$target.Browser] -and (Test-Path $chromiumRoots[$target.Browser]))
    if (-not $installed) { continue }

    $expectedFile = Join-Path $Dirs.Passwords "$($target.Browser)_Passwords.csv"
    try {
        Start-Process $target.Exe -ArgumentList $target.Url -ErrorAction Stop
    } catch {
        Write-Log "  -> Could not launch $($target.Browser): $($_.Exception.Message)"
        continue
    }

    [System.Windows.Forms.MessageBox]::Show(
        "$($target.Browser) is now open to its saved-password page.`n`n" +
        "1. Click Export / Download passwords (you may need to re-enter the Windows login password).`n" +
        "2. Save the CSV file as exactly:`n   $expectedFile`n`n" +
        "Click OK once the file has been saved.",
        "Export $($target.Browser) Passwords",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $found = Test-Path $expectedFile
    $passwordManifest += [PSCustomObject]@{
        Browser  = $target.Browser
        Expected = $expectedFile
        Found    = $found
    }
    Write-Log "  -> $($target.Browser): file found = $found"
}

$passwordManifest | Export-Csv -Path (Join-Path $Reports 'PasswordExportManifest.csv') -NoTypeInformation -Encoding UTF8

# ---------------------------------------------------------------------------
# 8. OneNote files (OneNote is cloud-native, so this is a local safety-net
#    copy, not a substitute for confirming notebooks re-open on the new PC)
# ---------------------------------------------------------------------------
Write-Log "Collecting OneNote files..."

$oneNoteRoots = @(
    (Join-Path $env:USERPROFILE 'Documents'),
    $Desktop,
    (Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe\LocalState'),
    (Join-Path $env:LOCALAPPDATA 'Microsoft\OneNote')
)
# Any "OneDrive - <Company>" or plain "OneDrive" folders under the user profile.
Get-ChildItem -Path $env:USERPROFILE -Directory -Filter 'OneDrive*' -ErrorAction SilentlyContinue |
    ForEach-Object { $oneNoteRoots += $_.FullName }

$oneNoteCount = 0
foreach ($root in ($oneNoteRoots | Select-Object -Unique | Where-Object { Test-Path $_ })) {
    Get-ChildItem -Path $root -Recurse -Include '*.one', '*.onetoc2' -ErrorAction SilentlyContinue -File |
        ForEach-Object {
            $relative = $_.FullName.Substring($root.Length).TrimStart('\')
            $rootLabel = Split-Path $root -Leaf
            $dest = Join-Path $Dirs.OneNote (Join-Path $rootLabel $relative)
            New-Item -ItemType Directory -Path (Split-Path $dest) -Force | Out-Null
            Copy-Item -Path $_.FullName -Destination $dest -Force -ErrorAction SilentlyContinue
            $oneNoteCount++
        }
}
Write-Log ("  -> {0} OneNote files copied to TempOneNote" -f $oneNoteCount)
Write-Log "  -> NOTE: notebooks live in OneDrive/SharePoint by design. After sign-in on the new PC, open OneNote and confirm all notebooks listed under File > Open appear and sync (green checkmark) before treating this as verified."

# ---------------------------------------------------------------------------
Write-Log "=== Export complete. Everything is under: $Root ==="
Write-Log "Because this sits on the Desktop, OneDrive Known Folder Move will carry it to the new PC - no manual copy needed."
Write-Host "`nDone. Review $LogPath for a full summary." -ForegroundColor Green
