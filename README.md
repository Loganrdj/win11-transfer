# Win11 Transfer Scripts

Two PowerShell scripts for migrating a user from an old Windows 10/11 PC to a
new Windows 11 PC. Both scripts must run **on Windows** (PowerShell 5.1+ or
PowerShell 7+) — they are not bash and cannot run under WSL/macOS/Linux, since
they rely on the registry, WMI/CIM, and browser-native APIs.

## 1. On the OLD PC: `Export-Win11Transfer.ps1`

Run as the user being migrated (elevation optional, but gets a fuller
installed-apps list). It builds `Desktop\ITWin11Transfer\` containing:

| Path | Contents |
|---|---|
| `Reports\InstalledApplications.csv` | Names of desktop apps from the uninstall registry (64-bit, 32-bit, per-user) - name only |
| `Reports\InstalledApps_UWP.csv` | Names of Store/UWP apps - name only |
| `Reports\Printers.csv` | All installed printers |
| `Reports\MappedDrives_FileServerShortcuts.csv` | SMB mappings, persistent drive letters, and Desktop/Documents `.lnk` shortcuts pointing at `\\server\share` |
| `Reports\SharePointLibraries.csv` | OneDrive-synced SharePoint/Teams libraries + relevant Quick Access pins |
| `Reports\BookmarkSources.csv` | Manifest consumed by the import script |
| `Reports\PasswordExportManifest.csv` | Which browsers' password exports were completed |
| `TempDownloads\` | Mirror of the user's Downloads folder |
| `TempBookmarks\` | Bookmarks JSON (Chromium browsers) / `places.sqlite` (Firefox) per profile |
| `TempPasswords\` | CSV exports from each browser's **built-in** password exporter — the script opens each installed browser and types its password-settings URL into the address bar (typed, not passed as a launch argument, since Chromium ignores internal `chrome://`/`edge://` URLs handed to an already-running instance that way), then you click Export manually |
| `TempOneNote\` | Any `.one` / `.onetoc2` files found under Documents, Desktop, OneDrive folders, and the OneNote local caches |

Because everything lands under `Desktop\`, **OneDrive Known Folder Move
carries the whole folder to the new PC automatically** — no manual zip/copy
step required.

**Security note:** `TempPasswords` contains plaintext credentials once
exported. Delete it (see step 3 below) as soon as the Keeper import is
confirmed.

**Slow machine?** The address-bar navigation paces itself off `-NavigationDelayMs`
(default `1200`). If a browser opens but the typed URL doesn't land correctly on an
older/loaded machine, try a larger value:

```powershell
.\Export-Win11Transfer.ps1 -NavigationDelayMs 2500
```

## 2. On the NEW PC: `Import-Win11Transfer.ps1`

Run after signing into OneDrive and confirming `Desktop\ITWin11Transfer` has
synced down.

```powershell
.\Import-Win11Transfer.ps1 -Browser Edge -RemovePasswordFilesAfterConfirm
```

- `-Browser` (required): `Chrome`, `Edge`, `Opera`, `OperaGX`, `Brave`, or `Vivaldi`.
- `-RemovePasswordFilesAfterConfirm` (optional): after you confirm the Keeper
  import worked, deletes the plaintext CSVs in `TempPasswords`.
- Run **as Administrator** to force-install the Keeper extension machine-wide
  (HKLM policy). Without elevation it falls back to a current-user-only
  install (HKCU) — fine for a single-user machine.

What it does:
1. Closes the target browser and replaces its `Bookmarks` file (or guides you
   through the Firefox `places.sqlite` restore, which can't be automated the
   same way).
2. Writes an `ExtensionInstallForcelist` policy so Keeper Password Manager
   installs automatically on next browser launch, then restarts the browser.
   Verify the `KeeperExtensionId` default (`bfogiafebfohielmmehodmfbbebbbpei`)
   against the current Chrome Web Store listing before relying on it — the
   script prints the ID it used so you can confirm the install matches.
3. Opens Keeper's vault and `TempPasswords` in File Explorer, then shows
   step-by-step instructions. **Logging into Keeper (with 2FA) and the actual
   CSV import can't be scripted** — that part is manual by design.

Downloads and OneNote files need **no import step** — they arrived on the new
PC's Desktop automatically via OneDrive.

## Not automated (manual verification only)

- **Outlook** signature/calendar sync — out of scope per request.
- **OneNote sync** — notebooks are cloud-native; the export script grabs a
  local safety-net copy, but you must still open OneNote on the new PC and
  confirm every notebook listed under *File > Open* shows a synced (green
  checkmark) state.
- **Opera extension policy support** is inconsistent across versions; if the
  forced install doesn't appear, install Keeper manually from the Chrome Web
  Store inside Opera.
