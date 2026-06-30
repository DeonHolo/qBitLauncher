<p align="center">
  <img src="https://i.imgur.com/0epFbuH.png" alt="qBitLauncher Logo" width="128">
</p>

# qBitLauncher

An automated post-download workflow manager for qBittorrent and Windows Explorer. 

qBitLauncher eliminates the need to manually navigate directories and extract archives after a torrent completes. It automatically intercepts finished downloads, handles extraction, and surfaces the primary executables in a clean GUI.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

### What It Does
- **Automated Extraction Intercept**: Hooks into qBittorrent's "Run on torrent finished" event to instantly prompt extraction for `.zip`, `.rar`, `.7z`, `.iso`, and `.img` archives.
- **Smart Executable Discovery**: Recursively scans extracted directories to surface primary binaries (`.exe`, `.bat`), sorting them logically by directory depth.
- **One-Click Cleanup**: Interfaces with qBittorrent's Local Web API to seamlessly remove torrents and optionally purge downloaded source data once installation is complete.
- **Explorer Integration**: Integrates directly into the Windows Shift+Right-Click context menu for on-demand execution on any local folder or archive.


<p align="center">
  <img src="https://i.imgur.com/Y3974UK.png" alt="qBitLauncher GUI Screenshot" width="650">
</p>

## Features

- 🗜️ **External extraction** - Extracts ZIP, RAR, 7z, ISO, and IMG archives through the installed 7-Zip or WinRAR GUI
- 📂 **Custom extraction path** - Choose where to extract with modern folder picker
- 🔍 **Smart executable discovery** - Finds .exe files with icons, sorted by folder depth
- ✏️ **Inline rename** - Rename executables directly (F2, double-click, or right-click)
- 📋 **Activity Log** - Real-time log panel showing all actions
- 🎨 **Themed GUI** - Dracula dark theme with Light mode option
- 🛡️ **Admin launch** - Run selected executables with UAC elevation
- 🔔 **Action feedback** - Themed dialogs and sound effects for all actions
- 📦 **Auto-update** - Checks for updates on startup
- 🧹 **qBittorrent cleanup** - Remove the torrent from qBittorrent after extraction/installation, with optional data deletion

## Requirements

- Windows 10/11
- PowerShell 5.1+
- [7-Zip](https://www.7-zip.org/) (recommended) or [WinRAR](https://www.win-rar.com/)

## Installation

The easiest way to install qBitLauncher is to download the compiled executable.

1. **Download the latest `.exe`** from the [GitHub Releases](https://github.com/DeonHolo/qBitLauncher/releases/latest) page.
2. **Run it** to launch the Setup window and install it to a permanent location.
   > **Note:** Because this is an open-source tool without an expensive code-signing certificate, Windows SmartScreen may show an "unrecognized app" warning. Click **More Info** -> **Run anyway**.
3. **Click "Install to qBittorrent"** from the Dashboard.

**For Developers (Source Code):**
```bash
git clone https://github.com/DeonHolo/qBitLauncher.git
```
*You can run `qBitLauncher.ps1` directly, or compile your own `.exe` using the included `Compile.ps1` script.*

### qBittorrent Integration

The easiest way to integrate with qBittorrent is to use the built-in Dashboard:
1. Double-click `qBitLauncher.exe` (or run `qBitLauncher.ps1` directly).
2. Click **Install to qBittorrent** (make sure qBittorrent is closed first).

**Manual Setup:**
1. qBittorrent → **Tools** → **Options** → **Downloads**
2. Under **"Run external program"**, enable **"Run on torrent finished"**
3. Set command:
   ```
   powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\qBitLauncher.ps1" "%F" "%I" "%N"
   ```
> **Note:** If using the compiled `.exe`, the manual command is just `"C:\path\to\qBitLauncher.exe" "%F" "%I" "%N"`.
> `%I` passes the torrent hash for qBittorrent cleanup. `%N` passes the torrent name for clearer confirmation dialogs.

### Context Menu Integration (Optional)

Add "Open with qBitLauncher" to your **Shift+Right-click** context menu for folders and archive files.

The easiest way to integrate with context menus is via the Dashboard:
1. Double-click `qBitLauncher.exe` (or run `qBitLauncher.ps1` directly).
2. Click **Add Shift+Right-Click Context Menus**.

*(To remove the integration later, you can simply click **Remove Context Menus** in the Dashboard).*

**Supported locations:**
- Folders
- Archive files: `.zip`, `.rar`, `.7z`, `.iso`, `.img`

> **Note:** The menu item only appears with **Shift+Right-click** to keep your regular context menu clean.

## GUI Actions

| Button | Action | Shortcut |
|--------|--------|----------|
| **Run** | Launch as admin | `Alt+R` |
| **Shortcut** | Create desktop shortcut | `Alt+S` |
| **Open Folder** | Open in Explorer | `Alt+O` |
| **Rename** | Rename executable | `Alt+N`, `F2`, double-click |
| **Settings** | Configure theme | `Alt+T` |
| **Remove Torrent** | Remove from qBittorrent; optionally delete downloaded files | None |
| **Close** | Close window | `Alt+C` |

## Configuration

Edit via **Settings** button or `config.json`:
```json
{
  "Theme": "Dracula",
  "QbittorrentWebUrl": "http://localhost:8080",
  "QbittorrentUsername": "",
  "QbittorrentPassword": ""
}
```

**Themes**: `Dracula`, `Light`

**Remove Torrent** uses qBittorrent's local API. qBitLauncher will offer to enable that API the first time cleanup needs it; you do not need to keep a WebUI page open. If qBittorrent requires authentication for localhost, enter the API username/password in **Settings**.

## Supported Extensions

| Type | Extensions |
|------|------------|
| Runnables | `exe`, `bat`, `cmd` |
| Archives | `iso`, `zip`, `rar`, `7z`, `img` |
| Media | `mp4`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `webm`, `mp3`, `flac`, `wav`, `aac`, `ogg`, `m4a` |

## Logging

- **GUI**: Real-time Activity Log panel
- **File**: `qBitLauncher_log.txt`

## Development, Auto-Versioning & Releases

This project uses automated GitHub Actions workflows to manage versions and releases:

1. **Auto-Versioning (`version-bump.yml`)**: 
   When pushing to the `main` branch, include one of the following keywords in your commit message to automatically bump the version in `qBitLauncher.ps1`:
   - `feat:` - Increments the MINOR version (e.g., 1.2.3 → 1.3.0). Rolls over to a new MAJOR version if MINOR exceeds 9.
   - `fix:` - Increments the PATCH version (e.g., 1.2.3 → 1.2.4). Rolls over to a new MINOR version if PATCH exceeds 9.
   - `BREAKING:` - Increments the MAJOR version (e.g., 1.2.3 → 2.0.0).

2. **Auto-Releases (`build-release.yml`)**:
   Whenever a new version tag (e.g., `v3.1.0`) is pushed, GitHub Actions automatically uses `PS2EXE` to compile the `.ps1` script (along with the `.ico` logo) into a standalone `qBitLauncher.exe` and attaches it to a new GitHub Release.

**The Built-in Updater**:
The script's built-in auto-updater seamlessly handles both raw `.ps1` and compiled `.exe` users. 
- For `.ps1` users, it fetches the latest raw code from the repository.
- For `.exe` users, it pulls the latest `qBitLauncher.exe` directly from the GitHub Releases page, using a background script to seamlessly hot-swap the executable.

## License

MIT License - see [LICENSE](LICENSE)
