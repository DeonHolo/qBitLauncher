<p align="center">
  <img src="https://i.imgur.com/0epFbuH.png" alt="qBitLauncher Logo" width="128">
</p>

# qBitLauncher

A PowerShell post-download handler for qBittorrent that automatically extracts archives and launches executables with a clean GUI.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

- 🗜️ **Auto-extraction** - Extracts ZIP, RAR, 7z, ISO, and IMG archives
- 📂 **Custom extraction path** - Choose where to extract with modern folder picker
- 📊 **Progress bar** - Visual feedback during extraction with cancel support
- 🔍 **Smart executable discovery** - Finds .exe files with icons, sorted by folder depth
- ✏️ **Inline rename** - Rename executables directly (F2, double-click, or right-click)
- 📋 **Activity Log** - Real-time log panel showing all actions
- 🎨 **Themed GUI** - Dracula dark theme with Light mode option
- 🛡️ **Run as Administrator** - One-click UAC elevation for installers
- 🔔 **Toast notifications** - Windows notifications for all actions
- 📦 **Auto-update** - Checks for updates on startup

## Requirements

- Windows 10/11
- PowerShell 5.1+
- [7-Zip](https://www.7-zip.org/) (recommended) or [WinRAR](https://www.win-rar.com/)

## Installation

**Clone:**
```bash
git clone https://github.com/DeonHolo/qBitLauncher.git
```

**Or download** `qBitLauncher.ps1` from the [repo](https://github.com/DeonHolo/qBitLauncher).

### qBittorrent Integration

1. qBittorrent → **Tools** → **Options** → **Downloads**
2. Enable **"Run external program on torrent finished"**
3. Set command:
   ```
   powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\qBitLauncher.ps1" "%F"
   ```

## Usage

When a torrent completes:
- **Archives**: Extract → show executables
- **Executables**: Show selection GUI
- **Media**: Open containing folder

### GUI Actions

| Button | Action | Shortcut |
|--------|--------|----------|
| **Run** | Launch as admin | `Alt+R` |
| **Shortcut** | Create desktop shortcut | `Alt+S` |
| **Open Folder** | Open in Explorer | `Alt+O` |
| **Rename** | Rename executable | `Alt+N`, `F2`, double-click |
| **Settings** | Configure theme | `Alt+T` |
| **Close** | Close window | `Alt+C` |

## Configuration

Edit via **Settings** button or `config.json`:
```json
{
  "Theme": "Dracula",
  "Notifications": true
}
```

**Themes**: `Dracula`, `Light`

## Supported Extensions

| Type | Extensions |
|------|------------|
| Archives | `iso`, `zip`, `rar`, `7z`, `img` |
| Media | `mp4`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `webm`, `mp3`, `flac`, `wav`, `aac`, `ogg`, `m4a` |

## Logging

- **GUI**: Real-time Activity Log panel
- **File**: `qBitLauncher_log.txt`

## License

MIT License - see [LICENSE](LICENSE)
