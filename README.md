# qBitLauncher

A PowerShell post-download handler for qBittorrent that automatically extracts archives and launches executables with a clean GUI.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

- 🗜️ **Auto-extraction** - Extracts ZIP, RAR, 7z, ISO, and IMG archives
- 🔍 **Smart executable discovery** - Finds all .exe files and sorts by folder depth
- 🎨 **Themed GUI** - Dark mode UI matching qBittorrent's style
- 🛡️ **Run as Administrator** - One-click UAC elevation for installers
- 🔔 **Toast notifications** - Windows notifications for all actions
- 📁 **Multiple actions** - Run, create desktop shortcut, or open folder

## Requirements

- Windows 10/11
- PowerShell 5.1+
- **One of the following extractors:**
  - [7-Zip](https://www.7-zip.org/)
  - [WinRAR](https://www.win-rar.com/)

## Installation

### qBittorrent Integration

1. Open qBittorrent → **Tools** → **Options** → **Downloads**
2. Enable **"Run external program on torrent finished"**
3. Set the command:
   ```
   powershell.exe -ExecutionPolicy Bypass -File "D:\path\to\qBitLauncher.ps1" "%F"
   ```
   Replace `D:\path\to\` with the actual path to the script.

## Usage

When a torrent completes, the script will:

1. **For archives**: Prompt to extract, then show executable selection
2. **For folders with executables**: Show executable selection directly
3. **For media files**: Open the containing folder

### GUI Actions

| Button | Action |
|--------|--------|
| **Run** | Launch with administrator privileges (UAC prompt) |
| **Create Shortcut** | Create desktop shortcut |
| **Open Folder** | Open containing folder in Explorer |

## Configuration

### Themes

Edit line 27 to change the theme:
```powershell
$Global:ThemeSelection = 'qBitDark'  # Options: 'qBitDark', 'Dark', 'Light'
```

### Supported Extensions

| Type | Extensions |
|------|------------|
| Archives | `iso`, `zip`, `rar`, `7z`, `img` |
| Media | `mp4`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `webm`, `mp3`, `flac`, `wav`, `aac`, `ogg`, `m4a` |

## Logging

Logs are written to:
```
%TEMP%\qBitLauncher_log.txt
```

Fallback location (if temp fails):
```
C:\Users\Public\Documents\qBitLauncher_fallback_log.txt
```

## License

MIT License - see [LICENSE](LICENSE) for details.
