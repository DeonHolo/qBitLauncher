# qBitLauncher

<p align="center">
  <img src="https://i.imgur.com/0epFbuH.png" alt="qBitLauncher Logo" width="128">
</p>

A PowerShell post-download handler for qBittorrent that automatically extracts archives and launches executables with a clean GUI.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

- 🗜️ **Auto-extraction** - Extracts ZIP, RAR, 7z, ISO, and IMG archives
- 📂 **Custom extraction path** - Choose where to extract files with folder browser
- 📊 **Progress bar** - Visual feedback during extraction with percentage for 7-Zip
- 🔍 **Smart executable discovery** - Finds all .exe files with icons and sorts by folder depth
- 🎨 **Themed GUI** - Dracula dark theme with Light mode option (and much more to come!)
- 🛡️ **Run as Administrator** - One-click UAC elevation for installers
- 🔔 **Toast notifications** - Windows notifications for all actions
- 📁 **Multiple actions** - Run, create desktop shortcut, or open folder


## Requirements

- Windows 10/11
- PowerShell 5.1+
- **One of the following extractors:**
  - [7-Zip](https://www.7-zip.org/) (recommended - shows extraction progress in GUI)
  - [WinRAR](https://www.win-rar.com/)

## Installation

### Getting the Script

**Option 1: Clone the repository**
```bash
git clone https://github.com/DeonHolo/qBitLauncher.git
```

**Option 2: Download directly**
- Download `qBitLauncher.ps1` from the [repo](https://github.com/DeonHolo/qBitLauncher)

### qBittorrent Integration

1. Open qBittorrent → **Tools** → **Options** → **Downloads**
2. Enable **"Run external program on torrent finished"**
3. Set the command:
   ```
   powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\qBitLauncher.ps1" "%F"
   ```
   Replace `C:\path\to\` with the actual path to the script.

## Usage

When a torrent completes, the script will:

1. **For archives**: Prompt to extract (with custom path option), then show executable selection
2. **For folders with executables**: Show executable selection directly with icons
3. **For media files**: Open the containing folder

### GUI Actions

| Button | Action |
|--------|--------|
| **Run** | Launch with administrator privileges (UAC prompt) |
| **Shortcut** | Create desktop shortcut |
| **Open Folder** | Open containing folder in Explorer |
| **Settings** | Configure theme preferences |
| **Close** | Close the window |

All actions play a sound effect and show a confirmation popup.

## Configuration

Settings are accessible via the **Settings** button in the main window, or edit `config.json`:

```json
{
  "Theme": "Dracula"
}
```

### Themes

Available themes: `Dracula`, `Light`

### Supported Extensions

| Type | Extensions |
|------|------------|
| Archives | `iso`, `zip`, `rar`, `7z`, `img` |
| Media | `mp4`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `webm`, `mp3`, `flac`, `wav`, `aac`, `ogg`, `m4a` |

## Logging

Logs are written to `qBitLauncher_log.txt` in the script folder.

## License

MIT License - see [LICENSE](LICENSE) for details.
