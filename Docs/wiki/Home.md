# qBitLauncher wiki

qBitLauncher opens downloads and folders, hands archives to 7-Zip or WinRAR, and shows the runnable files it finds. These pages cover extraction, desktop shortcuts, saved preferences, and cleanup.

| Page | Contents |
| --- | --- |
| [Usage](../usage.md) | Opening downloads, naming shortcuts, launch settings, and cleanup choices |
| [Architecture](Architecture.md) | Source organization, workflow, stored data, and packaging |
| [Interface style](Interface-Style.md) | Compact Dracula direction, colors, hierarchy, and control behavior |
| [Upgrade status](Upgrade-Backlog.md) | What was implemented from the audit and remaining limits |
| [Implementation and validation](2026-11-03-upgrade-implementation.md) | Changes, tests, local build, and publishing status |
| [Scanning and extraction handoff](2026-11-10-scanning-fix.md) | Large-folder scan changes, extraction output checks, and validation limits |
| [Troubleshooting](Troubleshooting.md) | Missing files, cleanup scope, shortcuts, credentials, and updates |
| [Development](../development.md) | Build commands, checks, CI, and releases |
| [Original repository audit](Repository-Status.md) | Historical snapshot before the upgrades |

qBitLauncher is a Windows PowerShell/Windows Forms utility packaged as a standalone EXE. It accepts completed content from qBittorrent or a local folder, hands archives to 7-Zip or WinRAR, and shows the runnable files it finds.

## Two important choices

**Naming changes the future shortcut label only.** F2, double-click, and Shortcut Name save a name for a desktop `.lnk`. The original EXE, BAT, or CMD keeps its filename and contents.

**Cleanup follows the user's intent.** Extracted files may be a disposable installer or the actual playable game. Tick source files, extracted files, the qBittorrent entry, or all three. Removing the source alone keeps extracted content. Local files go to the Recycle Bin by default; permanent deletion requires a separate choice and confirmation.

## Repository and releases

Source, documentation, and screenshots live in [DeonHolo/qBitLauncher](https://github.com/DeonHolo/qBitLauncher). Windows CI validates changes on `main`; the release workflow builds the EXE, standalone script, and checksums. Download application builds from [Releases](https://github.com/DeonHolo/qBitLauncher/releases/latest). This wiki is stored in the repository under `Docs/wiki/`.

Update these Markdown pages alongside code changes. Keep passwords, personal runtime logs, and generated binaries out of source control. The root script is generated from `Source/`; edit those development files and run `Build.ps1`.
