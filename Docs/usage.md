# Usage

[Wiki home](wiki/Home.md) · [Troubleshooting](wiki/Troubleshooting.md) · [Implementation report](wiki/2026-11-03-upgrade-implementation.md)

These instructions describe the current source. Download the matching application from [Releases](https://github.com/DeonHolo/qBitLauncher/releases/latest); older releases may have different controls.

## Open a download or folder

Launch `qBitLauncher.exe` without arguments to open setup or the dashboard. Use **Choose a folder...**, select a recent folder and **Open recent folder**, or start through qBittorrent/Explorer integration. Missing recent folders can be forgotten or selected again at their new location.

The dashboard can install/update qBittorrent integration and add Shift+right-click context menus. Use **Remove integrations...** to remove the qBittorrent completion command or Explorer menu entries. qBittorrent must be completely closed before its INI integration is changed.

qBittorrent supplies the content path, torrent hash, and torrent name. A torrent hash enables torrent removal; local file actions also work for manual and Explorer folder workflows.

Different games open and extract independently. A large extraction or an open result window does not hold up a smaller game. Opening the same source twice reports **Already open** to avoid two extractions writing to the same location. There is no download queue.

## Choose an archive and extraction destination

The folder scan prefers useful shallow archive payloads and avoids common nested asset archives. When several payloads are plausible, choose one or choose executable scanning. An already runnable folder can go straight to executable selection instead of prompting about a nested `Plugins\icons.zip`.

For extraction, select the destination, reuse existing files, or choose **Skip Extraction** to scan the original folder. The destination base is remembered. RAR prefers WinRAR; the other supported formats prefer 7-Zip GUI, with the other installed tool available as fallback. The external extractor owns extraction progress and cancellation.

## Find and launch a runnable

Search matches the saved shortcut label and original full path. Results include separate columns for **Shortcut name**, **Original file**, **Type**, and **Folder**. The type/rank is a filename/path heuristic, not proof that a file is a game or installer. Enable **Show helpers / uninstallers** when a maintenance tool is the intended target.

Select a file and click **Run**. **More → Launch options** reveals saved launch arguments and the administrator checkbox. Arguments and admin mode are saved per original file path. Administrator launch is initially enabled, matching the earlier application behavior; **More → Settings** changes the default for files without a saved override. The primary button says **Run as admin** when that file is set to elevate.

The preferred executable is remembered per folder. The result window can be resized and remembers window size and column widths. Hover a file or folder for the complete path.

## Name the future desktop shortcut

**F2, double-click, right-click → Edit shortcut name, or Shortcut Name changes only a saved label. It never renames the EXE, BAT, or CMD file.**

The name edits in its existing cell. **Enter** saves, **Escape** cancels, and clicking outside saves. The editor uses the list colors and hides the old label while typing.

For example, give `setup.exe` the label `Example Game Installer`. Its on-disk path remains `setup.exe`; clicking **Shortcut** later creates `Example Game Installer.lnk`. The stored label remains available after reopening the same original file path. Changing the label does not rename an existing desktop shortcut either.

**Shortcut** uses the selected target, its containing folder as working directory, and the current launch arguments. Unicode names and destination folders are supported. If the desktop label already exists, choose **Yes** to replace it, **No** to keep both with a numbered name, or **Cancel** to create nothing.

The admin checkbox controls **Run** in qBitLauncher; it does not mark the resulting desktop link to always elevate.

## Cleanup: choose what to keep

Open **Cleanup** and tick any combination of the three items. Each file location appears once; hover a truncated path to read it in full. Review the summary, then click **Remove selected**:

| Choice | Effect |
| --- | --- |
| **Source files** | Removes the source files. Unticked extracted/selected content is kept, including when it is nested inside the source. |
| **Extracted files** | Removes the extracted tree, including its game or installer. Requires acknowledgement when it includes extracted/selected content. |
| **qBittorrent entry** | Stops seeding and removes the torrent entry. File deletion is controlled by the other two checkboxes. Disabled for manual folders without a torrent hash. |

To remove **source + extracted + qBittorrent entry**, tick all three and acknowledge removing the extracted app/installer. The entry is removed first to release seeding handles; the local file plan then removes each location once, including overlapping paths.

The program does not infer whether extracted files are disposable installers or the actual game. Typical choices are:

- **Portable/playable game in the extracted folder:** source-only cleanup keeps it.
- **Installer already used, with the installed game elsewhere:** extracted-only cleanup can remove that installer tree; choose both if the original source is also unwanted.
- **Everything in these displayed source/extraction paths is unwanted:** choose both and acknowledge the preview.

Local cleanup uses the Recycle Bin by default. **Permanently delete files (skip Recycle Bin)** requires another confirmation. Recycle Bin behavior depends on the Windows destination/volume. Removing only the qBittorrent entry leaves all files on disk. If entry removal fails, local cleanup does not start; if file removal fails afterward, the activity log records that the entry was already removed.

Path checks reject protected system/app roots, preserve junctions rather than following them, and recheck targets before removal. If files or links changed after the preview, prepare a fresh preview. If source-only has nothing separable to remove, choose another scope deliberately rather than assuming the extracted folder is disposable.

## Settings, errors, and keyboard controls

**More → Settings** includes theme, local API URL/account, extraction base, default admin mode, and background update checks. API passwords are encrypted for the current Windows user; use Settings to re-enter credentials after moving portable config to another account.

**More → Open Log** opens the current per-user log. **More → Copy Error** copies the last captured error with credential redaction. The activity panel shows actions for the current window; rotated log history is under `%LOCALAPPDATA%\qBitLauncher\logs`.

| Action | Control |
| --- | --- |
| Run | `Alt+R` |
| Create desktop shortcut | `Alt+S` |
| Edit future shortcut name | `F2`, double-click, right-click action, or `Alt+N` |
| Open folder | `Alt+O` |
| Settings | `Alt+T` |
| Close | `Alt+C` |

Automatic update checks run in the background. When a version is available, Settings offers installation; **Check for updates** forces a check without waiting for the normal cache interval.
