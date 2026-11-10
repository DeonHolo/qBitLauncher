# Troubleshooting

These notes describe local **3.4.0**. The [original audit](Repository-Status.md) records older behavior.

## Naming a game without renaming its EXE

Select the file and use F2, double-click, right-click → Edit shortcut name, or Shortcut Name. This saves a future desktop label only. The Original file column and file on disk stay unchanged. Click Shortcut to create the desktop link. Changing the label later does not rename an existing `.lnk`.

If a desktop name is taken, choose replace, keep both, or cancel. Non-English labels and paths use Windows' Unicode Shell Link interface; no system locale change or executable rename is needed. The admin checkbox controls Run inside qBitLauncher, rather than an always-elevated flag on desktop links.

## What should cleanup remove?

The app cannot know whether an extracted folder contains the game itself or an installer you already used. Review the exact paths:

- **Keep the extracted game:** choose source-only; extracted files and the selected app stay.
- **Remove an installer after installing elsewhere:** choose extracted-only and acknowledge removing that displayed tree.
- **Remove the whole download and extracted content:** choose both and acknowledge the preview.
- **Stop tracking the torrent but keep everything:** choose torrent-only.

If extraction and source are the same folder, source-only may have nothing separate to remove. Choose a broader scope deliberately if that entire folder is unwanted. Local cleanup defaults to the Recycle Bin; permanent deletion needs an explicit extra choice. qBittorrent's torrent-data deletion is a separate permanent action.

Junctions are retained instead of traversed. If a folder changes after the preview, recreate the preview. Protected system/application roots cannot be cleanup targets.

## The desired executable is missing

Enable Show helpers/uninstallers and clear Search. Ranking can mistake uncommon names for support tools. Files in inaccessible directories or behind junctions may be skipped. Open Folder to inspect them. If a remembered file moved, choose its new folder/path and save its preferences there.

## An archive choice appears

Root/shallow payload archives take precedence over deep asset archives. If multiple payloads are plausible, choose the intended archive or Show executables. In the extraction dialog, Skip Extraction scans the original folder. Multipart sets start from the first supported volume; keep the remaining volumes beside it for the external extractor.

## Open several games at once

The current local version has no queue. Separate games extract independently, even while another result window stays open. An **Already open** message applies only to the same source path. If an old window still says **Download queued**, that window belongs to the earlier executable; reopen that source with the rebuilt local application.

## A scan appears stuck after extraction

The updated scan displays checked files, found runnables, the current folder, and elapsed time. Use **Use files found** to stop the remaining scan and open available results, or **Cancel** to leave it. Scan completion, cancellation, counts, skipped locations, and the destination are written to the log.

Check the destination shown by qBitLauncher: it may differ from extracting beside the source archive in Explorer. If output validation reports missing or short files, keep the original archive and retry extraction into that destination. The extractor's exit code alone does not establish complete output. When archive metadata cannot be listed, the log explicitly says that file sizes were not verified.

## Launch arguments or administrator mode

Arguments and admin mode are remembered per file. New files initially request elevation; Settings changes that default. Cancelling UAC offers a non-admin retry. BAT/CMD launches use a Unicode helper and preserve literal percent characters in target paths. Programs may impose their own filename, permission, or argument requirements.

## qBittorrent API cannot connect or authenticate

Check Settings for the URL and existing qBittorrent account. Repair shows its changes, creates a backup, enables the API and loopback binding, and preserves credentials/authentication/HTTPS. If HTTPS is enabled in qBittorrent, the URL must match it. A 401/403 response requires the correct existing credentials; repair does not reset them.

Passwords in config are encrypted for the Windows user. After moving a portable config to another account, re-enter the password. Failed decryption retains the previous ciphertext until replacement.

## Updates or builds fail

Automatic checks run in the background and are cached after success; Settings can force a check. Only a published eligible release with the requested asset and verifiable SHA-256 is installed. A checksum failure keeps the existing app. Update handoff records and backups can help diagnose replacement/restart errors.

Build with Windows PowerShell 5.1 using `Build.ps1`. Close the running app before replacing its EXE; the build does not terminate it. Edit files under `Source/`, then regenerate the root script. Open Log and Copy Error provide diagnostics with credential redaction.
