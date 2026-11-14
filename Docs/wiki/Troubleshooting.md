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

Root/shallow payload archives take precedence over deep asset archives. If multiple payloads are plausible, choose the intended archive or Show executables. In the extraction dialog, Open extracted files opens the selected destination; Scan source folder scans the original location. Multipart sets start from the first supported volume; keep the remaining volumes beside it for the external extractor.

## Open several games at once

The current local version has no queue. Separate games extract independently, even while another result window stays open. An **Already open** message applies only to the same source path. If an old window still says **Download queued**, that window belongs to the earlier executable; reopen that source with the rebuilt local application.

## A scan appears stuck after extraction

The updated scan displays checked files, found runnables, the current folder, and elapsed time. Use **Use files found** to stop the remaining scan and open available results, or **Cancel** to leave it. Scan completion, cancellation, counts, skipped locations, and the destination are written to the log.

Check the destination shown by qBitLauncher: it may differ from extracting beside the source archive in Explorer. If output validation reports missing or short files, keep the original archive and retry extraction into that destination. The extractor's exit code alone does not establish complete output. When archive metadata cannot be listed, the log explicitly says that file sizes were not verified.

## Return after closing a game window

Open qBitLauncher without a file and use **Recent sessions → Reopen session**. Saved sessions retain the source archive/download, game folder, and torrent identity. Closing the window leaves that record in `config.json`; **Forget session** removes the record without deleting files.

For an extraction made before session history existed, open its archive, choose the existing destination, and click **Open extracted files**. **Scan source folder** scans the archive's containing location and can cover all of `D:\Games` when the archive is stored there.

## Launch arguments or administrator mode

Arguments and admin mode are remembered per file. New files initially request elevation; Settings changes that default. Cancelling UAC offers a non-admin retry. BAT/CMD launches use a Unicode helper and preserve literal percent characters in target paths. Programs may impose their own filename, permission, or argument requirements.

## qBittorrent API cannot connect or authenticate

Open More → Settings and enter the API URL and the username and password set in qBittorrent's Web UI options. Use **Test connection** to check those values without changing torrents or files. If HTTPS is enabled in qBittorrent, the URL must use HTTPS too.

When cleanup encounters a rejected login, Settings opens with the error. Correct the credentials and choose **Save and retry cleanup**. The launcher verifies access before saving and continuing with your original cleanup choices. Cancel leaves local files untouched. Other open launcher windows reload saved credentials before their next cleanup attempt.

A 403 response during login can mean qBittorrent temporarily blocked access after too many failed attempts. Wait for the block to expire before testing again. The launcher preserves qBittorrent's password and access rules.

If qBittorrent shows **“WebUI configuration failed. Reason: Credentials are not set”**, open Tools → Options → Web UI. Under Authentication, set a username and use Change the current password to set a password, then click Apply. Enter the same username and password in qBitLauncher → More → Settings and retry cleanup.

Repair previews its changes, creates a backup, and enables the API on the loopback address while preserving credentials, authentication rules, and HTTPS. If it asks you to exit qBittorrent, use File → Exit or the tray menu so it can save the settings safely.

Passwords in config are encrypted for the Windows user. After moving a portable config to another account, re-enter the password. Failed decryption retains the previous ciphertext until replacement.

## Updates or builds fail

Automatic checks run in the background and are cached after success; Settings can force a check. Only a published eligible release with the requested asset and verifiable SHA-256 is installed. A checksum failure keeps the existing app. Update handoff records and backups can help diagnose replacement/restart errors.

Build with Windows PowerShell 5.1 using `Build.ps1`. Close the running app before replacing its EXE; the build does not terminate it. Edit files under `Source/`, then regenerate the root script. Open Log and Copy Error provide diagnostics with credential redaction.
