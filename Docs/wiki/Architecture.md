# Architecture

## Source and distribution

`Build.ps1` bundles the definition files into `Source/App.ps1` at its marker and writes the standalone root `qBitLauncher.ps1` in UTF-8 with a BOM. PS2EXE compiles that generated script. Distributed users need neither the development source directory nor additional PowerShell modules.

| File | Responsibility |
| --- | --- |
| `Source/App.ps1` | Startup, themes, integration, installer, extraction dialogs, and main routing |
| `Source/Services.ps1` | Settings, Windows DPAPI secrets, remembered state, log rotation, and duplicate-source locks |
| `Source/Updates.ps1` | Release metadata, background checks, checksum verification, replacement and rollback |
| `Source/Workflow.ps1` | Cleanup planning/execution, path containment, discovery/ranking, labels, Unicode links, launch handoff, and API repair helpers |
| `Source/UI.ps1` | Dashboard, selection/settings/cleanup forms, background tasks, and UI timers |
| `qBitLauncher.ps1` | Generated standalone source; do not maintain independent edits here |
| `Build.ps1`, `Compile.ps1` | Validation, standalone bundling, staged compilation, backups, and checksums |
| `Tests/` | Isolated behavioral checks and compiled-app startup verification |

## Workflow

Input can come from qBittorrent's `%F`, `%I`, `%N` completion arguments, Explorer context menus, or the dashboard's recent/manual folder selection. Different games run independently without a queue or shared processing slot. Requests for the same canonical source path share a nonblocking exclusive file handle. Windows releases the handle when its owner exits, including after a crash.

Folder traversal runs in a worker, skips reparse points, and collects runnables and payload archives. Shallow archives are preferred; multipart continuation volumes and common nested asset archives are excluded from payload selection. A runnable folder without a root payload archive can go directly to selection. Ambiguous payloads get a choice dialog.

Extraction still uses the installed tool's GUI and progress/cancellation. RAR prefers WinRAR; other formats prefer 7-Zip. After extraction, a cancellable scan finds runnables. The selection form ranks likely applications/installers ahead of maintenance tools, supports name/path search, and exposes helpers through Show helpers/uninstallers. Icons load from a worker into a UI-owned queue and are cached by file identity/mtime.

Aliases, arguments, administrator mode, recent folders, preferred executable, window size, and column widths are remembered. Labels never change the original file. EXEs launch through Windows shell execution; BAT/CMD targets use an encoded Unicode PowerShell handoff so literal percent characters in their paths survive cmd's parsing.

## Cleanup model

`Get-CleanupPlan` produces an explicit list of approved targets and retained locations. Source-only splits a source tree around retained extracted/selected content. Extracted-only and Everything deliberately permit removal of the selected app, with acknowledgement. Overlapping roots are pruned so a parent/child target is not removed twice.

The UI offers independent source, extracted, and qBittorrent-entry checkboxes. All three can run together. Entry removal uses the API with `deleteFiles=false` before local cleanup to release seeding handles; the local plan owns file deletion, preventing double deletion. An API failure stops before local removal. Completed steps remain in the activity log if a later step fails.

Canonical paths account for junction targets. Traversal does not descend through links; cleanup retains reparse points, rejects protected roots, and revalidates targets after the preview. Local removal defaults to Windows' Recycle Bin. Permanent local deletion needs a separate choice and confirmation.

These checks cannot identify the installed location of a game outside the displayed folders. Cleanup does not infer whether an installer is still needed.

## Storage

`config.json` remains beside the application for portable compatibility. Under a short settings-write mutex, each window merges its changes since loading with the latest disk state, preserving another game's preferences and settings. Saved API passwords are encrypted with user-scoped Windows DPAPI; plaintext settings migrate on load. A config moved to another Windows user requires password re-entry. A decryption failure preserves the existing ciphertext instead of silently erasing it.

Logs append under `%LOCALAPPDATA%\qBitLauncher\logs` with 2 MB rotation, three archives, interprocess coordination, and credential redaction. Per-source lock files, icon cache, and update handoff files also live under the per-user application directory. `QBITLAUNCHER_DATA_DIRECTORY` redirects that runtime directory for isolated diagnostics. `QBITLAUNCHER_SKIP_UPDATE_CHECK=1` suppresses update networking in verification runs.

qBittorrent integration changes its INI file with a backup; Explorer entries are per-user registry keys. API repair enables the service and binds loopback while preserving existing username, password, authentication, and HTTPS settings.

## Updates and CI

Automatic release checks use a background runspace and a 12-hour cache after a successful check. Version and downloadable asset come from the same published release. SHA-256 verification is mandatory before replacement. A unique Unicode helper waits for the old process, replaces the file on its volume, keeps a backup, and can roll back a failed restart.

Windows CI parses/bundles, runs regression checks, compiles with a pinned PS2EXE version, and creates executable/source/checksum artifacts. The release workflow validates before creating a version commit/tag and publishing. See [development](../development.md) for the exact commands and triggers.
