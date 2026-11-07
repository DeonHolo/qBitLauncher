# Upgrade status

The **3 November 2026 local 3.4.0 implementation** addresses the audit backlog. The user's clarified requirements govern two changes: naming is an alias only, and deliberate deletion of extracted installers/games remains available.

| ID | Status | Implemented behavior |
| --- | --- | --- |
| U01 | Implemented | Source-only, extracted-only, and both cleanup scopes; target preview; retained content; canonical containment; junction handling; Recycle Bin default; explicit app-deletion acknowledgement and permanent-delete confirmation |
| U02 | Implemented | API repair preserves credentials/authentication/HTTPS, binds loopback, and keeps an INI backup; saved app passwords use user-scoped DPAPI with migration |
| U03 | Implemented with clarified scope | F2/double-click/Shortcut Name changes only a saved future shortcut label for EXE/BAT/CMD. Original files never rename. Desktop link conflicts offer replace, keep both, or cancel. |
| U04 | Implemented | Same-release version/asset checks, published SHA-256 verification, Unicode handoff, unique helper/payload, original arguments, backup, and rollback |
| U05 | Implemented | Shallow payload selection, ambiguous archive choice, first multipart volumes, runnable-first handling for nested asset ZIPs |
| U06 | Implemented | Search over aliases and original paths, runnable ranking/type, helpers hidden by default with an explicit toggle |
| U07 | Implemented | Cancellable background inventory scans, queued/cached icon loading, background automatic/manual metadata checks, responsive initial API operations and cleanup progress |
| U08 | Implemented | Append/rotate logs, writer coordination, credential redaction, Open Log, and Copy Error |
| U09 | Implemented | Reliable compiled/source execution-path detection, preserved reinstall settings, source-mode shell shortcuts, terminating copy/build errors, staged atomic EXE replacement without killing a running app |
| U10 | Implemented | Recent folders and forgetting missing entries; preferred executable, alias, arguments, and admin mode per path; extraction base and window state remembered |
| U11 | Implemented | Resizable layout containers for the main/dashboard/settings/cleanup forms, DPI mode/manifest, full-path tooltips, remembered size and columns |
| U12 | Implemented | Windows checks/build workflow, pre-release gate, pinned PS2EXE, source version in binary metadata, executable/source/checksum assets |
| U13 | Revised after feedback | Independent concurrent game workflows, duplicate-source suppression across entry points, automatic lock release, and settings merges across windows; queue removed |
| U14 | Implemented | Separate service/workflow/update/UI source files with a build step preserving a self-contained distributed script/EXE |

## Practical limits and follow-ups

- Windows CI validates source changes; the release workflow publishes application builds after a versioned feature or fix push to `main`.
- Filename ranking is a heuristic. The app does not reliably distinguish a playable game from a disposable installer, and cleanup asks the user to choose.
- Recycle Bin support depends on Windows, the volume, and destination. Permanent local deletion requires its own checkbox and confirmation; removing only the qBittorrent entry keeps files.
- Extraction is owned by the installed 7-Zip/WinRAR GUI. Large real archives and a real torrent-completion/cleanup event were not run against personal downloads during verification.
- UI fixtures check resizing and DPI configuration. Physical testing across different display scales remains useful.
- Aliases and launch preferences follow original paths; moving a program requires selecting it again. Existing desktop shortcuts are not renamed when the saved future label changes.
- The cached icon directory is keyed by file metadata; a future maintenance action could prune obsolete cache entries if it becomes large.
- A release hash establishes consistency with the published asset; Authenticode signing remains a separate distribution improvement.

See the [validation report](2026-11-03-upgrade-implementation.md) and [usage guide](../usage.md).
