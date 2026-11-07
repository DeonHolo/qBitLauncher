# Repository status: 3 November 2026

**Historical audit snapshot before the 3.4.0 upgrades.** The subsequent [implementation report](2026-11-03-upgrade-implementation.md) and [upgrade status](Upgrade-Backlog.md) describe current local behavior. Findings and local version/build statements below refer to the original audit.

## Verified GitHub connection

| Check | Result |
| --- | --- |
| Repository | [DeonHolo/qBitLauncher](https://github.com/DeonHolo/qBitLauncher), public and not archived |
| Local branch | `main`, tracking `origin/main` |
| Fetch and push remote | `https://github.com/DeonHolo/qBitLauncher.git` |
| Local HEAD before this audit | `2a88ed7da7f8ad62cfb31a43798fb71b85061163` |
| Live remote `main` | Same commit, verified with `git ls-remote` |
| Connected account | `DeonHolo`; repository metadata confirms admin and push permissions |
| Latest release | [v3.3.7](https://github.com/DeonHolo/qBitLauncher/releases/tag/v3.3.7), published 1 July 2026 at 02:41 Singapore time |
| Latest Actions run | [Successful run for the current main commit](https://github.com/DeonHolo/qBitLauncher/actions/runs/28469599179) |

Git's remote configuration, public read access, and the connected GitHub account's permissions were checked. A terminal Git push was not performed, so its separate credential flow was not exercised.

The original local EXE's SHA-256 matched the published v3.3.7 asset: `86ea660c15f29878dbeac8eab4906227301c6acd5e2b528d78584570d099d8d8`.

## Installed integration

The local qBittorrent `[AutoRun]` section is enabled and supplies the executable with `%F`, `%I`, and `%N`. Sampled per-user Explorer commands for directories, directory backgrounds, and ZIP files point to the same executable. These settings were inspected without editing qBittorrent or the registry. An actual torrent-completion event was not triggered.

## Problems confirmed during review

- **Unicode shortcuts:** the original WScript implementation threw `Value does not fall within the expected range` for Chinese, Japanese, Korean, Cyrillic, and emoji shortcut filenames on this PC. English, accented Latin, and bracketed names passed the same probe.
- **Ignored documentation:** `.gitignore` contained `docs/`, which also excluded new files under `Docs/` on this checkout. `usage.md` and `development.md` existed locally but were not tracked.
- **Stale development guide:** it described two workflow files and numeric rollover behavior that the actual single workflow does not implement.
- **Cleanup containment:** the source deletion guard compares source and extraction-root equality, without protecting a source directory that contains the extraction root.
- **API repair settings:** the repair function writes a fixed username/password-hash pair and disables localhost authentication, without explicitly binding the listener to loopback.
- **Rename extension:** the inline rename handler appends `.exe` even for `.bat` and `.cmd` files.
- **Responsiveness:** update checking, directory traversal, icon loading, and API requests run synchronously. The logger reads and rewrites the entire existing file for every entry; the local log was approximately 1 MB.

The [upgrade backlog](Upgrade-Backlog.md) gives concrete remedies. Findings outside the shortcut/docs work remain open.

## Local changes and validation

`New-WindowsShortcut` now uses `IShellLinkW` and `IPersistFile.Save` directly. Both the runnable-list shortcut button and installer shortcuts use this helper. The display name, target, and working directory remain Unicode; executables do not need to be renamed.

The regression script passed **37 checks in each of three environments**: Windows PowerShell 5.1, PowerShell 7.6.5, and a PS2EXE-compiled test harness running Windows PowerShell 5.1. It checks Unicode filenames, target folders, shortcut destination folders, combined paths, accents, emoji, brackets, spaces, ampersands, percent signs, `.bat`/`.cmd` targets, overwrite behavior, and missing-input errors. Explorer's automation API independently reads the saved links back. Test targets are never launched, and no shortcuts are written to the user's desktop.

The full app also compiled successfully. The local `qBitLauncher.exe` was rebuilt, with the original preserved as `qBitLauncher.exe.pre-unicode-20261103.bak`. The interactive full-app GUI, extraction, updates, installer, and destructive cleanup were not exercised by these tests.

The source version remains `3.3.7`; this is a local patch, not a new published release. No commits, pushes, releases, or hosted GitHub Wiki changes were made. The pre-existing `.token-optimizer/` directory was left in place.
