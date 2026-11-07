# Local 3.4.0 upgrade — 3 November 2026

The audit upgrades were implemented locally with two explicit product decisions from the user:

1. **Names are future shortcut labels.** Editing a label must never rename an EXE, BAT, or CMD.
2. **Deleting extracted content remains available.** An extracted folder can be an installer or a playable app; the user chooses what to keep. Cleanup offers separate scopes, exact previews, and acknowledgement for removing the extracted/selected content.

The [upgrade status](Upgrade-Backlog.md) maps all fourteen audit items to implementation. The [usage guide](../usage.md) describes the resulting controls.

## Validation

All checks use temporary fixtures and isolated configuration rather than personal downloads. The complete Windows PowerShell build gate passes the source suites:

| Suite | Checks | Evidence |
| --- | --- | --- |
| Build | 19 | Bundling, source validation, atomic staging, hash/version behavior |
| Launch | 43 | Actual fixture BAT/CMD execution with Unicode, spaces, brackets, ampersands, apostrophes, bangs, and literal percent-variable names; native EXE arguments and admin preference |
| Services | 54 | DPAPI migration/encryption, saved state, redacted log rotation/concurrent writers, independent concurrent workflows, duplicate-source locks, and settings merges across windows |
| Cleanup selection | 34 | All seven checklist combinations, explicit API keep-files flag, removal order, acknowledgements, permanent choice, and partial failures using stubbed side effects |
| Shortcuts | 38 | Explorer independently reads Unicode link/target/folder paths and saved arguments; missing targets rejected |
| UI | 85 | Controlled forms exercise filtering, helper visibility, shortcut aliases, original file preservation, preference restoration, combined cleanup checklist/acknowledgement, settings, dashboard, compact More options, collapsed-column recovery, and real inline editing with Enter/Escape/focus loss |
| Updates | 62 | Published metadata/checksum fixtures, real Unicode helper replacement, positional argument quoting, corruption rejection, rollback, background disposal |
| Workflow | 94 | Source-only preservation, acknowledged deliberate deletion, containment/case/sibling paths, stale previews/junction substitution, root protection, ranking/archives, and repair credentials |

Total: **429 focused assertions** in the integrated Windows PowerShell 5.1 run after the 7 November revisions. Services, updates, and workflow suites also passed on PowerShell 7 during the initial implementation. The full application is compiled from the generated source with version metadata and a DPI manifest. `Tests/Verify-AppStartup.ps1 -Parallel` separately checks the actual compiled application from an isolated Unicode installation path, with two independent game windows open simultaneously.

## Visual refinement after feedback

The user requested the previous compact feel and fewer bright native bars. The main form uses themed headers/selection, dark title-bar colors, soft field surfaces, fitting columns, and a quiet activity panel. Launch options, Settings, and log utilities live under More. Original file identity remains visible and F2 still edits only a future shortcut label.

After the user corrected the palette, every original theme token was restored from the pre-upgrade script: background `#282A36`, controls `#202020`, buttons `#2F3440`, border `#44475A`, accent `#6272A4`, text `#F8F8F2`, and secondary text `#969696`. The added violet colors and filled primary button were removed, including custom-control defaults and hover states. Real form previews were checked; physical multi-display DPI checks remain separate.

On 6 November, the remaining native white header gap was fixed by painting the complete header. Column and panel dividers are visible, and collapsed saved columns recover usable widths. Cleanup now has independent source/extracted/qBittorrent-entry checkboxes, including all three together. Each location appears once, and the summary explains the effects. The workflow queue was removed: different games open and extract independently, while per-source locks still prevent duplicate extraction. Settings saves merge changes across windows.

On 7 November, inline shortcut editing was reproduced with a native white field starting before the drawn label while the original name remained visible beneath it. A themed editor now occupies the existing name cell after the icon, suppresses the old label while typing, and stays within the column when resized. Enter, Escape, and focus loss are exercised against the actual editor. Original files remain unchanged.

## Local files and publishing

`qBitLauncher.exe` is the local runtime used by the existing integration paths. Build replacement preserves the previous executable in a unique `.pre-build-...bak`. `qBitLauncher.exe.pre-unicode-20261103.bak` preserves the original published binary; `qBitLauncher.ps1.pre-upgrades-20261103.bak` preserves the earlier local script.

Development files are separated under `Source/`; `Build.ps1` bundles the standalone root script, runs checks, compiles, and writes `SHA256SUMS.txt`. Windows CI and the automatic release workflow use this same build. The 7 November publishing update adds current application and cleanup screenshots and a rewritten README.

Verification used isolated files and settings. It did not delete personal downloads or change actual qBittorrent cleanup/configuration settings. Existing integration settings were retained. Live release installation, a real torrent completion/cleanup, large real extraction, and physical multi-monitor DPI testing remain outside the fixture checks. Source publication and release builds are handled separately by GitHub workflows; this report records the local implementation and its test evidence.

The [repository audit](Repository-Status.md) is historical: its pre-upgrade problems and v3.3.7 local version statements describe the initial snapshot, not current 3.4.0 behavior.
