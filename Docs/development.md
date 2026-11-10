# Development, testing, and releases

[Wiki home](wiki/Home.md) · [Architecture](wiki/Architecture.md) · [Implementation report](wiki/2026-11-03-upgrade-implementation.md)

## Development source and standalone distribution

Edit `Source/App.ps1` for startup, integration, installer, and main routing; the other source files hold definitions:

1. `Source/Services.ps1`: settings, DPAPI, logging, remembered state, and duplicate-source locks. Settings saves merge each window's changes with the latest disk state.
2. `Source/Updates.ps1`: release metadata, background checks, checksums, and replacement/restart handoff.
3. `Source/Workflow.ps1`: cleanup planning, path checks, the native directory-scanning task, archive/runnable discovery, labels, shortcuts, and repair helpers.
4. `Source/Extraction.ps1`: metadata-verification entry points for expected extracted paths and sizes.
5. `Source/UI.ps1`: Windows Forms and responsive worker orchestration. File scans poll a native task directly; other background work uses PowerShell runspaces.

`Build.ps1` inserts these definition files, in that order, at `# <BUNDLED_FUNCTIONS>` in the app template. The root `qBitLauncher.ps1` is generated output and is checked in for standalone source users. Do not maintain changes only in that generated file: the next build regenerates it.

The distributed script and EXE do not need the loose `Source/` files. The generated script is UTF-8 with a BOM so Windows PowerShell 5.1 reads Unicode source correctly. Build parsing reads source text as UTF-8 before validation, avoiding the ANSI interpretation of BOM-less development files.

## Build commands

Run from the project root using Windows PowerShell 5.1:

```powershell
# Generate the standalone script, run all Test-*.ps1 checks, then build the EXE.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build.ps1

# Generate and validate source without compiling an EXE.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build.ps1 -SkipCompile

# Stage release files in a separate directory.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build.ps1 -OutputPath C:\Temp\qBitLauncher-build\qBitLauncher.exe -SourceOutputPath C:\Temp\qBitLauncher-build\qBitLauncher.ps1
```

`-SkipTests` is available when the same sources were just validated. `Compile.ps1` is a noninteractive compatibility wrapper, accepting `-SkipTests` and `-OutputPath`.

The build validates syntax and duplicate top-level function names, rejects startup side effects in definition files, and runs every `Tests/Test-*.ps1` in a fresh Windows PowerShell process with `-SourcePath` pointing at the generated script. PS2EXE is pinned to **1.0.17** and installed for the current user when missing.

Compilation uses a plain temporary compiler directory to work around PS2EXE's wildcard handling in manifest paths. The result is copied to staging on the destination volume, validated, then atomically installed. A unique `.pre-build-...bak` preserves an existing EXE. A running app is not killed; close it before rebuilding to its path. The binary receives version metadata, STA mode, and a DPI manifest. `qBitLauncher.ps1` and `SHA256SUMS.txt` are placed alongside the output EXE.

## Focused checks

The scripts use assertions directly; Pester is not required.

| Script | What it exercises |
| --- | --- |
| `Tests/Test-Workflow.ps1` | Cleanup preservation/acknowledgement, containment and reparse handling, labels, ranking, archive selection, repair settings |
| `Tests/Test-Services.ps1` | Secret migration/encryption, saved state, log rotation/redaction, simultaneous workflows, duplicate-source protection, and settings saves from older windows |
| `Tests/Test-CleanupSelection.ps1` | All seven checklist combinations, entry removal before file cleanup, Recycle Bin/permanent choices, and partial failures with stubbed side effects |
| `Tests/Test-Updates.ps1` | Release/asset consistency, checksums, argument quoting, Unicode replacement, rollback and helper execution in fixtures |
| `Tests/Test-Shortcuts.ps1` | Explorer-readback verification of Unicode links, targets, arguments, destination folders, and missing inputs |
| `Tests/Test-Build.ps1` | Bundling, syntax/marker checks, UTF-8 output, version metadata constraints, atomic publication and hashes |
| `Tests/Test-UI.ps1` | Controlled WinForms interactions for selection, cleanup choices, settings, and setup |
| `Tests/Test-Launch.ps1` | Actual disposable BAT/CMD launches with special-character paths, saved arguments, and native EXE start-info validation |
| `Tests/Test-Scanning.ps1` | Streaming discovery across 12,005 files, a locked sparse 12 GB payload, completion, cancellation, junction cycles, missing roots, and found-file snapshots |
| `Tests/Test-Extraction.ps1` | Missing/short output, incorrect destinations, expected file sizes, archive-list timeouts, asynchronous verification, and cancellation during listing |

Run a focused check after generating the current standalone script:

```powershell
powershell.exe -NoProfile -File .\Tests\Test-Workflow.ps1 -SourcePath .\qBitLauncher.ps1
```

Tests import definitions rather than app startup, isolate state in temporary directories, and do not operate on the user's real downloads. UI checks use controlled forms and fixture content. Actual torrent completion, real release installation, and physical multi-display scaling are separate runtime checks; the [implementation report](wiki/2026-11-03-upgrade-implementation.md) records their limits.

After compilation, verify real app startup using `powershell.exe -NoProfile -File .\Tests\Verify-AppStartup.ps1 -ExecutablePath .\qBitLauncher.exe -Parallel`. It opens an isolated portable copy with updates disabled, checks that a Unicode runnable folder containing an unrelated nested archive reaches the selection UI, then opens a second game while the first result window stays open. Both close without launching targets or changing integrations.

Add `-ExtractArchive` to exercise an actual 7-Zip extraction of a disposable archive, then verify that expected output checks and scanning reach executable selection. This optional check needs an installed 7-Zip GUI and CLI. It never extracts a personal archive or launches fixture executables.

## CI and releases

`.github/workflows/ci.yml` validates and compiles on pull requests and pushes to `main`, then uploads the EXE, standalone source, and checksums as a build artifact.

`.github/workflows/release-pipeline.yml` runs on pushes to `main`. One Windows job validates before versioning and publication. It reads the version from `Source/App.ps1` and examines commits since the matching version tag; the push's commit messages are the fallback when that tag does not exist.

| Keyword, in precedence order | Version change |
| --- | --- |
| `BREAKING:` | Increment major; reset minor and patch |
| `feat:` | Increment minor; reset patch |
| `fix:` | Increment patch |
| None | Validation only; no release |

Matching uses literal keyword searches. There is no rollover at 9: `3.3.9` becomes `3.3.10` for a patch bump.

When a bump is needed, the workflow checks that `main` still matches the expected commit, updates the app template, regenerates the standalone source, and stages the build. It commits both versioned files, creates a tag, and atomically pushes the branch and tag. Release concurrency and the atomic push prevent an old build from overwriting newer `main`; bot pushes do not run the release loop again. Publication uses GitHub CLI to upload `qBitLauncher.exe`, `qBitLauncher.ps1`, and `SHA256SUMS.txt` for the verified tag.

Local building does not create a commit, push, tag, release, or hosted Wiki. Pushing to `main` runs Windows validation and the automatic release workflow described above. The project wiki is versioned under `Docs/wiki/` with the source.

## Update delivery

Both executable and script users check the latest eligible published release. The requested asset and version come from the same release. Downloads must match a published SHA-256 digest or matching `SHA256SUMS.txt` entry; an unverified asset is not installed automatically.

Checks run in the background, normally at most once per 12 hours after a successful check. Settings can disable automatic checks or force a manual check. Replacement uses a unique PowerShell helper and UTF-8 JSON payload, rather than an ASCII batch file. The previous application is retained as `.previous.bak`, and a failed replacement/restart can roll back. Helper status is written under the per-user `updates` directory.
