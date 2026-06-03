# qBitLauncher Extraction and Admin Launch Refactor

Date: 2026-06-03

## What Started This

The launcher had useful core workflow, but three UX problems were getting in the way:

- The Run button did not actually launch with administrator privileges, so installer-style executables such as FitGirl setup files failed unless opened manually from Explorer.
- The extraction UI was too slow, especially for very large ZIP files. qBitLauncher tried PowerShell native ZIP extraction first and also spent time maintaining its own progress UI.
- The extraction window's click-away/minimize behavior was finicky. This was mostly noticeable because the custom extraction UI stayed open for a long time.

## UX Decisions

### Extraction

The user should not have to choose between 7-Zip and WinRAR. That choice is internal plumbing.

The visible UX should stay simple:

1. qBitLauncher detects an archive.
2. The user confirms the destination path.
3. The user clicks one Extract action.
4. qBitLauncher launches the real installed extractor GUI.
5. When extraction finishes, qBitLauncher scans the extracted folder.
6. qBitLauncher shows the executable list.

Extractor priority:

- RAR files use WinRAR first, then 7-Zip GUI as fallback.
- ZIP, 7z, ISO, and IMG files use 7-Zip GUI first, then WinRAR as fallback.
- If neither extractor exists, qBitLauncher tells the user to install 7-Zip or WinRAR.

The old qBitLauncher extraction progress window should not be used. The real extractor window is faster, more familiar, and already supports progress and cancellation.

### Run Button

The UI should have only one Run button.

That button should always launch the selected executable with UAC elevation. Users do not need separate Run and Run as Admin choices because the main use case includes installers and setup files that often need admin rights.

### Window Behavior

The old auto-minimize issue was tied to the custom extraction progress form. Since that form is removed, the behavior should be revisited only if it still feels annoying after real extractor handoff is in place.

## Code Changes Made

- Added GUI extractor discovery for 7-Zip and WinRAR.
- Added automatic extractor selection by archive type.
- Replaced qBitLauncher's custom extraction implementation with external GUI extractor handoff.
- Removed PowerShell native `Expand-Archive` from the extraction path.
- Removed the old qBitLauncher extraction progress form and its update helper.
- Changed the single Run button to launch with `-Verb RunAs`.
- Updated README feature wording to describe external extraction and admin launch.

## Important Follow-Ups

- Rebuild `qBitLauncher.exe` from the updated script if using the compiled app.
- Test a large ZIP with 7-Zip GUI handoff.
- Test a RAR archive with WinRAR handoff.
- Test a FitGirl or similar setup executable from the EXE list and confirm UAC appears.
- If executable discovery remains noisy, add ranking/filtering later so setup/game executables appear above crash handlers, uninstallers, redists, and helper binaries.

## Rollback Notes

The work was done on branch `codex-external-extractor-admin-run`.

The repository has a GitHub remote:

```text
https://github.com/DeonHolo/qBitLauncher.git
```

If this refactor behaves badly, switch back to `main` or restore the previous commit from Git.
