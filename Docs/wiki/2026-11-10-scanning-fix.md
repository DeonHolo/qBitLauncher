# Scanning and extraction handoff, 10 November 2026

The user reported a scan that stayed open after large extractions. The runtime log recorded an extractor exit followed by more than 17 minutes without a scan-completion record before cancellation. A second extraction also reached scanning. Fresh reads of those extracted folders later completed quickly, so file size alone did not reproduce the original stall. The previous log text treated an extractor exit code as successful extraction without checking its output.

The scan now uses a dedicated .NET task instead of a PowerShell background runspace. It streams directory metadata, keeps junctions out of traversal, and publishes completion and progress directly. The dialog shows checked files, folders, found runnables, the current path, and elapsed time. Cancellation returns without waiting for disk I/O. **Use files found** explicitly stops the remaining scan and opens a partial result.

After extraction, or when reusing existing output, the scan can list the archive through the installed 7-Zip CLI and compare expected filenames and sizes with the chosen destination. This runs in the same cancellable native task. Listing has a 15-second limit; missing or short output produces an incomplete-output error with the destination. If listing is unavailable, the log records that the output was not verified. These are metadata checks, not CRC checks or proof that an executable will work.

Validation used isolated settings and disposable files:

- The full build passed 467 assertions across ten suites.
- A streaming scan completed 12,005 fixture files, including a locked sparse 12 GB payload, in about 0.06 seconds on this machine. That measures metadata scanning, not extraction throughput.
- Checks cover junction cycles, missing roots, cancellation, partial results, UI completion and errors, short or missing extracted output, listing timeouts, and cancellation during verification.
- The compiled app extracted a disposable archive through 7-Zip, checked 501 expected files, reached executable selection, and kept a second game workflow independent. Ten compiled startup/extraction checks passed.
- Read-only checks of the reported extracted folders completed quickly with the new scanner. The ISO file list matched six expected file sizes in qBitLauncher's chosen output folder. The user's manual extraction shown in the screenshot used the original source folder instead; no personal files were moved, removed, or re-extracted during verification.

The intermittent stall was not reproduced in a fresh scan before the change. The revised task and completion path is covered by the tests above; a real multi-minute extraction remains a separate follow-up check. This fix is a local build until committed and published.
