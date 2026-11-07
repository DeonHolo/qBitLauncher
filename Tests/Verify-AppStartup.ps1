# Exercise the actual compiled app against an isolated portable installation.
param([Parameter(Mandatory = $true)][string]$ExecutablePath, [string]$ScreenshotDirectory, [switch]$Parallel)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class QBitStartupProbe {
    delegate bool Callback(IntPtr handle, IntPtr state);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, Callback callback, IntPtr state);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window, StringBuilder value, int length);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window, StringBuilder value, int length);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr window, int message, IntPtr first, IntPtr second);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr window, int message, IntPtr first, IntPtr second);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr window, IntPtr device, uint flags);
    public static void Capture(IntPtr window, string path, int width, int height) {
        using (System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(width,height)) {
            using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(bitmap)) {
                IntPtr device = graphics.GetHdc();
                try { if (!PrintWindow(window,device,0)) throw new InvalidOperationException("The app window could not be captured."); }
                finally { graphics.ReleaseHdc(device); }
            }
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
    }
    public static IntPtr FindList(IntPtr parent) {
        IntPtr result = IntPtr.Zero;
        EnumChildWindows(parent, delegate(IntPtr handle, IntPtr state) {
            StringBuilder name = new StringBuilder(256); GetClassName(handle,name,name.Capacity);
            if (name.ToString().IndexOf("SysListView32",StringComparison.OrdinalIgnoreCase)>=0) result = handle;
            return true;
        },IntPtr.Zero);
        return result;
    }
    public static bool HasLabel(IntPtr parent, string text) {
        bool found = false;
        EnumChildWindows(parent, delegate(IntPtr handle, IntPtr state) {
            StringBuilder name = new StringBuilder(256); GetWindowText(handle,name,name.Capacity);
            if (name.ToString() == text) found = true;
            return true;
        },IntPtr.Zero);
        return found;
    }
    public static int CountItems(IntPtr list) { return SendMessage(list,0x1004,IntPtr.Zero,IntPtr.Zero).ToInt32(); }
}
'@
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-startup-tests-' + [Guid]::NewGuid().ToString('N'))
$appProcess = $null
$parallelProcess = $null
try {
    $unicode = -join [char[]]@(0x6e38, 0x620f)
    $install = Join-Path $testRoot ('Install ' + $unicode + ' [test]')
    $fixture = Join-Path $testRoot ('Example ' + $unicode)
    [IO.Directory]::CreateDirectory($install) | Out-Null
    [IO.Directory]::CreateDirectory($fixture) | Out-Null
    $app = Join-Path $install 'qBitLauncher.exe'; [IO.File]::Copy([IO.Path]::GetFullPath($ExecutablePath), $app)
    $target = Join-Path $fixture ($unicode + '.exe'); [IO.File]::Copy((Join-Path $env:WINDIR 'System32\whoami.exe'), $target)
    [IO.File]::Copy((Join-Path $env:WINDIR 'System32\whoami.exe'), (Join-Path $fixture 'UnityCrashHandler64.exe'))
    $plugin = Join-Path $fixture 'Plugins'; [IO.Directory]::CreateDirectory($plugin) | Out-Null; [IO.File]::WriteAllText((Join-Path $plugin 'icons.zip'), 'fixture asset archive')
    [IO.File]::WriteAllText((Join-Path $install 'config.json'), '{"Theme":"Dracula","CheckForUpdates":false}', [Text.Encoding]::UTF8)
    $info = New-Object Diagnostics.ProcessStartInfo; $info.FileName = $app; $info.UseShellExecute = $false; $info.CreateNoWindow = $true
    $info.Arguments = '"' + $fixture + '"'; $info.WorkingDirectory = $install
    $info.EnvironmentVariables['QBITLAUNCHER_SKIP_UPDATE_CHECK'] = '1'; $info.EnvironmentVariables['QBITLAUNCHER_DATA_DIRECTORY'] = Join-Path $testRoot 'appdata'
    $appProcess = [Diagnostics.Process]::Start($info)
    $deadline = [DateTime]::UtcNow.AddSeconds(25); $window = $null; $list = [IntPtr]::Zero
    while ([DateTime]::UtcNow -lt $deadline) {
        $appProcess.Refresh()
        if ($appProcess.HasExited) { throw "Compiled app exited before opening its window (code $($appProcess.ExitCode))." }
        if ($appProcess.MainWindowHandle -ne [IntPtr]::Zero) {
            $list = [QBitStartupProbe]::FindList($appProcess.MainWindowHandle)
            if ($list -ne [IntPtr]::Zero -and [QBitStartupProbe]::CountItems($list) -eq 1) {
                $window = [Windows.Automation.AutomationElement]::FromHandle($appProcess.MainWindowHandle)
                break
            }
        }
        Start-Sleep -Milliseconds 150
    }
    if ($list -eq [IntPtr]::Zero -or [QBitStartupProbe]::CountItems($list) -ne 1) {
        $titles = $appProcess.MainWindowTitle
        $logPath = Join-Path $testRoot 'appdata\logs\qBitLauncher.log'
        $logTail = if ([IO.File]::Exists($logPath)) { (Get-Content -LiteralPath $logPath -Tail 10 -Encoding UTF8) -join "`n" } else { 'No isolated runtime log was written.' }
        throw "The compiled app did not reach executable selection. Windows: $titles`n$logTail"
    }
    if (-not [QBitStartupProbe]::HasLabel($appProcess.MainWindowHandle, 'Search')) { throw 'The compiled app did not expose the search label.' }
    if ($Parallel) {
        $smaller = Join-Path $testRoot ('Smaller game ' + $unicode); [IO.Directory]::CreateDirectory($smaller) | Out-Null
        [IO.File]::Copy((Join-Path $env:WINDIR 'System32\whoami.exe'), (Join-Path $smaller 'Smaller.exe'))
        $info.Arguments = '"' + $smaller + '"'; $parallelProcess = [Diagnostics.Process]::Start($info)
        $deadline = [DateTime]::UtcNow.AddSeconds(25); $parallelList = [IntPtr]::Zero
        while ([DateTime]::UtcNow -lt $deadline) {
            $appProcess.Refresh(); $parallelProcess.Refresh()
            if ($appProcess.HasExited -or $parallelProcess.HasExited) { throw 'A compiled workflow exited while checking simultaneous games.' }
            if ($parallelProcess.MainWindowHandle -ne [IntPtr]::Zero) {
                $parallelList = [QBitStartupProbe]::FindList($parallelProcess.MainWindowHandle)
                if ($parallelList -ne [IntPtr]::Zero -and [QBitStartupProbe]::CountItems($parallelList) -eq 1) { break }
            }
            Start-Sleep -Milliseconds 150
        }
        if ($parallelList -eq [IntPtr]::Zero -or [QBitStartupProbe]::CountItems($parallelList) -ne 1 -or [QBitStartupProbe]::CountItems($list) -ne 1) { throw 'The second game was blocked by the first open result window.' }
        if (-not [QBitStartupProbe]::HasLabel($parallelProcess.MainWindowHandle, 'Search')) { throw 'The second compiled workflow did not reach its result window.' }
        [void][QBitStartupProbe]::PostMessage($parallelProcess.MainWindowHandle, 0x10, [IntPtr]::Zero, [IntPtr]::Zero)
        if (-not $parallelProcess.WaitForExit(8000) -or $parallelProcess.ExitCode -ne 0) { throw 'The second compiled workflow did not close cleanly.' }
    }
    if ($ScreenshotDirectory) {
        [IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($ScreenshotDirectory)) | Out-Null
        $bounds = $window.Current.BoundingRectangle
        [QBitStartupProbe]::Capture($appProcess.MainWindowHandle, (Join-Path $ScreenshotDirectory 'qBitLauncher-3.4.0.png'), [int]$bounds.Width, [int]$bounds.Height)
    }
    [void][QBitStartupProbe]::PostMessage($appProcess.MainWindowHandle, 0x10, [IntPtr]::Zero, [IntPtr]::Zero)
    if (-not $appProcess.WaitForExit(8000)) { throw 'The compiled app did not close cleanly.' }
    if ($appProcess.ExitCode -ne 0) { throw "The compiled app exited with code $($appProcess.ExitCode)." }
    if (-not [IO.File]::Exists($target)) { throw 'Startup or closing changed the original fixture target.' }
    [pscustomobject]@{ Passed = $true; Suite = 'Compiled app startup'; Checks = $(if ($Parallel) { 8 } else { 5 }); UnicodeInstall = $true; NestedAssetArchiveSkipped = $true; IndependentGames = [bool]$Parallel }
}
finally {
    if ($appProcess) { if (-not $appProcess.HasExited) { $appProcess.Kill(); $appProcess.WaitForExit() }; $appProcess.Dispose() }
    if ($parallelProcess) { if (-not $parallelProcess.HasExited) { $parallelProcess.Kill(); $parallelProcess.WaitForExit() }; $parallelProcess.Dispose() }
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('qbitlauncher-startup-tests-') -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
