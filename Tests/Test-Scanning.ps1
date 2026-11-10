param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
. (Join-Path $repository 'Source\Workflow.ps1')
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-scan-tests-' + [Guid]::NewGuid().ToString('N'))
$checks = 0; $scan = $null; $largeFile = $null; $junction = $null
function Assert-Scan { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message }; $script:checks++ }
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
public static class ScanTestFixtures {
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool DeviceIoControl(SafeFileHandle file,uint code,IntPtr input,int inputSize,IntPtr output,int outputSize,out int returned,IntPtr overlapped);
    public static void Populate(string root,int count) {
        Directory.CreateDirectory(root);
        for(int i=0;i<count;i++) { string folder=Path.Combine(root,"assets",(i%80).ToString()); Directory.CreateDirectory(folder); File.WriteAllText(Path.Combine(folder,i.ToString()+".dat"),"scan fixture"); }
    }
    public static FileStream SparsePayload(string path) {
        FileStream stream=new FileStream(path,FileMode.CreateNew,FileAccess.ReadWrite,FileShare.None);
        int returned;
        if(!DeviceIoControl(stream.SafeFileHandle,0x900C4,IntPtr.Zero,0,IntPtr.Zero,0,out returned,IntPtr.Zero)) { stream.Dispose(); throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error()); }
        stream.SetLength(12L*1024*1024*1024); return stream;
    }
}
'@
try {
    $fixture = Join-Path $testRoot ('Large [fixture] ' + (-join [char[]]@(0x6e38, 0x620f)))
    [ScanTestFixtures]::Populate($fixture, 12000)
    foreach ($name in @('Game.exe', 'launch.CMD', 'UnityCrashHandler64.exe', 'payload.zip')) { [IO.File]::WriteAllText((Join-Path $fixture $name), 'fixture') }
    $largeFile = [ScanTestFixtures]::SparsePayload((Join-Path $fixture 'setup-1.bin'))
    Initialize-FolderScanner
    $scan = New-Object QBitLauncher.FolderScan($fixture)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not $scan.Operation.IsCompleted -and $watch.Elapsed.TotalSeconds -lt 20) { Start-Sleep -Milliseconds 10 }
    Assert-Scan $scan.Operation.IsCompleted 'A large directory scan did not finish within 20 seconds.'
    $raw = $scan.Operation.GetAwaiter().GetResult()
    $watch.Stop()
    $inventory = ConvertTo-FolderInventory $raw
    Assert-Scan ($inventory.Complete -and $scan.Progress.Complete) 'Successful scans must report completion independently of PowerShell worker polling.'
    Assert-Scan ($scan.Progress.Files -eq 12005 -and $scan.Progress.Folders -eq 82) 'Progress counters did not cover the complete large fixture.'
    Assert-Scan ($inventory.Runnables.Count -eq 3 -and $inventory.Archives.Count -eq 1) 'Streaming discovery lost runnables or root payloads.'
    Assert-Scan ($inventory.Runnables[-1].Name -eq 'UnityCrashHandler64.exe') 'Streaming discovery changed the helper ranking.'
    Assert-Scan ($largeFile.Length -eq 12GB -and $watch.Elapsed.TotalSeconds -lt 20) 'Scanning should finish with a locked 12 GB payload without opening its contents.'
    Assert-Scan ($scan.Snapshot($false).Runnables.Count -eq 3) 'Found-file snapshots lost runnable results.'
    $scan.Dispose(); $scan = $null

    $scan = New-Object QBitLauncher.FolderScan($fixture)
    $cancelWatch = [Diagnostics.Stopwatch]::StartNew(); $scan.Cancel(); $scan.Dispose()
    Assert-Scan ($cancelWatch.Elapsed.TotalMilliseconds -lt 500) 'Cancelling a scan blocked while waiting for the worker.'
    $deadline = [DateTime]::UtcNow.AddSeconds(2)
    while (-not $scan.Operation.IsCompleted -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 10 }
    Assert-Scan $scan.Operation.IsCompleted 'Cancelled scan did not settle.'
    $scan = $null

    $outside = Join-Path $testRoot 'outside'; [IO.Directory]::CreateDirectory($outside) | Out-Null
    [IO.File]::WriteAllText((Join-Path $outside 'outside.exe'), 'fixture')
    $linkRoot = Join-Path $testRoot 'links'; [IO.Directory]::CreateDirectory($linkRoot) | Out-Null
    [IO.File]::WriteAllText((Join-Path $linkRoot 'inside.exe'), 'fixture')
    $junction = Join-Path $linkRoot 'loop'
    [void](New-Item -ItemType Junction -Path $junction -Target $linkRoot -ErrorAction Stop)
    $otherLink = Join-Path $linkRoot 'outside'
    [void](New-Item -ItemType Junction -Path $otherLink -Target $outside -ErrorAction Stop)
    try {
        $inventory = Get-FolderInventory $linkRoot
        Assert-Scan ($inventory.Runnables.Count -eq 1 -and $inventory.Skipped.Count -eq 2) 'Scanning followed a junction cycle or entered an outside folder.'
    } finally { if ([IO.Directory]::Exists($otherLink)) { [IO.Directory]::Delete($otherLink, $false) } }
    [IO.Directory]::Delete($junction, $false); $junction = $null

    $empty = Join-Path $testRoot 'empty'; [IO.Directory]::CreateDirectory($empty) | Out-Null
    $inventory = Get-FolderInventory $empty
    Assert-Scan ($inventory.Complete -and $inventory.Runnables.Count -eq 0) 'An empty folder scan should complete with zero results.'
    $failed = $false
    try { $null = Get-FolderInventory (Join-Path $testRoot 'missing') } catch { $failed = $true }
    Assert-Scan $failed 'A missing scan root was silently reported as a successful scan.'
    $result = [pscustomobject]@{ Passed = $true; Checks = $checks; Suite = 'Scanning'; PowerShellVersion = $PSVersionTable.PSVersion.ToString(); Files = 12005; LargePayloadGB = 12; ScanSeconds = [Math]::Round($watch.Elapsed.TotalSeconds, 3) }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { $result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $checks; Error = $_.Exception.Message; Suite = 'Scanning' } | ConvertTo-Json)); exit 1 }
    throw
}
finally {
    if ($scan) { $scan.Dispose() }
    if ($largeFile) { $largeFile.Dispose() }
    if ($junction -and [IO.Directory]::Exists($junction)) { [IO.Directory]::Delete($junction, $false) }
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('qbitlauncher-scan-tests-') -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
