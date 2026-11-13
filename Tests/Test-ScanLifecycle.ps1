# Exercise the real scan controller with headless controls. No windows are opened.
param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
. (Join-Path $PSScriptRoot '..\Source\UI.ps1')
$script:checks = 0
$script:messages = New-Object System.Collections.ArrayList
$script:appearanceFails = $false
$Global:MainForm = $null
$Global:CurrentTheme = @{ FormBack = [Drawing.Color]::FromArgb(40, 42, 54); TextFore = [Drawing.Color]::White }
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
namespace ScanLifecycleFixture {
    public class Collection : List<object> {
        public void Add(object value, int column, int row) { Add(value); }
    }
    public class Control : IDisposable {
        public object Name, Text, ClientSize, MinimumSize, AutoScaleMode, StartPosition, Font, Dock, Padding, Tag, FlowDirection, CancelButton, ForeColor, BackColor;
        public bool MaximizeBox, AutoEllipsis, UseMnemonic, WrapContents, AutoSize, Enabled, IsDisposed, Closed;
        public int ColumnCount, RowCount;
        public Collection Controls = new Collection(), ColumnStyles = new Collection(), RowStyles = new Collection();
        EventHandler shown;
        public void Add_Shown(EventHandler handler) { shown += handler; }
        public void Add_Click(EventHandler handler) { }
        public void Close() { Closed = true; }
        public void Dispose() { IsDisposed = true; }
        public void ShowDialog() {
            State.Shows++;
            State.Window = this;
            if (!State.WithholdShown && shown != null) {
                try { shown(this, EventArgs.Empty); } catch (Exception) { }
            }
            for (int i = 0; i < 5 && !Closed; i++) {
                if (State.Advance != null) State.Advance(i);
                foreach (Timer timer in State.Timers.ToArray()) if (timer.Enabled) timer.Tick();
            }
            if (!Closed) throw new TimeoutException("Scan window never started polling; it still says Starting scan.");
        }
        public void ShowDialog(object owner) { ShowDialog(); }
    }
    public class Timer : IDisposable {
        public int Interval;
        public bool Enabled, Disposed;
        EventHandler tick;
        public Timer() { State.Timers.Add(this); }
        public void Add_Tick(EventHandler handler) { tick += handler; }
        public void Start() { Enabled = true; }
        public void Stop() { Enabled = false; }
        public void Tick() { if (tick != null) tick(this, EventArgs.Empty); }
        public void Dispose() { Disposed = true; Enabled = false; }
    }
    public static class State {
        public static bool WithholdShown;
        public static int Shows;
        public static Control Window;
        public static Action<int> Advance;
        public static readonly List<Timer> Timers = new List<Timer>();
    }
}
'@
function Assert-Lifecycle { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message }; $script:checks++ }
function Initialize-AppPresentation { }
function Set-ThemedButton { param($Button) }
function Set-FormIcon { param($Form); if ($script:appearanceFails) { throw 'fixture appearance failure' } }
function Write-LogMessage { param($Message); [void]$script:messages.Add($Message) }
function Get-7ZipGuiPath { return $null }
function Start-FolderScan { param($RootFolder, $ArchivePath, $ListToolPath); return $script:scan }
function ConvertTo-FolderInventory { param($ScanResult); return $ScanResult }
function New-Object {
    param([Parameter(Position = 0)][string]$TypeName, [Parameter(Position = 1)][object[]]$ArgumentList)
    if ($TypeName -eq 'Windows.Forms.Timer') { return [ScanLifecycleFixture.Timer]::new() }
    if ($TypeName -like 'Windows.Forms.*') { return [ScanLifecycleFixture.Control]::new() }
    return Microsoft.PowerShell.Utility\New-Object -TypeName $TypeName -ArgumentList $ArgumentList
}
function New-LifecycleScan {
    $completion = [Threading.Tasks.TaskCompletionSource[object]]::new()
    $raw = [pscustomobject]@{ Complete = $true; Runnables = @('fixture.exe'); Skipped = @(); OutputCheck = $null }
    $item = [pscustomobject]@{
        Operation = $completion.Task; Completion = $completion; Raw = $raw; Disposed = $false; Cancelled = $false
        Progress = [pscustomobject]@{ Phase = 'Scanning files'; Files = 4; Folders = 2; Runnables = 1; Skipped = 0; CurrentPath = 'C:\scan-fixture' }
    }
    $item | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $this.Disposed = $true }
    $item | Add-Member -MemberType ScriptMethod -Name Cancel -Value { $this.Cancelled = $true }
    $item | Add-Member -MemberType ScriptMethod -Name Snapshot -Value { param($complete); return $this.Raw }
    return $item
}
try {
    foreach ($case in @('NoShownEvent', 'AppearanceFailure', 'Normal')) {
        [ScanLifecycleFixture.State]::Timers.Clear(); [ScanLifecycleFixture.State]::Shows = 0
        [ScanLifecycleFixture.State]::WithholdShown = $case -eq 'NoShownEvent'
        $script:appearanceFails = $case -eq 'AppearanceFailure'
        $script:scan = New-LifecycleScan
        [ScanLifecycleFixture.State]::Advance = { param($step); if ($step -eq 1) { $script:scan.Completion.SetResult($script:scan.Raw) } }
        $inventory = Get-InventoryResponsive -Path 'C:\scan-fixture'
        Assert-Lifecycle ($inventory.Complete -and $inventory.Runnables.Count -eq 1) "$case must still return the discovered files."
        Assert-Lifecycle ($script:scan.Disposed -and [ScanLifecycleFixture.State]::Timers[0].Disposed) "$case leaked the scanner or timer."
        if ($case -eq 'AppearanceFailure') {
            Assert-Lifecycle ([ScanLifecycleFixture.State]::Window.BackColor -eq $Global:CurrentTheme.FormBack) 'The scan must keep its selected background when icon/title styling fails.'
            Assert-Lifecycle (@($script:messages | Where-Object { $_ -match 'appearance.*fixture appearance failure' }).Count -eq 1) 'An appearance failure must be logged instead of disabling scan progress silently.'
        }
    }
    [ScanLifecycleFixture.State]::Timers.Clear(); [ScanLifecycleFixture.State]::Shows = 0
    $script:appearanceFails = $false; $script:scan = New-LifecycleScan
    $script:scan.Completion.SetResult($script:scan.Raw)
    $inventory = Get-InventoryResponsive -Path 'C:\scan-fixture'
    Assert-Lifecycle ($inventory.Complete -and [ScanLifecycleFixture.State]::Shows -eq 0) 'A completed scan should return its results without opening a progress window.'
    Assert-Lifecycle $script:scan.Disposed 'A completed scan leaked its worker.'

    foreach ($case in @('BeforeWindow', 'DuringPolling')) {
        [ScanLifecycleFixture.State]::Timers.Clear(); [ScanLifecycleFixture.State]::Shows = 0
        $script:scan = New-LifecycleScan
        if ($case -eq 'BeforeWindow') { $script:scan.Completion.SetException([IO.IOException]::new('fixture disk failure')) }
        [ScanLifecycleFixture.State]::Advance = { param($step); if ($step -eq 1) { $script:scan.Completion.SetException([IO.IOException]::new('fixture disk failure')) } }
        $failure = ''
        try { $null = Get-InventoryResponsive -Path 'C:\scan-fixture' } catch { $failure = $_.Exception.Message }
        Assert-Lifecycle ($failure -match 'fixture disk failure') "$case must surface the actual scan error."
        Assert-Lifecycle ($script:scan.Disposed -and [ScanLifecycleFixture.State]::Timers[0].Disposed) "$case leaked a failed scanner or timer."
    }

    $result = @{ Passed = $true; Checks = $script:checks; Suite = 'ScanLifecycle'; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { [pscustomobject]$result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'ScanLifecycle' } | ConvertTo-Json)); exit 1 }
    throw
}
finally { [ScanLifecycleFixture.State]::Advance = $null }
