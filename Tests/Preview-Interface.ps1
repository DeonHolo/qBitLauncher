# Capture the real forms with disposable, representative files and isolated state.
param([Parameter(Mandatory = $true)][string]$ScreenshotDirectory, [ValidateSet('Dracula', 'Light')][string]$Theme = 'Dracula', [ValidateSet('Selection', 'Cleanup')][string]$View = 'Selection', [switch]$CollapsedColumns, [switch]$EditShortcut, [string]$FixtureParent)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
public static class QBitClientPreview {
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr window, IntPtr device, uint flags);
    public static void Capture(IntPtr window, string path, int width, int height) {
        using (Bitmap bitmap = new Bitmap(width,height)) {
            using (Graphics graphics = Graphics.FromImage(bitmap)) {
                IntPtr device = graphics.GetHdc();
                try { if (!PrintWindow(window,device,1)) throw new InvalidOperationException("Client preview capture failed."); }
                finally { graphics.ReleaseHdc(device); }
            }
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
    }
}
'@
$repository = Split-Path $PSScriptRoot -Parent
$source = Join-Path $repository 'qBitLauncher.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join "`n") }
foreach ($definition in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    Invoke-Expression ('function global:' + $definition.Name + ' ' + $definition.Body.Extent.Text)
}
$palette = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -eq '$Global:Themes' } | Select-Object -First 1
Invoke-Expression $palette.Extent.Text
$tempBase = if ($FixtureParent) { [IO.Path]::GetFullPath($FixtureParent) } else { [IO.Path]::GetFullPath([IO.Path]::GetTempPath()) }
if (-not [IO.Directory]::Exists($tempBase)) { throw 'The preview fixture parent must be an existing directory.' }
$suffix = [Guid]::NewGuid().ToString('N')
if ($FixtureParent) { $suffix = $suffix.Substring(0, 12) }
$testRoot = Join-Path $tempBase ('qbitlauncher-preview-' + $suffix)
$timer = $null
try {
    $fixture = Join-Path $testRoot 'Moonlit Garden'; [IO.Directory]::CreateDirectory($fixture) | Out-Null
    $Global:DataDirectory = Join-Path $testRoot 'appdata'; $Global:ConfigFile = Join-Path $testRoot 'config.json'; $Global:LogFile = Join-Path $testRoot 'preview.log'
    $Global:ThemeSelection = $Theme; $Global:CurrentTheme = $Global:Themes[$Theme]; $Global:UserSettings = New-DefaultUserSettings
    $Global:UserSettings.Theme = $Theme; $Global:UserSettings.DefaultAdmin = $false; $Global:UserSettings.CheckForUpdates = $false
    $Global:MainForm = $null; $Global:AppIcon = $null; $Global:PendingUpdateCheck = $null; $Global:AvailableVersion = $null; $Global:ScriptDir = $repository
    $Global:ScriptVersion = [regex]::Match([IO.File]::ReadAllText($source), '(?m)^\$Global:ScriptVersion\s*=\s*"([^\"]+)"').Groups[1].Value
    $Global:filePathFromQB = $fixture
    if ($CollapsedColumns) { $Global:UserSettings.WindowState = @{ Columns = @(0, 130, 0, 1000) } }
    $archive = Join-Path $testRoot 'Moonlit Garden.rar'; [IO.File]::WriteAllText($archive, 'Disposable UI preview')
    $files = @()
    foreach ($name in @('MoonlitGarden.exe', 'MoonlitGarden_DX12.exe', 'UnityCrashHandler64.exe', 'uninstall.exe')) {
        $path = Join-Path $fixture $name; [IO.File]::Copy((Join-Path $env:WINDIR 'System32\whoami.exe'), $path); $files += Get-Item -LiteralPath $path
    }
    $scriptFile = Join-Path $fixture 'Launch game.cmd'; [IO.File]::WriteAllText($scriptFile, '@echo off'); $files += Get-Item -LiteralPath $scriptFile
    [void](Set-FilePreference -Path $files[0].FullName -Preference @{ Label = 'Moonlit Garden'; AsAdmin = $false })
    [void](Set-FilePreference -Path $files[1].FullName -Preference @{ Label = 'Moonlit Garden · DirectX 12'; AsAdmin = $false })
    [void](Set-FilePreference -Path $fixture -Preference @{ PreferredExecutable = $files[0].FullName })
    $fileName = if ($View -eq 'Cleanup') { 'qBitLauncher-Cleanup-' + $Theme + '.png' } elseif ($EditShortcut) { 'qBitLauncher-Editing-' + $Theme + '.png' } elseif ($CollapsedColumns) { 'qBitLauncher-Columns-' + $Theme + '.png' } else { 'qBitLauncher-' + $Theme + '.png' }
    $state = @{ Ticks = 0; Error = $null; Saved = $false; Selected = $false; Editing = $false }
    $timer = New-Object Windows.Forms.Timer; $timer.Interval = 140
    $timer.Add_Tick({
        try {
            $form = if ($View -eq 'Cleanup') { [Windows.Forms.Application]::OpenForms | Where-Object { $_.Name -eq 'CleanupForm' } | Select-Object -First 1 } else { $Global:MainForm }
            if (-not $form -or $form.Tag -ne 'Ready') { return }
            if ($EditShortcut -and -not $state.Editing) {
                $state.Editing = $true
                $list = $form.Controls.Find('RunnableList', $true)[0]; $list.Focus() | Out-Null; $list.SelectedItems[0].BeginEdit()
                return
            }
            if ($View -eq 'Cleanup') {
                if (-not $state.Selected) {
                    foreach ($name in @('RemoveSourceFiles', 'RemoveExtractedFiles', 'RemoveTorrentEntry')) { $form.Controls.Find($name, $true)[0].Checked = $true }
                    $state.Selected = $true
                }
                if ($form.Controls.Find('CleanupPreview', $true)[0].Text -eq 'Checking selected locations...') { return }
            }
            $state.Ticks++
            if ($state.Ticks -lt 4) { return }
            if ($EditShortcut) {
                $editors = $form.Controls.Find('ShortcutNameEditor', $true)
                if ($editors.Count) { $editors[0].SelectionStart = $editors[0].TextLength; $editors[0].SelectionLength = 0 }
            }
            if ($View -eq 'Cleanup') { $form.Controls.Find('AcknowledgeAppDeletion', $true)[0].Checked = $true }
            $timer.Stop()
            $debug = New-Object 'System.Collections.Generic.List[object]'
            $pending = New-Object 'System.Collections.Generic.Stack[System.Windows.Forms.Control]'; $pending.Push($form)
            while ($pending.Count) {
                $control = $pending.Pop()
                $debug.Add([pscustomobject]@{ Type = $control.GetType().Name; Name = $control.Name; Text = $control.Text; Visible = $control.Visible; Bounds = $control.Bounds.ToString(); Parent = $(if ($control.Parent) { $control.Parent.GetType().Name } else { '' }) })
                foreach ($child in $control.Controls) { $pending.Push($child) }
            }
            [IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($ScreenshotDirectory)) | Out-Null
            [IO.File]::WriteAllText((Join-Path $ScreenshotDirectory 'preview-controls.json'), ($debug.ToArray() | ConvertTo-Json -Depth 4))
            [QBitClientPreview]::Capture($form.Handle, (Join-Path $ScreenshotDirectory $fileName), $form.ClientSize.Width, $form.ClientSize.Height)
            $state.Saved = $true; $form.Close()
        } catch { $state.Error = $_; $timer.Stop(); if ($Global:MainForm) { $Global:MainForm.Close() } }
    }.GetNewClosure())
    $timer.Start()
    if ($View -eq 'Cleanup') { $null = Show-CleanupConfirmForm -SourcePath $archive -RootFolder $fixture -SelectedPath $files[0].FullName -TorrentHash ('a' * 40) -TorrentLabel 'Moonlit Garden.rar' }
    else { Show-ExecutableSelectionForm -FoundExecutables $files -WindowTitle 'qBitLauncher' -RootFolder $fixture }
    if ($state.Error) { throw $state.Error }
    if (-not $state.Saved) { throw 'No preview was captured.' }
    Join-Path $ScreenshotDirectory $fileName
} finally {
    if ($timer) { $timer.Stop(); $timer.Dispose() }
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('qbitlauncher-preview-') -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
