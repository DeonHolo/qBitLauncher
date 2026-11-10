# Noninteractive WinForms smoke checks. Import definitions only; never run app startup.
param(
    [string]$SourcePath,
    [string]$ResultPath,
    [ValidateSet('All', 'Selection', 'Cleanup', 'Settings', 'Setup', 'Scanning')][string]$Case = 'All'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
if ([string]::IsNullOrWhiteSpace($SourcePath)) { $SourcePath = Join-Path $PSScriptRoot '..\qBitLauncher.ps1' }
$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-ui-tests-' + [Guid]::NewGuid().ToString('N'))
$Global:UIHarnessState = @{ Checks = 0; ThreadError = $null }
$importedFunctions = New-Object System.Collections.ArrayList
$originalFunctions = @{}
$globals = @{}
foreach ($name in @('DataDirectory', 'ConfigFile', 'UserSettings', 'Themes', 'ThemeSelection', 'CurrentTheme', 'MainForm', 'ScriptDir', 'AppIcon', 'PendingUpdateCheck', 'AvailableVersion', 'ScriptVersion', 'filePathFromQB', 'PasswordDecryptionFailed', 'ProtectedSettingsPassword', 'SettingsBaseline', 'SettingsBaselinePath', 'UIHarnessState')) {
    if ($name -eq 'UIHarnessState') { continue }
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $globals[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}

function global:Assert-UICheck {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $Global:UIHarnessState.Checks++
}

function global:Get-UITestControl {
    param([Windows.Forms.Control]$Form, [string]$Name)
    $controls = $Form.Controls.Find($Name, $true)
    if (-not $controls.Count) { throw "The form is missing control '$Name'." }
    return $controls[0]
}

function global:Get-UITestForm {
    param([string]$Name)
    foreach ($form in @([Windows.Forms.Application]::OpenForms)) {
        if ($form.Name -eq $Name) { return $form }
    }
    return $null
}

function global:Close-UITestForms {
    foreach ($form in @([Windows.Forms.Application]::OpenForms)) {
        if (-not $form.IsDisposed) { $form.Close() }
    }
}

function global:Select-UITestRunnable {
    param([Windows.Forms.ListView]$List, [string]$Path)
    foreach ($item in $List.Items) { $item.Selected = $false }
    foreach ($item in $List.Items) {
        if ([string]::Equals([string]$item.Tag, $Path, [StringComparison]::OrdinalIgnoreCase)) {
            $item.Selected = $true
            $item.Focused = $true
            return $item
        }
    }
    throw "The runnable list is missing a fixture: $Path"
}

function Invoke-UITestCase {
    param([string]$Name, [scriptblock]$Action, [scriptblock]$OnTick)
    $state = @{ Step = 0; Deadline = [DateTime]::UtcNow.AddSeconds(15); Error = $null; Finished = $false; Result = $null }
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = 100
    $timer.Add_Tick({
        if ($state.Finished) { return }
        try {
            if ($Global:UIHarnessState.ThreadError) { throw $Global:UIHarnessState.ThreadError }
            if ([DateTime]::UtcNow -gt $state.Deadline) { throw "Timed out while checking $Name (step $($state.Step))." }
            & $OnTick $state
        }
        catch {
            $state.Error = $_
            $state.Finished = $true
            $timer.Stop()
            Close-UITestForms
        }
    }.GetNewClosure())
    $timer.Start()
    try {
        $state.Result = & $Action
        if ($state.Error) { throw $state.Error }
        if ($Global:UIHarnessState.ThreadError) { throw $Global:UIHarnessState.ThreadError }
        Assert-UICheck $state.Finished ("The $Name form closed before all smoke checks ran.")
        return $state.Result
    }
    finally { $timer.Stop(); $timer.Dispose(); Close-UITestForms }
}

$threadHandler = [Threading.ThreadExceptionEventHandler]{
    param($sender, $eventArgs)
    $Global:UIHarnessState.ThreadError = $eventArgs.Exception
    Close-UITestForms
}
[Windows.Forms.Application]::SetUnhandledExceptionMode([Windows.Forms.UnhandledExceptionMode]::CatchException)
[Windows.Forms.Application]::add_ThreadException($threadHandler)

try {
    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile([IO.Path]::GetFullPath($SourcePath), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw ($parseErrors.Message -join "`n") }
    $definitions = $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)
    foreach ($definition in $definitions) {
        if (-not $originalFunctions.ContainsKey($definition.Name)) {
            $existing = Get-Command $definition.Name -CommandType Function -ErrorAction SilentlyContinue
            $originalFunctions[$definition.Name] = $(if ($existing) { $existing.ScriptBlock } else { $null })
            [void]$importedFunctions.Add($definition.Name)
        }
        Invoke-Expression ('function global:' + $definition.Name + ' ' + $definition.Body.Extent.Text)
    }
    # Any unexpected action is an immediate failure, rather than a desktop/API side effect.
    foreach ($name in @('Invoke-CleanupPlan', 'Remove-QBittorrentTorrent', 'Install-QBittorrentIntegration', 'Install-ContextMenu', 'Remove-QBittorrentIntegration', 'Remove-ContextMenu', 'New-WindowsShortcut')) {
        Set-Item -Path ('Function:global:' + $name) -Value { throw 'This UI smoke test forbids file deletion, program launches, desktop shortcuts and integration changes.' }
    }
    Set-Item -Path Function:global:Invoke-ActionSound -Value { param($Type) }
    Set-Item -Path Function:global:Show-ThemedMessageBox -Value { param($Message, $Title, $Buttons, $Icon) throw ("Unexpected UI message: $Title - $Message") }

    [IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $Global:DataDirectory = Join-Path $testRoot 'appdata'
    $Global:ConfigFile = Join-Path $testRoot 'config.json'
    $script:LogFile = Join-Path $testRoot 'ui.log'
    $Global:ScriptDir = $repositoryRoot
    $Global:ScriptVersion = '3.3.7'
    $Global:AppIcon = $null
    $Global:MainForm = $null
    $Global:PendingUpdateCheck = $null
    $Global:AvailableVersion = $null
    $Global:ThemeSelection = 'Dracula'
    $colors = @{
        FormBack = [Drawing.Color]::FromArgb(40, 42, 54); TextFore = [Drawing.Color]::White
        ControlBack = [Drawing.Color]::FromArgb(32, 32, 32); ButtonBack = [Drawing.Color]::FromArgb(47, 52, 64)
        Border = [Drawing.Color]::Gray; Accent = [Drawing.Color]::Cyan; SecondaryText = [Drawing.Color]::LightGray
    }
    $Global:Themes = @{ Dracula = $colors; Light = $colors }
    $Global:CurrentTheme = $colors
    $Global:UserSettings = New-DefaultUserSettings
    $Global:UserSettings.CheckForUpdates = $false
    $Global:PasswordDecryptionFailed = $false
    $Global:ProtectedSettingsPassword = ''

    $sourceDirectory = Join-Path $testRoot 'Download [fixture]'
    $extractedDirectory = Join-Path $sourceDirectory 'Extracted'
    [IO.Directory]::CreateDirectory($extractedDirectory) | Out-Null
    $archive = Join-Path $sourceDirectory 'installer.zip'
    [IO.File]::WriteAllText($archive, 'UI fixture; not a real archive')
    $unicodeName = -join [char[]]@(0x6e38, 0x620f)
    $game = Join-Path $extractedDirectory 'Game.exe'
    $unicodeGame = Join-Path $extractedDirectory ($unicodeName + '.exe')
    $bat = Join-Path $extractedDirectory 'Start [game].bat'
    $cmd = Join-Path $extractedDirectory 'Play [game].cmd'
    $helper = Join-Path $extractedDirectory 'UnityCrashHandler64.exe'
    foreach ($file in @($game, $unicodeGame, $helper)) { [IO.File]::Copy((Join-Path $env:WINDIR 'System32\whoami.exe'), $file) }
    foreach ($file in @($bat, $cmd)) { [IO.File]::WriteAllText($file, '@echo off') }
    $fixtures = @($game, $unicodeGame, $bat, $cmd, $helper)
    $originalBytes = @{}
    foreach ($file in $fixtures) { $originalBytes[$file] = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file)) }
    $files = @($fixtures | ForEach-Object { Get-Item -LiteralPath $_ })
    $Global:filePathFromQB = $sourceDirectory
    [void](Set-FilePreference -Path $game -Preference @{ Label = 'Original preference'; Arguments = '--windowed'; AsAdmin = $false })
    [void](Set-FilePreference -Path $extractedDirectory -Preference @{ PreferredExecutable = $game })
    [void](Add-RecentFolder $extractedDirectory)

    if ($Case -in @('All', 'Selection')) {
        $aliases = @{ $game = ('Game ' + $unicodeName); $bat = 'My launch script'; $cmd = 'My command script' }
        $selectionTick = {
            param($state)
            $form = Get-UITestForm 'ExecutableSelectionForm'
            if (-not $form) { return }
            if ($form.Tag -ne 'Ready') { return }
            $form.Opacity = 0
            $list = Get-UITestControl $form 'RunnableList'
            $search = Get-UITestControl $form 'SearchBox'
            $showAll = Get-UITestControl $form 'ShowAllCheckBox'
            $arguments = Get-UITestControl $form 'LaunchArgumentsTextBox'
            $admin = Get-UITestControl $form 'RunAsAdminCheckBox'
            if ($state.Step -eq 0) {
                Assert-UICheck ($form.FormBorderStyle -eq 'Sizable' -and $form.MaximizeBox) 'The executable window must be resizable.'
                Assert-UICheck ($form.AutoScaleMode -eq 'Dpi') 'The executable window must use DPI scaling.'
                Assert-UICheck ($list.LabelEdit -and -not $showAll.Checked -and $list.Items.Count -eq 4) ("Default results: expected 4 files with editing enabled and helpers hidden; got Count=$($list.Items.Count), LabelEdit=$($list.LabelEdit), ShowAll=$($showAll.Checked).")
                Assert-UICheck ($list.Columns[0].Text -eq 'Shortcut name' -and $list.Columns[1].Text -eq 'Original file') 'The UI must distinguish shortcut labels from original filenames.'
                Assert-UICheck ($list.OwnerDraw -and $list.BorderStyle -eq 'None') 'The result list should draw its themed header and selection without a bright native frame.'
                Assert-UICheck ((($list.Columns | Measure-Object -Property Width -Sum).Sum) -le $list.ClientSize.Width) 'The result columns should fit without a horizontal scrollbar.'
                $list.SetPreferredColumns(@(0, 10, 0, 800))
                Assert-UICheck ($list.Columns[0].Width -ge 80 -and $list.Columns[1].Width -ge 80 -and $list.Columns[2].Width -ge 60) 'Collapsed saved columns must recover visible shortcut, filename, and type headers.'
                $list.SetPreferredColumns(@(180, 170, 95))
                Assert-UICheck ($search.BorderStyle -eq 'None' -and $search.Parent.Parent.ClientSize.Width -ge $search.Parent.Right) 'Search should use the themed field and stay inside its layout.'
                $more = Get-UITestControl $form 'MoreButton'
                Assert-UICheck ($more.Visible -and -not $arguments.Visible) 'Secondary launch options should be tucked under More by default.'
                $more.ContextMenuStrip.Items[0].PerformClick()
                Assert-UICheck ($arguments.Visible -and $admin.Visible) 'Launch options must remain reachable from More.'
                $more.ContextMenuStrip.Items[0].PerformClick()
                Assert-UICheck (-not $arguments.Visible) 'Collapsing launch options should restore the compact layout.'
                $search.Text = 'UnityCrashHandler'
                Assert-UICheck ($list.Items.Count -eq 0) 'A hidden helper appeared without Show all.'
                $showAll.Checked = $true
                Assert-UICheck ($list.Items.Count -eq 1 -and [string]$list.Items[0].Tag -ceq $helper) 'Show all must make the filtered helper reachable.'
                $search.Text = $unicodeName
                Assert-UICheck ($list.Items.Count -eq 1 -and [string]$list.Items[0].Tag -ceq $unicodeGame) 'Search must match Unicode runnable names.'
                $search.Text = ''
                $showAll.Checked = $false
                $method = $list.GetType().GetMethod('OnAfterLabelEdit', [Reflection.BindingFlags]'Instance,NonPublic')
                foreach ($path in @($game, $bat, $cmd)) {
                    $item = Select-UITestRunnable $list $path
                    $eventArgs = [Windows.Forms.LabelEditEventArgs]::new($item.Index, $aliases[$path])
                    $eventArguments = New-Object object[] 1
                    $eventArguments[0] = $eventArgs.PSObject.BaseObject
                    [void]$method.Invoke($list, $eventArguments)
                    Assert-UICheck $eventArgs.CancelEdit 'The native file-label edit must be cancelled after saving the shortcut alias.'
                    Assert-UICheck ((Get-FilePreference $path).Label -ceq $aliases[$path]) 'A shortcut name was not saved for its original file.'
                }
                $state.Step = 1
                return
            }
            if ($state.Step -eq 1) {
                foreach ($path in @($game, $bat, $cmd)) {
                    $item = Select-UITestRunnable $list $path
                    Assert-UICheck ($item.Text -ceq $aliases[$path]) 'The shortcut label did not update in the list.'
                    Assert-UICheck ($item.SubItems[1].Text -ceq [IO.Path]::GetFileName($path)) 'The original filename column changed after editing an alias.'
                    Assert-UICheck ([IO.File]::Exists($path) -and [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) -ceq $originalBytes[$path]) 'Editing a shortcut label changed an EXE/BAT/CMD file.'
                }
                $null = Select-UITestRunnable $list $game
                Assert-UICheck ($arguments.Text -ceq '--windowed' -and -not $admin.Checked) 'Selecting a runnable must restore its launch preferences.'
                $arguments.Text = '--windowed --ui-smoke'
                $admin.Checked = $true
                $list.Columns[0].Width = 225
                $form.Size = New-Object Drawing.Size(1180, 760)
                $state.Step = 2
                return
            }
            $state.Finished = $true
            $form.Close()
        }.GetNewClosure()
        $selectionAction = { Show-ExecutableSelectionForm -FoundExecutables $files -WindowTitle 'qBitLauncher UI smoke fixture' -RootFolder $extractedDirectory }.GetNewClosure()
        $null = Invoke-UITestCase -Name 'executable selection' -Action $selectionAction -OnTick $selectionTick
        Assert-UICheck ((Get-FilePreference $game).Arguments -ceq '--windowed --ui-smoke' -and (Get-FilePreference $game).AsAdmin) 'Closing the form must persist launch preferences.'
        Assert-UICheck ($Global:UserSettings.WindowState.Width -ge 780 -and $Global:UserSettings.WindowState.Height -ge 540 -and $Global:UserSettings.WindowState.Columns[0] -eq 225) 'Closing the form must persist window and column dimensions.'
        Get-UserSettings
        Assert-UICheck ((Get-FilePreference $bat).Label -ceq $aliases[$bat] -and (Get-FilePreference $cmd).Label -ceq $aliases[$cmd]) 'BAT/CMD aliases did not survive reloading settings.'
        $reopenTick = {
            param($state)
            $form = Get-UITestForm 'ExecutableSelectionForm'
            if (-not $form -or $form.Tag -ne 'Ready') { return }
            $form.Opacity = 0
            $list = Get-UITestControl $form 'RunnableList'
            Assert-UICheck ($list.Items.Count -eq 1 -and $list.Items[0].Text -ceq $aliases[$game]) 'Reopening a single runnable must restore its shortcut name.'
            Assert-UICheck ((Get-UITestControl $form 'LaunchArgumentsTextBox').Text -ceq '--windowed --ui-smoke') 'Reopening did not restore saved launch arguments.'
            Assert-UICheck ((Get-UITestControl $form 'RunAsAdminCheckBox').Checked) 'Reopening did not restore saved admin mode.'
            Assert-UICheck ($list.Columns[0].Width -eq 225) 'Reopening did not restore column widths.'
            $state.Finished = $true
            $form.Close()
        }.GetNewClosure()
        $reopenAction = { Show-ExecutableSelectionForm -FoundExecutables @((Get-Item -LiteralPath $game)) -RootFolder $extractedDirectory }.GetNewClosure()
        $null = Invoke-UITestCase -Name 'remembered selection' -Action $reopenAction -OnTick $reopenTick
        $editingTick = {
            param($state)
            $form = Get-UITestForm 'ExecutableSelectionForm'
            if (-not $form -or $form.Tag -ne 'Ready') { return }
            $form.Opacity = 0
            $list = Get-UITestControl $form 'RunnableList'
            $item = $list.Items[0]
            if ($state.Step -eq 0) {
                $list.Focus() | Out-Null
                $item.BeginEdit()
                $state.Step = 1; return
            }
            if ($state.Step -eq 1) {
                $editor = Get-UITestControl $form 'ShortcutNameEditor'
                Assert-UICheck ($list.IsShortcutNameEditing -and $form.Controls.Find('ShortcutNameEditor', $true).Count -eq 1) 'Inline editing must create exactly one name editor.'
                Assert-UICheck ($editor.Parent.Parent -eq $list -and $editor.Parent.Top -ge $item.Bounds.Top -and $editor.Parent.Bottom -le $item.Bounds.Bottom) 'The name editor must stay inside the existing row.'
                $list.Update()
                $hasIcon = $list.SmallImageList -and $item.ImageIndex -ge 0 -and $item.ImageIndex -lt $list.SmallImageList.Images.Count
                $textLeft = $item.Bounds.Left + 10
                if ($hasIcon) { $textLeft += [Math]::Min(20, $item.Bounds.Height - 4) + 8 }
                Assert-UICheck ($editor.Parent.Left -eq $textLeft -and $editor.Parent.Right -le $item.Bounds.Left + $list.Columns[0].Width) 'The editor must align with the displayed label, including before optional icons finish loading, and stay inside its column.'
                Assert-UICheck ($editor.BackColor.ToArgb() -eq $list.BackColor.ToArgb() -and $editor.ForeColor.ToArgb() -eq $list.ForeColor.ToArgb() -and $editor.BorderStyle -eq 'None') 'Inline editing must use the list colors without a native white frame.'
                $editor.Text = 'Typing a new ' + $unicodeName + ' label'
                Assert-UICheck ($item.Text -ceq $aliases[$game] -and (Get-FilePreference $game).Label -ceq $aliases[$game]) 'Typing must not save before the edit is committed.'
                $method = [Windows.Forms.Control].GetMethod('OnKeyDown', [Reflection.BindingFlags]'Instance,NonPublic')
                $key = New-Object Windows.Forms.KeyEventArgs([Windows.Forms.Keys]::Escape)
                $args = New-Object object[] 1; $args[0] = $key.PSObject.BaseObject
                [void]$method.Invoke($editor, $args)
                Assert-UICheck (-not $list.IsShortcutNameEditing -and $form.Controls.Find('ShortcutNameEditor', $true).Count -eq 0 -and (Get-FilePreference $game).Label -ceq $aliases[$game]) 'Escape must remove the editor and preserve the saved label.'
                $list.BeginShortcutNameEdit($item)
                $editor = Get-UITestControl $form 'ShortcutNameEditor'; $editor.Text = 'Saved with Enter ' + $unicodeName
                $key = New-Object Windows.Forms.KeyEventArgs([Windows.Forms.Keys]::Enter); $args[0] = $key.PSObject.BaseObject
                [void]$method.Invoke($editor, $args)
                Assert-UICheck (-not $list.IsShortcutNameEditing -and $key.SuppressKeyPress -and (Get-FilePreference $game).Label -ceq ('Saved with Enter ' + $unicodeName)) 'Enter must save the shortcut label and close the editor without an extra keypress.'
                $state.Step = 2; return
            }
            if ($state.Step -eq 2) {
                Assert-UICheck ($item.Text -ceq ('Saved with Enter ' + $unicodeName)) 'After Enter, the row must display the new label once.'
                $list.BeginShortcutNameEdit($item)
                $editor = Get-UITestControl $form 'ShortcutNameEditor'; $editor.Text = 'Saved on focus change ' + $unicodeName
                $list.Columns[0].Width = 180
                Assert-UICheck ($editor.Parent.Right -le $item.Bounds.Left + $list.Columns[0].Width) 'Resizing the column must keep the active editor inside it.'
                (Get-UITestControl $form 'SearchBox').Focus() | Out-Null
                Assert-UICheck (-not $list.IsShortcutNameEditing -and (Get-FilePreference $game).Label -ceq ('Saved on focus change ' + $unicodeName)) 'Clicking outside must commit the edit and remove the overlay.'
                $state.Step = 3; return
            }
            Assert-UICheck ($item.Text -ceq ('Saved on focus change ' + $unicodeName) -and [Convert]::ToBase64String([IO.File]::ReadAllBytes($game)) -ceq $originalBytes[$game]) 'Inline renaming must update only the displayed shortcut label, leaving the executable intact.'
            $state.Finished = $true; $form.Close()
        }.GetNewClosure()
        $null = Invoke-UITestCase -Name 'inline shortcut editor' -Action $reopenAction -OnTick $editingTick
    }

    if ($Case -in @('All', 'Scanning')) {
        $nativeScanFactory = (Get-Command Start-FolderScan).ScriptBlock
        function New-UITestScan {
            $completion = New-Object 'Threading.Tasks.TaskCompletionSource[object]'
            $stub = [pscustomobject]@{
                Operation = $completion.Task; Completion = $completion; Cancelled = $false; Disposed = $false
                Progress = [pscustomobject]@{ Files = 0; Folders = 0; Runnables = 0; Skipped = 0; CurrentPath = $extractedDirectory }
                Partial = [pscustomobject]@{ RootFolder = $extractedDirectory; Runnables = @((Get-Item -LiteralPath $game)); AllArchives = @(); Media = @(); Skipped = @(); Complete = $false }
            }
            $stub | Add-Member -MemberType ScriptMethod -Name Cancel -Value { $this.Cancelled = $true }
            $stub | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $this.Cancelled = $true; $this.Disposed = $true }
            $stub | Add-Member -MemberType ScriptMethod -Name Snapshot -Value { param($complete) return $this.Partial }
            return $stub
        }
        Set-Item -Path Function:global:Start-FolderScan -Value { param($RootFolder, $ArchivePath, $ListToolPath) return $Global:UIHarnessState.Scan }
        try {
            $Global:UIHarnessState.Scan = New-UITestScan
            $scanTick = {
                param($state)
                $form = Get-UITestForm 'FolderScanForm'
                if (-not $form -or $form.Tag -ne 'Ready') { return }
                $form.Opacity = 0
                if ($state.Step -eq 0) {
                    Assert-UICheck (-not (Get-UITestControl $form 'UseFoundFiles').Enabled) 'Partial results must be unavailable before a runnable is found.'
                    $Global:UIHarnessState.Scan.Progress.Files = 12000; $Global:UIHarnessState.Scan.Progress.Folders = 82; $Global:UIHarnessState.Scan.Progress.Runnables = 1
                    $state.Step = 1; return
                }
                if (-not (Get-UITestControl $form 'ScanStatus').Text.Contains('12,000')) { return }
                Assert-UICheck ((Get-UITestControl $form 'ScanLocation').Text -ceq $extractedDirectory) 'Scan progress must show the actual folder being scanned.'
                Assert-UICheck ((Get-UITestControl $form 'UseFoundFiles').Enabled -and (Get-UITestControl $form 'CancelScan').Enabled) 'A running scan must allow cancellation and using found files.'
                $raw = $Global:UIHarnessState.Scan.Partial.PSObject.Copy(); $raw.Complete = $true
                $Global:UIHarnessState.Scan.Completion.SetResult($raw)
                $state.Finished = $true
            }.GetNewClosure()
            $scanAction = { Get-InventoryResponsive -Path $extractedDirectory }.GetNewClosure()
            $inventory = Invoke-UITestCase -Name 'scan progress and completion' -Action $scanAction -OnTick $scanTick
            Assert-UICheck ($inventory.Complete -and $inventory.Runnables.Count -eq 1 -and $Global:UIHarnessState.Scan.Disposed) 'Completed scans must return results and release their worker without waiting again.'

            $Global:UIHarnessState.Scan = New-UITestScan
            $cancelTick = {
                param($state)
                $form = Get-UITestForm 'FolderScanForm'
                if (-not $form -or $form.Tag -ne 'Ready') { return }
                $form.Opacity = 0; $state.Finished = $true; (Get-UITestControl $form 'CancelScan').PerformClick()
            }.GetNewClosure()
            $cancelAction = { try { Get-InventoryResponsive $extractedDirectory } catch [OperationCanceledException] { return 'cancelled' } }.GetNewClosure()
            $cancelled = Invoke-UITestCase -Name 'scan cancellation' -Action $cancelAction -OnTick $cancelTick
            Assert-UICheck ($cancelled -eq 'cancelled' -and $Global:UIHarnessState.Scan.Disposed) 'Cancel must return even while the scan task is still pending.'

            $Global:UIHarnessState.Scan = New-UITestScan; $Global:UIHarnessState.Scan.Progress.Runnables = 1
            $partialTick = {
                param($state)
                $form = Get-UITestForm 'FolderScanForm'
                if (-not $form -or $form.Tag -ne 'Ready' -or -not (Get-UITestControl $form 'UseFoundFiles').Enabled) { return }
                $form.Opacity = 0; $state.Finished = $true; (Get-UITestControl $form 'UseFoundFiles').PerformClick()
            }.GetNewClosure()
            $partial = Invoke-UITestCase -Name 'use found files' -Action $scanAction -OnTick $partialTick
            Assert-UICheck (-not $partial.Complete -and $partial.Runnables.Count -eq 1 -and $Global:UIHarnessState.Scan.Cancelled) 'Using found files must explicitly stop the remaining scan and return a marked partial result.'

            $Global:UIHarnessState.Scan = New-UITestScan
            $errorTick = {
                param($state)
                $form = Get-UITestForm 'FolderScanForm'
                if (-not $form -or $form.Tag -ne 'Ready') { return }
                $form.Opacity = 0; $Global:UIHarnessState.Scan.Completion.SetException([IO.IOException]::new('scan fixture failed')); $state.Finished = $true
            }.GetNewClosure()
            $errorAction = { try { Get-InventoryResponsive $extractedDirectory } catch { return $_.Exception.Message } }.GetNewClosure()
            $errorText = Invoke-UITestCase -Name 'scan failure' -Action $errorAction -OnTick $errorTick
            Assert-UICheck ($errorText.Contains('scan fixture failed') -and $Global:UIHarnessState.Scan.Disposed) 'A failed scan must close with an actionable error instead of waiting forever.'
        } finally { Set-Item -Path Function:global:Start-FolderScan -Value $nativeScanFactory }
    }

    if ($Case -in @('All', 'Cleanup')) {
        $cleanupTick = {
            param($state)
            $form = Get-UITestForm 'CleanupForm'
            if (-not $form -or $form.Tag -ne 'Ready') { return }
            $form.Opacity = 0
            $preview = Get-UITestControl $form 'CleanupPreview'
            $confirm = Get-UITestControl $form 'ConfirmCleanup'
            $ack = Get-UITestControl $form 'AcknowledgeAppDeletion'
            $source = Get-UITestControl $form 'RemoveSourceFiles'
            $extracted = Get-UITestControl $form 'RemoveExtractedFiles'
            $torrent = Get-UITestControl $form 'RemoveTorrentEntry'
            $permanent = Get-UITestControl $form 'PermanentlyDeleteFiles'
            if ($preview.Text -eq 'Checking selected locations...') { return }
            if ($state.Step -eq 0) {
                Assert-UICheck (-not $source.Checked -and -not $extracted.Checked -and -not $torrent.Checked -and -not $confirm.Enabled) 'Cleanup should start with an explicit empty selection.'
                Assert-UICheck ((Get-UITestControl $form 'RemoveSourceFilesLocation').Text -ceq $sourceDirectory) 'The source location should appear once as a full path.'
                Assert-UICheck ((Get-UITestControl $form 'RemoveExtractedFilesLocation').Text -ceq $extractedDirectory) 'The extracted location should appear as a full path.'
                $source.Checked = $true; $state.Step = 1; return
            }
            if ($state.Step -eq 1) {
                Assert-UICheck ($preview.Text.Contains($archive) -and $preview.Text.Contains('Keep extracted files')) 'Source-only cleanup must preview the installer and keep the extracted game.'
                Assert-UICheck ($confirm.Enabled -and -not $ack.Visible) 'Keeping extracted files should not request app-deletion acknowledgement.'
                Assert-UICheck ([IO.File]::Exists($archive) -and [IO.File]::Exists($game)) 'Preparing a cleanup preview must never delete files.'
                $extracted.Checked = $true
                $torrent.Checked = $true
                Assert-UICheck (-not $confirm.Enabled) 'Cleanup must wait for its replacement preview before continuing.'
                $state.Step = 2
                return
            }
            if ($state.Step -eq 2) {
                Assert-UICheck ($source.Checked -and $extracted.Checked -and $torrent.Checked -and $preview.Text.Contains('source and extracted files') -and $preview.Text.Contains('entry from qBittorrent')) 'All three cleanup items must be selectable and previewed together.'
                Assert-UICheck (-not $preview.Text.Contains('UI fixture torrent') -and -not $preview.Text.Contains($sourceDirectory)) 'Cleanup must not repeat the torrent name and paths already shown above.'
                Assert-UICheck ($ack.Visible -and -not $ack.Checked -and -not $confirm.Enabled) 'Deleting the extracted app must require explicit acknowledgement.'
                $ack.Checked = $true
                Assert-UICheck $confirm.Enabled 'Intentional app deletion should be available after acknowledgement.'
                $ack.Checked = $false
                Assert-UICheck (-not $confirm.Enabled) 'Removing acknowledgement must disable app deletion again.'
                Assert-UICheck ($permanent.Enabled -and -not $permanent.Checked) 'Combined cleanup must default to the Recycle Bin.'
                $ack.Checked = $true
                $state.Finished = $true
                $confirm.PerformClick()
            }
        }.GetNewClosure()
        $cleanupAction = { Show-CleanupConfirmForm -TorrentLabel 'UI fixture torrent' -TorrentHash ('a' * 40) -SourcePath $sourceDirectory -RootFolder $extractedDirectory -SelectedPath $game }.GetNewClosure()
        $cleanupResult = Invoke-UITestCase -Name 'cleanup scope and acknowledgement' -Action $cleanupAction -OnTick $cleanupTick
        Assert-UICheck ($cleanupResult.RemoveSource -and $cleanupResult.RemoveExtracted -and $cleanupResult.RemoveTorrent -and $cleanupResult.Plan.Mode -eq 'Everything' -and $cleanupResult.Acknowledged) 'The combined checklist selection was not returned as one action.'
        Assert-UICheck ([IO.File]::Exists($archive) -and [IO.File]::Exists($game)) 'Choosing a preview must not execute file deletion.'
        $torrentOnlyTick = {
            param($state)
            $form = Get-UITestForm 'CleanupForm'
            if (-not $form -or $form.Tag -ne 'Ready') { return }
            $form.Opacity = 0
            $torrent = Get-UITestControl $form 'RemoveTorrentEntry'
            $torrent.Checked = $true
            Assert-UICheck ((Get-UITestControl $form 'ConfirmCleanup').Enabled -and -not (Get-UITestControl $form 'AcknowledgeAppDeletion').Visible -and -not (Get-UITestControl $form 'PermanentlyDeleteFiles').Enabled) 'Removing just the entry must leave file deletion disabled.'
            $state.Finished = $true; (Get-UITestControl $form 'ConfirmCleanup').PerformClick()
        }.GetNewClosure()
        $torrentOnly = Invoke-UITestCase -Name 'torrent entry only' -Action $cleanupAction -OnTick $torrentOnlyTick
        Assert-UICheck ($torrentOnly.RemoveTorrent -and -not $torrentOnly.RemoveSource -and -not $torrentOnly.RemoveExtracted -and -not $torrentOnly.Plan) 'Torrent-only cleanup must not carry a file deletion plan.'
        $cancelTick = {
            param($state)
            $form = Get-UITestForm 'CleanupForm'
            if (-not $form -or $form.Tag -ne 'Ready') { return }
            $form.Opacity = 0
            Assert-UICheck (-not (Get-UITestControl $form 'RemoveTorrentEntry').Enabled) 'Manual folders must show that no torrent entry is available.'
            $state.Finished = $true; $form.Close()
        }.GetNewClosure()
        $cancelAction = { Show-CleanupConfirmForm -SourcePath $sourceDirectory -RootFolder $extractedDirectory }.GetNewClosure()
        $cancelResult = Invoke-UITestCase -Name 'manual cleanup cancellation' -Action $cancelAction -OnTick $cancelTick
        Assert-UICheck ($null -eq $cancelResult) 'Closing cleanup should cancel without a result.'
    }

    if ($Case -in @('All', 'Settings')) {
        $settingsTick = {
            param($state)
            $form = Get-UITestForm 'SettingsForm'
            if (-not $form) { return }
            $form.Opacity = 0
            Assert-UICheck ($form.AutoScaleMode -eq 'Dpi') 'The settings form must use DPI scaling.'
            $all = @($form.Controls | ForEach-Object { $_.Controls })
            $text = @($all | ForEach-Object { $_.Text }) -join ' '
            Assert-UICheck ($text -match 'update' -and $text -match 'admin') 'Settings should expose update and default admin preferences.'
            $state.Finished = $true
            $form.Close()
        }.GetNewClosure()
        $null = Invoke-UITestCase -Name 'settings' -Action { Show-SettingsForm } -OnTick $settingsTick
    }

    if ($Case -in @('All', 'Setup')) {
        $setupTick = {
            param($state)
            $form = Get-UITestForm 'DashboardForm'
            if (-not $form) { return }
            $form.Opacity = 0
            Assert-UICheck ($form.AutoScaleMode -eq 'Dpi') 'The setup form must use DPI scaling.'
            Assert-UICheck ($form.FormBorderStyle -eq 'Sizable') 'The setup form should adapt to the display.'
            $recent = Get-UITestControl $form 'RecentFolders'
            Assert-UICheck ($recent.Items.Count -gt 0) 'The dashboard should expose the remembered fixture folder.'
            $state.Finished = $true
            $form.Close()
        }.GetNewClosure()
        $null = Invoke-UITestCase -Name 'setup dashboard' -Action { Show-SetupMenu } -OnTick $setupTick
    }

    foreach ($file in $fixtures) { Assert-UICheck ([IO.File]::Exists($file) -and [Convert]::ToBase64String([IO.File]::ReadAllBytes($file)) -ceq $originalBytes[$file]) 'UI smoke checks changed an original target file.' }
    $result = [pscustomobject]@{ Passed = $true; Checks = $Global:UIHarnessState.Checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString(); Suite = 'UI' }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) }
    else { $result }
}
catch {
    if ($ResultPath) {
        [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $Global:UIHarnessState.Checks; Error = $_.Exception.Message; Suite = 'UI' } | ConvertTo-Json))
        exit 1
    }
    throw
}
finally {
    Close-UITestForms
    [Windows.Forms.Application]::remove_ThreadException($threadHandler)
    foreach ($name in $importedFunctions) {
        if ($originalFunctions[$name]) { Set-Item -Path ('Function:global:' + $name) -Value $originalFunctions[$name] }
        else { Remove-Item -LiteralPath ('Function:global:' + $name) -ErrorAction SilentlyContinue }
    }
    foreach ($name in $globals.Keys) {
        if ($globals[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $globals[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    Remove-Variable -Name UIHarnessState -Scope Global -ErrorAction SilentlyContinue
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    if ($resolvedRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedRoot).StartsWith('qbitlauncher-ui-tests-') -and [IO.Directory]::Exists($resolvedRoot)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
