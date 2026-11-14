# Execute the real startup routing against isolated fixtures with UI actions stubbed.
param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot '..\Source\Services.ps1')
. (Join-Path $PSScriptRoot '..\Source\Workflow.ps1')
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-session-flow-' + [Guid]::NewGuid().ToString('N'))
$script:checks = 0; $script:scans = New-Object System.Collections.ArrayList; $script:messages = New-Object System.Collections.ArrayList
$script:selected = $null; $script:extractions = 0; $script:cancelScan = $false; $script:confirmExtraction = $false
$script:setupCalls = 0
function Assert-SessionFlow { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message }; $script:checks++ }
function Show-SetupMenu { $script:setupCalls++; $Global:PendingSessionId = $script:sessionId; return [Windows.Forms.DialogResult]::OK }
function Show-ExecutableSelectionForm {
    param($FoundExecutables, $WindowTitle, $RootFolder, $TorrentHash, $TorrentName, $SessionId)
    $script:selected = @{ Root = $RootFolder; Hash = $TorrentHash; Name = $TorrentName; Id = $SessionId; Source = $filePathFromQB; Count = $FoundExecutables.Count }
}
function Get-AllRunnables {
    param($RootFolderPath, $VerifyArchivePath)
    [void]$script:scans.Add(@{ Root = $RootFolderPath; Archive = $VerifyArchivePath })
    if ($script:cancelScan) { throw [OperationCanceledException]::new('fixture scan cancelled') }
    return Get-Item -LiteralPath (Join-Path $RootFolderPath 'Game.exe')
}
function Show-ExtractionConfirmForm {
    param($ArchivePath, $DefaultDestination)
    $script:defaultDestination = $DefaultDestination
    return @{ Confirmed = $script:confirmExtraction; UseExisting = -not $script:confirmExtraction; SkippedExtraction = $false; DestinationPath = $script:chosenRoot }
}
function Expand-ArchiveFile { param($ArchivePath, $DestinationPath); $script:extractions++; return $DestinationPath }
function Show-ThemedMessageBox { param($Message, $Title, $Icon); [void]$script:messages.Add(@{ Text = $Message; Title = $Title }) }
function Start-Process { throw 'Session reopening must not launch Explorer or an extractor.' }
function Select-ExtractionPath { throw 'A saved session must not open a folder picker.' }
function New-SessionForm {
    $control = [pscustomobject]@{ DialogResult = [Windows.Forms.DialogResult]::None; Closed = $false }
    $control | Add-Member -MemberType ScriptMethod -Name Close -Value { $this.Closed = $true }
    return $control
}
try {
    [void][IO.Directory]::CreateDirectory($testRoot)
    $Global:ConfigFile = Join-Path $testRoot 'config.json'; $Global:DataDirectory = Join-Path $testRoot 'appdata'
    $Global:PendingUpdateCheck = $null; $Global:WorkflowLock = $null
    $ArchiveExtensions = @('zip', 'rar', '7z', 'iso', 'img')
    Get-UserSettings; [void](Save-UserSettings)
    $source = Join-Path $testRoot 'Original [download].rar'; [IO.File]::WriteAllText($source, 'fixture archive')
    $root = Join-Path $testRoot 'Specific extracted game'; [void][IO.Directory]::CreateDirectory($root)
    $game = Join-Path $root 'Game.exe'; [IO.File]::WriteAllText($game, 'fixture runnable')
    $entry = Save-LauncherSession -SourcePath $source -RootFolder $root -ArchivePath $source -TorrentHash ('a' * 40) -TorrentName 'Fixture game'
    $script:sessionId = $entry.Id; $script:chosenRoot = $root
    [void](Set-FilePreference -Path $game -Preference @{ Label = 'My saved shortcut'; Arguments = '--saved' })
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Source\App.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors.Message -join "`n") }
    $main = @($ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] -and $_.Extent.Text.Contains('$jobKey') })[0]
    if (-not $main) { throw 'The application startup workflow was not found.' }
    $workflow = [scriptblock]::Create($main.Extent.Text)

    $SavedSessionId = $entry.Id
    $filePathFromQB = ''; $torrentHashFromQB = ''; $torrentNameFromQB = ''
    & $workflow
    Assert-SessionFlow ($script:setupCalls -eq 0 -and $script:selected.Id -eq $entry.Id -and $script:selected.Hash -eq ('a' * 40)) 'A child launched with a saved session ID must open that session directly without a dashboard or extraction prompt.'
    Assert-SessionFlow (@($ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'SavedSessionId' }).Count -eq 1) 'The standalone entry point must accept the saved session parameter used by separate windows.'
    $SavedSessionId = $null; $script:scans.Clear()
    $filePathFromQB = ''; $torrentHashFromQB = ''; $torrentNameFromQB = ''
    & $workflow
    Assert-SessionFlow ($script:selected.Root -ceq $root -and $script:selected.Source -ceq $source -and $script:selected.Hash -eq ('a' * 40) -and $script:selected.Id -eq $entry.Id) 'Reopening must restore the exact source, extracted folder, torrent hash, and session ID.'
    Assert-SessionFlow ($script:scans.Count -eq 1 -and $script:scans[0].Root -ceq $root -and $script:scans[0].Archive -ceq $source -and $script:extractions -eq 0) 'Reopening must scan only the saved folder and verify existing output without re-extracting.'
    Assert-SessionFlow ((Get-FilePreference $game).Label -ceq 'My saved shortcut') 'Reopening erased the saved shortcut name.'
    Assert-SessionFlow $Global:WorkflowLock.IsDisposed 'Reopening left a source lock behind after closing.'

    [IO.File]::Delete($source); $script:scans.Clear(); $script:selected = $null
    & $workflow
    Assert-SessionFlow ($script:selected.Root -ceq $root -and $script:scans.Count -eq 1 -and -not $script:scans[0].Archive) 'An absent source archive must not prevent reopening the extracted game or trigger a parent-folder scan.'
    Set-LauncherSessionTorrentRemoved -Id $entry.Id
    & $workflow
    Assert-SessionFlow (-not $script:selected.Hash -and $script:selected.Source -ceq $source) 'An already removed torrent entry was restored as an available cleanup target.'

    [IO.File]::Delete($game); [IO.Directory]::Delete($root); $script:scans.Clear(); $script:selected = $null; $script:messages.Clear()
    & $workflow
    Assert-SessionFlow ($script:scans.Count -eq 0 -and -not $script:selected -and $script:messages[0].Text -match 'moved or was deleted') 'A missing extraction destination must report its location without scanning the parent or reopening the archive.'
    [void][IO.Directory]::CreateDirectory($root); [IO.File]::WriteAllText($game, 'fixture runnable')
    [IO.File]::WriteAllText($source, 'fixture archive')
    $held = New-WorkflowLock -Key ('path:' + (Get-CanonicalPath $source).ToLowerInvariant())
    try {
        $script:scans.Clear(); $script:messages.Clear()
        & $workflow
        Assert-SessionFlow ($script:scans.Count -eq 0 -and $script:messages[0].Title -eq 'Already open') 'Reopening must honor an existing workflow for the same source.'
    } finally { Exit-WorkflowLock $held }

    # The archive dialog can reuse a known extraction without another extraction.
    $entry = Save-LauncherSession -SourcePath $source -RootFolder $root -ArchivePath $source -TorrentHash ('b' * 40) -TorrentName 'Fresh torrent'
    $filePathFromQB = $source; $torrentHashFromQB = ''; $torrentNameFromQB = ''; $script:scans.Clear()
    & $workflow
    Assert-SessionFlow ($script:defaultDestination -ceq $root -and $script:selected.Hash -eq ('b' * 40) -and $script:extractions -eq 0) ("Opening a known archive must restore its destination/context. Destination=$script:defaultDestination; hash=$($script:selected.Hash); extractions=$script:extractions; messages=" + ($script:messages | ConvertTo-Json -Compress))

    # Save the destination before scanning so closing/cancelling that stage is recoverable.
    $other = Join-Path $testRoot 'Next extraction'; [void][IO.Directory]::CreateDirectory($other); [IO.File]::WriteAllText((Join-Path $other 'Game.exe'), 'fixture')
    $script:chosenRoot = $other; $script:confirmExtraction = $true; $script:cancelScan = $true
    $filePathFromQB = $source; $torrentHashFromQB = ('c' * 40); $torrentNameFromQB = 'Next game'
    & $workflow
    Get-UserSettings
    Assert-SessionFlow ((Get-LauncherSession -SourcePath $source -RootFolder $other).TorrentHash -eq ('c' * 40)) 'A cancelled scan after extraction must still leave a reopenable session.'

    # Exercise dashboard and extraction buttons as callbacks without displaying controls.
    $uiAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Source\UI.ps1'), [ref]$tokens, [ref]$errors)
    $dashboard = $uiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Show-SetupMenu' }, $false)
    $click = $dashboard.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression -is [Management.Automation.Language.VariableExpressionAst] -and $node.Expression.VariablePath.UserPath -eq 'reopen' -and $node.Member.Value -eq 'Add_Click' }, $true)
    $sessionState = @{ Items = @($entry) }; $sessions = [pscustomobject]@{ SelectedIndex = 0 }; $form = New-SessionForm
    $Global:PendingSessionId = $null
    $reopenAction = Invoke-Expression $click.Arguments[0].Extent.Text
    & $reopenAction
    Assert-SessionFlow ($Global:PendingSessionId -eq $entry.Id -and $Global:PendingManualFolder -ceq $root -and $form.Closed) 'Reopen session must pass the selected identity to startup and close the dashboard.'

    $dialog = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Show-ExtractionConfirmForm' }, $false)
    $click = $dialog.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression -is [Management.Automation.Language.VariableExpressionAst] -and $node.Expression.VariablePath.UserPath -eq 'openExistingButton' -and $node.Member.Value -eq 'Add_Click' }, $true)
    $destTextBox = [pscustomobject]@{ Text = $root }; $form = New-SessionForm
    $existingAction = Invoke-Expression $click.Arguments[0].Extent.Text
    & $existingAction
    Assert-SessionFlow ($form.Closed -and $form.DialogResult -eq [Windows.Forms.DialogResult]::Retry) 'Open extracted files must choose the existing destination without invoking extraction.'
    $destTextBox.Text = Join-Path $testRoot 'missing'; $form = New-SessionForm; $script:messages.Clear()
    $existingAction = Invoke-Expression $click.Arguments[0].Extent.Text
    & $existingAction
    Assert-SessionFlow (-not $form.Closed -and $script:messages[0].Title -eq 'Extracted folder missing') 'An invalid extracted folder must keep the dialog open instead of scanning another location.'

    $script:cancelScan = $false; $script:scans.Clear()
    $single = Save-LauncherSession -SourcePath $game -RootFolder $root
    $script:sessionId = $single.Id; $filePathFromQB = ''; $torrentHashFromQB = ''; $torrentNameFromQB = ''
    & $workflow
    Assert-SessionFlow ($script:selected.Count -eq 1 -and $script:scans.Count -eq 0 -and $script:selected.Source -ceq $game) 'A session started from one runnable must reopen only that file, without scanning its parent folder.'
    [IO.File]::Delete($game); $script:messages.Clear(); $script:selected = $null
    & $workflow
    Assert-SessionFlow (-not $script:selected -and $script:scans.Count -eq 0 -and $script:messages[0].Text -match 'saved application has moved or was deleted') 'A missing single-file session must not fall back to scanning the parent folder.'

    $result = @{ Passed = $true; Checks = $script:checks; Suite = 'SessionWorkflow'; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { [pscustomobject]$result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'SessionWorkflow' } | ConvertTo-Json)); exit 1 }
    throw
}
finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('qbitlauncher-session-flow-') -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
