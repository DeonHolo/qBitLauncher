# Read-only connection and credential-recovery callbacks, without opening windows.
param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Source\UI.ps1')
$script:checks = 0; $script:saves = 0; $script:finished = $null
function Assert-Settings { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message }; $script:checks++ }
function New-SettingsControl {
    param([string]$Text = '')
    $control = [pscustomobject]@{ Text = $Text; Enabled = $true; Checked = $false; SelectedItem = 'Dracula'; Closed = $false; Callback = $null }
    $control | Add-Member -MemberType ScriptMethod -Name PerformClick -Value { if ($this.Enabled) { & $this.Callback } }
    $control | Add-Member -MemberType ScriptMethod -Name Close -Value { $this.Closed = $true }
    $control | Add-Member -MemberType ScriptMethod -Name Start -Value { }
    $control | Add-Member -MemberType ScriptMethod -Name Stop -Value { }
    return $control
}
try {
    # A real runspace must report the underlying error, not EndInvoke internals.
    $task = Start-AppWorker -Action { throw 'qBittorrent login failed. Fixture credentials rejected.' }
    $deadline = [DateTime]::UtcNow.AddSeconds(5)
    while (-not $task.Handle.IsCompleted -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 10 }
    $received = Receive-AppWorker $task
    Assert-Settings ($received.Error -ceq 'qBittorrent login failed. Fixture credentials rejected.' -and $task.Disposed) 'Worker failures must show the original error and release the worker.'
    Assert-Settings ([object]::ReferenceEquals($received, (Receive-AppWorker $task))) 'Reading a completed worker twice should retain its result.'

    function Start-AppWorker {
        param($Action, $Arguments, $Functions, $Variables)
        $script:requested = $Arguments[0]
        if ($Functions -contains 'Remove-QBittorrentTorrent' -or $Action.ToString() -match 'Remove-QBittorrentTorrent') { throw 'Settings must never remove a torrent.' }
        return @{ Fixture = $true }
    }
    function Receive-AppWorker { param($Task); return $script:finished }
    function Save-UserSettings { $script:saves++; return $true }
    function Show-ThemedMessageBox { throw 'Unexpected message in a settings callback.' }

    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Source\UI.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors.Message -join "`n") }
    $settings = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Show-SettingsForm' }, $false)
    $url = New-SettingsControl 'http://127.0.0.1:8080'
    $username = New-SettingsControl 'fixture-user'; $password = New-SettingsControl 'old fixture password'
    $save = New-SettingsControl; $test = New-SettingsControl; $timer = New-SettingsControl; $form = New-SettingsControl
    $connectionStatus = New-SettingsControl; $info = New-SettingsControl; $destination = New-SettingsControl
    $theme = New-SettingsControl; $defaultAdmin = New-SettingsControl; $autoUpdate = New-SettingsControl
    $ConnectionError = 'qBittorrent login failed.'
    $Global:UserSettings = @{}; $Global:Themes = @{ Dracula = @{} }
    $state = @{ Saved = $false; Check = $null; ApiTask = $null; ApiCandidate = $null; Verified = $null; SaveAfterTest = $false }
    foreach ($name in @('getApiSettings', 'testConnection')) {
        $assignment = $settings.Find({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left -is [Management.Automation.Language.VariableExpressionAst] -and $node.Left.VariablePath.UserPath -eq $name }, $true)
        Invoke-Expression $assignment.Extent.Text
    }
    $saveEvent = $settings.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression -is [Management.Automation.Language.VariableExpressionAst] -and $node.Expression.VariablePath.UserPath -eq 'save' -and $node.Member.Value -eq 'Add_Click' }, $true)
    $save.Callback = Invoke-Expression $saveEvent.Arguments[0].Extent.Text
    $timerEvent = $settings.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression -is [Management.Automation.Language.VariableExpressionAst] -and $node.Expression.VariablePath.UserPath -eq 'timer' -and $node.Member.Value -eq 'Add_Tick' }, $true)
    $poll = Invoke-Expression $timerEvent.Arguments[0].Extent.Text

    $save.PerformClick()
    Assert-Settings ($state.ApiTask -and -not $save.Enabled -and -not $test.Enabled -and -not $password.Enabled -and $script:saves -eq 0) 'Recovery must test the entered credentials before saving or retrying cleanup.'
    & $poll
    Assert-Settings ($state.ApiTask -and $script:saves -eq 0) 'An unfinished connection check must keep the save pending.'
    $script:finished = @{ Values = @(); Error = 'qBittorrent login failed. Credentials rejected.' }
    & $poll
    Assert-Settings (-not $state.Saved -and -not $form.Closed -and $script:saves -eq 0 -and $info.Text -match 'login failed' -and $password.Enabled) 'Rejected credentials must remain editable without saving or continuing cleanup.'

    $script:finished = $null
    $password.Text = '  fixture & + = % ' + [char]0x6e38 + '  '
    $save.PerformClick()
    Assert-Settings ($script:requested.QbittorrentPassword -ceq $password.Text) 'Connection tests must preserve password whitespace, symbols, and non-English characters.'
    $script:finished = @{ Values = @('v5.2.4'); Error = $null }
    & $poll
    Assert-Settings ($state.Saved -and $form.Closed -and $script:saves -eq 1 -and $Global:UserSettings.QbittorrentPassword -ceq $password.Text) 'A verified correction must save the exact password and resume the pending cleanup.'

    $state.Saved = $false; $form.Closed = $false; $script:finished = $null
    & $testConnection $false
    $script:finished = @{ Values = @('v5.2.4'); Error = $null }
    & $poll
    Assert-Settings ($connectionStatus.Text -match 'Connected' -and -not $state.Saved -and -not $form.Closed -and $script:saves -eq 1) 'Test connection must report success without saving or closing Settings.'
    $password.Text = 'edited after testing'; $script:finished = $null
    $save.PerformClick()
    Assert-Settings ($state.ApiTask -and $script:saves -eq 1) 'Editing a verified password must require a new check before cleanup resumes.'

    $result = @{ Passed = $true; Checks = $script:checks; Suite = 'ApiSettings'; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { [pscustomobject]$result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'ApiSettings' } | ConvertTo-Json)); exit 1 }
    throw
}
