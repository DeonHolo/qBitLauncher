# Window launch arguments use hidden, file-only probes. No app windows are opened.
param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
. (Join-Path $PSScriptRoot '..\Source\Services.ps1')
. (Join-Path $PSScriptRoot '..\Source\Workflow.ps1')
. (Join-Path $PSScriptRoot '..\Source\Updates.ps1')
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-session-windows-' + [Guid]::NewGuid().ToString('N'))
$script:checks = 0; $moreMenu = $null; $savedSessionsMenu = $null
function Assert-Window { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message }; $script:checks++ }
function Wait-ProbeFile {
    param([string]$Path)
    $deadline = [DateTime]::UtcNow.AddSeconds(8)
    while (-not [IO.File]::Exists($Path) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 20 }
    if (-not [IO.File]::Exists($Path)) { throw 'The hidden launch-argument probe did not finish.' }
}
try {
    [void][IO.Directory]::CreateDirectory($testRoot)
    $Global:ConfigFile = Join-Path $testRoot 'config.json'; $Global:DataDirectory = Join-Path $testRoot 'appdata'
    $Global:CurrentTheme = @{ TextFore = [Drawing.Color]::White }
    Get-UserSettings
    $firstRoot = Join-Path $testRoot 'First game'; $otherRoot = Join-Path $testRoot 'Other game'
    [void][IO.Directory]::CreateDirectory($firstRoot); [void][IO.Directory]::CreateDirectory($otherRoot)
    $first = Save-LauncherSession -SourcePath $firstRoot -RootFolder $firstRoot -TorrentHash ('a' * 40) -TorrentName 'First game'
    $other = Save-LauncherSession -SourcePath $otherRoot -RootFolder $otherRoot -TorrentHash ('b' * 40) -TorrentName 'Other game'
    $probeDirectory = Join-Path $testRoot ("Install's [test] & " + [char]0x6e38)
    [void][IO.Directory]::CreateDirectory($probeDirectory)
    $probe = Join-Path $probeDirectory 'Launcher probe.exe'
    $compiledProbe = Join-Path $testRoot 'probe.exe'
    Add-Type -TypeDefinition @'
public static class SessionLaunchProbe {
    public static void Main(string[] args) {
        string file = System.IO.Path.ChangeExtension(System.Reflection.Assembly.GetExecutingAssembly().Location, ".args");
        System.IO.File.WriteAllLines(file, args);
    }
}
'@ -OutputAssembly $compiledProbe -OutputType ConsoleApplication
    [IO.File]::Copy($compiledProbe, $probe)
    $Global:IsCompiledExe = $true; $Global:ExecutablePath = $probe
    $info = New-LauncherSessionStartInfo -SessionId $other.Id
    Assert-Window ($info.FileName -ceq $probe -and $info.WorkingDirectory -ceq $probeDirectory -and -not $info.UseShellExecute -and $info.CreateNoWindow) 'A compiled session should start the same launcher from its own directory without a console.'
    Start-LauncherSessionWindow -SessionId $other.Id
    $report = [IO.Path]::ChangeExtension($probe, '.args'); Wait-ProbeFile $report
    $actual = [IO.File]::ReadAllLines($report)
    Assert-Window ($actual.Count -eq 2 -and $actual[0] -ceq '-SavedSessionId' -and $actual[1] -ceq $other.Id) 'The compiled launch must pass the selected saved session directly, without reopening a source archive.'

    $scriptProbe = Join-Path $probeDirectory 'Launcher probe.ps1'
    [IO.File]::WriteAllText($scriptProbe, 'param([string]$SavedSessionId) [IO.File]::WriteAllText((Join-Path $PSScriptRoot ''script.args''), $SavedSessionId)', (New-Object Text.UTF8Encoding($true)))
    $Global:IsCompiledExe = $false; $Global:ExecutablePath = $scriptProbe
    $info = New-LauncherSessionStartInfo -SessionId $other.Id
    Assert-Window ($info.FileName -match 'WindowsPowerShell.+powershell\.exe$' -and $info.WindowStyle -eq [Diagnostics.ProcessWindowStyle]::Hidden -and $info.Arguments -match '-STA') 'Script sessions must run in a separate STA PowerShell process with its console hidden.'
    Start-LauncherSessionWindow -SessionId $other.Id
    $report = Join-Path $probeDirectory 'script.args'; Wait-ProbeFile $report
    Assert-Window ([IO.File]::ReadAllText($report) -ceq $other.Id) 'A script launcher path with spaces, brackets, apostrophes, and Unicode was not passed literally.'
    $rejected = $false
    try { $null = New-LauncherSessionStartInfo -SessionId 'unknown' } catch { $rejected = $true }
    Assert-Window $rejected 'An unknown saved session must fail before starting another process.'

    # Populate and click the real More submenu in memory, without showing it.
    $script:opened = New-Object System.Collections.ArrayList; $script:logs = New-Object System.Collections.ArrayList; $script:errors = New-Object System.Collections.ArrayList
    function Start-LauncherSessionWindow { param($SessionId); [void]$script:opened.Add($SessionId) }
    function Show-ThemedMessageBox { param($Message, $Title, $Icon); [void]$script:errors.Add($Message) }
    $Global:MainForm = [pscustomobject]@{ Name = 'Original session'; Closed = $false }
    $originalForm = $Global:MainForm
    $Global:PendingSessionId = 'unchanged'; $Global:PendingManualFolder = $firstRoot
    $filePathFromQB = $firstRoot; $SessionId = $first.Id
    $addLog = { param($message); [void]$script:logs.Add($message) }
    $savedSessionsMenu = New-Object Windows.Forms.ToolStripMenuItem
    $moreMenu = New-Object Windows.Forms.ContextMenuStrip
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Source\UI.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors.Message -join "`n") }
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Show-ExecutableSelectionForm' }, $false)
    $opening = $definition.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression -is [Management.Automation.Language.VariableExpressionAst] -and $node.Expression.VariablePath.UserPath -eq 'savedSessionsMenu' -and $node.Member.Value -eq 'Add_DropDownOpening' }, $true)
    $refresh = Invoke-Expression $opening.Arguments[0].Extent.Text
    & $refresh
    Assert-Window ($savedSessionsMenu.DropDownItems.Count -eq 1 -and $savedSessionsMenu.DropDownItems[0].Tag -ceq $other.Id -and $savedSessionsMenu.DropDownItems[0].ToolTipText -ceq $otherRoot) 'More must list other saved sessions with their destination, excluding the current session.'
    $savedSessionsMenu.DropDownItems[0].PerformClick()
    Assert-Window ($script:opened.Count -eq 1 -and $script:opened[0] -ceq $other.Id -and $script:errors.Count -eq 0) 'Choosing a saved game must start its separate launcher window.'
    Assert-Window ([object]::ReferenceEquals($Global:MainForm, $originalForm) -and -not $originalForm.Closed -and $filePathFromQB -ceq $firstRoot -and $Global:PendingSessionId -ceq 'unchanged' -and $Global:PendingManualFolder -ceq $firstRoot) 'Opening another session must not replace the current form, source, or pending workflow.'
    $null = Remove-LauncherSession -Id $other.Id
    & $refresh
    Assert-Window ($savedSessionsMenu.DropDownItems.Count -eq 1 -and -not $savedSessionsMenu.DropDownItems[0].Enabled) 'A refreshed menu must show an empty state rather than a forgotten session.'

    $result = @{ Passed = $true; Checks = $script:checks; Suite = 'SessionWindows'; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { [pscustomobject]$result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'SessionWindows' } | ConvertTo-Json)); exit 1 }
    throw
}
finally {
    if ($savedSessionsMenu) { $savedSessionsMenu.Dispose() }; if ($moreMenu) { $moreMenu.Dispose() }
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('qbitlauncher-session-windows-') -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
