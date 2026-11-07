# Isolated updater checks. No request to GitHub and no live application replacement.
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'Source\Updates.ps1')
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-update-tests-' + [Guid]::NewGuid().ToString('N'))
$script:checks = 0
$script:savedSettings = 0
$previousSkipCheck = $env:QBITLAUNCHER_SKIP_UPDATE_CHECK
$env:QBITLAUNCHER_SKIP_UPDATE_CHECK = '0'
function Write-LogMessage { param([string]$Message) }
function Save-UserSettings { $script:savedSettings++; return $true }
function Assert-Update { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message }; $script:checks++ }
function Assert-UpdateThrows {
    param([scriptblock]$Action, [string]$Message)
    $threw = $false
    try { $null = & $Action } catch { $threw = $true }
    Assert-Update $threw $Message
}

try {
    [IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $Global:IsCompiledExe = $true
    $Global:ScriptVersion = '3.3.7'
    $Global:UserSettings = @{ CheckForUpdates = $true; LastUpdateCheckUtc = '' }
    $hash = ('a' * 64)
    $release = [pscustomobject]@{
        tag_name = 'v3.10.2'; draft = $false; prerelease = $false; html_url = 'https://github.com/DeonHolo/qBitLauncher/releases/tag/v3.10.2'
        assets = @(
            [pscustomobject]@{ name = 'qBitLauncher.exe'; digest = ('sha256:' + $hash); browser_download_url = 'https://github.com/DeonHolo/qBitLauncher/releases/download/v3.10.2/qBitLauncher.exe' },
            [pscustomobject]@{ name = 'qBitLauncher.ps1'; browser_download_url = 'https://github.com/DeonHolo/qBitLauncher/releases/download/v3.10.2/qBitLauncher.ps1' },
            [pscustomobject]@{ name = 'SHA256SUMS.txt'; browser_download_url = 'https://github.com/DeonHolo/qBitLauncher/releases/download/v3.10.2/SHA256SUMS.txt' }
        )
    }
    $candidate = Get-ReleaseUpdateCandidate -Release $release -IsCompiled $true
    Assert-Update ($candidate.Version -eq '3.10.2') 'Release tag version was not parsed.'
    Assert-Update ($candidate.AssetName -eq 'qBitLauncher.exe') 'Compiled updater selected the wrong asset.'
    Assert-Update ($candidate.DownloadUrl -like '*/v3.10.2/qBitLauncher.exe') 'Version and asset did not come from the same release.'
    Assert-Update ($candidate.SHA256 -ceq $hash) 'Asset digest was not retained.'
    $sourceCandidate = Get-ReleaseUpdateCandidate -Release $release -IsCompiled $false
    Assert-Update ($sourceCandidate.AssetName -eq 'qBitLauncher.ps1') 'Source updater selected an EXE.'
    $script:checksumsFixture = ($hash + '  qBitLauncher.ps1' + "`n" + ('b' * 64) + '  qBitLauncher.exe')
    function Invoke-WebRequest { param($Uri, [switch]$UseBasicParsing, $Headers, $TimeoutSec, $ErrorAction); return [pscustomobject]@{ Content = $script:checksumsFixture } }
    Assert-Update ((Get-UpdateExpectedSHA256 -Candidate $sourceCandidate) -ceq $hash) 'Published checksum list was not parsed.'
    $script:checksumsFixture = ($hash + '  other-file.ps1')
    Assert-UpdateThrows { Get-UpdateExpectedSHA256 -Candidate $sourceCandidate } 'Missing checksum entry was accepted.'
    $withoutDigest = Get-ReleaseUpdateCandidate -Release $release -IsCompiled $false
    $withoutDigest.ChecksumsUrl = $null
    Assert-UpdateThrows { Get-UpdateExpectedSHA256 -Candidate $withoutDigest } 'An unverifiable update was accepted.'
    $noExe = [pscustomobject]@{ tag_name = 'v99.0.0'; assets = @(); draft = $false; prerelease = $false }
    Assert-Update ($null -eq (Get-ReleaseUpdateCandidate -Release $noExe -IsCompiled $true)) 'A release without an EXE was offered to compiled users.'
    $release.draft = $true
    Assert-Update ($null -eq (Get-ReleaseUpdateCandidate -Release $release)) 'A draft release was offered.'
    $release.draft = $false
    $release.prerelease = $true
    Assert-Update ($null -eq (Get-ReleaseUpdateCandidate -Release $release)) 'A prerelease was offered as stable.'
    $release.prerelease = $false
    $release.assets[0].browser_download_url = 'https://example.com/other.exe'
    Assert-UpdateThrows { Get-ReleaseUpdateCandidate -Release $release } 'An unrelated download URL was accepted.'
    $release.assets[0].browser_download_url = 'https://github.com/DeonHolo/qBitLauncher/releases/download/v3.10.2/qBitLauncher.exe'

    # Native Windows parsing checks the quoting independently of the formatter.
    if (-not ('QBitLauncher.UpdateArgvReader' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace QBitLauncher {
    public static class UpdateArgvReader {
        [DllImport("shell32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        private static extern IntPtr CommandLineToArgvW(string line, out int count);
        [DllImport("kernel32.dll")] private static extern IntPtr LocalFree(IntPtr memory);
        public static string[] Read(string line) {
            int count; IntPtr block=CommandLineToArgvW(line,out count);
            if(block==IntPtr.Zero) throw new InvalidOperationException("Could not parse arguments.");
            try { string[] result=new string[count]; for(int i=0;i<count;i++) result[i]=Marshal.PtrToStringUni(Marshal.ReadIntPtr(block,i*IntPtr.Size)); return result; }
            finally { LocalFree(block); }
        }
    }
}
'@ -ErrorAction Stop
    }
    $unicode = -join [char[]]@(0x6e38, 0x620f, 0x0418, 0x0433, 0x0440, 0x0430)
    $arguments = @('', 'plain', 'has spaces', '100% & parentheses()', "apostrophe's", 'quote"inside', 'two\\slashes', 'trailing\', 'space trailing\', 'slashes\\"quote', $unicode, ('emoji ' + [char]::ConvertFromUtf32(0x1f3ae)))
    $argumentString = ConvertTo-WindowsArgumentString -Arguments $arguments
    $parsed = [QBitLauncher.UpdateArgvReader]::Read('"fixture.exe" ' + $argumentString)
    Assert-Update ($parsed.Count -eq ($arguments.Count + 1)) 'Quoting changed the argument count.'
    for ($i = 0; $i -lt $arguments.Count; $i++) { Assert-Update ($parsed[$i + 1] -ceq $arguments[$i]) ('Argument changed at index ' + $i) }

    $unicodeDir = Join-Path $testRoot ($unicode + " 100% & apostrophe's")
    [IO.Directory]::CreateDirectory($unicodeDir) | Out-Null
    $target = Join-Path $unicodeDir 'application.ps1'
    $stage = Join-Path $unicodeDir 'new.ps1'
    $backup = Join-Path $unicodeDir 'previous.bak'
    [IO.File]::WriteAllText($target, 'old application')
    [IO.File]::WriteAllText($stage, ('new application ' + $unicode), [Text.Encoding]::UTF8)
    $expected = Get-UpdateFileSHA256 -Path $stage
    Assert-Update (Test-UpdateAssetHash -Path $stage -ExpectedSHA256 $expected) 'Known file hash did not verify.'
    Assert-UpdateThrows { Test-UpdateAssetHash -Path $stage -ExpectedSHA256 'bad' } 'Malformed hash was accepted.'
    $payload = @{ TargetPath = $target; StagedPath = $stage; BackupPath = $backup; SHA256 = ('0' * 64); ParentProcessId = 0; IsCompiled = $false; Arguments = $arguments; SkipRestart = $true }
    Assert-UpdateThrows { Invoke-UpdateReplacement -Payload $payload } 'A corrupt update was accepted.'
    Assert-Update ([IO.File]::ReadAllText($target) -ceq 'old application') 'Checksum failure changed the current application.'
    Assert-Update ([IO.File]::Exists($stage)) 'Checksum failure consumed the staged file.'
    $payload.SHA256 = $expected
    $handoff = New-UpdateHandoff -Payload $payload -Directory (Join-Path $unicodeDir 'handoff')
    $tokens = $null; $parseErrors = $null
    $null = [Management.Automation.Language.Parser]::ParseFile($handoff.ScriptPath, [ref]$tokens, [ref]$parseErrors)
    Assert-Update ($parseErrors.Count -eq 0) 'The generated handoff script does not parse.'
    $payloadReadback = [IO.File]::ReadAllText($handoff.PayloadPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Update ($payloadReadback.TargetPath -ceq $target) 'Unicode target changed in the handoff JSON.'
    for ($i = 0; $i -lt $arguments.Count; $i++) { Assert-Update ($payloadReadback.Arguments[$i] -ceq $arguments[$i]) ('Handoff argument changed at index ' + $i) }
    $decodedCommand = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($handoff.EncodedCommand))
    Assert-Update ($decodedCommand.Contains($unicode)) 'Encoded command lost Unicode path characters.'
    Assert-Update ($decodedCommand -notmatch 'cmd\.exe|\.bat') 'Handoff invokes a batch command processor.'

    # Execute the actual encoded-command helper against disposable fixture files.
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $start.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ' + $handoff.EncodedCommand
    $start.UseShellExecute = $false; $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $process = [Diagnostics.Process]::Start($start)
    try {
        if (-not $process.WaitForExit(15000)) { $process.Kill(); throw 'The fixture update handoff timed out.' }
        Assert-Update ($process.ExitCode -eq 0) 'The fixture update handoff failed.'
    }
    finally { $process.Dispose() }
    $handoffResult = [IO.File]::ReadAllText($handoff.PayloadPath + '.result.json', [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-Update ([bool]$handoffResult.Updated) 'The helper did not report successful replacement.'
    Assert-Update ([IO.File]::ReadAllText($target) -ceq ('new application ' + $unicode)) 'New Unicode content was not installed.'
    Assert-Update ([IO.File]::ReadAllText($backup) -ceq 'old application') 'The old application backup was not preserved.'
    Assert-Update (-not [IO.File]::Exists($stage)) 'Atomic replacement did not consume the staged file.'

    # Invalid replacement executable fails to start and must restore old bytes.
    $badTarget = Join-Path $unicodeDir 'rollback.exe'
    $badStage = Join-Path $unicodeDir 'invalid.exe'
    $badBackup = Join-Path $unicodeDir 'rollback.bak'
    [IO.File]::WriteAllText($badTarget, 'old valid fixture')
    [IO.File]::WriteAllText($badStage, 'not an executable')
    $rollbackPayload = @{ TargetPath = $badTarget; StagedPath = $badStage; BackupPath = $badBackup; SHA256 = (Get-UpdateFileSHA256 $badStage); ParentProcessId = 0; IsCompiled = $true; Arguments = @(); SkipRestart = $false }
    Assert-UpdateThrows { Invoke-UpdateReplacement -Payload $rollbackPayload } 'Failed restart did not report an error.'
    Assert-Update ([IO.File]::ReadAllText($badTarget) -ceq 'old valid fixture') 'Failed restart did not restore the old file.'

    # Replace only the metadata function with a literal fixture, including inside
    # the background runspace. No GitHub request is made in this test process.
    $fixtureJson = $release | ConvertTo-Json -Depth 8 -Compress
    Set-Item -LiteralPath function:Get-UpdateReleaseInfo -Value ([scriptblock]::Create("return ('" + $fixtureJson.Replace("'", "''") + "' | ConvertFrom-Json)"))
    Assert-Update ((Test-UpdateAvailable) -eq '3.10.2') 'Manual availability comparison failed.'
    $Global:ScriptVersion = '99.0.0'
    Assert-Update ($null -eq (Test-UpdateAvailable)) 'An older version was offered as an update.'
    $Global:ScriptVersion = '3.3.7'
    $Global:UserSettings.CheckForUpdates = $false
    Assert-Update ($null -eq (Start-BackgroundUpdateCheck)) 'Disabled automatic checks were started.'
    $Global:UserSettings.CheckForUpdates = $true
    $Global:UserSettings.LastUpdateCheckUtc = [DateTime]::UtcNow.ToString('o')
    Assert-Update ($null -eq (Start-BackgroundUpdateCheck)) 'The 12-hour check cache was ignored.'
    $check = Start-BackgroundUpdateCheck -Force
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    $result = $null
    while (-not $result -and [DateTime]::UtcNow -lt $deadline) { $result = Receive-BackgroundUpdateCheck -Check $check; if (-not $result) { Start-Sleep -Milliseconds 30 } }
    Assert-Update ([bool]$result.Completed) 'Background result did not complete.'
    Assert-Update ($result.Version -eq '3.10.2' -and -not $result.Error) 'Background candidate did not match the fixture release.'
    Assert-Update ([bool]$check.Disposed) 'Background runspace was not disposed.'
    Assert-Update ((Receive-BackgroundUpdateCheck -Check $check).Version -eq $result.Version) 'Reading the completed result twice failed.'
    Stop-BackgroundUpdateCheck -Check $check
    Stop-BackgroundUpdateCheck -Check $check
    Assert-Update ($script:savedSettings -ge 1) 'Background completion did not save its cache timestamp.'
    [pscustomobject]@{ Passed = $true; Checks = $script:checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
}
finally {
    $env:QBITLAUNCHER_SKIP_UPDATE_CHECK = $previousSkipCheck
    if ($check) { Stop-BackgroundUpdateCheck -Check $check }
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    if ($resolvedRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedRoot).StartsWith('qbitlauncher-update-tests-') -and [IO.Directory]::Exists($resolvedRoot)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
