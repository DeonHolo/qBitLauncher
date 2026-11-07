# Execute only disposable fixture scripts to verify native launch argument handling.
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
. (Join-Path $repository 'Source\Workflow.ps1')
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-launch-tests-' + [Guid]::NewGuid().ToString('N'))
$checks = 0
try {
    [IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $unicode = -join [char[]]@(0x6e38, 0x620f)
    $names = @('plain', 'space [brackets] & more', 'literal %WINDIR%', ('Unicode ' + $unicode), "apostrophe's !bang!")
    foreach ($extension in @('.bat', '.cmd')) {
        foreach ($name in $names) {
            $folder = Join-Path $testRoot ($name + '-' + $extension.Substring(1))
            [IO.Directory]::CreateDirectory($folder) | Out-Null
            $path = Join-Path $folder ($name + $extension)
            $content = "@echo off`r`n> `"%~dp0result.txt`" echo %~1`r`nexit /b 0`r`n"
            [IO.File]::WriteAllText($path, $content, [Text.Encoding]::ASCII)
            $info = New-RunnableStartInfo -Path $path -Arguments '"fixture argument"' -AsAdmin $false
            if ($info.Verb) { throw 'Non-admin launch unexpectedly requested elevation.' }; $checks++
            if ($info.Arguments -notmatch '-EncodedCommand ') { throw 'Script launch did not use the Unicode handoff.' }; $checks++
            # The returned command is safe to run: its only target is this fixture.
            $info.UseShellExecute = $false; $info.CreateNoWindow = $true
            $process = [Diagnostics.Process]::Start($info)
            try {
                if (-not $process.WaitForExit(10000)) { $process.Kill(); throw 'Fixture launch timed out.' }
                if ($process.ExitCode -ne 0) { throw "Fixture script returned $($process.ExitCode): $path" }
            } finally { $process.Dispose() }
            $result = Join-Path $folder 'result.txt'
            if (-not [IO.File]::Exists($result) -or [IO.File]::ReadAllText($result).Trim() -cne 'fixture argument') { throw "The script path or arguments were corrupted: $path" }; $checks++
            if ([IO.File]::ReadAllText($path) -cne $content) { throw 'Launching changed the original script.' }; $checks++
        }
    }
    $exePath = Join-Path $testRoot ('original ' + $unicode + '.exe')
    [IO.File]::Copy((Join-Path $env:WINDIR 'System32\whoami.exe'), $exePath)
    $info = New-RunnableStartInfo -Path $exePath -Arguments '--example "two words"' -AsAdmin $false
    if ($info.FileName -cne $exePath -or $info.Arguments -cne '--example "two words"' -or $info.WorkingDirectory -cne $testRoot) { throw 'EXE launch lost its original path, arguments or working directory.' }; $checks++
    $adminInfo = New-RunnableStartInfo -Path $exePath -AsAdmin $true
    if ($adminInfo.Verb -ne 'runas' -or -not $adminInfo.UseShellExecute) { throw 'Admin launch preference was not respected.' }; $checks++
    $unsupportedRejected = $false
    $textPath = Join-Path $testRoot 'readme.txt'; [IO.File]::WriteAllText($textPath, 'fixture')
    try { New-RunnableStartInfo -Path $textPath | Out-Null } catch { $unsupportedRejected = $true }
    if (-not $unsupportedRejected) { throw 'An unsupported file was accepted as a runnable.' }; $checks++
    [pscustomobject]@{ Passed = $true; Checks = $checks; Suite = 'Launch'; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
}
finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('qbitlauncher-launch-tests-') -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Force -Recurse }
}
