# Session persistence uses isolated settings and disposable folders only.
param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Source\Services.ps1')
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-session-tests-' + [Guid]::NewGuid().ToString('N'))
$script:checks = 0
function Assert-Session { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message }; $script:checks++ }
function Assert-SessionFailure { param([scriptblock]$Action, [string]$Pattern); $failure = ''; try { & $Action | Out-Null } catch { $failure = $_.Exception.Message }; Assert-Session ($failure -match $Pattern) ('Expected a session error matching ' + $Pattern) }
try {
    [void][IO.Directory]::CreateDirectory($testRoot)
    $Global:ConfigFile = Join-Path $testRoot 'config.json'; $Global:DataDirectory = Join-Path $testRoot 'appdata'
    Get-UserSettings
    Assert-Session ($Global:UserSettings.SavedSessions -is [hashtable] -and @(Get-LauncherSessions).Count -eq 0) 'Existing installations must begin with an empty, mutable session history.'
    $source = Join-Path $testRoot ('Download [fixture] ' + [char]0x6e38 + '.rar')
    $root = Join-Path $testRoot ('Extracted [fixture] ' + [char]0x6e38)
    [IO.File]::WriteAllText($source, 'fixture archive'); [void][IO.Directory]::CreateDirectory($root)
    $game = Join-Path $root 'Game.exe'; [IO.File]::WriteAllText($game, 'fixture runnable')
    $name = 'Example ' + [char]0x6e38
    $first = Save-LauncherSession -SourcePath $source -RootFolder $root -ArchivePath $source -TorrentHash ('a' * 40) -TorrentName $name
    Get-UserSettings
    $loaded = Get-LauncherSessionRequest -Id $first.Id
    Assert-Session ($loaded.SourcePath -ceq $source -and $loaded.RootFolder -ceq $root -and $loaded.ArchivePath -ceq $source) 'A session lost its source, extraction destination, or archive after restarting.'
    Assert-Session ($loaded.TorrentHash -eq ('a' * 40) -and $loaded.TorrentName -ceq $name -and $loaded.Name -ceq $name) 'A session lost its torrent identity or Unicode display name.'
    Assert-Session ((Get-LauncherSession -SourcePath $source.ToUpperInvariant()).Id -eq $first.Id) 'Source lookup must ignore path casing.'
    Assert-Session ((Get-LauncherSession -RootFolder ($root + '\')).Id -eq $first.Id) 'Root lookup must ignore trailing separators.'
    $same = Save-LauncherSession -SourcePath $source.ToUpperInvariant() -RootFolder ($root + '\')
    Assert-Session ($same.Id -eq $first.Id -and @(Get-LauncherSessions).Count -eq 1 -and $same.TorrentHash -eq ('a' * 40)) 'Reopening a known location duplicated the session or erased its torrent context.'
    Assert-Session ($same.CreatedUtc -ceq $first.CreatedUtc) 'Reopening a session changed its creation date.'
    $otherRoot = Join-Path $testRoot 'Different extraction'; [void][IO.Directory]::CreateDirectory($otherRoot)
    $second = Save-LauncherSession -SourcePath $source -RootFolder $otherRoot -ArchivePath $source
    Assert-Session ($second.Id -ne $first.Id -and @(Get-LauncherSessions).Count -eq 2) 'Different extraction destinations must remain distinct sessions.'

    [IO.File]::Delete($source)
    Assert-Session ((Get-LauncherSessionRequest -Id $first.Id).Id -eq $first.Id) 'Removing the source archive must not prevent reopening an extracted game.'
    [IO.File]::Delete($game); [IO.Directory]::Delete($root)
    Assert-SessionFailure { Get-LauncherSessionRequest -Id $first.Id } 'moved or was deleted'
    Assert-SessionFailure { Get-LauncherSessionRequest -Id 'unknown' } 'no longer available'
    [void][IO.Directory]::CreateDirectory($root); [IO.File]::WriteAllText($game, 'fixture runnable')

    Set-LauncherSessionTorrentRemoved -Id $first.Id
    Get-UserSettings
    $removed = Get-LauncherSessionRequest -Id $first.Id
    Assert-Session ($removed.TorrentRemoved -and -not $removed.TorrentHash -and $removed.RootFolder -ceq $root) 'Removing a torrent must clear its saved cleanup identity and keep the game folder.'
    $again = Save-LauncherSession -SourcePath $source -RootFolder $root
    Assert-Session ($again.TorrentRemoved -and -not $again.TorrentHash) 'Reopening must not resurrect an already removed torrent entry.'

    # Simulate two windows that loaded the same baseline and saved different games.
    Get-UserSettings
    $olderSettings = ConvertTo-SettingsHashtable $Global:UserSettings
    $olderBaseline = ConvertTo-SettingsHashtable $Global:SettingsBaseline
    $extraRoot = Join-Path $testRoot 'Third game'; [void][IO.Directory]::CreateDirectory($extraRoot)
    $third = Save-LauncherSession -SourcePath $extraRoot -RootFolder $extraRoot
    $Global:UserSettings = $olderSettings; $Global:SettingsBaseline = $olderBaseline
    $fourthRoot = Join-Path $testRoot 'Fourth game'; [void][IO.Directory]::CreateDirectory($fourthRoot)
    $fourth = Save-LauncherSession -SourcePath $fourthRoot -RootFolder $fourthRoot
    Get-UserSettings
    Assert-Session ((Get-LauncherSession -Id $third.Id) -and (Get-LauncherSession -Id $fourth.Id)) 'Saving from another game window erased a concurrent session.'
    $entries = @(Get-LauncherSessions)
    Assert-Session ($entries[0].Id -eq $fourth.Id) 'The latest opened session must appear first.'

    $null = Remove-LauncherSession -Id $first.Id
    Get-UserSettings
    Assert-Session (-not (Get-LauncherSession -Id $first.Id) -and [IO.Directory]::Exists($root) -and [IO.File]::Exists($game)) 'Forgetting history must remove only the record, leaving game files intact.'
    Assert-Session ((Get-LauncherSession -Id $second.Id) -and (Get-LauncherSession -Id $third.Id)) 'Forgetting one session affected other games.'
    $Global:UserSettings.SavedSessions['invalid'] = @{ SourcePath = 'relative'; RootFolder = 'relative' }
    Assert-Session (@(Get-LauncherSessions | Where-Object { $_.Id -eq 'invalid' }).Count -eq 0) 'Invalid saved locations should not become reopening targets.'
    $result = @{ Passed = $true; Checks = $script:checks; Suite = 'Sessions'; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { [pscustomobject]$result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'Sessions' } | ConvertTo-Json)); exit 1 }
    throw
}
finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('qbitlauncher-session-tests-') -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
