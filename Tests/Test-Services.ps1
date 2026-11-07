param([string]$SourcePath, [string]$ResultPath)

$ErrorActionPreference = 'Stop'
$serviceSource = Join-Path $PSScriptRoot '..\Source\Services.ps1'
. $serviceSource
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-service-tests-' + [Guid]::NewGuid().ToString('N'))
$script:checks = 0
$children = New-Object System.Collections.ArrayList
$tickets = New-Object System.Collections.ArrayList
$savedGlobals = @{}
foreach ($name in @('DataDirectory', 'ConfigFile', 'UserSettings', 'Themes', 'ThemeSelection', 'CurrentTheme', 'PasswordDecryptionFailed', 'ProtectedSettingsPassword', 'LogMaxBytes', 'SettingsBaseline', 'SettingsBaselinePath')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $savedGlobals[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}

function Assert-Service {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}

function Get-TestLiteral {
    param([string]$Value)
    return "'" + $Value.Replace("'", "''") + "'"
}

function Start-ServiceTestChild {
    param([string]$Code, [string]$Name)
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Code))
    $executable = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $outPath = Join-Path $testRoot ($Name + '.stdout')
    $errPath = Join-Path $testRoot ($Name + '.stderr')
    $child = Start-Process -FilePath $executable -ArgumentList @('-NoProfile', '-NonInteractive', '-EncodedCommand', $encoded) -WindowStyle Hidden -RedirectStandardOutput $outPath -RedirectStandardError $errPath -PassThru
    $null = $child.Handle
    $child | Add-Member -NotePropertyName TestErrorPath -NotePropertyValue $errPath
    [void]$children.Add($child)
    return $child
}

function Complete-ServiceTestChild {
    param($Child)
    if (-not $Child.WaitForExit(15000)) { $Child.Kill(); throw 'A service test helper timed out.' }
    $Child.Refresh()
    Assert-Service ($Child.ExitCode -eq 0) ('A service test helper failed: ' + [IO.File]::ReadAllText($Child.TestErrorPath))
}

try {
    [IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $Global:DataDirectory = Join-Path $testRoot 'appdata'
    $Global:ConfigFile = Join-Path $testRoot 'config [test].json'
    $script:LogFile = Join-Path $testRoot 'services.log'
    $Global:Themes = @{ Dracula = @{}; Light = @{} }
    $Global:UserSettings = New-DefaultUserSettings
    Assert-Service ($Global:UserSettings.FilePreferences -is [hashtable]) 'File preferences must be mutable.'
    Assert-Service ($Global:UserSettings.DefaultAdmin -eq $true) 'The existing admin launch default must be preserved.'
    Assert-Service ((Get-AppDataDirectory) -ceq [IO.Path]::GetFullPath($Global:DataDirectory)) 'Tests must use their isolated data directory.'

    $secret = 'private secret ' + (-join [char[]]@(0x6e38, 0x620f))
    $protected = Protect-AppSecret $secret
    Assert-Service ($protected.StartsWith('dpapi:v1:') -and -not $protected.Contains($secret)) 'The protected credential must not contain plaintext.'
    Assert-Service ((Unprotect-AppSecret $protected) -ceq $secret) 'The protected credential did not round-trip.'
    Assert-Service ((Protect-AppSecret '') -ceq '') 'An empty password should stay empty.'
    $invalidRejected = $false
    try { $null = Unprotect-AppSecret 'invalid' } catch { $invalidRejected = $true }
    Assert-Service $invalidRejected 'An unknown credential format was accepted.'

    $legacy = @{ Theme = 'Light'; QbittorrentUsername = 'private-user'; QbittorrentPassword = $secret; DefaultAdmin = $false; FutureFeature = @{ Enabled = $false; Items = @('one') }; RecentFolders = @() }
    [IO.File]::WriteAllText($Global:ConfigFile, ($legacy | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($false)))
    Get-UserSettings
    Assert-Service ($Global:UserSettings.QbittorrentPassword -ceq $secret) 'Legacy password migration changed the runtime password.'
    $diskText = [IO.File]::ReadAllText($Global:ConfigFile)
    $disk = $diskText | ConvertFrom-Json
    Assert-Service (-not ($disk.PSObject.Properties.Name -contains 'QbittorrentPassword')) 'A plaintext password property was saved.'
    Assert-Service (-not $diskText.Contains($secret)) 'The saved settings contain the plaintext password.'
    Assert-Service ($disk.QbittorrentPasswordProtected.StartsWith('dpapi:v1:')) 'Legacy credentials were not migrated automatically.'
    Assert-Service ($Global:UserSettings.DefaultAdmin -eq $false) 'A saved false setting was replaced by a default.'
    Assert-Service ($Global:UserSettings.FutureFeature.Enabled -eq $false -and $Global:UserSettings.FutureFeature.Items.Count -eq 1) 'Future/nested settings were not preserved.'
    Assert-Service ($Global:UserSettings.RecentFolders.Count -eq 0) 'An empty settings array was not preserved.'
    $Global:UserSettings = $null
    Get-UserSettings
    Assert-Service ($Global:UserSettings.QbittorrentPassword -ceq $secret) 'The encrypted saved password did not reload.'
    Assert-Service ($Global:UserSettings.FutureFeature -is [hashtable]) 'Nested JSON objects must reload as mutable hashtables.'

    $target = Join-Path $testRoot 'Game [test].cmd'
    [IO.File]::WriteAllText($target, '@echo off')
    Assert-Service (Set-FilePreference -Path $target -Preference @{ ShortcutName = 'A nicer label' }) 'Shortcut label preferences did not save.'
    Assert-Service (Set-FilePreference -Path $target.ToUpperInvariant() -Preference @{ Arguments = '--windowed'; AsAdmin = $false }) 'Launch preferences did not save.'
    $preference = Get-FilePreference $target
    Assert-Service ($preference.ShortcutName -ceq 'A nicer label' -and $preference.Arguments -ceq '--windowed' -and $preference.AsAdmin -eq $false) 'File preference updates must merge case-insensitively.'
    Assert-Service ([IO.File]::Exists($target) -and [IO.File]::ReadAllText($target) -ceq '@echo off') 'Saving a display label must never modify the target.'
    $preference.ShortcutName = 'unsaved change'
    Assert-Service ((Get-FilePreference $target).ShortcutName -ceq 'A nicer label') 'Reading a preference must return a detached value.'
    Assert-Service (Add-RecentFolder $testRoot) 'A recent folder did not save.'
    Assert-Service (Add-RecentFolder $testRoot.ToUpperInvariant()) 'A recent folder update did not save.'
    Assert-Service ($Global:UserSettings.RecentFolders.Count -eq 1) 'Recent folders must not duplicate paths with different casing.'
    Assert-Service (Add-RecentFolder ([IO.Path]::GetPathRoot($testRoot))) 'A drive root could not be saved as a recent folder.'
    Assert-Service ($Global:UserSettings.RecentFolders[0] -ceq [IO.Path]::GetPathRoot($testRoot)) 'A recent drive root became a drive-relative path.'
    Assert-Service (Remove-RecentFolder ([IO.Path]::GetPathRoot($testRoot))) 'A recent folder could not be forgotten.'
    Assert-Service (Remove-FilePreference $target) 'A file preference could not be forgotten.'
    Assert-Service ((Get-FilePreference $target).Count -eq 0) 'A forgotten file preference remained stored.'

    $failedCiphertext = 'dpapi:v1:not-a-valid-ciphertext'
    [IO.File]::WriteAllText($Global:ConfigFile, (@{ QbittorrentPasswordProtected = $failedCiphertext } | ConvertTo-Json))
    Get-UserSettings
    Assert-Service ($Global:PasswordDecryptionFailed -and $Global:UserSettings.QbittorrentPassword -ceq '') 'An unreadable credential should request re-entry.'
    Assert-Service (Save-UserSettings) 'Unrelated settings should still be saveable with an unreadable credential.'
    $disk = [IO.File]::ReadAllText($Global:ConfigFile) | ConvertFrom-Json
    Assert-Service ($disk.QbittorrentPasswordProtected -ceq $failedCiphertext) 'An unreadable credential must not be silently erased.'
    $Global:UserSettings.QbittorrentPassword = $secret
    Assert-Service (Save-UserSettings) 'A re-entered credential could not be saved.'
    Assert-Service (-not $Global:PasswordDecryptionFailed) 'Re-entering a password should clear the decryption error.'

    $Global:UserSettings.QbittorrentUsername = 'private-user'
    Write-LogMessage ('Private API: private-user ' + $secret + ' password=abc; token=xyz')
    $logged = [IO.File]::ReadAllText($script:LogFile)
    Assert-Service (-not $logged.Contains($secret) -and -not $logged.Contains('private-user') -and -not $logged.Contains('abc') -and -not $logged.Contains('xyz')) 'A log entry exposed saved credentials.'
    Assert-Service ($logged.Contains('[redacted]')) 'Credential redaction should be visible.'

    $concurrentLog = Join-Path $testRoot 'concurrent.log'
    $prefix = '$ErrorActionPreference="Stop"; . ' + (Get-TestLiteral $serviceSource) + '; $Global:DataDirectory=' + (Get-TestLiteral $Global:DataDirectory) + '; $script:LogFile=' + (Get-TestLiteral $concurrentLog) + '; $Global:UserSettings=@{}; '
    $childOne = Start-ServiceTestChild -Code ($prefix + '0..24 | ForEach-Object { Write-LogMessage ("worker-one-"+$_) }; exit 0') -Name 'logger-one'
    $childTwo = Start-ServiceTestChild -Code ($prefix + '0..24 | ForEach-Object { Write-LogMessage ("worker-two-"+$_) }; exit 0') -Name 'logger-two'
    Complete-ServiceTestChild $childOne
    Complete-ServiceTestChild $childTwo
    $entries = [IO.File]::ReadAllLines($concurrentLog)
    Assert-Service ($entries.Count -eq 50) 'Concurrent append logging lost entries.'
    Assert-Service (@($entries | Select-String 'worker-one-').Count -eq 25 -and @($entries | Select-String 'worker-two-').Count -eq 25) 'Concurrent log writers did not retain their entries.'

    $script:LogFile = Join-Path $testRoot 'rotation.log'
    $Global:LogMaxBytes = 256
    1..30 | ForEach-Object { Write-LogMessage ('rotation-entry-' + $_ + ' padding-for-a-bounded-log') }
    Assert-Service ([IO.File]::Exists($script:LogFile + '.1') -and [IO.File]::Exists($script:LogFile + '.3')) 'Logs did not rotate into the expected archives.'
    Assert-Service (-not [IO.File]::Exists($script:LogFile + '.4')) 'The logger retained too many archives.'
    Assert-Service ((Get-Item -LiteralPath $script:LogFile).Length -le 256) 'The live log exceeded its configured size.'
    $Global:LogMaxBytes = 2MB

    $first = New-WorkflowLock -Key 'first-source'
    [void]$tickets.Add($first)
    $duplicate = New-WorkflowLock -Key 'FIRST-SOURCE'
    [void]$tickets.Add($duplicate)
    Assert-Service $duplicate.Duplicate 'The same source must not get a second extraction.'
    $second = New-WorkflowLock -Key 'second-source'
    [void]$tickets.Add($second)
    Assert-Service (-not $first.Duplicate -and -not $second.Duplicate -and $first.Handle -and $second.Handle) 'Different games must hold their own locks at the same time.'
    Exit-WorkflowLock $first
    Exit-WorkflowLock $first
    Exit-WorkflowLock $second

    $readyPath = Join-Path $testRoot 'parallel-child.ready'
    $releasePath = Join-Path $testRoot 'parallel-child.release'
    # Deliberately exit without disposing: Windows must release a crashed owner's handle.
    $childCode = '$ErrorActionPreference="Stop"; . ' + (Get-TestLiteral $serviceSource) + '; $Global:DataDirectory=' + (Get-TestLiteral $Global:DataDirectory) + '; $lock=New-WorkflowLock -Key "child-source"; if ($lock.Duplicate) { throw "Child source unavailable" }; [IO.File]::WriteAllText(' + (Get-TestLiteral $readyPath) + ',"ready"); $deadline=[DateTime]::UtcNow.AddSeconds(15); while (-not [IO.File]::Exists(' + (Get-TestLiteral $releasePath) + ')) { if ([DateTime]::UtcNow -gt $deadline) { throw "Parent did not release child" }; Start-Sleep -Milliseconds 50 }; exit 0'
    $parallelChild = Start-ServiceTestChild -Code $childCode -Name 'parallel-child'
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while (-not [IO.File]::Exists($readyPath)) {
        if ([DateTime]::UtcNow -gt $deadline -or $parallelChild.HasExited) { throw 'The parallel workflow helper did not become ready.' }
        Start-Sleep -Milliseconds 50
    }
    $duplicateChild = New-WorkflowLock -Key 'CHILD-SOURCE'
    [void]$tickets.Add($duplicateChild)
    Assert-Service $duplicateChild.Duplicate 'Duplicate detection did not work across processes.'
    $afterChild = New-WorkflowLock -Key 'after-child-source'
    [void]$tickets.Add($afterChild)
    Assert-Service (-not $afterChild.Duplicate -and $afterChild.Handle -and -not $parallelChild.HasExited) 'A smaller game must start while another process is still open.'
    [IO.File]::WriteAllText($releasePath, 'release')
    Complete-ServiceTestChild $parallelChild
    $recovered = New-WorkflowLock -Key 'child-source'
    [void]$tickets.Add($recovered)
    Assert-Service (-not $recovered.Duplicate -and $recovered.Handle) 'An exited process left its source locked.'
    Exit-WorkflowLock $afterChild

    # Both processes load the same snapshot, then save changes to different games.
    Get-UserSettings
    $parentBaseline = ConvertTo-SettingsHashtable $Global:UserSettings
    $settingsChild = '$ErrorActionPreference="Stop"; . ' + (Get-TestLiteral $serviceSource) + '; $Global:ConfigFile=' + (Get-TestLiteral $Global:ConfigFile) + '; Get-UserSettings; if (-not (Set-FilePreference -Path ' + (Get-TestLiteral (Join-Path $testRoot 'other game.exe')) + ' -Preference @{Label="Other game"})) { throw "Preference save failed" }; $Global:UserSettings.Theme="Dracula"; $Global:UserSettings.QbittorrentPassword="new credential"; if (-not (Save-UserSettings)) { throw "Settings save failed" }; exit 0'
    $settingsWriter = Start-ServiceTestChild -Code $settingsChild -Name 'parallel-settings'
    Complete-ServiceTestChild $settingsWriter
    Assert-Service (Set-FilePreference -Path $target -Preference @{ Label = 'Parent game' }) 'The older window could not save its own game.'
    Get-UserSettings
    Assert-Service ((Get-FilePreference (Join-Path $testRoot 'other game.exe')).Label -ceq 'Other game') 'Saving from an older window erased another game shortcut name.'
    Assert-Service ((Get-FilePreference $target).Label -ceq 'Parent game') 'The parent game preference was lost.'
    Assert-Service ($Global:UserSettings.Theme -eq 'Dracula' -and $Global:UserSettings.QbittorrentPassword -ceq 'new credential') 'An older window overwrote settings or credentials changed by another process.'

    $result = [pscustomobject]@{ Passed = $true; Checks = $script:checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString(); Suite = 'Services' }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) }
    else { $result }
}
catch {
    if ($ResultPath) {
        [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'Services' } | ConvertTo-Json))
        exit 1
    }
    throw
}
finally {
    foreach ($child in $children) {
        if (-not $child.HasExited) { $child.Kill(); $null = $child.WaitForExit(5000) }
        $child.Dispose()
    }
    foreach ($ticket in $tickets) { Exit-WorkflowLock $ticket }
    foreach ($name in $savedGlobals.Keys) {
        if ($savedGlobals[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $savedGlobals[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    if ($resolvedRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedRoot).StartsWith('qbitlauncher-service-tests-') -and [IO.Directory]::Exists($resolvedRoot)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
