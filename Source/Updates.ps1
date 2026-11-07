# Release metadata and Unicode-safe update handoff. Definitions only: safe to import.
function Get-UpdateReleaseInfo {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    return Invoke-RestMethod -Uri 'https://api.github.com/repos/DeonHolo/qBitLauncher/releases/latest' -Headers @{ 'User-Agent' = 'qBitLauncher'; Accept = 'application/vnd.github+json' } -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
}

function Get-ReleaseUpdateCandidate {
    param([Parameter(Mandatory = $true)]$Release, [bool]$IsCompiled = $Global:IsCompiledExe)

    if ($Release.draft -or $Release.prerelease) { return $null }
    $version = $null
    if (-not [Version]::TryParse(([string]$Release.tag_name -replace '^[vV]', ''), [ref]$version)) { return $null }
    $assetName = if ($IsCompiled) { 'qBitLauncher.exe' } else { 'qBitLauncher.ps1' }
    $asset = @($Release.assets | Where-Object { $_.name -ceq $assetName }) | Select-Object -First 1
    $sums = @($Release.assets | Where-Object { $_.name -ceq 'SHA256SUMS.txt' }) | Select-Object -First 1
    $url = if ($asset) { [string]$asset.browser_download_url } else { '' }
    if (-not $asset -and -not $IsCompiled -and $sums) {
        $url = 'https://raw.githubusercontent.com/DeonHolo/qBitLauncher/' + [Uri]::EscapeDataString([string]$Release.tag_name) + '/qBitLauncher.ps1'
    }
    if ([string]::IsNullOrWhiteSpace($url)) { return $null }
    if ($url -notmatch '^https://github\.com/DeonHolo/qBitLauncher/releases/download/' -and
        $url -notmatch '^https://raw\.githubusercontent\.com/DeonHolo/qBitLauncher/') {
        throw 'The release asset URL is outside the configured repository.'
    }
    $hash = $null
    if ($asset -and ([string]$asset.digest -match '^sha256:([a-fA-F0-9]{64})$')) { $hash = $Matches[1].ToLowerInvariant() }
    return [pscustomobject]@{
        Version = $version.ToString(); Tag = [string]$Release.tag_name; AssetName = $assetName
        DownloadUrl = $url; SHA256 = $hash; ChecksumsUrl = $(if ($sums) { [string]$sums.browser_download_url } else { $null })
        ReleaseUrl = [string]$Release.html_url
    }
}

function Get-RemoteVersion {
    try {
        $release = Get-UpdateReleaseInfo
        $candidate = Get-ReleaseUpdateCandidate -Release $release
        $Global:AvailableUpdate = $candidate
        if ($candidate) { return $candidate.Version }
    }
    catch { Write-LogMessage "Failed to check release updates: $($_.Exception.Message)" }
    return $null
}

function Test-UpdateAvailable {
    $version = Get-RemoteVersion
    if ($version -and [Version]$version -gt [Version]$Global:ScriptVersion) { return $version }
    return $null
}

function Start-BackgroundUpdateCheck {
    param([switch]$Force)
    if ($env:QBITLAUNCHER_SKIP_UPDATE_CHECK -eq '1') { return $null }
    if (-not $Force -and $Global:UserSettings -and $Global:UserSettings.ContainsKey('CheckForUpdates') -and -not $Global:UserSettings.CheckForUpdates) { return $null }
    if (-not $Force -and $Global:UserSettings -and $Global:UserSettings.LastUpdateCheckUtc) {
        $last = [DateTime]::MinValue
        if ([DateTime]::TryParse([string]$Global:UserSettings.LastUpdateCheckUtc, [ref]$last) -and $last.ToUniversalTime().AddHours(12) -gt [DateTime]::UtcNow) { return $null }
    }
    $powerShell = [PowerShell]::Create()
    try {
        $code = ${function:Get-UpdateReleaseInfo}.ToString()
        $null = $powerShell.AddScript($code)
        $handle = $powerShell.BeginInvoke()
        return [pscustomobject]@{ PowerShell = $powerShell; Handle = $handle; Disposed = $false; Result = $null }
    }
    catch { $powerShell.Dispose(); Write-LogMessage "Could not start background update check: $($_.Exception.Message)"; return $null }
}

function Receive-BackgroundUpdateCheck {
    param($Check)
    if (-not $Check) { return $null }
    if ($Check.Disposed) { return $Check.Result }
    if (-not $Check.Handle.IsCompleted) { return $null }
    $result = [pscustomobject]@{ Completed = $true; Version = $null; Error = $null; Release = $null }
    try {
        $items = $Check.PowerShell.EndInvoke($Check.Handle)
        if ($Check.PowerShell.HadErrors) { throw ($Check.PowerShell.Streams.Error | Select-Object -First 1) }
        $release = @($items) | Select-Object -First 1
        if (-not $release) { throw 'GitHub did not return release metadata.' }
        $candidate = Get-ReleaseUpdateCandidate -Release $release
        $Global:AvailableUpdate = $candidate
        $result.Release = $release
        if ($candidate -and [Version]$candidate.Version -gt [Version]$Global:ScriptVersion) { $result.Version = $candidate.Version }
        if ($Global:UserSettings) {
            $Global:UserSettings.LastUpdateCheckUtc = [DateTime]::UtcNow.ToString('o')
            $null = Save-UserSettings
        }
    }
    catch { $result.Error = $_.Exception.Message; Write-LogMessage "Background update check failed: $($result.Error)" }
    finally { $Check.PowerShell.Dispose(); $Check.Disposed = $true }
    $Check.Result = $result
    return $result
}

function Stop-BackgroundUpdateCheck {
    param($Check)
    if ($Check -and -not $Check.Disposed) {
        try { $Check.PowerShell.Stop() } catch { }
        $Check.PowerShell.Dispose()
        $Check.Disposed = $true
    }
}

function Get-UpdateFileSHA256 {
    param([Parameter(Mandatory = $true)][string]$Path)
    $hasher = [Security.Cryptography.SHA256]::Create()
    try {
        $stream = [IO.File]::OpenRead($Path)
        try { return [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-', '').ToLowerInvariant() }
        finally { $stream.Dispose() }
    }
    finally { $hasher.Dispose() }
}

function Test-UpdateAssetHash {
    param([string]$Path, [string]$ExpectedSHA256)
    $expected = $ExpectedSHA256 -replace '^sha256:', ''
    if ($expected -notmatch '^[a-fA-F0-9]{64}$') { throw 'No valid SHA-256 checksum is available for the release asset.' }
    if ((Get-UpdateFileSHA256 -Path $Path) -cne $expected.ToLowerInvariant()) { throw 'The downloaded update did not match its published SHA-256 checksum.' }
    return $true
}

function ConvertTo-WindowsArgumentString {
    param([AllowEmptyCollection()][AllowNull()][string[]]$Arguments)
    $parts = foreach ($argument in $Arguments) {
        $builder = New-Object Text.StringBuilder
        $null = $builder.Append('"')
        $slashes = 0
        foreach ($character in ([string]$argument).ToCharArray()) {
            if ($character -eq [char]92) { $slashes++; continue }
            if ($character -eq [char]34) {
                $null = $builder.Append([char]92, (2 * $slashes + 1)).Append([char]34)
            }
            else {
                if ($slashes) { $null = $builder.Append([char]92, $slashes) }
                $null = $builder.Append($character)
            }
            $slashes = 0
        }
        if ($slashes) { $null = $builder.Append([char]92, (2 * $slashes)) }
        $null = $builder.Append('"')
        $builder.ToString()
    }
    return ($parts -join ' ')
}

function Invoke-UpdateReplacement {
    param([Parameter(Mandatory = $true)]$Payload)
    $replacementMutex = New-Object Threading.Mutex($false, 'Local\qBitLauncher.Update')
    $hasReplacementLock = $false
    try {
        try { $hasReplacementLock = $replacementMutex.WaitOne(45000) }
        catch [Threading.AbandonedMutexException] { $hasReplacementLock = $true }
        if (-not $hasReplacementLock) { throw 'Another update did not finish before the replacement timeout.' }
        $target = [IO.Path]::GetFullPath([string]$Payload.TargetPath)
        $stage = [IO.Path]::GetFullPath([string]$Payload.StagedPath)
        $backup = [IO.Path]::GetFullPath([string]$Payload.BackupPath)
        $directory = [IO.Path]::GetDirectoryName($target)
        if ([IO.Path]::GetDirectoryName($stage) -ine $directory -or [IO.Path]::GetDirectoryName($backup) -ine $directory -or $stage -ieq $target -or $backup -ieq $target -or $backup -ieq $stage) {
            throw 'Update files must use distinct paths in the application directory.'
        }
        if (-not [IO.File]::Exists($target)) { throw 'The application to update no longer exists.' }
        $null = Test-UpdateAssetHash -Path $stage -ExpectedSHA256 ([string]$Payload.SHA256)
        $waitUntil = [DateTime]::UtcNow.AddSeconds(45)
        if ([int]$Payload.ParentProcessId -gt 0) {
            while ($true) {
                $running = $null
                try { $running = [Diagnostics.Process]::GetProcessById([int]$Payload.ParentProcessId) } catch { break }
                try {
                    if ($Payload.ParentStartTimeUtc -and $running.StartTime.ToUniversalTime().ToString('o') -ne [string]$Payload.ParentStartTimeUtc) { break }
                    if ($running.HasExited) { break }
                }
                finally { if ($running) { $running.Dispose() } }
                if ([DateTime]::UtcNow -gt $waitUntil) { throw 'The running application did not close before the update timeout.' }
                Start-Sleep -Milliseconds 150
            }
        }
        # Revalidate after waiting, and keep all replacement files on the same volume.
        $null = Test-UpdateAssetHash -Path $stage -ExpectedSHA256 ([string]$Payload.SHA256)
        [IO.File]::Replace($stage, $target, $backup, $true)
        try {
            if (-not $Payload.SkipRestart) {
                $info = New-Object Diagnostics.ProcessStartInfo
                $info.UseShellExecute = $false
                $info.WorkingDirectory = $directory
                $launchArguments = @($Payload.Arguments | ForEach-Object { [string]$_ })
                if ($Payload.IsCompiled) {
                    $info.FileName = $target
                    $info.Arguments = ConvertTo-WindowsArgumentString -Arguments $launchArguments
                }
                else {
                    $info.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
                    $info.CreateNoWindow = $true
                    $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
                    $info.Arguments = ConvertTo-WindowsArgumentString -Arguments (@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', $target) + $launchArguments)
                }
                $process = [Diagnostics.Process]::Start($info)
                if (-not $process) { throw 'The updated application could not be started.' }
                $process.Dispose()
            }
            return [pscustomobject]@{ Updated = $true; TargetPath = $target; BackupPath = $backup }
        }
        catch {
            $failedPath = $target + '.failed-update-' + [Guid]::NewGuid().ToString('N') + '.bak'
            [IO.File]::Replace($backup, $target, $failedPath, $true)
            throw "The updated application failed to start; the previous version was restored. $($_.Exception.Message)"
        }
    }
    finally {
        if ($hasReplacementLock) { try { $replacementMutex.ReleaseMutex() } catch { } }
        $replacementMutex.Dispose()
    }
}

function Get-UpdateHandoffSource {
    $builder = New-Object Text.StringBuilder
    $null = $builder.AppendLine('param([Parameter(Mandatory = $true)][string]$PayloadPath)')
    $null = $builder.AppendLine('$ErrorActionPreference = ''Stop''')
    foreach ($name in @('Get-UpdateFileSHA256', 'Test-UpdateAssetHash', 'ConvertTo-WindowsArgumentString', 'Invoke-UpdateReplacement')) {
        $definition = (Get-Command $name -CommandType Function -ErrorAction Stop).Definition
        $null = $builder.AppendLine(('function ' + $name + ' {')).AppendLine($definition).AppendLine('}')
    }
    $null = $builder.AppendLine(@'
try {
    $payload = [IO.File]::ReadAllText($PayloadPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $result = Invoke-UpdateReplacement -Payload $payload
    [IO.File]::WriteAllText($PayloadPath + '.result.json', ($result | ConvertTo-Json -Depth 8), [Text.Encoding]::UTF8)
}
catch {
    [IO.File]::WriteAllText($PayloadPath + '.result.json', (@{ Updated = $false; Error = $_.Exception.Message } | ConvertTo-Json), [Text.Encoding]::UTF8)
    exit 1
}
'@)
    return $builder.ToString()
}

function New-UpdateHandoff {
    param([Parameter(Mandatory = $true)]$Payload, [Parameter(Mandatory = $true)][string]$Directory)
    [IO.Directory]::CreateDirectory($Directory) | Out-Null
    $id = [Guid]::NewGuid().ToString('N')
    $scriptPath = Join-Path $Directory ('handoff-' + $id + '.ps1')
    $payloadPath = Join-Path $Directory ('handoff-' + $id + '.json')
    [IO.File]::WriteAllText($scriptPath, (Get-UpdateHandoffSource), (New-Object Text.UTF8Encoding($true)))
    [IO.File]::WriteAllText($payloadPath, ($Payload | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($true)))
    $command = '& ' + "'" + $scriptPath.Replace("'", "''") + "' -PayloadPath '" + $payloadPath.Replace("'", "''") + "'"
    return [pscustomobject]@{ ScriptPath = $scriptPath; PayloadPath = $payloadPath; EncodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command)) }
}

function Get-UpdateExpectedSHA256 {
    param([Parameter(Mandatory = $true)]$Candidate)
    if ($Candidate.SHA256 -match '^[a-fA-F0-9]{64}$') { return $Candidate.SHA256.ToLowerInvariant() }
    if ($Candidate.ChecksumsUrl -notmatch '^https://github\.com/DeonHolo/qBitLauncher/releases/download/') { throw 'This release has no published SHA-256 checksum. Download it manually or wait for a verified release.' }
    $response = Invoke-WebRequest -Uri $Candidate.ChecksumsUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'qBitLauncher' } -TimeoutSec 15 -ErrorAction Stop
    foreach ($line in ([string]$response.Content -split '\r?\n')) {
        if ($line -match '^([a-fA-F0-9]{64})\s+\*?(.+?)\s*$' -and $Matches[2] -ceq $Candidate.AssetName) { return $Matches[1].ToLowerInvariant() }
    }
    throw 'The release checksums do not include the requested application asset.'
}

function Update-Script {
    param([switch]$Restart)
    $updateMutex = $null
    $hasLock = $false
    $stagedPath = $null
    try {
        $updateMutex = New-Object Threading.Mutex($false, 'Local\qBitLauncher.Update')
        try { $hasLock = $updateMutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $hasLock = $true }
        if (-not $hasLock) { Write-LogMessage 'An update is already in progress.'; return $false }
        $candidate = $Global:AvailableUpdate
        if (-not $candidate) { $null = Get-RemoteVersion; $candidate = $Global:AvailableUpdate }
        if (-not $candidate -or [Version]$candidate.Version -le [Version]$Global:ScriptVersion) { return $false }
        $expected = Get-UpdateExpectedSHA256 -Candidate $candidate
        $targetPath = [IO.Path]::GetFullPath($Global:ExecutablePath)
        $stagedPath = Join-Path ([IO.Path]::GetDirectoryName($targetPath)) ('qBitLauncher.update-' + [Guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension($targetPath))
        Write-LogMessage "Downloading verified release $($candidate.Tag)..."
        Invoke-WebRequest -Uri $candidate.DownloadUrl -OutFile $stagedPath -Headers @{ 'User-Agent' = 'qBitLauncher' } -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop | Out-Null
        $null = Test-UpdateAssetHash -Path $stagedPath -ExpectedSHA256 $expected
        if (-not $Global:IsCompiledExe) {
            $tokens = $null; $errors = $null
            $null = [Management.Automation.Language.Parser]::ParseFile($stagedPath, [ref]$tokens, [ref]$errors)
            if ($errors.Count) { throw 'The verified source update could not be parsed.' }
        }
        $arguments = @()
        # Preserve positional empty values when later arguments are present.
        if ($filePathFromQB -or $torrentHashFromQB -or $torrentNameFromQB) { $arguments = @([string]$filePathFromQB, [string]$torrentHashFromQB, [string]$torrentNameFromQB) }
        $current = [Diagnostics.Process]::GetCurrentProcess()
        $payload = @{
            TargetPath = $targetPath; StagedPath = $stagedPath; BackupPath = $targetPath + '.previous.bak'
            SHA256 = $expected; ParentProcessId = $PID; ParentStartTimeUtc = $current.StartTime.ToUniversalTime().ToString('o')
            IsCompiled = [bool]$Global:IsCompiledExe; Arguments = $arguments; SkipRestart = (-not $Restart)
        }
        $current.Dispose()
        if (-not $Global:IsCompiledExe -and -not $Restart) {
            $payload.ParentProcessId = 0
            $null = Invoke-UpdateReplacement -Payload $payload
            $Global:ScriptVersion = $candidate.Version
            Write-LogMessage 'Source updated and previous version retained; reopen to load the new code.'
            return $true
        }
        $handoff = New-UpdateHandoff -Payload $payload -Directory (Join-Path (Get-AppDataDirectory) 'updates')
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $info.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand ' + $handoff.EncodedCommand
        $info.UseShellExecute = $false
        $info.CreateNoWindow = $true
        $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $helper = [Diagnostics.Process]::Start($info)
        if (-not $helper) { throw 'The update handoff could not be started.' }
        $helper.Dispose()
        Write-LogMessage "Update handoff started; status will be saved to $($handoff.PayloadPath).result.json"
        [Environment]::Exit(0)
    }
    catch {
        Write-LogMessage "Update failed: $($_.Exception.Message)"
        if ($stagedPath -and [IO.File]::Exists($stagedPath)) { try { [IO.File]::Delete($stagedPath) } catch { } }
        if (Get-Command Show-ThemedMessageBox -ErrorAction SilentlyContinue) {
            $null = Show-ThemedMessageBox -Message "The current application has been kept.`n$($_.Exception.Message)" -Title 'Update Failed' -Icon 'Warning'
        }
        return $false
    }
    finally {
        if ($hasLock -and $updateMutex) { try { $updateMutex.ReleaseMutex() } catch { } }
        if ($updateMutex) { $updateMutex.Dispose() }
    }
}
