# Side-effect-free definitions for settings, user state, logging and job coordination.
function Get-AppDataDirectory {
    $override = Get-Variable -Name DataDirectory -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if (-not [string]::IsNullOrWhiteSpace($override)) {
        return [IO.Path]::GetFullPath($override)
    }
    if (-not [string]::IsNullOrWhiteSpace($env:QBITLAUNCHER_DATA_DIRECTORY)) {
        return [IO.Path]::GetFullPath($env:QBITLAUNCHER_DATA_DIRECTORY)
    }
    $base = [Environment]::GetFolderPath('LocalApplicationData')
    if ([string]::IsNullOrWhiteSpace($base)) { throw 'The current user has no local application data directory.' }
    return Join-Path $base 'qBitLauncher'
}

function ConvertTo-SettingsHashtable {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Collections.IDictionary]) {
        $result = @{}
        foreach ($key in $Value.Keys) { $result[[string]$key] = ConvertTo-SettingsHashtable $Value[$key] }
        return $result
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $result = @{}
        foreach ($property in $Value.PSObject.Properties) { $result[$property.Name] = ConvertTo-SettingsHashtable $property.Value }
        return $result
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $result = New-Object System.Collections.ArrayList
        foreach ($entry in $Value) { [void]$result.Add((ConvertTo-SettingsHashtable $entry)) }
        return ,$result.ToArray()
    }
    return $Value
}

function New-DefaultUserSettings {
    return @{
        Theme = 'Dracula'
        QbittorrentWebUrl = 'http://localhost:8080'
        QbittorrentUsername = ''
        QbittorrentPassword = ''
        FilePreferences = @{}
        RecentFolders = @()
        LastExtractionPath = ''
        DefaultAdmin = $true
        CheckForUpdates = $true
        LastUpdateCheckUtc = ''
        WindowState = @{}
    }
}

function Protect-AppSecret {
    param([AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    Add-Type -AssemblyName System.Security -ErrorAction Stop
    $bytes = [Text.Encoding]::UTF8.GetBytes($Value)
    $entropy = [Text.Encoding]::UTF8.GetBytes('qBitLauncher.settings.v1')
    try {
        $protected = [Security.Cryptography.ProtectedData]::Protect($bytes, $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
        return 'dpapi:v1:' + [Convert]::ToBase64String($protected)
    }
    finally { [Array]::Clear($bytes, 0, $bytes.Length) }
}

function Unprotect-AppSecret {
    param([AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    if (-not $Value.StartsWith('dpapi:v1:', [StringComparison]::Ordinal)) { throw 'Unrecognized saved credential format.' }
    Add-Type -AssemblyName System.Security -ErrorAction Stop
    $entropy = [Text.Encoding]::UTF8.GetBytes('qBitLauncher.settings.v1')
    $bytes = [Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($Value.Substring(9)), $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    try { return [Text.Encoding]::UTF8.GetString($bytes) }
    finally { [Array]::Clear($bytes, 0, $bytes.Length) }
}

function Get-ServiceKeyHash {
    param([string]$Value)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Get-UserSettings {
    $defaults = New-DefaultUserSettings
    $Global:PasswordDecryptionFailed = $false
    $Global:ProtectedSettingsPassword = ''
    $migratePassword = $false
    if ([string]::IsNullOrWhiteSpace($Global:ConfigFile)) {
        $Global:ConfigFile = Join-Path (Get-AppDataDirectory) 'config.json'
    }
    if ([IO.File]::Exists($Global:ConfigFile)) {
        try {
            $loaded = ConvertTo-SettingsHashtable ([IO.File]::ReadAllText($Global:ConfigFile, [Text.Encoding]::UTF8) | ConvertFrom-Json)
            if ($loaded -isnot [System.Collections.IDictionary]) { throw 'Settings must be a JSON object.' }
            foreach ($key in $loaded.Keys) { $defaults[$key] = $loaded[$key] }
            if ($loaded.Contains('QbittorrentPasswordProtected') -and -not [string]::IsNullOrEmpty($loaded.QbittorrentPasswordProtected)) {
                $Global:ProtectedSettingsPassword = [string]$loaded.QbittorrentPasswordProtected
                try { $defaults.QbittorrentPassword = Unprotect-AppSecret $Global:ProtectedSettingsPassword }
                catch {
                    $Global:PasswordDecryptionFailed = $true
                    $defaults.QbittorrentPassword = ''
                    Write-LogMessage 'The saved API credential could not be decrypted for this Windows user. Re-enter it in Settings.'
                }
            }
            elseif ($loaded.Contains('QbittorrentPassword')) { $migratePassword = $true }
        }
        catch { Write-LogMessage ('Settings could not be loaded; using defaults. ' + $_.Exception.GetType().Name) }
    }
    foreach ($key in @('FilePreferences', 'WindowState')) {
        if ($defaults[$key] -isnot [System.Collections.IDictionary]) { $defaults[$key] = @{} }
    }
    if ($null -eq $defaults.RecentFolders -or $defaults.RecentFolders -is [string]) {
        $defaults.RecentFolders = @($defaults.RecentFolders | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    if ([string]::IsNullOrWhiteSpace($defaults.QbittorrentWebUrl)) { $defaults.QbittorrentWebUrl = 'http://localhost:8080' }
    $Global:UserSettings = $defaults
    $themes = Get-Variable -Name Themes -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ($themes -is [System.Collections.IDictionary]) {
        if (-not $themes.Contains($defaults.Theme)) { $defaults.Theme = 'Dracula' }
        $Global:ThemeSelection = $defaults.Theme
        $Global:CurrentTheme = $themes[$defaults.Theme]
    }
    $Global:SettingsBaseline = ConvertTo-SettingsHashtable $defaults
    $Global:SettingsBaselinePath = [IO.Path]::GetFullPath($Global:ConfigFile)
    if ($migratePassword) { $null = Save-UserSettings }
}

function Merge-SettingsChanges {
    param([System.Collections.IDictionary]$Latest, [System.Collections.IDictionary]$Baseline, [System.Collections.IDictionary]$Current)
    $merged = ConvertTo-SettingsHashtable $Latest
    foreach ($key in $Baseline.Keys) { if (-not $Current.Contains($key)) { $merged.Remove($key) } }
    foreach ($key in $Current.Keys) {
        if ($Current[$key] -is [System.Collections.IDictionary]) {
            $old = if ($Baseline[$key] -is [System.Collections.IDictionary]) { $Baseline[$key] } else { @{} }
            $disk = if ($merged[$key] -is [System.Collections.IDictionary]) { $merged[$key] } else { @{} }
            $merged[$key] = Merge-SettingsChanges -Latest $disk -Baseline $old -Current $Current[$key]
        }
        elseif (-not $Baseline.Contains($key) -or (ConvertTo-Json -InputObject $Current[$key] -Depth 32 -Compress) -cne (ConvertTo-Json -InputObject $Baseline[$key] -Depth 32 -Compress)) {
            if ($key -eq 'RecentFolders') {
                $removed = @($Baseline[$key] | Where-Object { $_ -notin $Current[$key] })
                $folders = New-Object 'System.Collections.Generic.List[string]'
                foreach ($entry in @($Current[$key]) + @($merged[$key])) {
                    if ($entry -and $entry -notin $removed -and $entry -notin $folders -and $folders.Count -lt 12) { $folders.Add($entry) }
                }
                $merged[$key] = $folders.ToArray()
            } else { $merged[$key] = $Current[$key] }
        }
    }
    return $merged
}

function Save-UserSettings {
    $mutex = $null
    $locked = $false
    $temporary = $null
    try {
        if ([string]::IsNullOrWhiteSpace($Global:ConfigFile)) { $Global:ConfigFile = Join-Path (Get-AppDataDirectory) 'config.json' }
        $path = [IO.Path]::GetFullPath($Global:ConfigFile)
        [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path)) | Out-Null
        $mutex = New-Object Threading.Mutex($false, ('Local\qBitLauncher.Settings.' + (Get-ServiceKeyHash $path.ToLowerInvariant())))
        try { $locked = $mutex.WaitOne(5000) }
        catch [Threading.AbandonedMutexException] { $locked = $true }
        if (-not $locked) { throw 'Settings are busy in another qBitLauncher process.' }
        $saved = ConvertTo-SettingsHashtable $Global:UserSettings
        if ($saved -isnot [System.Collections.IDictionary]) { throw 'No settings object is available.' }
        $password = [string]$saved.QbittorrentPassword
        $saved.Remove('QbittorrentPassword')
        $saved.Remove('QbittorrentPasswordProtected')
        $latest = @{}
        if ([IO.File]::Exists($path)) { $latest = ConvertTo-SettingsHashtable ([IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json) }
        $hasBaseline = $Global:SettingsBaseline -is [System.Collections.IDictionary] -and $Global:SettingsBaselinePath -eq $path
        if ($hasBaseline) {
            $baseline = ConvertTo-SettingsHashtable $Global:SettingsBaseline
            $baseline.Remove('QbittorrentPassword'); $baseline.Remove('QbittorrentPasswordProtected')
            $saved = Merge-SettingsChanges -Latest $latest -Baseline $baseline -Current $saved
        }
        $saved.Remove('QbittorrentPassword')
        if ($hasBaseline -and $password -ceq [string]$Global:SettingsBaseline.QbittorrentPassword -and $latest.Contains('QbittorrentPasswordProtected')) {
            $saved.QbittorrentPasswordProtected = $latest.QbittorrentPasswordProtected
        }
        elseif (-not [string]::IsNullOrEmpty($password)) {
            $saved.QbittorrentPasswordProtected = Protect-AppSecret $password
            $Global:PasswordDecryptionFailed = $false
        }
        elseif ($Global:PasswordDecryptionFailed -and -not [string]::IsNullOrEmpty($Global:ProtectedSettingsPassword)) {
            # Moving portable settings to a different Windows account must not erase ciphertext.
            $saved.QbittorrentPasswordProtected = $Global:ProtectedSettingsPassword
        }
        else { $saved.QbittorrentPasswordProtected = '' }
        $temporary = Join-Path ([IO.Path]::GetDirectoryName($path)) ('.qbit-settings-' + [Guid]::NewGuid().ToString('N') + '.tmp')
        [IO.File]::WriteAllText($temporary, ($saved | ConvertTo-Json -Depth 32), (New-Object Text.UTF8Encoding($false)))
        if ([IO.File]::Exists($path)) { [IO.File]::Replace($temporary, $path, [NullString]::Value) }
        else { [IO.File]::Move($temporary, $path) }
        $Global:ProtectedSettingsPassword = $saved.QbittorrentPasswordProtected
        $Global:SettingsBaseline = ConvertTo-SettingsHashtable $Global:UserSettings
        $Global:SettingsBaselinePath = $path
        return $true
    }
    catch {
        Write-LogMessage ('Settings could not be saved. ' + $_.Exception.GetType().Name)
        return $false
    }
    finally {
        if ($temporary -and [IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
        if ($locked -and $mutex) { $mutex.ReleaseMutex() }
        if ($mutex) { $mutex.Dispose() }
    }
}

function Protect-LogText {
    param([AllowEmptyString()][string]$Text)
    if ($null -eq $Text) { return '' }
    $settings = Get-Variable -Name UserSettings -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ($settings -is [System.Collections.IDictionary]) {
        foreach ($key in @('QbittorrentPassword', 'QbittorrentPasswordProtected', 'QbittorrentUsername')) {
            $secret = [string]$settings[$key]
            if (-not [string]::IsNullOrEmpty($secret)) { $Text = $Text.Replace($secret, '[redacted]') }
        }
    }
    $Text = [regex]::Replace($Text, '(?i)(password|token|authorization)(\s*[:=]\s*)([^\r\n,;]+)', '$1$2[redacted]')
    $Text = [regex]::Replace($Text, '(?i)(https?://)[^/@\s]+:[^/@\s]+@', '$1[redacted]@')
    return $Text
}

function Get-AppLogPath {
    $override = Get-Variable -Name LogFile -ValueOnly -ErrorAction SilentlyContinue
    if (-not [string]::IsNullOrWhiteSpace($override)) { return [IO.Path]::GetFullPath($override) }
    return Join-Path (Join-Path (Get-AppDataDirectory) 'logs') 'qBitLauncher.log'
}

function Write-LogMessage {
    param([AllowEmptyString()][string]$Message)
    $mutex = $null
    $locked = $false
    try {
        $path = Get-AppLogPath
        [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path)) | Out-Null
        $safeMessage = Protect-LogText $Message
        if ($safeMessage.Length -gt 32768) { $safeMessage = $safeMessage.Substring(0, 32768) + ' [truncated]' }
        $entry = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss.fff') + ' [' + $PID + '] ' + $safeMessage + [Environment]::NewLine
        $encoding = New-Object Text.UTF8Encoding($false)
        $maxBytes = 2MB
        $override = Get-Variable -Name LogMaxBytes -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        if ($override -and [long]$override -ge 128) { $maxBytes = [long]$override }
        $mutex = New-Object Threading.Mutex($false, ('Local\qBitLauncher.Log.' + (Get-ServiceKeyHash $path.ToLowerInvariant())))
        try { $locked = $mutex.WaitOne(1500) }
        catch [Threading.AbandonedMutexException] { $locked = $true }
        if (-not $locked) { return }
        if ([IO.File]::Exists($path) -and ((Get-Item -LiteralPath $path).Length + $encoding.GetByteCount($entry)) -gt $maxBytes) {
            if ([IO.File]::Exists($path + '.3')) { [IO.File]::Delete($path + '.3') }
            for ($archive = 2; $archive -ge 1; $archive--) {
                if ([IO.File]::Exists($path + '.' + $archive)) { [IO.File]::Move($path + '.' + $archive, $path + '.' + ($archive + 1)) }
            }
            [IO.File]::Move($path, $path + '.1')
        }
        [IO.File]::AppendAllText($path, $entry, $encoding)
    }
    catch { } # Logging must never block an app action with an error dialog.
    finally {
        if ($locked -and $mutex) { $mutex.ReleaseMutex() }
        if ($mutex) { $mutex.Dispose() }
    }
}

function Get-PreferencePathKey {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-NormalizedServicePath $Path).ToLowerInvariant()
}

function Get-NormalizedServicePath {
    param([Parameter(Mandatory = $true)][string]$Path)
    $fullPath = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.Length -gt $root.Length) { $fullPath = $fullPath.TrimEnd([char]92, [char]47) }
    return $fullPath
}

function Get-FilePreference {
    param([Parameter(Mandatory = $true)][string]$Path)
    $key = Get-PreferencePathKey $Path
    if ($Global:UserSettings.FilePreferences.Contains($key)) { return ConvertTo-SettingsHashtable $Global:UserSettings.FilePreferences[$key] }
    return @{}
}

function Set-FilePreference {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][hashtable]$Preference)
    $key = Get-PreferencePathKey $Path
    $merged = Get-FilePreference $Path
    foreach ($field in $Preference.Keys) { $merged[$field] = $Preference[$field] }
    $Global:UserSettings.FilePreferences[$key] = $merged
    return Save-UserSettings
}

function Remove-FilePreference {
    param([Parameter(Mandatory = $true)][string]$Path)
    $Global:UserSettings.FilePreferences.Remove((Get-PreferencePathKey $Path))
    return Save-UserSettings
}

function Add-RecentFolder {
    param([Parameter(Mandatory = $true)][string]$Path)
    $fullPath = Get-NormalizedServicePath $Path
    $folders = New-Object System.Collections.Generic.List[string]
    $folders.Add($fullPath)
    foreach ($entry in $Global:UserSettings.RecentFolders) {
        if (-not [string]::IsNullOrWhiteSpace($entry) -and -not [string]::Equals($entry, $fullPath, [StringComparison]::OrdinalIgnoreCase) -and $folders.Count -lt 12) {
            $folders.Add($entry)
        }
    }
    $Global:UserSettings.RecentFolders = $folders.ToArray()
    return Save-UserSettings
}

function Remove-RecentFolder {
    param([Parameter(Mandatory = $true)][string]$Path)
    $Global:UserSettings.RecentFolders = @($Global:UserSettings.RecentFolders | Where-Object { -not [string]::Equals($_, $Path, [StringComparison]::OrdinalIgnoreCase) })
    return Save-UserSettings
}

function New-WorkflowLock {
    param([Parameter(Mandatory = $true)][string]$Key)
    # Only the same source is exclusive. Different games never share a slot.
    $directory = Join-Path (Get-AppDataDirectory) 'workflows'
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $normalizedKey = $Key.Trim().ToLowerInvariant()
    if ([IO.Path]::IsPathRooted($Key)) { $normalizedKey = Get-PreferencePathKey $Key }
    $lockPath = Join-Path $directory ((Get-ServiceKeyHash $normalizedKey) + '.lock')
    $lock = [pscustomobject]@{ Duplicate = $false; Handle = $null; IsDisposed = $false }
    try { $lock.Handle = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
    catch [IO.IOException] { $lock.Duplicate = $true }
    return $lock
}

function Exit-WorkflowLock {
    param([AllowNull()]$Lock)
    if ($null -eq $Lock -or $Lock.IsDisposed) { return }
    if ($Lock.Handle) { $Lock.Handle.Dispose(); $Lock.Handle = $null }
    $Lock.IsDisposed = $true
    # Keep the empty file: removing its path could race the next owner.
}
