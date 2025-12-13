# qBitLauncher_v3.ps1

param(
    [string]$filePathFromQB 
)

# -------------------------
# GLOBAL INITIALIZATION
# -------------------------
# Load .NET assemblies at the start to make their types available globally.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide the PowerShell console window (show only GUI)
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consoleWindow = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consoleWindow, 0) | Out-Null  # 0 = SW_HIDE

# -------------------------
# Configuration
# -------------------------
# Log file in script folder for easy access
$LogFile = Join-Path $PSScriptRoot "qBitLauncher_log.txt"
$ArchiveExtensions = @('iso', 'zip', 'rar', '7z', 'img')
$MediaExtensions = @('mp4', 'mkv', 'avi', 'mov', 'wmv', 'flv', 'webm', 'mp3', 'flac', 'wav', 'aac', 'ogg', 'm4a')
$ProcessableExtensions = @('exe') + $ArchiveExtensions

# -------------------------
# qBittorrent Web API Configuration (for cleanup feature)
# -------------------------
$Global:QBitConfig = @{
    Enabled  = $true
    BaseUrl  = "http://localhost:8080"
    Username = ""  # Leave empty if "Bypass auth for localhost" is enabled
    Password = ""  # Leave empty if "Bypass auth for localhost" is enabled
}

# Store the original file path for cleanup feature
$Global:OriginalFilePath = $null

# -------------------------
# GUI: Theme and Color Definitions
# -------------------------
# Set the desired theme here: 'qBitDark', 'Dark', or 'Light'
$Global:ThemeSelection = 'qBitDark' 
$Global:Themes = @{
    Light    = @{
        FormBack    = [System.Drawing.Color]::FromArgb(240, 240, 240)
        TextFore    = [System.Drawing.Color]::Black
        ControlBack = [System.Drawing.Color]::White
        ButtonBack  = [System.Drawing.Color]::FromArgb(225, 225, 225)
        Border      = [System.Drawing.Color]::DimGray
        Accent      = [System.Drawing.Color]::DodgerBlue
    }
    Dark     = @{
        FormBack    = [System.Drawing.Color]::FromArgb(45, 45, 48)
        TextFore    = [System.Drawing.Color]::White
        ControlBack = [System.Drawing.Color]::FromArgb(30, 30, 30)
        ButtonBack  = [System.Drawing.Color]::FromArgb(63, 63, 70)
        Border      = [System.Drawing.Color]::FromArgb(85, 85, 91)
        Accent      = [System.Drawing.Color]::FromArgb(0, 122, 204)
    }
    qBitDark = @{
        FormBack    = [System.Drawing.Color]::FromArgb(45, 49, 58)  # Main dark background
        TextFore    = [System.Drawing.Color]::FromArgb(220, 220, 220) # Off-white text
        ControlBack = [System.Drawing.Color]::FromArgb(34, 38, 46)  # Slightly lighter control background
        ButtonBack  = [System.Drawing.Color]::FromArgb(59, 64, 75)  # Button background
        Border      = [System.Drawing.Color]::FromArgb(80, 85, 96)  # Border for buttons/controls
        Accent      = [System.Drawing.Color]::FromArgb(80, 128, 175) # Blue from progress bar
    }
}
$Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]

# -------------------------
# Helper: Logging (Verb-Noun: Write-LogMessage)
# -------------------------
function Write-LogMessage {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp - $Message"
    try { Add-Content -Path $LogFile -Value $LogEntry -ErrorAction Stop }
    catch {
        $FallbackLogDir = Join-Path $env:PUBLIC "Documents"; $FallbackLogFile = Join-Path $FallbackLogDir "qBitLauncher_fallback_log.txt"
        try { if (-not (Test-Path $FallbackLogDir)) { New-Item -ItemType Directory -Path $FallbackLogDir -Force -ErrorAction SilentlyContinue | Out-Null }; Add-Content -Path $FallbackLogFile -Value "$Timestamp - FALLBACK: $Message (Original log failed: $($_.Exception.Message))" -ErrorAction SilentlyContinue } catch {}
        Write-Warning "Failed to write to primary log file: $LogFile. Error: $($_.Exception.Message)"
    }
}

Write-LogMessage "--------------------------------------------------------"
Write-LogMessage "Script started: qBitLauncher.ps1"
Write-LogMessage "Received initial path from qBittorrent: '$filePathFromQB'"
Write-Host "qBitLauncher.ps1 started. Log file: $LogFile"

# -------------------------
# Configuration File System
# -------------------------
$Global:ConfigFile = Join-Path $PSScriptRoot "config.json"
$Global:UserSettings = @{
    Theme        = "qBitDark"
    SoundEnabled = $true
}

function Get-UserSettings {
    if (Test-Path $Global:ConfigFile) {
        try {
            $json = Get-Content $Global:ConfigFile -Raw | ConvertFrom-Json
            $Global:UserSettings.Theme = if ($json.Theme) { $json.Theme } else { "qBitDark" }
            $Global:UserSettings.SoundEnabled = if ($null -ne $json.SoundEnabled) { $json.SoundEnabled } else { $true }
            Write-LogMessage "Loaded settings from config.json"
        }
        catch {
            Write-LogMessage "Failed to load config.json, using defaults: $($_.Exception.Message)"
        }
    }
    # Apply theme from settings
    $Global:ThemeSelection = $Global:UserSettings.Theme
    $Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]
}

function Save-UserSettings {
    try {
        $Global:UserSettings | ConvertTo-Json | Set-Content $Global:ConfigFile -Encoding UTF8
        Write-LogMessage "Settings saved to config.json"
        return $true
    }
    catch {
        Write-LogMessage "Failed to save settings: $($_.Exception.Message)"
        return $false
    }
}

# Load settings on startup
Get-UserSettings

# -------------------------
# Helper: Play Sound Effect
# -------------------------
function Play-ActionSound {
    param(
        [ValidateSet('Success', 'Error', 'Notify')]
        [string]$Type = 'Success'
    )
    
    if (-not $Global:UserSettings.SoundEnabled) { return }
    
    try {
        switch ($Type) {
            'Success' { [System.Media.SystemSounds]::Asterisk.Play() }
            'Error' { [System.Media.SystemSounds]::Exclamation.Play() }
            'Notify' { [System.Media.SystemSounds]::Beep.Play() }
        }
    }
    catch {
        Write-LogMessage "Failed to play sound: $($_.Exception.Message)"
    }
}

# -------------------------
# Helper: Show Toast Notification (NEW)
# -------------------------
function Show-ToastNotification {
    param(
        [string]$Title = "qBitLauncher",
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )
    
    try {
        # Use Windows built-in notification via .NET
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
        
        $template = @"
<toast>
    <visual>
        <binding template="ToastText02">
            <text id="1">$Title</text>
            <text id="2">$Message</text>
        </binding>
    </visual>
</toast>
"@
        
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("qBitLauncher").Show($toast)
        Write-LogMessage "Toast notification shown: $Title - $Message"
    }
    catch {
        # Fallback to balloon tip if toast fails
        try {
            $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
            $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
            $notifyIcon.Visible = $true
            
            $iconType = switch ($Type) {
                'Warning' { [System.Windows.Forms.ToolTipIcon]::Warning }
                'Error' { [System.Windows.Forms.ToolTipIcon]::Error }
                default { [System.Windows.Forms.ToolTipIcon]::Info }
            }
            
            $notifyIcon.ShowBalloonTip(3000, $Title, $Message, $iconType)
            Start-Sleep -Milliseconds 3500
            $notifyIcon.Dispose()
            Write-LogMessage "Balloon notification shown: $Title - $Message"
        }
        catch {
            Write-LogMessage "Failed to show notification: $($_.Exception.Message)"
        }
    }
}

# -------------------------
# Helper: Find WinRAR.exe (Verb-Noun: Get-WinRARPath)
# -------------------------
function Get-WinRARPath {
    foreach ($path in @((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe' -ErrorAction SilentlyContinue).'(Default)', "$env:ProgramFiles\WinRAR\WinRAR.exe", "$env:ProgramFiles(x86)\WinRAR\WinRAR.exe")) {
        if ($path -and (Test-Path $path)) { Write-LogMessage "Found WinRAR at: $path"; return $path }
    }
    Write-LogMessage "WinRAR not found."; return $null
}

# -------------------------
# Helper: Find 7-Zip.exe (NEW - Verb-Noun: Get-7ZipPath)
# -------------------------
function Get-7ZipPath {
    foreach ($path in @(
            (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\7zFM.exe' -ErrorAction SilentlyContinue).'(Default)',
            "$env:ProgramFiles\7-Zip\7z.exe",
            "$env:ProgramFiles(x86)\7-Zip\7z.exe"
        )) {
        if ($path) {
            # If we found 7zFM.exe path, convert to 7z.exe (command line version)
            $cmdPath = $path -replace '7zFM\.exe$', '7z.exe'
            if (Test-Path $cmdPath) {
                Write-LogMessage "Found 7-Zip at: $cmdPath"
                return $cmdPath
            }
            if (Test-Path $path) {
                Write-LogMessage "Found 7-Zip at: $path"
                return $path
            }
        }
    }
    Write-LogMessage "7-Zip not found."; return $null
}

# -------------------------
# Helper: Select Extraction Path (Verb-Noun: Select-ExtractionPath)
# -------------------------
function Select-ExtractionPath {
    param([string]$DefaultPath)
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select extraction destination folder"
    $folderBrowser.SelectedPath = $DefaultPath
    $folderBrowser.ShowNewFolderButton = $true
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-LogMessage "User selected extraction path: $($folderBrowser.SelectedPath)"
        return $folderBrowser.SelectedPath
    }
    Write-LogMessage "User cancelled folder selection."
    return $null
}

# -------------------------
# Helper: Extract Icon from Executable (Verb-Noun: Get-ExecutableIcon)
# -------------------------
function Get-ExecutableIcon {
    param([string]$ExePath)
    try {
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ExePath)
        if ($icon) {
            return $icon.ToBitmap()
        }
    }
    catch {
        Write-LogMessage "Failed to extract icon from '$ExePath': $($_.Exception.Message)"
    }
    return $null
}

# -------------------------
# Helper: Show Extraction Progress Form (Verb-Noun: Show-ExtractionProgress)
# -------------------------
function Show-ExtractionProgress {
    param(
        [string]$ArchiveName,
        [scriptblock]$OnCancel = $null
    )
    $colors = $Global:CurrentTheme
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Extracting..."
    $form.Size = New-Object System.Drawing.Size(450, 150)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.TopMost = $true
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 15)
    $label.Size = New-Object System.Drawing.Size(400, 25)
    $label.Text = "Extracting: $ArchiveName"
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)
    
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 50)
    $progressBar.Size = New-Object System.Drawing.Size(400, 25)
    $progressBar.Style = 'Marquee'
    $progressBar.MarqueeAnimationSpeed = 30
    $form.Controls.Add($progressBar)
    
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 85)
    $statusLabel.Size = New-Object System.Drawing.Size(400, 20)
    $statusLabel.Text = "Please wait..."
    $statusLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($statusLabel)
    
    # Store references for external access
    $form.Tag = @{
        ProgressBar = $progressBar
        StatusLabel = $statusLabel
        MainLabel   = $label
    }
    
    return $form
}

# -------------------------
# Helper: Update Progress Form (Verb-Noun: Update-ProgressForm)
# -------------------------
function Update-ProgressForm {
    param(
        [System.Windows.Forms.Form]$Form,
        [int]$Percentage = -1,
        [string]$Status = $null
    )
    if (-not $Form -or $Form.IsDisposed) { return }
    
    $controls = $Form.Tag
    if ($Percentage -ge 0 -and $Percentage -le 100) {
        $controls.ProgressBar.Style = 'Continuous'
        $controls.ProgressBar.Value = $Percentage
    }
    if ($Status) {
        $controls.StatusLabel.Text = $Status
    }
    $Form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

# -------------------------
# qBittorrent API: Authenticate (Verb-Noun: Connect-QBittorrent)
# -------------------------
function Connect-QBittorrent {
    if (-not $Global:QBitConfig.Enabled) { return $null }
    
    $baseUrl = $Global:QBitConfig.BaseUrl.TrimEnd('/')
    
    try {
        # First try without authentication (bypass mode)
        $testResponse = Invoke-RestMethod -Uri "$baseUrl/api/v2/app/version" -Method Get -SessionVariable session -ErrorAction Stop
        Write-LogMessage "Connected to qBittorrent (bypass auth mode). Version: $testResponse"
        return $session
    }
    catch {
        # Try with credentials if bypass failed
        if ($Global:QBitConfig.Username -and $Global:QBitConfig.Password) {
            try {
                $loginUrl = "$baseUrl/api/v2/auth/login"
                $body = @{
                    username = $Global:QBitConfig.Username
                    password = $Global:QBitConfig.Password
                }
                $response = Invoke-WebRequest -Uri $loginUrl -Method Post -Body $body -SessionVariable session -ErrorAction Stop
                if ($response.Content -eq "Ok.") {
                    Write-LogMessage "Connected to qBittorrent with credentials."
                    return $session
                }
            }
            catch {
                Write-LogMessage "qBittorrent login failed: $($_.Exception.Message)"
            }
        }
        Write-LogMessage "Failed to connect to qBittorrent: $($_.Exception.Message)"
    }
    return $null
}

# -------------------------
# qBittorrent API: Find Torrent by Path (Verb-Noun: Find-TorrentByPath)
# -------------------------
function Find-TorrentByPath {
    param(
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session,
        [string]$FilePath
    )
    if (-not $Session) { return $null }
    
    $baseUrl = $Global:QBitConfig.BaseUrl.TrimEnd('/')
    
    try {
        $torrents = Invoke-RestMethod -Uri "$baseUrl/api/v2/torrents/info" -Method Get -WebSession $Session -ErrorAction Stop
        
        foreach ($torrent in $torrents) {
            $torrentPath = Join-Path $torrent.save_path $torrent.name
            # Check if the file path starts with or matches the torrent path
            if ($FilePath -like "$torrentPath*" -or $torrent.content_path -eq $FilePath) {
                Write-LogMessage "Found matching torrent: $($torrent.name) (Hash: $($torrent.hash))"
                return @{
                    Hash        = $torrent.hash
                    Name        = $torrent.name
                    SavePath    = $torrent.save_path
                    ContentPath = $torrent.content_path
                }
            }
        }
        Write-LogMessage "No matching torrent found for path: $FilePath"
    }
    catch {
        Write-LogMessage "Failed to get torrent list: $($_.Exception.Message)"
    }
    return $null
}

# -------------------------
# qBittorrent API: Delete Torrent (Verb-Noun: Remove-TorrentFromClient)
# -------------------------
function Remove-TorrentFromClient {
    param(
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session,
        [string]$Hash,
        [bool]$DeleteFiles = $true
    )
    if (-not $Session -or -not $Hash) { return $false }
    
    $baseUrl = $Global:QBitConfig.BaseUrl.TrimEnd('/')
    
    try {
        $deleteUrl = "$baseUrl/api/v2/torrents/delete"
        $body = @{
            hashes      = $Hash
            deleteFiles = $DeleteFiles.ToString().ToLower()
        }
        Invoke-RestMethod -Uri $deleteUrl -Method Post -Body $body -WebSession $Session -ErrorAction Stop
        Write-LogMessage "Deleted torrent with hash: $Hash (deleteFiles: $DeleteFiles)"
        return $true
    }
    catch {
        Write-LogMessage "Failed to delete torrent: $($_.Exception.Message)"
    }
    return $false
}

# -------------------------
# Helper: Extract Archive (Verb-Noun: Expand-ArchiveFile)
# -------------------------
function Expand-ArchiveFile {
    param(
        [string]$ArchivePath, 
        [string]$DestinationPath  # Now accepts direct destination path
    )
    $ArchiveType = [IO.Path]::GetExtension($ArchivePath).TrimStart('.').ToLowerInvariant()
    $archiveName = [IO.Path]::GetFileName($ArchivePath)
    Write-LogMessage "Attempting to extract '${ArchivePath}' (Type: ${ArchiveType})"
    Write-Host "Attempting to extract '${ArchivePath}'..."
    Write-LogMessage "Target extraction directory: '$DestinationPath'"

    if (-not (Test-Path -LiteralPath $DestinationPath)) {
        try { 
            New-Item -ItemType Directory -Path $DestinationPath -ErrorAction Stop | Out-Null
            Write-LogMessage "Created extraction directory: '$DestinationPath'" 
        } 
        catch { 
            $errMsg = "Failed to create extraction directory: '$DestinationPath'. Error: $($_.Exception.Message)"
            Write-Error $errMsg
            Write-LogMessage "ERROR: $errMsg"
            return $null 
        }
    }
    else { Write-LogMessage "Extraction directory '$DestinationPath' already exists." }

    # Show progress window
    $progressForm = Show-ExtractionProgress -ArchiveName $archiveName
    $progressForm.Show()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Try native PowerShell for ZIP first
        if ($ArchiveType -eq 'zip') {
            try { 
                Update-ProgressForm -Form $progressForm -Status "Using native PowerShell..."
                Write-Host "Using native PowerShell to extract ZIP..."
                Expand-Archive -LiteralPath $ArchivePath -DestinationPath $DestinationPath -Force -ErrorAction Stop
                $progressForm.Close()
                $progressForm.Dispose()
                Write-Host "ZIP extracted successfully."
                Write-LogMessage "ZIP extracted with Expand-Archive."
                Show-ToastNotification -Title "Extraction Complete" -Message "ZIP extracted to: $DestinationPath" -Type Info
                return $DestinationPath 
            } 
            catch { 
                Write-Warning "Native ZIP extraction failed. Trying other extractors..."
                Write-LogMessage "Native ZIP extraction failed. Trying other extractors." 
            }
        }
        
        # Try 7-Zip first (more common and handles more formats)
        $sevenZip = Get-7ZipPath
        if ($sevenZip) {
            Update-ProgressForm -Form $progressForm -Status "Extracting with 7-Zip..."
            Write-Host "Extracting with 7-Zip..."
            
            # Use -bsp1 for progress output
            $processArgs = "x `"$ArchivePath`" -o`"$DestinationPath`" -y -bsp1"
            
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $sevenZip
            $psi.Arguments = $processArgs
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.CreateNoWindow = $true
            
            $process = [System.Diagnostics.Process]::Start($psi)
            
            # Read output and update progress
            while (-not $process.HasExited) {
                $line = $process.StandardOutput.ReadLine()
                if ($line -match '(\d+)%') {
                    $percent = [int]$Matches[1]
                    Update-ProgressForm -Form $progressForm -Percentage $percent -Status "$percent% complete"
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
            $process.WaitForExit()
            
            $progressForm.Close()
            $progressForm.Dispose()
            
            if ($process.ExitCode -eq 0) { 
                Write-Host "$($ArchiveType.ToUpper()) extracted successfully with 7-Zip to $DestinationPath"
                Write-LogMessage "Extracted with 7-Zip successfully."
                Show-ToastNotification -Title "Extraction Complete" -Message "$($ArchiveType.ToUpper()) extracted to: $DestinationPath" -Type Info
                return $DestinationPath 
            }
            else {
                Write-Warning "7-Zip extraction failed with exit code $($process.ExitCode). Trying WinRAR..."
                Write-LogMessage "7-Zip extraction failed. Exit Code: $($process.ExitCode). Falling back to WinRAR."
                # Reopen progress for WinRAR
                $progressForm = Show-ExtractionProgress -ArchiveName $archiveName
                $progressForm.Show()
            }
        }
        
        # Fallback to WinRAR
        $winrar = Get-WinRARPath
        if (-not $winrar) { 
            $progressForm.Close()
            $progressForm.Dispose()
            $errMsg = "No archive extractor found (tried 7-Zip and WinRAR). Cannot extract '${ArchivePath}'. Please install 7-Zip or WinRAR."
            Write-Error $errMsg
            Write-LogMessage "ERROR: $errMsg"
            Show-ToastNotification -Title "Extraction Failed" -Message "No extractor found. Install 7-Zip or WinRAR." -Type Error
            return $null 
        }

        Update-ProgressForm -Form $progressForm -Status "Extracting with WinRAR..."
        Write-Host "Extracting with WinRAR..."
        $processArgs = @('x', "`"$ArchivePath`"", "`"$DestinationPath\`"", '-y', '-o+')
        $process = Start-Process -FilePath $winrar -ArgumentList $processArgs -NoNewWindow -Wait -PassThru
        
        $progressForm.Close()
        $progressForm.Dispose()
        
        if ($process.ExitCode -ne 0) { 
            $warnMsg = "WinRAR extraction might have failed for '${ArchivePath}'. Exit Code: $($process.ExitCode)."
            Write-Warning $warnMsg
            Write-LogMessage "WARNING: $warnMsg"
            Show-ToastNotification -Title "Extraction Warning" -Message "WinRAR reported exit code $($process.ExitCode)" -Type Warning
            return $null 
        } 
        else { 
            Write-Host "$($ArchiveType.ToUpper()) extracted successfully with WinRAR to $DestinationPath"
            Write-LogMessage "Extracted with WinRAR successfully."
            Show-ToastNotification -Title "Extraction Complete" -Message "$($ArchiveType.ToUpper()) extracted to: $DestinationPath" -Type Info
            return $DestinationPath 
        }
    }
    catch { 
        if ($progressForm -and -not $progressForm.IsDisposed) {
            $progressForm.Close()
            $progressForm.Dispose()
        }
        $errMsg = "Error during extraction process for '${ArchivePath}'. Error: $($_.Exception.Message)"
        Write-Error $errMsg
        Write-LogMessage "ERROR: $errMsg"
        Show-ToastNotification -Title "Extraction Error" -Message $_.Exception.Message -Type Error
        return $null 
    }
}

# -------------------------
# Helper: Find ALL Executables (Verb-Noun: Get-AllExecutables)
# -------------------------
function Get-AllExecutables {
    param([string]$RootFolderPath)
    Write-LogMessage "Searching for all executables (.exe) in '$RootFolderPath' (depth-first sort)."; Write-Host "Searching for all .exe files in '$RootFolderPath' and its subfolders..."
    $allExecutables = Get-ChildItem -LiteralPath $RootFolderPath -Filter *.exe -File -Recurse -ErrorAction SilentlyContinue
    if ($allExecutables) {
        $sortedExecutables = $allExecutables | Sort-Object @{Expression = { ($_.FullName.Split([IO.Path]::DirectorySeparatorChar).Count) } }, FullName
        Write-LogMessage "Found $($sortedExecutables.Count) executables."; return $sortedExecutables
    }
    Write-LogMessage "No .exe files found in '$RootFolderPath'."; Write-Warning "No .exe files found in '$RootFolderPath' or its subdirectories."; return $null
}

# ---------------------------------------------------
# GUI: Extraction Confirmation Form with Custom Path
# ---------------------------------------------------
function Show-ExtractionConfirmForm {
    param(
        [string]$ArchivePath,
        [string]$DefaultDestination,
        [string]$Title = "Confirm Extraction"
    )
    $colors = $Global:CurrentTheme
    
    # Result object
    $result = @{
        Confirmed       = $false
        DestinationPath = $DefaultDestination
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(700, 250)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    # Message label
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 15)
    $label.Size = New-Object System.Drawing.Size(660, 40)
    $label.Text = "An archive file was found. Proceed with extraction?`nFile: $([IO.Path]::GetFileName($ArchivePath))"
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)

    # Destination label
    $destLabel = New-Object System.Windows.Forms.Label
    $destLabel.Location = New-Object System.Drawing.Point(20, 65)
    $destLabel.Size = New-Object System.Drawing.Size(100, 25)
    $destLabel.Text = "Extract to:"
    $destLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($destLabel)

    # Destination textbox
    $destTextBox = New-Object System.Windows.Forms.TextBox
    $destTextBox.Location = New-Object System.Drawing.Point(120, 63)
    $destTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $destTextBox.Text = $DefaultDestination
    $destTextBox.BackColor = $colors.ControlBack
    $destTextBox.ForeColor = $colors.TextFore
    $destTextBox.BorderStyle = 'FixedSingle'
    $form.Controls.Add($destTextBox)

    # Browse button
    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(580, 61)
    $browseButton.Size = New-Object System.Drawing.Size(90, 28)
    $browseButton.Text = "&Browse..."
    $browseButton.BackColor = $colors.ButtonBack
    $browseButton.ForeColor = $colors.TextFore
    $browseButton.FlatStyle = 'Flat'
    $browseButton.FlatAppearance.BorderSize = 1
    $browseButton.FlatAppearance.BorderColor = $colors.Accent
    $browseButton.Add_Click({
            $selected = Select-ExtractionPath -DefaultPath $destTextBox.Text
            if ($selected) {
                $destTextBox.Text = $selected
            }
        })
    $form.Controls.Add($browseButton)

    # Info label
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(20, 100)
    $infoLabel.Size = New-Object System.Drawing.Size(660, 40)
    $infoLabel.Text = "Full archive path: $ArchivePath"
    $infoLabel.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150)
    $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($infoLabel)

    # Buttons
    $extractButton = New-Object System.Windows.Forms.Button
    $extractButton.Location = New-Object System.Drawing.Point(450, 160)
    $extractButton.Size = New-Object System.Drawing.Size(100, 35)
    $extractButton.Text = "&Extract"
    $extractButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(560, 160)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "&Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::No

    foreach ($button in @($extractButton, $cancelButton)) {
        $button.BackColor = $colors.ButtonBack
        $button.ForeColor = $colors.TextFore
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $colors.Accent
        $form.Controls.Add($button)
    }
    
    $form.AcceptButton = $extractButton
    $form.CancelButton = $cancelButton

    $dialogResult = $form.ShowDialog()
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $result.Confirmed = $true
        $result.DestinationPath = $destTextBox.Text
    }
    
    $form.Dispose()
    return $result
}

# ---------------------------------------------------
# GUI: Settings Form
# ---------------------------------------------------
function Show-SettingsForm {
    $colors = $Global:CurrentTheme
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Settings - qBitLauncher"
    $form.Size = New-Object System.Drawing.Size(400, 250)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    # Theme label
    $themeLabel = New-Object System.Windows.Forms.Label
    $themeLabel.Location = New-Object System.Drawing.Point(20, 25)
    $themeLabel.Size = New-Object System.Drawing.Size(100, 25)
    $themeLabel.Text = "Theme:"
    $themeLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($themeLabel)

    # Theme dropdown
    $themeCombo = New-Object System.Windows.Forms.ComboBox
    $themeCombo.Location = New-Object System.Drawing.Point(130, 22)
    $themeCombo.Size = New-Object System.Drawing.Size(220, 25)
    $themeCombo.DropDownStyle = 'DropDownList'
    $themeCombo.BackColor = $colors.ControlBack
    $themeCombo.ForeColor = $colors.TextFore
    $themeCombo.FlatStyle = 'Flat'
    [void]$themeCombo.Items.AddRange(@('qBitDark', 'Dark', 'Light'))
    $themeCombo.SelectedItem = $Global:UserSettings.Theme
    $form.Controls.Add($themeCombo)

    # Sound checkbox
    $soundCheckbox = New-Object System.Windows.Forms.CheckBox
    $soundCheckbox.Location = New-Object System.Drawing.Point(20, 70)
    $soundCheckbox.Size = New-Object System.Drawing.Size(300, 25)
    $soundCheckbox.Text = "Enable sound effects"
    $soundCheckbox.Checked = $Global:UserSettings.SoundEnabled
    $soundCheckbox.ForeColor = $colors.TextFore
    $soundCheckbox.FlatStyle = 'Flat'
    $form.Controls.Add($soundCheckbox)

    # Info label
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(20, 110)
    $infoLabel.Size = New-Object System.Drawing.Size(340, 40)
    $infoLabel.Text = "Theme changes apply to new windows.`nSettings are saved to config.json"
    $infoLabel.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150)
    $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($infoLabel)

    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(160, 160)
    $saveButton.Size = New-Object System.Drawing.Size(100, 35)
    $saveButton.Text = "&Save"
    $saveButton.Add_Click({
            $Global:UserSettings.Theme = $themeCombo.SelectedItem
            $Global:UserSettings.SoundEnabled = $soundCheckbox.Checked
            $Global:ThemeSelection = $Global:UserSettings.Theme
            $Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]
            if (Save-UserSettings) {
                Play-ActionSound -Type Success
                [System.Windows.Forms.MessageBox]::Show("Settings saved!", "Settings", 'OK', 'Information')
            }
            $form.Close()
        })

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(270, 160)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "&Cancel"
    $cancelButton.Add_Click({
            $form.Close()
        })

    foreach ($button in @($saveButton, $cancelButton)) {
        $button.BackColor = $colors.ButtonBack
        $button.ForeColor = $colors.TextFore
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $colors.Accent
        $form.Controls.Add($button)
    }
    
    $form.AcceptButton = $saveButton
    $form.CancelButton = $cancelButton

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# ---------------------------------------------------
# GUI: Main Executable Selection Form (with Icons)
# ---------------------------------------------------
function Show-ExecutableSelectionForm {
    param(
        [System.Management.Automation.PSObject[]]$FoundExecutables,
        [string]$WindowTitle = "qBitLauncher Action"
    )
    $colors = $Global:CurrentTheme

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(750, 450)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $colors.FormBack
    $font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Font = $font

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(710, 25)
    $label.Text = "Please select an executable and choose an action."
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)

    # Create ImageList for icons
    $imageList = New-Object System.Windows.Forms.ImageList
    $imageList.ImageSize = New-Object System.Drawing.Size(24, 24)
    $imageList.ColorDepth = 'Depth32Bit'

    # Create ListView instead of ListBox
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 40)
    $listView.Size = New-Object System.Drawing.Size(710, 300)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.MultiSelect = $false
    $listView.BackColor = $colors.ControlBack
    $listView.ForeColor = $colors.TextFore
    $listView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $listView.Font = $font
    $listView.SmallImageList = $imageList
    $listView.HeaderStyle = 'None'
    
    # Add column for the path
    $column = New-Object System.Windows.Forms.ColumnHeader
    $column.Width = 700
    [void]$listView.Columns.Add($column)
    
    # Add executables with icons
    $iconIndex = 0
    foreach ($exe in $FoundExecutables) {
        if ($exe -and $exe.FullName) {
            # Extract icon
            $icon = Get-ExecutableIcon -ExePath $exe.FullName
            if ($icon) {
                $imageList.Images.Add($icon)
                $item = New-Object System.Windows.Forms.ListViewItem($exe.FullName, $iconIndex)
                $iconIndex++
            }
            else {
                # Use default icon index -1 (no icon)
                $item = New-Object System.Windows.Forms.ListViewItem($exe.FullName)
            }
            $item.Tag = $exe.FullName
            [void]$listView.Items.Add($item)
        }
    }
    
    if ($listView.Items.Count -gt 0) {
        $listView.Items[0].Selected = $true
    }

    $form.Controls.Add($listView)

    # Track last action for return value
    $script:lastAction = 'None'

    # Helper to get selected executable
    $getSelectedExe = {
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select an executable first.", "No Selection", 'OK', 'Warning')
            return $null
        }
        $path = $listView.SelectedItems[0].Tag
        try {
            return Get-Item -LiteralPath $path -ErrorAction Stop
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Could not access: $path", "Error", 'OK', 'Error')
            return $null
        }
    }

    # Button row - Main actions (no DialogResult - form stays open)
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Location = New-Object System.Drawing.Point(10, 355)
    $runButton.Size = New-Object System.Drawing.Size(80, 35)
    $runButton.Text = "&Run"
    $runButton.Add_Click({
            $exe = & $getSelectedExe
            if ($exe) {
                try {
                    Start-Process -FilePath $exe.FullName -WorkingDirectory $exe.DirectoryName -Verb RunAs
                    Play-ActionSound -Type Success
                    Show-ToastNotification -Title "Launched" -Message "$($exe.Name)" -Type Info
                    Write-LogMessage "Launched: $($exe.FullName)"
                    [System.Windows.Forms.MessageBox]::Show("Launched: $($exe.Name)", "Application Started", 'OK', 'Information')
                }
                catch {
                    Play-ActionSound -Type Error
                    [System.Windows.Forms.MessageBox]::Show("Failed to launch: $($_.Exception.Message)", "Error", 'OK', 'Error')
                }
            }
        })
    
    $shortcutButton = New-Object System.Windows.Forms.Button
    $shortcutButton.Location = New-Object System.Drawing.Point(95, 355)
    $shortcutButton.Size = New-Object System.Drawing.Size(100, 35)
    $shortcutButton.Text = "&Shortcut"
    $shortcutButton.Add_Click({
            $exe = & $getSelectedExe
            if ($exe) {
                try {
                    $desktopPath = [System.Environment]::GetFolderPath('Desktop')
                    $shortcutName = $exe.BaseName + ".lnk"
                    $shortcutPath = Join-Path $desktopPath $shortcutName
                    $wshell = New-Object -ComObject WScript.Shell
                    $shortcut = $wshell.CreateShortcut($shortcutPath)
                    $shortcut.TargetPath = $exe.FullName
                    $shortcut.WorkingDirectory = $exe.DirectoryName
                    $shortcut.Save()
                    Play-ActionSound -Type Success
                    Show-ToastNotification -Title "Shortcut Created" -Message "$shortcutName on Desktop" -Type Info
                    Write-LogMessage "Shortcut created: $shortcutPath"
                    [System.Windows.Forms.MessageBox]::Show("Shortcut created on Desktop:`n$shortcutName", "Shortcut Created", 'OK', 'Information')
                }
                catch {
                    Play-ActionSound -Type Error
                    [System.Windows.Forms.MessageBox]::Show("Failed to create shortcut: $($_.Exception.Message)", "Error", 'OK', 'Error')
                }
            }
        })
    
    $exploreButton = New-Object System.Windows.Forms.Button
    $exploreButton.Location = New-Object System.Drawing.Point(200, 355)
    $exploreButton.Size = New-Object System.Drawing.Size(100, 35)
    $exploreButton.Text = "&Open Folder"
    $exploreButton.Add_Click({
            $exe = & $getSelectedExe
            if ($exe) {
                Start-Process explorer -ArgumentList "`"$($exe.DirectoryName)`""
                Play-ActionSound -Type Success
                Write-LogMessage "Opened folder: $($exe.DirectoryName)"
                [System.Windows.Forms.MessageBox]::Show("Folder opened in Explorer.", "Folder Opened", 'OK', 'Information')
            }
        })

    # Settings button
    $settingsButton = New-Object System.Windows.Forms.Button
    $settingsButton.Location = New-Object System.Drawing.Point(305, 355)
    $settingsButton.Size = New-Object System.Drawing.Size(100, 35)
    $settingsButton.Text = "Se&ttings"
    $settingsButton.Add_Click({
            $oldTheme = $Global:ThemeSelection
            Show-SettingsForm
            
            # If theme changed, refresh this form's colors
            if ($Global:ThemeSelection -ne $oldTheme) {
                $newColors = $Global:CurrentTheme
                $form.BackColor = $newColors.FormBack
                $label.ForeColor = $newColors.TextFore
                $listView.BackColor = $newColors.ControlBack
                $listView.ForeColor = $newColors.TextFore
                
                # Refresh all buttons
                foreach ($ctrl in $form.Controls) {
                    if ($ctrl -is [System.Windows.Forms.Button]) {
                        $ctrl.BackColor = $newColors.ButtonBack
                        $ctrl.ForeColor = $newColors.TextFore
                        $ctrl.FlatAppearance.BorderColor = $newColors.Accent
                    }
                }
                $form.Refresh()
            }
        })
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(620, 355)
    $closeButton.Size = New-Object System.Drawing.Size(100, 35)
    $closeButton.Text = "&Close"
    $closeButton.Add_Click({
            $form.Close()
        })

    foreach ($button in @($runButton, $shortcutButton, $exploreButton, $settingsButton, $closeButton)) {
        $button.BackColor = $colors.ButtonBack
        $button.ForeColor = $colors.TextFore
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $colors.Accent
        $form.Controls.Add($button)
    }

    $form.ActiveControl = $listView

    # Show form (blocks until closed)
    $form.ShowDialog() | Out-Null
    
    # Cleanup
    $imageList.Dispose()
    $form.Dispose()
}

# ---------------------------------------------------
# GUI: Cleanup Confirmation Form (with seeding message)
# ---------------------------------------------------
function Show-CleanupConfirmForm {
    param([string]$TorrentName)
    $colors = $Global:CurrentTheme
    
    $result = @{
        Confirmed     = $false
        RemoveTorrent = $true
        DeleteFiles   = $true
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Clean Up - qBitLauncher"
    $form.Size = New-Object System.Drawing.Size(550, 280)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    # Seeding message
    $seedingLabel = New-Object System.Windows.Forms.Label
    $seedingLabel.Location = New-Object System.Drawing.Point(20, 15)
    $seedingLabel.Size = New-Object System.Drawing.Size(500, 40)
    $seedingLabel.Text = "🌱 Seeding helps the community! But we understand if you need to save space."
    $seedingLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 180, 100)
    $seedingLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
    $form.Controls.Add($seedingLabel)

    # Torrent name
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Location = New-Object System.Drawing.Point(20, 60)
    $nameLabel.Size = New-Object System.Drawing.Size(500, 25)
    $nameLabel.Text = "Torrent: $TorrentName"
    $nameLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($nameLabel)

    # Checkbox: Remove torrent
    $removeTorrentCheckbox = New-Object System.Windows.Forms.CheckBox
    $removeTorrentCheckbox.Location = New-Object System.Drawing.Point(20, 100)
    $removeTorrentCheckbox.Size = New-Object System.Drawing.Size(300, 25)
    $removeTorrentCheckbox.Text = "Remove torrent from qBittorrent"
    $removeTorrentCheckbox.Checked = $true
    $removeTorrentCheckbox.ForeColor = $colors.TextFore
    $removeTorrentCheckbox.FlatStyle = 'Flat'
    $form.Controls.Add($removeTorrentCheckbox)

    # Checkbox: Delete files
    $deleteFilesCheckbox = New-Object System.Windows.Forms.CheckBox
    $deleteFilesCheckbox.Location = New-Object System.Drawing.Point(20, 130)
    $deleteFilesCheckbox.Size = New-Object System.Drawing.Size(300, 25)
    $deleteFilesCheckbox.Text = "Delete downloaded files"
    $deleteFilesCheckbox.Checked = $true
    $deleteFilesCheckbox.ForeColor = $colors.TextFore
    $deleteFilesCheckbox.FlatStyle = 'Flat'
    $form.Controls.Add($deleteFilesCheckbox)

    # Warning
    $warningLabel = New-Object System.Windows.Forms.Label
    $warningLabel.Location = New-Object System.Drawing.Point(20, 165)
    $warningLabel.Size = New-Object System.Drawing.Size(500, 25)
    $warningLabel.Text = "⚠️ This action cannot be undone!"
    $warningLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 180, 100)
    $form.Controls.Add($warningLabel)

    # Buttons
    $confirmButton = New-Object System.Windows.Forms.Button
    $confirmButton.Location = New-Object System.Drawing.Point(300, 200)
    $confirmButton.Size = New-Object System.Drawing.Size(100, 35)
    $confirmButton.Text = "&Confirm"
    $confirmButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(410, 200)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "&Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::No

    foreach ($button in @($confirmButton, $cancelButton)) {
        $button.BackColor = $colors.ButtonBack
        $button.ForeColor = $colors.TextFore
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $colors.Accent
        $form.Controls.Add($button)
    }
    
    $form.AcceptButton = $confirmButton
    $form.CancelButton = $cancelButton

    $dialogResult = $form.ShowDialog()
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $result.Confirmed = $true
        $result.RemoveTorrent = $removeTorrentCheckbox.Checked
        $result.DeleteFiles = $deleteFilesCheckbox.Checked
    }
    
    $form.Dispose()
    return $result
}

# ---------------------------------------------------
# Consolidated Action Handler (with Cleanup support)
# ---------------------------------------------------
function Invoke-UserAction {
    param(
        [hashtable]$GuiResult
    )
    
    $selectedExecutable = $GuiResult.SelectedExecutable
    
    switch ($GuiResult.DialogResult) {
        'OK' {
            # Run as Admin (default behavior)
            Write-LogMessage "User chose to run as ADMIN '$($selectedExecutable.FullName)'."
            Write-Host "Attempting to run as Administrator: '$($selectedExecutable.FullName)'..."
            try {
                Start-Process -FilePath $selectedExecutable.FullName -WorkingDirectory $selectedExecutable.DirectoryName -Verb RunAs
                Show-ToastNotification -Title "Launched" -Message "$($selectedExecutable.Name)" -Type Info
            }
            catch {
                $errMsg = "Error starting executable '$($selectedExecutable.FullName)': $($_.Exception.Message)"
                Write-Warning $errMsg
                Write-LogMessage "WARNING: $errMsg"
                Show-ToastNotification -Title "Launch Failed" -Message $errMsg -Type Error
            }
        }
        'Yes' {
            # Shortcut
            Write-LogMessage "User chose to create shortcut for '$($selectedExecutable.FullName)'."
            try {
                $desktopPath = [System.Environment]::GetFolderPath('Desktop')
                $shortcutName = $selectedExecutable.BaseName + ".lnk"
                $shortcutPath = Join-Path $desktopPath $shortcutName
                Write-LogMessage "Creating shortcut: '$shortcutPath'"
                $wshell = New-Object -ComObject WScript.Shell
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $selectedExecutable.FullName
                $shortcut.WorkingDirectory = $selectedExecutable.DirectoryName
                $shortcut.Save()
                Write-Host "Shortcut created on Desktop: $shortcutPath"
                Write-LogMessage "Shortcut created."
                Show-ToastNotification -Title "Shortcut Created" -Message "$shortcutName on Desktop" -Type Info
            }
            catch { 
                $errMsg = "Error creating shortcut: $($_.Exception.Message)"
                Write-Warning $errMsg
                Write-LogMessage "WARNING: $errMsg"
                Show-ToastNotification -Title "Shortcut Failed" -Message $errMsg -Type Error
            }
        }
        'Retry' {
            # Explore
            Write-LogMessage "User chose to open the folder."
            Write-Host "Opening folder."
            Start-Process explorer -ArgumentList "`"$($selectedExecutable.DirectoryName)`""
        }
        'Abort' {
            # Cleanup - Remove torrent and files from qBittorrent
            Write-LogMessage "User chose to clean up (remove torrent and files)."
            
            if (-not $Global:OriginalFilePath) {
                Write-Warning "Original file path not available for cleanup."
                Show-ToastNotification -Title "Cleanup Failed" -Message "Original file path not found" -Type Error
                return
            }
            
            # Connect to qBittorrent
            $session = Connect-QBittorrent
            if (-not $session) {
                Write-Warning "Could not connect to qBittorrent. Make sure Web UI is enabled."
                Show-ToastNotification -Title "Cleanup Failed" -Message "Cannot connect to qBittorrent Web UI" -Type Error
                return
            }
            
            # Find the torrent
            $torrent = Find-TorrentByPath -Session $session -FilePath $Global:OriginalFilePath
            if (-not $torrent) {
                Write-Warning "Could not find matching torrent in qBittorrent."
                Show-ToastNotification -Title "Cleanup Failed" -Message "Torrent not found in qBittorrent" -Type Warning
                return
            }
            
            # Show confirmation dialog
            $cleanupResult = Show-CleanupConfirmForm -TorrentName $torrent.Name
            
            if ($cleanupResult.Confirmed) {
                if ($cleanupResult.RemoveTorrent) {
                    $success = Remove-TorrentFromClient -Session $session -Hash $torrent.Hash -DeleteFiles $cleanupResult.DeleteFiles
                    if ($success) {
                        Write-Host "Torrent removed from qBittorrent."
                        $message = if ($cleanupResult.DeleteFiles) { "Torrent and files deleted" } else { "Torrent removed (files kept)" }
                        Show-ToastNotification -Title "Cleanup Complete" -Message $message -Type Info
                    }
                    else {
                        Show-ToastNotification -Title "Cleanup Failed" -Message "Could not remove torrent" -Type Error
                    }
                }
                else {
                    Write-Host "User chose not to remove the torrent."
                    Show-ToastNotification -Title "Cleanup Skipped" -Message "No changes made" -Type Info
                }
            }
            else {
                Write-Host "Cleanup cancelled by user."
            }
        }
        default {
            # Cancel or Closed
            Write-LogMessage "User cancelled or closed the selection window."
            Write-Host "Action cancelled."
        }
    }
}

# ===================================================================
# MAIN SCRIPT LOGIC STARTS HERE
# ===================================================================

if (-not (Test-Path -LiteralPath $filePathFromQB)) {
    $errMsg = "Error: Initial path not found - $filePathFromQB"
    Write-Error $errMsg
    Write-LogMessage "FATAL: $errMsg. Script exiting."
    Show-ToastNotification -Title "qBitLauncher Error" -Message "Path not found: $filePathFromQB" -Type Error
    Read-Host "Press Enter to exit..."
    exit 1
}

$mainFileToProcess = $null

if (Test-Path -LiteralPath $filePathFromQB -PathType Container) {
    $downloadFolder = $filePathFromQB
    Write-Host "Input path is a folder: '$downloadFolder'. Searching for primary file..."
    $mainFileToProcess = Get-ChildItem -LiteralPath $downloadFolder -File -Recurse | Where-Object { $ArchiveExtensions -contains $_.Extension.TrimStart('.').ToLowerInvariant() } | Sort-Object Length -Descending | Select-Object -First 1
    
    if ($mainFileToProcess) {
        Write-LogMessage "Found a primary archive file to process in folder: '$($mainFileToProcess.FullName)'"
    }
    else {
        Write-LogMessage "No archives found in folder. Searching for executables..."
        $allExecutables = Get-AllExecutables -RootFolderPath $downloadFolder
        if ($allExecutables) {
            $mainFileToProcess = $allExecutables
        }
        else {
            Write-LogMessage "No executables found. Checking for media files..."
            $foundMediaFile = Get-ChildItem -LiteralPath $downloadFolder -File -Recurse | Where-Object { $MediaExtensions -contains $_.Extension.TrimStart('.').ToLowerInvariant() } | Select-Object -First 1
            if ($foundMediaFile) {
                Write-Host "Found a media file: $($foundMediaFile.Name). Opening folder."
                Start-Process explorer -ArgumentList "`"$(Split-Path $foundMediaFile.FullName -Parent)`""
            }
            else {
                Write-Warning "No processable files found in '$downloadFolder'."
                Start-Process explorer -ArgumentList "`"$downloadFolder`""
            }
        }
    }
}
else {
    $mainFileToProcess = Get-Item -LiteralPath $filePathFromQB
    Write-LogMessage "Input is a single file: '$($mainFileToProcess.FullName)'"
}

if ($mainFileToProcess) {
    $firstFile = if ($mainFileToProcess -is [array]) { $mainFileToProcess[0] } else { $mainFileToProcess }
    $filePath = $firstFile.FullName
    $parentDir = $firstFile.DirectoryName
    $baseName = $firstFile.BaseName
    $ext = $firstFile.Extension.ToLowerInvariant().TrimStart('.')
    
    # Store original path for cleanup feature
    $Global:OriginalFilePath = $filePathFromQB

    if ($ArchiveExtensions -contains $ext) {
        Write-LogMessage "Processing archive: '$filePath'"
        Write-Host "`nFound an archive file: $filePath"
        
        # Default extraction path (beside the archive, in a subfolder named after archive)
        $defaultExtractPath = Join-Path $parentDir $baseName
        
        # Show extraction confirmation with custom path option
        $extractionResult = Show-ExtractionConfirmForm -ArchivePath $filePath -DefaultDestination $defaultExtractPath
        
        if ($extractionResult.Confirmed) {
            Write-LogMessage "User confirmed extraction to: $($extractionResult.DestinationPath)"
            $extractedDir = Expand-ArchiveFile -ArchivePath $filePath -DestinationPath $extractionResult.DestinationPath
            if ($extractedDir) {
                Write-Host "`nExtraction complete. Searching for executables..."
                $executablesInArchive = Get-AllExecutables -RootFolderPath $extractedDir
                if ($executablesInArchive) {
                    Show-ExecutableSelectionForm -FoundExecutables $executablesInArchive -WindowTitle "Archive Extracted"
                }
                else {
                    Write-Warning "No executables found in the extracted folder: $extractedDir"
                    Start-Process explorer -ArgumentList "`"$extractedDir`""
                }
            }
        }
        else {
            Write-LogMessage "User declined extraction. Opening folder: '$parentDir'"
            Start-Process explorer -ArgumentList "`"$parentDir`""
        }
    } 
    elseif ($ext -eq 'exe' -or $mainFileToProcess -is [array]) {
        $executables = if ($mainFileToProcess -is [array]) { $mainFileToProcess } else { @($mainFileToProcess) }
        Write-LogMessage "Processing one or more executables."
        
        Show-ExecutableSelectionForm -FoundExecutables $executables -WindowTitle "Executable Found"
    } 
    elseif ($MediaExtensions -contains $ext) {
        Write-LogMessage "File is a media file."
        Write-Host "Media file '${filePath}' is ready."
        Start-Process explorer -ArgumentList "`"$parentDir`""
        Write-Host "Opening containing folder: $parentDir"
    }
    else {
        Write-LogMessage "File is an unhandled type (.$ext)."
        Write-Warning "File type .${ext} is not handled explicitly."
        Start-Process explorer -ArgumentList "`"$parentDir`""
        Write-Host "Opening containing folder: $parentDir"
    }
}

Write-Host "`nScript actions complete."
Write-LogMessage "Script finished."
Write-LogMessage "--------------------------------------------------------`n"