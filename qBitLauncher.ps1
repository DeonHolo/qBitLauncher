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

# Set AppUserModelID for proper taskbar icon (separates from PowerShell)
Add-Type -Name Shell32 -Namespace Win32 -MemberDefinition '
[DllImport("shell32.dll", SetLastError = true)]
public static extern void SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);
'
[Win32.Shell32]::SetCurrentProcessExplicitAppUserModelID("qBitLauncher.App")

# -------------------------
# Configuration
# -------------------------
# Log file in script folder for easy access
$LogFile = Join-Path $PSScriptRoot "qBitLauncher_log.txt"
$ArchiveExtensions = @('iso', 'zip', 'rar', '7z', 'img')
$MediaExtensions = @('mp4', 'mkv', 'avi', 'mov', 'wmv', 'flv', 'webm', 'mp3', 'flac', 'wav', 'aac', 'ogg', 'm4a')

# Load app icon from embedded Base64 (no external file needed)
$Global:AppIcon = $null
$LogoBase64 = "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"
try {
    $iconBytes = [System.Convert]::FromBase64String($LogoBase64)
    $ms = New-Object System.IO.MemoryStream(, $iconBytes)
    $bitmap = [System.Drawing.Bitmap]::FromStream($ms)
    $Global:AppIcon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
    $ms.Dispose()
}
catch {
    Write-Warning "Could not load embedded app icon: $($_.Exception.Message)"
}

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
# Set the desired theme here: 'Dracula', or 'Light'
$Global:ThemeSelection = 'Dracula' 
$Global:Themes = @{
    Light   = @{
        FormBack    = [System.Drawing.Color]::FromArgb(240, 240, 240)
        TextFore    = [System.Drawing.Color]::Black
        ControlBack = [System.Drawing.Color]::White
        ButtonBack  = [System.Drawing.Color]::FromArgb(225, 225, 225)
        Border      = [System.Drawing.Color]::DimGray
        Accent      = [System.Drawing.Color]::DodgerBlue
    }
    Dracula = @{
        FormBack    = [System.Drawing.Color]::FromArgb(40, 42, 54)   # Background #282A36 (Shadow Grey)
        TextFore    = [System.Drawing.Color]::FromArgb(248, 248, 242) # Foreground #F8F8F2
        ControlBack = [System.Drawing.Color]::FromArgb(32, 32, 32)   # Carbon Black #202020
        ButtonBack  = [System.Drawing.Color]::FromArgb(47, 52, 64)   # Jet Black #2F3440
        Border      = [System.Drawing.Color]::FromArgb(68, 71, 90)   # Current Line #44475A
        Accent      = [System.Drawing.Color]::FromArgb(98, 114, 164)  # Comment #6272A4
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
    Theme = "Dracula"
}

function Get-UserSettings {
    if (Test-Path $Global:ConfigFile) {
        try {
            $json = Get-Content $Global:ConfigFile -Raw | ConvertFrom-Json
            $Global:UserSettings.Theme = if ($json.Theme) { $json.Theme } else { "Dracula" }
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
# Uses modern Windows 10/11 folder picker dialog via COM IFileOpenDialog
# -------------------------
function Select-ExtractionPath {
    param([string]$DefaultPath)
    
    # Add COM type definition for modern folder picker
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
public class FileOpenDialogRCW { }

[ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IFileOpenDialog {
    [PreserveSig] int Show([In] IntPtr hwndOwner);
    void SetFileTypes();
    void SetFileTypeIndex();
    void GetFileTypeIndex();
    void Advise();
    void Unadvise();
    void SetOptions([In] uint fos);
    void GetOptions();
    void SetDefaultFolder();
    void SetFolder([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi);
    void GetFolder();
    void GetCurrentSelection();
    void SetFileName();
    void GetFileName();
    void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
    void SetOkButtonLabel();
    void SetFileNameLabel();
    [PreserveSig] int GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
    void AddPlace();
    void SetDefaultExtension();
    void Close();
    void SetClientGuid();
    void ClearClientData();
    void SetFilter();
    void GetResults();
    void GetSelectedItems();
}

[ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IShellItem {
    void BindToHandler();
    void GetParent();
    [PreserveSig] int GetDisplayName([In] uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    void GetAttributes();
    void Compare();
}

public static class FolderPicker {
    public const uint FOS_PICKFOLDERS = 0x20;
    public const uint FOS_FORCEFILESYSTEM = 0x40;
    public const uint SIGDN_FILESYSPATH = 0x80058000;
    
    public static string ShowDialog(string title, string defaultPath) {
        var dialog = (IFileOpenDialog)new FileOpenDialogRCW();
        dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
        dialog.SetTitle(title);
        
        if (dialog.Show(IntPtr.Zero) == 0) {
            IShellItem result;
            if (dialog.GetResult(out result) == 0) {
                string path;
                result.GetDisplayName(SIGDN_FILESYSPATH, out path);
                return path;
            }
        }
        return null;
    }
}
"@ -ErrorAction SilentlyContinue

    try {
        $selectedPath = [FolderPicker]::ShowDialog("Select extraction destination folder", $DefaultPath)
        if ($selectedPath) {
            Write-LogMessage "User selected extraction path: $selectedPath"
            return $selectedPath
        }
    }
    catch {
        Write-LogMessage "Modern folder picker failed: $($_.Exception.Message)"
        # Fallback to classic dialog
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select extraction destination folder"
        $folderBrowser.SelectedPath = $DefaultPath
        $folderBrowser.ShowNewFolderButton = $true
        
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Write-LogMessage "User selected extraction path (fallback): $($folderBrowser.SelectedPath)"
            return $folderBrowser.SelectedPath
        }
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
    if ($Global:AppIcon) { $form.Icon = $Global:AppIcon }
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 15)
    $label.Size = New-Object System.Drawing.Size(400, 25)
    $label.Text = "Extracting: $ArchiveName"
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)
    
    # Custom themed progress bar (panel-based for color control)
    $progressPanel = New-Object System.Windows.Forms.Panel
    $progressPanel.Location = New-Object System.Drawing.Point(20, 50)
    $progressPanel.Size = New-Object System.Drawing.Size(400, 25)
    $progressPanel.BackColor = $colors.ControlBack
    $progressPanel.BorderStyle = 'FixedSingle'
    $form.Controls.Add($progressPanel)
    
    $progressFill = New-Object System.Windows.Forms.Panel
    $progressFill.Location = New-Object System.Drawing.Point(0, 0)
    $progressFill.Size = New-Object System.Drawing.Size(0, 23)
    $progressFill.BackColor = $colors.Accent
    $progressPanel.Controls.Add($progressFill)
    
    # Marquee animation timer for indeterminate progress
    $marqueeTimer = New-Object System.Windows.Forms.Timer
    $marqueeTimer.Interval = 30
    $marqueePos = 0
    $marqueeWidth = 80
    $marqueeTimer.Add_Tick({
            $script:marqueePos = ($script:marqueePos + 3) % (400 + $marqueeWidth)
            $startX = $script:marqueePos - $marqueeWidth
            if ($startX -lt 0) { $startX = 0 }
            $endX = [Math]::Min($script:marqueePos, 400)
            $progressFill.Location = New-Object System.Drawing.Point($startX, 0)
            $progressFill.Size = New-Object System.Drawing.Size(($endX - $startX), 23)
        }.GetNewClosure())
    $marqueeTimer.Start()
    
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 85)
    $statusLabel.Size = New-Object System.Drawing.Size(400, 20)
    $statusLabel.Text = "Please wait..."
    $statusLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($statusLabel)
    
    # Store references for external access
    $form.Tag = @{
        ProgressPanel = $progressPanel
        ProgressFill  = $progressFill
        MarqueeTimer  = $marqueeTimer
        StatusLabel   = $statusLabel
        MainLabel     = $label
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
        # Stop marquee animation when switching to percentage mode
        if ($controls.MarqueeTimer -and $controls.MarqueeTimer.Enabled) {
            $controls.MarqueeTimer.Stop()
        }
        # Set progress fill width based on percentage
        $fillWidth = [int]([Math]::Round(($Percentage / 100.0) * 398))
        $controls.ProgressFill.Location = New-Object System.Drawing.Point(0, 0)
        $controls.ProgressFill.Size = New-Object System.Drawing.Size($fillWidth, 23)
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
    if ($Global:AppIcon) { $form.Icon = $Global:AppIcon }

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
    $form.Size = New-Object System.Drawing.Size(400, 180)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    if ($Global:AppIcon) { $form.Icon = $Global:AppIcon }

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
    [void]$themeCombo.Items.AddRange(@('Dracula', 'Light'))
    $themeCombo.SelectedItem = $Global:UserSettings.Theme
    $form.Controls.Add($themeCombo)

    # Info label
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(20, 60)
    $infoLabel.Size = New-Object System.Drawing.Size(340, 40)
    $infoLabel.Text = "Theme changes apply to new windows.`nSettings are saved to config.json"
    $infoLabel.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150)
    $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($infoLabel)

    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(160, 100)
    $saveButton.Size = New-Object System.Drawing.Size(100, 35)
    $saveButton.Text = "&Save"
    $saveButton.Add_Click({
            $Global:UserSettings.Theme = $themeCombo.SelectedItem
            $Global:ThemeSelection = $Global:UserSettings.Theme
            $Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]
            if (Save-UserSettings) {
                Play-ActionSound -Type Success
                [System.Windows.Forms.MessageBox]::Show("Settings saved!", "Settings", 'OK', 'Information')
            }
            $form.Close()
        })

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(270, 100)
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
    if ($Global:AppIcon) { $form.Icon = $Global:AppIcon }

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
                    Show-ExecutableSelectionForm -FoundExecutables $executablesInArchive -WindowTitle "qBitLauncher"
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
        
        Show-ExecutableSelectionForm -FoundExecutables $executables -WindowTitle "qBitLauncher"
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