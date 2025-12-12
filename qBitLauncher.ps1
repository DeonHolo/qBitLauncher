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

# -------------------------
# Configuration
# -------------------------
# Use a universal path in the user's temp directory for portability.
$LogFile = Join-Path ([System.IO.Path]::GetTempPath()) "qBitLauncher_log.txt"
$ArchiveExtensions = @('iso', 'zip', 'rar', '7z', 'img')
$MediaExtensions = @('mp4', 'mkv', 'avi', 'mov', 'wmv', 'flv', 'webm', 'mp3', 'flac', 'wav', 'aac', 'ogg', 'm4a')
$ProcessableExtensions = @('exe') + $ArchiveExtensions

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
# Helper: Extract Archive (Verb-Noun: Expand-ArchiveFile)
# -------------------------
function Expand-ArchiveFile {
    param([string]$ArchivePath, [string]$ParentDirectory, [string]$BaseName)
    $ArchiveType = [IO.Path]::GetExtension($ArchivePath).TrimStart('.').ToLowerInvariant()
    Write-LogMessage "Attempting to extract '${ArchivePath}' (Type: ${ArchiveType})"; Write-Host "Attempting to extract '${ArchivePath}'..."
    $extractDir = Join-Path $ParentDirectory $BaseName 
    Write-LogMessage "Target extraction directory: '$extractDir'"

    if (-not (Test-Path -LiteralPath $extractDir)) {
        try { New-Item -ItemType Directory -Path $extractDir -ErrorAction Stop | Out-Null; Write-LogMessage "Created extraction directory: '$extractDir'" } 
        catch { $errMsg = "Failed to create extraction directory: '$extractDir'. Error: $($_.Exception.Message)"; Write-Error $errMsg; Write-LogMessage "ERROR: $errMsg"; return $null }
    }
    else { Write-LogMessage "Extraction directory '$extractDir' already exists." }

    try {
        # Try native PowerShell for ZIP first
        if ($ArchiveType -eq 'zip') {
            try { 
                Write-Host "Using native PowerShell to extract ZIP..."
                Expand-Archive -LiteralPath $ArchivePath -DestinationPath $extractDir -Force -ErrorAction Stop
                Write-Host "ZIP extracted successfully."
                Write-LogMessage "ZIP extracted with Expand-Archive."
                Show-ToastNotification -Title "Extraction Complete" -Message "ZIP extracted to: $extractDir" -Type Info
                return $extractDir 
            } 
            catch { Write-Warning "Native ZIP extraction failed. Trying other extractors..."; Write-LogMessage "Native ZIP extraction failed. Trying other extractors." }
        }
        
        # Try 7-Zip first (more common and handles more formats)
        $sevenZip = Get-7ZipPath
        if ($sevenZip) {
            Write-Host "Extracting with 7-Zip..."
            $processArgs = @('x', "`"$ArchivePath`"", "-o`"$extractDir`"", '-y')
            $process = Start-Process -FilePath $sevenZip -ArgumentList $processArgs -NoNewWindow -Wait -PassThru
            
            if ($process.ExitCode -eq 0) { 
                Write-Host "$($ArchiveType.ToUpper()) extracted successfully with 7-Zip to $extractDir"
                Write-LogMessage "Extracted with 7-Zip successfully."
                Show-ToastNotification -Title "Extraction Complete" -Message "$($ArchiveType.ToUpper()) extracted to: $extractDir" -Type Info
                return $extractDir 
            }
            else {
                Write-Warning "7-Zip extraction failed with exit code $($process.ExitCode). Trying WinRAR..."
                Write-LogMessage "7-Zip extraction failed. Exit Code: $($process.ExitCode). Falling back to WinRAR."
            }
        }
        
        # Fallback to WinRAR
        $winrar = Get-WinRARPath
        if (-not $winrar) { 
            $errMsg = "No archive extractor found (tried 7-Zip and WinRAR). Cannot extract '${ArchivePath}'. Please install 7-Zip or WinRAR."
            Write-Error $errMsg
            Write-LogMessage "ERROR: $errMsg"
            Show-ToastNotification -Title "Extraction Failed" -Message "No extractor found. Install 7-Zip or WinRAR." -Type Error
            return $null 
        }

        Write-Host "Extracting with WinRAR..."
        $processArgs = @('x', "`"$ArchivePath`"", "`"$extractDir\`"", '-y', '-o+')
        $process = Start-Process -FilePath $winrar -ArgumentList $processArgs -NoNewWindow -Wait -PassThru
        
        if ($process.ExitCode -ne 0) { 
            $warnMsg = "WinRAR extraction might have failed for '${ArchivePath}'. Exit Code: $($process.ExitCode)."
            Write-Warning $warnMsg
            Write-LogMessage "WARNING: $warnMsg"
            Show-ToastNotification -Title "Extraction Warning" -Message "WinRAR reported exit code $($process.ExitCode)" -Type Warning
            return $null 
        } 
        else { 
            Write-Host "$($ArchiveType.ToUpper()) extracted successfully with WinRAR to $extractDir"
            Write-LogMessage "Extracted with WinRAR successfully."
            Show-ToastNotification -Title "Extraction Complete" -Message "$($ArchiveType.ToUpper()) extracted to: $extractDir" -Type Info
            return $extractDir 
        }
    }
    catch { 
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
# GUI: Custom Confirmation Form (WIDER)
# ---------------------------------------------------
function Show-CustomConfirmForm {
    param(
        [string]$Message,
        [string]$Title = "Confirmation"
    )
    $colors = $Global:CurrentTheme

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(700, 180) 
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(660, 60)
    $label.Text = $Message
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)

    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(450, 100) 
    $yesButton.Size = New-Object System.Drawing.Size(100, 30)
    $yesButton.Text = "Yes"
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes

    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(560, 100)
    $noButton.Size = New-Object System.Drawing.Size(100, 30)
    $noButton.Text = "No"
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No

    foreach ($button in @($yesButton, $noButton)) {
        $button.BackColor = $colors.ButtonBack
        $button.ForeColor = $colors.TextFore
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $colors.Accent
        $form.Controls.Add($button)
    }
    
    $form.AcceptButton = $yesButton
    $form.CancelButton = $noButton

    $result = $form.ShowDialog()
    $form.Dispose()

    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

# ---------------------------------------------------
# GUI: Main Executable Selection Form (with Run as Admin)
# ---------------------------------------------------
function Show-ExecutableSelectionForm {
    param(
        [System.Management.Automation.PSObject[]]$FoundExecutables,
        [string]$WindowTitle = "qBitLauncher Action"
    )
    $colors = $Global:CurrentTheme

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(700, 420)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $colors.FormBack
    $font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Font = $font

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(660, 30)
    $label.Text = "Please select an executable and choose an action."
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 40)
    $listBox.Size = New-Object System.Drawing.Size(660, 280)
    $listBox.BackColor = $colors.ControlBack
    $listBox.ForeColor = $colors.TextFore
    $listBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $listBox.Font = $font 
    $listBox.HorizontalScrollbar = $true
    
    $itemAdded = $false
    foreach ($exe in $FoundExecutables) {
        if ($exe -and $exe.FullName) {
            [void]$listBox.Items.Add($exe.FullName)
            $itemAdded = $true
        }
    }
    
    if ($itemAdded) {
        $listBox.SelectedIndex = 0
    }

    $form.Controls.Add($listBox)

    # Button row - Main actions (Run always elevates to Admin)
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Location = New-Object System.Drawing.Point(10, 335)
    $runButton.Size = New-Object System.Drawing.Size(100, 30)
    $runButton.Text = "&Run"
    $runButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    $shortcutButton = New-Object System.Windows.Forms.Button
    $shortcutButton.Location = New-Object System.Drawing.Point(115, 335)
    $shortcutButton.Size = New-Object System.Drawing.Size(110, 30)
    $shortcutButton.Text = "Create &Shortcut"
    $shortcutButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    
    $exploreButton = New-Object System.Windows.Forms.Button
    $exploreButton.Location = New-Object System.Drawing.Point(230, 335)
    $exploreButton.Size = New-Object System.Drawing.Size(100, 30)
    $exploreButton.Text = "&Open Folder"
    $exploreButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry 
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(570, 335)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Text = "&Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    foreach ($button in @($runButton, $shortcutButton, $exploreButton, $cancelButton)) {
        $button.BackColor = $colors.ButtonBack
        $button.ForeColor = $colors.TextFore
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $colors.Accent
        $form.Controls.Add($button)
    }

    $form.AcceptButton = $runButton
    $form.CancelButton = $cancelButton
    $form.ActiveControl = $listBox

    $dialogResult = $form.ShowDialog()
    $form.Dispose()

    # Create the return object
    $returnInfo = @{ DialogResult = $dialogResult }

    # Only add the executable if the choice was positive and an item was selected.
    $positiveResults = @(
        [System.Windows.Forms.DialogResult]::OK, 
        [System.Windows.Forms.DialogResult]::Yes, 
        [System.Windows.Forms.DialogResult]::Retry
    )
    
    if (($dialogResult -in $positiveResults) -and $listBox.SelectedItem) {
        try {
            $selectedItem = Get-Item -LiteralPath $listBox.SelectedItem -ErrorAction Stop
            $returnInfo.Add('SelectedExecutable', $selectedItem)
        }
        catch {
            $errMsg = "FATAL: Failed to Get-Item on '$($listBox.SelectedItem)'. Error: $($_.Exception.Message)"
            Write-LogMessage $errMsg
            Write-Error $errMsg
            $returnInfo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        }
    }
    
    return $returnInfo
}

# ---------------------------------------------------
# Consolidated Action Handler (NEW - Eliminates duplicate switch blocks)
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

    if ($ArchiveExtensions -contains $ext) {
        Write-LogMessage "Processing archive: '$filePath'"
        Write-Host "`nFound an archive file: $filePath"
        if (Show-CustomConfirmForm -Message "An archive file was found. Proceed with extraction?`n`nFile: $filePath" -Title "Confirm Extraction") {
            Write-LogMessage "User confirmed extraction."
            $extractedDir = Expand-ArchiveFile -ArchivePath $filePath -ParentDirectory $parentDir -BaseName $baseName
            if ($extractedDir) {
                Write-Host "`nExtraction complete. Searching for executables..."
                $executablesInArchive = Get-AllExecutables -RootFolderPath $extractedDir
                if ($executablesInArchive) {
                    $guiResult = Show-ExecutableSelectionForm -FoundExecutables $executablesInArchive -WindowTitle "Archive Extracted"
                    Invoke-UserAction -GuiResult $guiResult
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
        
        $guiResult = Show-ExecutableSelectionForm -FoundExecutables $executables -WindowTitle "Executable Found"
        Invoke-UserAction -GuiResult $guiResult
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