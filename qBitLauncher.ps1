# qBitLauncher_v2.ps1
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
    Light = @{
        FormBack      = [System.Drawing.Color]::FromArgb(240, 240, 240)
        TextFore      = [System.Drawing.Color]::Black
        ControlBack   = [System.Drawing.Color]::White
        ButtonBack    = [System.Drawing.Color]::FromArgb(225, 225, 225)
        Border        = [System.Drawing.Color]::DimGray
        Accent        = [System.Drawing.Color]::DodgerBlue
    }
    Dark = @{
        FormBack      = [System.Drawing.Color]::FromArgb(45, 45, 48)
        TextFore      = [System.Drawing.Color]::White
        ControlBack   = [System.Drawing.Color]::FromArgb(30, 30, 30)
        ButtonBack    = [System.Drawing.Color]::FromArgb(63, 63, 70)
        Border        = [System.Drawing.Color]::FromArgb(85, 85, 91)
        Accent        = [System.Drawing.Color]::FromArgb(0, 122, 204)
    }
    qBitDark = @{
        FormBack      = [System.Drawing.Color]::FromArgb(45, 49, 58)  # Main dark background
        TextFore      = [System.Drawing.Color]::FromArgb(220, 220, 220) # Off-white text
        ControlBack   = [System.Drawing.Color]::FromArgb(34, 38, 46)  # Slightly lighter control background
        ButtonBack    = [System.Drawing.Color]::FromArgb(59, 64, 75)  # Button background
        Border        = [System.Drawing.Color]::FromArgb(80, 85, 96)  # Border for buttons/controls
        Accent        = [System.Drawing.Color]::FromArgb(80, 128, 175) # Blue from progress bar
    }
}
$Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]

# -------------------------
# Helper: Logging
# -------------------------
function Log-Message {
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

Log-Message "--------------------------------------------------------"
Log-Message "Script started: qBitLauncher.ps1"
Log-Message "Received initial path from qBittorrent: '$filePathFromQB'"
Write-Host "qBitLauncher.ps1 started. Log file: $LogFile"

# -------------------------
# Helper: Find WinRAR.exe
# -------------------------
function Get-WinRARPath {
    foreach ($path in @((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe' -ErrorAction SilentlyContinue).'(Default)', "$env:ProgramFiles\WinRAR\WinRAR.exe", "$env:ProgramFiles(x86)\WinRAR\WinRAR.exe")) {
        if ($path -and (Test-Path $path)) { Log-Message "Found WinRAR at: $path"; return $path }
    }
    Log-Message "WinRAR not found."; return $null
}

# -------------------------
# Helper: Extract Archive
# -------------------------
function Extract-Archive {
    param([string]$ArchivePath, [string]$ParentDirectory, [string]$BaseName)
    $ArchiveType = [IO.Path]::GetExtension($ArchivePath).TrimStart('.').ToLowerInvariant()
    Log-Message "Attempting to extract '${ArchivePath}' (Type: ${ArchiveType})"; Write-Host "Attempting to extract '${ArchivePath}'..."
    $extractDir = Join-Path $ParentDirectory $BaseName 
    Log-Message "Target extraction directory: '$extractDir'"

    if (-not (Test-Path -LiteralPath $extractDir)) {
        try { New-Item -ItemType Directory -Path $extractDir -ErrorAction Stop | Out-Null; Log-Message "Created extraction directory: '$extractDir'" } 
        catch { $errMsg = "Failed to create extraction directory: '$extractDir'. Error: $($_.Exception.Message)"; Write-Error $errMsg; Log-Message "ERROR: $errMsg"; return $null }
    } else { Log-Message "Extraction directory '$extractDir' already exists." }

    try {
        if ($ArchiveType -eq 'zip') {
            try { Write-Host "Using native PowerShell to extract ZIP..."; Expand-Archive -LiteralPath $ArchivePath -DestinationPath $extractDir -Force -ErrorAction Stop; Write-Host "ZIP extracted successfully."; Log-Message "ZIP extracted with Expand-Archive."; return $extractDir } 
            catch { Write-Warning "Native ZIP extraction failed. Falling back to WinRAR..."; Log-Message "Native ZIP extraction failed. Falling back to WinRAR." }
        }
        
        $winrar = Get-WinRARPath
        if (-not $winrar) { $errMsg = "WinRAR not found. Cannot extract '${ArchivePath}'. Please install WinRAR."; Write-Error $errMsg; Log-Message "ERROR: $errMsg"; return $null }

        Write-Host "Extracting with WinRAR..."
        $args = @('x', "`"$ArchivePath`"", "`"$extractDir\`"", '-y', '-o+');
        $process = Start-Process -FilePath $winrar -ArgumentList $args -NoNewWindow -Wait -PassThru
        
        if ($process.ExitCode -ne 0) { $warnMsg = "WinRAR extraction might have failed for '${ArchivePath}'. Exit Code: $($process.ExitCode)."; Write-Warning $warnMsg; Log-Message "WARNING: $warnMsg"; return $null } 
        else { Write-Host "$($ArchiveType.ToUpper()) extracted successfully to $extractDir"; return $extractDir }
    } catch { $errMsg = "Error during extraction process for '${ArchivePath}'. Error: $($_.Exception.Message)"; Write-Error $errMsg; Log-Message "ERROR: $errMsg"; return $null }
}

# -------------------------
# Helper: Find ALL Executables
# -------------------------
function Find-AllExecutables {
    param([string]$RootFolderPath)
    Log-Message "Searching for all executables (.exe) in '$RootFolderPath' (depth-first sort)."; Write-Host "Searching for all .exe files in '$RootFolderPath' and its subfolders..."
    $allExecutables = Get-ChildItem -LiteralPath $RootFolderPath -Filter *.exe -File -Recurse -ErrorAction SilentlyContinue
    if ($allExecutables) {
        $sortedExecutables = $allExecutables | Sort-Object @{Expression = {($_.FullName.Split([IO.Path]::DirectorySeparatorChar).Count)}}, FullName
        Log-Message "Found $($sortedExecutables.Count) executables."; return $sortedExecutables
    }
    Log-Message "No .exe files found in '$RootFolderPath'."; Write-Warning "No .exe files found in '$RootFolderPath' or its subdirectories."; return $null
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
    # --- FIX: Made form wider (700px) ---
    $form.Size = New-Object System.Drawing.Size(700, 180) 
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
     # --- FIX: Made label wider (660px) ---
    $label.Size = New-Object System.Drawing.Size(660, 60)
    $label.Text = $Message
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)

    $yesButton = New-Object System.Windows.Forms.Button
    # --- FIX: Repositioned buttons for new width ---
    $yesButton.Location = New-Object System.Drawing.Point(450, 100) 
    $yesButton.Size = New-Object System.Drawing.Size(100, 30)
    $yesButton.Text = "Yes"
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes

    $noButton = New-Object System.Windows.Forms.Button
    # --- FIX: Repositioned buttons for new width ---
    $noButton.Location = New-Object System.Drawing.Point(560, 100)
    $noButton.Size = New-Object System.Drawing.Size(100, 30)
    $noButton.Text = "No"
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No

    foreach ($button in @($yesButton, $noButton)) {
        $button.BackColor = $colors.ButtonBack
        $button.ForeColor = $colors.TextFore
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $colors.Accent # Using accent color for borders
        $form.Controls.Add($button)
    }
    
    $form.AcceptButton = $yesButton
    $form.CancelButton = $noButton

    $result = $form.ShowDialog()
    $form.Dispose()

    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

# ---------------------------------------------------
# GUI: Main Executable Selection Form (SCROLLBAR FIX)
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
    
    # --- FIX: Added HorizontalScrollbar ---
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

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Location = New-Object System.Drawing.Point(10, 335)
    $runButton.Size = New-Object System.Drawing.Size(120, 30)
    $runButton.Text = "&Run Selected"
    $runButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    $shortcutButton = New-Object System.Windows.Forms.Button
    $shortcutButton.Location = New-Object System.Drawing.Point(140, 335)
    $shortcutButton.Size = New-Object System.Drawing.Size(120, 30)
    $shortcutButton.Text = "Create &Shortcut"
    $shortcutButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    
    $exploreButton = New-Object System.Windows.Forms.Button
    $exploreButton.Location = New-Object System.Drawing.Point(270, 335)
    $exploreButton.Size = New-Object System.Drawing.Size(120, 30)
    $exploreButton.Text = "&Open Folder"
    $exploreButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry 
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(550, 335)
    $cancelButton.Size = New-Object System.Drawing.Size(120, 30)
    $cancelButton.Text = "&Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    foreach ($button in @($runButton, $shortcutButton, $exploreButton, $cancelButton)) {
        $button.BackColor = $colors.ButtonBack
        $button.ForeColor = $colors.TextFore
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 1
        $button.FlatAppearance.BorderColor = $colors.Accent # Using accent color for borders
        $form.Controls.Add($button)
    }

    $form.AcceptButton = $runButton
    $form.CancelButton = $cancelButton
    $form.ActiveControl = $listBox

    # --------- START: ROBUST RETURN LOGIC (FIXED) ---------
    
    $dialogResult = $form.ShowDialog()
    $form.Dispose()

    # Create the return object *first*
    $returnInfo = @{ DialogResult = $dialogResult }

    # Only add the executable if the choice was positive
    # and an item was actually selected.
    $positiveResults = @(
        [System.Windows.Forms.DialogResult]::OK, 
        [System.Windows.Forms.DialogResult]::Yes, 
        [System.Windows.Forms.DialogResult]::Retry
    )
    
    if (($dialogResult -in $positiveResults) -and $listBox.SelectedItem) {
        try {
            # Use -ErrorAction Stop to force a terminating error if the path is bad
            $selectedItem = Get-Item -LiteralPath $listBox.SelectedItem -ErrorAction Stop
            $returnInfo.Add('SelectedExecutable', $selectedItem)
        } catch {
            $errMsg = "FATAL: Failed to Get-Item on '$($listBox.SelectedItem)'. Error: $($_.Exception.Message)"
            Log-Message $errMsg
            Write-Error $errMsg
            # Failed to get the item, so treat it as a cancel.
            $returnInfo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        }
    }
    
    return $returnInfo
    
    # --------- END: ROBUST RETURN LOGIC (FIXED) ---------
}


# ===================================================================
# MAIN SCRIPT LOGIC STARTS HERE
# ===================================================================

if (-not (Test-Path -LiteralPath $filePathFromQB)) {
    $errMsg = "Error: Initial path not found - $filePathFromQB"; Write-Error $errMsg; Log-Message "FATAL: $errMsg. Script exiting."; Read-Host "Press Enter to exit..."; exit 1
}

$mainFileToProcess = $null

if (Test-Path -LiteralPath $filePathFromQB -PathType Container) {
    $downloadFolder = $filePathFromQB
    Write-Host "Input path is a folder: '$downloadFolder'. Searching for primary file..."
    $mainFileToProcess = Get-ChildItem -LiteralPath $downloadFolder -File -Recurse | Where-Object { $ArchiveExtensions -contains $_.Extension.TrimStart('.').ToLowerInvariant() } | Sort-Object Length -Descending | Select-Object -First 1
    
    if ($mainFileToProcess) {
        Log-Message "Found a primary archive file to process in folder: '$($mainFileToProcess.FullName)'"
    } else {
        Log-Message "No archives found in folder. Searching for executables..."
        $allExecutables = Find-AllExecutables -RootFolderPath $downloadFolder
        if ($allExecutables) {
            $mainFileToProcess = $allExecutables
        } else {
            Log-Message "No executables found. Checking for media files..."
            $foundMediaFile = Get-ChildItem -LiteralPath $downloadFolder -File -Recurse | Where-Object { $MediaExtensions -contains $_.Extension.TrimStart('.').ToLowerInvariant() } | Select-Object -First 1
            if ($foundMediaFile) {
                Write-Host "Found a media file: $($foundMediaFile.Name). Opening folder."
                Start-Process explorer -ArgumentList "`"$(Split-Path $foundMediaFile.FullName -Parent)`""
            } else {
                Write-Warning "No processable files found in '$downloadFolder'."
                Start-Process explorer -ArgumentList "`"$downloadFolder`""
            }
        }
    }
} else {
    $mainFileToProcess = Get-Item -LiteralPath $filePathFromQB
    Log-Message "Input is a single file: '$($mainFileToProcess.FullName)'"
}

if ($mainFileToProcess) {
    $firstFile = if ($mainFileToProcess -is [array]) { $mainFileToProcess[0] } else { $mainFileToProcess }
    $filePath = $firstFile.FullName
    $parentDir = $firstFile.DirectoryName
    $baseName = $firstFile.BaseName
    $ext = $firstFile.Extension.ToLowerInvariant().TrimStart('.')

    if ($ArchiveExtensions -contains $ext) {
        Log-Message "Processing archive: '$filePath'"; Write-Host "`nFound an archive file: $filePath"
        if (Show-CustomConfirmForm -Message "An archive file was found. Proceed with extraction?`n`nFile: $filePath" -Title "Confirm Extraction") {
            Log-Message "User confirmed extraction."
            $extractedDir = Extract-Archive -ArchivePath $filePath -ParentDirectory $parentDir -BaseName $baseName
            if ($extractedDir) {
                Write-Host "`nExtraction complete. Searching for executables..."
                $executablesInArchive = Find-AllExecutables -RootFolderPath $extractedDir
                if ($executablesInArchive) {
                    $guiResult = Show-ExecutableSelectionForm -FoundExecutables $executablesInArchive -WindowTitle "Archive Extracted"
                    $selectedExecutable = $guiResult.SelectedExecutable
                    
                    switch ($guiResult.DialogResult) {
                        'OK' { # Run
                            Log-Message "User chose to run '$($selectedExecutable.FullName)'."
                            Write-Host "Attempting to run '$($selectedExecutable.FullName)'..."
                            try { Push-Location -LiteralPath $selectedExecutable.DirectoryName; & $selectedExecutable.FullName; Pop-Location } 
                            catch { $errMsg = "Error starting executable '$($selectedExecutable.FullName)': $($_.Exception.Message)"; Write-Warning $errMsg; Log-Message "WARNING: $errMsg" }
                        }
                        'Yes' { # Shortcut
                            Log-Message "User chose to create shortcut for '$($selectedExecutable.FullName)'."
                            try {
                                $desktopPath = [System.Environment]::GetFolderPath('Desktop'); $shortcutName = $selectedExecutable.BaseName + ".lnk"; $shortcutPath = Join-Path $desktopPath $shortcutName
                                Log-Message "Creating shortcut: '$shortcutPath'"; $wshell = New-Object -ComObject WScript.Shell; $shortcut = $wshell.CreateShortcut($shortcutPath)
                                $shortcut.TargetPath = $selectedExecutable.FullName; $shortcut.WorkingDirectory = $selectedExecutable.DirectoryName; $shortcut.Save()
                                Write-Host "Shortcut created on Desktop: $shortcutPath"; Log-Message "Shortcut created."
                            } catch { $errMsg = "Error creating shortcut: $($_.Exception.Message)"; Write-Warning $errMsg; Log-Message "WARNING: $errMsg" }
                        }
                        'Retry' { # Explore
                            Log-Message "User chose to open the folder."; Write-Host "Opening folder."
                            Start-Process explorer -ArgumentList "`"$($selectedExecutable.DirectoryName)`""
                        }
                        default { # Cancel or Closed
                            Log-Message "User cancelled or closed the selection window."
                            Write-Host "Action cancelled."
                        }
                    }
                } else {
                    Write-Warning "No executables found in the extracted folder: $extractedDir"
                    Start-Process explorer -ArgumentList "`"$extractedDir`""
                }
            }
        } else {
            Log-Message "User declined extraction. Opening folder: '$parentDir'"; Start-Process explorer -ArgumentList "`"$parentDir`""
        }
    } 
    elseif ($ext -eq 'exe' -or $mainFileToProcess -is [array]) {
        $executables = if ($mainFileToProcess -is [array]) { $mainFileToProcess } else { @($mainFileToProcess) }
        Log-Message "Processing one or more executables."
        
        $guiResult = Show-ExecutableSelectionForm -FoundExecutables $executables -WindowTitle "Executable Found"
        $selectedExecutable = $guiResult.SelectedExecutable

        switch ($guiResult.DialogResult) {
            'OK' { # Run
                Log-Message "User chose to run '$($selectedExecutable.FullName)'."
                Write-Host "Attempting to run '$($selectedExecutable.FullName)'..."
                try { Push-Location -LiteralPath $selectedExecutable.DirectoryName; & $selectedExecutable.FullName; Pop-Location } 
                catch { $errMsg = "Error starting executable '$($selectedExecutable.FullName)': $($_.Exception.Message)"; Write-Warning $errMsg; Log-Message "WARNING: $errMsg" }
            }
            'Yes' { # Shortcut
                Log-Message "User chose to create shortcut for '$($selectedExecutable.FullName)'."
                try {
                    $desktopPath = [System.Environment]::GetFolderPath('Desktop'); $shortcutName = $selectedExecutable.BaseName + ".lnk"; $shortcutPath = Join-Path $desktopPath $shortcutName
                    Log-Message "Creating shortcut: '$shortcutPath'"; $wshell = New-Object -ComObject WScript.Shell; $shortcut = $wshell.CreateShortcut($shortcutPath)
                    $shortcut.TargetPath = $selectedExecutable.FullName; $shortcut.WorkingDirectory = $selectedExecutable.DirectoryName; $shortcut.Save()
                    Write-Host "Shortcut created on Desktop: $shortcutPath"; Log-Message "Shortcut created."
                } catch { $errMsg = "Error creating shortcut: $($_.Exception.Message)"; Write-Warning $errMsg; Log-Message "WARNING: $errMsg" }
            }
            'Retry' { # Explore
                Log-Message "User chose to open the folder."; Write-Host "Opening folder."
                Start-Process explorer -ArgumentList "`"$($selectedExecutable.DirectoryName)`""
            }
            default { # Cancel or Closed
                Log-Message "User cancelled or closed the selection window."
                Write-Host "Action cancelled."
            }
        }
    } 
    elseif ($MediaExtensions -contains $ext) {
        Log-Message "File is a media file."; Write-Host "Media file '${filePath}' is ready."
        Start-Process explorer -ArgumentList "`"$parentDir`""; Write-Host "Opening containing folder: $parentDir"
    }
    else {
        Log-Message "File is an unhandled type (.$ext)."; Write-Warning "File type .${ext} is not handled explicitly."
        Start-Process explorer -ArgumentList "`"$parentDir`""; Write-Host "Opening containing folder: $parentDir"
    }
}

Write-Host "`nScript actions complete."
Log-Message "Script finished."
Log-Message "--------------------------------------------------------`n"