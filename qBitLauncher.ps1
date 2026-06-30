# qBitLauncher.ps1
# Trigger version bump

param(
    [string]$filePathFromQB,
    [string]$torrentHashFromQB,
    [string]$torrentNameFromQB
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

# Add SendMessage for proper taskbar icon support
Add-Type -Name User32Icon -Namespace Win32 -MemberDefinition '
[DllImport("user32.dll", CharSet = CharSet.Auto)]
public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
public const int WM_SETICON = 0x80;
public const int ICON_SMALL = 0;
public const int ICON_BIG = 1;
'

# -------------------------
# Configuration
# -------------------------
# Version and update settings
$Global:ScriptVersion = "3.0.0"
$Global:MainForm = $null
$Global:GitHubRawUrl = "https://raw.githubusercontent.com/DeonHolo/qBitLauncher/main/qBitLauncher.ps1"
$Global:GitHubCommitsUrl = "https://github.com/DeonHolo/qBitLauncher/commits/main"

# Determine script directory (handle PS2EXE compiled EXE where $PSScriptRoot is empty)
$Global:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } 
elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
elseif ([System.AppDomain]::CurrentDomain.BaseDirectory) { [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\') }
else { [Environment]::CurrentDirectory }

# Log file in script folder for easy access
$LogFile = Join-Path $Global:ScriptDir "qBitLauncher_log.txt"
$ArchiveExtensions = @('iso', 'zip', 'rar', '7z', 'img')
$MediaExtensions = @('mp4', 'mkv', 'avi', 'mov', 'wmv', 'flv', 'webm', 'mp3', 'flac', 'wav', 'aac', 'ogg', 'm4a')

# Load app icon from embedded Base64 ICO (no external file needed)
$Global:AppIcon = $null
$LogoIcoBase64 = "AAABAAEAHiAAAAEAIACoDwAAFgAAACgAAAAeAAAAQAAAAAEAIAAAAAAAAA8AAMMOAADDDgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANHT2wD47t0A8undBfnw3yD68d9G+u/dZvvu3HT779x3+vDeZvnw30P4798b8+rfA/fu4wD///8B+e7bMvPix5Pu2LbI7tm4yfTjyZL67dgu6///APbw5AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADs5doA///gAPfu2hT68NxZ+OrUp/Dcwdrozq7y4sGf/N+5lP7fupX+48Oh/OnQsfHx38TU+ezWlvrw3EP16M1a5sij3NGdav/HiU//x4tQ/9Cdav/myaXZ9+vVTPjPjgAAAAAAAAAAAAAAAAAAAAAAAAAAAPfs2QD269kK+e/bW/Xmzcjmyqf606N4/8OEUf+6dD7/uXA4/7luOP+5bTn/uW87/7t1Qf/DhlT/1Kd9/+3Vt/TlyKL0woJH/752N//AeDn/v3o9/797Pv/GiE7/5Mag3Pjs1TkAAAAAAAAAAAAAAADu5NsA/PHbAPnu3B/469Wf6s+v99Gcbv+/eUH/u3E4/7twOf+6bzn/uW83/7luOP+4bjj/uG05/7ltOv+4bDf/uG44/9qyiv/Ijlj/x5Bf/8+jef/Bf0b/07CP/9i6nP/Gj13/ypJc//DdwawAAAAAAAAAAO/m2AD77tgA+O3aLvbnz8Pgu5T/x4RN/8F2Pf/Adj7/vnU9/7xzPP+8cjr/u3A6/7pvOf+4bjf/uG03/7hsOf+3bDf/vntI/9y3j/++eT7/wX9F/9e3mP/Rpn7/woBH/8KAR/+/ej7/wn5C/+fKpusAAAAA6+XbAPnt2AD57two9ufOytywhv/Hf0b/xXxC/8R7Qv/DeUD/wXc//792Pf++dDv/vXM7/7xyO/+7cTr/unA4/7luN/+3bDX/v39N/9y2iv+/eTr/vXU2/8iRX//hzLb/wYBG/751N/+/djn/wHo+/+TCmfcAAAAA9urXAPft2xL36tKy4LeO/8qCSf/If0X/x39E/8V+Q//EfEL/wnlA/8N8RP/aqnr/0Jdk/750PP+9czv/vHI6/7txOf+5bzf/vnpF/967k/+/fUD/xIdQ/9u/o//Ik2P/v3c4/793Ov+/djn/xIJH/+jOqt7z6doA09jzAfnu2nzpyKX9zohQ/8yCSf/Lgkj/yYBH/8d/Rf/FfkT/w3xC/8aBSv/x1q//3rOH/8B3Pv/Adj7/vnQ8/71zO/+8czr/u3I6/9arfv/Nl2H/xotV/8mSYf++dTf/wHc6/8B3O/++dDf/0qFx/fPkzI3469cA+O7dMfPexOTWmmb/zoVL/86ES//Ng0r/zIFI/8uASP/Jf0b/xn1E/8iETP/x1K7/3rKH/8J4P//BeED/wHc//791Pf++dDz/vXM5/8KBTP/hvpf/y5Fc/7x0Nf+/dTb/vnQ3/792O//MlWL/6tGwwPrw3B/e3+4B+OzYkeS8kv/Rik3/0IlN/8+HTP/OhUv/zYNJ/8yBR//Kf0b/yX5F/8qGTv/x1a//37OH/8N6Qf/DekH/wng+/8F3Pv/AdTz/vnM6/71xN//HiFX/3bWK/9mwhP/Uom//0p9s/9mvhP/s1Lb/+OzXcfvt1QD47+Al9OHH2tmeZ//TjE//0otP/9CJTv/Ph0z/0ZBX/9igbP/Zom7/0ZJc/8yJUv/w1Kz/37WI/8V8Q//FgUr/yI1d/8J7Qv/Egk3/zpty/9CifP/FjV//v3xG/8yTYv/QoXP/0KBw/8eLWP/Omm7/9eXPtPnw5Av5791p7M6q+tWRVf/UjVH/041R/9GLT//cp3T/89ez//TbtP/y17D/8tey/+G1hv/x1K3/4beK/8d+Q//Olmf/7eLW/9Osiv/q3Mz/8erh/+/p3//y8Or/4cix/797R/+5bjb/u3A3/7pwNf++e0X/7NS46fvx4Tf469eu5LmM/9aPUv/Vj1P/1I1R/9icZv/127f/7cqh/9ecZf/Sk1v/151m/+zKoP/55cX/4bWJ/8l+RP/QmGr/+PXx//Xx7P/atpb/yo5b/8iJVP/Xr4z/9fTv/9q5nf+/dT3/vnQ7/71zOv+8czr/4b2Z/Prv3XD259De4Kp3/9iRVP/XkVX/1Y5S/+W5i//34L7/2Jxl/9CKTv/QiU//zoZL/9mjcP/55MP/4raJ/8yARv/SmWz/+ff0/+nTv//Igkn/xn5D/8Z+Qv/EfEH/4Mev//Lr4//HhlP/wHU8/790PP+9cjn/16l+//ns2aT14sj736Ns/9qTV//Zk1f/15NX/+3Lov/w0qz/1JJX/9KOUv/SjVH/0YtQ/9SUW//02rT/47mL/86ESf/Tm23/+Pbz/963lP/KgEX/yoFH/8iARf/GfUD/16yH//f38//Ol2r/wnc9/8J4Pv+/dDr/0Jts//jp08X04MT/36Jq/9yWWf/alFj/2pdb//HTrP/uzKL/1ZJW/9SRVP/TkFP/0o5S/9STWv/z2LL/5LuN/8+HSv/VnW//+PXx/92yi//Mgkf/zIJJ/8qBR//If0P/1KV7//f49P/SoXf/w3o//8N7QP/Bdzz/zpZl//fmz9P04cX/4aRr/96YW//dl1r/3Jhc//HSqv/vz6b/2JRY/9eTV//Wklb/1JBT/9aXXf/02rb/5buP/9CJTP/WoHH/+Pby/960jv/NhEj/zYRK/8yCSf/KgEX/1qV8//b49P/ToXf/xXxB/8R9Qv/Dej7/z5hn//fmz9H148n35Kpy/+GbXf/gmlz/3phb/+7Inf/027b/3Zxh/9qUWP/alFn/2JNW/9yfZ//53bv/57uO/9OLUP/ZonP/+ff0/+PCpP/Phkv/z4dM/86FS//Mgkf/3LWU//T07//Rl2j/yH5D/8d+RP/Fe0H/1KBx//jp08D36NHZ6LR//+OeXv/inl//4Jtd/+i1gv/55MT/57WB/9yWWf/blVf/25ZZ/+i/kv/75sb/6LyO/9aOUv/bpHT/+Pj0//Hn2//Ul2P/z4ZL/8+FSf/RkVv/7+LU/+vdzf/MiE//yoFH/8mARv/HfUP/26+F//nt2Z757Nis7MKV/+WgYf/koWL/4qBg/+KiZf/wzqL/9+C9/+3GmP/ouon/7ceb//LVr//13rv/6cCR/9iRU//dpnX/+PXv//Do3f/s387/4sGj/+G+n//r3sz/9fPs/9qqgf/NhUn/zIRK/8qCSP/JgEf/5cKe/Pvw3W76791r8dOu++WkZP/lomL/5KJi/+OgYP/jpWn/7cOT//TVq//z17D/7sqe/+GmbP/nuIb/5LJ//9mUVv/eqHf/+PXw/+G7lP/ivZj/79/P//Ll2f/r1sH/2ad5/9CKT//QiE3/zoZM/8yESv/OjVb/79a66fvw4Db47t8o9uLI3eiwdv/mpGP/5qRj/+WjY//koWH/46Bf/+OiZf/ioWT/4Jxe/9+aXP/emVv/3Jhb/9qWWf/fqnj/+fby/+S4kP/XkFH/1pRX/9aUW//VkFb/04xP/9KMUP/Rik//0IhO/86FSv/Zpnf/9+jTuPju4Azs6OYC+e3Zle/JnP/npmT/56Zl/+akZP/mpGP/5aNj/+SiYv/joGL/4p9h/+GeX//gnF3/3ppc/9yXWf/hrHr/+ffy/+W7k//ak1T/2ZNV/9eSVP/WkVX/1Y9U/9SOUv/TjVH/0oxQ/9GLUP/px6T8+u/dZPvs2AD57NkA+e/eN/bhw+nqsnX/6ahl/+inZf/npmX/5qVk/+WlZP/ko2P/46Fi/+OgYf/hn2D/4J5e/9+bW//irXv/+fjz/+a8lP/blVb/2pVY/9mUV//YlFb/15JW/9aQVP/Uj1L/041Q/92qef/25c7M+O3dGPbp1wD169oA8OniAvnt2Yzy0qr/66xq/+qpZv/pqGb/56Zm/+emZP/mpWT/5aRj/+WiYv/joWL/4aBg/+GeXf/iqnX/793J/+S0hf/dl1j/25ZZ/9qVWP/alVf/2ZRX/9iTVf/XkFP/15hd/+/Vtff68N9e/fDcAPLp2gAAAAAA+ezaAPnu3h356tLF8cWU/+yqaP/rqmf/6qhn/+mnZv/op2b/56Zl/+akZP/komP/46Jj/+OhYv/ioGL/4Z9i/+CdX//fmlz/3Zla/9yXWv/cllr/2pVZ/9qUV//Zl1v/6cWe//jr15707eAJ9evaAAAAAAAAAAAA8ejcAPvv1wD57ts5+OfM2/HCj//tq2j/7atn/+upaP/qqGf/6adm/+enZf/mpWT/5aNk/+WiY//koWL/46Bh/+KeX//gnF7/35tc/96aW//dmFv/3JZZ/9qWWf/nwJP/+OrVvvjs3x7469wA3t/eAAAAAAAAAAAAAAAAAPXq3AD/+NQA+u/bRPnoy9ryx5b/7q9s/+2rZ//sqmf/6qln/+moZv/op2X/56Zk/+akZP/lo2P/5KJi/+OhYP/in1//4J1e/9+cXf/fmln/359j/+vFnP/36NDA+e7cJ/vu2QDt5t0AAAAAAAAAAAAAAAAAAAAAAAAAAAD269oA//ncAPrv2zX669LB9NSr/u63e//trGj/7Ktm/+uqZ//qqWb/6ahl/+inZf/npWT/5qRj/+WjYv/koWD/4Z5d/+GfX//ksHn/8dax+vnt1qj57t0h/O/cAOzm4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA8unlAPzw2gD6790Z+u3ZhfjkxufzzqH/77p//+2ubf/sqmb/66hl/+qnZP/ppmP/6KRi/+ekYv/mqGn/6LV///DOpv735crX+u/bafjt2w357tsA6ubdAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD57N8A+OvhA/rv3TD67diK+efN0/bcuPXz0Kb+8ceY/+7BkP/uwpD/8MeZ//LQp/313Lrv+OjPxfvv23X67t4g79jjAPjq3wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPft4gD26+QC+vDfHPrx3k/779yB+u3YpPns1bf67da0+u7Znfrw3HX78N9A+e/gEubm4gDz7eEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/wAgM/wAABPwAAAD4AAAA8AAAAOAAAADAAAAAgAAAAIAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAAAASAAAAMwAAADOAAABzwAAA8+AAAfPwAAPz+AAP8/4AP/A=="
try {
    $iconBytes = [System.Convert]::FromBase64String($LogoIcoBase64)
    $iconStream = New-Object System.IO.MemoryStream(, $iconBytes)
    $Global:AppIcon = New-Object System.Drawing.Icon($iconStream)
    # Icon loaded successfully (silent - no output in compiled EXE)
}
catch {
    Write-Warning "Could not load embedded app icon: $($_.Exception.Message)"
}

# Helper function to set form icon properly (including taskbar)
function Set-FormIcon {
    param([System.Windows.Forms.Form]$Form)
    if ($Global:AppIcon) {
        $Form.Icon = $Global:AppIcon
        # Also set via SendMessage for proper taskbar display
        $handle = $Form.Handle
        [Win32.User32Icon]::SendMessage($handle, [Win32.User32Icon]::WM_SETICON, [IntPtr][Win32.User32Icon]::ICON_BIG, $Global:AppIcon.Handle) | Out-Null
        [Win32.User32Icon]::SendMessage($handle, [Win32.User32Icon]::WM_SETICON, [IntPtr][Win32.User32Icon]::ICON_SMALL, $Global:AppIcon.Handle) | Out-Null
    }
}

# Helper function to apply consistent theme styling to buttons
function Set-ThemedButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [hashtable]$Colors = $Global:CurrentTheme
    )
    $Button.BackColor = $Colors.ButtonBack
    $Button.ForeColor = $Colors.TextFore
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $Colors.Accent
}

# Helper function for destructive actions that need stronger visual separation
function Set-DestructiveButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [hashtable]$Colors = $Global:CurrentTheme
    )

    Set-ThemedButton -Button $Button -Colors $Colors
    if ($Global:ThemeSelection -eq 'Light') {
        $Button.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 245)
        $Button.ForeColor = [System.Drawing.Color]::FromArgb(125, 30, 30)
        $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(192, 92, 92)
    }
    else {
        $Button.BackColor = $Colors.ButtonBack
        $Button.ForeColor = $Colors.TextFore
        $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(181, 91, 91)
    }
}

# -------------------------
# Auto-Update Functions
# -------------------------
function Get-RemoteVersion {
    try {
        # Fetch only the first part of the remote script to get version
        $response = Invoke-WebRequest -Uri $Global:GitHubRawUrl -UseBasicParsing -TimeoutSec 10
        $content = $response.Content
        if ($content -match '\$Global:ScriptVersion\s*=\s*"([^"]+)"') {
            return $Matches[1]
        }
    }
    catch {
        Write-LogMessage "Failed to check for updates: $($_.Exception.Message)"
    }
    return $null
}

function Test-UpdateAvailable {
    $remoteVersion = Get-RemoteVersion
    if (-not $remoteVersion) { return $null }
    
    try {
        $local = [Version]$Global:ScriptVersion
        $remote = [Version]$remoteVersion
        if ($remote -gt $local) {
            return $remoteVersion
        }
    }
    catch {
        Write-LogMessage "Version comparison failed: $($_.Exception.Message)"
    }
    return $null
}

function Update-Script {
    param([switch]$Restart)
    
    $updateMutex = $null
    $hasUpdateLock = $false

    try {
        $updateMutex = New-Object System.Threading.Mutex($false, "Local\qBitLauncher.Update")
        try {
            $hasUpdateLock = $updateMutex.WaitOne(0)
        }
        catch [System.Threading.AbandonedMutexException] {
            $hasUpdateLock = $true
            Write-LogMessage "Recovered abandoned update lock."
        }

        if (-not $hasUpdateLock) {
            Write-LogMessage "Update skipped because another qBitLauncher instance is already updating."
            return $false
        }

        Write-LogMessage "Downloading update from GitHub..."
        $response = Invoke-WebRequest -Uri $Global:GitHubRawUrl -UseBasicParsing -TimeoutSec 30
        $newContent = $response.Content
        $newContent = $newContent.TrimStart([char]0xFEFF)
        
        # Verify it looks like valid PowerShell
        if ($newContent -notmatch '\$Global:ScriptVersion') {
            throw "Downloaded content doesn't appear to be valid qBitLauncher script"
        }
        
        # Get current script path
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { $scriptPath = $MyInvocation.PSCommandPath }
        if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot "qBitLauncher.ps1" }
        
        # Backup current script
        $backupPath = "$scriptPath.bak"
        Copy-Item -Path $scriptPath -Destination $backupPath -Force
        Write-LogMessage "Backup created: $backupPath"
        
        # Write new content
        [IO.File]::WriteAllText($scriptPath, $newContent, [System.Text.Encoding]::UTF8)
        Write-LogMessage "Script updated successfully!"
        
        # Update the in-memory version so current GUI session shows correct version
        if ($newContent -match '\$Global:ScriptVersion\s*=\s*"([^"]+)"') {
            $Global:ScriptVersion = $Matches[1]
            Write-LogMessage "In-memory version updated to: $($Global:ScriptVersion)"
        }
        
        
        if ($Restart) {
            # Restart the script
            Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" -WindowStyle Hidden
            exit
        }
        return $true
    }
    catch {
        Write-LogMessage "Update failed: $($_.Exception.Message)"
        return $false
    }
    finally {
        if ($hasUpdateLock -and $updateMutex) {
            try { $updateMutex.ReleaseMutex() | Out-Null } catch {}
        }
        if ($updateMutex) {
            $updateMutex.Dispose()
        }
    }
}

function Show-UpdatePrompt {
    param([string]$NewVersion)
    
    $colors = $Global:CurrentTheme
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Update Available"
    $form.Size = New-Object System.Drawing.Size(400, 220)
    $form.StartPosition = 'CenterScreen'
    $form.Add_Load({ Set-FormPositionNearOwner -Form $this })
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Add_Shown({ Set-FormIcon -Form $this })

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(350, 50)
    $label.Text = "A new version of qBitLauncher is available!`n`nCurrent: v$($Global:ScriptVersion)  >>  New: v$NewVersion"
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)

    # Changelog link
    $linkLabel = New-Object System.Windows.Forms.LinkLabel
    $linkLabel.Location = New-Object System.Drawing.Point(20, 75)
    $linkLabel.Size = New-Object System.Drawing.Size(350, 20)
    $linkLabel.Text = "View changelog (commits)"
    $linkLabel.LinkColor = $colors.Accent
    $linkLabel.ActiveLinkColor = $colors.TextFore
    $linkLabel.Add_LinkClicked({ Start-Process $Global:GitHubCommitsUrl })
    $form.Controls.Add($linkLabel)

    $updateBtn = New-Object System.Windows.Forms.Button
    $updateBtn.Location = New-Object System.Drawing.Point(180, 120)
    $updateBtn.Size = New-Object System.Drawing.Size(90, 35)
    $updateBtn.Text = "&Update"
    $updateBtn.DialogResult = [System.Windows.Forms.DialogResult]::Yes

    $skipBtn = New-Object System.Windows.Forms.Button
    $skipBtn.Location = New-Object System.Drawing.Point(280, 120)
    $skipBtn.Size = New-Object System.Drawing.Size(90, 35)
    $skipBtn.Text = "&Skip"
    $skipBtn.DialogResult = [System.Windows.Forms.DialogResult]::No

    foreach ($btn in @($updateBtn, $skipBtn)) {
        Set-ThemedButton -Button $btn -Colors $colors
        $form.Controls.Add($btn)
    }

    $form.AcceptButton = $updateBtn
    $form.CancelButton = $skipBtn

    $result = $form.ShowDialog()
    $form.Dispose()

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Update script but don't restart - let GUI continue normally
        Update-Script
    }
}

# -------------------------
# GUI: Theme and Color Definitions
# -------------------------
# Set the desired theme here: 'Dracula', or 'Light'
$Global:ThemeSelection = 'Dracula' 
$Global:Themes = @{
    Light   = @{
        FormBack      = [System.Drawing.Color]::FromArgb(240, 240, 240)
        TextFore      = [System.Drawing.Color]::Black
        ControlBack   = [System.Drawing.Color]::White
        ButtonBack    = [System.Drawing.Color]::FromArgb(225, 225, 225)
        Border        = [System.Drawing.Color]::DimGray
        Accent        = [System.Drawing.Color]::DodgerBlue
        SecondaryText = [System.Drawing.Color]::FromArgb(100, 100, 100)
    }
    Dracula = @{
        FormBack      = [System.Drawing.Color]::FromArgb(40, 42, 54)   # Background #282A36 (Shadow Grey)
        TextFore      = [System.Drawing.Color]::FromArgb(248, 248, 242) # Foreground #F8F8F2
        ControlBack   = [System.Drawing.Color]::FromArgb(32, 32, 32)   # Carbon Black #202020
        ButtonBack    = [System.Drawing.Color]::FromArgb(47, 52, 64)   # Jet Black #2F3440
        Border        = [System.Drawing.Color]::FromArgb(68, 71, 90)   # Current Line #44475A
        Accent        = [System.Drawing.Color]::FromArgb(98, 114, 164)  # Comment #6272A4
        SecondaryText = [System.Drawing.Color]::FromArgb(150, 150, 150)  # Muted text
    }
}
$Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]

# -------------------------
# Helper: Logging (Verb-Noun: Write-LogMessage)
# -------------------------
function Write-LogMessage {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd hh:mm tt"
    $LogEntry = "$Timestamp - $Message"
    try {
        # Prepend new entries to top of log file (newest first)
        $existingContent = if (Test-Path $LogFile) { Get-Content $LogFile -Raw -ErrorAction SilentlyContinue } else { "" }
        $newContent = "$LogEntry`r`n$existingContent"
        [System.IO.File]::WriteAllText($LogFile, $newContent)
    }
    catch {
        $FallbackLogDir = Join-Path $env:PUBLIC "Documents"; $FallbackLogFile = Join-Path $FallbackLogDir "qBitLauncher_fallback_log.txt"
        try { if (-not (Test-Path $FallbackLogDir)) { New-Item -ItemType Directory -Path $FallbackLogDir -Force -ErrorAction SilentlyContinue | Out-Null }; Add-Content -Path $FallbackLogFile -Value "$Timestamp - FALLBACK: $Message (Original log failed: $($_.Exception.Message))" -ErrorAction SilentlyContinue } catch {}
        Write-Warning "Failed to write to primary log file: $LogFile. Error: $($_.Exception.Message)"
    }
}

Write-LogMessage "--------------------------------------------------------"
Write-LogMessage "Script started: qBitLauncher.ps1"
Write-LogMessage "Received initial path from qBittorrent: '$filePathFromQB'"
if (-not [string]::IsNullOrWhiteSpace($torrentHashFromQB)) {
    Write-LogMessage "Received torrent hash from qBittorrent: '$torrentHashFromQB'"
}
if (-not [string]::IsNullOrWhiteSpace($torrentNameFromQB)) {
    Write-LogMessage "Received torrent name from qBittorrent: '$torrentNameFromQB'"
}
# Write-Host suppressed for PS2EXE compatibility

# -------------------------
# Configuration File System
# -------------------------
$Global:ConfigFile = Join-Path $Global:ScriptDir "config.json"
$Global:UserSettings = @{
    Theme                 = "Dracula"
    QbittorrentWebUrl     = "http://localhost:8080"
    QbittorrentUsername   = ""
    QbittorrentPassword   = ""
}

function Get-UserSettings {
    if (Test-Path $Global:ConfigFile) {
        try {
            $json = Get-Content $Global:ConfigFile -Raw | ConvertFrom-Json
            $Global:UserSettings.Theme = if ($json.Theme) { $json.Theme } else { "Dracula" }
            $Global:UserSettings.QbittorrentWebUrl = if ($json.QbittorrentWebUrl) { $json.QbittorrentWebUrl } else { "http://localhost:8080" }
            $Global:UserSettings.QbittorrentUsername = if ($json.QbittorrentUsername) { $json.QbittorrentUsername } else { "" }
            $Global:UserSettings.QbittorrentPassword = if ($json.QbittorrentPassword) { $json.QbittorrentPassword } else { "" }
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
# Helper: Play Sound Effect (Verb-Noun: Invoke-ActionSound)
# -------------------------
function Invoke-ActionSound {
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
# Helper: Center a form on the same monitor as the main form
# -------------------------
function Set-FormPositionNearOwner {
    param([System.Windows.Forms.Form]$Form)

    $owner = $Global:MainForm
    if ($owner -and $owner.IsHandleCreated) {
        $screen = [System.Windows.Forms.Screen]::FromControl($owner)
    }
    else {
        $screen = [System.Windows.Forms.Screen]::FromPoint([System.Windows.Forms.Cursor]::Position)
    }

    $wa = $screen.WorkingArea
    $x = $wa.X + [math]::Max(0, ($wa.Width  - $Form.Width)  / 2)
    $y = $wa.Y + [math]::Max(0, ($wa.Height - $Form.Height) / 2)
    $Form.StartPosition = 'Manual'
    $Form.Location = New-Object System.Drawing.Point([int]$x, [int]$y)
}

# -------------------------
# Helper: Themed Message Box (replaces standard MessageBox)
# -------------------------
function Show-ThemedMessageBox {
    param(
        [string]$Message,
        [string]$Title = "qBitLauncher",
        [ValidateSet('OK', 'OKCancel', 'YesNo', 'YesNoCancel')]
        [string]$Buttons = 'OK',
        [ValidateSet('None', 'Information', 'Warning', 'Error', 'Question')]
        [string]$Icon = 'None'
    )
    
    $colors = $Global:CurrentTheme
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(400, 180)
    $form.StartPosition = 'CenterScreen'
    $form.Add_Load({ Set-FormPositionNearOwner -Form $this })
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Add_Shown({ Set-FormIcon -Form $this })
    
    # Icon
    $iconSize = 32
    $iconLeft = 20
    $textLeft = 65
    
    if ($Icon -ne 'None') {
        $iconPicture = New-Object System.Windows.Forms.PictureBox
        $iconPicture.Location = New-Object System.Drawing.Point($iconLeft, 25)
        $iconPicture.Size = New-Object System.Drawing.Size($iconSize, $iconSize)
        $iconPicture.SizeMode = 'CenterImage'
        
        $systemIcon = switch ($Icon) {
            'Information' { [System.Drawing.SystemIcons]::Information }
            'Warning' { [System.Drawing.SystemIcons]::Warning }
            'Error' { [System.Drawing.SystemIcons]::Error }
            'Question' { [System.Drawing.SystemIcons]::Question }
            default { $null }
        }
        
        if ($systemIcon) {
            $iconPicture.Image = $systemIcon.ToBitmap()
        }
        $form.Controls.Add($iconPicture)
    }
    else {
        $textLeft = 20
    }
    
    # Message label
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($textLeft, 25)
    $label.Size = New-Object System.Drawing.Size((350 - $textLeft), 60)
    $label.Text = $Message
    $label.UseMnemonic = $false
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)
    
    # Auto-size form height based on message length
    $textFlags = [System.Windows.Forms.TextFormatFlags]::WordBreak -bor [System.Windows.Forms.TextFormatFlags]::NoPrefix
    $textSize = [System.Windows.Forms.TextRenderer]::MeasureText($Message, $form.Font, [System.Drawing.Size]::new((350 - $textLeft), 0), $textFlags)
    $requiredHeight = [Math]::Max(180, $textSize.Height + 130)
    $form.Size = New-Object System.Drawing.Size(400, $requiredHeight)
    $label.Size = New-Object System.Drawing.Size((350 - $textLeft), ($requiredHeight - 120))
    
    # Buttons
    $buttonY = $requiredHeight - 80
    $buttonWidth = 90
    $buttonHeight = 35
    
    switch ($Buttons) {
        'OK' {
            $okBtn = New-Object System.Windows.Forms.Button
            $okBtn.Location = New-Object System.Drawing.Point(150, $buttonY)
            $okBtn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $okBtn.Text = "&OK"
            $okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
            Set-ThemedButton -Button $okBtn -Colors $colors
            $form.Controls.Add($okBtn)
            $form.AcceptButton = $okBtn
        }
        'OKCancel' {
            $okBtn = New-Object System.Windows.Forms.Button
            $okBtn.Location = New-Object System.Drawing.Point(100, $buttonY)
            $okBtn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $okBtn.Text = "&OK"
            $okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
            Set-ThemedButton -Button $okBtn -Colors $colors
            $form.Controls.Add($okBtn)
            
            $cancelBtn = New-Object System.Windows.Forms.Button
            $cancelBtn.Location = New-Object System.Drawing.Point(200, $buttonY)
            $cancelBtn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $cancelBtn.Text = "&Cancel"
            $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            Set-ThemedButton -Button $cancelBtn -Colors $colors
            $form.Controls.Add($cancelBtn)
            
            $form.AcceptButton = $okBtn
            $form.CancelButton = $cancelBtn
        }
        'YesNo' {
            $yesBtn = New-Object System.Windows.Forms.Button
            $yesBtn.Location = New-Object System.Drawing.Point(100, $buttonY)
            $yesBtn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $yesBtn.Text = "&Yes"
            $yesBtn.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            Set-ThemedButton -Button $yesBtn -Colors $colors
            $form.Controls.Add($yesBtn)
            
            $noBtn = New-Object System.Windows.Forms.Button
            $noBtn.Location = New-Object System.Drawing.Point(200, $buttonY)
            $noBtn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $noBtn.Text = "&No"
            $noBtn.DialogResult = [System.Windows.Forms.DialogResult]::No
            Set-ThemedButton -Button $noBtn -Colors $colors
            $form.Controls.Add($noBtn)
            
            $form.AcceptButton = $yesBtn
            $form.CancelButton = $noBtn
        }
        'YesNoCancel' {
            $yesBtn = New-Object System.Windows.Forms.Button
            $yesBtn.Location = New-Object System.Drawing.Point(55, $buttonY)
            $yesBtn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $yesBtn.Text = "&Yes"
            $yesBtn.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            Set-ThemedButton -Button $yesBtn -Colors $colors
            $form.Controls.Add($yesBtn)
            
            $noBtn = New-Object System.Windows.Forms.Button
            $noBtn.Location = New-Object System.Drawing.Point(155, $buttonY)
            $noBtn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $noBtn.Text = "&No"
            $noBtn.DialogResult = [System.Windows.Forms.DialogResult]::No
            Set-ThemedButton -Button $noBtn -Colors $colors
            $form.Controls.Add($noBtn)
            
            $cancelBtn = New-Object System.Windows.Forms.Button
            $cancelBtn.Location = New-Object System.Drawing.Point(255, $buttonY)
            $cancelBtn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $cancelBtn.Text = "&Cancel"
            $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            Set-ThemedButton -Button $cancelBtn -Colors $colors
            $form.Controls.Add($cancelBtn)
            
            $form.AcceptButton = $yesBtn
            $form.CancelButton = $cancelBtn
        }
    }
    
    $owner = $Global:MainForm
    $result = if ($owner) { $form.ShowDialog($owner) } else { $form.ShowDialog() }
    $form.Dispose()
    return $result
}

# -------------------------
# GUI: qBittorrent Removal Confirmation
# -------------------------
function Show-CleanupConfirmForm {
    param(
        [string]$TorrentLabel,
        [string]$TorrentHash,
        [string]$FilePath,
        [string]$SourcePath,
        [bool]$SourceCanDelete = $false
    )

    $colors = $Global:CurrentTheme
    $hasTorrent = -not [string]::IsNullOrWhiteSpace($TorrentHash)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Cleanup"
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Add_Shown({ Set-FormIcon -Form $this })

    # Tooltip provider for hover info
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.InitialDelay = 300
    $toolTip.ReshowDelay = 200

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(24, 18)
    $titleLabel.AutoSize = $true
    $titleLabel.Text = "What would you like to clean up?"
    $titleLabel.ForeColor = $colors.TextFore
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLabel)

    # Build info text — show file location, torrent info, and source path
    $infoLines = @()
    if (-not [string]::IsNullOrWhiteSpace($FilePath)) {
        $infoLines += "File Location: $FilePath"
    }
    if ($hasTorrent) {
        if (-not [string]::IsNullOrWhiteSpace($TorrentLabel)) {
            $infoLines += "Torrent: $TorrentLabel"
        }
        $infoLines += "Hash: $TorrentHash"
    }
    if ($SourceCanDelete -and -not [string]::IsNullOrWhiteSpace($SourcePath)) {
        $sourceFileName = [System.IO.Path]::GetFileName($SourcePath.TrimEnd('\'))
        $infoLines += "Source: $sourceFileName"
    }
    $infoText = ($infoLines -join "`n")

    # Measure info text to auto-size the details area
    $measureFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $maxDetailsWidth = 640
    $textFlags = [System.Windows.Forms.TextFormatFlags]::WordBreak -bor [System.Windows.Forms.TextFormatFlags]::NoPrefix
    $measuredSize = [System.Windows.Forms.TextRenderer]::MeasureText($infoText, $measureFont, [System.Drawing.Size]::new($maxDetailsWidth, 0), $textFlags)
    $detailsHeight = [Math]::Max(50, $measuredSize.Height + 10)

    $detailsLabel = New-Object System.Windows.Forms.Label
    $detailsLabel.Location = New-Object System.Drawing.Point(24, 50)
    $detailsLabel.Size = New-Object System.Drawing.Size($maxDetailsWidth, $detailsHeight)
    $detailsLabel.Text = $infoText
    $detailsLabel.UseMnemonic = $false
    $detailsLabel.ForeColor = $colors.SecondaryText
    $form.Controls.Add($detailsLabel)

    $selection = @{ Value = "Cancel" }

    # Layout: buttons start below the details area
    $buttonY1 = 50 + $detailsHeight + 16
    $buttonY2 = $buttonY1 + 46
    $btnHeight = 38
    $btnSpacing = 12

    # --- Row 1: Torrent options (evenly spaced across the form) ---
    $removeTorrentFilesButton = New-Object System.Windows.Forms.Button
    $removeTorrentFilesButton.Location = New-Object System.Drawing.Point(24, $buttonY1)
    $removeTorrentFilesButton.Size = New-Object System.Drawing.Size(230, $btnHeight)
    $removeTorrentFilesButton.Text = "Remove &Torrent + Delete Files"
    Set-DestructiveButton -Button $removeTorrentFilesButton -Colors $colors
    $removeTorrentFilesButton.Add_Click({
            $selection.Value = "DeleteTorrentAndFiles"
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            $form.Close()
        }.GetNewClosure())
    $form.Controls.Add($removeTorrentFilesButton)

    $removeTorrentOnlyButton = New-Object System.Windows.Forms.Button
    $removeTorrentOnlyButton.Location = New-Object System.Drawing.Point((24 + 230 + $btnSpacing), $buttonY1)
    $removeTorrentOnlyButton.Size = New-Object System.Drawing.Size(200, $btnHeight)
    $removeTorrentOnlyButton.Text = "Remove Torrent &Only"
    Set-ThemedButton -Button $removeTorrentOnlyButton -Colors $colors
    $removeTorrentOnlyButton.Add_Click({
            $selection.Value = "DeleteTorrentOnly"
            $form.DialogResult = [System.Windows.Forms.DialogResult]::No
            $form.Close()
        }.GetNewClosure())
    $form.Controls.Add($removeTorrentOnlyButton)

    # --- Row 2: Source deletion + Cancel ---
    $deleteSourceButton = New-Object System.Windows.Forms.Button
    $deleteSourceButton.Location = New-Object System.Drawing.Point(24, $buttonY2)
    $deleteSourceButton.Size = New-Object System.Drawing.Size(160, $btnHeight)
    $deleteSourceButton.Text = "Delete &Source"
    if ($SourceCanDelete) {
        Set-DestructiveButton -Button $deleteSourceButton -Colors $colors
        # Show full source path on hover instead of cramming it into the button
        $toolTip.SetToolTip($deleteSourceButton, $SourcePath)
    }
    else {
        Set-ThemedButton -Button $deleteSourceButton -Colors $colors
        $deleteSourceButton.Enabled = $false
    }
    $deleteSourceButton.Add_Click({
            $selection.Value = "DeleteSource"
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Retry
            $form.Close()
        }.GetNewClosure())
    $form.Controls.Add($deleteSourceButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Size = New-Object System.Drawing.Size(110, $btnHeight)
    $cancelButton.Text = "&Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    Set-ThemedButton -Button $cancelButton -Colors $colors
    $form.Controls.Add($cancelButton)

    $form.CancelButton = $cancelButton

    # --- Auto-size form to fit content ---
    $formContentWidth = 24 + 230 + $btnSpacing + 200 + 24
    $formHeight = $buttonY2 + $btnHeight + 50  # padding below last row + title bar
    $form.ClientSize = New-Object System.Drawing.Size($formContentWidth, ($buttonY2 + $btnHeight + 20))

    # Position cancel button at the right edge
    $cancelButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 110 - 24), $buttonY2)

    $owner = $Global:MainForm
    $dialogResult = if ($owner) { $form.ShowDialog($owner) } else { $form.ShowDialog() }
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
        $selection.Value = "Cancel"
    }
    $toolTip.Dispose()
    $form.Dispose()

    return $selection.Value
}


# -------------------------
# Helper: qBittorrent Local API
# -------------------------
function Remove-QBittorrentTorrent {
    param(
        [string]$TorrentHash,
        [bool]$DeleteFiles
    )

    if ([string]::IsNullOrWhiteSpace($TorrentHash)) {
        throw "No qBittorrent torrent hash was provided. Update the qBittorrent command to pass %I."
    }

    $baseUrl = $Global:UserSettings.QbittorrentWebUrl
    if ([string]::IsNullOrWhiteSpace($baseUrl)) {
        throw "qBittorrent local API URL is empty. Set it in Settings."
    }

    $baseUrl = $baseUrl.Trim().TrimEnd('/')
    $parsedUrl = $null
    if (-not [System.Uri]::TryCreate($baseUrl, [System.UriKind]::Absolute, [ref]$parsedUrl)) {
        throw "qBittorrent local API URL is invalid: $baseUrl"
    }

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

    if (-not [string]::IsNullOrWhiteSpace($Global:UserSettings.QbittorrentUsername)) {
        $loginUri = "$baseUrl/api/v2/auth/login"
        $loginBody = @{
            username = $Global:UserSettings.QbittorrentUsername
            password = $Global:UserSettings.QbittorrentPassword
        }

        $loginResponse = Invoke-WebRequest -Uri $loginUri -Method Post -Body $loginBody -ContentType 'application/x-www-form-urlencoded' -WebSession $session -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop
        if ($loginResponse.Content -notmatch '^Ok\.?$') {
            throw "qBittorrent login failed. Check the local API username and password in Settings."
        }
    }

    $deleteUri = "$baseUrl/api/v2/torrents/delete"
    $deleteBody = @{
        hashes = $TorrentHash.Trim()
    }
    if ($DeleteFiles) {
        $deleteBody.Add('deleteFiles', 'true')
    }

    try {
        Invoke-WebRequest -Uri $deleteUri -Method Post -Body $deleteBody -ContentType 'application/x-www-form-urlencoded' -WebSession $session -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop | Out-Null
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match 'Unable to connect|actively refused|timed out|No connection could be made|remote server') {
            throw "Could not connect to qBittorrent's local API at $baseUrl."
        }
        if ($message -match 'Unauthorized|Forbidden|401|403') {
            throw "qBittorrent rejected the request. Enter the local API username/password in qBitLauncher Settings."
        }
        throw
    }
}

function Get-QBittorrentApiBaseUrl {
    $baseUrl = $Global:UserSettings.QbittorrentWebUrl
    if ([string]::IsNullOrWhiteSpace($baseUrl)) {
        return "http://localhost:8080"
    }

    return $baseUrl.Trim().TrimEnd('/')
}

function Get-QBittorrentApiUri {
    $baseUrl = Get-QBittorrentApiBaseUrl
    $parsedUrl = $null
    if ([System.Uri]::TryCreate($baseUrl, [System.UriKind]::Absolute, [ref]$parsedUrl)) {
        return $parsedUrl
    }

    return $null
}

function Test-QBittorrentApiUrlIsLocal {
    $parsedUrl = Get-QBittorrentApiUri
    return ($parsedUrl -and $parsedUrl.IsLoopback)
}

function Get-QBittorrentApiPort {
    $parsedUrl = Get-QBittorrentApiUri
    if ($parsedUrl -and $parsedUrl.Port -gt 0) {
        return $parsedUrl.Port
    }

    return 8080
}

function Get-QBittorrentConfigPath {
    if ([string]::IsNullOrWhiteSpace($env:APPDATA)) {
        return $null
    }

    return (Join-Path $env:APPDATA "qBittorrent\qBittorrent.ini")
}

function Get-QBittorrentWebApiState {
    $configPath = Get-QBittorrentConfigPath
    $state = [pscustomobject]@{
        ConfigPath  = $configPath
        Exists      = $false
        Enabled     = $null
        Port        = $null
        CloseToTray = $null
    }

    if ([string]::IsNullOrWhiteSpace($configPath) -or -not (Test-Path -LiteralPath $configPath)) {
        return $state
    }

    $state.Exists = $true
    try {
        $lines = Get-Content -LiteralPath $configPath -ErrorAction Stop
        foreach ($line in $lines) {
            if ($line -match '^WebUI\\Enabled=(.+)$') {
                $state.Enabled = ($Matches[1].Trim().ToLowerInvariant() -eq 'true')
            }
            elseif ($line -match '^WebUI\\Port=(\d+)$') {
                $state.Port = [int]$Matches[1]
            }
            elseif ($line -match '^General\\CloseToTray=(.+)$') {
                $state.CloseToTray = ($Matches[1].Trim().ToLowerInvariant() -eq 'true')
            }
        }
    }
    catch {
        Write-LogMessage "Failed to inspect qBittorrent config: $($_.Exception.Message)"
    }

    return $state
}

function Set-QBittorrentIniValue {
    param(
        [string[]]$Lines,
        [string]$Key,
        [string]$Value
    )

    $list = New-Object System.Collections.Generic.List[string]
    $pattern = '^' + [regex]::Escape($Key) + '='
    
    foreach ($line in $Lines) {
        if ($line -notmatch $pattern) {
            $list.Add($line)
        }
    }

    $resultArray = $list.ToArray()
    $prefIndex = [array]::IndexOf($resultArray, '[Preferences]')
    if ($prefIndex -ge 0) {
        $list.Insert($prefIndex + 1, "$Key=$Value")
    }
    else {
        $list.Add('[Preferences]')
        $list.Add("$Key=$Value")
    }

    return ,$list.ToArray()
}

function Enable-QBittorrentLocalApiConfig {
    param(
        [int]$Port = 8080
    )

    $state = Get-QBittorrentWebApiState
    if (-not $state.Exists) {
        throw "qBittorrent config was not found at $($state.ConfigPath). Start qBittorrent once, then try again."
    }

    $lines = @(Get-Content -LiteralPath $state.ConfigPath -ErrorAction Stop)
    $backupPath = "$($state.ConfigPath).qBitLauncher.bak"
    Copy-Item -LiteralPath $state.ConfigPath -Destination $backupPath -Force -ErrorAction Stop

    $lines = Set-QBittorrentIniValue -Lines $lines -Key 'WebUI\Enabled' -Value 'true'
    $lines = Set-QBittorrentIniValue -Lines $lines -Key 'WebUI\Port' -Value ([string]$Port)
    $lines = Set-QBittorrentIniValue -Lines $lines -Key 'WebUI\HTTPS\Enabled' -Value 'false'
    $lines = Set-QBittorrentIniValue -Lines $lines -Key 'WebUI\LocalHostAuth' -Value 'false'
    $lines = Set-QBittorrentIniValue -Lines $lines -Key 'WebUI\Username' -Value 'admin'
    $lines = Set-QBittorrentIniValue -Lines $lines -Key 'WebUI\Password_PBKDF2' -Value '"@ByteArray(ARQ77eY1NUZaQsuDHbIMCA==:0WMRkYTUWVT9wVvdDtHAjU9b3b7uB8NR1Gur2hmQCvCDpm39Q+PsJRJPaCU51dEiz+dTzh8qbPsL8WkFljQYFQ==)"'

    [System.IO.File]::WriteAllLines($state.ConfigPath, [string[]]$lines, [System.Text.Encoding]::UTF8)
    Write-LogMessage "Enabled qBittorrent local API in config: $($state.ConfigPath)"
    return $backupPath
}

function Get-QBittorrentExecutablePath {
    try {
        $runningProcess = Get-Process -Name qbittorrent -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($runningProcess -and $runningProcess.Path -and (Test-Path -LiteralPath $runningProcess.Path)) {
            return $runningProcess.Path
        }
    }
    catch {
        Write-LogMessage "Could not read running qBittorrent path: $($_.Exception.Message)"
    }

    try {
        $command = Get-Command qbittorrent.exe -ErrorAction SilentlyContinue
        if ($command -and $command.Source -and (Test-Path -LiteralPath $command.Source)) {
            return $command.Source
        }
    }
    catch {}

    $candidates = @(
        (Join-Path $env:ProgramFiles "qBittorrent\qbittorrent.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "qBittorrent\qbittorrent.exe")
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Test-QBittorrentProcessRunning {
    return (@(Get-Process -Name qbittorrent -ErrorAction SilentlyContinue).Count -gt 0)
}

function Stop-QBittorrentGracefully {
    param(
        [int]$TimeoutSeconds = 15
    )

    $processes = @(Get-Process -Name qbittorrent -ErrorAction SilentlyContinue)
    if ($processes.Count -eq 0) {
        return @{ Closed = $true; WasRunning = $false }
    }

    foreach ($process in $processes) {
        try {
            [void]$process.CloseMainWindow()
        }
        catch {
            Write-LogMessage "Could not ask qBittorrent to close: $($_.Exception.Message)"
        }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
        $remaining = @(Get-Process -Name qbittorrent -ErrorAction SilentlyContinue)
        if ($remaining.Count -eq 0) {
            return @{ Closed = $true; WasRunning = $true }
        }
    } while ((Get-Date) -lt $deadline)

    # Graceful close failed (elevated process or CloseToTray). Try elevated taskkill.
    Write-LogMessage "Graceful close timed out. Attempting elevated taskkill."
    try {
        $killProc = Start-Process -FilePath "taskkill" -ArgumentList "/F /IM qbittorrent.exe" -Verb RunAs -Wait -PassThru -WindowStyle Hidden
        Start-Sleep -Seconds 2
        $remaining = @(Get-Process -Name qbittorrent -ErrorAction SilentlyContinue)
        if ($remaining.Count -eq 0) {
            Write-LogMessage "Elevated taskkill succeeded."
            return @{ Closed = $true; WasRunning = $true }
        }
    }
    catch {
        Write-LogMessage "Elevated taskkill failed: $($_.Exception.Message)"
    }

    return @{ Closed = $false; WasRunning = $true }
}

function Show-QBittorrentExitWaitForm {
    param(
        [bool]$CloseToTray = $false
    )

    $colors = $Global:CurrentTheme
    $result = @{ Value = "Cancel" }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Waiting for qBittorrent"
    $form.Size = New-Object System.Drawing.Size(540, 265)
    $form.StartPosition = 'CenterScreen'
    $form.Add_Load({ Set-FormPositionNearOwner -Form $this })
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Add_Shown({ Set-FormIcon -Form $this })

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(22, 18)
    $titleLabel.Size = New-Object System.Drawing.Size(480, 25)
    $titleLabel.Text = "qBitLauncher is still holding this torrent session."
    $titleLabel.ForeColor = $colors.TextFore
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLabel)

    $messageLabel = New-Object System.Windows.Forms.Label
    $messageLabel.Location = New-Object System.Drawing.Point(22, 52)
    $messageLabel.Size = New-Object System.Drawing.Size(480, 92)
    $trayText = if ($CloseToTray) { "qBittorrent is set to close to the tray, so use its tray icon/menu and choose Exit.`n`n" } else { "" }
    $messageLabel.Text = "${trayText}Exit qBittorrent once. As soon as it is closed, qBitLauncher will enable cleanup and retry Remove Torrent for this same game."
    $messageLabel.UseMnemonic = $false
    $messageLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($messageLabel)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(22, 150)
    $statusLabel.Size = New-Object System.Drawing.Size(480, 24)
    $statusLabel.Text = "Waiting for qBittorrent to exit..."
    $statusLabel.ForeColor = $colors.SecondaryText
    $form.Controls.Add($statusLabel)

    $checkButton = New-Object System.Windows.Forms.Button
    $checkButton.Location = New-Object System.Drawing.Point(265, 185)
    $checkButton.Size = New-Object System.Drawing.Size(115, 35)
    $checkButton.Text = "Check &Again"
    Set-ThemedButton -Button $checkButton -Colors $colors
    $checkButton.Add_Click({
            if (-not (Test-QBittorrentProcessRunning)) {
                $result.Value = "Exited"
                $form.Close()
            }
            else {
                $statusLabel.Text = "Still running. Exit qBittorrent from the tray/menu."
                Invoke-ActionSound -Type Notify
            }
        }.GetNewClosure())
    $form.Controls.Add($checkButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(390, 185)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 35)
    $cancelButton.Text = "&Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    Set-ThemedButton -Button $cancelButton -Colors $colors
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    $waitSeconds = @{ Value = 0 }
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
            if (-not (Test-QBittorrentProcessRunning)) {
                $result.Value = "Exited"
                $timer.Stop()
                $form.Close()
                return
            }

            $waitSeconds.Value++
            $statusLabel.Text = "Still running. Waiting... $($waitSeconds.Value)s"
        }.GetNewClosure())

    $form.Add_Shown({
            if (-not (Test-QBittorrentProcessRunning)) {
                $result.Value = "Exited"
                $form.Close()
                return
            }
            $timer.Start()
        }.GetNewClosure())
    $form.Add_FormClosed({
            $timer.Stop()
            $timer.Dispose()
        }.GetNewClosure())

    $owner = $Global:MainForm
    if ($owner) { [void]$form.ShowDialog($owner) } else { [void]$form.ShowDialog() }
    $form.Dispose()
    return $result.Value
}

function Start-QBittorrentIfAvailable {
    param(
        [string]$ExecutablePath
    )

    if ([string]::IsNullOrWhiteSpace($ExecutablePath) -or -not (Test-Path -LiteralPath $ExecutablePath)) {
        return $false
    }

    Start-Process -FilePath $ExecutablePath | Out-Null
    Write-LogMessage "Started qBittorrent: $ExecutablePath"
    return $true
}

function Test-QBittorrentApiEndpointAvailable {
    $baseUrl = Get-QBittorrentApiBaseUrl
    try {
        Invoke-WebRequest -Uri "$baseUrl/api/v2/app/version" -Method Get -UseBasicParsing -TimeoutSec 3 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        $response = $_.Exception.Response
        if ($response -and ($response.StatusCode -eq 401 -or $response.StatusCode -eq 403)) {
            return $true
        }
        return $false
    }
}

function Wait-QBittorrentApiEndpoint {
    param(
        [int]$TimeoutSeconds = 12
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        if (Test-QBittorrentApiEndpointAvailable) {
            return $true
        }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 700
    } while ((Get-Date) -lt $deadline)

    return $false
}

function Test-QBittorrentConnectionFailureMessage {
    param([string]$Message)
    return ($Message -match 'Could not connect|Unable to connect|actively refused|timed out|No connection could be made|remote server')
}

function Invoke-QBittorrentLocalApiRepairFlow {
    param(
        [string]$ErrorMessage,
        [scriptblock]$RetryAction
    )

    if (-not (Test-QBittorrentConnectionFailureMessage -Message $ErrorMessage)) {
        return @{ Status = "NotApplicable"; Message = $ErrorMessage }
    }

    if (-not (Test-QBittorrentApiUrlIsLocal)) {
        return @{ Status = "NotApplicable"; Message = "qBitLauncher can only auto-enable qBittorrent's API for localhost URLs. Current URL: $(Get-QBittorrentApiBaseUrl)" }
    }

    $state = Get-QBittorrentWebApiState
    if (-not $state.Exists) {
        return @{ Status = "NotApplicable"; Message = "Could not find qBittorrent.ini at $($state.ConfigPath). Start qBittorrent once, then try again." }
    }

    # If the WebUI is enabled in the config but the API is not responding,
    # try to start qBittorrent automatically and wait for the API to come up.
    if ($state.Enabled -eq $true -and -not (Test-QBittorrentApiEndpointAvailable)) {
        Write-LogMessage "qBittorrent WebUI is enabled in config but API not responding. Attempting to start qBittorrent process."
        $exePath = Get-QBittorrentExecutablePath
        if ($exePath) {
            Write-LogMessage "Attempting to start qBittorrent at: $exePath"
            Start-QBittorrentIfAvailable -ExecutablePath $exePath | Out-Null
            if (Wait-QBittorrentApiEndpoint -TimeoutSeconds 15) {
                try {
                    & $RetryAction
                    return @{ Status = "RetrySucceeded"; Message = "qBittorrent was started and the API responded. Retry succeeded." }
                }
                catch {
                    Write-LogMessage "Retry after auto-start failed: $($_.Exception.Message)"
                    return @{ Status = "RetryFailed"; Message = $_.Exception.Message }
                }
            }
            else {
                Write-LogMessage "qBittorrent auto-started but API did not respond within timeout. Proceeding to config repair."
            }
        }
        else {
            Write-LogMessage "Could not determine qBittorrent executable path to auto-start."
        }
    }

    $port = Get-QBittorrentApiPort
    $enableResult = Show-ThemedMessageBox -Message "qBittorrent's local API is turned off.`n`nqBitLauncher can enable it on localhost port $port and then retry removal. This is a one-time setup and it does not open a browser page.`n`nEnable it now?" -Title "Set Up qBittorrent Cleanup" -Buttons 'YesNo' -Icon 'Question'
    if ($enableResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return @{ Status = "Declined"; Message = "qBittorrent local API setup was cancelled." }
    }

    $exePath = Get-QBittorrentExecutablePath
    $closeResult = Stop-QBittorrentGracefully -TimeoutSeconds 5
    if (-not $closeResult.Closed) {
        $waitResult = Show-QBittorrentExitWaitForm -CloseToTray ($state.CloseToTray -eq $true)
        if ($waitResult -ne "Exited") {
            return @{ Status = "PendingExit"; Message = "qBittorrent local API setup was cancelled while waiting for qBittorrent to exit." }
        }
    }

    try {
        $backupPath = Enable-QBittorrentLocalApiConfig -Port $port
        if ($backupPath) {
            Write-LogMessage "qBittorrent config backup created: $backupPath"
        }
    }
    catch {
        return @{ Status = "Failed"; Message = $_.Exception.Message }
    }

    [void](Start-QBittorrentIfAvailable -ExecutablePath $exePath)
    if (-not (Wait-QBittorrentApiEndpoint -TimeoutSeconds 15)) {
        Show-ThemedMessageBox -Message "qBitLauncher enabled qBittorrent's local API, but it is not responding yet.`n`nStart or reopen qBittorrent, then click Remove Torrent again." -Title "qBittorrent Setup Saved" -Icon 'Information'
        return @{ Status = "PendingRestart"; Message = "qBittorrent local API was enabled, but qBittorrent needs to be reopened before cleanup can run." }
    }

    try {
        & $RetryAction
        return @{ Status = "RetrySucceeded"; Message = "qBittorrent local API was enabled and the torrent was removed." }
    }
    catch {
        return @{ Status = "RetryFailed"; Message = $_.Exception.Message }
    }
}

# -------------------------
# Helper: Validate Extraction Path (Verb-Noun: Test-ExtractionPath)
# -------------------------
function Test-ExtractionPath {
    param(
        [string]$Path,
        [long]$RequiredSpaceBytes = 0
    )
    
    # Check if path is empty
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @{ Valid = $false; Error = "Please enter a destination path." }
    }
    
    # Check if path format is valid and drive exists
    try {
        $drive = [System.IO.Path]::GetPathRoot($Path)
        if (-not $drive -or -not (Test-Path $drive)) {
            return @{ Valid = $false; Error = "Drive '$drive' does not exist or is not accessible." }
        }
    }
    catch {
        return @{ Valid = $false; Error = "Invalid path format: $($_.Exception.Message)" }
    }
    
    # Check disk space if required
    if ($RequiredSpaceBytes -gt 0) {
        try {
            $driveLetter = $drive[0]
            $driveInfo = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue
            if ($driveInfo -and $driveInfo.Free -lt $RequiredSpaceBytes) {
                $freeGB = [math]::Round($driveInfo.Free / 1GB, 2)
                $reqGB = [math]::Round($RequiredSpaceBytes / 1GB, 2)
                return @{ Valid = $false; Error = "Insufficient disk space. Need approximately ${reqGB}GB, only ${freeGB}GB available." }
            }
        }
        catch {
            Write-LogMessage "Could not check disk space: $($_.Exception.Message)"
        }
    }
    
    # Check write permission
    try {
        if (Test-Path $Path) {
            # Directory exists, try to create a temp file
            $testFile = Join-Path $Path ".qbit_write_test_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
            [System.IO.File]::WriteAllText($testFile, "test")
            Remove-Item $testFile -Force -ErrorAction SilentlyContinue
        }
        else {
            # Directory doesn't exist, check if we can create it
            $parent = Split-Path $Path -Parent
            if ($parent -and (Test-Path $parent)) {
                $testFile = Join-Path $parent ".qbit_write_test_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
                [System.IO.File]::WriteAllText($testFile, "test")
                Remove-Item $testFile -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        return @{ Valid = $false; Error = "Cannot write to this location. It may be read-only or require administrator privileges." }
    }
    
    return @{ Valid = $true; Error = $null }
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
# Helper: Find 7-Zip GUI extractor (Verb-Noun: Get-7ZipGuiPath)
# -------------------------
function Get-7ZipGuiPath {
    foreach ($path in @(
            (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\7zFM.exe' -ErrorAction SilentlyContinue).'(Default)',
            "$env:ProgramFiles\7-Zip\7zG.exe",
            "$env:ProgramFiles(x86)\7-Zip\7zG.exe"
        )) {
        if ($path) {
            $guiPath = $path -replace '7zFM\.exe$', '7zG.exe' -replace '7z\.exe$', '7zG.exe'
            if (Test-Path $guiPath) {
                Write-LogMessage "Found 7-Zip GUI at: $guiPath"
                return $guiPath
            }
            if ((Split-Path $path -Leaf) -ieq '7zG.exe' -and (Test-Path $path)) {
                Write-LogMessage "Found 7-Zip GUI at: $path"
                return $path
            }
        }
    }
    Write-LogMessage "7-Zip GUI not found."; return $null
}

# -------------------------
# Helper: Choose GUI extractors automatically (Verb-Noun: Get-ArchiveExtractorCandidates)
# -------------------------
function Get-ArchiveExtractorCandidates {
    param([string]$ArchiveType)

    $sevenZipGui = Get-7ZipGuiPath
    $winrar = Get-WinRARPath

    $sevenZipCandidate = if ($sevenZipGui) {
        [PSCustomObject]@{ Name = "7-Zip"; Kind = "7ZipGui"; Path = $sevenZipGui }
    }
    else { $null }

    $winrarCandidate = if ($winrar) {
        [PSCustomObject]@{ Name = "WinRAR"; Kind = "WinRAR"; Path = $winrar }
    }
    else { $null }

    if ($ArchiveType -eq 'rar') {
        return @($winrarCandidate, $sevenZipCandidate) | Where-Object { $null -ne $_ }
    }

    return @($sevenZipCandidate, $winrarCandidate) | Where-Object { $null -ne $_ }
}

# -------------------------
# Helper: Build GUI extractor command line (Verb-Noun: Get-ArchiveExtractorArguments)
# -------------------------
function Get-ArchiveExtractorArguments {
    param(
        [PSCustomObject]$Extractor,
        [string]$ArchivePath,
        [string]$DestinationPath
    )

    switch ($Extractor.Kind) {
        "7ZipGui" {
            return "x -y -aoa -o`"$DestinationPath`" `"$ArchivePath`""
        }
        "WinRAR" {
            $winrarDestination = $DestinationPath
            if (-not $winrarDestination.EndsWith('\')) {
                $winrarDestination = "$winrarDestination\"
            }
            return "x -y -o+ `"$ArchivePath`" `"$winrarDestination`""
        }
        default {
            throw "Unsupported extractor kind: $($Extractor.Kind)"
        }
    }
}

# -------------------------
# Helper: Interpret extractor exit code (Verb-Noun: Test-ArchiveExtractionSucceeded)
# -------------------------
function Test-ArchiveExtractionSucceeded {
    param([int]$ExitCode)
    return ($ExitCode -eq 0 -or $ExitCode -eq 1)
}

# -------------------------
# Helper: Select Extraction Path (Verb-Noun: Select-ExtractionPath)
# Uses modern Windows 10/11 folder picker dialog via COM IFileOpenDialog
# -------------------------
function Select-ExtractionPath {
    param(
        [string]$DefaultPath,
        [switch]$OpenToParent  # If set, opens to parent directory of DefaultPath
    )
    
    # Determine the folder to open the dialog at
    $initialFolder = $DefaultPath
    if ($OpenToParent -and $DefaultPath -and (Test-Path $DefaultPath -ErrorAction SilentlyContinue)) {
        $parentPath = Split-Path -Parent $DefaultPath
        if ($parentPath -and (Test-Path $parentPath -ErrorAction SilentlyContinue)) {
            $initialFolder = $parentPath
        }
    }
    # Also navigate to parent if the specified folder doesn't exist yet
    elseif ($DefaultPath -and -not (Test-Path $DefaultPath -ErrorAction SilentlyContinue)) {
        $parentPath = Split-Path -Parent $DefaultPath
        if ($parentPath -and (Test-Path $parentPath -ErrorAction SilentlyContinue)) {
            $initialFolder = $parentPath
        }
    }
    
    # Check if our FolderPicker type already exists to avoid redefining it
    if (-not ([System.Management.Automation.PSTypeName]'FolderPickerDialog').Type) {
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

public static class FolderPickerDialog {
    public const uint FOS_PICKFOLDERS = 0x20;
    public const uint FOS_FORCEFILESYSTEM = 0x40;
    public const uint SIGDN_FILESYSPATH = 0x80058000;
    
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHCreateItemFromParsingName(
        string pszPath, IntPtr pbc, 
        [MarshalAs(UnmanagedType.LPStruct)] Guid riid, 
        out IShellItem ppv);
    
    private static readonly Guid IShellItemGuid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
    
    public static string ShowDialog(string title, string initialFolder) {
        var dialog = (IFileOpenDialog)new FileOpenDialogRCW();
        dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
        dialog.SetTitle(title);
        
        // Clear dialog's memory of the last folder to ensure we always open at initialFolder
        try { dialog.ClearClientData(); } catch { }
        
        // Set initial folder if path exists
        if (!string.IsNullOrEmpty(initialFolder) && System.IO.Directory.Exists(initialFolder)) {
            IShellItem folder;
            if (SHCreateItemFromParsingName(initialFolder, IntPtr.Zero, IShellItemGuid, out folder) == 0) {
                dialog.SetFolder(folder);
            }
        }
        
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
"@ -ErrorAction Stop
    }

    try {
        $selectedPath = [FolderPickerDialog]::ShowDialog("Select extraction destination folder", $initialFolder)
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
        # For fallback, also use initialFolder (parent) if appropriate
        $folderBrowser.SelectedPath = $initialFolder
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
# Helper: Extract Archive (Verb-Noun: Expand-ArchiveFile)
# Hands extraction to the user's installed extractor GUI, then resumes qBitLauncher.
# -------------------------
function Expand-ArchiveFile {
    param(
        [string]$ArchivePath,
        [string]$DestinationPath
    )
    $ArchiveType = [IO.Path]::GetExtension($ArchivePath).TrimStart('.').ToLowerInvariant()
    Write-LogMessage "Attempting to extract '${ArchivePath}' (Type: ${ArchiveType})"
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

    $extractors = @(Get-ArchiveExtractorCandidates -ArchiveType $ArchiveType)
    if (-not $extractors -or $extractors.Count -eq 0) {
        $errMsg = "No GUI archive extractor found. Please install 7-Zip or WinRAR."
        Write-LogMessage "ERROR: $errMsg"
        Show-ThemedMessageBox -Message $errMsg -Title "Extractor Not Found" -Icon 'Error' | Out-Null
        return $null
    }

    foreach ($extractor in $extractors) {
        $process = $null
        try {
            $arguments = Get-ArchiveExtractorArguments -Extractor $extractor -ArchivePath $ArchivePath -DestinationPath $DestinationPath
            Write-LogMessage "Launching $($extractor.Name) GUI extractor: $($extractor.Path)"
            Write-Host "Launching $($extractor.Name) to extract '${ArchivePath}'..."

            $process = Start-Process -FilePath $extractor.Path -ArgumentList $arguments -PassThru -Wait -ErrorAction Stop
            $exitCode = $process.ExitCode
            Write-LogMessage "$($extractor.Name) exited with code $exitCode."

            if (Test-ArchiveExtractionSucceeded -ExitCode $exitCode) {
                if ($exitCode -eq 1) {
                    Write-LogMessage "$($extractor.Name) completed with warnings. Continuing to executable discovery."
                }
                Write-LogMessage "Extracted with $($extractor.Name) GUI successfully."
                return $DestinationPath
            }

            if ($exitCode -eq 255) {
                Write-LogMessage "Extraction cancelled by user in $($extractor.Name)."
                return $null
            }

            Write-LogMessage "$($extractor.Name) extraction failed with exit code $exitCode. Trying fallback extractor if available."
        }
        catch {
            Write-LogMessage "$($extractor.Name) extraction could not start or complete: $($_.Exception.Message)"
        }
    }

    $errMsg = "Extraction failed. qBitLauncher tried the installed GUI extractors but none completed successfully."
    Write-LogMessage "ERROR: $errMsg"
    Show-ThemedMessageBox -Message $errMsg -Title "Extraction Failed" -Icon 'Error' | Out-Null
    return $null
}

# -------------------------
# Helper: Find ALL Runnables (Verb-Noun: Get-AllRunnables)
# Finds .exe, .bat, and .cmd files for game installers and scripts
# -------------------------
function Get-AllRunnables {
    param([string]$RootFolderPath)
    Write-LogMessage "Searching for runnables (.exe, .bat, .cmd) in '$RootFolderPath' (depth-first sort)."
    Write-Host "Searching for .exe, .bat, .cmd files in '$RootFolderPath' and its subfolders..."
    
    # Get all runnable file types (use Where-Object for reliable filtering)
    $runnableExtensions = @('.exe', '.bat', '.cmd')
    $allRunnables = Get-ChildItem -LiteralPath $RootFolderPath -File -Recurse -ErrorAction SilentlyContinue | 
    Where-Object { $runnableExtensions -contains $_.Extension.ToLowerInvariant() }
    
    if ($allRunnables) {
        # Sort by folder depth (shallower first), then by name
        $sortedRunnables = $allRunnables | Sort-Object @{Expression = { ($_.FullName.Split([IO.Path]::DirectorySeparatorChar).Count) } }, FullName
        Write-LogMessage "Found $($sortedRunnables.Count) runnable files."
        return $sortedRunnables
    }
    Write-LogMessage "No runnable files found in '$RootFolderPath'."
    Write-Warning "No .exe, .bat, or .cmd files found in '$RootFolderPath' or its subdirectories."
    return $null
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
        Confirmed         = $false
        DestinationPath   = $DefaultDestination
        SkippedExtraction = $false
        UseExisting       = $false
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(700, 250)
    $form.StartPosition = 'CenterScreen'
    $form.Add_Load({ Set-FormPositionNearOwner -Form $this })
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Add_Shown({ Set-FormIcon -Form $this })

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
    Set-ThemedButton -Button $browseButton -Colors $colors
    $browseButton.Add_Click({
            $selected = Select-ExtractionPath -DefaultPath $destTextBox.Text -OpenToParent
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
    $infoLabel.ForeColor = $colors.SecondaryText
    $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($infoLabel)

    # Buttons
    $extractButton = New-Object System.Windows.Forms.Button
    $extractButton.Location = New-Object System.Drawing.Point(300, 160)
    $extractButton.Size = New-Object System.Drawing.Size(100, 35)
    $extractButton.Text = "&Extract"
    # Don't set DialogResult - we'll handle validation first
    $extractButton.Add_Click({
            # Validate path before accepting
            $archiveSize = (Get-Item $ArchivePath -ErrorAction SilentlyContinue).Length
            $validation = Test-ExtractionPath -Path $destTextBox.Text -RequiredSpaceBytes ($archiveSize * 3)
            if (-not $validation.Valid) {
                Invoke-ActionSound -Type Error
                Show-ThemedMessageBox -Message $validation.Error -Title "Invalid Path" -Icon 'Warning'
                return
            }
            
            # Check if destination folder already exists and has files
            $destPath = $destTextBox.Text
            if ((Test-Path -LiteralPath $destPath) -and (Get-ChildItem -LiteralPath $destPath -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1)) {
                $overwriteResult = Show-ThemedMessageBox -Message "The destination folder already contains files:`n$destPath`n`nDo you want to overwrite existing files?" -Title "Folder Exists" -Buttons 'YesNoCancel' -Icon 'Question'
                if ($overwriteResult -eq [System.Windows.Forms.DialogResult]::No) {
                    # User chose No - use existing files, skip extraction
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::Retry
                    $form.Close()
                    return
                }
                elseif ($overwriteResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
                    # User cancelled - stay on dialog
                    return
                }
                # User chose Yes - proceed with extraction/overwrite
            }
            
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            $form.Close()
        }.GetNewClosure())

    # Skip Extraction button - for when archive is irrelevant (e.g., ZIP in crack folder)
    $skipToExeButton = New-Object System.Windows.Forms.Button
    $skipToExeButton.Location = New-Object System.Drawing.Point(410, 160)
    $skipToExeButton.Size = New-Object System.Drawing.Size(120, 35)
    $skipToExeButton.Text = "&Skip Extraction"
    $skipToExeButton.DialogResult = [System.Windows.Forms.DialogResult]::Ignore

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(540, 160)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "&Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::No

    foreach ($button in @($extractButton, $skipToExeButton, $cancelButton)) {
        Set-ThemedButton -Button $button -Colors $colors
        $form.Controls.Add($button)
    }
    
    $form.AcceptButton = $extractButton
    $form.CancelButton = $cancelButton

    $owner = $Global:MainForm
    $dialogResult = if ($owner) { $form.ShowDialog($owner) } else { $form.ShowDialog() }
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $result.Confirmed = $true
        $result.DestinationPath = $destTextBox.Text
    }
    elseif ($dialogResult -eq [System.Windows.Forms.DialogResult]::Retry) {
        # User chose to use existing extracted files
        $result.UseExisting = $true
        $result.DestinationPath = $destTextBox.Text
    }
    elseif ($dialogResult -eq [System.Windows.Forms.DialogResult]::Ignore) {
        $result.SkippedExtraction = $true
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
    $form.Size = New-Object System.Drawing.Size(500, 430)
    $form.StartPosition = 'CenterScreen'
    $form.Add_Load({ Set-FormPositionNearOwner -Form $this })
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Add_Shown({ Set-FormIcon -Form $this })

    # Theme label
    $themeLabel = New-Object System.Windows.Forms.Label
    $themeLabel.Location = New-Object System.Drawing.Point(20, 25)
    $themeLabel.Size = New-Object System.Drawing.Size(130, 25)
    $themeLabel.Text = "Theme:"
    $themeLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($themeLabel)

    # Theme dropdown
    $themeCombo = New-Object System.Windows.Forms.ComboBox
    $themeCombo.Location = New-Object System.Drawing.Point(160, 22)
    $themeCombo.Size = New-Object System.Drawing.Size(280, 25)
    $themeCombo.DropDownStyle = 'DropDownList'
    $themeCombo.BackColor = $colors.ControlBack
    $themeCombo.ForeColor = $colors.TextFore
    $themeCombo.FlatStyle = 'Flat'
    [void]$themeCombo.Items.AddRange(@('Dracula', 'Light'))
    $themeCombo.SelectedItem = $Global:UserSettings.Theme
    $form.Controls.Add($themeCombo)

    # Version label
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Location = New-Object System.Drawing.Point(20, 55)
    $versionLabel.Size = New-Object System.Drawing.Size(180, 25)
    $versionLabel.Text = "Version: v$($Global:ScriptVersion)"
    $versionLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($versionLabel)

    # Check for Updates button
    $updateButton = New-Object System.Windows.Forms.Button
    $updateButton.Location = New-Object System.Drawing.Point(270, 50)
    $updateButton.Size = New-Object System.Drawing.Size(170, 30)
    $updateButton.Text = "Check for &Updates"
    Set-ThemedButton -Button $updateButton -Colors $colors
    $updateButton.Add_Click({
            $updateButton.Enabled = $false
            $updateButton.Text = "Checking..."
            [System.Windows.Forms.Application]::DoEvents()
        
            $newVersion = Test-UpdateAvailable
            if ($newVersion) {
                $form.Close()
                Show-UpdatePrompt -NewVersion $newVersion
            }
            else {
                Invoke-ActionSound -Type Success
                Show-ThemedMessageBox -Message "You're running the latest version (v$($Global:ScriptVersion))!" -Title "Up to Date" -Icon 'Information'
                $updateButton.Text = "Check for &Updates"
                $updateButton.Enabled = $true
            }
        })
    $form.Controls.Add($updateButton)

    $qbLabel = New-Object System.Windows.Forms.Label
    $qbLabel.Location = New-Object System.Drawing.Point(20, 98)
    $qbLabel.Size = New-Object System.Drawing.Size(420, 22)
    $qbLabel.Text = "qBittorrent Local API"
    $qbLabel.ForeColor = $colors.TextFore
    $qbLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($qbLabel)

    $webUrlLabel = New-Object System.Windows.Forms.Label
    $webUrlLabel.Location = New-Object System.Drawing.Point(20, 130)
    $webUrlLabel.Size = New-Object System.Drawing.Size(130, 25)
    $webUrlLabel.Text = "API URL:"
    $webUrlLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($webUrlLabel)

    $webUrlTextBox = New-Object System.Windows.Forms.TextBox
    $webUrlTextBox.Location = New-Object System.Drawing.Point(160, 128)
    $webUrlTextBox.Size = New-Object System.Drawing.Size(280, 25)
    $webUrlTextBox.Text = $Global:UserSettings.QbittorrentWebUrl
    $webUrlTextBox.BackColor = $colors.ControlBack
    $webUrlTextBox.ForeColor = $colors.TextFore
    $webUrlTextBox.BorderStyle = 'FixedSingle'
    $form.Controls.Add($webUrlTextBox)

    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Location = New-Object System.Drawing.Point(20, 165)
    $usernameLabel.Size = New-Object System.Drawing.Size(130, 25)
    $usernameLabel.Text = "Username:"
    $usernameLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($usernameLabel)

    $usernameTextBox = New-Object System.Windows.Forms.TextBox
    $usernameTextBox.Location = New-Object System.Drawing.Point(160, 163)
    $usernameTextBox.Size = New-Object System.Drawing.Size(280, 25)
    $usernameTextBox.Text = $Global:UserSettings.QbittorrentUsername
    $usernameTextBox.BackColor = $colors.ControlBack
    $usernameTextBox.ForeColor = $colors.TextFore
    $usernameTextBox.BorderStyle = 'FixedSingle'
    $form.Controls.Add($usernameTextBox)

    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Location = New-Object System.Drawing.Point(20, 200)
    $passwordLabel.Size = New-Object System.Drawing.Size(130, 25)
    $passwordLabel.Text = "Password:"
    $passwordLabel.ForeColor = $colors.TextFore
    $form.Controls.Add($passwordLabel)

    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.Location = New-Object System.Drawing.Point(160, 198)
    $passwordTextBox.Size = New-Object System.Drawing.Size(280, 25)
    $passwordTextBox.Text = $Global:UserSettings.QbittorrentPassword
    $passwordTextBox.UseSystemPasswordChar = $true
    $passwordTextBox.BackColor = $colors.ControlBack
    $passwordTextBox.ForeColor = $colors.TextFore
    $passwordTextBox.BorderStyle = 'FixedSingle'
    $form.Controls.Add($passwordTextBox)

    # Info label
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(20, 238)
    $infoLabel.Size = New-Object System.Drawing.Size(440, 50)
    $infoLabel.Text = "Pass %I from qBittorrent to enable torrent removal.`nThe local API can be auto-enabled when cleanup first needs it."
    $infoLabel.ForeColor = $colors.SecondaryText
    $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($infoLabel)

    # Keyboard shortcuts section
    $shortcutsLabel = New-Object System.Windows.Forms.Label
    $shortcutsLabel.Location = New-Object System.Drawing.Point(20, 295)
    $shortcutsLabel.Size = New-Object System.Drawing.Size(440, 50)
    $shortcutsLabel.Text = "Keyboard Shortcuts (hold Alt key):`nAlt+R: Run   |   Alt+S: Shortcut   |   Alt+O: Open Folder`nAlt+T: Settings   |   Alt+C: Close"
    $shortcutsLabel.ForeColor = $colors.SecondaryText
    $shortcutsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($shortcutsLabel)

    # Buttons
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(260, 345)
    $saveButton.Size = New-Object System.Drawing.Size(100, 35)
    $saveButton.Text = "&Save"
    $saveButton.Add_Click({
            $webUrl = $webUrlTextBox.Text.Trim()
            $parsedUrl = $null
            if ([string]::IsNullOrWhiteSpace($webUrl) -or -not [System.Uri]::TryCreate($webUrl, [System.UriKind]::Absolute, [ref]$parsedUrl)) {
                Invoke-ActionSound -Type Error
                Show-ThemedMessageBox -Message "Enter a valid qBittorrent local API URL, such as http://localhost:8080." -Title "Invalid Settings" -Icon 'Warning'
                return
            }

            $Global:UserSettings.Theme = $themeCombo.SelectedItem
            $Global:UserSettings.QbittorrentWebUrl = $webUrl
            $Global:UserSettings.QbittorrentUsername = $usernameTextBox.Text.Trim()
            $Global:UserSettings.QbittorrentPassword = $passwordTextBox.Text
            $Global:ThemeSelection = $Global:UserSettings.Theme
            $Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]
            if (Save-UserSettings) {
                Invoke-ActionSound -Type Success
                Show-ThemedMessageBox -Message "Settings saved!" -Title "Settings" -Icon 'Information'
            }
            $form.Close()
        })

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(370, 345)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "&Cancel"
    $cancelButton.Add_Click({
            $form.Close()
        })

    foreach ($button in @($saveButton, $cancelButton)) {
        Set-ThemedButton -Button $button -Colors $colors
        $form.Controls.Add($button)
    }
    
    $form.AcceptButton = $saveButton
    $form.CancelButton = $cancelButton

    $owner = $Global:MainForm
    if ($owner) { $form.ShowDialog($owner) | Out-Null } else { $form.ShowDialog() | Out-Null }
    $form.Dispose()
}


# ---------------------------------------------------
# GUI: Main Executable Selection Form (with Icons, Log Panel, Rename)
# ---------------------------------------------------
function Show-ExecutableSelectionForm {
    param(
        [System.Management.Automation.PSObject[]]$FoundExecutables,
        [string]$WindowTitle = "qBitLauncher Action",
        [string]$RootFolder = $null,  # Root folder for Open Folder button
        [string]$TorrentHash = $null,
        [string]$TorrentName = $null
    )
    $colors = $Global:CurrentTheme
    
    # Determine root folder from first executable if not provided
    if (-not $RootFolder -and $FoundExecutables -and $FoundExecutables.Count -gt 0) {
        $RootFolder = $FoundExecutables[0].DirectoryName
    }

    # Expanded form size for split layout
    $form = New-Object System.Windows.Forms.Form
    $Global:MainForm = $form
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(950, 520)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.BackColor = $colors.FormBack
    $font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Font = $font
    $form.Add_Shown({ Set-FormIcon -Form $this })
    $toolTip = New-Object System.Windows.Forms.ToolTip

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(600, 25)
    $label.Text = "Select an executable. Double-click or F2 to rename."
    $label.ForeColor = $colors.TextFore
    $form.Controls.Add($label)

    # === LEFT SIDE: Executable ListView ===
    $imageList = New-Object System.Windows.Forms.ImageList
    $imageList.ImageSize = New-Object System.Drawing.Size(24, 24)
    $imageList.ColorDepth = 'Depth32Bit'

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 40)
    $listView.Size = New-Object System.Drawing.Size(620, 370)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.MultiSelect = $false
    $listView.BackColor = $colors.ControlBack
    $listView.ForeColor = $colors.TextFore
    $listView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $listView.Font = $font
    $listView.SmallImageList = $imageList
    $listView.LabelEdit = $true  # Enable inline editing for rename
    $listView.HeaderStyle = 'None'  # Hide column headers for cleaner look
    
    # Two columns: Name (first, editable) and Folder Path (second, info only)
    $colName = New-Object System.Windows.Forms.ColumnHeader
    $colName.Text = "Name"
    $colName.Width = 180
    $colPath = New-Object System.Windows.Forms.ColumnHeader
    $colPath.Text = "Folder"
    $colPath.Width = 420
    [void]$listView.Columns.Add($colName)
    [void]$listView.Columns.Add($colPath)

    $form.Controls.Add($listView)

    # === RIGHT SIDE: Activity Log Panel ===
    $logLabel = New-Object System.Windows.Forms.Label
    $logLabel.Location = New-Object System.Drawing.Point(645, 40)
    $logLabel.Size = New-Object System.Drawing.Size(280, 20)
    $logLabel.Text = "Activity Log"
    $logLabel.ForeColor = $colors.TextFore
    $logLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($logLabel)

    $logBox = New-Object System.Windows.Forms.TextBox
    $logBox.Location = New-Object System.Drawing.Point(645, 65)
    $logBox.Size = New-Object System.Drawing.Size(280, 345)
    $logBox.Multiline = $true
    $logBox.ReadOnly = $true
    $logBox.ScrollBars = 'Vertical'
    $logBox.BackColor = $colors.ControlBack
    $logBox.ForeColor = $colors.SecondaryText
    $logBox.BorderStyle = 'FixedSingle'
    $logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($logBox)

    # Helper to add log entries
    $addLogEntry = {
        param([string]$Message)
        $timestamp = Get-Date -Format "h:mm tt"
        $entry = "[$timestamp] $Message`r`n"
        $logBox.AppendText($entry)
        Write-LogMessage $Message
    }

    $removeListItem = {
        param([System.Windows.Forms.ListViewItem]$Item)
        if ($Item) {
            $index = $Item.Index
            $listView.Items.Remove($Item)
            if ($listView.Items.Count -gt 0) {
                $nextIndex = [Math]::Min($index, $listView.Items.Count - 1)
                $listView.Items[$nextIndex].Selected = $true
            }
        }
    }

    # Log form open
    & $addLogEntry "Form opened"
    if (-not [string]::IsNullOrWhiteSpace($TorrentHash)) {
        & $addLogEntry "qBittorrent torrent linked: $TorrentHash"
    }

    # Add executables with icons (two columns)
    $iconIndex = 0
    foreach ($exe in $FoundExecutables) {
        if ($exe -and $exe.FullName) {
            $icon = Get-ExecutableIcon -ExePath $exe.FullName
            $folderPath = $exe.DirectoryName
            $fileName = $exe.Name
            
            if ($icon) {
                $imageList.Images.Add($icon)
                $item = New-Object System.Windows.Forms.ListViewItem($fileName, $iconIndex)
                $iconIndex++
            }
            else {
                $item = New-Object System.Windows.Forms.ListViewItem($fileName)
            }
            $item.SubItems.Add($folderPath) | Out-Null
            $item.Tag = $exe.FullName
            [void]$listView.Items.Add($item)
        }
    }
    
    if ($listView.Items.Count -gt 0) {
        $listView.Items[0].Selected = $true
        & $addLogEntry "Found $($listView.Items.Count) executable(s)"
    }

    # === Context Menu for Right-Click ===
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $renameMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $renameMenuItem.Text = "Rename"
    $renameMenuItem.Add_Click({
            if ($listView.SelectedItems.Count -gt 0) {
                $listView.SelectedItems[0].BeginEdit()
            }
        })
    $contextMenu.Items.Add($renameMenuItem) | Out-Null
    $listView.ContextMenuStrip = $contextMenu

    # === Custom Column Resizing with Visible Divider ===
    # Permanently disable horizontal scrollbar using ListView style
    Add-Type -Name ListViewStyleHelper -Namespace Win32 -MemberDefinition '
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        public const int GWL_STYLE = -16;
        public const int LVS_NOHSCROLL = 0x8000;
    ' -ErrorAction SilentlyContinue
    
    # Add LVS_NOHSCROLL style to permanently remove horizontal scrollbar
    $currentStyle = [Win32.ListViewStyleHelper]::GetWindowLong($listView.Handle, [Win32.ListViewStyleHelper]::GWL_STYLE)
    [Win32.ListViewStyleHelper]::SetWindowLong($listView.Handle, [Win32.ListViewStyleHelper]::GWL_STYLE, $currentStyle -bor [Win32.ListViewStyleHelper]::LVS_NOHSCROLL) | Out-Null
    
    # Set column widths - use very conservative sizing
    $usableWidth = $listView.ClientSize.Width - 20  # Large buffer to be safe
    $initialNameWidth = [math]::Floor($usableWidth / 2)
    $listView.Columns[0].Width = $initialNameWidth
    $listView.Columns[1].Width = $usableWidth - $initialNameWidth
    
    # Create visible divider line (simple solid line)
    $dividerPanel = New-Object System.Windows.Forms.Panel
    $dividerPanel.Size = New-Object System.Drawing.Size(2, $listView.Height)
    $dividerPanel.Location = New-Object System.Drawing.Point(($listView.Left + $initialNameWidth), $listView.Top)
    $dividerPanel.BackColor = $colors.Border
    $dividerPanel.Cursor = [System.Windows.Forms.Cursors]::VSplit
    $form.Controls.Add($dividerPanel)
    $dividerPanel.BringToFront()
    
    # State tracking for drag operation
    $resizeState = @{
        Active     = $false
        StartX     = 0
        StartWidth = 0
    }
    
    # Drag limits
    $minWidth = 80
    $maxWidth = $usableWidth - 80
    
    # Update divider position
    $updateDivider = {
        $dividerPanel.Left = $listView.Left + $listView.Columns[0].Width - 1
    }
    
    # Mouse events on the divider panel for dragging
    $dividerPanel.Add_MouseDown({
            param($eventSender, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $resizeState.Active = $true
                $resizeState.StartX = [System.Windows.Forms.Cursor]::Position.X
                $resizeState.StartWidth = $listView.Columns[0].Width
                $dividerPanel.Capture = $true
                $listView.SuspendLayout()  # Prevent flicker
            }
        }.GetNewClosure())
    
    $dividerPanel.Add_MouseMove({
            param($eventSender, $e)
            if ($resizeState.Active) {
                $currentX = [System.Windows.Forms.Cursor]::Position.X
                $delta = $currentX - $resizeState.StartX
                $newWidth = $resizeState.StartWidth + $delta
            
                # Apply limits
                if ($newWidth -lt $minWidth) { $newWidth = $minWidth }
                if ($newWidth -gt $maxWidth) { $newWidth = $maxWidth }
            
                $listView.Columns[0].Width = $newWidth
                $listView.Columns[1].Width = $usableWidth - $newWidth
                & $updateDivider
            }
        }.GetNewClosure())
    
    $dividerPanel.Add_MouseUp({
            param($eventSender, $e)
            if ($resizeState.Active) {
                $resizeState.Active = $false
                $dividerPanel.Capture = $false
                $listView.ResumeLayout()  # Resume layout after drag
            }
        }.GetNewClosure())

    # === Event Handlers ===
    # (No selection logging - it was redundant)

    # Double-click to start rename
    $listView.Add_DoubleClick({
            if ($listView.SelectedItems.Count -gt 0) {
                $listView.SelectedItems[0].BeginEdit()
            }
        })

    # F2 key to start rename
    $listView.Add_KeyDown({
            param($eventSender, $e)
            if ($e.KeyCode -eq 'F2' -and $listView.SelectedItems.Count -gt 0) {
                $listView.SelectedItems[0].BeginEdit()
                $e.Handled = $true
            }
        })

    # After inline edit - perform actual file rename
    $listView.Add_AfterLabelEdit({
            param($eventSender, $e)
        
            if ($null -eq $e.Label -or $e.Label -eq "") {
                $e.CancelEdit = $true
                return
            }
        
            $item = $listView.Items[$e.Item]
            $oldPath = $item.Tag
            $oldName = [System.IO.Path]::GetFileName($oldPath)
            $newName = $e.Label
        
            # Validate new name
            $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            foreach ($char in $invalidChars) {
                if ($newName.Contains($char)) {
                    Show-ThemedMessageBox -Message "Invalid character in filename: $char" -Title "Rename Error" -Icon 'Error'
                    $e.CancelEdit = $true
                    return
                }
            }
        
            # Ensure .exe extension
            if (-not $newName.EndsWith(".exe", [StringComparison]::OrdinalIgnoreCase)) {
                $newName = $newName + ".exe"
            }
        
            # Use case-SENSITIVE comparison to allow case-only renames (e.g., "dogpile" -> "Dogpile")
            if ($newName -ceq $oldName) {
                $e.CancelEdit = $true
                return
            }
        
            $directory = [System.IO.Path]::GetDirectoryName($oldPath)
            $newPath = Join-Path $directory $newName
            
            # Check if this is a case-only rename (same name, different case)
            $isCaseOnlyRename = $newName -ieq $oldName
        
            # Check if target exists (but skip this check for case-only renames on case-insensitive NTFS)
            if (-not $isCaseOnlyRename -and (Test-Path -LiteralPath $newPath)) {
                Show-ThemedMessageBox -Message "A file with that name already exists." -Title "Rename Error" -Icon 'Error'
                $e.CancelEdit = $true
                return
            }
        
            try {
                if ($isCaseOnlyRename) {
                    # For case-only renames, use a two-step process via temp name
                    $tempName = [System.IO.Path]::GetRandomFileName() + ".exe"
                    Rename-Item -LiteralPath $oldPath -NewName $tempName -ErrorAction Stop
                    $tempPath = Join-Path $directory $tempName
                    Rename-Item -LiteralPath $tempPath -NewName $newName -ErrorAction Stop
                }
                else {
                    Rename-Item -LiteralPath $oldPath -NewName $newName -ErrorAction Stop
                }
                $item.Tag = $newPath
                # Note: First column (Name) is automatically updated by LabelEdit
                Invoke-ActionSound -Type Success
                & $addLogEntry "Renamed: $oldName -> $newName"
            }
            catch {
                Show-ThemedMessageBox -Message "Failed to rename: $($_.Exception.Message)" -Title "Rename Error" -Icon 'Error'
                $e.CancelEdit = $true
                & $addLogEntry "Rename failed: $oldName"
            }
        })

    # Helper to get selected executable
    $getSelectedExe = {
        if ($listView.SelectedItems.Count -eq 0) {
            Show-ThemedMessageBox -Message "Please select an executable first." -Title "No Selection" -Icon 'Warning' | Out-Null
            return $null
        }
        $item = $listView.SelectedItems[0]
        $path = $item.Tag
        try {
            return Get-Item -LiteralPath $path -ErrorAction Stop
        }
        catch {
            & $addLogEntry "Missing file removed from list: $($item.Text) - $path"
            & $removeListItem $item
            Show-ThemedMessageBox -Message "File not found:`n$path`n`nIt may have been moved or deleted." -Title "File Missing" -Icon 'Error' | Out-Null
            return $null
        }
    }

    $testUserCancelledLaunch = {
        param([System.Exception]$Exception)

        $currentException = $Exception
        while ($currentException) {
            if (($currentException -is [System.ComponentModel.Win32Exception]) -and $currentException.NativeErrorCode -eq 1223) {
                return $true
            }
            if ($currentException.Message -match 'operation was canceled by the user') {
                return $true
            }
            $currentException = $currentException.InnerException
        }

        return $false
    }

    $testExecutableRequiresAdmin = {
        param([System.IO.FileInfo]$File)

        if ($File.Extension -ine '.exe') {
            return $false
        }

        try {
            $bytes = [System.IO.File]::ReadAllBytes($File.FullName)
            $fileText = [System.Text.Encoding]::ASCII.GetString($bytes)
            return ($fileText -match 'requestedExecutionLevel[^>]+requireAdministrator')
        }
        catch {
            return $false
        }
    }

    $getInternetZoneId = {
        param([System.IO.FileInfo]$File)

        try {
            $zoneStream = Get-Content -LiteralPath $File.FullName -Stream Zone.Identifier -ErrorAction Stop
            $zoneText = $zoneStream -join "`n"
            if ($zoneText -match 'ZoneId=(\d+)') {
                return [int]$Matches[1]
            }
        }
        catch {
            return $null
        }

        return $null
    }

    $getSignatureStatus = {
        param([System.IO.FileInfo]$File)

        if ($File.Extension -ine '.exe') {
            return "NotChecked"
        }

        try {
            return (Get-AuthenticodeSignature -LiteralPath $File.FullName).Status.ToString()
        }
        catch {
            return "Unknown"
        }
    }

    $getLaunchDiagnostics = {
        param([System.IO.FileInfo]$File)

        $zoneId = & $getInternetZoneId $File
        $signatureStatus = & $getSignatureStatus $File

        [PSCustomObject]@{
            RequiresAdmin    = (& $testExecutableRequiresAdmin $File)
            ZoneId           = $zoneId
            IsInternetMarked = ($null -ne $zoneId -and $zoneId -ge 3)
            SignatureStatus  = $signatureStatus
            IsUnsigned       = ($signatureStatus -eq "NotSigned")
        }
    }

    $formatLaunchBlockMessage = {
        param(
            [string]$Intro,
            [PSCustomObject]$Diagnostics
        )

        $lines = @($Intro)

        if ($Diagnostics.RequiresAdmin) {
            $lines += "This installer declares that it requires administrator privileges."
        }
        if ($Diagnostics.IsInternetMarked) {
            $lines += "Windows marks this file as downloaded from the Internet (ZoneId=$($Diagnostics.ZoneId))."
        }
        if ($Diagnostics.IsUnsigned) {
            $lines += "The file is not digitally signed."
        }
        elseif ($Diagnostics.SignatureStatus -and $Diagnostics.SignatureStatus -ne "NotChecked" -and $Diagnostics.SignatureStatus -ne "Valid") {
            $lines += "Signature status: $($Diagnostics.SignatureStatus)."
        }
        if ($Diagnostics.RequiresAdmin) {
            $lines += ""
            $lines += "Because it requires admin, running without admin will still trigger Windows elevation/trust checks."
        }

        return ($lines -join "`n")
    }

    $clearInternetMark = {
        param([System.IO.FileInfo]$File)

        Unblock-File -LiteralPath $File.FullName -ErrorAction Stop
    }

    $startRunnable = {
        param(
            [System.IO.FileInfo]$File,
            [bool]$AsAdmin = $true
        )

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.UseShellExecute = $true
        $startInfo.WorkingDirectory = $File.DirectoryName
        if ($AsAdmin) {
            $startInfo.Verb = "runas"
        }

        $extension = $File.Extension.ToLowerInvariant()
        if ($extension -in @('.bat', '.cmd')) {
            $commandProcessor = if ([string]::IsNullOrWhiteSpace($env:ComSpec)) { Join-Path $env:SystemRoot "System32\cmd.exe" } else { $env:ComSpec }
            $startInfo.FileName = $commandProcessor
            $startInfo.Arguments = "/d /c `"`"$($File.FullName)`"`""
        }
        else {
            $startInfo.FileName = $File.FullName
        }

        $process = [System.Diagnostics.Process]::Start($startInfo)
        if (-not $process) {
            throw "Windows did not return a process handle."
        }
    }

    # === Button Row ===
    $buttonY = 425
    
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Location = New-Object System.Drawing.Point(10, $buttonY)
    $runButton.Size = New-Object System.Drawing.Size(80, 35)
    $runButton.Text = "&Run"
    $runButton.Add_Click({
            $exe = & $getSelectedExe
            if ($exe) {
                try {
                    & $startRunnable $exe $true
                    Invoke-ActionSound -Type Success
                    & $addLogEntry "Launched as admin: $($exe.Name)"
                }
                catch {
                    if (& $testUserCancelledLaunch $_.Exception) {
                        Invoke-ActionSound -Type Notify
                        $diagnostics = & $getLaunchDiagnostics $exe
                        & $addLogEntry "Admin launch cancelled: $($exe.Name) - RequiresAdmin=$($diagnostics.RequiresAdmin); ZoneId=$($diagnostics.ZoneId); Signature=$($diagnostics.SignatureStatus)"

                        if ($diagnostics.RequiresAdmin) {
                            $blockMessage = & $formatLaunchBlockMessage "Windows cancelled the administrator launch." $diagnostics

                            if ($diagnostics.IsInternetMarked) {
                                $unblockMessage = "$blockMessage`n`nRemove the Internet block marker and try again as administrator?`nOnly do this if you trust this installer."
                                $unblockResult = Show-ThemedMessageBox -Message $unblockMessage -Title "Launch Blocked by Windows" -Buttons 'YesNo' -Icon 'Warning'
                                if ($unblockResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                                    try {
                                        & $clearInternetMark $exe
                                        & $addLogEntry "Removed Internet block marker: $($exe.Name)"
                                        & $startRunnable $exe $true
                                        Invoke-ActionSound -Type Success
                                        & $addLogEntry "Launched as admin after unblock: $($exe.Name)"
                                    }
                                    catch {
                                        Invoke-ActionSound -Type Error
                                        if (& $testUserCancelledLaunch $_.Exception) {
                                            & $addLogEntry "Launch still cancelled after unblock: $($exe.Name)"
                                            Show-ThemedMessageBox -Message "Windows still cancelled the launch.`n`nThis installer requires administrator approval. Approve the Windows UAC or SmartScreen prompt to run it." -Title "Launch Cancelled" -Icon 'Warning'
                                        }
                                        else {
                                            & $addLogEntry "Launch failed after unblock: $($exe.Name) - $($_.Exception.Message)"
                                            Show-ThemedMessageBox -Message "Failed to launch after unblock: $($_.Exception.Message)" -Title "Error" -Icon 'Error'
                                        }
                                    }
                                }
                                else {
                                    & $addLogEntry "Launch cancelled; Internet block marker kept: $($exe.Name)"
                                }
                            }
                            else {
                                Show-ThemedMessageBox -Message $blockMessage -Title "Launch Blocked by Windows" -Icon 'Warning'
                            }
                            return
                        }

                        $fallbackResult = Show-ThemedMessageBox -Message "Administrator launch was cancelled.`n`nRun without administrator privileges instead?" -Title "Launch Cancelled" -Buttons 'YesNo' -Icon 'Question'
                        if ($fallbackResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                            try {
                                & $startRunnable $exe $false
                                Invoke-ActionSound -Type Success
                                & $addLogEntry "Launched without admin: $($exe.Name)"
                            }
                            catch {
                                Invoke-ActionSound -Type Error
                                & $addLogEntry "Launch failed without admin: $($exe.Name) - $($_.Exception.Message)"
                                if (& $testUserCancelledLaunch $_.Exception) {
                                    $fallbackDiagnostics = & $getLaunchDiagnostics $exe
                                    $fallbackMessage = & $formatLaunchBlockMessage "Windows cancelled the non-admin launch." $fallbackDiagnostics
                                    Show-ThemedMessageBox -Message $fallbackMessage -Title "Launch Blocked by Windows" -Icon 'Warning'
                                }
                                else {
                                    Show-ThemedMessageBox -Message "Failed to launch without admin: $($_.Exception.Message)" -Title "Error" -Icon 'Error'
                                }
                            }
                        }
                        else {
                            & $addLogEntry "Launch cancelled: $($exe.Name)"
                        }
                        return
                    }

                    Invoke-ActionSound -Type Error
                    & $addLogEntry "Launch failed: $($exe.Name) - $($_.Exception.Message)"
                    Show-ThemedMessageBox -Message "Failed to launch: $($_.Exception.Message)" -Title "Error" -Icon 'Error'
                }
            }
        })
    
    $shortcutButton = New-Object System.Windows.Forms.Button
    $shortcutButton.Location = New-Object System.Drawing.Point(95, $buttonY)
    $shortcutButton.Size = New-Object System.Drawing.Size(90, 35)
    $shortcutButton.Text = "&Shortcut"
    $shortcutButton.Add_Click({
            # Check selection first with specific error message
            if ($listView.SelectedItems.Count -eq 0) {
                Invoke-ActionSound -Type Error
                Show-ThemedMessageBox -Message "Please select an executable to create a shortcut for." -Title "No Selection" -Icon 'Warning'
                return
            }
            
            $exePath = $listView.SelectedItems[0].Tag
            try {
                $exe = Get-Item -LiteralPath $exePath -ErrorAction Stop
                $desktopPath = [System.Environment]::GetFolderPath('Desktop')
                $shortcutDisplayName = $exe.BaseName
                $shortcutName = $exe.BaseName + ".lnk"
                $shortcutPath = Join-Path $desktopPath $shortcutName
                $wshell = New-Object -ComObject WScript.Shell
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $exe.FullName
                $shortcut.WorkingDirectory = $exe.DirectoryName
                $shortcut.Save()
                Invoke-ActionSound -Type Success
                & $addLogEntry "Shortcut: $shortcutDisplayName"
                Show-ThemedMessageBox -Message "Shortcut created on Desktop:`n$shortcutDisplayName" -Title "Success" -Icon 'Information'
            }
            catch {
                Invoke-ActionSound -Type Error
                & $addLogEntry "Shortcut failed: $($_.Exception.Message)"
                Show-ThemedMessageBox -Message "Failed to create shortcut: $($_.Exception.Message)" -Title "Error" -Icon 'Error'
            }
        })
    
    $exploreButton = New-Object System.Windows.Forms.Button
    $exploreButton.Location = New-Object System.Drawing.Point(190, $buttonY)
    $exploreButton.Size = New-Object System.Drawing.Size(100, 35)
    $exploreButton.Text = "&Open Folder"
    $exploreButton.Add_Click({
            # Open selected item's folder if selected, otherwise open root folder
            if ($listView.SelectedItems.Count -gt 0) {
                $exePath = $listView.SelectedItems[0].Tag
                $folderToOpen = [System.IO.Path]::GetDirectoryName($exePath)
            }
            else {
                $folderToOpen = $RootFolder
            }
            
            # Use -LiteralPath to handle paths with special characters like brackets []
            if ($folderToOpen -and (Test-Path -LiteralPath $folderToOpen)) {
                Start-Process explorer -ArgumentList "`"$folderToOpen`""
                & $addLogEntry "Opened: $folderToOpen"
            }
            else {
                Show-ThemedMessageBox -Message "Folder not found." -Title "Error" -Icon 'Error'
            }
        })

    $renameButton = New-Object System.Windows.Forms.Button
    $renameButton.Location = New-Object System.Drawing.Point(295, $buttonY)
    $renameButton.Size = New-Object System.Drawing.Size(90, 35)
    $renameButton.Text = "Re&name"
    $renameButton.Add_Click({
            if ($listView.SelectedItems.Count -gt 0) {
                $listView.SelectedItems[0].BeginEdit()
            }
        })

    $settingsButton = New-Object System.Windows.Forms.Button
    $settingsButton.Location = New-Object System.Drawing.Point(390, $buttonY)
    $settingsButton.Size = New-Object System.Drawing.Size(90, 35)
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
                $logLabel.ForeColor = $newColors.TextFore
                $logBox.BackColor = $newColors.ControlBack
                $logBox.ForeColor = $newColors.SecondaryText
            
                foreach ($ctrl in $form.Controls) {
                    if ($ctrl -is [System.Windows.Forms.Button]) {
                        $ctrl.BackColor = $newColors.ButtonBack
                        $ctrl.ForeColor = $newColors.TextFore
                        $ctrl.FlatAppearance.BorderColor = $newColors.Accent
                    }
                }
                Set-DestructiveButton -Button $cleanupButton -Colors $newColors
                $form.Refresh()
                & $addLogEntry "Theme changed"
            }
        })

    # === Cleanup Button (unified: torrent removal + source deletion) ===
    $cleanupButton = New-Object System.Windows.Forms.Button
    $cleanupButton.Location = New-Object System.Drawing.Point(650, $buttonY)
    $cleanupButton.Size = New-Object System.Drawing.Size(175, 35)
    $cleanupButton.Text = "Cleanup"
    $cleanupState = @{
        Busy           = $false
        TorrentRemoved = $false
        SourceDeleted  = $false
    }
    Set-DestructiveButton -Button $cleanupButton -Colors $colors
    $toolTip.SetToolTip($cleanupButton, "Remove torrent or delete source files.")

    # Pre-calculate source deletion eligibility
    $sourcePath = $filePathFromQB
    $sourceIsValid = (-not [string]::IsNullOrWhiteSpace($sourcePath)) -and (Test-Path -LiteralPath $sourcePath)
    $sourceIsSameAsRoot = $sourceIsValid -and $RootFolder -and ($sourcePath.TrimEnd('\') -eq $RootFolder.TrimEnd('\'))
    $sourceCanDelete = $sourceIsValid -and -not $sourceIsSameAsRoot

    $cleanupButton.Add_Click({
            if ($cleanupState.Busy) {
                return
            }

            # Only build torrent label when there's actually a torrent hash
            $torrentLabel = $null
            if (-not [string]::IsNullOrWhiteSpace($TorrentHash)) {
                $torrentLabel = if (-not [string]::IsNullOrWhiteSpace($TorrentName)) {
                    $TorrentName
                }
                elseif (-not [string]::IsNullOrWhiteSpace($RootFolder)) {
                    [System.IO.Path]::GetFileName($RootFolder.TrimEnd('\'))
                }
                else {
                    "Current torrent"
                }
            }

            # Check if source can still be deleted (it may have been deleted already)
            $currentSourceCanDelete = $sourceCanDelete -and -not $cleanupState.SourceDeleted -and (Test-Path -LiteralPath $sourcePath)

            $choice = Show-CleanupConfirmForm -TorrentLabel $torrentLabel -TorrentHash $TorrentHash -FilePath $filePathFromQB -SourcePath $sourcePath -SourceCanDelete $currentSourceCanDelete
            if ($choice -eq "Cancel") {
                & $addLogEntry "Cleanup cancelled"
                return
            }

            # --- Handle Delete Source (no torrent check needed) ---
            if ($choice -eq "DeleteSource") {
                if ($cleanupState.SourceDeleted) {
                    Show-ThemedMessageBox -Message "The source has already been deleted." -Title "Cleanup" -Icon 'Information'
                    return
                }
                if (-not (Test-Path -LiteralPath $sourcePath)) {
                    $cleanupState.SourceDeleted = $true
                    Show-ThemedMessageBox -Message "Source no longer exists:`n$sourcePath" -Title "Cleanup" -Icon 'Information'
                    return
                }

                $typeLabel = if (Test-Path -LiteralPath $sourcePath -PathType Container) { "folder" } else { "file" }
                $confirmResult = Show-ThemedMessageBox -Message "Delete the source $typeLabel`?`n`n$sourcePath`n`nThis cannot be undone." -Title "Delete Source" -Buttons 'YesNo' -Icon 'Warning'
                if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                    & $addLogEntry "Delete source cancelled"
                    return
                }

                try {
                    if (Test-Path -LiteralPath $sourcePath -PathType Container) {
                        Remove-Item -LiteralPath $sourcePath -Recurse -Force -ErrorAction Stop
                    }
                    else {
                        Remove-Item -LiteralPath $sourcePath -Force -ErrorAction Stop
                    }
                    Invoke-ActionSound -Type Success
                    $sourceName = [System.IO.Path]::GetFileName($sourcePath.TrimEnd('\'))
                    & $addLogEntry "Deleted source: $sourcePath"
                    $cleanupState.SourceDeleted = $true
                    Show-ThemedMessageBox -Message "Source deleted:`n$sourceName" -Title "Cleanup" -Icon 'Information'
                }
                catch {
                    Invoke-ActionSound -Type Error
                    & $addLogEntry "Delete source failed: $($_.Exception.Message)"
                    Show-ThemedMessageBox -Message "Could not delete source:`n$($_.Exception.Message)" -Title "Cleanup" -Icon 'Error'
                }
                return
            }

            # --- Handle Torrent options (check hash first) ---
            if ($cleanupState.TorrentRemoved) {
                Show-ThemedMessageBox -Message "This torrent was already removed from qBittorrent." -Title "Cleanup" -Icon 'Information'
                return
            }
            if ([string]::IsNullOrWhiteSpace($TorrentHash)) {
                $commandExample = 'powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\qBitLauncher.ps1" "%F" "%I" "%N"'
                Show-ThemedMessageBox -Message "No torrent hash available. This file may not have been launched from qBittorrent (e.g. opened via right-click context menu).`n`nTo enable torrent removal, launch from qBittorrent with this command:`n$commandExample" -Title "Cleanup" -Icon 'Information'
                return
            }

            # Seeding confirmation before torrent removal
            $seedConfirm = Show-ThemedMessageBox -Message "Please consider seeding!`n`nAre you sure you want to remove this torrent from qBittorrent?" -Title "Cleanup" -Buttons 'YesNo' -Icon 'Question'
            if ($seedConfirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                & $addLogEntry "Torrent removal cancelled (seeding prompt)"
                return
            }

            $deleteFiles = ($choice -eq "DeleteTorrentAndFiles")
            $oldText = $cleanupButton.Text
            $cleanupState.Busy = $true
            $cleanupButton.Text = "Removing..."
            [System.Windows.Forms.Application]::DoEvents()

            try {
                Remove-QBittorrentTorrent -TorrentHash $TorrentHash -DeleteFiles $deleteFiles
                Invoke-ActionSound -Type Success
                $mode = if ($deleteFiles) { "with files" } else { "torrent only" }
                & $addLogEntry "Removed from qBittorrent ($mode): $torrentLabel"
                $cleanupState.TorrentRemoved = $true
                $cleanupButton.Text = $oldText
                Show-ThemedMessageBox -Message "Removed from qBittorrent:`n$torrentLabel" -Title "Cleanup" -Icon 'Information'
            }
            catch {
                Invoke-ActionSound -Type Error
                $errorMessage = $_.Exception.Message
                & $addLogEntry "qBittorrent removal failed: $errorMessage"
                $cleanupButton.Text = $oldText

                $repairResult = Invoke-QBittorrentLocalApiRepairFlow -ErrorMessage $errorMessage -RetryAction {
                    Remove-QBittorrentTorrent -TorrentHash $TorrentHash -DeleteFiles $deleteFiles
                }

                switch ($repairResult.Status) {
                    "RetrySucceeded" {
                        Invoke-ActionSound -Type Success
                        $mode = if ($deleteFiles) { "with files" } else { "torrent only" }
                        & $addLogEntry "Enabled qBittorrent local API and removed ($mode): $torrentLabel"
                        $cleanupState.TorrentRemoved = $true
                        Show-ThemedMessageBox -Message "qBittorrent cleanup is set up now, and this torrent was removed:`n$torrentLabel" -Title "Cleanup" -Icon 'Information'
                    }
                    "PendingExit" {
                        & $addLogEntry "qBittorrent must exit before local API setup can continue"
                    }
                    "PendingRestart" {
                        & $addLogEntry "qBittorrent local API enabled; waiting for qBittorrent to reopen"
                    }
                    "Declined" {
                        & $addLogEntry "qBittorrent local API setup cancelled"
                    }
                    default {
                        $message = $repairResult.Message
                        if ([string]::IsNullOrWhiteSpace($message)) {
                            $message = $errorMessage
                        }
                        Show-ThemedMessageBox -Message "Could not remove the torrent:`n$message" -Title "Cleanup" -Icon 'Error'
                    }
                }
            }
            finally {
                $cleanupState.Busy = $false
            }
        }.GetNewClosure())
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(835, $buttonY)
    $closeButton.Size = New-Object System.Drawing.Size(90, 35)
    $closeButton.Text = "&Close"
    $closeButton.Add_Click({
            $form.Close()
        })

    foreach ($button in @($runButton, $shortcutButton, $exploreButton, $renameButton, $settingsButton, $closeButton)) {
        Set-ThemedButton -Button $button -Colors $colors
        $form.Controls.Add($button)
    }
    $form.Controls.Add($cleanupButton)

    $form.ActiveControl = $listView

    # Show form (blocks until closed)
    $form.ShowDialog() | Out-Null
    
    # Cleanup
    $Global:MainForm = $null
    $toolTip.Dispose()
    $imageList.Dispose()
    $form.Dispose()
}

# ===================================================================
# MAIN SCRIPT LOGIC STARTS HERE
# ===================================================================

# Check for updates on startup (non-blocking)
Write-LogMessage "Checking for updates..."
$startupNewVersion = Test-UpdateAvailable
if ($startupNewVersion) {
    Write-LogMessage "Update available: v$startupNewVersion"
    Show-UpdatePrompt -NewVersion $startupNewVersion
}

# Test mode: If no path provided (e.g., double-clicking EXE), show folder picker for testing
if ([string]::IsNullOrWhiteSpace($filePathFromQB)) {
    Write-LogMessage "No path provided - entering test mode"
    
    # Show folder picker for testing
    $testPath = Select-ExtractionPath -DefaultPath ([Environment]::GetFolderPath('Desktop'))
    
    if ($testPath) {
        $filePathFromQB = $testPath
        Write-LogMessage "Test mode: User selected folder '$testPath'"
    }
    else {
        Write-LogMessage "Test mode: User cancelled folder selection. Exiting."
        exit 0
    }
}

if (-not (Test-Path -LiteralPath $filePathFromQB)) {
    $errMsg = "Error: Initial path not found - $filePathFromQB"
    Write-LogMessage "FATAL: $errMsg. Script exiting."
    Show-ThemedMessageBox -Message "Path not found:`n$filePathFromQB" -Title "qBitLauncher Error" -Icon 'Error'
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
        Write-LogMessage "No archives found in folder. Searching for runnables..."
        $allRunnables = Get-AllRunnables -RootFolderPath $downloadFolder
        if ($allRunnables) {
            $mainFileToProcess = $allRunnables
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
        
        # Default extraction path (beside the archive, in a subfolder named after archive)
        $defaultExtractPath = Join-Path $parentDir $baseName
        
        # Show extraction confirmation with custom path option
        $extractionResult = Show-ExtractionConfirmForm -ArchivePath $filePath -DefaultDestination $defaultExtractPath
        
        if ($extractionResult.Confirmed) {
            Write-LogMessage "User confirmed extraction to: $($extractionResult.DestinationPath)"
            $extractedDir = Expand-ArchiveFile -ArchivePath $filePath -DestinationPath $extractionResult.DestinationPath
            if ($extractedDir) {
                Write-Host "`nExtraction complete. Searching for runnables..."
                $runnablesInArchive = Get-AllRunnables -RootFolderPath $extractedDir
                if ($runnablesInArchive) {
                    Show-ExecutableSelectionForm -FoundExecutables $runnablesInArchive -WindowTitle "qBitLauncher" -RootFolder $extractedDir -TorrentHash $torrentHashFromQB -TorrentName $torrentNameFromQB
                }
                else {
                    Write-Warning "No executables found in the extracted folder: $extractedDir"
                    Start-Process explorer -ArgumentList "`"$extractedDir`""
                }
            }
        }
        elseif ($extractionResult.UseExisting) {
            # User chose to use existing extracted files instead of re-extracting
            $existingDir = $extractionResult.DestinationPath
            Write-LogMessage "User chose to use existing files in: '$existingDir'"
            Write-Host "`nUsing existing extracted files. Searching for runnables..."
            $runnablesInExisting = Get-AllRunnables -RootFolderPath $existingDir
            if ($runnablesInExisting) {
                Show-ExecutableSelectionForm -FoundExecutables $runnablesInExisting -WindowTitle "qBitLauncher" -RootFolder $existingDir -TorrentHash $torrentHashFromQB -TorrentName $torrentNameFromQB
            }
            else {
                Write-Warning "No executables found in existing folder: $existingDir"
                Start-Process explorer -ArgumentList "`"$existingDir`""
            }
        }
        elseif ($extractionResult.SkippedExtraction) {
            # User chose to skip extraction and go directly to EXE handler
            # Use the original download folder (root) instead of the archive's parent directory
            $searchFolder = if ($downloadFolder) { $downloadFolder } else { $parentDir }
            Write-LogMessage "User skipped extraction. Searching for executables in: '$searchFolder'"
            $runnablesInFolder = Get-AllRunnables -RootFolderPath $searchFolder
            if ($runnablesInFolder) {
                Show-ExecutableSelectionForm -FoundExecutables $runnablesInFolder -WindowTitle "qBitLauncher" -RootFolder $searchFolder -TorrentHash $torrentHashFromQB -TorrentName $torrentNameFromQB
            }
            else {
                Write-Warning "No executables found in folder: $searchFolder"
                Start-Process explorer -ArgumentList "`"$searchFolder`""
            }
        }
        else {
            Write-LogMessage "User declined extraction. Opening folder: '$parentDir'"
            Start-Process explorer -ArgumentList "`"$parentDir`""
        }
    } 
    elseif ($ext -in @('exe', 'bat', 'cmd') -or $mainFileToProcess -is [array]) {
        $runnables = if ($mainFileToProcess -is [array]) { $mainFileToProcess } else { @($mainFileToProcess) }
        Write-LogMessage "Processing one or more runnable files."
        
        Show-ExecutableSelectionForm -FoundExecutables $runnables -WindowTitle "qBitLauncher" -RootFolder $parentDir -TorrentHash $torrentHashFromQB -TorrentName $torrentNameFromQB
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
