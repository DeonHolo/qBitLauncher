# qBitLauncher.ps1

param(
    [string]$filePathFromQB,
    [string]$torrentHashFromQB,
    [string]$torrentNameFromQB
)

# <BUNDLED_FUNCTIONS>

# -------------------------
# GLOBAL INITIALIZATION
# -------------------------
# Determine true execution context (Fix for PS2EXE $PSCommandPath unreliability on new PCs)
$processPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$Global:IsCompiledExe = ($processPath -notmatch 'powershell(_ise)?\.exe$' -and $processPath -notmatch 'pwsh\.exe$')

# Determine script directory (handle PS2EXE compiled EXE where $PSScriptRoot is empty)
$Global:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent }
elseif ([System.AppDomain]::CurrentDomain.BaseDirectory) { [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\') }
else { [Environment]::CurrentDirectory }

$Global:ExecutablePath = if ($Global:IsCompiledExe) { $processPath } else { (Join-Path $Global:ScriptDir "qBitLauncher.ps1") }

# Load .NET assemblies at the start to make their types available globally.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not $Global:IsCompiledExe) {
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
    try {
        [Win32.Shell32]::SetCurrentProcessExplicitAppUserModelID("qBitLauncher.App")
    } catch {}

    # Win32 API for setting form icons properly to avoid taskbar glitches
    Add-Type -Name User32Icon -Namespace Win32 -MemberDefinition '
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
    public const int WM_SETICON = 128;
    public const int ICON_SMALL = 0;
    public const int ICON_BIG = 1;
    '
}

# -------------------------
# Configuration
# -------------------------
$ProgressPreference = 'SilentlyContinue'

# Version and update settings
$Global:ScriptVersion = "3.4.1"
$Global:MainForm = $null
$Global:GitHubRawUrl = "https://raw.githubusercontent.com/DeonHolo/qBitLauncher/main/qBitLauncher.ps1"
$Global:GitHubCommitsUrl = "https://github.com/DeonHolo/qBitLauncher/commits/main"

# Log file in script folder for easy access
$LogFile = Join-Path (Join-Path (Get-AppDataDirectory) "logs") "qBitLauncher.log"
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
    Write-LogMessage "Could not load embedded app icon: $($_.Exception.Message)"
}

# Helper function to set form icon properly (including taskbar)
function Set-FormIcon {
    param([System.Windows.Forms.Form]$Form)
    Set-AppFormAppearance $Form
    if ($Global:AppIcon) {
        $Form.Icon = $Global:AppIcon
        # Also set via SendMessage for proper taskbar display
        if ("Win32.User32Icon" -as [type]) {
            $handle = $Form.Handle
            [Win32.User32Icon]::SendMessage($handle, [Win32.User32Icon]::WM_SETICON, [IntPtr][Win32.User32Icon]::ICON_BIG, $Global:AppIcon.Handle) | Out-Null
            [Win32.User32Icon]::SendMessage($handle, [Win32.User32Icon]::WM_SETICON, [IntPtr][Win32.User32Icon]::ICON_SMALL, $Global:AppIcon.Handle) | Out-Null
        }
    }
}

# Helper function to apply consistent theme styling to buttons
function Set-ThemedButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [hashtable]$Colors = $Global:CurrentTheme
    )
    $Button.UseVisualStyleBackColor = $false
    $Button.BackColor = $Colors.ButtonBack
    $Button.ForeColor = $Colors.TextFore
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $Colors.Accent
    $Button.FlatAppearance.MouseOverBackColor = $Colors.Border
    $Button.FlatAppearance.MouseDownBackColor = $Colors.ButtonBack
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



function Show-UpdatePrompt {
    param([string]$NewVersion)

    $colors = $Global:CurrentTheme

    $form = New-Object System.Windows.Forms.Form
    $form.AutoScaleMode = "Dpi"
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
        # The verified replacement helper preserves this workflow's arguments.
        Update-Script -Restart
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
        ControlBack   = [System.Drawing.Color]::FromArgb(32, 32, 32)
        ButtonBack    = [System.Drawing.Color]::FromArgb(47, 52, 64)
        Border        = [System.Drawing.Color]::FromArgb(68, 71, 90)
        Accent        = [System.Drawing.Color]::FromArgb(98, 114, 164)
        SecondaryText = [System.Drawing.Color]::FromArgb(150, 150, 150)
    }
}
$Global:CurrentTheme = $Global:Themes[$Global:ThemeSelection]

# -------------------------
# Helper: Logging (Verb-Noun: Write-LogMessage)
# -------------------------

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
$Global:UserSettings = New-DefaultUserSettings



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
    $form.AutoScaleMode = "Dpi"
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
        deleteFiles = $DeleteFiles.ToString().ToLowerInvariant()
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
    $form.AutoScaleMode = "Dpi"
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
    $enableResult = Show-ThemedMessageBox -Message "qBittorrent's local API is unavailable.`n`nApply these changes to qBittorrent.ini?`nWebUI\Enabled=true`nWebUI\Port=$port`nWebUI\Address=127.0.0.1`n`nExisting username, password, HTTPS, and authentication settings will be preserved. A backup is kept before editing.`n`nApply and retry removal?" -Title "Set Up qBittorrent Cleanup" -Buttons 'YesNo' -Icon 'Question'
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
        if (Test-Path -LiteralPath $Path) {
            # Directory exists, try to create a temp file
            $testFile = Join-Path $Path ".qbit_write_test_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
            [System.IO.File]::WriteAllText($testFile, "test")
            Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
        }
        else {
            # Directory doesn't exist, check if we can create it
            $parent = Split-Path $Path -Parent
            if ($parent -and (Test-Path -LiteralPath $parent)) {
                $testFile = Join-Path $parent ".qbit_write_test_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
                [System.IO.File]::WriteAllText($testFile, "test")
                Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
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
        if ($parentPath -and (Test-Path -LiteralPath $parentPath -ErrorAction SilentlyContinue)) {
            $initialFolder = $parentPath
        }
    }
    # Also navigate to parent if the specified folder doesn't exist yet
    elseif ($DefaultPath -and -not (Test-Path $DefaultPath -ErrorAction SilentlyContinue)) {
        $parentPath = Split-Path -Parent $DefaultPath
        if ($parentPath -and (Test-Path -LiteralPath $parentPath -ErrorAction SilentlyContinue)) {
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
            Write-LogMessage "Launching $($extractor.Name) to extract '${ArchivePath}'..."

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
    $form.AutoScaleMode = "Dpi"
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
            $archiveSize = (Get-Item -LiteralPath $ArchivePath -ErrorAction SilentlyContinue).Length
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


# Use the Unicode Shell Link interface directly. WScript.Shell rejects shortcut
# paths containing characters outside the active Windows ANSI code page.

# ---------------------------------------------------
# GUI: Main Executable Selection Form (with Icons, Log Panel, Rename)
# ---------------------------------------------------

# -------------------------
# First-Time Setup GUI and Integration
# -------------------------

function Install-QBittorrentIntegration {
    # Check if qbittorrent is running
    $qbProcess = Get-Process -Name "qbittorrent" -ErrorAction SilentlyContinue
    if ($qbProcess) {
        $msg = "qBittorrent is currently running.`n`nPlease close qBittorrent completely (File -> Exit) before installing the integration, otherwise it may overwrite the changes when it closes."
        Show-ThemedMessageBox -Message $msg -Title "Close qBittorrent" -Icon 'Warning'
        return $false
    }

    $iniPath = Join-Path $env:APPDATA "qBittorrent\qBittorrent.ini"
    if (-not (Test-Path $iniPath)) {
        Show-ThemedMessageBox -Message "Could not find qBittorrent.ini at:`n$iniPath`n`nPlease ensure qBittorrent is installed and has been run at least once." -Title "Error" -Icon 'Error'
        return $false
    }

    try {
        $command = if ($Global:IsCompiledExe) {
            "`"$Global:ExecutablePath`" `"%F`" `"%I`" `"%N`""
        } else {
            "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$Global:ExecutablePath`" `"%F`" `"%I`" `"%N`""
        }

        # Qt's QSettings INI parser treats backslashes and quotes as escape characters.
        # We must escape them so the literal path is preserved in qBittorrent.
        $command = $command.Replace('\', '\\').Replace('"', '\"')

        # Parse and update the ini file
        $iniLines = [System.IO.File]::ReadAllLines($iniPath)

        # Check current state first
        $isAlreadyInstalled = $false
        $inAutoRunCheck = $false
        foreach ($line in $iniLines) {
            if ($line -match '^\[(.*)\]$') {
                $inAutoRunCheck = ($line -eq '[AutoRun]')
            } elseif ($inAutoRunCheck) {
                if ($line -match '^program\s*=(.*)$') {
                    if ($Matches[1] -match 'qBitLauncher') {
                        $isAlreadyInstalled = $true
                        break
                    }
                }
            }
        }

        if ($isAlreadyInstalled) {
            $confirm = Show-ThemedMessageBox -Message "qBitLauncher is already integrated with qBittorrent.`n`nDo you want to reinstall or update the integration?" -Title "Already Installed" -Icon 'Information' -Buttons 'YesNo'
            if ($confirm -ne 'Yes') {
                return $true
            }
        }

        $newLines = @()
        $inAutoRun = $false
        $foundAutoRun = $false
        $updatedEnabled = $false
        $updatedProgram = $false

        foreach ($line in $iniLines) {
            if ($line -match '^\[(.*)\]$') {
                if ($inAutoRun) {
                    if (-not $updatedEnabled) { $newLines += "enabled=true" }
                    if (-not $updatedProgram) { $newLines += "program=$command" }
                    $inAutoRun = $false
                }

                if ($line -eq '[AutoRun]') {
                    $inAutoRun = $true
                    $foundAutoRun = $true
                }
                $newLines += $line
            }
            elseif ($inAutoRun) {
                if ($line -match '^enabled\s*=') {
                    $newLines += "enabled=true"
                    $updatedEnabled = $true
                }
                elseif ($line -match '^program\s*=') {
                    $newLines += "program=$command"
                    $updatedProgram = $true
                }
                else {
                    $newLines += $line
                }
            }
            else {
                $newLines += $line
            }
        }

        if ($inAutoRun) {
            if (-not $updatedEnabled) { $newLines += "enabled=true" }
            if (-not $updatedProgram) { $newLines += "program=$command" }
        }
        elseif (-not $foundAutoRun) {
            $newLines += ""
            $newLines += "[AutoRun]"
            $newLines += "enabled=true"
            $newLines += "program=$command"
        }

        $integrationBackup = $iniPath + ".qBitLauncher-integration-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".bak"
        [IO.File]::Copy($iniPath, $integrationBackup, $true)
        [System.IO.File]::WriteAllLines($iniPath, $newLines)

        Show-ThemedMessageBox -Message "Successfully integrated qBitLauncher into qBittorrent!`n`nqBittorrent will now automatically launch this tool when a download completes." -Title "Success" -Icon 'Information'
        return $true
    }
    catch {
        Show-ThemedMessageBox -Message "Failed to update qBittorrent configuration:`n$($_.Exception.Message)" -Title "Error" -Icon 'Error'
        return $false
    }
}

function Install-ContextMenu {
    $commandStr = if ($Global:IsCompiledExe) {
        "`"$Global:ExecutablePath`" `"%V`""
    } else {
        "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Global:ExecutablePath`" `"%V`""
    }

    $fileCommandStr = $commandStr.Replace("%V", "%1")

    try {
        # Using HKCU:\Software\Classes instead of HKCR avoids the need for Administrator privileges
        $basePaths = @(
            "HKCU:\Software\Classes\Directory\shell\qBitLauncher",
            "HKCU:\Software\Classes\Directory\Background\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.zip\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.rar\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.7z\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.iso\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.img\shell\qBitLauncher"
        )

        foreach ($path in $basePaths) {
            if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
            New-ItemProperty -Path $path -Name "(default)" -Value "Open with qBitLauncher" -Force | Out-Null

            $icon = if ($Global:IsCompiledExe) { "$Global:ExecutablePath,0" } else { "powershell.exe,0" }
            New-ItemProperty -Path $path -Name "Icon" -Value $icon -Force | Out-Null
            New-ItemProperty -Path $path -Name "Extended" -Value "" -Force | Out-Null

            $cmdPath = "$path\command"
            if (-not (Test-Path $cmdPath)) { New-Item -Path $cmdPath -Force | Out-Null }

            $finalCmd = if ($path -match "Directory") { $commandStr } else { $fileCommandStr }
            New-ItemProperty -Path $cmdPath -Name "(default)" -Value $finalCmd -Force | Out-Null
        }

        Show-ThemedMessageBox -Message "Successfully added 'Open with qBitLauncher' to Shift+Right-click menus!" -Title "Success" -Icon 'Information'
        return $true
    }
    catch {
        Show-ThemedMessageBox -Message "Failed to install context menus:`n$($_.Exception.Message)" -Title "Error" -Icon 'Error'
        return $false
    }
}

function Remove-QBittorrentIntegration {
    # Check if qbittorrent is running
    $qbProcess = Get-Process -Name "qbittorrent" -ErrorAction SilentlyContinue
    if ($qbProcess) {
        $msg = "qBittorrent is currently running.`n`nPlease close qBittorrent completely (File -> Exit) before uninstalling the integration, otherwise it may overwrite the changes when it closes."
        Show-ThemedMessageBox -Message $msg -Title "Close qBittorrent" -Icon 'Warning'
        return $false
    }

    $iniPath = Join-Path $env:APPDATA "qBittorrent\qBittorrent.ini"
    if (-not (Test-Path $iniPath)) {
        Show-ThemedMessageBox -Message "Could not find qBittorrent.ini at:`n$iniPath" -Title "Error" -Icon 'Error'
        return $false
    }

    try {
        $iniLines = [System.IO.File]::ReadAllLines($iniPath)

        # Check current state first
        $isAlreadyInstalled = $false
        $inAutoRunCheck = $false
        foreach ($line in $iniLines) {
            if ($line -match '^\[(.*)\]$') {
                $inAutoRunCheck = ($line -eq '[AutoRun]')
            } elseif ($inAutoRunCheck) {
                if ($line -match '^program\s*=(.*)$') {
                    if ($Matches[1] -match 'qBitLauncher') {
                        $isAlreadyInstalled = $true
                        break
                    }
                }
            }
        }

        if (-not $isAlreadyInstalled) {
            Show-ThemedMessageBox -Message "qBitLauncher is not currently integrated with qBittorrent.`n`nNothing to remove." -Title "Info" -Icon 'Information'
            return $true
        }

        $confirm = Show-ThemedMessageBox -Message "This will remove qBitLauncher from qBittorrent's Run-on-Completion settings.`n`nAre you sure you want to proceed?" -Title "Remove Integration" -Icon 'Warning' -Buttons 'YesNo'
        if ($confirm -ne 'Yes') {
            return $true
        }

        $newLines = @()
        $inAutoRun = $false
        $updated = $false

        foreach ($line in $iniLines) {
            if ($line -match '^\[(.*)\]$') {
                $inAutoRun = ($line -eq '[AutoRun]')
                $newLines += $line
            }
            elseif ($inAutoRun) {
                if ($line -match '^enabled\s*=') {
                    $newLines += "enabled=false"
                    $updated = $true
                }
                elseif ($line -match '^program\s*=') {
                    $newLines += "program="
                    $updated = $true
                }
                else {
                    $newLines += $line
                }
            }
            else {
                $newLines += $line
            }
        }

        if ($updated) {
            $integrationBackup = $iniPath + ".qBitLauncher-integration-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".bak"
        [IO.File]::Copy($iniPath, $integrationBackup, $true)
        [System.IO.File]::WriteAllLines($iniPath, $newLines)
            Show-ThemedMessageBox -Message "Successfully removed integration from qBittorrent!" -Title "Success" -Icon 'Information'
        } else {
            Show-ThemedMessageBox -Message "No integration found in qBittorrent settings." -Title "Info" -Icon 'Information'
        }
        return $true
    }
    catch {
        Show-ThemedMessageBox -Message "Failed to update qBittorrent configuration:`n$($_.Exception.Message)" -Title "Error" -Icon 'Error'
        return $false
    }
}

function Remove-ContextMenu {
    try {
        $basePaths = @(
            "HKCU:\Software\Classes\Directory\shell\qBitLauncher",
            "HKCU:\Software\Classes\Directory\Background\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.zip\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.rar\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.7z\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.iso\shell\qBitLauncher",
            "HKCU:\Software\Classes\SystemFileAssociations\.img\shell\qBitLauncher",

            # Legacy paths from older versions (HKEY_CLASSES_ROOT)
            "Registry::HKEY_CLASSES_ROOT\Directory\shell\qBitLauncher",
            "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\qBitLauncher",
            "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.zip\shell\qBitLauncher",
            "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.rar\shell\qBitLauncher",
            "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.7z\shell\qBitLauncher",
            "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.iso\shell\qBitLauncher",
            "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.img\shell\qBitLauncher"
        )

        $removed = $false
        foreach ($path in $basePaths) {
            if (Test-Path $path) {
                Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
                $removed = $true
            }
        }

        if ($removed) {
            Show-ThemedMessageBox -Message "Successfully removed 'Open with qBitLauncher' from Context Menus!" -Title "Success" -Icon 'Information'
        } else {
            Show-ThemedMessageBox -Message "Context menus were not currently installed." -Title "Info" -Icon 'Information'
        }
        return $true
    }
    catch {
        Show-ThemedMessageBox -Message "Failed to remove context menus:`n$($_.Exception.Message)" -Title "Error" -Icon 'Error'
        return $false
    }
}

function Show-InstallerGUI {
    $colors = $Global:CurrentTheme

    $form = New-Object System.Windows.Forms.Form
    $form.AutoScaleMode = "Dpi"
    $form.Text = "qBitLauncher Setup"
    $form.Size = New-Object System.Drawing.Size(500, 360)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.BackColor = $colors.FormBack
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Add_Shown({ Set-FormIcon -Form $this })

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Welcome to qBitLauncher"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = $colors.TextFore
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(400, 35)
    $form.Controls.Add($titleLabel)

    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "Choose an installation folder. It is highly recommended to install qBitLauncher to a permanent location so integrations don't break if you clear your Downloads folder."
    $descLabel.ForeColor = $colors.TextFore
    $descLabel.Location = New-Object System.Drawing.Point(24, 60)
    $descLabel.Size = New-Object System.Drawing.Size(440, 60)
    $form.Controls.Add($descLabel)

    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Text = "Destination Folder:"
    $pathLabel.ForeColor = $colors.TextFore
    $pathLabel.Location = New-Object System.Drawing.Point(24, 130)
    $pathLabel.Size = New-Object System.Drawing.Size(440, 20)
    $form.Controls.Add($pathLabel)

    $pathBox = New-Object System.Windows.Forms.TextBox
    $pathBox.Text = Join-Path $env:LOCALAPPDATA "qBitLauncher"
    $pathBox.Location = New-Object System.Drawing.Point(24, 155)
    $pathBox.Size = New-Object System.Drawing.Size(350, 25)
    $pathBox.BackColor = $colors.ControlBack
    $pathBox.ForeColor = $colors.TextFore
    $pathBox.BorderStyle = 'FixedSingle'
    $form.Controls.Add($pathBox)

    $browseBtn = New-Object System.Windows.Forms.Button
    $browseBtn.Text = "Browse..."
    $browseBtn.Location = New-Object System.Drawing.Point(380, 154)
    $browseBtn.Size = New-Object System.Drawing.Size(80, 27)
    Set-ThemedButton -Button $browseBtn -Colors $colors
    $browseBtn.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select Installation Folder"
        $dialog.SelectedPath = $pathBox.Text
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $pathBox.Text = $dialog.SelectedPath
        }
    })
    $form.Controls.Add($browseBtn)

    $desktopShortcut = New-Object System.Windows.Forms.CheckBox
    $desktopShortcut.Text = "Create Desktop Shortcut"
    $desktopShortcut.Checked = $true
    $desktopShortcut.ForeColor = $colors.TextFore
    $desktopShortcut.Location = New-Object System.Drawing.Point(24, 195)
    $desktopShortcut.Size = New-Object System.Drawing.Size(200, 25)
    $form.Controls.Add($desktopShortcut)

    $startShortcut = New-Object System.Windows.Forms.CheckBox
    $startShortcut.Text = "Create Start Menu Shortcut"
    $startShortcut.Checked = $true
    $startShortcut.ForeColor = $colors.TextFore
    $startShortcut.Location = New-Object System.Drawing.Point(230, 195)
    $startShortcut.Size = New-Object System.Drawing.Size(220, 25)
    $form.Controls.Add($startShortcut)

    $installBtn = New-Object System.Windows.Forms.Button
    $installBtn.Text = "Install"
    $installBtn.Location = New-Object System.Drawing.Point(340, 260)
    $installBtn.Size = New-Object System.Drawing.Size(120, 35)
    $installBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    Set-ThemedButton -Button $installBtn -Colors $colors
    $installBtn.BackColor = $colors.Accent
    $installBtn.ForeColor = $colors.FormBack

    $script:result = 'Cancel'

    $installBtn.Add_Click({
        try {
            $targetDir = [IO.Path]::GetFullPath($pathBox.Text)
            [IO.Directory]::CreateDirectory($targetDir) | Out-Null
            $scriptPath = $Global:ExecutablePath
            if (-not [IO.File]::Exists($scriptPath)) { throw "The running application could not be located: $scriptPath" }
            $fileName = [IO.Path]::GetFileName($scriptPath)
            $targetFile = Join-Path $targetDir $fileName
            if (-not [string]::Equals($scriptPath, $targetFile, [StringComparison]::OrdinalIgnoreCase)) {
                [IO.File]::Copy($scriptPath, $targetFile, $true)
            }

            # Create shortcuts
            $shortcutTarget = if ($Global:IsCompiledExe) { $targetFile } else { Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe' }
            $shortcutArguments = if ($Global:IsCompiledExe) { '' } else { '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $targetFile + '"' }
            if ($desktopShortcut.Checked) {
                $desktopPath = [Environment]::GetFolderPath('Desktop')
                New-WindowsShortcut -ShortcutPath (Join-Path $desktopPath "qBitLauncher.lnk") -TargetPath $shortcutTarget -WorkingDirectory $targetDir -Arguments $shortcutArguments
            }
            if ($startShortcut.Checked) {
                $startPath = Join-Path ([Environment]::GetFolderPath('Programs')) "qBitLauncher"
                if (-not (Test-Path -LiteralPath $startPath)) { [IO.Directory]::CreateDirectory($startPath) | Out-Null }
                New-WindowsShortcut -ShortcutPath (Join-Path $startPath "qBitLauncher.lnk") -TargetPath $shortcutTarget -WorkingDirectory $targetDir -Arguments $shortcutArguments
            }

            # Preserve existing settings on reinstall and carry portable settings forward.
            $targetConfig = Join-Path $targetDir "config.json"
            if (-not [IO.File]::Exists($targetConfig)) {
                if ([IO.File]::Exists($Global:ConfigFile)) { [IO.File]::Copy($Global:ConfigFile, $targetConfig) }
                else { [IO.File]::WriteAllText($targetConfig, "{}", [Text.Encoding]::UTF8) }
            }
            if ($Global:IsCompiledExe) { Start-Process -FilePath $targetFile -ErrorAction Stop }
            else { Start-Process -FilePath powershell.exe -WindowStyle Hidden -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "' + $targetFile + '"') -ErrorAction Stop }

            $script:result = 'Installed'
            $form.Close()
        } catch {
            Show-ThemedMessageBox -Message "Installation failed: $($_.Exception.Message)" -Title "Error" -Icon 'Error'
        }
    })
    $form.Controls.Add($installBtn)

    $portableBtn = New-Object System.Windows.Forms.Button
    $portableBtn.Text = "Run Portably (Skip Install)"
    $portableBtn.Location = New-Object System.Drawing.Point(24, 260)
    $portableBtn.Size = New-Object System.Drawing.Size(200, 35)
    Set-ThemedButton -Button $portableBtn -Colors $colors
    $portableBtn.Add_Click({
        $script:result = 'Portable'
        $form.Close()
    })
    $form.Controls.Add($portableBtn)

    $form.ShowDialog() | Out-Null
    $form.Dispose()

    return $result
}


# ===================================================================
# MAIN SCRIPT LOGIC STARTS HERE
# ===================================================================

# Startup checks run in a worker; no network request blocks the first window.
$Global:PendingUpdateCheck = Start-BackgroundUpdateCheck
$Global:WorkflowLock = $null
try {
    if ([string]::IsNullOrWhiteSpace($filePathFromQB)) {
        if (-not (Test-Path -LiteralPath $Global:ConfigFile)) {
            $installChoice = Show-InstallerGUI
            if ($installChoice -eq 'Installed' -or $installChoice -eq 'Cancel') { return }
            [void](Save-UserSettings)
        }
        do { $runManual = Show-SetupMenu } while ($runManual -eq [Windows.Forms.DialogResult]::Retry)
        if ($runManual -ne [Windows.Forms.DialogResult]::OK) { return }
        $manualPath = $Global:PendingManualFolder
        if (-not $manualPath) {
            $initialFolder = if ($Global:UserSettings.RecentFolders.Count) { $Global:UserSettings.RecentFolders[0] } else { [Environment]::GetFolderPath('Desktop') }
            $manualPath = Select-ExtractionPath -DefaultPath $initialFolder
        }
        if (-not $manualPath) { return }
        $filePathFromQB = $manualPath
    }
    if (-not (Test-Path -LiteralPath $filePathFromQB)) { throw "Path not found: $filePathFromQB" }
    $filePathFromQB = [IO.Path]::GetFullPath($filePathFromQB)
    # Different games run independently. Only duplicate source requests conflict.
    $jobKey = 'path:' + (Get-CanonicalPath $filePathFromQB).ToLowerInvariant()
    $jobLabel = if ($torrentNameFromQB) { $torrentNameFromQB } else { [IO.Path]::GetFileName($filePathFromQB.TrimEnd([char]92)) }
    $Global:WorkflowLock = New-WorkflowLock -Key $jobKey
    if ($Global:WorkflowLock.Duplicate) {
        Show-ThemedMessageBox -Message "This source is already open in qBitLauncher:`n$filePathFromQB" -Title 'Already open' -Icon Information | Out-Null
        return
    }
    Get-UserSettings
    $contentItem = Get-Item -LiteralPath $filePathFromQB -ErrorAction Stop
    $rootFolder = if ($contentItem.PSIsContainer) { $contentItem.FullName } else { $contentItem.DirectoryName }
    [void](Add-RecentFolder $rootFolder)
    $inventory = $null; $archive = $null; $runnables = @()
    if ($contentItem.PSIsContainer) {
        $inventory = Get-InventoryResponsive $rootFolder
        $primary = Get-PrimaryContent $inventory
        if ($primary.Kind -eq 'Runnables') { $runnables = @($primary.Files) }
        elseif ($primary.Kind -eq 'Archives') {
            if ($primary.Files.Count -gt 1 -or $inventory.Runnables.Count) {
                $choice = Show-ArchiveChoice -Files $primary.Files -HasRunnables ($inventory.Runnables.Count -gt 0)
                if (-not $choice) { return }
                if ($choice -is [string] -and $choice -eq 'Runnables') { $runnables = @($inventory.Runnables) } else { $archive = $choice }
            } else { $archive = $primary.Files[0] }
        } else {
            Show-ThemedMessageBox -Message 'No executable or supported archive was found. You can open this folder in Explorer.' -Title 'Folder scanned' -Icon Information | Out-Null
            Start-Process explorer.exe -ArgumentList ('"' + $rootFolder + '"'); return
        }
    } elseif ($contentItem.Extension -in @('.exe', '.bat', '.cmd')) { $runnables = @($contentItem) }
    elseif ($contentItem.Extension.TrimStart('.').ToLowerInvariant() -in $ArchiveExtensions -or $contentItem.Name -match '(?i)\.(?:7z|zip)\.001$') { $archive = $contentItem }
    else { Start-Process explorer.exe -ArgumentList ('"' + $rootFolder + '"'); return }
    if ($archive) {
        $baseName = $archive.BaseName -replace '(?i)\.(?:7z|zip)$', '' -replace '(?i)\.part0*1$', ''
        $baseFolder = if ($Global:UserSettings.LastExtractionPath) { $Global:UserSettings.LastExtractionPath } else { $archive.DirectoryName }
        $extraction = Show-ExtractionConfirmForm -ArchivePath $archive.FullName -DefaultDestination (Join-Path $baseFolder $baseName)
        if ($extraction.Confirmed) {
            $destination = Expand-ArchiveFile -ArchivePath $archive.FullName -DestinationPath $extraction.DestinationPath
            if (-not $destination) { return }
            $rootFolder = $destination
            $Global:UserSettings.LastExtractionPath = [IO.Path]::GetDirectoryName($destination.TrimEnd([char]92)); [void](Save-UserSettings)
            $runnables = @(Get-AllRunnables $rootFolder)
        } elseif ($extraction.UseExisting) {
            $rootFolder = $extraction.DestinationPath; $runnables = @(Get-AllRunnables $rootFolder)
        } elseif ($extraction.SkippedExtraction) {
            $runnables = if ($inventory) { @($inventory.Runnables) } else { @(Get-AllRunnables $rootFolder) }
        } else { return }
    }
    if ($runnables.Count) {
        Show-ExecutableSelectionForm -FoundExecutables $runnables -WindowTitle ('qBitLauncher — ' + $jobLabel) -RootFolder $rootFolder -TorrentHash $torrentHashFromQB -TorrentName $torrentNameFromQB
    } else {
        Show-ThemedMessageBox -Message 'No runnable files were found in the selected folder.' -Title 'Scan complete' -Icon Information | Out-Null
        Start-Process explorer.exe -ArgumentList ('"' + $rootFolder + '"')
    }
}
catch [OperationCanceledException] { Write-LogMessage 'Workflow cancelled.' }
catch { Write-LogMessage ('Workflow failed: ' + $_.Exception.Message); Show-ThemedMessageBox -Message (Protect-LogText $_.Exception.Message) -Title 'qBitLauncher' -Icon Error | Out-Null }
finally {
    if ($Global:WorkflowLock) { Exit-WorkflowLock $Global:WorkflowLock }
    if ($Global:PendingUpdateCheck) { Stop-BackgroundUpdateCheck $Global:PendingUpdateCheck; $Global:PendingUpdateCheck = $null }
    Write-LogMessage 'Workflow finished.'
}
