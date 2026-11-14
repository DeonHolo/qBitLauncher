param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
. (Join-Path $PSScriptRoot '..\Source\UI.ps1')
$script:checks = 0
$script:calls = New-Object System.Collections.ArrayList
$script:messages = New-Object System.Collections.ArrayList
$script:failApi = $false
$script:failFiles = $false
$script:apiError = 'API fixture unavailable'
$script:settingsSaved = $false
$script:settingsOpened = 0
$script:latestSettings = $null
$script:removedSessions = New-Object System.Collections.ArrayList
$Global:UserSettings = @{}
function Get-UserSettings { if ($script:latestSettings) { $Global:UserSettings = $script:latestSettings } }
function Set-LauncherSessionTorrentRemoved { param($Id); [void]$script:removedSessions.Add($Id) }
function Show-SettingsForm {
    param($ConnectionError)
    if ($ConnectionError -notmatch 'login failed|requires valid Web UI credentials') { throw 'Credential recovery must show the actual authentication error.' }
    $script:settingsOpened++
    if ($script:settingsSaved) { $script:failApi = $false }
    return $script:settingsSaved
}

function Assert-CleanupSelection {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}
function Invoke-ResponsiveTask {
    param($Title, [scriptblock]$Action, [object[]]$Arguments, $Functions, $Variables, [switch]$NoCancel)
    if (-not $NoCancel) { throw 'Cleanup must finish its active operation before closing.' }
    & $Action @Arguments
}
function Remove-QBittorrentTorrent {
    param($TorrentHash, [bool]$DeleteFiles)
    [void]$script:calls.Add(@{ Kind = 'Torrent'; Hash = $TorrentHash; DeleteFiles = $DeleteFiles; Username = $Global:UserSettings.QbittorrentUsername })
    if ($script:failApi) { throw $script:apiError }
}
function Invoke-QBittorrentLocalApiRepairFlow { return @{ Status = 'Cancelled'; Message = 'API fixture unavailable' } }
function Invoke-CleanupPlan {
    param($Plan, [switch]$AcknowledgeSelectedFiles, [switch]$PermanentlyDelete)
    [void]$script:calls.Add(@{ Kind = 'Files'; Plan = $Plan; Acknowledged = [bool]$AcknowledgeSelectedFiles; Permanent = [bool]$PermanentlyDelete })
    if ($script:failFiles) { throw 'File fixture unavailable' }
    return @{ Deleted = $Plan.Targets; Recycled = -not $PermanentlyDelete }
}

try {
    foreach ($mask in 1..7) {
        $script:calls.Clear(); $script:messages.Clear()
        $source = [bool]($mask -band 1); $extracted = [bool]($mask -band 2); $torrent = [bool]($mask -band 4)
        $plan = if ($source -or $extracted) { @{ RequiresAcknowledgement = $extracted; Targets = @('disposable location'); Mode = 'fixture' } } else { $null }
        $choice = @{ RemoveSource = $source; RemoveExtracted = $extracted; RemoveTorrent = $torrent; Plan = $plan; Acknowledged = $extracted; Permanent = $false }
        $result = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40) -OnMessage { param($message) [void]$script:messages.Add($message) }
        $api = @($script:calls | Where-Object { $_.Kind -eq 'Torrent' }); $files = @($script:calls | Where-Object { $_.Kind -eq 'Files' })
        Assert-CleanupSelection ($api.Count -eq [int]$torrent -and $files.Count -eq [int]($source -or $extracted)) "Combination $mask skipped an item or ran an unselected action."
        Assert-CleanupSelection ($result.TorrentRemoved -eq $torrent) "Combination $mask reported the wrong torrent result."
        if ($torrent) { Assert-CleanupSelection (-not $api[0].DeleteFiles -and $api[0].Hash -eq ('a' * 40)) 'Removing the entry must keep qBittorrent file deletion off.' }
        if ($files.Count) { Assert-CleanupSelection (-not $files[0].Permanent -and $files[0].Acknowledged -eq $extracted) 'Local cleanup lost the Recycle Bin or acknowledgement choice.' }
        if ($torrent -and $files.Count) { Assert-CleanupSelection ($script:calls[0].Kind -eq 'Torrent') 'Stop seeding and release file handles before local cleanup.' }
    }
    $choice = @{ RemoveSource = $true; RemoveExtracted = $true; RemoveTorrent = $true; Plan = @{ RequiresAcknowledgement = $true; Targets = @('disposable location') }; Acknowledged = $false; Permanent = $false }
    $script:calls.Clear(); $rejected = $false
    try { $null = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40) } catch { $rejected = $true }
    Assert-CleanupSelection ($rejected -and $script:calls.Count -eq 0) 'Missing app-deletion acknowledgement must block every side effect.'
    $choice.Acknowledged = $true
    $rejected = $false
    try { $null = Invoke-CleanupSelection -Choice $choice -TorrentHash '' } catch { $rejected = $true }
    Assert-CleanupSelection ($rejected -and $script:calls.Count -eq 0) 'A missing torrent identity must be detected before deleting files.'
    $script:failApi = $true; $rejected = $false
    try { $null = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40) } catch { $rejected = $true }
    Assert-CleanupSelection ($rejected -and @($script:calls | Where-Object { $_.Kind -eq 'Files' }).Count -eq 0) 'Failed entry removal must leave local files alone.'
    $script:failApi = $false; $script:failFiles = $true; $script:calls.Clear(); $script:messages.Clear(); $rejected = $false
    try { $null = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40) -SessionId 'fixture-session' -OnMessage { param($message) [void]$script:messages.Add($message) } } catch { $rejected = $true }
    Assert-CleanupSelection ($rejected -and $script:messages.Contains('qBittorrent entry removed')) 'A local-file failure must preserve the reported successful entry removal.'
    Assert-CleanupSelection ($script:removedSessions.Contains('fixture-session')) 'The saved session must remember torrent removal even when local file cleanup fails afterward.'
    $script:failFiles = $false; $script:calls.Clear(); $choice.RemoveTorrent = $false; $choice.Permanent = $true
    $null = Invoke-CleanupSelection -Choice $choice
    Assert-CleanupSelection ($script:calls.Count -eq 1 -and $script:calls[0].Permanent) 'The explicit permanent-file choice must remain available.'
    $choice.RemoveTorrent = $true; $script:calls.Clear(); $script:failApi = $true
    $script:apiError = 'qBittorrent login failed. Credentials rejected.'; $script:settingsSaved = $true
    $null = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40)
    Assert-CleanupSelection ($script:settingsOpened -eq 1 -and $script:calls.Count -eq 3 -and $script:calls[0].Kind -eq 'Torrent' -and $script:calls[1].Kind -eq 'Torrent' -and $script:calls[2].Kind -eq 'Files') 'Correcting a rejected login must retry the entry first, then carry out the original local plan.'
    Assert-CleanupSelection ($script:calls[2].Permanent -and $script:calls[2].Acknowledged -and $script:calls[2].Plan -eq $choice.Plan) 'Credential recovery must retain the original deletion choices and acknowledgement.'
    $script:failApi = $true; $script:settingsSaved = $false; $script:calls.Clear(); $cancelled = $false
    try { $null = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40) } catch [OperationCanceledException] { $cancelled = $true }
    Assert-CleanupSelection ($cancelled -and $script:calls.Count -eq 1 -and $script:calls[0].Kind -eq 'Torrent') 'Cancelling credential correction must leave files untouched and avoid another login.'
    $script:apiError = 'qBittorrent blocked the login request (HTTP 403). Temporary block.'
    $beforeSettings = $script:settingsOpened; $script:calls.Clear(); $rejected = $false
    try { $null = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40) } catch { $rejected = $true }
    Assert-CleanupSelection ($rejected -and $script:settingsOpened -eq $beforeSettings -and $script:calls.Count -eq 1) 'A temporary login block must not start another login attempt or delete files.'
    $script:failApi = $false; $script:latestSettings = @{ QbittorrentUsername = 'updated-in-another-window' }; $script:calls.Clear()
    $null = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40)
    Assert-CleanupSelection ($script:calls[0].Username -ceq 'updated-in-another-window') 'Cleanup must reload credentials saved by another open launcher window.'
    # Verify the real API wrapper's form body as well as the orchestration calls.
    $tokens = $null; $errors = $null
    $appAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Source\App.ps1'), [ref]$tokens, [ref]$errors)
    foreach ($name in @('Remove-QBittorrentTorrent', 'Connect-QBittorrentApi', 'Invoke-QBittorrentApiRequest')) {
        $apiDefinition = $appAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $false)[0]
        Invoke-Expression $apiDefinition.Extent.Text
    }
    function Invoke-WebRequest {
        param($Uri, $Method, $Body, $ContentType, $WebSession, $Headers, [switch]$UseBasicParsing, $TimeoutSec, $ErrorAction)
        if ($Uri -match '/app/version$') { return [pscustomobject]@{ Content = 'v5.2.4' } }
        $script:apiBody = $Body
    }
    $Global:UserSettings = @{ QbittorrentWebUrl = 'http://127.0.0.1:8080'; QbittorrentUsername = '' }
    Remove-QBittorrentTorrent -TorrentHash ('b' * 40) -DeleteFiles $false
    Assert-CleanupSelection ($script:apiBody.deleteFiles -ceq 'false' -and $script:apiBody.hashes -eq ('b' * 40)) 'Entry-only API calls must explicitly keep downloaded files.'
    Remove-QBittorrentTorrent -TorrentHash ('b' * 40) -DeleteFiles $true
    Assert-CleanupSelection ($script:apiBody.deleteFiles -ceq 'true') 'The API wrapper must serialize an explicit deletion flag correctly.'

    # Select all can toggle permanent deletion while its preview worker is still
    # running. Exercise the actual preview closures without displaying a form.
    $uiAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Source\UI.ps1'), [ref]$tokens, [ref]$errors)
    $cleanupDefinition = $uiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Show-CleanupConfirmForm' }, $false)
    $state = @{ Task = [object]::new(); Plan = $null; Error = $false; NeedsAck = $false }
    $source = [pscustomobject]@{ Checked = $true }; $extracted = [pscustomobject]@{ Checked = $true }; $torrent = [pscustomobject]@{ Checked = $true }
    $ack = [pscustomobject]@{ Checked = $true; Visible = $false }
    $permanent = [pscustomobject]@{ Checked = $true; Enabled = $false }
    $preview = [pscustomobject]@{ Text = 'Checking selected locations...' }; $confirm = [pscustomobject]@{ Enabled = $true }
    $layout = [pscustomobject]@{ RowStyles = @(@{ Height = 0 }, @{ Height = 0 }, @{ Height = 0 }, @{ Height = 0 }) }
    foreach ($name in @('updateConfirm', 'renderPreview')) {
        $assignment = $cleanupDefinition.Find({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left -is [Management.Automation.Language.VariableExpressionAst] -and $node.Left.VariablePath.UserPath -eq $name }, $true)
        if (-not $assignment) { throw "Missing cleanup preview callback: $name" }
        Invoke-Expression $assignment.Extent.Text
    }
    & $renderPreview
    Assert-CleanupSelection ($preview.Text -ceq 'Checking selected locations...' -and -not $confirm.Enabled) 'Permanent deletion changes must keep the preview pending until its worker finishes.'
    Assert-CleanupSelection ($ack.Checked -and $permanent.Checked) 'Waiting for the preview must retain explicit Select all choices.'
    $state.Task = $null; $state.NeedsAck = $true
    $state.Plan = @{ Targets = @(@{ CanonicalPath = 'fixture'; Path = 'fixture' }); ApprovedRoots = @(); Retained = @(); Mode = 'Everything' }
    & $renderPreview
    Assert-CleanupSelection ($preview.Text -match 'Permanently delete source and extracted files' -and $preview.Text -match 'entry from qBittorrent') 'A completed combined preview must show the selected removal method and scope.'
    Assert-CleanupSelection ($source.Checked -and $extracted.Checked -and $torrent.Checked -and $ack.Checked -and $permanent.Checked -and $confirm.Enabled) 'All five choices must remain selected and become usable only after the preview completes.'
    $result = @{ Passed = $true; Checks = $script:checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString(); Suite = 'CleanupSelection' }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { [pscustomobject]$result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'CleanupSelection' } | ConvertTo-Json)); exit 1 }
    throw
}
