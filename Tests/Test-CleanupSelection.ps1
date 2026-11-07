param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
. (Join-Path $PSScriptRoot '..\Source\UI.ps1')
$script:checks = 0
$script:calls = New-Object System.Collections.ArrayList
$script:messages = New-Object System.Collections.ArrayList
$script:failApi = $false
$script:failFiles = $false
$Global:UserSettings = @{}

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
    [void]$script:calls.Add(@{ Kind = 'Torrent'; Hash = $TorrentHash; DeleteFiles = $DeleteFiles })
    if ($script:failApi) { throw 'API fixture unavailable' }
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
    try { $null = Invoke-CleanupSelection -Choice $choice -TorrentHash ('a' * 40) -OnMessage { param($message) [void]$script:messages.Add($message) } } catch { $rejected = $true }
    Assert-CleanupSelection ($rejected -and $script:messages.Contains('qBittorrent entry removed')) 'A local-file failure must preserve the reported successful entry removal.'
    $script:failFiles = $false; $script:calls.Clear(); $choice.RemoveTorrent = $false; $choice.Permanent = $true
    $null = Invoke-CleanupSelection -Choice $choice
    Assert-CleanupSelection ($script:calls.Count -eq 1 -and $script:calls[0].Permanent) 'The explicit permanent-file choice must remain available.'
    # Verify the real API wrapper's form body as well as the orchestration calls.
    $tokens = $null; $errors = $null
    $appAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Source\App.ps1'), [ref]$tokens, [ref]$errors)
    $apiDefinition = $appAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Remove-QBittorrentTorrent' }, $false)[0]
    Invoke-Expression $apiDefinition.Extent.Text
    function Invoke-WebRequest {
        param($Uri, $Method, $Body, $ContentType, $WebSession, [switch]$UseBasicParsing, $TimeoutSec, $ErrorAction)
        $script:apiBody = $Body
    }
    $Global:UserSettings = @{ QbittorrentWebUrl = 'http://127.0.0.1:8080'; QbittorrentUsername = '' }
    Remove-QBittorrentTorrent -TorrentHash ('b' * 40) -DeleteFiles $false
    Assert-CleanupSelection ($script:apiBody.deleteFiles -ceq 'false' -and $script:apiBody.hashes -eq ('b' * 40)) 'Entry-only API calls must explicitly keep downloaded files.'
    Remove-QBittorrentTorrent -TorrentHash ('b' * 40) -DeleteFiles $true
    Assert-CleanupSelection ($script:apiBody.deleteFiles -ceq 'true') 'The API wrapper must serialize an explicit deletion flag correctly.'
    $result = @{ Passed = $true; Checks = $script:checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString(); Suite = 'CleanupSelection' }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { [pscustomobject]$result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'CleanupSelection' } | ConvertTo-Json)); exit 1 }
    throw
}
