# API and configuration fixtures only. No live qBittorrent requests or desktop interaction.
param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot '..\Source\Workflow.ps1')
$script:checks = 0
$script:requests = New-Object System.Collections.ArrayList
$script:dialogCount = 0
$script:startCount = 0
$script:processRunning = $true
$script:httpFailure = 0
$script:failLogin = $false
$script:failConnect = $false
$script:loginContent = "Ok.`n"
$transport = $null
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-api-tests-' + [Guid]::NewGuid().ToString('N'))
$script:configPath = Join-Path $testRoot 'qBittorrent.ini'

function Assert-Api {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}
function Get-ApiFailure {
    param([scriptblock]$Action)
    try { & $Action | Out-Null } catch { return $_.Exception.Message }
    throw 'The fixture request unexpectedly succeeded.'
}
function Set-ApiFixture {
    param([string[]]$Lines)
    [IO.File]::WriteAllLines($script:configPath, $Lines)
}
function Write-LogMessage { param($Message) }
function Get-QBittorrentConfigPath { return $script:configPath }
function Test-QBittorrentProcessRunning { return $script:processRunning }
function Test-QBittorrentApiEndpointAvailable { return $false }
function Get-QBittorrentExecutablePath { return 'fixture-only.exe' }
function Start-QBittorrentIfAvailable { param($ExecutablePath) $script:startCount++; return $true }
function Wait-QBittorrentApiEndpoint { param($TimeoutSeconds) return $false }
function Enable-QBittorrentLocalApiConfig { throw 'A regression fixture must never edit real configuration.' }
function Show-ThemedMessageBox {
    param($Message, $Title, $Buttons, $Icon)
    $script:dialogCount++; $script:lastDialog = $Message
    return [Windows.Forms.DialogResult]::No
}
function Invoke-WebRequest {
    param($Uri, $Method, $Body, $ContentType, $WebSession, $Headers, [switch]$UseBasicParsing, $TimeoutSec, $ErrorAction)
    [void]$script:requests.Add(@{ Uri = $Uri; Body = $Body; Session = $WebSession; Headers = $Headers })
    if ($script:failConnect) { throw 'Unable to connect to the remote server' }
    $isLogin = $Uri -match '/auth/login$'
    if ($script:httpFailure -and $isLogin -eq $script:failLogin) { throw (New-Object QBApiFixtureException $script:httpFailure) }
    return [pscustomobject]@{ Content = $(if ($isLogin) { $script:loginContent } else { '' }) }
}

try {
    [void][IO.Directory]::CreateDirectory($testRoot)
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Source\App.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors.Message -join "`n") }
    foreach ($name in @('Remove-QBittorrentTorrent', 'Get-QBittorrentApiBaseUrl', 'Get-QBittorrentApiUri', 'Test-QBittorrentApiUrlIsLocal', 'Get-QBittorrentApiPort', 'Get-QBittorrentWebApiState', 'Set-QBittorrentIniValue', 'Test-QBittorrentConnectionFailureMessage', 'Invoke-QBittorrentLocalApiRepairFlow', 'Stop-QBittorrentGracefully')) {
        $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $false)
        if (-not $definition) { throw "Missing function: $name" }
        Invoke-Expression $definition.Extent.Text
    }
    Add-Type -TypeDefinition @'
public class QBApiFixtureResponse { public int StatusCode { get; set; } }
public class QBApiFixtureException : System.Exception {
    public QBApiFixtureResponse Response { get; private set; }
    public QBApiFixtureException(int status) : base("The remote server returned an error.") {
        Response = new QBApiFixtureResponse { StatusCode = status };
    }
}
public sealed class QBApiTransportFixture : System.IDisposable {
    private readonly System.Net.Sockets.TcpListener listener;
    private readonly System.Threading.Thread thread;
    public readonly System.Collections.Generic.List<string> Requests = new System.Collections.Generic.List<string>();
    public int Port { get; private set; }
    public string Error { get; private set; }
    public QBApiTransportFixture() {
        listener = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 0);
        listener.Start();
        Port = ((System.Net.IPEndPoint)listener.LocalEndpoint).Port;
        thread = new System.Threading.Thread(Run);
        thread.IsBackground = true;
        thread.Start();
    }
    private void Run() {
        try {
            for (int i = 0; i < 2; i++) {
                using (var client = listener.AcceptTcpClient())
                using (var stream = client.GetStream())
                using (var reader = new System.IO.StreamReader(stream, System.Text.Encoding.ASCII, false, 1024, true)) {
                    stream.ReadTimeout = 5000;
                    var request = new System.Text.StringBuilder();
                    int length = 0;
                    for (string line; !string.IsNullOrEmpty(line = reader.ReadLine()); ) {
                        request.AppendLine(line);
                        if (line.StartsWith("Content-Length:", System.StringComparison.OrdinalIgnoreCase))
                            length = int.Parse(line.Substring(15).Trim());
                        if (line.Equals("Expect: 100-continue", System.StringComparison.OrdinalIgnoreCase)) {
                            byte[] interim = System.Text.Encoding.ASCII.GetBytes("HTTP/1.1 100 Continue\r\n\r\n");
                            stream.Write(interim, 0, interim.Length); stream.Flush();
                        }
                    }
                    var body = new char[length];
                    for (int read = 0; read < length; ) {
                        int count = reader.Read(body, read, length - read);
                        if (count == 0) throw new System.IO.EndOfStreamException();
                        read += count;
                    }
                    request.Append(body); Requests.Add(request.ToString());
                    string content = i == 0 ? "Ok.\n" : "";
                    string cookie = i == 0 ? "Set-Cookie: SID=fixture-session; Path=/\r\n" : "";
                    byte[] response = System.Text.Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\n" + cookie + "Content-Type: text/plain; charset=UTF-8\r\nContent-Length: " + content.Length + "\r\nConnection: close\r\n\r\n" + content);
                    stream.Write(response, 0, response.Length); stream.Flush();
                }
            }
        }
        catch (System.Exception ex) { Error = ex.Message; }
    }
    public void Dispose() { listener.Stop(); thread.Join(2000); }
}
'@
    $Global:UserSettings = @{ QbittorrentWebUrl = 'http://127.0.0.1:8080'; QbittorrentUsername = ''; QbittorrentPassword = '' }
    $legacy = @('[Preferences]', 'WebUI\Enabled=true', 'WebUI\Port=8181', 'WebUI\Address=127.0.0.1', 'WebUI\Username=fixture', 'WebUI\Password_PBKDF2=fixture-hash', 'General\CloseToTray=true')
    Set-ApiFixture $legacy
    $state = Get-QBittorrentWebApiState
    Assert-Api ($state.Exists -and $state.Enabled -and $state.Port -eq 8181 -and $state.ConfigurationSection -eq 'Preferences') 'Legacy Web UI settings were not read.'
    Assert-Api ($state.HasPassword -and $state.Username -eq 'fixture' -and $state.CloseToTray) 'Legacy authentication/tray settings were not read.'

    # Old repair keys can coexist with the current section. They must not override it.
    $modern = $legacy + @('[WebUI]', 'Enabled=false', 'Port=9191', 'Address=*', 'Username=modern-fixture', 'Password_PBKDF2=@ByteArray(fixture-hash)', 'LocalHostAuth=true', 'HTTPS\Enabled=true', 'HostHeaderValidation=true', 'CSRFProtection=true', '[Other]', 'Enabled=preserve', 'Port=preserve')
    Set-ApiFixture $modern
    $state = Get-QBittorrentWebApiState
    Assert-Api (-not $state.Enabled -and $state.Port -eq 9191 -and $state.Address -eq '*' -and $state.ConfigurationSection -eq 'WebUI') 'Legacy or unrelated keys overrode the current Web UI section.'
    Assert-Api ($state.HasPassword -and $state.Username -eq 'modern-fixture') 'Current authentication was not read from the Web UI section.'
    $repaired = Get-QBittorrentRepairLines -Lines $modern -Port 8080
    Set-ApiFixture $repaired
    $state = Get-QBittorrentWebApiState
    Assert-Api ($state.Enabled -and $state.Port -eq 8080 -and $state.Address -eq '127.0.0.1') 'Current-format repair did not update the active Web UI section.'
    foreach ($line in @('WebUI\Port=8181', 'Username=modern-fixture', 'Password_PBKDF2=@ByteArray(fixture-hash)', 'LocalHostAuth=true', 'HTTPS\Enabled=true', 'HostHeaderValidation=true', 'CSRFProtection=true', 'Enabled=preserve', 'Port=preserve')) {
        Assert-Api ($repaired -ccontains $line) ('Repair changed an unrelated setting: ' + $line)
    }
    Assert-Api (@($repaired | Where-Object { $_ -ceq '[WebUI]' }).Count -eq 1) 'Repair duplicated the current Web UI section.'

    foreach ($passwordLine in @('', 'Password_PBKDF2=', 'Password_PBKDF2=@ByteArray()')) {
        Set-ApiFixture ($legacy + @('[WebUI]', 'Enabled=true', 'Username=admin', $passwordLine))
        $state = Get-QBittorrentWebApiState
        Assert-Api (-not $state.HasPassword) 'A missing current password was hidden by the legacy hash or empty Qt value.'
        $before = [IO.File]::ReadAllText($script:configPath)
        $repair = Invoke-QBittorrentLocalApiRepairFlow -ErrorMessage 'Could not connect to the local API' -RetryAction { throw 'Missing credentials must not retry deletion.' }
        Assert-Api ($repair.Status -eq 'CredentialsRequired' -and $repair.Message -match 'Authentication' -and $repair.Message -match 'Settings') 'Missing credentials must explain both qBittorrent and launcher settings.'
        Assert-Api ($script:dialogCount -eq 0 -and $script:startCount -eq 0 -and [IO.File]::ReadAllText($script:configPath) -ceq $before) 'Missing credentials must not start, stop, prompt for config repair, or edit configuration.'
    }
    Set-ApiFixture @('[WebUI]', 'Enabled=true', 'Username=', 'Password_PBKDF2=fixture-hash')
    $repair = Invoke-QBittorrentLocalApiRepairFlow -ErrorMessage 'Could not connect' -RetryAction { throw 'Empty usernames must not retry deletion.' }
    Assert-Api ($repair.Status -eq 'CredentialsRequired') 'An empty username was not diagnosed.'

    Set-ApiFixture ($legacy + @('[WebUI]', 'Enabled=true', 'Username=admin', 'Password_PBKDF2=fixture-hash'))
    $repair = Invoke-QBittorrentLocalApiRepairFlow -ErrorMessage 'Could not connect' -RetryAction { throw 'Declined repair must not retry deletion.' }
    Assert-Api ($repair.Status -eq 'Declined' -and $script:startCount -eq 0) 'An already-running qBittorrent was launched again.'
    Assert-Api ($script:lastDialog -match '\[WebUI\]' -and $script:lastDialog -notmatch 'WebUI\\Enabled') 'The repair preview used obsolete keys for the current format.'

    foreach ($message in @('The remote server returned an error: (401) Unauthorized.', 'The remote server returned an error: (403) Forbidden.', 'API returned HTTP 404', 'API returned HTTP 500', 'qBittorrent login failed. Check credentials.')) {
        Assert-Api (-not (Test-QBittorrentConnectionFailureMessage $message)) 'An HTTP/authentication failure was mistaken for a missing API.'
    }
    Assert-Api (Test-QBittorrentConnectionFailureMessage 'Unable to connect to the remote server') 'A genuine connection failure was not recognized.'

    foreach ($status in @(401, 403, 404, 500)) {
        foreach ($login in @($false, $true)) {
            $script:requests.Clear(); $script:httpFailure = $status; $script:failLogin = $login
            $Global:UserSettings.QbittorrentUsername = $(if ($login) { 'fixture' } else { '' })
            $failure = Get-ApiFailure { Remove-QBittorrentTorrent -TorrentHash ('a' * 40) -DeleteFiles $false }
            if ($status -in @(401, 403)) { Assert-Api ($failure -match 'username/password') 'A rejected login/removal must request the existing credentials.' }
            else { Assert-Api ($failure -match "HTTP $status") 'An HTTP failure lost its status.' }
            Assert-Api (-not (Test-QBittorrentConnectionFailureMessage $failure)) 'HTTP failure incorrectly triggered connection repair.'
            Assert-Api ($script:requests.Count -eq 1) 'A failed login was followed by a delete request.'
        }
    }
    $script:httpFailure = 0; $script:loginContent = 'Fails.'; $script:requests.Clear()
    $failure = Get-ApiFailure { Remove-QBittorrentTorrent -TorrentHash ('a' * 40) -DeleteFiles $false }
    Assert-Api ($failure -match 'login failed' -and $script:requests.Count -eq 1) 'An unsuccessful login body was accepted.'

    $script:loginContent = "Ok.`n"; $script:requests.Clear()
    $Global:UserSettings.QbittorrentWebUrl = ' http://127.0.0.1:8080/proxy/ '
    $Global:UserSettings.QbittorrentPassword = 'fixture-secret'
    Remove-QBittorrentTorrent -TorrentHash (' ' + ('a' * 40) + ' ') -DeleteFiles $false
    Assert-Api ($script:requests.Count -eq 2 -and $script:requests[0].Uri -eq 'http://127.0.0.1:8080/proxy/api/v2/auth/login' -and $script:requests[1].Uri -eq 'http://127.0.0.1:8080/proxy/api/v2/torrents/delete') 'Authenticated removal did not use the configured URL.'
    Assert-Api ([object]::ReferenceEquals($script:requests[0].Session, $script:requests[1].Session)) 'Removal lost the login session.'
    Assert-Api ($script:requests[0].Body.username -eq 'fixture' -and $script:requests[0].Body.password -ceq 'fixture-secret') 'Login did not use the saved credentials.'
    Assert-Api ($script:requests[1].Body.hashes -eq ('a' * 40) -and $script:requests[1].Body.deleteFiles -ceq 'false') 'Entry removal changed its identity or enabled file deletion.'
    foreach ($request in $script:requests) {
        Assert-Api ($request.Headers.Origin -ceq 'http://127.0.0.1:8080' -and $request.Headers.Referer -ceq 'http://127.0.0.1:8080/proxy/') 'API requests did not identify the configured origin for CSRF validation.'
    }
    $script:failConnect = $true
    $failure = Get-ApiFailure { Remove-QBittorrentTorrent -TorrentHash ('a' * 40) -DeleteFiles $false }
    Assert-Api ($failure -match 'Could not connect' -and (Test-QBittorrentConnectionFailureMessage $failure)) 'Connection failures must remain eligible for repair.'
    $script:failConnect = $false
    $Global:UserSettings.QbittorrentWebUrl = 'file:///C:/fixture'
    $failure = Get-ApiFailure { Remove-QBittorrentTorrent -TorrentHash ('a' * 40) -DeleteFiles $false }
    Assert-Api ($failure -match 'URL is invalid') 'A non-HTTP API URL was accepted.'

    # Exercise the real Windows HTTP cmdlet against an isolated loopback fixture.
    Remove-Item -Path Function:\Invoke-WebRequest
    $transport = New-Object QBApiTransportFixture
    $Global:UserSettings.QbittorrentWebUrl = 'http://127.0.0.1:' + $transport.Port
    Remove-QBittorrentTorrent -TorrentHash ('a' * 40) -DeleteFiles $false
    Assert-Api ($transport.Requests.Count -eq 2 -and -not $transport.Error) 'The Windows HTTP transport did not complete login and removal.'
    Assert-Api ($transport.Requests[1] -match 'Cookie: SID=fixture-session' -and $transport.Requests[1] -match 'deleteFiles=false') 'The real removal request lost the login cookie or enabled file deletion.'
    foreach ($request in $transport.Requests) {
        Assert-Api ($request -match ('Origin: ' + [regex]::Escape($Global:UserSettings.QbittorrentWebUrl)) -and $request -match ('Referer: ' + [regex]::Escape($Global:UserSettings.QbittorrentWebUrl + '/'))) 'The Windows HTTP cmdlet did not send the configured origin and referer.'
    }
    $transport.Dispose(); $transport = $null

    # A process that ignores graceful close must remain running until the user exits it.
    $script:closeRequests = 0; $script:processPresent = $true
    $script:fakeProcess = [pscustomobject]@{ Id = 1 }
    $script:fakeProcess | Add-Member -MemberType ScriptMethod -Name CloseMainWindow -Value { $script:closeRequests++; return $true }
    function Get-Process { param($Name, $ErrorAction) if ($script:processPresent) { return $script:fakeProcess } }
    function Start-Process { throw 'Repair must never force-kill a process.' }
    function Start-Sleep { param($Milliseconds) }
    $stopped = Stop-QBittorrentGracefully -TimeoutSeconds 0
    Assert-Api (-not $stopped.Closed -and $stopped.WasRunning -and $script:closeRequests -eq 1) 'A running process was forced closed or reported closed before exit.'
    $script:processPresent = $false
    $stopped = Stop-QBittorrentGracefully -TimeoutSeconds 0
    Assert-Api ($stopped.Closed -and -not $stopped.WasRunning) 'An absent process was reported as running.'

    $result = @{ Passed = $true; Checks = $script:checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString(); Suite = 'QBittorrentApi' }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { [pscustomobject]$result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'QBittorrentApi' } | ConvertTo-Json)); exit 1 }
    throw
}
finally {
    if ($transport) { $transport.Dispose() }
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    if ($resolvedRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedRoot).StartsWith('qbitlauncher-api-tests-') -and [IO.Directory]::Exists($resolvedRoot)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
