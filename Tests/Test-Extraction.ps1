param([string]$SourcePath, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Source\Workflow.ps1')
. (Join-Path $PSScriptRoot '..\Source\Extraction.ps1')
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-extraction-tests-' + [Guid]::NewGuid().ToString('N'))
$script:checks = 0
function Assert-Extraction { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message }; $script:checks++ }
try {
    $destination = Join-Path $testRoot ('Output [test] ' + (-join [char[]]@(0x6e38, 0x620f)))
    [IO.Directory]::CreateDirectory($destination) | Out-Null
    $file = Join-Path $destination 'setup.exe'; [IO.File]::WriteAllBytes($file, [byte[]]@(1,2,3))
    $payload = Join-Path $destination 'setup-1.bin'; [IO.File]::WriteAllBytes($payload, [byte[]]@(4,5,6,7))
    $manifest = @(@{ Path = 'setup.exe'; Length = 3 }, @{ Path = 'setup-1.bin'; Length = 4 })
    $result = Test-ArchiveOutput -Manifest $manifest -DestinationPath $destination
    Assert-Extraction ($result.Valid -and $result.Checked -eq 2) 'Complete extracted files were not accepted.'
    [IO.File]::WriteAllBytes($payload, [byte[]]@())
    $result = Test-ArchiveOutput -Manifest $manifest -DestinationPath $destination
    Assert-Extraction (-not $result.Valid -and $result.Problems[0].Contains('Incomplete: setup-1.bin')) 'An empty installer payload was reported as extracted successfully.'
    [IO.File]::WriteAllBytes($payload, [byte[]]@(4,5))
    Assert-Extraction (-not (Test-ArchiveOutput -Manifest $manifest -DestinationPath $destination).Valid) 'A partially extracted nonempty payload was accepted.'
    [IO.File]::Delete($payload)
    $result = Test-ArchiveOutput -Manifest $manifest -DestinationPath $destination
    Assert-Extraction (-not $result.Valid -and $result.Problems[0].Contains('Missing: setup-1.bin')) 'Missing extracted data was accepted.'
    Assert-Extraction (-not (Test-ArchiveOutput -Manifest $manifest -DestinationPath (Join-Path $testRoot 'wrong destination')).Valid) 'A different output folder was accepted as the extraction destination.'
    $threw = $false
    try { $null = Test-ArchiveOutput -Manifest @(@{ Path = '..\outside.exe'; Length = 3 }) -DestinationPath $destination } catch { $threw = $true }
    Assert-Extraction $threw 'Archive paths outside the destination must not be checked as valid output.'
    Assert-Extraction (-not (Test-ArchiveOutput -Manifest @() -DestinationPath $destination).Valid) 'An empty file list cannot establish extraction success.'
    Assert-Extraction ($null -eq (Get-ArchiveOutputManifest -ArchivePath 'missing.zip' -ListToolPath '')) 'A missing listing tool should leave the result unverified.'
    $zip = Join-Path $testRoot 'content.zip'
    $archive = [IO.Compression.ZipFile]::Open($zip, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $manifest) {
            $item = $archive.CreateEntry($entry.Path)
            $stream = $item.Open()
            try { $stream.Write((New-Object byte[] ([int]$entry.Length)), 0, [int]$entry.Length) } finally { $stream.Dispose() }
        }
        [void]$archive.CreateEntry('empty/'); [void]$archive.CreateEntry('empty/zero.txt')
    } finally { $archive.Dispose() }
    $listTool = Join-Path $env:ProgramFiles '7-Zip\7z.exe'
    if (Test-Path -LiteralPath $listTool) {
        $actual = Get-ArchiveOutputManifest -ArchivePath $zip -ListToolPath $listTool
        Assert-Extraction ($actual.Count -eq 3 -and @($actual | Where-Object { $_.Path -eq 'setup-1.bin' -and $_.Length -eq 4 }).Count -eq 1) 'The listing parser lost expected file sizes or included directory entries.'
        Assert-Extraction (@($actual | Where-Object { $_.Length -eq 0 -and $_.Path.Replace('/', '\') -eq 'empty\zero.txt' }).Count -eq 1) 'Valid zero-length files must remain in the manifest.'
        $native = Start-FolderScan -RootFolder $destination -ArchivePath $zip -ListToolPath $listTool
        try {
            $raw = $native.Operation.GetAwaiter().GetResult()
            Assert-Extraction (-not $raw.OutputCheck.Valid -and $raw.OutputCheck.Problems.Count -eq 2) 'The asynchronous post-extraction scan accepted missing installer data.'
        } finally { $native.Dispose() }
        [IO.File]::WriteAllBytes($payload, (New-Object byte[] 4))
        [IO.Directory]::CreateDirectory((Join-Path $destination 'empty')) | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $destination 'empty\zero.txt'), [byte[]]@())
        $native = Start-FolderScan -RootFolder $destination -ArchivePath $zip -ListToolPath $listTool
        try {
            $raw = $native.Operation.GetAwaiter().GetResult()
            Assert-Extraction ($raw.Complete -and $raw.OutputCheck.Valid -and $raw.OutputCheck.Checked -eq 3) 'The asynchronous scanner did not validate complete extraction output.'
        } finally { $native.Dispose() }
    }
    # Return a bounded failure if a listing process never finishes or asks for input.
    $helper = Join-Path $testRoot 'slow-list.exe'
    Add-Type -TypeDefinition 'public static class SlowArchiveList { public static void Main(string[] args) { System.Threading.Thread.Sleep(30000); } }' -OutputAssembly $helper -OutputType ConsoleApplication
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $timeout = Get-ArchiveOutputManifest -ArchivePath $zip -ListToolPath $helper -TimeoutMilliseconds 100
    Assert-Extraction ($null -eq $timeout -and $watch.Elapsed.TotalSeconds -lt 3) 'An unresponsive metadata-listing process blocked extraction verification.'
    $native = Start-FolderScan -RootFolder $destination -ArchivePath $zip -ListToolPath $helper
    try {
        $deadline = [DateTime]::UtcNow.AddSeconds(3)
        while ($native.Progress.Phase -ne 'Checking archive file list' -and -not $native.Operation.IsCompleted -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 10 }
        Assert-Extraction ($native.Progress.Phase -eq 'Checking archive file list') 'Archive verification must publish its phase while the listing process is running.'
        $watch.Restart(); $native.Cancel()
        $deadline = [DateTime]::UtcNow.AddSeconds(2)
        while (-not $native.Operation.IsCompleted -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 10 }
        Assert-Extraction ($native.Operation.IsCompleted -and $watch.Elapsed.TotalSeconds -lt 2) 'Cancelling verification did not stop its metadata process promptly.'
    } finally { $native.Dispose() }
    $result = [pscustomobject]@{ Passed = $true; Checks = $script:checks; Suite = 'Extraction'; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) } else { $result }
}
catch {
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message; Suite = 'Extraction' } | ConvertTo-Json)); exit 1 }
    throw
}
finally {
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('qbitlauncher-extraction-tests-') -and [IO.Directory]::Exists($resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
