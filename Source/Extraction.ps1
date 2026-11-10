# Synchronous entry points for metadata verification. The scan UI uses the same
# native operations in its cancellable task, without blocking the UI thread.
function Get-ArchiveOutputManifest {
    param([string]$ArchivePath, [string]$ListToolPath, [int]$TimeoutMilliseconds = 15000)
    Initialize-FolderScanner
    $entries = [QBitLauncher.FolderScan]::ReadArchiveManifest($ArchivePath, $ListToolPath, $TimeoutMilliseconds, [Threading.CancellationToken]::None)
    if ($null -eq $entries) { return $null }
    return ,$entries
}

function Test-ArchiveOutput {
    param([object[]]$Manifest, [string]$DestinationPath)
    Initialize-FolderScanner
    $entries = New-Object 'System.Collections.Generic.List[QBitLauncher.ArchiveOutputFile]'
    foreach ($entry in $Manifest) { $entries.Add((New-Object QBitLauncher.ArchiveOutputFile([string]$entry.Path, [long]$entry.Length))) }
    return [QBitLauncher.FolderScan]::ValidateArchiveOutput($entries.ToArray(), $DestinationPath, [Threading.CancellationToken]::None, $null)
}
