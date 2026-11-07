# Runs without opening qBitLauncher, changing the desktop, or launching targets.
param(
    [string]$SourcePath,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = Join-Path $PSScriptRoot '..\qBitLauncher.ps1'
}

$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-shortcut-tests-' + [Guid]::NewGuid().ToString('N'))
$shell = $null
$script:checks = 0

function Assert-Shortcut {
    param([string]$LinkPath, [string]$ExpectedTarget, [string]$ExpectedDirectory)

    if (-not [IO.File]::Exists($LinkPath)) { throw "Shortcut was not saved: $LinkPath" }

    # Read through Explorer's automation API independently of the writer.
    $folder = $shell.NameSpace([IO.Path]::GetDirectoryName($LinkPath))
    $item = $folder.ParseName([IO.Path]::GetFileName($LinkPath))
    if (-not $item.IsLink) { throw "Windows did not recognize the shortcut: $LinkPath" }
    $link = $item.GetLink
    try {
        if ($link.Path -cne $ExpectedTarget) { throw "Shortcut target changed: $LinkPath -> $($link.Path)" }
        if ($link.WorkingDirectory -cne $ExpectedDirectory) { throw "Shortcut working directory changed: $LinkPath" }
        $script:checks++
    }
    finally {
        foreach ($comObject in @($link, $item, $folder)) {
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) | Out-Null
        }
    }
}

try {
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile([IO.Path]::GetFullPath($SourcePath), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw ($parseErrors.Message -join "`n") }
    $definition = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-WindowsShortcut'
    }, $false)
    if (-not $definition) { throw 'New-WindowsShortcut was not found in the source.' }
    # Import only this function; the app's startup/update/integration code never runs.
    Invoke-Expression $definition.Extent.Text

    [IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $shell = New-Object -ComObject Shell.Application
    $names = @(
        'English',
        (-join [char[]]@(0x6e38, 0x620f)),
        (-join [char[]]@(0x30b2, 0x30fc, 0x30e0)),
        (-join [char[]]@(0x0418, 0x0433, 0x0440, 0x0430)),
        (-join [char[]]@(0xac8c, 0xc784)),
        ('Caf' + [char]0xe9),
        'game [test] & 100%',
        ('game ' + [char]::ConvertFromUtf32(0x1f3ae))
    )
    $systemExe = Join-Path $env:WINDIR 'System32\whoami.exe'

    foreach ($name in $names) {
        foreach ($placement in @('FileName', 'TargetFolder', 'ShortcutFolder', 'Combined')) {
            $caseRoot = Join-Path $testRoot ([Guid]::NewGuid().ToString('N'))
            $targetFolderName = if ($placement -in @('TargetFolder', 'Combined')) { $name } else { 'target' }
            $linkFolderName = if ($placement -in @('ShortcutFolder', 'Combined')) { $name } else { 'links' }
            $fileName = if ($placement -in @('FileName', 'Combined')) { $name } else { 'game' }
            $targetDir = Join-Path (Join-Path $caseRoot 'files') $targetFolderName
            $linkDir = Join-Path (Join-Path $caseRoot 'shortcuts') $linkFolderName
            [IO.Directory]::CreateDirectory($targetDir) | Out-Null
            [IO.Directory]::CreateDirectory($linkDir) | Out-Null
            $target = Join-Path $targetDir ($fileName + '.exe')
            [IO.File]::Copy($systemExe, $target)
            $linkPath = Join-Path $linkDir ($fileName + '.lnk')
            New-WindowsShortcut -ShortcutPath $linkPath -TargetPath $target
            Assert-Shortcut -LinkPath $linkPath -ExpectedTarget $target -ExpectedDirectory $targetDir
        }
    }

    # Batch runnable types and an explicitly supplied working directory.
    foreach ($extension in @('.bat', '.cmd')) {
        $target = Join-Path $testRoot ('script [test]' + $extension)
        [IO.File]::WriteAllText($target, '@echo off')
        $linkPath = Join-Path $testRoot ($extension.TrimStart('.') + '.lnk')
        New-WindowsShortcut -ShortcutPath $linkPath -TargetPath $target -WorkingDirectory $testRoot
        Assert-Shortcut -LinkPath $linkPath -ExpectedTarget $target -ExpectedDirectory $testRoot
    }

    # Preserve the existing button's overwrite behavior and surface invalid input.
    $target = Join-Path $testRoot 'replacement.exe'
    [IO.File]::Copy($systemExe, $target)
    New-WindowsShortcut -ShortcutPath $linkPath -TargetPath $target
    Assert-Shortcut -LinkPath $linkPath -ExpectedTarget $target -ExpectedDirectory $testRoot

    # Launch arguments are stored verbatim, including Unicode and shell symbols.
    $expectedArguments = '--title "' + (-join [char[]]@(0x6e38, 0x620f)) + '" --path "C:\100% & more\"'
    New-WindowsShortcut -ShortcutPath $linkPath -TargetPath $target -Arguments $expectedArguments
    $argumentFolder = $shell.NameSpace($testRoot)
    $argumentItem = $argumentFolder.ParseName([IO.Path]::GetFileName($linkPath))
    $argumentLink = $argumentItem.GetLink
    try {
        if ($argumentLink.Arguments -cne $expectedArguments) { throw 'Shortcut arguments were changed.' }
        $script:checks++
    } finally {
        foreach ($comObject in @($argumentLink, $argumentItem, $argumentFolder)) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) | Out-Null }
    }

    $missingTargetRejected = $false
    try { New-WindowsShortcut -ShortcutPath $linkPath -TargetPath (Join-Path $testRoot 'missing.exe') }
    catch { $missingTargetRejected = $true }
    if (-not $missingTargetRejected) { throw 'A missing target was accepted.' }
    $script:checks++

    $missingDirectoryRejected = $false
    try { New-WindowsShortcut -ShortcutPath $linkPath -TargetPath $target -WorkingDirectory (Join-Path $testRoot 'missing') }
    catch { $missingDirectoryRejected = $true }
    if (-not $missingDirectoryRejected) { throw 'A missing working directory was accepted.' }
    $script:checks++

    $result = [pscustomobject]@{ Passed = $true; Checks = $script:checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
    if ($ResultPath) { [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json)) }
    else { $result }
}
catch {
    if ($ResultPath) {
        [IO.File]::WriteAllText($ResultPath, (@{ Passed = $false; Checks = $script:checks; Error = $_.Exception.Message } | ConvertTo-Json))
        exit 1
    }
    throw
}
finally {
    if ($shell) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell) | Out-Null }
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    if ($resolvedRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedRoot).StartsWith('qbitlauncher-shortcut-tests-') -and
        [IO.Directory]::Exists($resolvedRoot)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
