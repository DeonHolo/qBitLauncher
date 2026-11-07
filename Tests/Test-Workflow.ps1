# Core workflow checks use disposable fixtures. No game or user configuration is changed.
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
. (Join-Path $repository 'Source\Services.ps1')
. (Join-Path $repository 'Source\Workflow.ps1')
$Global:ScriptDir = $repository
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-workflow-tests-' + [Guid]::NewGuid().ToString('N'))
$script:checks = 0
$script:junctions = New-Object 'System.Collections.Generic.List[string]'
function Assert-Workflow {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}
function Assert-WorkflowThrows {
    param([scriptblock]$Action, [string]$Message)
    $threw = $false
    try { $null = & $Action } catch { $threw = $true }
    Assert-Workflow $threw $Message
}
function New-WorkflowFixture {
    param([string]$RelativePath, [string]$Content = 'fixture')
    $path = Join-Path $testRoot $RelativePath
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path)) | Out-Null
    [IO.File]::WriteAllText($path, $Content, [Text.Encoding]::UTF8)
    return $path
}
function Assert-CleanupPreserves {
    param($Plan, [string]$Path)
    $overlaps = @($Plan.Targets | Where-Object { (Test-PathContains -Parent $_.Path -Child $Path) -or (Test-PathContains -Parent $Path -Child $_.Path) })
    Assert-Workflow ($overlaps.Count -eq 0) ('A source-only target overlaps a preserved path: ' + $Path)
}

try {
    [IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $selected = New-WorkflowFixture 'source\Extracted\game.exe'
    $archive = New-WorkflowFixture 'source\installer.rar'
    $notes = New-WorkflowFixture 'source\installer-files\instructions.txt'
    $source = Join-Path $testRoot 'source'
    $extracted = Join-Path $source 'Extracted'
    $sibling = Join-Path $testRoot 'source2'
    [IO.Directory]::CreateDirectory($sibling) | Out-Null
    Assert-Workflow (Test-PathContains -Parent $source -Child $selected) 'A real descendant was not recognized.'
    Assert-Workflow (Test-PathContains -Parent ($source.ToUpperInvariant() + '\') -Child $selected) 'Case or a trailing separator changed containment.'
    Assert-Workflow (-not (Test-PathContains -Parent $source -Child $sibling)) 'A sibling with the same prefix was treated as a descendant.'
    Assert-Workflow (-not (Test-PathContains -Parent $selected -Child $source)) 'Containment was reversed.'

    $sourcePlan = Get-CleanupPlan -SourcePath $source -ExtractedPath $extracted -SelectedPath $selected -Mode SourceOnly
    Assert-Workflow (-not $sourcePlan.RequiresAcknowledgement) 'Source-only unexpectedly affects the selected application.'
    Assert-Workflow ($sourcePlan.Targets.Count -eq 2) 'Source-only did not identify both independent installer targets.'
    Assert-CleanupPreserves -Plan $sourcePlan -Path $extracted
    Assert-CleanupPreserves -Plan $sourcePlan -Path $selected
    Assert-Workflow (@($sourcePlan.Retained | Where-Object { $_ -eq $extracted }).Count -eq 1) 'The extracted tree was not listed as retained.'

    $everything = Get-CleanupPlan -SourcePath $source -ExtractedPath $extracted -SelectedPath $selected -Mode Everything
    Assert-Workflow ($everything.Targets.Count -eq 1) 'Everything did not prune overlapping parent/child targets.'
    Assert-Workflow ($everything.Targets[0].CanonicalPath -eq (Get-CanonicalPath $source)) 'Everything did not include the source tree.'
    Assert-Workflow ([bool]$everything.RequiresAcknowledgement) 'Removing the selected game did not require acknowledgement.'
    Assert-WorkflowThrows { Invoke-CleanupPlan -Plan $everything -PermanentlyDelete } 'Everything deleted the selected app without acknowledgement.'
    Assert-Workflow ([IO.File]::Exists($selected)) 'A rejected cleanup changed the selected file.'

    $extractedPlan = Get-CleanupPlan -SourcePath $source -ExtractedPath $extracted -SelectedPath $selected -Mode ExtractedOnly
    Assert-Workflow ($extractedPlan.Targets.Count -eq 1) 'Extracted-only did not identify its own root.'
    Assert-Workflow ($extractedPlan.Targets[0].CanonicalPath -eq (Get-CanonicalPath $extracted)) 'Extracted-only included an installer source path.'
    Assert-Workflow ([bool]$extractedPlan.RequiresAcknowledgement) 'Extracted-only lacked the selected-app acknowledgement.'
    Assert-WorkflowThrows { Invoke-CleanupPlan -Plan $extractedPlan -PermanentlyDelete } 'Extracted-only ran without acknowledgement.'

    $sameRoot = Get-CleanupPlan -SourcePath ($source + '\') -ExtractedPath $source.ToUpperInvariant() -SelectedPath $selected -Mode SourceOnly
    Assert-Workflow ($sameRoot.Targets.Count -eq 0) 'Equal source/extracted roots were not preserved in source-only mode.'
    $sameEverything = Get-CleanupPlan -SourcePath $source -ExtractedPath $source -SelectedPath $selected -Mode Everything
    Assert-Workflow ($sameEverything.Targets.Count -eq 1 -and $sameEverything.RequiresAcknowledgement) 'Equal roots prevented intentional delete-everything.'

    $selectedOnly = New-WorkflowFixture 'selected-only\launch.cmd'
    $selectedNotes = New-WorkflowFixture 'selected-only\installer.txt'
    $selectedPlan = Get-CleanupPlan -SourcePath (Split-Path $selectedOnly -Parent) -SelectedPath $selectedOnly -Mode SourceOnly
    Assert-CleanupPreserves -Plan $selectedPlan -Path $selectedOnly
    Assert-Workflow ($selectedPlan.Targets.Count -eq 1 -and $selectedPlan.Targets[0].Path -eq $selectedNotes) 'Protecting one runnable also hid unrelated source files.'

    # Validate forbidden roots and forged previews before invoking deletion.
    foreach ($protected in @([IO.Path]::GetPathRoot($testRoot), $env:USERPROFILE, $env:WINDIR, $repository)) {
        Assert-WorkflowThrows { Test-CleanupRoot -Path $protected } ('A protected root was allowed: ' + $protected)
    }
    $outsideFixture = New-WorkflowFixture 'outside-approved\keep.txt'
    $forged = [pscustomobject]@{
        Mode = 'Everything'; RequiresAcknowledgement = $false; ApprovedRoots = @($source)
        Targets = @([pscustomobject]@{ Path = $outsideFixture; CanonicalPath = (Get-CanonicalPath $outsideFixture); IsDirectory = $false })
    }
    Assert-WorkflowThrows { Invoke-CleanupPlan -Plan $forged -PermanentlyDelete } 'A forged target outside the approved root was deleted.'
    Assert-Workflow ([IO.File]::Exists($outsideFixture)) 'A rejected forged preview changed the outside fixture.'
    $changed = [pscustomobject]@{
        Mode = 'Everything'; RequiresAcknowledgement = $false; ApprovedRoots = @($source)
        Targets = @([pscustomobject]@{ Path = $archive; CanonicalPath = $outsideFixture; IsDirectory = $false })
    }
    Assert-WorkflowThrows { Invoke-CleanupPlan -Plan $changed -PermanentlyDelete } 'A changed target canonical path was accepted.'
    Assert-Workflow ([IO.File]::Exists($archive)) 'A changed-preview rejection removed an installer file.'

    # Actual deletion is limited to the tested disposable root, with no Recycle Bin.
    $fullFixture = [IO.Path]::GetFullPath($testRoot)
    Assert-Workflow ($fullFixture.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($fullFixture).StartsWith('qbitlauncher-workflow-tests-')) 'Unsafe fixture directory.'
    $result = Invoke-CleanupPlan -Plan $sourcePlan -PermanentlyDelete
    Assert-Workflow ($result.Deleted.Count -eq 2) 'Source-only did not execute its previewed targets.'
    Assert-Workflow (-not [IO.File]::Exists($archive) -and -not [IO.File]::Exists($notes)) 'Installer files were not removed.'
    Assert-Workflow ([IO.File]::Exists($selected)) 'Source-only deletion removed the selected application.'
    $intentional = Get-CleanupPlan -SourcePath $source -ExtractedPath $extracted -SelectedPath $selected -Mode ExtractedOnly
    $null = Invoke-CleanupPlan -Plan $intentional -AcknowledgeSelectedFiles -PermanentlyDelete
    Assert-Workflow (-not [IO.Directory]::Exists($extracted)) 'Acknowledged extracted-only deletion was blocked.'
    Assert-Workflow ([IO.Directory]::Exists($source)) 'Extracted-only also removed the source directory.'
    $allTarget = New-WorkflowFixture 'intentional-all\Extracted\selected.bat'
    $allSource = Split-Path (Split-Path $allTarget -Parent) -Parent
    $allPlan = Get-CleanupPlan -SourcePath $allSource -ExtractedPath (Split-Path $allTarget -Parent) -SelectedPath $allTarget -Mode Everything
    $null = Invoke-CleanupPlan -Plan $allPlan -AcknowledgeSelectedFiles -PermanentlyDelete
    Assert-Workflow (-not [IO.Directory]::Exists($allSource)) 'Acknowledged delete-everything remained blocked.'

    # A junction points only to another disposable fixture. The plan and executor
    # must preserve the link and its target rather than recurse into it.
    $linkedFile = New-WorkflowFixture 'linked-target\do-not-touch.txt'
    $linkRoot = Join-Path $testRoot 'junction-source'
    [IO.Directory]::CreateDirectory($linkRoot) | Out-Null
    $linkPath = Join-Path $linkRoot 'linked-folder'
    $junction = New-Item -ItemType Junction -Path $linkPath -Target (Split-Path $linkedFile -Parent) -ErrorAction Stop
    $script:junctions.Add($linkPath)
    $regular = New-WorkflowFixture 'junction-source\installer.zip'
    Assert-Workflow ((Get-CanonicalPath $linkPath) -eq (Get-CanonicalPath (Split-Path $linkedFile -Parent))) 'Junction target was not resolved canonically.'
    $junctionPlan = Get-CleanupPlan -SourcePath $linkRoot -Mode Everything
    Assert-Workflow (@($junctionPlan.Retained | Where-Object { $_ -eq $linkPath }).Count -eq 1) 'A junction was not retained.'
    Assert-Workflow ($junctionPlan.Targets.Count -eq 1 -and $junctionPlan.Targets[0].Path -eq $regular) 'A junction-containing directory was selected wholesale.'
    $null = Invoke-CleanupPlan -Plan $junctionPlan -AcknowledgeSelectedFiles -PermanentlyDelete
    Assert-Workflow ([IO.File]::Exists($linkedFile)) 'Cleanup followed a junction into its linked target.'
    Assert-Workflow ([IO.Directory]::Exists($linkPath)) 'Cleanup unexpectedly deleted the retained junction.'
    $linkInventory = Get-FolderInventory -RootFolder $linkRoot
    Assert-Workflow ($linkInventory.Skipped -contains $linkPath) 'Discovery did not report the skipped junction.'

    # Changes after preview are checked again before deletion. Both linked and
    # unlinked sentinels live under the checked disposable fixture root.
    $lateSentinel = New-WorkflowFixture 'changed-after-preview\keep.txt'
    $lateRoot = Split-Path $lateSentinel -Parent
    $latePlan = Get-CleanupPlan -SourcePath $lateRoot -Mode Everything
    $lateLink = Join-Path $lateRoot 'new-junction'
    $null = New-Item -ItemType Junction -Path $lateLink -Target (Split-Path $linkedFile -Parent) -ErrorAction Stop
    $script:junctions.Add($lateLink)
    Assert-WorkflowThrows { Invoke-CleanupPlan -Plan $latePlan -PermanentlyDelete } 'A junction added after preview was not rejected.'
    Assert-Workflow ([IO.File]::Exists($lateSentinel) -and [IO.File]::Exists($linkedFile)) 'A changed-tree rejection deleted fixture files.'
    $replacedRoot = Join-Path $testRoot 'replaced-after-preview'
    [IO.Directory]::CreateDirectory($replacedRoot) | Out-Null
    $replacedPlan = Get-CleanupPlan -SourcePath $replacedRoot -Mode Everything
    [IO.Directory]::Delete($replacedRoot, $false)
    $null = New-Item -ItemType Junction -Path $replacedRoot -Target (Split-Path $linkedFile -Parent) -ErrorAction Stop
    $script:junctions.Add($replacedRoot)
    Assert-WorkflowThrows { Invoke-CleanupPlan -Plan $replacedPlan -PermanentlyDelete } 'A cleanup root replaced by a junction was accepted.'
    Assert-Workflow ([IO.File]::Exists($linkedFile)) 'A root substitution deleted linked content.'

    # Shortcut labels are metadata, with literal paths and collision-safe names.
    $unicode = (-join [char[]]@(0x6e38, 0x620f, 0x30b2, 0x30fc, 0x30e0))
    Assert-Workflow ((ConvertTo-ShortcutLabel (' ' + $unicode + '.lnk ')) -ceq $unicode) 'Unicode shortcut label was changed.'
    Assert-Workflow ((ConvertTo-ShortcutLabel 'Game [test] & 100%') -ceq 'Game [test] & 100%') 'Valid special characters were removed from a label.'
    Assert-Workflow ((ConvertTo-ShortcutLabel 'console') -ceq 'console') 'An ordinary label was confused with a reserved device name.'
    foreach ($invalid in @('', 'CON', 'nul.txt', 'COM1', 'LPT9.lnk', '..', 'bad.', 'folder\name', 'a/b', 'a:b', 'a?b', 'a*b', ('control' + [char]1), ('a' * 181))) {
        Assert-WorkflowThrows { ConvertTo-ShortcutLabel $invalid } ('Invalid shortcut label was accepted: ' + $invalid)
    }
    $linkDir = Join-Path $testRoot 'desktop-fixture'
    [IO.Directory]::CreateDirectory($linkDir) | Out-Null
    $firstLink = Get-AvailableShortcutPath -Directory $linkDir -Label $unicode
    Assert-Workflow ([IO.Path]::GetFileName($firstLink) -ceq ($unicode + '.lnk')) 'First shortcut path was not the requested label.'
    [IO.File]::WriteAllText($firstLink, 'existing shortcut fixture')
    $secondLink = Get-AvailableShortcutPath -Directory $linkDir -Label $unicode
    Assert-Workflow ([IO.Path]::GetFileName($secondLink) -ceq ($unicode + ' (2).lnk')) 'Keep-both did not select the second name.'
    [IO.File]::WriteAllText($secondLink, 'second existing shortcut fixture')
    Assert-Workflow ([IO.Path]::GetFileName((Get-AvailableShortcutPath -Directory $linkDir -Label $unicode)) -ceq ($unicode + ' (3).lnk')) 'Keep-both did not skip all existing labels.'

    # Discovery prefers a real application over nested asset archives, retains
    # runnable scripts, and identifies only the first archive volume.
    $game = New-WorkflowFixture 'ready\Game.exe'
    $batch = New-WorkflowFixture 'ready\launch.bat'
    $command = New-WorkflowFixture 'ready\launch.cmd'
    $helper = New-WorkflowFixture 'ready\UnityCrashHandler64.exe'
    $asset = New-WorkflowFixture 'ready\Plugins\icons.zip'
    $readyRoot = Split-Path $game -Parent
    $inventory = Get-FolderInventory -RootFolder $readyRoot
    Assert-Workflow ($inventory.Runnables.Count -eq 4) 'EXE/BAT/CMD discovery missed runnable types.'
    Assert-Workflow ($inventory.AllArchives.Count -eq 1 -and $inventory.Archives.Count -eq 0) 'Nested plugin archive was treated as a payload.'
    Assert-Workflow ((Get-PrimaryContent $inventory).Kind -eq 'Runnables') 'A ready application prompted for its nested asset archive.'
    Assert-Workflow ($inventory.Runnables[-1].Name -eq 'UnityCrashHandler64.exe') 'Crash helper was not ranked below launchers.'
    Assert-Workflow (-not (Get-RunnableRank -File (Get-Item -LiteralPath $game) -RootFolder $readyRoot).Helper) 'Ordinary main executable was classified as a helper.'
    Assert-Workflow ((Get-RunnableRank -File (Get-Item -LiteralPath $helper) -RootFolder $readyRoot).Helper) 'Crash handler was not classified as a helper.'
    Assert-Workflow ((Get-RunnableRank -File (Get-Item -LiteralPath $batch) -RootFolder $readyRoot).Reason -eq 'Launch script') 'BAT was not described as a launch script.'
    Assert-Workflow ([IO.File]::Exists($game) -and [IO.File]::Exists($batch) -and [IO.File]::Exists($command)) 'Label/discovery operations renamed a runnable.'
    foreach ($gameName in @('Crashlands.exe', 'Crash Bandicoot.exe')) {
        $crashGame = New-WorkflowFixture ('ready\' + $gameName)
        Assert-Workflow (-not (Get-RunnableRank -File (Get-Item -LiteralPath $crashGame) -RootFolder $readyRoot).Helper) ('A game title was mistaken for a crash handler: ' + $gameName)
    }
    $rootArchive = New-WorkflowFixture 'ready\payload.zip'
    $inventory = Get-FolderInventory -RootFolder $readyRoot
    Assert-Workflow ((Get-PrimaryContent $inventory).Kind -eq 'Archives') 'A root payload archive was ignored.'
    $otherRootArchive = New-WorkflowFixture 'ready\other.7z'
    $inventory = Get-FolderInventory -RootFolder $readyRoot
    Assert-Workflow ((Get-PrimaryContent $inventory).Files.Count -eq 2) 'Ambiguous root payloads did not remain available for a choice.'
    $part1 = New-WorkflowFixture 'volumes\payload.part01.rar'
    $null = New-WorkflowFixture 'volumes\payload.part02.rar'
    $null = New-WorkflowFixture 'volumes\payload.part10.rar'
    $null = New-WorkflowFixture 'volumes\payload.part101.rar'
    $first7z = New-WorkflowFixture 'volumes\other.7z.001'
    $null = New-WorkflowFixture 'volumes\other.7z.002'
    $volumes = Get-FolderInventory -RootFolder (Split-Path $part1 -Parent)
    Assert-Workflow ($volumes.Archives.Count -eq 2) 'Multipart discovery did not isolate first volumes.'
    Assert-Workflow (@($volumes.Archives | Where-Object { $_.FullName -eq $part1 }).Count -eq 1) 'RAR first volume was missing.'
    Assert-Workflow (@($volumes.Archives | Where-Object { $_.FullName -eq $first7z }).Count -eq 1) '7z first volume was missing.'

    # Import only the existing INI transformer; never execute app startup.
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repository 'Source\App.ps1'), [ref]$tokens, [ref]$parseErrors)
    $iniFunction = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Set-QBittorrentIniValue' }, $false)
    if (-not $iniFunction) { throw 'The INI transformer could not be found.' }
    Invoke-Expression $iniFunction.Extent.Text
    $ini = @('[Other]', 'UnrelatedKey=preserve', '[Preferences]', 'WebUI\Enabled=false', 'WebUI\Port=1234', 'WebUI\Address=*', 'WebUI\Username=my-account', 'WebUI\Password_PBKDF2=my-preserved-hash', 'WebUI\LocalHostAuth=true', 'WebUI\HTTPS\Enabled=true', 'WebUI\HostHeaderValidation=true')
    $repaired = Get-QBittorrentRepairLines -Lines $ini -Port 8089
    foreach ($line in @('UnrelatedKey=preserve', 'WebUI\Username=my-account', 'WebUI\Password_PBKDF2=my-preserved-hash', 'WebUI\LocalHostAuth=true', 'WebUI\HTTPS\Enabled=true', 'WebUI\HostHeaderValidation=true')) {
        Assert-Workflow ($repaired -ccontains $line) ('API repair changed an unrelated/authentication setting: ' + $line)
    }
    Assert-Workflow ($repaired -ccontains 'WebUI\Enabled=true') 'API repair did not enable the local endpoint.'
    Assert-Workflow ($repaired -ccontains 'WebUI\Port=8089') 'API repair did not select the requested port.'
    Assert-Workflow ($repaired -ccontains 'WebUI\Address=127.0.0.1') 'API repair did not bind loopback.'
    Assert-Workflow (@($repaired | Where-Object { $_ -match '^WebUI\\Port=' }).Count -eq 1) 'API repair duplicated a setting.'
    Assert-WorkflowThrows { Get-QBittorrentRepairLines -Lines $ini -Port 0 } 'Invalid port zero was accepted.'
    Assert-WorkflowThrows { Get-QBittorrentRepairLines -Lines $ini -Port 65536 } 'Invalid port above 65535 was accepted.'
    [pscustomobject]@{ Passed = $true; Checks = $script:checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
}
finally {
    # Unlink known junctions without recursion before deleting our checked root.
    foreach ($linkPath in $script:junctions) {
        if ([IO.Directory]::Exists($linkPath)) { [IO.Directory]::Delete($linkPath, $false) }
    }
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    if ($resolvedRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedRoot).StartsWith('qbitlauncher-workflow-tests-') -and [IO.Directory]::Exists($resolvedRoot)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
