# Fixture checks for the bundler and atomic publication helpers; never builds or launches the app.
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$tokens = $null
$parseErrors = $null
$buildAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $projectRoot 'Build.ps1'), [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join '; ') }
foreach ($definition in $buildAst.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
    Invoke-Expression $definition.Extent.Text
}

$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ('qbitlauncher-build-tests-' + [Guid]::NewGuid().ToString('N'))
$script:checks = 0
function Assert-BuildTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}
function Assert-BuildRejects {
    param([scriptblock]$Action, [string]$Message)
    $rejected = $false
    try { & $Action | Out-Null } catch { $rejected = $true }
    Assert-BuildTest -Condition $rejected -Message $Message
}

try {
    $fixtureSource = Join-Path $testRoot 'Source'
    [IO.Directory]::CreateDirectory($fixtureSource) | Out-Null
    $template = "param([string]`$InputPath)`n# <BUNDLED_FUNCTIONS>`n`$Global:ScriptVersion = `"1.2.3`"`n"
    [IO.File]::WriteAllText((Join-Path $fixtureSource 'App.ps1'), $template)
    $moduleNames = @('Services', 'Updates', 'Workflow', 'UI')
    foreach ($name in $moduleNames) {
        [IO.File]::WriteAllText((Join-Path $fixtureSource ($name + '.ps1')), ('function Test' + $name + ' { return 1 }'))
    }
    $bundle = Get-BundledApplicationSource -ProjectRoot $testRoot
    Assert-BuildTest ($bundle.Contains('function TestServices') -and $bundle.Contains('function TestUpdates') -and $bundle.Contains('function TestWorkflow') -and $bundle.Contains('function TestUI')) 'The bundle omitted a development module.'
    Assert-BuildTest (-not $bundle.Contains('# <BUNDLED_FUNCTIONS>')) 'The bundle still contains its insertion marker.'
    Assert-BuildTest ($bundle.IndexOf('function TestServices') -lt $bundle.IndexOf('function TestUpdates') -and $bundle.IndexOf('function TestUpdates') -lt $bundle.IndexOf('function TestWorkflow') -and $bundle.IndexOf('function TestWorkflow') -lt $bundle.IndexOf('function TestUI')) 'Module order changed.'
    Assert-BuildTest ((Get-BuildVersion -SourceText $bundle).ToString() -eq '1.2.3') 'The version was not read from the template.'
    Assert-BuildRejects { Get-BuildVersion -SourceText '$Global:ScriptVersion = "70000.1.2"' } 'Invalid Windows version metadata was accepted.'
    Assert-BuildRejects { Assert-BuildScriptSyntax -Text 'function Broken {' } 'Invalid PowerShell syntax was accepted.'

    $sourceOutput = Join-Path $testRoot 'standalone.ps1'
    $unicodeText = $bundle + '# ' + (-join [char[]]@(0x6e38, 0x620f))
    Write-BuildSource -Text $unicodeText -Path $sourceOutput
    $bytes = [IO.File]::ReadAllBytes($sourceOutput)
    Assert-BuildTest ($bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf) 'Generated source needs a UTF-8 BOM for Windows PowerShell 5.1.'
    Assert-BuildTest ([IO.File]::ReadAllText($sourceOutput) -ceq $unicodeText) 'Unicode source changed during publication.'
    Write-BuildSource -Text $bundle -Path $sourceOutput
    Assert-BuildTest ([IO.File]::ReadAllText($sourceOutput) -ceq $bundle) 'Atomic source replacement failed.'

    [IO.File]::WriteAllText((Join-Path $fixtureSource 'UI.ps1'), 'Write-Output "unexpected startup"')
    Assert-BuildRejects { Get-BundledApplicationSource -ProjectRoot $testRoot } 'Startup side effects were accepted in a definition-only module.'
    [IO.File]::WriteAllText((Join-Path $fixtureSource 'UI.ps1'), 'function TestServices { return 2 }')
    Assert-BuildRejects { Get-BundledApplicationSource -ProjectRoot $testRoot } 'Duplicate bundled functions were accepted.'
    [IO.File]::WriteAllText((Join-Path $fixtureSource 'App.ps1'), $template.Replace('# <BUNDLED_FUNCTIONS>', ''))
    Assert-BuildRejects { Get-BundledApplicationSource -ProjectRoot $testRoot } 'A missing bundle marker was accepted.'

    $destination = Join-Path $testRoot 'qbit-build-fixture.bin'
    [IO.File]::WriteAllText($destination, 'original')
    $staged = Join-Path $testRoot 'staged.bin'
    Assert-BuildRejects { Install-BuildArtifact -StagedPath $staged -DestinationPath $destination } 'A missing staged build was accepted.'
    Assert-BuildTest ([IO.File]::ReadAllText($destination) -ceq 'original') 'Failed publication damaged the existing executable.'
    [IO.File]::WriteAllText($staged, 'replacement')
    $backup = Install-BuildArtifact -StagedPath $staged -DestinationPath $destination
    Assert-BuildTest ([IO.File]::ReadAllText($destination) -ceq 'replacement' -and [IO.File]::ReadAllText($backup) -ceq 'original') 'Publication did not preserve a rollback copy.'
    [IO.File]::WriteAllText($staged, '')
    Assert-BuildRejects { Install-BuildArtifact -StagedPath $staged -DestinationPath $destination } 'An empty staged executable was accepted.'
    Assert-BuildTest ([IO.File]::ReadAllText($destination) -ceq 'replacement') 'An empty staged build replaced a working executable.'

    $hashFile = Join-Path $testRoot 'hash.txt'
    [IO.File]::WriteAllBytes($hashFile, [Text.Encoding]::ASCII.GetBytes('abc'))
    Assert-BuildTest ((Get-BuildFileHash -Path $hashFile) -ceq 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad') 'SHA-256 output does not match the known vector.'

    if ($SourcePath) {
        $generatedText = [IO.File]::ReadAllText([IO.Path]::GetFullPath($SourcePath))
        Assert-BuildScriptSyntax -Text $generatedText | Out-Null
        Assert-BuildTest (-not $generatedText.Contains('# <BUNDLED_FUNCTIONS>')) 'The actual generated application still depends on a bundle marker.'
    }
    [pscustomobject]@{ Passed = $true; Checks = $script:checks; PowerShellVersion = $PSVersionTable.PSVersion.ToString() }
}
finally {
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    if ($resolvedRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedRoot).StartsWith('qbitlauncher-build-tests-') -and [IO.Directory]::Exists($resolvedRoot)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
