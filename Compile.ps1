# Compatibility entry point for existing developer instructions.
[CmdletBinding()]
param(
    [switch]$SkipTests,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$buildArguments = @{ SkipTests = $SkipTests }
if ($OutputPath) { $buildArguments.OutputPath = $OutputPath }
& (Join-Path $PSScriptRoot 'Build.ps1') @buildArguments
