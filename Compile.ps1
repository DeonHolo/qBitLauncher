# Compile.ps1
# Helper script to compile qBitLauncher.ps1 into qBitLauncher.exe with the correct icon.

if (-not (Get-Command ps2exe -ErrorAction SilentlyContinue)) {
    Write-Host "Installing PS2EXE module..."
    Install-Module -Name PS2EXE -Scope CurrentUser -Force
}

$scriptDir = $PSScriptRoot
$inFile = Join-Path $scriptDir "qBitLauncher.ps1"
$outFile = Join-Path $scriptDir "qBitLauncher.exe"
$iconFile = Join-Path $scriptDir "qBitLauncher Logo.ico"

Write-Host "Closing any running instances of qBitLauncher.exe..."
Stop-Process -Name "qBitLauncher" -Force -ErrorAction SilentlyContinue

Write-Host "Compiling $inFile to $outFile..."
ps2exe -inputFile $inFile -outputFile $outFile -iconFile $iconFile -noConsole

if ($LASTEXITCODE -eq 0 -or $?) {
    Write-Host "`nCompilation Successful!" -ForegroundColor Green
} else {
    Write-Host "`nCompilation Failed." -ForegroundColor Red
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
