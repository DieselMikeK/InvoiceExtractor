$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

Write-Host 'Running syntax checks...'
python -m py_compile update_utils.py updater_app.py invoice_extractor_gui.py

Write-Host 'Building updater helper...'
python -m PyInstaller --noconfirm InvoiceExtractorUpdater.spec

Write-Host 'Building main application...'
python -m PyInstaller --noconfirm InvoiceExtractor.spec

$version = (Get-Content VERSION -Raw).Trim()
$mainExe = Join-Path $repoRoot 'dist\InvoiceExtractor.exe'
$updaterExe = Join-Path $repoRoot 'dist\InvoiceExtractorUpdater.exe'
$releaseTemplatePath = Join-Path $repoRoot 'dist\release-template.json'
$runtimeUpdateDir = Join-Path $repoRoot 'update'
$runtimeUpdaterExe = Join-Path $runtimeUpdateDir 'InvoiceExtractorUpdater.exe'

if (-not (Test-Path $mainExe)) {
    throw 'Main build did not produce dist\InvoiceExtractor.exe'
}
if (-not (Test-Path $updaterExe)) {
    throw 'Updater build did not produce dist\InvoiceExtractorUpdater.exe'
}

$sha256 = (Get-FileHash -Algorithm SHA256 $mainExe).Hash.ToLowerInvariant()
$releaseTemplate = [ordered]@{
    version      = $version
    download_url = ''
    sha256       = $sha256
    notes        = ''
    published_at = ''
}
$releaseTemplate | ConvertTo-Json | Set-Content -Path $releaseTemplatePath -Encoding UTF8

New-Item -ItemType Directory -Force -Path $runtimeUpdateDir | Out-Null
foreach ($legacyPath in @(
    (Join-Path $runtimeUpdateDir 'python'),
    (Join-Path $runtimeUpdateDir 'build_log.txt'),
    (Join-Path $runtimeUpdateDir 'updater.py'),
    (Join-Path $runtimeUpdateDir 'updater.bat')
)) {
    if (Test-Path $legacyPath) {
        Remove-Item -Recurse -Force $legacyPath
    }
}
Copy-Item -Force $updaterExe $runtimeUpdaterExe

$workspaceRoot = Split-Path -Parent $repoRoot
Copy-Item -Force $mainExe (Join-Path $workspaceRoot 'InvoiceExtractor.exe')
$rootUpdaterExe = Join-Path $workspaceRoot 'InvoiceExtractorUpdater.exe'
if (Test-Path $rootUpdaterExe) {
    Remove-Item -Force $rootUpdaterExe
}

Write-Host ''
Write-Host "Build complete."
Write-Host "Main app:      $mainExe"
Write-Host "Updater helper:$runtimeUpdaterExe"
Write-Host "Release JSON:  $releaseTemplatePath"
Write-Host "SHA-256:       $sha256"
