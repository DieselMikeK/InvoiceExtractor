$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

Write-Host 'Running syntax checks...'
python -m py_compile update_utils.py updater_app.py invoice_extractor_gui.py
if ($LASTEXITCODE -ne 0) {
    throw 'Syntax checks failed.'
}

Write-Host 'Building updater helper...'
python -m PyInstaller --noconfirm InvoiceExtractorUpdater.spec
if ($LASTEXITCODE -ne 0) {
    throw 'Updater build failed.'
}

Write-Host 'Building main application...'
python -m PyInstaller --noconfirm InvoiceExtractor.spec
if ($LASTEXITCODE -ne 0) {
    throw 'Main application build failed.'
}

Write-Host 'Verifying packaged modules...'
$archiveListing = python -m PyInstaller.utils.cliutils.archive_viewer -r dist\InvoiceExtractor.exe
if ($LASTEXITCODE -ne 0) {
    throw 'Failed to inspect packaged modules.'
}
foreach ($requiredModule in @(
    'invoice_parser',
    'spreadsheet_writer',
    'gmail_client',
    'skunexus_client',
    'shopify_client',
    'update_utils'
)) {
    $moduleMatch = $archiveListing | Select-String -SimpleMatch "'$requiredModule'"
    if (-not $moduleMatch) {
        throw "Built InvoiceExtractor.exe is missing required bundled module '$requiredModule'."
    }
}

$version = (Get-Content VERSION -Raw).Trim()
$mainExe = Join-Path $repoRoot 'dist\InvoiceExtractor.exe'
$updaterExe = Join-Path $repoRoot 'dist\InvoiceExtractorUpdater.exe'
$releaseAssetsDir = Join-Path $repoRoot 'dist\release-assets'
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
$releaseFiles = [System.Collections.ArrayList]::new()
$null = $releaseFiles.Add([ordered]@{
    relative_path = 'InvoiceExtractor.exe'
    asset_name    = 'InvoiceExtractor.exe'
    sha256        = $sha256
})

if (Test-Path $releaseAssetsDir) {
    Remove-Item -Recurse -Force $releaseAssetsDir
}
New-Item -ItemType Directory -Force -Path $releaseAssetsDir | Out-Null

# Curated payload files that should live beside the installed app.
# User-specific folders such as required/, build/, and training/ are intentionally excluded.
foreach ($payloadFile in @(
    [ordered]@{
        source_path   = Join-Path $repoRoot 'vendors.csv'
        relative_path = 'app/vendors.csv'
        asset_name    = 'app-vendors.csv'
    },
    [ordered]@{
        source_path   = $updaterExe
        relative_path = 'update/InvoiceExtractorUpdater.exe'
        asset_name    = 'update-InvoiceExtractorUpdater.exe'
    }
)) {
    if (-not (Test-Path $payloadFile.source_path)) {
        throw "Missing release payload file $($payloadFile.source_path)"
    }

    $assetTarget = Join-Path $releaseAssetsDir $payloadFile.asset_name
    Copy-Item -Force $payloadFile.source_path $assetTarget

    $payloadSha256 = (Get-FileHash -Algorithm SHA256 $assetTarget).Hash.ToLowerInvariant()
    $null = $releaseFiles.Add([ordered]@{
        relative_path = $payloadFile.relative_path
        asset_name    = $payloadFile.asset_name
        sha256        = $payloadSha256
    })
}

$releaseTemplate = [ordered]@{
    version      = $version
    download_url = ''
    sha256       = $sha256
    notes        = ''
    published_at = ''
    files        = $releaseFiles
}
$releaseTemplate | ConvertTo-Json -Depth 5 | Set-Content -Path $releaseTemplatePath -Encoding UTF8

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
Write-Host "Release assets:$releaseAssetsDir"
Write-Host "Release JSON:  $releaseTemplatePath"
Write-Host "SHA-256:       $sha256"
