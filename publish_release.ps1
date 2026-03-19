param(
    [string]$Ref = 'main',
    [string]$Notes = '',
    [switch]$NoWait
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

$null = Get-Command gh -ErrorAction Stop

Write-Host "Checking GitHub CLI authentication..."
gh auth status
if ($LASTEXITCODE -ne 0) {
    throw "GitHub CLI is not authenticated. Run 'gh auth login' once on this machine."
}

Write-Host "Dispatching release workflow for ref '$Ref'..."
if ([string]::IsNullOrWhiteSpace($Notes)) {
    gh workflow run release.yml --ref $Ref
} else {
    gh workflow run release.yml --ref $Ref -f "notes=$Notes"
}
if ($LASTEXITCODE -ne 0) {
    throw "Failed to dispatch the release workflow."
}

if ($NoWait) {
    Write-Host "Release workflow submitted."
    exit 0
}

Start-Sleep -Seconds 5

$runId = $null
$runUrl = ''

for ($attempt = 1; $attempt -le 12 -and -not $runId; $attempt++) {
    $runs = gh run list `
        --workflow release.yml `
        --branch $Ref `
        --limit 5 `
        --json databaseId,createdAt,event,status,url | ConvertFrom-Json

    $run = $runs |
        Where-Object { $_.event -eq 'workflow_dispatch' } |
        Sort-Object createdAt -Descending |
        Select-Object -First 1

    if ($run) {
        $runId = [string]$run.databaseId
        $runUrl = [string]$run.url
        break
    }

    Start-Sleep -Seconds 5
}

if (-not $runId) {
    throw "The workflow was dispatched, but no matching run was found."
}

Write-Host "Watching GitHub Actions run $runId..."
gh run watch $runId --exit-status
if ($LASTEXITCODE -ne 0) {
    if ($runUrl) {
        throw "Release workflow failed. See $runUrl"
    }
    throw "Release workflow failed."
}

if ($runUrl) {
    Write-Host "Release completed successfully: $runUrl"
} else {
    Write-Host "Release completed successfully."
}
