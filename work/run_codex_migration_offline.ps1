$ErrorActionPreference = 'Stop'

$migrationScript = Join-Path $PSScriptRoot 'migrate_codex_data_offline.ps1'
$logRoot = 'D:\CodexMigrationLogs'
$started = Get-Date -Format 'yyyyMMdd-HHmmss'
$logPath = Join-Path $logRoot "codex_migration_$started.log"

if (-not (Test-Path -LiteralPath $migrationScript)) {
    throw "Migration script not found: $migrationScript"
}

New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
Start-Transcript -LiteralPath $logPath -Force | Out-Null

try {
    Write-Host ''
    Write-Host 'Codex offline migration: Plan -> Migrate -> Verify' -ForegroundColor Cyan
    Write-Host 'Do not continue unless the Codex desktop app is fully closed.' -ForegroundColor Yellow
    Write-Host "Log: $logPath"
    Write-Host ''

    & $migrationScript -Action Plan

    Write-Host ''
    $confirmation = Read-Host 'Type MIGRATE to start; type anything else to cancel'
    if ($confirmation -cne 'MIGRATE') {
        Write-Host 'Cancelled. No migration changes were made.' -ForegroundColor Yellow
        exit 2
    }

    Write-Host ''
    Write-Host 'Starting migration...' -ForegroundColor Cyan
    & $migrationScript -Action Migrate

    Write-Host ''
    Write-Host 'Migration finished. Keeping Codex closed and starting verification...' -ForegroundColor Cyan
    & $migrationScript -Action Verify

    Write-Host ''
    Write-Host 'Migration and offline verification succeeded.' -ForegroundColor Green
    Write-Host 'You may now open Codex and verify old and new conversations.' -ForegroundColor Green
    Write-Host 'Do not remove the C-drive .pre_migration backups yet.' -ForegroundColor Yellow
}
catch {
    Write-Host ''
    Write-Host 'Migration or verification failed.' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "Review the log: $logPath" -ForegroundColor Yellow
    Write-Host 'Keep Codex closed. Do not manually move or delete any migration files.' -ForegroundColor Yellow
    exit 1
}
finally {
    Stop-Transcript | Out-Null
}

