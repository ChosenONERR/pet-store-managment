$ErrorActionPreference = 'Stop'

$migrationScript = Join-Path $PSScriptRoot 'migrate_codex_data_offline.ps1'
$logRoot = 'D:\CodexMigrationLogs'
$started = Get-Date -Format 'yyyyMMdd-HHmmss'
$logPath = Join-Path $logRoot "codex_verify_$started.log"

New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
Start-Transcript -LiteralPath $logPath -Force | Out-Null
try {
    Write-Host 'Codex migration verification only. No files will be moved or deleted.' -ForegroundColor Cyan
    Write-Host 'The Codex desktop app must be fully closed.' -ForegroundColor Yellow
    Write-Host "Log: $logPath"
    & $migrationScript -Action Verify
    Write-Host ''
    Write-Host 'Migration verification succeeded.' -ForegroundColor Green
    Write-Host 'You may now reopen Codex.' -ForegroundColor Green
}
catch {
    Write-Host ''
    Write-Host 'Migration verification failed.' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "Review the log: $logPath" -ForegroundColor Yellow
    Write-Host 'Do not delete or manually move any migration data.' -ForegroundColor Yellow
    exit 1
}
finally {
    Stop-Transcript | Out-Null
}

