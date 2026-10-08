$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$manifestPath = 'D:\CodexData\migration_manifest.json'
$archiveRecordPath = 'D:\CodexData\source-backups\20260918-105345\archive_record.json'
$expectedSuffix = '.pre_migration_20260918-105345'
$allowedSourceRoot = [IO.Path]::GetFullPath($env:USERPROFILE).TrimEnd('\') + '\'

function Remove-ExtendedPathDirectory([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    $extended = if ($full.StartsWith('\\')) { '\\?\UNC\' + $full.TrimStart('\') } else { '\\?\' + $full }
    [IO.Directory]::Delete($extended,$true)
}

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Migration manifest not found: $manifestPath" }
if (-not (Test-Path -LiteralPath $archiveRecordPath -PathType Leaf)) { throw "Archive record not found: $archiveRecordPath" }
$migration = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$record = Get-Content -LiteralPath $archiveRecordPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not (Test-Path -LiteralPath $record.archive -PathType Leaf)) { throw "Archive not found: $($record.archive)" }

Write-Host 'Verifying archive SHA-256 before cleanup...' -ForegroundColor Cyan
$archiveHash = (Get-FileHash -LiteralPath $record.archive -Algorithm SHA256).Hash
if ($archiveHash -ne $record.archive_sha256) { throw "Archive SHA-256 mismatch: $($record.archive)" }

$validated = @()
foreach ($entry in @($migration.entries)) {
    $sourceItem = Get-Item -LiteralPath $entry.source -Force -ErrorAction Stop
    if ($sourceItem.LinkType -ne 'Junction') { throw "Active source is not a junction: $($entry.source)" }
    $actualTarget = [string]@($sourceItem.Target)[0]
    if ([IO.Path]::GetFullPath($actualTarget).TrimEnd('\') -ne [IO.Path]::GetFullPath([string]$entry.target).TrimEnd('\')) {
        throw "Active junction target mismatch: $($entry.source)"
    }

    $backupFull = [IO.Path]::GetFullPath([string]$entry.backup).TrimEnd('\')
    if (-not $backupFull.StartsWith($allowedSourceRoot,[StringComparison]::OrdinalIgnoreCase)) {
        throw "Backup is outside the allowed user root: $backupFull"
    }
    if (-not $backupFull.EndsWith($expectedSuffix,[StringComparison]::OrdinalIgnoreCase)) {
        throw "Backup suffix mismatch: $backupFull"
    }
    if (Test-Path -LiteralPath $backupFull -PathType Container) {
        $item = Get-Item -LiteralPath $backupFull -Force
        if ($item.LinkType) { throw "Refusing to remove a linked backup path: $backupFull" }
        $validated += $backupFull
    }
}

Write-Host ("Removing {0} remaining validated backup directories with extended-path support..." -f $validated.Count) -ForegroundColor Yellow
foreach ($backup in $validated) {
    Write-Host "  $backup"
    Remove-ExtendedPathDirectory $backup
}

$remaining = @($migration.entries | Where-Object { Test-Path -LiteralPath $_.backup })
if ($remaining.Count) { throw "Backup cleanup incomplete: $($remaining.backup -join '; ')" }

$migration | Add-Member -NotePropertyName source_backups_archived -NotePropertyValue $true -Force
$migration | Add-Member -NotePropertyName source_archive -NotePropertyValue ([string]$record.archive) -Force
$migration | Add-Member -NotePropertyName source_archive_sha256 -NotePropertyValue ([string]$record.archive_sha256) -Force
$migration | Add-Member -NotePropertyName source_archive_manifest -NotePropertyValue ([string]$record.source_manifest) -Force
$migration | Add-Member -NotePropertyName source_backups_removed_at -NotePropertyValue ([datetime]::Now.ToString('o')) -Force
$migration.source_backups_preserved = $false
$tempManifest = "$manifestPath.$([guid]::NewGuid().ToString('N')).tmp"
[IO.File]::WriteAllText($tempManifest,($migration | ConvertTo-Json -Depth 9),[Text.UTF8Encoding]::new($false))
Move-Item -LiteralPath $tempManifest -Destination $manifestPath -Force

[ordered]@{
    ok = $true
    archive = [string]$record.archive
    archive_gb = [math]::Round((Get-Item -LiteralPath $record.archive).Length / 1GB,3)
    archive_sha256 = [string]$record.archive_sha256
    remaining_backups_removed = $validated
    all_c_drive_source_backups_removed = $true
    active_junctions_untouched = $true
} | ConvertTo-Json -Depth 6

