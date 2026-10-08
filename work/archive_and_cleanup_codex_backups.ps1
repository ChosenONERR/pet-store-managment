$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$manifestPath = 'D:\CodexData\migration_manifest.json'
$archiveRoot = 'D:\CodexData\source-backups\20260918-105345'
$archivePath = Join-Path $archiveRoot 'codex_pre_migration_20260918-105345.tar.gz'
$partialArchive = "$archivePath.partial"
$fileManifestPath = Join-Path $archiveRoot 'source_file_manifest.json'
$archiveRecordPath = Join-Path $archiveRoot 'archive_record.json'
$manifestBackupPath = Join-Path $archiveRoot 'migration_manifest_before_cleanup.json'
$allowedSourceRoot = [IO.Path]::GetFullPath($env:USERPROFILE).TrimEnd('\') + '\'
$expectedSuffix = '.pre_migration_20260918-105345'
$tarExe = Join-Path $env:SystemRoot 'System32\tar.exe'
$excludedCacheRegex = '(?i)\\edge-profile-[^\\]+\\'
$archiveExclusions = @('*/edge-profile-*')

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Migration manifest not found: $manifestPath" }
if (-not (Test-Path -LiteralPath $tarExe -PathType Leaf)) { throw "tar.exe not found: $tarExe" }

$migration = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$entries = @($migration.entries)
if ($entries.Count -ne 7) { throw "Expected 7 migration backup entries, found $($entries.Count)." }

$validated = @()
foreach ($entry in $entries) {
    $backupFull = [IO.Path]::GetFullPath([string]$entry.backup).TrimEnd('\')
    if (-not $backupFull.StartsWith($allowedSourceRoot,[StringComparison]::OrdinalIgnoreCase)) {
        throw "Backup is outside the allowed user root: $backupFull"
    }
    if (-not $backupFull.EndsWith($expectedSuffix,[StringComparison]::OrdinalIgnoreCase)) {
        throw "Backup does not have the expected migration suffix: $backupFull"
    }
    $item = Get-Item -LiteralPath $backupFull -Force -ErrorAction Stop
    if (-not $item.PSIsContainer -or $item.LinkType) { throw "Backup must be a normal directory: $backupFull" }

    $sourceItem = Get-Item -LiteralPath $entry.source -Force -ErrorAction Stop
    if ($sourceItem.LinkType -ne 'Junction') { throw "Active source is not a junction: $($entry.source)" }
    $actualTarget = [string]@($sourceItem.Target)[0]
    if ([IO.Path]::GetFullPath($actualTarget).TrimEnd('\') -ne [IO.Path]::GetFullPath([string]$entry.target).TrimEnd('\')) {
        throw "Active junction target mismatch: $($entry.source)"
    }
    $validated += [pscustomobject]@{name=[string]$entry.name;backup=$backupFull}
}

New-Item -ItemType Directory -Path $archiveRoot -Force | Out-Null
if (Test-Path -LiteralPath $archivePath) { throw "Final archive already exists; refusing to overwrite: $archivePath" }
if (Test-Path -LiteralPath $partialArchive) { Remove-Item -LiteralPath $partialArchive -Force }

Copy-Item -LiteralPath $manifestPath -Destination $manifestBackupPath -Force

Write-Host 'Building SHA-256 source manifest...' -ForegroundColor Cyan
$fileRows = @()
$totalBytes = [int64]0
foreach ($root in $validated) {
    $rootFiles = @(Get-ChildItem -LiteralPath $root.backup -Recurse -File -Force -ErrorAction Stop | Where-Object { $_.FullName -notmatch $excludedCacheRegex })
    Write-Host ("  {0}: {1} files" -f $root.name,$rootFiles.Count)
    foreach ($file in $rootFiles) {
        $relative = $file.FullName.Substring($allowedSourceRoot.Length).Replace('\','/')
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $fileRows += [pscustomobject]@{path=$relative;bytes=[int64]$file.Length;sha256=$hash}
        $totalBytes += [int64]$file.Length
    }
}
$sourceManifest = [ordered]@{
    version = 1
    created_at = [datetime]::Now.ToString('o')
    file_count = $fileRows.Count
    total_bytes = $totalBytes
    files = $fileRows
}
[IO.File]::WriteAllText($fileManifestPath,($sourceManifest | ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))

$relativeRoots = @($validated | ForEach-Object { $_.backup.Substring($allowedSourceRoot.Length) })
Write-Host ("Creating compressed archive for {0:N2} GB..." -f ($totalBytes / 1GB)) -ForegroundColor Cyan
& $tarExe -czf $partialArchive '--exclude=*/edge-profile-*' -C $env:USERPROFILE @relativeRoots
if ($LASTEXITCODE -ne 0) { throw "tar creation failed with exit code $LASTEXITCODE" }
Move-Item -LiteralPath $partialArchive -Destination $archivePath

Write-Host 'Testing the complete archive stream and file listing...' -ForegroundColor Cyan
$archiveListing = @(& $tarExe -tzf $archivePath)
if ($LASTEXITCODE -ne 0) { throw "tar verification failed with exit code $LASTEXITCODE" }
$listed = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($line in $archiveListing) {
    $normalized = ([string]$line).Replace('\','/').TrimEnd('/')
    if ($normalized) { $null = $listed.Add($normalized) }
}
$missingFromArchive = @($fileRows | Where-Object { -not $listed.Contains([string]$_.path) })
if ($missingFromArchive.Count) { throw "Archive listing is missing $($missingFromArchive.Count) files; first=$($missingFromArchive[0].path)" }

$archiveHash = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash
$archiveItem = Get-Item -LiteralPath $archivePath -Force
$record = [ordered]@{
    version = 1
    created_at = [datetime]::Now.ToString('o')
    archive = $archivePath
    archive_bytes = [int64]$archiveItem.Length
    archive_sha256 = $archiveHash
    source_manifest = $fileManifestPath
    source_file_count = $fileRows.Count
    source_total_bytes = $totalBytes
    tar_full_stream_test = 'passed'
    all_manifest_files_listed = $true
    excluded_recreatable_cache_patterns = $archiveExclusions
    source_backups = @($validated.backup)
}
[IO.File]::WriteAllText($archiveRecordPath,($record | ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))

Write-Host 'Archive verified. Removing only the seven validated C-drive backup directories...' -ForegroundColor Yellow
foreach ($root in $validated) {
    $current = Get-Item -LiteralPath $root.backup -Force -ErrorAction Stop
    $resolved = [IO.Path]::GetFullPath($current.FullName).TrimEnd('\')
    if ($resolved -ne $root.backup -or $current.LinkType -or -not $resolved.EndsWith($expectedSuffix,[StringComparison]::OrdinalIgnoreCase)) {
        throw "Backup changed after validation; refusing cleanup: $($root.backup)"
    }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
$remaining = @($validated | Where-Object { Test-Path -LiteralPath $_.backup })
if ($remaining.Count) { throw "Some source backups remain after cleanup: $($remaining.backup -join '; ')" }

$migration | Add-Member -NotePropertyName source_backups_archived -NotePropertyValue $true -Force
$migration | Add-Member -NotePropertyName source_archive -NotePropertyValue $archivePath -Force
$migration | Add-Member -NotePropertyName source_archive_sha256 -NotePropertyValue $archiveHash -Force
$migration | Add-Member -NotePropertyName source_archive_manifest -NotePropertyValue $fileManifestPath -Force
$migration | Add-Member -NotePropertyName source_backups_removed_at -NotePropertyValue ([datetime]::Now.ToString('o')) -Force
$migration.source_backups_preserved = $false
$tempManifest = "$manifestPath.$([guid]::NewGuid().ToString('N')).tmp"
[IO.File]::WriteAllText($tempManifest,($migration | ConvertTo-Json -Depth 9),[Text.UTF8Encoding]::new($false))
Move-Item -LiteralPath $tempManifest -Destination $manifestPath -Force

[ordered]@{
    ok = $true
    archive = $archivePath
    archive_gb = [math]::Round($archiveItem.Length / 1GB,3)
    archive_sha256 = $archiveHash
    source_gb_archived = [math]::Round($totalBytes / 1GB,3)
    source_gb_removed_from_c = [math]::Round((($migration.copy_results | Measure-Object bytes -Sum).Sum) / 1GB,3)
    excluded_recreatable_cache_patterns = $archiveExclusions
    source_backup_directories_removed = @($validated.backup)
    active_junctions_untouched = $true
} | ConvertTo-Json -Depth 6
