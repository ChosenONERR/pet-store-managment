param(
    [ValidateSet('Plan','Migrate','Verify','Rollback')]
    [string]$Action = 'Plan',
    [string]$TargetRoot = 'D:\CodexData'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

function Get-DirectoryFiles([string]$Root) {
    if (-not (Test-Path -LiteralPath $Root -PathType Container)) { return @() }
    # Ignore transient/broken file links inside browser and runtime caches. Robocopy
    # already excludes junction traversal, and these reparse points are not data files.
    return @(Get-ChildItem -LiteralPath $Root -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object { -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) })
}

function Get-DirectorySummary([string]$Root) {
    $files = @(Get-DirectoryFiles $Root)
    return [pscustomobject]@{
        files = $files.Count
        bytes = [double](($files | Measure-Object Length -Sum).Sum)
    }
}

function Assert-CodexStopped {
    $userRoot = [IO.Path]::GetFullPath($env:USERPROFILE).TrimEnd('\') + '\'
    $localRuntimeRoot = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'OpenAI\Codex')).TrimEnd('\') + '\'
    $codexDataMarkers = @(
        $userRoot + '.codex\',
        $userRoot + '.cache\codex-runtimes\',
        $localRuntimeRoot
    )
    $busy = @(Get-CimInstance Win32_Process | Where-Object {
        $executable = [string]$_.ExecutablePath
        $commandLine = [string]$_.CommandLine
        $_.ProcessId -ne $PID -and (
            $_.Name -match '^(Codex|codex|ChatGPT)(\.exe)?$' -or
            @($codexDataMarkers | Where-Object {
                ($executable -and $executable.StartsWith($_,[StringComparison]::OrdinalIgnoreCase)) -or
                ($commandLine -and $commandLine.IndexOf($_,[StringComparison]::OrdinalIgnoreCase) -ge 0)
            }).Count -gt 0
        )
    })
    if ($busy.Count -gt 0) {
        $details = ($busy | ForEach-Object { "$($_.Name) PID=$($_.ProcessId)" }) -join '; '
        throw "Codex is still running. Exit Codex before migration: $details"
    }
}

function Assert-BackupPreserved([string]$Backup, [string]$Target, [ValidateSet('AppendOrEqual','ExactFiles','PresenceOnly','MutableCache')][string]$Policy) {
    if ($Policy -eq 'MutableCache') {
        # The migration copy was already verified byte-for-byte before switching.
        # After Codex starts, cache package versions may be upgraded or evicted, so
        # requiring every old relative path to remain would report false data loss.
        $backupSummary = Get-DirectorySummary $Backup
        $targetSummary = Get-DirectorySummary $Target
        return [pscustomobject]@{
            files = $targetSummary.files
            policy = $Policy
            grown = 0
            hash_checked = 0
            baseline_files = $backupSummary.files
            target_files = $targetSummary.files
        }
    }

    $backupMap = Get-RelativeFileMap $Backup
    $targetMap = Get-RelativeFileMap $Target
    $missing = @($backupMap.Keys | Where-Object { -not $targetMap.ContainsKey($_) })
    if ($missing.Count) { throw "Backup files missing from target: count=$($missing.Count), first=$($missing[0])" }

    if ($Policy -eq 'PresenceOnly') {
        return [pscustomobject]@{files=$backupMap.Count;policy=$Policy;grown=0;hash_checked=0;baseline_files=$backupMap.Count;target_files=$targetMap.Count}
    }

    $grown = 0
    $hashChecked = 0
    foreach ($relative in $backupMap.Keys) {
        $backupFile = $backupMap[$relative]
        $targetFile = $targetMap[$relative]
        if ($Policy -eq 'AppendOrEqual' -and $targetFile.Length -gt $backupFile.Length) {
            $grown++
            continue
        }
        if ($targetFile.Length -ne $backupFile.Length) {
            throw "Unexpected size change: $relative, backup=$($backupFile.Length), target=$($targetFile.Length)"
        }
        $backupHash = (Get-FileHash -LiteralPath $backupFile.FullName -Algorithm SHA256).Hash
        $targetHash = (Get-FileHash -LiteralPath $targetFile.FullName -Algorithm SHA256).Hash
        $hashChecked++
        if ($backupHash -ne $targetHash) { throw "SHA-256 mismatch: $relative" }
    }
    return [pscustomobject]@{files=$backupMap.Count;policy=$Policy;grown=$grown;hash_checked=$hashChecked;baseline_files=$backupMap.Count;target_files=$targetMap.Count}
}

function Assert-PathInside([string]$Path, [string]$Root, [string]$Label) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    $fullRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    if (-not $fullPath.StartsWith($fullRoot,[StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label path is outside the allowed root: $fullPath"
    }
}

function Get-RelativeFileMap([string]$Root) {
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $map = @{}
    foreach ($file in @(Get-DirectoryFiles $rootFull)) {
        $relative = $file.FullName.Substring($rootFull.Length).TrimStart('\')
        $map[$relative] = $file
    }
    return $map
}

function Assert-DirectoriesEqual([string]$Source, [string]$Target, [switch]$Hash) {
    $sourceMap = Get-RelativeFileMap $Source
    $targetMap = Get-RelativeFileMap $Target
    $missing = @($sourceMap.Keys | Where-Object { -not $targetMap.ContainsKey($_) })
    $extra = @($targetMap.Keys | Where-Object { -not $sourceMap.ContainsKey($_) })
    $sizeMismatch = @($sourceMap.Keys | Where-Object {
        $targetMap.ContainsKey($_) -and $sourceMap[$_].Length -ne $targetMap[$_].Length
    })
    if ($missing.Count -or $extra.Count -or $sizeMismatch.Count) {
        throw "Directory manifests differ: missing=$($missing.Count), extra=$($extra.Count), size=$($sizeMismatch.Count)"
    }
    if ($Hash) {
        foreach ($relative in $sourceMap.Keys) {
            $sourceHash = (Get-FileHash -LiteralPath $sourceMap[$relative].FullName -Algorithm SHA256).Hash
            $targetHash = (Get-FileHash -LiteralPath $targetMap[$relative].FullName -Algorithm SHA256).Hash
            if ($sourceHash -ne $targetHash) { throw "SHA-256 mismatch: $relative" }
        }
    }
    return [pscustomobject]@{ files=$sourceMap.Count; bytes=[double](($sourceMap.Values | Measure-Object Length -Sum).Sum); hash_checked=[bool]$Hash }
}

function Remove-VerifiedJunction([string]$Path, [string]$ExpectedTarget) {
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if ($item.LinkType -ne 'Junction') { throw "Refusing to remove a non-junction path: $Path" }
    $actualTarget = [string]@($item.Target)[0]
    if ([IO.Path]::GetFullPath($actualTarget).TrimEnd('\') -ne [IO.Path]::GetFullPath($ExpectedTarget).TrimEnd('\')) {
        throw "Junction target mismatch: $Path -> $actualTarget"
    }
    [IO.Directory]::Delete($Path)
}

$mappings = @(
    [pscustomobject]@{ name='sessions'; source=(Join-Path $env:USERPROFILE '.codex\sessions'); target=(Join-Path $TargetRoot '.codex\sessions') },
    [pscustomobject]@{ name='generated_images'; source=(Join-Path $env:USERPROFILE '.codex\generated_images'); target=(Join-Path $TargetRoot '.codex\generated_images') },
    [pscustomobject]@{ name='visualizations'; source=(Join-Path $env:USERPROFILE '.codex\visualizations'); target=(Join-Path $TargetRoot '.codex\visualizations') },
    [pscustomobject]@{ name='sqlite'; source=(Join-Path $env:USERPROFILE '.codex\sqlite'); target=(Join-Path $TargetRoot '.codex\sqlite') },
    [pscustomobject]@{ name='sandbox_bin'; source=(Join-Path $env:USERPROFILE '.codex\.sandbox-bin'); target=(Join-Path $TargetRoot '.codex\sandbox-bin') },
    [pscustomobject]@{ name='plugin_cache'; source=(Join-Path $env:USERPROFILE '.codex\plugins\cache'); target=(Join-Path $TargetRoot '.codex\plugin-cache') },
    [pscustomobject]@{ name='runtime_cache'; source=(Join-Path $env:USERPROFILE '.cache\codex-runtimes'); target=(Join-Path $TargetRoot 'cache\codex-runtimes') },
    [pscustomobject]@{ name='web_data'; source=(Join-Path $env:APPDATA 'Codex\web'); target=(Join-Path $TargetRoot 'appdata\roaming\web') },
    [pscustomobject]@{ name='logs'; source=(Join-Path $env:LOCALAPPDATA 'Codex\Logs'); target=(Join-Path $TargetRoot 'appdata\local\Logs') }
)

$targetFull = [IO.Path]::GetFullPath($TargetRoot).TrimEnd('\')
if ([IO.Path]::GetPathRoot($targetFull) -ne 'D:\') { throw 'This script only allows a target on drive D.' }
foreach ($mapping in $mappings) {
    Assert-PathInside $mapping.source $env:USERPROFILE 'Source'
    Assert-PathInside $mapping.target $targetFull 'Target'
}

$plan = @($mappings | ForEach-Object {
    $item = Get-Item -LiteralPath $_.source -Force -ErrorAction SilentlyContinue
    $summary = Get-DirectorySummary $_.source
    [pscustomobject]@{
        name = $_.name
        source = $_.source
        target = $_.target
        exists = [bool]$item
        link_type = if ($item) { [string]$item.LinkType } else { '' }
        files = $summary.files
        bytes = $summary.bytes
        gb = [math]::Round($summary.bytes / 1GB, 3)
    }
})

if ($Action -eq 'Plan') {
    $drive = [IO.DriveInfo]::new('D')
    [ordered]@{
        ok = $true
        action = 'Plan'
        target_root = $targetFull
        total_gb = [math]::Round((($plan | Measure-Object bytes -Sum).Sum) / 1GB, 3)
        d_free_gb = [math]::Round($drive.AvailableFreeSpace / 1GB, 3)
        close_codex_before_migrate = $true
        keeps_source_backups = $true
        mappings = $plan
    } | ConvertTo-Json -Depth 6
    exit 0
}

$manifestPath = Join-Path $targetFull 'migration_manifest.json'

if ($Action -eq 'Rollback') {
    Assert-CodexStopped
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Migration manifest not found: $manifestPath" }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($manifest.source_backups_archived -eq $true) {
        throw "Source backups are archived at $($manifest.source_archive). Restore and verify the archive before rollback."
    }
    foreach ($entry in @($manifest.entries) | Sort-Object sequence -Descending) {
        if (Test-Path -LiteralPath $entry.source) {
            Remove-VerifiedJunction $entry.source $entry.target
        }
        if (-not (Test-Path -LiteralPath $entry.backup -PathType Container)) { throw "Source backup not found: $($entry.backup)" }
        Move-Item -LiteralPath $entry.backup -Destination $entry.source
    }
    [ordered]@{ok=$true;action='Rollback';restored=@($manifest.entries.name);target_data_preserved=$true} | ConvertTo-Json -Depth 5
    exit 0
}

if ($Action -eq 'Verify') {
    Assert-CodexStopped
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Migration manifest not found: $manifestPath" }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $archiveVerified = $false
    if ($manifest.source_backups_archived -eq $true) {
        if (-not (Test-Path -LiteralPath $manifest.source_archive -PathType Leaf)) { throw "Source backup archive not found: $($manifest.source_archive)" }
        $currentArchiveHash = (Get-FileHash -LiteralPath $manifest.source_archive -Algorithm SHA256).Hash
        if ($currentArchiveHash -ne $manifest.source_archive_sha256) { throw "Source backup archive SHA-256 mismatch: $($manifest.source_archive)" }
        $archiveVerified = $true
    }
    $verified = @()
    foreach ($entry in @($manifest.entries)) {
        $item = Get-Item -LiteralPath $entry.source -Force -ErrorAction Stop
        if ($item.LinkType -ne 'Junction') { throw "Path is not a junction: $($entry.source)" }
        $actualTarget = [string]@($item.Target)[0]
        if ([IO.Path]::GetFullPath($actualTarget).TrimEnd('\') -ne [IO.Path]::GetFullPath($entry.target).TrimEnd('\')) {
            throw "Junction target is incorrect: $($entry.source)"
        }
        if (-not (Test-Path -LiteralPath $entry.target -PathType Container)) { throw "Target directory not found: $($entry.target)" }
        $copyResult = @($manifest.copy_results | Where-Object name -eq $entry.name)
        if ($copyResult.Count -ne 1) { throw "Pre-switch copy verification record missing: $($entry.name)" }

        if (Test-Path -LiteralPath $entry.backup -PathType Container) {
            $policy = switch ($entry.name) {
                'sessions' { 'AppendOrEqual' }
                'generated_images' { 'ExactFiles' }
                'sqlite' { 'PresenceOnly' }
                'sandbox_bin' { 'MutableCache' }
                'plugin_cache' { 'MutableCache' }
                'runtime_cache' { 'MutableCache' }
                default { 'PresenceOnly' }
            }
            $check = Assert-BackupPreserved $entry.backup $entry.target $policy
        } elseif ($archiveVerified) {
            $targetSummary = Get-DirectorySummary $entry.target
            $check = [pscustomobject]@{files=$targetSummary.files;policy='ArchivedBaseline';grown=0;hash_checked=0;baseline_files=$copyResult[0].files;target_files=$targetSummary.files}
        } else {
            throw "Source backup not found and no verified archive is available: $($entry.backup)"
        }
        $verified += [pscustomobject]@{
            name=$entry.name
            files_at_migration=$copyResult[0].files
            bytes_at_migration=$copyResult[0].bytes
            pre_switch_sha256_verified=$true
            current_policy=$check.policy
            current_files_checked=$check.files
            current_hashes_checked=$check.hash_checked
            grown_files=$check.grown
            baseline_files=$check.baseline_files
            target_files=$check.target_files
            target=$entry.target
        }
    }
    $sqliteEntry = @($manifest.entries | Where-Object name -eq 'sqlite')
    $sqliteQuickCheck = 'not_available'
    if ($sqliteEntry.Count -eq 1) {
        $sqliteExe = Get-Command sqlite3.exe -ErrorAction SilentlyContinue
        $dbPath = Join-Path $sqliteEntry[0].target 'codex-dev.db'
        if ($sqliteExe -and (Test-Path -LiteralPath $dbPath -PathType Leaf)) {
            $sqliteQuickCheck = [string](& $sqliteExe.Source $dbPath 'PRAGMA quick_check;')
            if ($sqliteQuickCheck.Trim() -ne 'ok') { throw "SQLite quick_check failed: $sqliteQuickCheck" }
        }
    }
    [ordered]@{ok=$true;action='Verify';verification_model='pre-switch exact SHA-256 plus post-switch preservation audit';sqlite_quick_check=$sqliteQuickCheck;verified=$verified;source_backups_preserved=[bool]$manifest.source_backups_preserved;source_backups_archived=[bool]$manifest.source_backups_archived;source_archive_verified=$archiveVerified} | ConvertTo-Json -Depth 7
    exit 0
}

Assert-CodexStopped
$existing = @($plan | Where-Object { $_.exists })
$requiredBytes = [double](($existing | Measure-Object bytes -Sum).Sum)
$targetDrive = [IO.DriveInfo]::new('D')
if (-not $targetDrive.IsReady -or $targetDrive.AvailableFreeSpace -lt ($requiredBytes + 2GB)) {
    throw "Not enough free space on D:. Data=$([math]::Round($requiredBytes/1GB,2))GB; reserve=2GB."
}

New-Item -ItemType Directory -Path $targetFull -Force | Out-Null
$copyResults = @()
foreach ($mapping in $mappings) {
    if (-not (Test-Path -LiteralPath $mapping.source -PathType Container)) { continue }
    $sourceItem = Get-Item -LiteralPath $mapping.source -Force
    if ($sourceItem.LinkType) { throw "Source is already a link; refusing duplicate migration: $($mapping.source)" }
    New-Item -ItemType Directory -Path $mapping.target -Force | Out-Null
    $logPath = Join-Path $targetFull ("copy_{0}.log" -f $mapping.name)
    & robocopy $mapping.source $mapping.target /E /COPY:DAT /DCOPY:DAT /R:2 /W:1 /XJ /NFL /NDL /NP /LOG:$logPath | Out-Null
    $robocopyCode = $LASTEXITCODE
    if ($robocopyCode -ge 8) { throw "Copy failed: $($mapping.name), robocopy=$robocopyCode" }
    $check = Assert-DirectoriesEqual $mapping.source $mapping.target -Hash
    $copyResults += [pscustomobject]@{name=$mapping.name;files=$check.files;bytes=$check.bytes;robocopy_code=$robocopyCode;log=$logPath}
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$switched = @()
try {
    $sequence = 0
    foreach ($mapping in $mappings) {
        if (-not (Test-Path -LiteralPath $mapping.source -PathType Container)) { continue }
        $sequence++
        $backup = "$($mapping.source).pre_migration_$stamp"
        if (Test-Path -LiteralPath $backup) { throw "Backup path already exists: $backup" }
        Move-Item -LiteralPath $mapping.source -Destination $backup
        try {
            $null = New-Item -ItemType Junction -Path $mapping.source -Target $mapping.target -ErrorAction Stop
        } catch {
            Move-Item -LiteralPath $backup -Destination $mapping.source
            throw
        }
        $switched += [pscustomobject]@{sequence=$sequence;name=$mapping.name;source=$mapping.source;target=$mapping.target;backup=$backup}
    }
} catch {
    foreach ($entry in @($switched) | Sort-Object sequence -Descending) {
        if (Test-Path -LiteralPath $entry.source) { Remove-VerifiedJunction $entry.source $entry.target }
        if (Test-Path -LiteralPath $entry.backup -PathType Container) { Move-Item -LiteralPath $entry.backup -Destination $entry.source }
    }
    throw
}

$manifest = [ordered]@{
    version = 1
    migrated_at = [datetime]::Now.ToString('o')
    target_root = $targetFull
    entries = $switched
    copy_results = $copyResults
    source_backups_preserved = $true
}
$tempManifest = "$manifestPath.$([guid]::NewGuid().ToString('N')).tmp"
[IO.File]::WriteAllText($tempManifest,($manifest | ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
Move-Item -LiteralPath $tempManifest -Destination $manifestPath -Force

[ordered]@{
    ok = $true
    action = 'Migrate'
    target_root = $targetFull
    migrated = @($switched.name)
    source_backups_preserved = $true
    manifest = $manifestPath
    next_step = 'Keep Codex closed and run -Action Verify first. Reopen Codex only after verification succeeds.'
} | ConvertTo-Json -Depth 6
