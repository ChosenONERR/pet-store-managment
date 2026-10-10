$ErrorActionPreference = 'Stop'

function Get-FinanceTransactionDirectory([string]$StateRoot) {
    $path = Join-Path $StateRoot 'transactions'
    if (-not (Test-Path -LiteralPath $path)) { $null = New-Item -ItemType Directory -Path $path }
    return $path
}

function Get-FinanceTransactionPath([string]$StateRoot, [string]$TransactionId) {
    $safeId = [regex]::Replace($TransactionId, '[^A-Za-z0-9_.-]', '_')
    return (Join-Path (Get-FinanceTransactionDirectory $StateRoot) "$safeId.json")
}

function Invoke-WithFinanceJournalLock([scriptblock]$Operation) {
    $mutex = [Threading.Mutex]::new($false, 'Local\PetstarFinanceTransactionJournalV2')
    $taken = $false
    try {
        try { $taken = $mutex.WaitOne([timespan]::FromSeconds(30)) }
        catch [Threading.AbandonedMutexException] { $taken = $true }
        if (-not $taken) { throw '经营事务日志锁等待超时' }
        return (& $Operation)
    } finally {
        if ($taken) { try { $mutex.ReleaseMutex() } catch {} }
        $mutex.Dispose()
    }
}

function Read-FinanceTransactionJournal([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    return ([System.IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json)
}

function Get-FinanceTransactionStatusFast([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $stream = [IO.File]::OpenRead($Path)
    try {
        $length = [math]::Min(65536, [int64]$stream.Length)
        $buffer = New-Object byte[] ([int]$length)
        $read = $stream.Read($buffer, 0, $buffer.Length)
        $prefix = [Text.Encoding]::UTF8.GetString($buffer, 0, $read)
        if ($prefix -match '"status"\s*:\s*"([^"]+)"') { return [string]$Matches[1] }
        return $null
    } finally {
        $stream.Dispose()
    }
}

function Write-FinanceJournalAtomically([string]$Path, [string]$Content) {
    $temp = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [System.IO.File]::WriteAllText($temp, $Content, [Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $temp -Destination $Path -Force
    } finally {
        # 磁盘写满或进程异常时，不能把不完整的临时日志永久留在状态目录。
        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
    }
}

function Write-FinanceTransactionJournal([string]$Path, $Journal) {
    Invoke-WithFinanceJournalLock {
        $json = $Journal | ConvertTo-Json -Depth 50
        Write-FinanceJournalAtomically $Path ($json + [Environment]::NewLine)
    } | Out-Null
}

function Update-FinanceTransactionJournal([string]$Path, [scriptblock]$Update) {
    return (Invoke-WithFinanceJournalLock {
        $journal = Read-FinanceTransactionJournal $Path
        if ($null -eq $journal) { throw "找不到经营事务日志：$(Split-Path $Path -Leaf)" }
        $updated = & $Update $journal
        if ($null -eq $updated) { $updated = $journal }
        Write-FinanceJournalAtomically $Path (($updated | ConvertTo-Json -Depth 50) + [Environment]::NewLine)
        return $updated
    })
}

function ConvertTo-CompactedFinanceTransactionJournal($Journal, [string]$FinalStatus, [string]$ErrorMessage) {
    if ($FinalStatus -notin @('committed','rolled_back')) { throw "事务终态不可压缩：$FinalStatus" }
    $entries = @($Journal.entries)
    $manifest = [Collections.Generic.List[object]]::new()
    foreach ($entry in $entries) {
        $rows = @($entry.cells)
        $columns = if ($rows.Count -gt 0) { @($rows[0]).Count } else { 0 }
        $manifest.Add([ordered]@{
            sequence = $entry.sequence
            workbook = [string]$entry.workbook
            sheet = [string]$entry.sheet
            range = [string]$entry.range
            row_count = $rows.Count
            column_count = $columns
            captured_at = $entry.captured_at
        })
    }
    $result = [ordered]@{}
    foreach ($property in $Journal.PSObject.Properties) {
        if ($property.Name -in @('entries','entry_count','entry_manifest','committed_at','rolled_back_at')) { continue }
        $result[$property.Name] = $property.Value
    }
    $now = [datetime]::Now.ToString('o')
    $result['status'] = $FinalStatus
    $result['updated_at'] = $now
    if ($FinalStatus -eq 'committed') { $result['committed_at'] = $now }
    if ($FinalStatus -eq 'rolled_back') { $result['rolled_back_at'] = $now }
    if ($ErrorMessage) { $result['error'] = $ErrorMessage }
    $result['entry_count'] = $entries.Count
    $result['entry_manifest'] = $manifest.ToArray()
    # 终态事务禁止再次回滚；只保留审计范围，不再保留完整单元格快照。
    return [pscustomobject]$result
}

function ConvertTo-RestorableCell($Cell) {
    $result = [ordered]@{}
    if ($null -ne $Cell) {
        foreach ($name in @('formula','value','cell_styles','note','data_validation','rich_text')) {
            $property = $Cell.PSObject.Properties[$name]
            if ($null -ne $property -and $null -ne $property.Value) { $result[$name] = $property.Value }
        }
    }
    if ($result.Contains('formula')) { $null = $result.Remove('value') }
    elseif ($result.Contains('value') -and $result['value'] -is [string] -and [string]$result['value'] -ne '') {
        # Feishu cells-get renders many numeric cells as strings even when their
        # number format is numeric. Writing those strings back during rollback
        # makes SUMIFS ignore them. Restore the storage type from number_format;
        # IDs/text columns use General/@ and therefore remain strings.
        $numberFormat = $null
        if ($result.Contains('cell_styles') -and $null -ne $result['cell_styles']) {
            $formatProperty = $result['cell_styles'].PSObject.Properties['number_format']
            if ($null -ne $formatProperty) { $numberFormat = [string]$formatProperty.Value }
        }
        if ($numberFormat -and $numberFormat -notin @('General','@') -and $numberFormat -match '[#0ymd]') {
            $numericValue = 0.0
            $textValue = [string]$result['value']
            if ([double]::TryParse($textValue.Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$numericValue)) {
                $result['value'] = $numericValue
            } elseif ($numberFormat -match '(?i)y|m|d') {
                $dateValue = [datetime]::MinValue
                if ([datetime]::TryParse($textValue, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$dateValue)) {
                    $result['value'] = ($dateValue.Date - [datetime]'1899-12-30').TotalDays
                }
            }
        }
    }
    if (-not $result.Contains('formula') -and -not $result.Contains('value') -and -not $result.Contains('rich_text')) {
        $result['value'] = ''
    }
    return $result
}

function ConvertFrom-ColumnLetters([string]$Letters) {
    $value = 0
    foreach ($character in $Letters.ToUpperInvariant().ToCharArray()) { $value = ($value * 26) + ([int][char]$character - [int][char]'A' + 1) }
    return $value
}

function ConvertTo-AbsoluteColumnIndex($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0
    if ([int]::TryParse([string]$Value, [ref]$number)) { return $number }
    $letters = ([string]$Value).Trim()
    if ($letters -match '^[A-Za-z]+$') { return ConvertFrom-ColumnLetters $letters }
    throw "无法识别飞书列坐标：$Value"
}

function ConvertTo-RestorableMatrix($RangeResult, [string]$RequestedRange) {
    if ($RequestedRange -notmatch '^([A-Z]+)(\d+):([A-Z]+)(\d+)$') { throw "事务快照范围格式不受支持：$RequestedRange" }
    $startColumn = ConvertFrom-ColumnLetters $Matches[1]
    $startRow = [int]$Matches[2]
    $endColumn = ConvertFrom-ColumnLetters $Matches[3]
    $endRow = [int]$Matches[4]
    $rowCount = $endRow - $startRow + 1
    $columnCount = $endColumn - $startColumn + 1
    $matrix = [Collections.Generic.List[object]]::new()
    for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
        $row = [Collections.Generic.List[object]]::new()
        for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) { $row.Add([ordered]@{value=''}) }
        $matrix.Add($row.ToArray())
    }
    $returnedRows = @($RangeResult.row_indices)
    $returnedColumns = @($RangeResult.col_indices)
    # Different PowerShell/ConvertFrom-Json versions may return a one-row range as
    # either [cell, cell] or [[cell, cell]]. Normalize both shapes without adding
    # an extra array level, otherwise a real one-row before-image becomes blank.
    $rawCells = @($RangeResult.cells)
    if ($returnedRows.Count -eq 1) {
        if ($rawCells.Count -eq 1 -and $rawCells[0] -is [System.Array]) {
            $sourceRows = @($rawCells)
        } else {
            $sourceRows = @(,$rawCells)
        }
    } elseif ($returnedColumns.Count -eq 1) {
        $normalizedRows = [Collections.Generic.List[object]]::new()
        foreach ($item in $rawCells) {
            if ($item -is [System.Array]) { $normalizedRows.Add($item) }
            else { $normalizedRows.Add(@($item)) }
        }
        $sourceRows = $normalizedRows.ToArray()
    } else {
        $sourceRows = $rawCells
    }
    for ($sourceRowIndex = 0; $sourceRowIndex -lt $sourceRows.Count; $sourceRowIndex++) {
        $absoluteRow = if ($returnedRows.Count -eq $sourceRows.Count -and $null -ne $returnedRows[$sourceRowIndex]) { [int]$returnedRows[$sourceRowIndex] } else { $startRow + $sourceRowIndex }
        $targetRow = $absoluteRow - $startRow
        if ($targetRow -lt 0 -or $targetRow -ge $rowCount) { continue }
        $sourceCells = @($sourceRows[$sourceRowIndex])
        for ($sourceColumnIndex = 0; $sourceColumnIndex -lt $sourceCells.Count; $sourceColumnIndex++) {
            # 飞书真实接口返回 A/B/C，测试桩和部分旧版本也可能返回 1/2/3。
            $absoluteColumn = if ($returnedColumns.Count -eq $sourceCells.Count -and $null -ne $returnedColumns[$sourceColumnIndex]) { ConvertTo-AbsoluteColumnIndex $returnedColumns[$sourceColumnIndex] } else { $startColumn + $sourceColumnIndex }
            $targetColumn = $absoluteColumn - $startColumn
            if ($targetColumn -lt 0 -or $targetColumn -ge $columnCount) { continue }
            $matrix[$targetRow][$targetColumn] = ConvertTo-RestorableCell $sourceCells[$sourceColumnIndex]
        }
    }
    # 禁止 PowerShell 管道把单行二维数组再次摊平成一维单元格数组。
    return ,$matrix.ToArray()
}

function Get-FinanceRollbackChunks($Entry, [int]$MaximumRows = 100) {
    $range = [string]$Entry.range
    if ($range -notmatch '^([A-Z]+)(\d+):([A-Z]+)(\d+)$') { throw "回滚范围格式不受支持：$range" }
    $startLetters = $Matches[1]
    $startRow = [int]$Matches[2]
    $endLetters = $Matches[3]
    $endRow = [int]$Matches[4]
    $rows = @($Entry.cells)
    $expectedRows = $endRow - $startRow + 1
    if ($rows.Count -ne $expectedRows) { throw "回滚快照行数不匹配：$range，期望$expectedRows，实际$($rows.Count)" }
    $chunks = [Collections.Generic.List[object]]::new()
    for ($offset = 0; $offset -lt $rows.Count; $offset += $MaximumRows) {
        $count = [math]::Min($MaximumRows, $rows.Count - $offset)
        $chunkRows = [Collections.Generic.List[object]]::new()
        for ($index = 0; $index -lt $count; $index++) { $chunkRows.Add($rows[$offset + $index]) }
        $chunkStartRow = $startRow + $offset
        $chunkEndRow = $chunkStartRow + $count - 1
        $chunks.Add([pscustomobject]@{
            range = "$startLetters$($chunkStartRow):$endLetters$($chunkEndRow)"
            cells = $chunkRows.ToArray()
        })
    }
    return $chunks.ToArray()
}

function New-FinanceTransaction(
    [string]$StateRoot,
    [string]$Description,
    [string]$TransactionId,
    [switch]$AcquireLock,
    [hashtable]$WorkbookTokens,
    [string]$LarkCli,
    [int]$LockTimeoutSeconds = 600
) {
    if (-not $TransactionId) { $TransactionId = 'TX-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant() }
    $context = [pscustomobject]@{
        id = $TransactionId
        state_root = $StateRoot
        path = Get-FinanceTransactionPath $StateRoot $TransactionId
        owns_lock = $false
        mutex = $null
    }
    if ($AcquireLock) {
        $mutex = [Threading.Mutex]::new($false, 'Local\PetstarFinanceMutationV2')
        $taken = $false
        $effectiveLockTimeoutSeconds = [math]::Max(0, $LockTimeoutSeconds)
        try { $taken = $mutex.WaitOne([timespan]::FromSeconds($effectiveLockTimeoutSeconds)) }
        catch [Threading.AbandonedMutexException] { $taken = $true }
        if (-not $taken) { $mutex.Dispose(); throw '另一个经营数据事务仍在执行，本次未写入' }
        $context.owns_lock = $true
        $context.mutex = $mutex
        if ($null -ne $WorkbookTokens -and $LarkCli) {
            $repair = Repair-PendingFinanceTransactions $StateRoot $WorkbookTokens $LarkCli
            if (-not $repair.ok) {
                try { $mutex.ReleaseMutex() } catch {}
                $mutex.Dispose()
                throw "存在未恢复的经营事务，已阻止新的写入：$($repair.error)"
            }
        }
    }
    $journal = [ordered]@{
        transaction_id = $TransactionId
        description = $Description
        status = 'active'
        started_at = [datetime]::Now.ToString('o')
        updated_at = [datetime]::Now.ToString('o')
        entries = @()
        processed_event_ids = @()
        processed_events = @()
        error = $null
        rollback_error = $null
    }
    Write-FinanceTransactionJournal $context.path $journal
    $env:PETSTAR_FINANCE_TRANSACTION_ID = $TransactionId
    $env:PETSTAR_FINANCE_TRANSACTION_STATE_ROOT = $StateRoot
    return $context
}

function Join-FinanceTransaction([string]$StateRoot) {
    $transactionId = [string]$env:PETSTAR_FINANCE_TRANSACTION_ID
    if (-not $transactionId) { return $null }
    $effectiveStateRoot = if ($env:PETSTAR_FINANCE_TRANSACTION_STATE_ROOT) { [string]$env:PETSTAR_FINANCE_TRANSACTION_STATE_ROOT } else { $StateRoot }
    $path = Get-FinanceTransactionPath $effectiveStateRoot $transactionId
    $journal = Read-FinanceTransactionJournal $path
    if ($null -eq $journal -or [string]$journal.status -ne 'active') { throw "经营事务不可加入：$transactionId" }
    return [pscustomobject]@{id=$transactionId;state_root=$effectiveStateRoot;path=$path;owns_lock=$false;mutex=$null}
}

function Register-FinanceRangeBeforeImage(
    $Context,
    [string]$WorkbookKey,
    [string]$SpreadsheetToken,
    [string]$SheetName,
    [string]$Range,
    [string]$LarkCli
) {
    if ($null -eq $Context) { return }
    $current = Read-FinanceTransactionJournal $Context.path
    if ($null -eq $current -or [string]$current.status -ne 'active') { throw "经营事务已不在写入状态：$($Context.id)" }
    $alreadyCaptured = @($current.entries | Where-Object {
        [string]$_.workbook -eq $WorkbookKey -and [string]$_.sheet -eq $SheetName -and [string]$_.range -eq $Range
    }).Count -gt 0
    if ($alreadyCaptured) { return }
    $raw = & $LarkCli sheets +cells-get --as user --spreadsheet-token $SpreadsheetToken --sheet-name $SheetName --range $Range --include 'value,formula,style,comment,data_validation' --max-chars 5000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "事务写前快照失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "事务写前快照不完整：$SheetName $Range" }
    $matrix = ConvertTo-RestorableMatrix $result.data.ranges[0] $Range
    $null = Update-FinanceTransactionJournal $Context.path {
        param($journal)
        if ([string]$journal.status -ne 'active') { throw "经营事务已不在写入状态：$($Context.id)" }
        $entries = @($journal.entries)
        $entry = [ordered]@{
            sequence = $entries.Count + 1
            workbook = $WorkbookKey
            sheet = $SheetName
            range = $Range
            cells = $matrix
            captured_at = [datetime]::Now.ToString('o')
        }
        $journal.entries = @($entries + $entry)
        $journal.updated_at = [datetime]::Now.ToString('o')
        return $journal
    }
}

function Register-FinanceProcessedEvent($Context, [string]$EventId, [string]$EventAction, [string]$Reference) {
    if ($null -eq $Context -or -not $EventId) { return }
    $null = Update-FinanceTransactionJournal $Context.path {
        param($journal)
        $journal.processed_event_ids = @(@($journal.processed_event_ids) + $EventId | Select-Object -Unique)
        $records = @($journal.processed_events | Where-Object { $_ -and [string]$_.event_id -ne $EventId })
        $records += [ordered]@{
            event_id = $EventId
            action = $EventAction
            reference = $Reference
            processed_at = [datetime]::Now.ToString('o')
        }
        $journal | Add-Member -NotePropertyName processed_events -NotePropertyValue @($records) -Force
        $journal.updated_at = [datetime]::Now.ToString('o')
        return $journal
    }
}

function Test-FinanceCommittedEvent([string]$StateRoot, [string]$EventId) {
    if (-not $EventId) { return $false }
    # Standalone business events use a deterministic transaction id. Reading that
    # one journal is a safe recovery fallback when the compact processed-events
    # log was not appended. Never parse every transaction snapshot here: sync
    # journals contain full workbook backups and made each preview CPU-bound.
    $path = Get-FinanceTransactionPath $StateRoot ("event-$EventId")
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    try {
        $journal = Read-FinanceTransactionJournal $path
        return ([string]$journal.status -eq 'committed' -and $EventId -in @($journal.processed_event_ids))
    } catch {}
    return $false
}

function Set-FinanceTransactionStatus($Context, [string]$Status, [string]$ErrorMessage) {
    if ($null -eq $Context) { return }
    $null = Update-FinanceTransactionJournal $Context.path {
        param($journal)
        $journal.status = $Status
        $journal.updated_at = [datetime]::Now.ToString('o')
        if ($Status -eq 'committed') { $journal | Add-Member -NotePropertyName committed_at -NotePropertyValue ([datetime]::Now.ToString('o')) -Force }
        if ($Status -eq 'rolled_back') { $journal | Add-Member -NotePropertyName rolled_back_at -NotePropertyValue ([datetime]::Now.ToString('o')) -Force }
        if ($ErrorMessage) { $journal.error = $ErrorMessage }
        return $journal
    }
}

function Invoke-FinanceTransactionRollback(
    $Context,
    [hashtable]$WorkbookTokens,
    [string]$LarkCli,
    [string]$ErrorMessage
) {
    if ($null -eq $Context) { return [pscustomobject]@{ok=$true;status='not_started';transaction_id=$null} }
    $journal = Read-FinanceTransactionJournal $Context.path
    if ($null -eq $journal) { throw "经营事务日志丢失：$($Context.id)" }
    if ([string]$journal.status -eq 'committed') { throw "已提交事务禁止回滚：$($Context.id)" }
    if ([string]$journal.status -eq 'rolled_back') { return [pscustomobject]@{ok=$true;status='rolled_back';transaction_id=$Context.id;restored=0} }
    try {
        $entries = @($journal.entries)
        [array]::Reverse($entries)
        $restored = 0
        $restoredChunks = 0
        foreach ($entry in $entries) {
            $workbookKey = [string]$entry.workbook
            if (-not $WorkbookTokens.ContainsKey($workbookKey) -or -not $WorkbookTokens[$workbookKey]) { throw "回滚缺少工作簿映射：$workbookKey" }
            foreach ($chunk in @(Get-FinanceRollbackChunks $entry)) {
                $operations = @(@{shortcut='+cells-set';input=@{sheet_name=[string]$entry.sheet;range=[string]$chunk.range;cells=@($chunk.cells)}})
                # PowerShell 管道会把单元素数组拆成对象；飞书明确要求 operations 为 JSON 数组。
                $payload = ConvertTo-Json -InputObject $operations -Depth 50 -Compress
                # 飞书 cells-set 对 value='' 的空白单元格并不总是执行清除。
                # 先显式清空内容，再恢复完整写前快照，才能撤销新增的尾部行。
                $clearRanges = ConvertTo-Json -InputObject @("$([string]$entry.sheet)!$([string]$chunk.range)") -Compress
                $lastError = $null
                $succeeded = $false
                for ($attempt = 1; $attempt -le 3; $attempt++) {
                    $clearRaw = $clearRanges | & $LarkCli sheets +cells-batch-clear --as user --spreadsheet-token ([string]$WorkbookTokens[$workbookKey]) --ranges - --scope content --yes --format json | Out-String
                    $clearExitCode = $LASTEXITCODE
                    $clearSucceeded = $false
                    if ($clearExitCode -eq 0) {
                        try {
                            $clearResult = $clearRaw | ConvertFrom-Json
                            $clearSucceeded = [bool]$clearResult.ok
                            if (-not $clearSucceeded) { $lastError = '飞书清空返回 ok=false' }
                        } catch { $lastError = [string]$_.Exception.Message }
                    } else {
                        $lastError = "飞书清空 CLI 退出码 $clearExitCode"
                    }
                    if ($clearSucceeded) {
                        $raw = $payload | & $LarkCli sheets +batch-update --as user --spreadsheet-token ([string]$WorkbookTokens[$workbookKey]) --operations - --yes --format json | Out-String
                        if ($LASTEXITCODE -eq 0) {
                            try {
                                $result = $raw | ConvertFrom-Json
                                if ($result.ok) { $succeeded = $true; break }
                                $lastError = '飞书恢复返回 ok=false'
                            } catch { $lastError = [string]$_.Exception.Message }
                        } else {
                            $lastError = "飞书恢复 CLI 退出码 $LASTEXITCODE"
                        }
                    }
                    if ($attempt -lt 3) { Start-Sleep -Seconds $attempt }
                }
                if (-not $succeeded) { throw "回滚写入失败：$workbookKey/$($entry.sheet)!$($chunk.range)；$lastError" }
                $restoredChunks++
            }
            $restored++
        }
        $eventIds = @($journal.processed_event_ids | Where-Object { $_ })
        $eventLogPath = Join-Path $Context.state_root 'processed_events.ndjson'
        if ($eventIds.Count -gt 0 -and (Test-Path -LiteralPath $eventLogPath)) {
            $kept = [Collections.Generic.List[string]]::new()
            foreach ($line in Get-Content -LiteralPath $eventLogPath) {
                if (-not $line) { continue }
                $remove = $false
                try { $remove = ([string](($line | ConvertFrom-Json).event_id)) -in $eventIds } catch {}
                if (-not $remove) { $kept.Add($line) }
            }
            [System.IO.File]::WriteAllLines($eventLogPath, $kept.ToArray(), [Text.UTF8Encoding]::new($false))
        }
        $null = Update-FinanceTransactionJournal $Context.path {
            param($current)
            ConvertTo-CompactedFinanceTransactionJournal $current 'rolled_back' $ErrorMessage
        }
        return [pscustomobject]@{ok=$true;status='rolled_back';transaction_id=$Context.id;restored=$restored;restored_chunks=$restoredChunks}
    } catch {
        $rollbackError = [string]$_.Exception.Message
        $null = Update-FinanceTransactionJournal $Context.path {
            param($current)
            $current.status = 'rollback_pending'
            $current.error = $ErrorMessage
            $current.rollback_error = $rollbackError
            $current.updated_at = [datetime]::Now.ToString('o')
            return $current
        }
        return [pscustomobject]@{ok=$false;status='rollback_pending';transaction_id=$Context.id;error=$rollbackError}
    }
}

function Complete-FinanceTransaction($Context) {
    if ($null -eq $Context) { return }
    # 提交与删除回滚载荷合并成一次原子替换，避免先重写一份几十 MB 的已提交快照。
    $journal = Update-FinanceTransactionJournal $Context.path {
        param($current)
        ConvertTo-CompactedFinanceTransactionJournal $current 'committed' $null
    }
    $records = @($journal.processed_events | Where-Object { $_ -and $_.event_id })
    if ($records.Count -gt 0) {
        try {
            $eventLogPath = Join-Path $Context.state_root 'processed_events.ndjson'
            $existing = @{}
            if (Test-Path -LiteralPath $eventLogPath) {
                foreach ($line in Get-Content -LiteralPath $eventLogPath) {
                    if (-not $line) { continue }
                    try { $existing[[string](($line | ConvertFrom-Json).event_id)] = $true } catch {}
                }
            }
            foreach ($record in $records) {
                if ($existing.ContainsKey([string]$record.event_id)) { continue }
                [System.IO.File]::AppendAllText(
                    $eventLogPath,
                    (($record | ConvertTo-Json -Compress) + [Environment]::NewLine),
                    [Text.UTF8Encoding]::new($false)
                )
                $existing[[string]$record.event_id] = $true
            }
        } catch {
            # 已提交事务日志本身也是幂等事实源；本地摘要日志失败不能把已提交账目误判为失败。
            $eventLogError = [string]$_.Exception.Message
            $null = Update-FinanceTransactionJournal $Context.path {
                param($current)
                $current | Add-Member -NotePropertyName event_log_error -NotePropertyValue $eventLogError -Force
                $current.updated_at = [datetime]::Now.ToString('o')
                return $current
            }
        }
    }
}

function Repair-PendingFinanceTransactions(
    [string]$StateRoot,
    [hashtable]$WorkbookTokens,
    [string]$LarkCli
) {
    $directory = Get-FinanceTransactionDirectory $StateRoot
    $repaired = [Collections.Generic.List[string]]::new()
    foreach ($path in Get-ChildItem -LiteralPath $directory -Filter '*.json' -File -ErrorAction SilentlyContinue) {
        $status = Get-FinanceTransactionStatusFast $path.FullName
        if ($status -notin @('active','rollback_pending')) { continue }
        try { $journal = Read-FinanceTransactionJournal $path.FullName } catch { continue }
        $context = [pscustomobject]@{
            id = [string]$journal.transaction_id
            state_root = $StateRoot
            path = $path.FullName
            owns_lock = $false
            mutex = $null
        }
        $result = Invoke-FinanceTransactionRollback $context $WorkbookTokens $LarkCli '启动新事务前恢复上次未完成写入'
        if (-not $result.ok) {
            return [pscustomobject]@{ok=$false;repaired=$repaired.ToArray();error=$result.error;transaction_id=$result.transaction_id}
        }
        $repaired.Add([string]$context.id)
    }
    return [pscustomobject]@{ok=$true;repaired=$repaired.ToArray();error=$null}
}

function Close-FinanceTransaction($Context) {
    if ($null -eq $Context) { return }
    if ($Context.owns_lock -and $null -ne $Context.mutex) {
        try { $Context.mutex.ReleaseMutex() } catch {}
        $Context.mutex.Dispose()
    }
    # 加入父事务的子脚本与父脚本在同一个 PowerShell 进程运行。
    # 子脚本结束时不得清除父事务环境，否则重试会错误创建新事务并把父事务当作残留事务回滚。
    if ($Context.owns_lock -and [string]$env:PETSTAR_FINANCE_TRANSACTION_ID -eq [string]$Context.id) {
        Remove-Item Env:PETSTAR_FINANCE_TRANSACTION_ID -ErrorAction SilentlyContinue
        Remove-Item Env:PETSTAR_FINANCE_TRANSACTION_STATE_ROOT -ErrorAction SilentlyContinue
    }
}
