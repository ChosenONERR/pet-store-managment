$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $projectRoot 'agent\finance_transaction.ps1')

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('petstar-finance-transaction-test-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $testRoot
try {
    $context = New-FinanceTransaction $testRoot 'storage regression test' 'TX-TEST-COMPACT'
    $journal = Read-FinanceTransactionJournal $context.path
    $journal.entries = @([pscustomobject]@{
        sequence = 1
        workbook = 'source'
        sheet = 'Sheet1'
        range = 'A2:B3'
        cells = @(
            @([pscustomobject]@{value='a'},[pscustomobject]@{value='b'}),
            @([pscustomobject]@{value='c'},[pscustomobject]@{value='d'})
        )
        captured_at = [datetime]::Now.ToString('o')
    })
    $journal.processed_event_ids = @('EVENT-1')
    $journal.processed_events = @([pscustomobject]@{event_id='EVENT-1';action='test';reference='test';processed_at=[datetime]::Now.ToString('o')})
    Write-FinanceTransactionJournal $context.path $journal
    Complete-FinanceTransaction $context

    $committed = Read-FinanceTransactionJournal $context.path
    if ([string]$committed.status -ne 'committed') { throw 'Transaction did not commit' }
    if ($null -ne $committed.PSObject.Properties['entries']) { throw 'Committed journal retained cell payloads' }
    if ([int]$committed.entry_count -ne 1 -or @($committed.entry_manifest).Count -ne 1) { throw 'Committed journal lost its audit manifest' }
    if ('EVENT-1' -notin @($committed.processed_event_ids)) { throw 'Committed journal lost idempotency state' }
    if ((Get-ChildItem -LiteralPath (Get-FinanceTransactionDirectory $testRoot) -Filter '*.tmp' -File).Count -ne 0) { throw 'Atomic journal write left a temp file' }

    $rows = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt 205; $index++) {
        $rows.Add(@([pscustomobject]@{value=$index},[pscustomobject]@{value=$index}))
    }
    $chunks = @(Get-FinanceRollbackChunks ([pscustomobject]@{range='A2:B206';cells=$rows.ToArray()}))
    if ($chunks.Count -ne 3) { throw "Expected 3 rollback chunks, got $($chunks.Count)" }
    if ([string]$chunks[0].range -ne 'A2:B101' -or [string]$chunks[2].range -ne 'A202:B206') { throw 'Rollback chunk ranges are incorrect' }
    if (@($chunks[0].cells).Count -ne 100 -or @($chunks[2].cells).Count -ne 5) { throw 'Rollback chunk row counts are incorrect' }

    [ordered]@{
        ok = $true
        committed_journal_bytes = (Get-Item -LiteralPath $context.path).Length
        audit_entries = [int]$committed.entry_count
        rollback_chunks = $chunks.Count
        orphan_temp_files = 0
    } | ConvertTo-Json -Compress
} finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
