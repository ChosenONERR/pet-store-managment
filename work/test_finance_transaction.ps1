param([string]$ProjectRoot)
$ErrorActionPreference = 'Stop'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
. (Join-Path $ProjectRoot 'agent\finance_transaction.ps1')

$testRoot = Join-Path $env:TEMP ("petstar-finance-tx-" + [guid]::NewGuid().ToString('N'))
$stateRoot = Join-Path $testRoot 'state'
$fakeState = Join-Path $testRoot 'fake_lark.json'
$fakeCli = Join-Path $ProjectRoot 'work\fake_lark_cli_for_transaction_test.ps1'
$null = New-Item -ItemType Directory -Path $stateRoot -Force
$env:PETSTAR_FAKE_LARK_STATE = $fakeState

function Write-FakeState($Cells) {
    $state = [ordered]@{ranges=[ordered]@{'source|Sheet1|A1:B2'=@($Cells)}}
    [IO.File]::WriteAllText($fakeState,(($state|ConvertTo-Json -Depth 50)+[Environment]::NewLine),[Text.UTF8Encoding]::new($false))
}
function Read-FakeCells {
    $state = [IO.File]::ReadAllText($fakeState,[Text.Encoding]::UTF8)|ConvertFrom-Json
    return @($state.ranges.PSObject.Properties['source|Sheet1|A1:B2'].Value)
}

try {
    $before = @(
        @(@{value='alpha';cell_styles=@{bold=$true}},@{formula='=1+1';cell_styles=@{number_format='0.00'}}),
        @(@{value=''},@{value=7})
    )
    Write-FakeState $before
    $tx = New-FinanceTransaction $stateRoot 'rollback-test' 'rollback-test' -AcquireLock -WorkbookTokens @{source='source'} -LarkCli $fakeCli
    Register-FinanceRangeBeforeImage $tx 'source' 'source' 'Sheet1' 'A1:B2' $fakeCli
    Write-FakeState @(@(@{value='changed'},@{value=99}),@(@{value='new'},@{value=100}))
    $rollback = Invoke-FinanceTransactionRollback $tx @{source='source'} $fakeCli 'fault injection'
    Close-FinanceTransaction $tx
    if (-not $rollback.ok -or $rollback.restored -ne 1) { throw '事务回滚没有成功' }
    $rollbackState = [IO.File]::ReadAllText($fakeState,[Text.Encoding]::UTF8)|ConvertFrom-Json
    if ('source|Sheet1|A1:B2' -notin @($rollbackState.clear_calls)) { throw '事务回滚没有先显式清空目标范围' }
    $restored = Read-FakeCells
    if ([string]$restored[0][0].value -ne 'alpha' -or [string]$restored[0][1].formula -ne '=1+1' -or [double]$restored[1][1].value -ne 7) { throw '写前快照未完整恢复' }
    $sparseRange = [pscustomobject]@{cells=@(@(@{value='only'}));row_indices=@(1);col_indices=@()}
    $expanded = ConvertTo-RestorableMatrix $sparseRange 'A1:B2'
    if (@($expanded).Count -ne 2 -or @($expanded[0]).Count -ne 2 -or [string]$expanded[1][1].value -ne '') { throw '稀疏快照没有扩展为完整目标范围' }

    $letterRange = [pscustomobject]@{cells=@(@([pscustomobject]@{value='left'},[pscustomobject]@{value='right'}));row_indices=@(2);col_indices=@('A','B')}
    $letterExpanded = ConvertTo-RestorableMatrix $letterRange 'A2:B2'
    if ([string]$letterExpanded[0][0].value -ne 'left' -or [string]$letterExpanded[0][1].value -ne 'right') { throw '飞书字母列坐标没有正确解析' }
    $liveOneRowCells = [object[]]@(,[object[]]@([pscustomobject]@{value='live-left'},[pscustomobject]@{formula='=1+2';value=3}))
    $liveOneRowRange = [pscustomobject]@{cells=$liveOneRowCells;row_indices=@(86);col_indices=@('A','B')}
    $liveOneRowExpanded = ConvertTo-RestorableMatrix $liveOneRowRange 'A86:B86'
    if ([string]$liveOneRowExpanded[0][0].value -ne 'live-left' -or [string]$liveOneRowExpanded[0][1].formula -ne '=1+2') { throw '真实接口单行嵌套数组被错误保存为空白快照' }

    Write-FakeState $before
    $parent = New-FinanceTransaction $stateRoot 'parent-context-test' 'parent-context-test' -AcquireLock -WorkbookTokens @{source='source'} -LarkCli $fakeCli
    $child = Join-FinanceTransaction $stateRoot
    Close-FinanceTransaction $child
    if ([string]$env:PETSTAR_FINANCE_TRANSACTION_ID -ne 'parent-context-test') { throw '子脚本错误清除了父事务上下文' }
    Complete-FinanceTransaction $parent
    Close-FinanceTransaction $parent
    if ($env:PETSTAR_FINANCE_TRANSACTION_ID) { throw '父事务结束后没有清理事务上下文' }

    Write-FakeState $before
    # Business-event idempotency uses the deterministic event-<event_id> journal path.
    $tx2 = New-FinanceTransaction $stateRoot 'commit-test' 'event-event-001' -AcquireLock -WorkbookTokens @{source='source'} -LarkCli $fakeCli
    Register-FinanceProcessedEvent $tx2 'event-001' 'SetSkuCost' 'kw1109'
    if (Test-Path (Join-Path $stateRoot 'processed_events.ndjson')) { throw '事务提交前不应写入幂等事件日志' }
    Complete-FinanceTransaction $tx2
    Close-FinanceTransaction $tx2
    if (-not (Test-FinanceCommittedEvent $stateRoot 'event-001')) { throw '已提交事务未被幂等检查识别' }
    $event = Get-Content (Join-Path $stateRoot 'processed_events.ndjson') -Raw | ConvertFrom-Json
    if ([string]$event.event_id -ne 'event-001') { throw '提交后幂等事件日志不正确' }

    Write-FakeState $before
    $stale = New-FinanceTransaction $stateRoot 'stale-test' 'stale-test'
    Register-FinanceRangeBeforeImage $stale 'source' 'source' 'Sheet1' 'A1:B2' $fakeCli
    Write-FakeState @(@(@{value='partial'},@{value=88}),@(@{value='partial'},@{value=77}))
    Close-FinanceTransaction $stale
    $next = New-FinanceTransaction $stateRoot 'repair-test' 'repair-test' -AcquireLock -WorkbookTokens @{source='source'} -LarkCli $fakeCli
    $staleJournal = Read-FinanceTransactionJournal (Get-FinanceTransactionPath $stateRoot 'stale-test')
    if ([string]$staleJournal.status -ne 'rolled_back') { throw '启动新写入前没有恢复上次活动事务' }
    $afterRepair = Read-FakeCells
    if ([string]$afterRepair[0][0].value -ne 'alpha') { throw '活动事务自动恢复结果不正确' }
    Complete-FinanceTransaction $next
    Close-FinanceTransaction $next

    Write-FakeState $before
    $pending = New-FinanceTransaction $stateRoot 'rollback-pending-test' 'rollback-pending-test'
    Register-FinanceRangeBeforeImage $pending 'source' 'source' 'Sheet1' 'A1:B2' $fakeCli
    Write-FakeState @(@(@{value='partial'},@{value=1}),@(@{value='partial'},@{value=2}))
    $env:PETSTAR_FAKE_LARK_FAIL_BATCH = '1'
    $pendingResult = Invoke-FinanceTransactionRollback $pending @{source='source'} $fakeCli 'injected failure'
    Close-FinanceTransaction $pending
    if ($pendingResult.ok -or $pendingResult.status -ne 'rollback_pending') { throw '回滚失败没有进入阻断状态' }
    $blocked = $false
    try { $null = New-FinanceTransaction $stateRoot 'must-block' 'must-block' -AcquireLock -WorkbookTokens @{source='source'} -LarkCli $fakeCli } catch { $blocked = $true }
    if (-not $blocked) { throw '存在未恢复事务时仍允许新写入' }
    Remove-Item Env:PETSTAR_FAKE_LARK_FAIL_BATCH -ErrorAction SilentlyContinue
    $recovered = New-FinanceTransaction $stateRoot 'recover-pending' 'recover-pending' -AcquireLock -WorkbookTokens @{source='source'} -LarkCli $fakeCli
    if ([string](Read-FinanceTransactionJournal (Get-FinanceTransactionPath $stateRoot 'rollback-pending-test')).status -ne 'rolled_back') { throw '回滚待恢复事务没有在下一次写入前修复' }
    Complete-FinanceTransaction $recovered
    Close-FinanceTransaction $recovered

    [ordered]@{ok=$true;rollback='clear-then-restored';real_column_coordinates='A-B-supported';live_one_row_shape='preserved';nested_context='parent-preserved';commit='event-after-commit';stale_recovery='before-new-write';rollback_pending='blocks-and-recovers';global_lock='single-mutation-path'} | ConvertTo-Json -Compress
} finally {
    Remove-Item Env:PETSTAR_FAKE_LARK_STATE -ErrorAction SilentlyContinue
    Remove-Item Env:PETSTAR_FAKE_LARK_FAIL_BATCH -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
