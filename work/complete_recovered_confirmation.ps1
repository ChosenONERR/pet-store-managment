param(
    [Parameter(Mandatory=$true)][string]$ChangeId,
    [Parameter(Mandatory=$true)][string]$ProcessedEventId,
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$stateRoot = Join-Path $ProjectRoot 'agent\state'
$pendingPath = Join-Path $stateRoot 'pending_confirmations.json'
$processedPath = Join-Path $stateRoot 'processed_events.ndjson'
if (-not (Test-Path -LiteralPath $pendingPath)) { throw '待确认状态文件不存在' }
if (-not (Test-Path -LiteralPath $processedPath)) { throw '幂等事件日志不存在' }

$processed = Select-String -LiteralPath $processedPath -SimpleMatch ('"event_id":"{0}"' -f $ProcessedEventId)
if (-not $processed) { throw "事件 $ProcessedEventId 尚未确认执行成功，拒绝清理待确认项" }

$mutex = [Threading.Mutex]::new($false, 'Local\PetstarLarkPendingConfirmations')
$lockTaken = $false
try {
    try { $lockTaken = $mutex.WaitOne([timespan]::FromSeconds(15)) }
    catch [Threading.AbandonedMutexException] { $lockTaken = $true }
    if (-not $lockTaken) { throw '待确认队列正忙' }

    $state = Get-Content -LiteralPath $pendingPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $state.active -or [string]$state.active.change_id -ne $ChangeId) {
        throw "当前活动项不是 $ChangeId，拒绝修改队列"
    }

    $backupPath = Join-Path $stateRoot ("pending_confirmations_recovery_{0}_{1}.json" -f $ChangeId,[datetime]::Now.ToString('yyyyMMdd_HHmmss'))
    Copy-Item -LiteralPath $pendingPath -Destination $backupPath

    $queue = @($state.queue | Where-Object { $null -ne $_ })
    if ($queue.Count -gt 0) {
        $state.active = $queue[0]
        $state.queue = if ($queue.Count -gt 1) { @($queue[1..($queue.Count-1)]) } else { @() }
    } else {
        $state.active = $null
        $state.queue = @()
    }
    $state.updated_at = [datetime]::Now.ToString('o')
    $tempPath = "$pendingPath.$([guid]::NewGuid().ToString('N')).tmp"
    [IO.File]::WriteAllText($tempPath,($state | ConvertTo-Json -Depth 40),[Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tempPath -Destination $pendingPath -Force

    [ordered]@{ok=$true;change_id=$ChangeId;backup=$backupPath;next_change_id=$(if($state.active){$state.active.change_id}else{$null})} | ConvertTo-Json -Compress
} finally {
    if ($lockTaken) { try { $mutex.ReleaseMutex() } catch {} }
    $mutex.Dispose()
}
