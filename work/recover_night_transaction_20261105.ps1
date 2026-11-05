param([switch]$Apply)
$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'agent\finance_transaction.ps1')
$state = Join-Path $root 'agent\state'
$id = 'TX-20261105-000009-CD6BC333'
$path = Get-FinanceTransactionPath $state $id
$journal = Read-FinanceTransactionJournal $path
if ($null -eq $journal -or $journal.transaction_id -ne $id) { throw '指定恢复日志缺失或身份不符' }
if ($journal.status -eq 'rolled_back') { @{ok=$true;already_restored=$true;transaction_id=$id}|ConvertTo-Json; return }
if ($journal.status -ne 'rollback_pending' -or $journal.description -ne '定时经营同步 2026-11-04 night') { throw '事务状态或用途变化，禁止重跑' }
$entries = @($journal.entries)
if ($entries.Count -ne 27 -or @($entries | Where-Object {$_.workbook -ne 'source' -or $null -eq $_.cells}).Count) { throw '恢复范围与已核实的27份快照不符' }
if (-not $Apply) {
    [ordered]@{ok=$true;preview=$true;transaction_id=$id;ranges=@($entries | Select-Object sheet,range);external_writes=0}|ConvertTo-Json -Depth 5
    return
}
$cli = Join-Path $root 'agent\lark_cli_noexit.ps1'
$meta = Get-Content -LiteralPath (Join-Path $root 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$book = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
function Read-RecoveryRange([string]$Sheet,[string]$Range) {
    $raw = & $cli sheets +cells-get --as user --spreadsheet-token $book --sheet-name $Sheet --range $Range --include 'value,formula,style' --max-chars 5000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "恢复回读失败：$Sheet $Range" }
    $r = $raw | ConvertFrom-Json
    if (-not $r.ok -or $r.data.has_more -or $r.data.ranges[0].truncated) { throw '恢复回读不完整' }
    return $r.data.ranges[0]
}
function Recovery-CellSignature($Cell) {
    if ($null -eq $Cell) { return '' }
    if ($Cell -is [Collections.IDictionary]) {
        if ($Cell.Contains('formula') -and $Cell['formula']) { return 'formula:'+[string]$Cell['formula'] }
        if (-not $Cell.Contains('value') -or $null -eq $Cell['value'] -or [string]$Cell['value'] -eq '') { return '' }
        return 'value:'+[string]$Cell['value']
    }
    $f = $Cell.PSObject.Properties['formula']
    if ($null -ne $f -and $f.Value) { return 'formula:'+[string]$f.Value }
    $v = $Cell.PSObject.Properties['value']
    if ($null -eq $v -or $null -eq $v.Value -or [string]$v.Value -eq '') { return '' }
    return 'value:'+[string]$v.Value
}
$sync = [Threading.Mutex]::new($false,'Local\PetstarSyncAndNotify')
$finance = [Threading.Mutex]::new($false,'Local\PetstarFinanceMutationV2')
$syncHeld=$false;$financeHeld=$false
try {
    try {$syncHeld=$sync.WaitOne(0)} catch [Threading.AbandonedMutexException] {$syncHeld=$true}
    if (-not $syncHeld) { throw '已有同步运行，未开始恢复' }
    try {$financeHeld=$finance.WaitOne([timespan]::FromSeconds(15))} catch [Threading.AbandonedMutexException] {$financeHeld=$true}
    if (-not $financeHeld) { throw '已有经营写入，未开始恢复' }
    $null=Read-RecoveryRange '驾驶舱' 'A1:A1'
    $backup=Join-Path $state ($id+'.recovery-beforeimage')
    if(-not (Test-Path -LiteralPath $backup)){Copy-Item -LiteralPath $path -Destination $backup}
    $context=[pscustomobject]@{id=$id;state_root=$state;path=$path;owns_lock=$false;mutex=$null}
    $result=Invoke-FinanceTransactionRollback $context @{source=$book} $cli '店主确认恢复2026-11-04未完成日结'
    if (-not $result.ok) { throw "事务恢复未完成：$($result.error)" }
    foreach ($entry in $entries) {
        $actual=ConvertTo-RestorableMatrix (Read-RecoveryRange $entry.sheet $entry.range) $entry.range
        $expected=@($entry.cells)
        for($i=0;$i -lt $expected.Count;$i++) {
            for($k=0;$k -lt @($expected[$i]).Count;$k++) {
                if ((Recovery-CellSignature $expected[$i][$k]) -ne (Recovery-CellSignature $actual[$i][$k])) { throw "恢复快照不一致：$($entry.sheet) $($entry.range)，行偏移$i 列偏移$k" }
            }
        }
    }
    foreach($check in @(@('订单汇总','K2:L460'),@('库存快照','H2:M2000'))) {
        $raw=& $cli sheets +formula-verify --as user --spreadsheet-token $book --sheet-name $check[0] --range $check[1] --exit-on-error --format json | Out-String
        if($LASTEXITCODE -ne 0 -or ($raw|ConvertFrom-Json).data.status -ne 'success'){throw "恢复后公式检查失败：$($check[0])"}
    }
    [ordered]@{ok=$true;transaction_id=$id;status=$result.status;snapshots_verified=$entries.Count;formula_verification='success'}|ConvertTo-Json
} finally {
    if($financeHeld){$finance.ReleaseMutex()};$finance.Dispose()
    if($syncHeld){$sync.ReleaseMutex()};$sync.Dispose()
}
