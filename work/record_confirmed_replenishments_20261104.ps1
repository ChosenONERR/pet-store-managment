param([switch]$Apply)
$ErrorActionPreference='Stop'
$OutputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'agent/todo_store.ps1')
. (Join-Path $root 'agent/finance_transaction.ps1')
$ctx=Get-TodoContext $root
$state=Join-Path $root 'agent/state'
$hash='bc161ddec7bc1365f4611a61eacebeb5f57d9c1c84b55860b95cb6430eb3c3af'
$eventId='confirmed-replenishment-20261104-'+$hash.Substring(0,12)
$txId='recovery-'+$eventId
$msgMutex=[Threading.Mutex]::new($false,('Local\PetstarLarkMessage_'+$hash.Substring(0,32)))
$locked=$false;$tx=$null;$committed=$false
function Read-Cells([string]$Sheet,[string]$Range,[string]$Include='value,formula,style'){
    $r=Invoke-TodoLark $ctx @('sheets','+cells-get','--as','user','--spreadsheet-token',$ctx.token,'--sheet-name',$Sheet,'--range',$Range,'--include',$Include,'--max-chars','1000000','--format','json')
    if($r.data.has_more -or $r.data.ranges[0].truncated){throw "读取不完整：$Sheet $Range"}
    return $r.data.ranges[0]
}
function Mark-Recovered {
    $mutex=[Threading.Mutex]::new($false,'Local\PetstarLarkHandledMessageLedger');$taken=$false
    try{
        try{$taken=$mutex.WaitOne(15000)}catch [Threading.AbandonedMutexException]{$taken=$true}
        if(-not $taken){throw '经营登记已提交，但原消息恢复标记待重试'}
        $path=Join-Path $state 'handled_inbound_messages.ndjson'
        $exists=@(Get-Content -LiteralPath $path -Encoding utf8|ForEach-Object{try{$_|ConvertFrom-Json}catch{}}|Where-Object{$_.message_hash -eq $hash -and $_.status -eq 'completed'}).Count
        if(-not $exists){$record=[ordered]@{message_hash=$hash;status='completed';source='confirmed_desktop_recovery';create_time='1793725500000';recorded_at=[datetime]::Now.ToString('o');transaction_id=$txId};[IO.File]::AppendAllText($path,(($record|ConvertTo-Json -Compress)+[Environment]::NewLine),[Text.UTF8Encoding]::new($false))}
    }finally{if($taken){$mutex.ReleaseMutex()};$mutex.Dispose()}
}
try{
    if($Apply){try{$locked=$msgMutex.WaitOne(30000)}catch [Threading.AbandonedMutexException]{$locked=$true};if(-not $locked){throw '原消息仍在处理，本次未写入'}}
    $journalPath=Get-FinanceTransactionPath $state $txId
    if((Get-FinanceTransactionStatusFast $journalPath) -eq 'committed'){
        if($Apply){Mark-Recovered};[pscustomobject]@{ok=$true;duplicate=$true;transaction=$txId}|ConvertTo-Json;return
    }
    if($Apply){$tx=New-FinanceTransaction $state '店主确认盒装规格及补货前0盒；6笔采购231元' $txId -AcquireLock -WorkbookTokens @{source=$ctx.token} -LarkCli $ctx.cli -LockTimeoutSeconds 30}
    $header=(Read-Cells 'SKU成本主表' 'A1:Z1').cells[0];$cols=@{}
    for($c=0;$c -lt @($header).Count;$c++){$cols[(Get-TodoCell $header $c)]=$c}
    foreach($name in @('商品ID（美团SKU ID）','当前移动平均成本','成本账数量')){if(-not $cols.ContainsKey($name)){throw "缺少必要表头：$name"}}
    $boxes=Read-Cells 'SKU成本主表' 'A1450:Z1451'
    $ids=@('54650922708','54648071635')
    for($i=0;$i -lt 2;$i++){$cell=$boxes.cells[$i];if((Get-TodoCell $cell $cols['商品ID（美团SKU ID）']) -ne $ids[$i] -or [double](Get-TodoCell $cell $cols['成本账数量']) -ne 3 -or (Get-TodoCell $cell $cols['当前移动平均成本'])){throw '盒装档案或数量成本已变化，停止本次专项恢复'}}
    $existing=Read-Cells '补货记录' 'A2:T1001' 'value'
    $day=([datetime]'2026-11-04'-[datetime]'1899-12-30').TotalDays
    foreach($row in $existing.cells){if((Get-TodoCell $row 1) -in @('kw4171','kw4172','kw1033','kw1018','kw1013','kw1034') -and (Get-TodoCell $row 0) -in @([string]$day,'2026-11-04')){throw '本日已有对应货号补货，请先核实是否重复'}}
    $ops=@()
    foreach($ref in @('sku_ref:1451:54648071635','sku_ref:1450:54650922708')){
        $ops+=,[ordered]@{action='SetCostBookQuantity';arguments=@{goods_code=$ref;quantity=0;date='2026-11-04';note='店主确认本次补货前0盒；快照已有3盒即本次采购，防止重复计入';event_id="$eventId-quantity-$ref"}}
    }
    $rows=@(@('sku_ref:1451:54648071635',3,21.35,'kw4171'),@('sku_ref:1450:54650922708',3,21.35,'kw4172'),@('kw1033',1,39.41,'kw1033'),@('kw1018',1,39.41,'kw1018'),@('kw1013',2,54.74,'kw1013'),@('kw1034',2,54.74,'kw1034'))
    foreach($item in $rows){$ops+=,[ordered]@{action='Replenishment';arguments=@{goods_code=$item[0];quantity=$item[1];purchase_amount=$item[2];payment_amount=$item[2];payment_date='2026-11-04';date='2026-11-04';note='店主确认本批货款211元＋运费20元＝231元；运费按货款比例分摊；仅登记补货及当前成本';event_id="$eventId-$($item[3])"}}}
    $requestPath=Join-Path $state ($eventId+'.json')
    [IO.File]::WriteAllText($requestPath,(ConvertTo-Json -InputObject $ops -Depth 8),[Text.UTF8Encoding]::new($false))
    $exe='C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
    $entry=Join-Path $root 'agent/record_finance_event.ps1'
    $previewRaw=& $exe -NoProfile -ExecutionPolicy Bypass -File $entry -Action BatchSkuMaintenance -OperationsPath $requestPath -EventId $eventId -DryRun 2>&1|Out-String
    if($LASTEXITCODE -ne 0){throw "预演失败：$previewRaw"}
    $preview=$previewRaw|ConvertFrom-Json
    $items=@($preview.result.operations|Where-Object{$_.PSObject.Properties['purchase_amount']})
    if(-not $preview.ok -or $items.Count -ne 6 -or [math]::Abs(($items|Measure-Object purchase_amount -Sum).Sum-231) -gt 0.000001){throw '六笔补货预演合计不正确'}
    foreach($item in $items|Where-Object{$_.goods_code -in @('kw4171','kw4172')}){if($item.after_quantity -ne 3 -or [math]::Abs($item.after_average_cost-21.35/3) -gt 0.000001){throw '盒装数量或成本预演不正确'}}
    if(-not $Apply){[pscustomobject]@{ok=$true;dry_run=$true;items=$items;total=231;historical_changes=$false}|ConvertTo-Json -Depth 7;return}
    $old=Read-Cells 'SKU成本主表' 'A591:Z592'
    foreach($item in $items){$r=[int]$item.row;Register-FinanceRangeBeforeImage $tx 'source' $ctx.token '补货记录' "A${r}:U${r}" $ctx.cli;$null=Invoke-TodoLark $ctx @('sheets','+range-copy','--as','user','--spreadsheet-token',$ctx.token,'--sheet-name','补货记录','--source-range','A2:U2','--target-range',"A${r}:U${r}",'--paste-type','formats','--format','json')}
    $raw=& $exe -NoProfile -ExecutionPolicy Bypass -File $entry -Action BatchSkuMaintenance -OperationsPath $requestPath -EventId $eventId 2>&1|Out-String
    if($LASTEXITCODE -ne 0){throw "执行失败：$raw"}
    $result=$raw|ConvertFrom-Json
    if(-not $result.ok -or $result.cost_snapshot_rebuild -or $result.daily_ledger_sync){throw '执行结果失败或超出普通补货范围'}
    $written=@($result.result.operations|Where-Object{$_.PSObject.Properties['purchase_amount']})
    if($written.Count -ne 6){throw '执行结果缺少补货项目'}
    $oldAfter=Read-Cells 'SKU成本主表' 'A591:Z592'
    if((ConvertTo-Json -InputObject $old.cells -Depth 15 -Compress) -ne (ConvertTo-Json -InputObject $oldAfter.cells -Depth 15 -Compress)){throw '旧罐装档案发生变化'}
    $paid=0.0;foreach($item in $written){$r=[int]$item.row;$check=Read-Cells '补货记录' "A${r}:U${r}";$cells=$check.cells[0];if((Get-TodoCell $cells 1) -ne $item.goods_code -or [math]::Abs([double](Get-TodoCell $cells 5)-$item.purchase_amount) -gt 0.000001){throw '补货回读不一致'};$paid += [double](Get-TodoCell $cells 12);if(-not (Get-TodoCell $cells 11)){throw '付款日期缺失'}}
    if([math]::Abs($paid-231) -gt 0.000001){throw '付款合计不等于231元'}
    Complete-FinanceTransaction $tx;$committed=$true
    Mark-Recovered
    [pscustomobject]@{ok=$true;applied=$true;items=$written;total=231;payment_total=$paid;old_profiles_unchanged=$true;historical_changes=$false;transaction=$txId}|ConvertTo-Json -Depth 7
}catch{
    if($tx -and -not $committed){$rollback=Invoke-FinanceTransactionRollback $tx @{source=$ctx.token} $ctx.cli $_.Exception.Message;if(-not $rollback.ok){throw '专项恢复失败且回滚待处理，请检查事务日志'}}
    throw
}finally{if($tx){Close-FinanceTransaction $tx};if($locked){$msgMutex.ReleaseMutex()};$msgMutex.Dispose()}
