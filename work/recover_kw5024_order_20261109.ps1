# One-off owner-authorized recovery. Default is read-only; never blindly rerun.
param([switch]$Apply,[string]$ProjectRoot=(Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference='Stop'
$OutputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$agentRoot=Join-Path $ProjectRoot 'agent'
. (Join-Path $agentRoot 'runtime_paths.ps1')
Initialize-PetstarLarkUserEnvironment
. (Join-Path $agentRoot 'finance_transaction.ps1')
. (Join-Path $agentRoot 'order_actual_overrides.ps1')
. (Join-Path $agentRoot 'refund_adjustments.ps1')
$cli=Join-Path $agentRoot 'lark_cli_noexit.ps1'
$stateRoot=Join-Path $agentRoot 'state'
$tokens=@{}
foreach($entry in @(@('source','simple_finance_meta.json'),@('ledger','daily_ledger_meta.json'))){
    $meta=Get-Content (Join-Path $ProjectRoot ('work/'+$entry[1])) -Raw -Encoding UTF8|ConvertFrom-Json
    $tokens[$entry[0]]=([uri]$meta.workbook_url).Segments[-1].Trim('/')
}
$orderNo='3402379981240466045';$reference='sku_ref:657:51185971728'
$eventPrefix='repair-kw5024-3402379981240466045-20261109'
function Invoke-RecoveryCli([string[]]$Arguments){
    $global:LASTEXITCODE=0
    $raw=& $cli @Arguments|Out-String
    if($LASTEXITCODE-ne0){throw '飞书调用失败，未将结果当作成功'}
    $r=$raw|ConvertFrom-Json;if(-not$r.ok){throw '飞书返回失败'};return $r
}
function ReadRows($Book,$Sheet,$Range){
    $r=Invoke-RecoveryCli @('sheets','+cells-get','--as','user','--spreadsheet-token',$tokens[$Book],'--sheet-name',$Sheet,'--range',$Range,'--include','value,formula','--max-chars','3000000','--format','json')
    if($r.data.has_more -or $r.data.ranges[0].truncated){throw "读取不完整：$Sheet"}
    return $r.data.ranges[0]
}
function Cell($Row,[int]$Index){
    if($null-eq$Row -or $Index-ge@($Row).Count -or $null-eq$Row[$Index]){return $null}
    $property=$Row[$Index].PSObject.Properties['value'];if($null-ne$property){return $property.Value};return $null
}
function Verify($Book,$Sheet,$Range){
    $r=Invoke-RecoveryCli @('sheets','+formula-verify','--as','user','--spreadsheet-token',$tokens[$Book],'--sheet-name',$Sheet,'--range',$Range,'--exit-on-error','--format','json')
    if($r.data.status-ne'success'){throw "公式验证失败：$Sheet $Range"}
}
function CheckValues($Row,[int]$Income,[int]$Cost,[int]$Fee,[int]$Profit){
    foreach($pair in @(@($Income,36.75),@($Cost,29.8),@($Fee,6.37),@($Profit,0.58))){
        $value=Cell $Row ([int]$pair[0])
        if($null-eq$value -or [math]::Abs([double]$value-[double]$pair[1])-gt0.00001){throw '订单收入、成本、配送费或利润回读不符'}
    }
}
function InvokeEvent([hashtable]$Arguments){
    $global:LASTEXITCODE=0
    $raw=& (Join-Path $agentRoot 'record_finance_event.ps1') @Arguments -ProjectRoot $ProjectRoot|Out-String
    if($LASTEXITCODE-ne0){throw '补记子步骤失败'}
    $r=$raw|ConvertFrom-Json;if(-not$r.ok){throw '补记子步骤未返回成功'}
    return $r
}
$transaction=$null;$closed=$false
try{
    if($Apply){$transaction=New-FinanceTransaction $stateRoot 'KW5024指定订单成本及多送24片补记' $null -AcquireLock -WorkbookTokens $tokens -LarkCli $cli -LockTimeoutSeconds 1200}
    $summary=(ReadRows 'source' '订单汇总' 'A476:Q476').cells[0]
    if((Cell $summary 2)-ne$orderNo -or (Cell $summary 0)-ne'2026-11-08' -or (Cell $summary 3)-ne'已完成' -or (Cell $summary 4)-ne'自然'){throw '目标订单行、日期或状态已变化'}
    if([double](Cell $summary 7)-ne36.75 -or [double](Cell $summary 9)-ne6.37 -or [double](Cell $summary 5)-ne39.2){throw '原成交金额或配送费变化，停止补记'}
    $master=(ReadRows 'source' 'SKU成本主表' 'A657:Z657').cells[0]
    if((Cell $master 0)-ne'kw5024' -or (Cell $master 16)-ne'51185971728' -or (Cell $master 18)-ne'L码60*60（40片起购）'){throw '商品安全定位失效'}
    if(@(Get-OrderActualOverride $ProjectRoot $orderNo).Count -or @(Get-RefundEventsForOrder $ProjectRoot $orderNo).Count){throw '订单已有履约覆盖或退款，停止叠加补记'}
    $items=ReadRows 'source' '订单商品' 'A1534:Q1535'
    $qty=0.0;$sales=0.0
    foreach($row in $items.cells){
        if((Cell $row 2)-ne$orderNo -or (Cell $row 16)-ne'51185971728' -or (Cell $row 5)-ne'kw5024' -or [double](Cell $row 9)-ne0.7){throw '拆分明细已变化'}
        $qty += [double](Cell $row 8);$sales += [double](Cell $row 10)
    }
    if($qty-ne56 -or [math]::Abs($sales-39.2)-gt0.00001){throw '原购买数量或销售额不符'}
    $gifts=ReadRows 'source' '赠品记录' 'A2:L2001'
    $existing=@($gifts.cells|Where-Object{(Cell $_ 3)-eq$orderNo})
    if($existing.Count){
        if($existing.Count-ne1 -or (Cell $existing[0] 11)-ne($eventPrefix+'-gift') -or [double](Cell $existing[0] 5)-ne24){throw '订单已有其他赠品，不可重复补记'}
        CheckValues $summary 7 10 9 11
        if([double](Cell $master 5)-ne0.3725 -or [double](Cell $master 6)-ne0.3725){throw '已补记但主表成本不符'}
        if($Apply){Complete-FinanceTransaction $transaction;Close-FinanceTransaction $transaction;$closed=$true}
        [ordered]@{ok=$true;already_recovered=$true;order_no=$orderNo;cost=29.8;profit=0.58}|ConvertTo-Json -Compress
        return
    }
    $plan=[ordered]@{order_no=$orderNo;date='2026-11-08';paid_quantity=56;free_quantity=24;delivered_quantity=80;unit_cost=0.3725;product_sales=39.2;income=36.75;delivery_fee=6.37;cost=29.8;profit=0.58}
    if(-not$Apply){[ordered]@{ok=$true;preview=$true;plan=$plan;writes=0}|ConvertTo-Json -Depth 4;return}
    $null=InvokeEvent @{Action='SetSkuCost';GoodsCode=$reference;Amount=0.3725;OrderNo=$orderNo;Date='2026-11-08';EventId=($eventPrefix+'-cost');Note='店主确认每片0.3725元，仅修正本单成本，不倒改其他历史订单'}
    $null=InvokeEvent @{Action='AddOrderGift';OrderNo=$orderNo;Date='2026-11-08';Items=($reference+'×24');EventId=($eventPrefix+'-gift');Note='店主确认购买56片，实际交付80片；多送24片无额外收入，销售额和配送费不变，不改美团库存'}
    $summary=(ReadRows 'source' '订单汇总' 'A476:Q476').cells[0]
    CheckValues $summary 7 10 9 11
    if([double](Cell $summary 5)-ne39.2){throw '补记不应改变原商品销售额'}
    $master=(ReadRows 'source' 'SKU成本主表' 'F657:G657').cells[0]
    if([double](Cell $master 0)-ne0.3725 -or [double](Cell $master 1)-ne0.3725){throw '主表成本回读不符'}
    Verify 'source' '订单汇总' 'K476:L476'
    $incomeRows=ReadRows 'ledger' '11月收入' 'A2:K1001'
    $incomeMatches=@(for($i=0;$i-lt$incomeRows.cells.Count;$i++){if((Cell $incomeRows.cells[$i] 1)-eq$orderNo){[pscustomobject]@{row=$incomeRows.row_indices[$i];cells=$incomeRows.cells[$i]}}})
    if($incomeMatches.Count-ne1){throw '11月收入中订单未唯一匹配'}
    CheckValues $incomeMatches[0].cells 3 4 5 6
    $r=$incomeMatches[0].row;Verify 'ledger' '11月收入' "E${r}:G${r}"
    Complete-FinanceTransaction $transaction;Close-FinanceTransaction $transaction;$closed=$true
    [ordered]@{ok=$true;plan=$plan;transaction_id=$transaction.id;income_row=$r;gift_event_id=($eventPrefix+'-gift');formula_verification='success';inventory_changed=$false}|ConvertTo-Json -Depth 5
}catch{
    $reason=$_.Exception.Message
    if($null-ne$transaction -and -not$closed){$rollback=Invoke-FinanceTransactionRollback $transaction $tokens $cli $reason;Close-FinanceTransaction $transaction;$closed=$true;if(-not$rollback.ok){throw "$reason；回滚尚未完成，禁止重试写入"}}
    throw $reason
}finally{if($null-ne$transaction -and -not$closed){Close-FinanceTransaction $transaction}}
