param([switch]$Apply,[string]$ProjectRoot=(Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$agentRoot=Join-Path $ProjectRoot 'agent'
$cli=Join-Path $agentRoot 'lark_cli_noexit.ps1'
$stateRoot=Join-Path $agentRoot 'state'
. (Join-Path $agentRoot 'finance_transaction.ps1')
. (Join-Path $agentRoot 'refund_adjustments.ps1')
$sourceMeta=Get-Content -LiteralPath (Join-Path $ProjectRoot 'work/simple_finance_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$ledgerMeta=Get-Content -LiteralPath (Join-Path $ProjectRoot 'work/daily_ledger_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$sourceToken=([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken=([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$targets=@(
    @{order='3602371202674777015';date='2026-11-02';title='麦富迪羊奶粉30g幼猫哺乳期孕期宠物专用保健品补充营养30g一罐 (30g一罐)';spec='30g一罐';cost=7.0},
    @{order='2102365133521804156';date='2026-10-29';title='【主食猫条】蓝氏 全价主食高肉主食猫条 16gx12 支／盒 (16g*12支/盒)';spec='16g*12支/盒';cost=17.64}
)
function Invoke-RepairCli([string[]]$Arguments,[string]$Payload){
    $global:LASTEXITCODE=0
    $raw=if($PSBoundParameters.ContainsKey('Payload')){$Payload|& $cli @Arguments|Out-String}else{& $cli @Arguments|Out-String}
    if($LASTEXITCODE-ne0){throw "飞书调用失败：$raw"}
    $result=$raw|ConvertFrom-Json
    if(-not $result.ok){throw "飞书调用失败：$raw"}
    return $result
}
function Read-RepairRows($Token,$Sheet,$Range){
    $r=Invoke-RepairCli @('sheets','+cells-get','--as','user','--spreadsheet-token',$Token,'--sheet-name',$Sheet,'--range',$Range,'--include','value,formula','--max-chars','3000000','--format','json')
    if($r.data.has_more -or $r.data.ranges[0].truncated){throw "读取不完整：$Sheet"}
    return $r.data.ranges[0]
}
function Cell($Row,[int]$Index,[string]$Field='value'){
    if($null-eq$Row -or $Index-ge@($Row).Count){return $null}
    $c=@($Row)[$Index];if($null-eq$c){return $null}
    $p=$c.PSObject.Properties[$Field];if($null-ne$p){return $p.Value}
    return $null
}
function Number($Value){
    $n=0.0
    if([string]::IsNullOrWhiteSpace([string]$Value) -or -not[double]::TryParse([string]$Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){throw "金额或数量不明确：$Value"}
    return $n
}
function Verify-RepairFormulas($Token,$Sheet,$Range){
    $r=Invoke-RepairCli @('sheets','+formula-verify','--as','user','--spreadsheet-token',$Token,'--sheet-name',$Sheet,'--range',$Range,'--format','json','--exit-on-error')
    if($r.data.status-ne'success'){throw "公式校验失败：$Sheet $Range"}
}
$transaction=$null;$closed=$false
try{
    if($Apply){$transaction=New-FinanceTransaction $stateRoot '店主确认两笔未编码规格的订单成本' $null -AcquireLock -WorkbookTokens @{source=$sourceToken;ledger=$ledgerToken} -LarkCli $cli}
    $orders=Read-RepairRows $sourceToken '订单汇总' 'A2:Q2001'
    $items=Read-RepairRows $sourceToken '订单商品' 'A2:Q4001'
    $gifts=Read-RepairRows $sourceToken '赠品记录' 'A2:L2001'
    $plans=@()
    foreach($target in $targets){
        $orderMatches=@(for($i=0;$i-lt$orders.cells.Count;$i++){if((Cell $orders.cells[$i] 2)-eq$target.order){[pscustomobject]@{row=$orders.row_indices[$i];cells=$orders.cells[$i]}}})
        if($orderMatches.Count-ne1){throw '无法唯一定位订单'}
        $order=$orderMatches[0]
        $dateText=[string](Cell $order.cells 0)
        $date=[datetime]::MinValue;$serial=0.0
        if([double]::TryParse($dateText,[ref]$serial) -and $serial-gt30000){$date=[datetime]::FromOADate($serial)}elseif(-not[datetime]::TryParse($dateText,[ref]$date)){throw '订单日期无法识别'}
        if($date.ToString('yyyy-MM-dd')-ne$target.date -or (Cell $order.cells 3)-eq'已取消' -or (Cell $order.cells 4)-ne'自然'){throw '订单日期、类型或状态与请求不一致'}
        if(@(Get-RefundEventsForOrder $ProjectRoot $target.order).Count){throw '订单存在售后记录，需另核实净售出数量'}
        if(-not(Cell $order.cells 10 'formula') -or -not(Cell $order.cells 11 'formula')){throw '订单汇总成本和利润缺少公式，未写入'}
        $matches=@();$orderCost=0.0
        for($i=0;$i-lt$items.cells.Count;$i++){
            $item=$items.cells[$i];if((Cell $item 2)-ne$target.order){continue}
            if((Cell $item 6)-eq$target.title -and (Cell $item 7)-eq$target.spec){
                if((Cell $item 5) -or (Cell $item 16)){throw '商品已绑定货号或ID，须重新核实'}
                $qty=Number (Cell $item 8);if($qty-ne1){throw '指定规格数量不是1，未写入'}
                $old=Cell $item 11
                if($null-ne$old -and [string]$old-ne'' -and [math]::Abs((Number $old)-$target.cost)-gt0.0001){throw '原单位成本与预期不符，未写入'}
                $matches+= [int]$items.row_indices[$i]
                $orderCost+=$target.cost*$qty
            }else{
                # Freeze every other item snapshot; do not use today's SKU costs.
                $orderCost+=Number (Cell $item 12)
            }
        }
        if($matches.Count-ne1){throw "商品名和规格无法唯一定位：$($target.order)"}
        foreach($gift in $gifts.cells){
            if((Cell $gift 3)-ne$target.order -or (Cell $gift 2)-in@('待同步','订单已取消待处理')){continue}
            if((Cell $gift 8)-eq'成本待核对'){throw '赠品成本未确认'}
            $orderCost+=Number (Cell $gift 7)
        }
        $income=Number (Cell $order.cells 7);$delivery=Number (Cell $order.cells 9)
        $plans+=[pscustomobject]@{order_no=$target.order;date=$target.date;title=$target.title;unit_cost=$target.cost;item_row=$matches[0];summary_row=$order.row;income=$income;delivery=$delivery;order_cost=[math]::Round($orderCost,4);profit=[math]::Round($income-$delivery-$orderCost,4)}
    }
    if(-not$Apply){[ordered]@{ok=$true;preview=$true;plans=$plans;writes=0}|ConvertTo-Json -Depth 6;exit 0}
    foreach($plan in $plans){
        $r=$plan.item_row;$range="L${r}:N${r}"
        Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '订单商品' $range $cli
        $matrix=,@(@{value=$plan.unit_cost;cell_styles=@{number_format='0.####'}},@{formula="=I${r}*L${r}"},@{formula="=IF(K${r}=`"`",I${r}*J${r},K${r})-M${r}"})
        $payload=ConvertTo-Json -InputObject $matrix -Depth 8 -Compress
        $null=Invoke-RepairCli @('sheets','+cells-set','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','订单商品','--range',$range,'--cells','-','--format','json') $payload
        Verify-RepairFormulas $sourceToken '订单商品' $range
        $r=$plan.summary_row
        Verify-RepairFormulas $sourceToken '订单汇总' "K${r}:L${r}"
        $readback=Read-RepairRows $sourceToken '订单汇总' "K${r}:L${r}"
        if([math]::Abs((Number (Cell $readback.cells[0] 0))-$plan.order_cost)-gt0.0001 -or [math]::Abs((Number (Cell $readback.cells[0] 1))-$plan.profit)-gt0.0001){throw '订单成本/利润回读不符'}
    }
    $raw=& 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -NoProfile -ExecutionPolicy Bypass -File (Join-Path $agentRoot 'sync_daily_ledger.ps1') -ProjectRoot $ProjectRoot -MonthFilter '2026-10,2026-11' | Out-String
    if($LASTEXITCODE-ne0){throw "对应月份简账刷新失败：$raw"}
    $ledger=$raw|ConvertFrom-Json;if(-not$ledger.ok){throw '简账刷新未成功'}
    foreach($plan in $plans){
        $sheet="$([int]$plan.date.Substring(5,2))月收入"
        $incomeRows=Read-RepairRows $ledgerToken $sheet 'A2:K1001'
        $matches=@($incomeRows.cells|Where-Object{(Cell $_ 1)-eq$plan.order_no})
        if($matches.Count-ne1 -or [math]::Abs((Number (Cell $matches[0] 4))-$plan.order_cost)-gt0.0001 -or [math]::Abs((Number (Cell $matches[0] 6))-$plan.profit)-gt0.0001){throw '收入表成本或利润回读不符'}
    }
    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction;$closed=$true
    [ordered]@{ok=$true;plans=$plans;sku_master_changed=$false;inventory_changed=$false;refreshed_months=@('2026-10','2026-11')}|ConvertTo-Json -Depth 6
}catch{
    $reason=$_.Exception.Message
    if($null-ne$transaction -and -not$closed){$rollback=Invoke-FinanceTransactionRollback $transaction @{source=$sourceToken;ledger=$ledgerToken} $cli $reason;Close-FinanceTransaction $transaction;$closed=$true;if(-not$rollback.ok){throw "$reason；回滚待处理"}}
    throw $reason
}finally{if($null-ne$transaction -and -not$closed){Close-FinanceTransaction $transaction}}
