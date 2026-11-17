param([switch]$Apply,[string]$ProjectRoot=(Split-Path $PSScriptRoot -Parent))
# Owner-confirmed one-off recovery. Do not rerun blindly. No platform writes.
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false);$OutputEncoding=[Text.UTF8Encoding]::new($false)
$cli=Join-Path $ProjectRoot 'agent/lark_cli_noexit.ps1';$stateRoot=Join-Path $ProjectRoot 'agent/state'
. (Join-Path $ProjectRoot 'agent/finance_transaction.ps1')
. (Join-Path $ProjectRoot 'agent/expense_accounting.ps1')
$s=Get-Content -LiteralPath (Join-Path $ProjectRoot 'work/simple_finance_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$l=Get-Content -LiteralPath (Join-Path $ProjectRoot 'work/daily_ledger_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$source=([uri]$s.workbook_url).Segments[-1].Trim('/');$ledger=([uri]$l.workbook_url).Segments[-1].Trim('/')
$id='3402391430051537057';$other='3202391603486436114';$day='2026-11-16'
function Invoke-RecoveryCli([string[]]$Arguments,[string]$Payload){
    $global:LASTEXITCODE=0;$raw=if($PSBoundParameters.ContainsKey('Payload')){$Payload|& $cli @Arguments|Out-String}else{& $cli @Arguments|Out-String}
    if($LASTEXITCODE-ne0){throw '飞书调用失败，未输出凭据或完整响应'}
    $r=$raw|ConvertFrom-Json;if(-not$r.ok){throw '飞书返回失败'};return $r
}
function Read-Rows($Token,$Sheet,$Range){
    $r=Invoke-RecoveryCli @('sheets','+cells-get','--as','user','--spreadsheet-token',$Token,'--sheet-name',$Sheet,'--range',$Range,'--include','value,formula','--max-chars','4000000','--format','json')
    if($r.data.has_more-or$r.data.ranges[0].truncated){throw "读取不完整：$Sheet"};return $r.data.ranges[0]
}
function Val($Row,[int]$Index){if($null-ne$Row-and$Index-lt@($Row).Count){return @($Row)[$Index].value};return $null}
function Number($Value){$n=0.0;if([string]::IsNullOrWhiteSpace([string]$Value)-or-not[double]::TryParse([string]$Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)-or[double]::IsNaN($n)-or[double]::IsInfinity($n)){throw '金额未知，停止恢复'};return $n}
function Near($Actual,[double]$Expected,$Label){if([math]::Abs((Number $Actual)-$Expected)-gt0.0001){throw "核对失败：$Label"}}
function DateKey($Value){$n=0.0;if([double]::TryParse([string]$Value,[ref]$n)-and$n-gt30000){return [datetime]::FromOADate($n).ToString('yyyy-MM-dd')};return ([datetime]$Value).ToString('yyyy-MM-dd')}
function Write-Batch($Token,$Writes){$payload=ConvertTo-Json -InputObject @($Writes) -Depth 12 -Compress;$null=Invoke-RecoveryCli @('sheets','+cells-set','--as','user','--spreadsheet-token',$Token,'--writes','-','--format','json') $payload}
function Verify($Token,$Sheet,$Range){$r=Invoke-RecoveryCli @('sheets','+formula-verify','--as','user','--spreadsheet-token',$Token,'--sheet-name',$Sheet,'--range',$Range,'--exit-on-error','--format','json');if($r.data.status-ne'success'){throw "公式未通过：$Sheet"}}
function Day-Profit($Income,$Expenses){
    $gross=0.0;$count=0
    foreach($r in $Income.cells){if(-not(Val $r 1)-or(Val $r 1)-eq'订单编号'){continue};if((DateKey (Val $r 2))-ne$day){continue};$gross+=Number (Val $r 6);$count++}
    $entries=@(foreach($r in $Expenses.cells){if(-not(Val $r 1)){continue};$date=Val $r 0;if(-not$date-or[ string]$date-eq'日期'){continue};if((DateKey $date)-ne$day){continue};[pscustomobject]@{project=Val $r 1;fee=Val $r 2;order_no=Val $r 3;source=Val $r 4;note=Val $r 5}})
    $fees=Get-PetstarExpenseSummary -Entries $entries
    return [pscustomobject]@{date=$day;natural_and_offline_orders=$count;gross=[math]::Round($gross,4);expenses=$fees.total;net=[math]::Round($gross-$fees.total,4);pending=$fees.pending}
}
$transaction=$null;$closed=$false
try{
    if($Apply){$transaction=New-FinanceTransaction $stateRoot '店主确认11月16日3号与5号均为干预单' $null -AcquireLock -WorkbookTokens @{source=$source;ledger=$ledger} -LarkCli $cli}
    $orders=Read-Rows $source '订单汇总' 'A510:Q525';$interventions=Read-Rows $source '干预订单' 'A2:M1001'
    $items=Read-Rows $source '订单商品' 'A1638:Q1660'
    $income=Read-Rows $ledger '11月收入' 'A1:K1001';$expenses=Read-Rows $ledger '11月支出' 'A1:F1001'
    $target=@($orders.cells|Where-Object{(Val $_ 2)-eq$id});$fifth=@($interventions.cells|Where-Object{(Val $_ 3)-eq$other})
    if($target.Count-ne1-or$fifth.Count-ne1-or(Val $fifth[0] 2)-ne'黄慧娴'){throw '无法核实3号/5号登记身份'}
    if((DateKey (Val $target[0] 0))-ne$day-or(Val $target[0] 16)-ne'美团3号单'){throw '目标日期或当日单号不符'}
    Near (Val $target[0] 7) 92.17 '3号实收';Near (Val $target[0] 8) 5.87 '3号平台费用';Near (Val $target[0] 9) 0 '3号自配送'
    $active=@(for($i=0;$i-lt$interventions.cells.Count;$i++){if((Val $interventions.cells[$i] 3)-eq$id){$i}})
    if($active.Count-gt1){throw '3号干预记录重复'}
    $newRecord=$active.Count-eq0
    $slot=if($newRecord){@(for($i=0;$i-lt$interventions.cells.Count;$i++){if(-not(Val $interventions.cells[$i] 3)){$i;break}})[0]}else{$active[0]}
    if($null-eq$slot){throw '干预表无空位'};$r=[int]$interventions.row_indices[$slot]
    $sold=@();foreach($row in $items.cells){if((Val $row 2)-ne$id){continue};Near (Val $row 11) 0 '3号商品单位成本';Near (Val $row 12) 0 '3号商品成本';$sold+="$(Val $row 5)×$(Number (Val $row 8))"}
    if($sold.Count-ne7){throw '3号商品明细数量异常'}
    $before=Day-Profit $income $expenses
    if(-not$Apply){[ordered]@{ok=$true;preview=$true;order_no=$id;person='姐姐';new_intervention_row=$r;keep_order=$other;keep_person='黄慧娴';platform_fee=5.87;delivery_fee=0;before=$before;writes=0}|ConvertTo-Json -Depth 6;exit 0}
    if($newRecord){
        Register-FinanceRangeBeforeImage $transaction 'source' $source '干预订单' "A${r}:M${r}" $cli
        $note='店主2026-11-17明确确认：11月16日3号单和5号单都是干预单；本单补单人姐姐，5号单黄慧娴。保留原撤销审计，恢复本单有效登记。自己配送0元，库存已由店主手工恢复；不执行平台库存操作。'
        $row=@(@{value=([datetime]$day).ToOADate()},@{value='美团3号单'},@{value='姐姐'},@{value=$id},@{value=Number (Val $target[0] 6)},@{value='自配送'},@{value=5.87},@{value=0},@{formula="=IF(D${r}=`"`",`"`",IFERROR(G${r}*1,0)+IFERROR(H${r}*1,0))"},@{value='已恢复'},@{value=''},@{value=($sold-join'、')},@{value=$note})
        Write-Batch $source @(@{sheet_name='干预订单';range="A${r}:M${r}";cells=,$row})
    }else{
        Near (Val $interventions.cells[$slot] 6) 5.87 '现有平台费用';Near (Val $interventions.cells[$slot] 7) 0 '现有配送费用'
        if((Val $interventions.cells[$slot] 2)-ne'姐姐'){
            Register-FinanceRangeBeforeImage $transaction 'source' $source '干预订单' "C${r}:C${r}" $cli
            Write-Batch $source @(@{sheet_name='干预订单';range="C${r}:C${r}";cells=,@(@{value='姐姐'})})
        }
    }
    Verify $source '干预订单' "A${r}:M${r}"
    $raw=& 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe' -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $ProjectRoot 'agent/sync_daily_ledger.ps1') -ProjectRoot $ProjectRoot -MonthFilter '2026-11'|Out-String
    if($LASTEXITCODE-ne0){throw '11月账表刷新失败'};$refreshed=$raw|ConvertFrom-Json;if(-not$refreshed.ok){throw '账表未返回成功'}
    $updated=Read-Rows $ledger '11月收入' 'A1:K1001';$precisionWrites=@()
    foreach($old in $income.cells){
        $n=[string](Val $old 1);if(-not$n-or$n-eq'订单编号'-or$n-eq$id){continue}
        $indices=@(for($i=0;$i-lt$updated.cells.Count;$i++){if((Val $updated.cells[$i] 1)-eq$n){$i}})
        if($indices.Count-ne1){throw "非目标订单缺失：$n"};$i=$indices[0];$now=$updated.cells[$i]
        Near (Val $now 3) (Number (Val $old 3)) '保留其他订单收入';Near (Val $now 5) (Number (Val $old 5)) '保留其他订单配送费'
        $oldCost=Val $old 4;$newCost=Val $now 4
        if([string]::IsNullOrWhiteSpace([string]$oldCost)){if(-not[string]::IsNullOrWhiteSpace([string]$newCost)){throw '非目标缺失成本被重算'};continue}
        if([math]::Abs((Number $newCost)-(Number $oldCost))-gt0.0000001){
            $lr=[int]$updated.row_indices[$i];Register-FinanceRangeBeforeImage $transaction 'ledger' $ledger '11月收入' "E${lr}:E${lr}" $cli
            $precisionWrites+=@{sheet_name='11月收入';range="E${lr}:E${lr}";cells=,@(@{value=Number $oldCost})}
        }
    }
    if($precisionWrites.Count){Write-Batch $ledger $precisionWrites}
    $updated=Read-Rows $ledger '11月收入' 'A1:K1001';$newExpenses=Read-Rows $ledger '11月支出' 'A1:F1001'
    if(@($updated.cells|Where-Object{(Val $_ 1)-in@($id,$other)}).Count){throw '干预单仍出现在自然收入表'}
    foreach($old in $income.cells){$n=[string](Val $old 1);if(-not$n-or$n-eq'订单编号'-or$n-eq$id){continue};$now=@($updated.cells|Where-Object{(Val $_ 1)-eq$n});if($now.Count-ne1){throw '非目标订单未保留'};foreach($j in @(3,4,5,6)){if([string](Val $old $j)-ne[string](Val $now[0] $j)){Near (Val $now[0] $j) (Number (Val $old $j)) '非目标金额'}}}
    foreach($n in @($id,$other)){$fee=@($newExpenses.cells|Where-Object{(Val $_ 3)-eq$n-and(Val $_ 1)-eq'补单平台扣款'});$deliveryFee=@($newExpenses.cells|Where-Object{(Val $_ 3)-eq$n-and(Val $_ 1)-eq'补单配送费'});if($fee.Count-ne1-or$deliveryFee.Count-ne1){throw '干预费用遗漏或重复'};Near (Val $fee[0] 2) $(if($n-eq$id){5.87}else{2.29}) '干预平台费用';Near (Val $deliveryFee[0] 2) 0 '干预配送费'}
    $final=Read-Rows $source '干预订单' 'A2:M1001'
    $a=@($final.cells|Where-Object{(Val $_ 3)-eq$id});$b=@($final.cells|Where-Object{(Val $_ 3)-eq$other})
    if($a.Count-ne1-or$b.Count-ne1-or(Val $a[0] 2)-ne'姐姐'-or($b[0]|ConvertTo-Json -Depth 5 -Compress)-ne($fifth[0]|ConvertTo-Json -Depth 5 -Compress)){throw '两笔干预登记核对失败'}
    $sourceReadback=Read-Rows $source '订单汇总' 'A515:Q517'
    foreach($row in $sourceReadback.cells){if((Val $row 2)-in@($id,$other)){if((Val $row 4)-ne'干预'){throw '源订单分类未更新'}}}
    $after=Day-Profit $updated $newExpenses
    $hadIncome=@($income.cells|Where-Object{(Val $_ 1)-eq$id}).Count-gt0
    if($hadIncome){Near $after.gross ($before.gross-92.17) '日毛利更正';Near $after.expenses ($before.expenses+5.87) '日支出更正';Near $after.net ($before.net-98.04) '日利润更正'}
    Verify $source '订单汇总' 'E515:L517';Verify $ledger '11月收入' 'D2:J1001';Verify $ledger '11月支出' 'A2:F1001'
    Complete-FinanceTransaction $transaction;Close-FinanceTransaction $transaction;$closed=$true
    $result=[ordered]@{ok=$true;transaction=$transaction.id;intervention_row=$r;restored_order=$id;person='姐姐';unchanged_order=$other;unchanged_person='黄慧娴';before=$before;after=$after;platform_fee=5.87;delivery_fee=0;preserved_precision_rows=$precisionWrites.Count;inventory_changed=$false;formula_verification='success'}
    [IO.File]::WriteAllText((Join-Path $stateRoot ('restore_two_interventions_20261117_'+$transaction.id+'.json')),($result|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    $result|ConvertTo-Json -Depth 8
}catch{
    $reason=$_.Exception.Message
    if($transaction-and-not$closed){$rollback=Invoke-FinanceTransactionRollback $transaction @{source=$source;ledger=$ledger} $cli $reason;Close-FinanceTransaction $transaction;$closed=$true;if(-not$rollback.ok){throw "$reason；回滚待处理"}}
    throw $reason
}finally{if($transaction-and-not$closed){Close-FinanceTransaction $transaction}}
