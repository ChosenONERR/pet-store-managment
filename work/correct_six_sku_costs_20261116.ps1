param([switch]$Apply,[string]$ProjectRoot=(Split-Path $PSScriptRoot -Parent))
# One-off owner-authorized correction. Default read-only; do not rerun blindly.
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$OutputEncoding=[Text.UTF8Encoding]::new($false)
$cli=Join-Path $ProjectRoot 'agent/lark_cli_noexit.ps1'
$stateRoot=Join-Path $ProjectRoot 'agent/state'
. (Join-Path $ProjectRoot 'agent/finance_transaction.ps1')
$sourceMeta=Get-Content -LiteralPath (Join-Path $ProjectRoot 'work/simple_finance_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$ledgerMeta=Get-Content -LiteralPath (Join-Path $ProjectRoot 'work/daily_ledger_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$sourceToken=([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken=([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$prices=[ordered]@{kw1090=17.78;kw1107=17.78;kw1108=22.23;kw1091=22.23;kw1048=6.12;kw1114=24.44}
function Invoke-CorrectionCli([string[]]$Arguments,[string]$Payload){
    $global:LASTEXITCODE=0
    $raw=if($PSBoundParameters.ContainsKey('Payload')){$Payload|& $cli @Arguments|Out-String}else{& $cli @Arguments|Out-String}
    if($LASTEXITCODE-ne0){throw '飞书调用失败，未输出授权或完整响应'}
    $r=$raw|ConvertFrom-Json
    if(-not$r.ok){throw '飞书返回失败'}
    return $r
}
function Read-CorrectionRows($Token,$Sheet,$Range){
    $r=Invoke-CorrectionCli @('sheets','+cells-get','--as','user','--spreadsheet-token',$Token,'--sheet-name',$Sheet,'--range',$Range,'--include','value,formula','--max-chars','6000000','--format','json')
    if($r.data.has_more -or $r.data.ranges[0].truncated){throw "读取不完整：$Sheet"}
    return $r.data.ranges[0]
}
function C($Row,[int]$Index,[string]$Field='value'){
    if($null-eq$Row -or $Index-ge@($Row).Count){return $null}
    $cell=@($Row)[$Index];if($null-eq$cell){return $null}
    $p=$cell.PSObject.Properties[$Field];if($null-ne$p){return $p.Value}
    return $null
}
function Date-Key($Value){
    if($null-eq$Value -or [string]$Value-eq''){return ''}
    $n=0.0;$d=[datetime]::MinValue
    if([double]::TryParse([string]$Value,[ref]$n)-and$n-gt30000){return [datetime]::FromOADate($n).ToString('yyyy-MM-dd')}
    if([datetime]::TryParse([string]$Value,[ref]$d)){return $d.ToString('yyyy-MM-dd')}
    throw '日期无法识别'
}
function In-Scope($Value){return (Date-Key $Value)-match '^2026-(10|11)-'}
function Code($Value){return ([string]$Value).Trim().ToLowerInvariant()}
function Num($Value){
    $n=0.0
    if($null-eq$Value -or [string]$Value-eq'' -or -not[double]::TryParse([string]$Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)-or[double]::IsNaN($n)-or[double]::IsInfinity($n)){throw '数量或成本不明确，停止纠正'}
    return $n
}
function Near($Actual,[double]$Expected,[string]$Label){if([math]::Abs((Num $Actual)-$Expected)-gt0.0001){throw "回读不一致：$Label"}}
function Get-CorrectionPlan($Profiles,$Sales,$OfflineSales,$Orders){
    $expected=[ordered]@{kw1090='51185971712';kw1107='51185971714';kw1108='51189679003';kw1091='51189679004';kw1048='51186449393';kw1114='52586891723'}
    $selected=@();$byId=@{};$byCode=@{}
    foreach($code in $expected.Keys){
        $match=@($Profiles|Where-Object{$_.code-eq$code-and$_.product_id-eq$expected[$code]})
        if($match.Count-ne1){throw "无法唯一定位指定商品：$code"}
        $selected+=$match[0];$byId[$expected[$code]]=$match[0];$byCode[$code]=$match[0]
    }
    $lines=@();$orderPlans=@{};$skipped=@()
    foreach($s in $Sales){
        if(-not(In-Scope $s.date)){throw '超出已授权月份'}
        if($s.product_id-eq'51185971723'-and$s.code-eq'kw1107'){continue} # Unrelated 2kg dog food.
        $profile=if($s.product_id){$byId[$s.product_id]}else{$byCode[$s.code]}
        if($null-eq$profile){throw "商品身份不明确：$($s.order_no)"}
        if($s.code-ne$profile.code){throw '商品ID与货号不一致'}
        if(-not$s.product_id-and(-not([string]$s.title).Contains([string]$profile.title)-or-not([string]$s.title).Contains([string]$profile.spec))){throw '缺少商品ID且名称/规格不能核实'}
        if($s.status-eq'已取消'-or$s.type-eq'干预'){$skipped+=$s;continue}
        if($s.type-ne'自然'-or$s.status-eq'unknown'){throw '订单状态或类型不明确'}
        $qty=Num $s.quantity;if($qty-le0){throw '销售数量非正数'}
        $newCost=[math]::Round($qty*$profile.new_cost,4);$delta=$newCost-(Num $s.old_cost)
        $lines+=[pscustomobject]@{row=[int]$s.row;order_no=$s.order_no;date=$s.date;daily_no=$s.daily_no;code=$s.code;product_id=$profile.product_id;quantity=$qty;old_unit=Num $s.old_unit;old_cost=Num $s.old_cost;new_unit=$profile.new_cost;new_cost=$newCost;cost_delta=$delta;changed=([math]::Abs($delta)-gt0.0001-or[math]::Abs((Num $s.old_unit)-$profile.new_cost)-gt0.0001)}
        if(-not$orderPlans.ContainsKey($s.order_no)){
            $indices=@(for($j=1;$j-lt$Orders.cells.Count;$j++){if([string](C $Orders.cells[$j] 2)-eq$s.order_no){$j}})
            if($indices.Count-ne1){throw '订单汇总无法唯一定位'}
            $index=$indices[0];$o=$Orders.cells[$index]
            if((Date-Key (C $o 0))-ne$s.date-or(C $o 3)-eq'已取消'-or(C $o 4)-ne'自然'){throw '订单汇总状态不一致'}
            if(-not(C $o 10 'formula')-or-not(C $o 11 'formula')){throw '订单汇总缺少成本/利润公式'}
            $orderPlans[$s.order_no]=[pscustomobject]@{order_no=$s.order_no;date=$s.date;daily_no=$s.daily_no;row=[int]$Orders.row_indices[$index];income=Num (C $o 7);delivery=Num (C $o 9);old_cost=Num (C $o 10);old_profit=Num (C $o 11);cost_delta=0.0;new_cost=0.0;new_profit=0.0}
        }
        $orderPlans[$s.order_no].cost_delta+=$delta
    }
    foreach($p in $orderPlans.Values){$p.new_cost=[math]::Round($p.old_cost+$p.cost_delta,4);$p.new_profit=[math]::Round($p.income-$p.delivery-$p.new_cost,4)}
    $offlinePlans=@()
    foreach($o in $OfflineSales){
        if(-not(In-Scope $o.date)){throw '线下单超出月份范围'}
        $parts=@([regex]::Matches($o.items,'(?i)\b((?:kw|csx)\d+)\s*[×x*]\s*(\d+(?:\.\d+)?)'))
        if($parts.Count-ne1-or-not$byCode.ContainsKey((Code $parts[0].Groups[1].Value))){throw '线下混合商品缺少分项历史成本，不盲目重算其他商品'}
        $code=Code $parts[0].Groups[1].Value;$qty=Num $parts[0].Groups[2].Value
        if($qty-le0){throw '线下数量非正数'}
        $newCost=[math]::Round($qty*$byCode[$code].new_cost,2)
        $offlinePlans+=[pscustomobject]@{row=$o.row;date=$o.date;order_no=$o.order_no;code=$code;quantity=$qty;new_unit=$byCode[$code].new_cost;income=Num $o.income;delivery=Num $o.delivery;old_cost=Num $o.old_cost;new_cost=$newCost;old_profit=Num $o.old_profit;new_profit=[math]::Round((Num $o.income)-(Num $o.delivery)-$newCost,2);cost_delta=$newCost-(Num $o.old_cost);note=$o.note}
    }
    return [pscustomobject]@{profiles=$selected;lines=$lines;orders=@($orderPlans.Values|Sort-Object date,order_no);offline=$offlinePlans;skipped=$skipped}
}
function Write-CorrectionBatch($Token,$Writes){
    $payload=ConvertTo-Json -InputObject @($Writes) -Depth 12 -Compress
    $null=Invoke-CorrectionCli @('sheets','+cells-set','--as','user','--spreadsheet-token',$Token,'--writes','-','--format','json') $payload
}
function Get-LedgerCostCorrection($Row,$Plan){
    Near (C $Row 3) $Plan.income '原月表实收';Near (C $Row 5) $Plan.delivery '原月表运费'
    if(-not(C $Row 6 'formula')){throw '原收入表缺少利润公式'}
    $cost=(Num (C $Row 4))+$Plan.cost_delta
    return [pscustomobject]@{cost=$cost;profit=$Plan.income-$Plan.delivery-$cost}
}
function Verify-Correction($Token,$Sheet,$Range){
    foreach($area in @($Range-split','|Where-Object{$_})){
        $r=Invoke-CorrectionCli @('sheets','+formula-verify','--as','user','--spreadsheet-token',$Token,'--sheet-name',$Sheet,'--range',$area,'--exit-on-error','--format','json')
        if($r.data.status-ne'success'){throw "公式校验未通过：$Sheet"}
    }
}
$transaction=$null;$closed=$false
try{
if($Apply){$transaction=New-FinanceTransaction $stateRoot '店主确认六项进货成本及2026-10/11销售纠正' $null -AcquireLock -WorkbookTokens @{source=$sourceToken;ledger=$ledgerToken} -LarkCli $cli}
$master=Read-CorrectionRows $sourceToken 'SKU成本主表' 'A1:Z3000'
$orders=Read-CorrectionRows $sourceToken '订单汇总' 'A1:Q2001'
$items=Read-CorrectionRows $sourceToken '订单商品' 'A1:Q4001'
$offline=Read-CorrectionRows $sourceToken '线下订单' 'A1:J1001'
$gifts=Read-CorrectionRows $sourceToken '赠品记录' 'A1:L2001'
$headerChecks=@(@{data=$master;indices=@{0='当前KW';5='基准单位成本';6='当前移动平均成本';10='成本核对状态';16='商品ID（美团SKU ID）'}},@{data=$items;indices=@{2='平台订单号';5='货号';8='数量';11='单位成本';12='商品成本';16='商品ID（美团SKU ID）'}},@{data=$orders;indices=@{2='平台订单号';7='商家预计收入';9='配送费';10='商品成本';11='订单毛利'}},@{data=$offline;indices=@{2='订单编号';3='实际收入';4='配送费';5='货物成本';7='售出货品（KW）'}})
foreach($check in $headerChecks){foreach($j in $check.indices.Keys){if([string](C $check.data.cells[0] ([int]$j))-ne$check.indices[$j]){throw '表头与安全写入位置不一致'}}}
$profiles=@();$ids=@{}
for($i=1;$i-lt$master.cells.Count;$i++){
    $r=$master.cells[$i];$code=Code (C $r 0)
    if(-not$prices.Contains($code)){continue}
    $id=[string](C $r 16);$ids[$id]=$code
    $profiles+=[pscustomobject]@{code=$code;row=$master.row_indices[$i];product_id=$id;title=C $r 2;spec=C $r 18;old_baseline=C $r 5;old_moving=C $r 6;status=C $r 10;new_cost=$prices[$code]}
}
$sales=@()
for($i=1;$i-lt$items.cells.Count;$i++){
    $r=$items.cells[$i];$code=Code (C $r 5);$id=[string](C $r 16)
    if(-not(In-Scope (C $r 0)) -or (-not$prices.Contains($code)-and-not$ids.ContainsKey($id))){continue}
    $orderNo=[string](C $r 2);$o=@($orders.cells|Where-Object{[string](C $_ 2)-eq$orderNo})
    $sales+=[pscustomobject]@{row=$items.row_indices[$i];date=Date-Key (C $r 0);order_no=$orderNo;code=$code;product_id=$id;title=C $r 6;spec=C $r 7;quantity=C $r 8;old_unit=C $r 11;old_cost=C $r 12;status=$(if($o.Count-eq1){C $o[0] 3}else{'unknown'});type=$(if($o.Count-eq1){C $o[0] 4}else{'unknown'});daily_no=$(if($o.Count-eq1){C $o[0] 16}else{''})}
}
$offlineSales=@();$giftSales=@()
for($i=1;$i-lt$offline.cells.Count;$i++){
    $r=$offline.cells[$i]
    if((In-Scope (C $r 0))-and([string](C $r 7)-match '(?i)\b(?:kw1090|kw1107|kw1108|kw1091|kw1048|kw1114)\b')){
        $offlineSales+=[pscustomobject]@{row=$offline.row_indices[$i];date=Date-Key (C $r 0);order_no=C $r 2;items=C $r 7;income=C $r 3;delivery=C $r 4;old_cost=C $r 5;old_profit=C $r 6;note=C $r 9}
    }
}
for($i=1;$i-lt$gifts.cells.Count;$i++){
    $r=$gifts.cells[$i]
    if((In-Scope (C $r 0))-and([string](C $r 4)-match '(?i)(?:kw1090|kw1107|kw1108|kw1091|kw1048|kw1114)\b')){
        $giftSales+=[pscustomobject]@{row=$gifts.row_indices[$i];cells=$r}
    }
}
$plan=Get-CorrectionPlan $profiles $sales $offlineSales $orders
if($giftSales.Count){throw '这些商品存在赠品记录，须先核实赠品成本范围'}
. (Join-Path $ProjectRoot 'agent/refund_adjustments.ps1')
. (Join-Path $ProjectRoot 'agent/order_actual_overrides.ps1')
foreach($o in $plan.orders){if(@(Get-RefundEventsForOrder $ProjectRoot $o.order_no).Count-or@(Get-OrderActualOverride $ProjectRoot $o.order_no).Count){throw '指定商品订单有售后/履约覆盖，须先核实实际数量'}}
if(-not$Apply){[ordered]@{ok=$true;preview=$true;plan=$plan;writes=0}|ConvertTo-Json -Depth 10;exit 0}
$marker='店主按2026-09-09进货单纠正六项成本（2026-11-16），仅重算2026-10/11指定商品'
$writes=@();$changedOrders=@($plan.orders|Where-Object{[math]::Abs($_.cost_delta)-gt0.0001})
foreach($p in $plan.profiles){
    $r=[int]$p.row;$raw=@($master.cells[[array]::IndexOf(@($master.row_indices),$r)])
    $same=([math]::Abs((Num $p.old_baseline)-$p.new_cost)-le0.0001-and[math]::Abs((Num $p.old_moving)-$p.new_cost)-le0.0001)
    if($same-and([string](C $raw 14)).Contains($marker)){continue}
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' "F${r}:O${r}" $cli
    $status=if($p.status-match'当前KW冲突'){'当前KW冲突待核对（成本已确认）'}else{'用户确认'}
    $oldNote=[string](C $raw 14);$note=(@($oldNote,$marker)|Where-Object{$_}|Select-Object -Unique)-join'；'
    $writes+=@{sheet_name='SKU成本主表';range="F${r}:G${r}";cells=,@(@{value=$p.new_cost;cell_styles=@{number_format='0.####'}},@{value=$p.new_cost;cell_styles=@{number_format='0.####'}})}
    $writes+=@{sheet_name='SKU成本主表';range="K${r}:K${r}";cells=,@(@{value=$status})}
    $writes+=@{sheet_name='SKU成本主表';range="M${r}:O${r}";cells=,@(@{value=([datetime]::Today-[datetime]'1899-12-30').TotalDays},@{value='店主确认含分摊进货运费的销售单位成本'},@{value=$note})}
}
foreach($p in @($plan.lines|Where-Object{$_.changed})){
    $r=$p.row;Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '订单商品' "L${r}:N${r}" $cli
    $writes+=@{sheet_name='订单商品';range="L${r}:N${r}";cells=,@(@{value=$p.new_unit;cell_styles=@{number_format='0.####'}},@{formula="=I${r}*L${r}"},@{formula="=IF(K${r}=`"`",I${r}*J${r},K${r})-M${r}"})}
}
foreach($p in $plan.offline){
    if([math]::Abs($p.cost_delta)-le0.0001){continue}
    $r=$p.row;Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '线下订单' "F${r}:J${r}" $cli
    $note=(@($p.note,$marker)|Where-Object{$_}|Select-Object -Unique)-join'；'
    $writes+=@{sheet_name='线下订单';range="F${r}:G${r}";cells=,@(@{value=$p.new_cost;cell_styles=@{number_format='0.####'}},@{formula="=D${r}-E${r}-F${r}"})}
    $writes+=@{sheet_name='线下订单';range="J${r}:J${r}";cells=,@(@{value=$note})}
}
$oldLedger=@{}
foreach($month in @('10','11')){$oldLedger[$month]=Read-CorrectionRows $ledgerToken "${month}月收入" 'A1:K1001'}
if($writes.Count){Write-CorrectionBatch $sourceToken $writes}
Verify-Correction $sourceToken '订单商品' ((@($plan.lines|Where-Object{$_.changed}|ForEach-Object{"L$($_.row):N$($_.row)"}))-join',')
Verify-Correction $sourceToken '订单汇总' ((@($plan.orders|ForEach-Object{"K$($_.row):L$($_.row)"}))-join',')
foreach($p in $plan.offline){Verify-Correction $sourceToken '线下订单' "F$($p.row):G$($p.row)"}
$newOrders=Read-CorrectionRows $sourceToken '订单汇总' 'A1:Q2001'
$newItems=Read-CorrectionRows $sourceToken '订单商品' 'A1:Q4001'
$newMaster=Read-CorrectionRows $sourceToken 'SKU成本主表' 'A1:Z3000'
$newOffline=Read-CorrectionRows $sourceToken '线下订单' 'A1:J1001'
$lineMap=@{};foreach($p in $plan.lines){$lineMap[[int]$p.row]=$p}
$profileMap=@{};foreach($p in $plan.profiles){$profileMap[[int]$p.row]=$p}
$offlineMap=@{};foreach($p in $plan.offline){$offlineMap[[string]$p.order_no]=$p}
for($i=1;$i-lt$master.cells.Count;$i++){
    $r=[int]$master.row_indices[$i];$match=@($newMaster.cells[[array]::IndexOf(@($newMaster.row_indices),$r)])
    if($profileMap.ContainsKey($r)){Near (C $match 5) $profileMap[$r].new_cost "主表$r";Near (C $match 6) $profileMap[$r].new_cost "主表$r"}
    for($j=0;$j-lt26;$j++){if($profileMap.ContainsKey($r)-and$j-in@(5,6,10,12,13,14)){continue};if([string](C $master.cells[$i] $j)-ne[string](C $match $j)){throw "主表非授权字段改变：$r/$j"}}
}
for($i=1;$i-lt$items.cells.Count;$i++){
    $r=[int]$items.row_indices[$i];$match=@($newItems.cells[[array]::IndexOf(@($newItems.row_indices),$r)])
    if($lineMap.ContainsKey($r)){Near (C $match 11) $lineMap[$r].new_unit "商品$r";Near (C $match 12) $lineMap[$r].new_cost "商品$r"}
    for($j=0;$j-lt17;$j++){if($lineMap.ContainsKey($r)-and$j-in@(11,12,13)){continue};if([string](C $items.cells[$i] $j)-ne[string](C $match $j)){throw "非授权商品字段改变：$r/$j"}}
}
foreach($p in $plan.orders){
    $match=@($newOrders.cells|Where-Object{[string](C $_ 2)-eq$p.order_no});if($match.Count-ne1){throw '订单回读缺失'}
    Near (C $match[0] 7) $p.income '订单实收';Near (C $match[0] 9) $p.delivery '订单配送费';Near (C $match[0] 10) $p.new_cost '订单成本';Near (C $match[0] 11) $p.new_profit '订单利润'
}
foreach($p in $plan.offline){$match=@($newOffline.cells|Where-Object{[string](C $_ 2)-eq$p.order_no});if($match.Count-ne1){throw '线下单回读缺失'};Near (C $match[0] 3) $p.income '线下实收';Near (C $match[0] 4) $p.delivery '线下配送费';Near (C $match[0] 5) $p.new_cost '线下成本';Near (C $match[0] 6) $p.new_profit '线下利润'}
$monthResults=@();$allPlans=@($plan.orders)+@($plan.offline);$ledgerWrites=@()
# Source summary values are displayed to cents, while the month ledger keeps
# frozen fractional-cent costs. Apply only the confirmed item delta to each
# existing ledger cost; never regenerate unrelated rows or round their costs.
foreach($p in $allPlans){
    $month=$p.date.Substring(5,2);$data=$oldLedger[$month]
    $indices=@(for($i=1;$i-lt$data.cells.Count;$i++){if([string](C $data.cells[$i] 1)-eq$p.order_no){$i}})
    if($indices.Count-ne1){throw '目标订单在收入表中无法唯一定位'}
    $index=$indices[0];$row=$data.cells[$index];$r=[int]$data.row_indices[$index]
    $amounts=Get-LedgerCostCorrection $row $p
    $newLedgerCost=$amounts.cost;$newLedgerProfit=$amounts.profit
    $p|Add-Member NoteProperty ledger_row $r
    $p|Add-Member NoteProperty ledger_new_cost $newLedgerCost
    $p|Add-Member NoteProperty ledger_new_profit $newLedgerProfit
    if([math]::Abs($p.cost_delta)-le0.0001){continue}
    Register-FinanceRangeBeforeImage $transaction 'ledger' $ledgerToken "${month}月收入" "E${r}:E${r}" $cli
    $ledgerWrites+=@{sheet_name="${month}月收入";range="E${r}:E${r}";cells=,@(@{value=$newLedgerCost})}
}
if($ledgerWrites.Count){Write-CorrectionBatch $ledgerToken $ledgerWrites}
foreach($month in @('10','11')){
    $updated=Read-CorrectionRows $ledgerToken "${month}月收入" 'A1:K1001'
    $monthPlan=@($allPlans|Where-Object{$_.date.Substring(5,2)-eq$month});$targets=@{};foreach($p in $monthPlan){$targets[$p.order_no]=$p}
    foreach($p in $monthPlan){$m=@($updated.cells|Where-Object{[string](C $_ 1)-eq$p.order_no});if($m.Count-ne1){throw '月份收入记录不唯一'};Near (C $m[0] 3) $p.income '月表收入';Near (C $m[0] 5) $p.delivery '月表运费';Near (C $m[0] 4) $p.ledger_new_cost '月表成本';Near (C $m[0] 6) $p.ledger_new_profit '月表利润'}
    foreach($before in $oldLedger[$month].cells){$number=[string](C $before 1);if(-not$number-or$number-eq'订单编号'-or$targets.ContainsKey($number)){continue};$m=@($updated.cells|Where-Object{[string](C $_ 1)-eq$number});if($m.Count-ne1){throw '非目标订单在月表中缺失'};foreach($j in @(3,4,5,6)){if([string](C $before $j)-ne[string](C $m[0] $j)){throw "非目标订单金额改变：$number"}}}
    Verify-Correction $ledgerToken "${month}月收入" 'D2:J1001'
    $monthResults+=[pscustomobject]@{month="2026-$month";related_orders=$monthPlan.Count;cost_delta=[math]::Round(($monthPlan|Measure-Object cost_delta -Sum).Sum,4);profit_delta=-[math]::Round(($monthPlan|Measure-Object cost_delta -Sum).Sum,4)}
}
Verify-Correction $sourceToken '驾驶舱' 'C4:D7'
Complete-FinanceTransaction $transaction;Close-FinanceTransaction $transaction;$closed=$true
$result=[ordered]@{ok=$true;transaction=$transaction.id;plan=$plan;months=$monthResults;source_write_ranges=$writes.Count;ledger_write_ranges=$ledgerWrites.Count;formula_verification='success';income_changed=$false;delivery_changed=$false;inventory_changed=$false}
$reportPath=Join-Path $stateRoot ("six_sku_cost_correction_20261116_$($transaction.id).json")
[IO.File]::WriteAllText($reportPath,($result|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
$result|ConvertTo-Json -Depth 10
}catch{
    $reason=$_.Exception.Message
    if($transaction-and-not$closed){$rollback=Invoke-FinanceTransactionRollback $transaction @{source=$sourceToken;ledger=$ledgerToken} $cli $reason;Close-FinanceTransaction $transaction;$closed=$true;if(-not$rollback.ok){throw "$reason；回滚待处理"}}
    throw $reason
}finally{if($transaction-and-not$closed){Close-FinanceTransaction $transaction}}
