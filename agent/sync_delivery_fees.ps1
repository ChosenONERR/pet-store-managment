param(
    [string]$Date = ([datetime]::Today.ToString('yyyy-MM-dd')),
    [string]$ProjectRoot,
    [string]$WorkRoot,
    [switch]$Apply
)
$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
. (Join-Path $PSScriptRoot 'finance_transaction.ps1')
. (Join-Path $PSScriptRoot 'intervention_record.ps1')
. (Join-Path $PSScriptRoot 'order_actual_overrides.ps1')
$lark = Join-Path $PSScriptRoot 'lark_cli_noexit.ps1'
function Read-DeliveryJson([string]$Path,[string]$Label) {
    try { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
    catch { throw "DELIVERY_LOCAL_JSON_FAILED: $Label 无法按UTF-8读取或解析；请检查该本地文件，原配送费保持不动" }
}
$meta = Read-DeliveryJson (Join-Path $ProjectRoot 'work\simple_finance_meta.json') '经营工作簿绑定'
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$stateRoot = Join-Path $ProjectRoot 'agent\state'
$pendingPath = Join-Path $stateRoot 'delivery_fee_pending.json'
$previousDay = ([datetime]::ParseExact($Date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)).AddDays(-1).ToString('yyyy-MM-dd')
$transaction = $null
$writesStarted = $false
function Invoke-DeliveryLark([scriptblock]$Command) {
    $raw = & $Command | Out-String
    if ($LASTEXITCODE -ne 0) { throw '配送费步骤飞书请求失败，未确认写入' }
    $result = $raw | ConvertFrom-Json
    if ($result.ok -ne $true) { throw '配送费步骤飞书返回失败' }
    return $result
}
function Get-DeliveryRows([string]$Token,[string]$Sheet,[string]$Range) {
    $r = Invoke-DeliveryLark { & $lark sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json }
    if ($r.data.has_more) { throw '配送费读取不完整，不作费用调整' }
    return ,@($r.data.ranges[0].cells)
}
function Get-DeliveryCell($Row,[int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell=@($Row)[$Index]
    if($null -eq $cell){return $null}
    $p=$cell.PSObject.Properties['value']
    if($null -eq $p){return $null}
    return $p.Value
}
function Convert-DeliveryDay($Value) {
    $n=0.0
    if ([double]::TryParse([string]$Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)) {
        return ([datetime]'1899-12-30').AddDays([math]::Floor($n)).ToString('yyyy-MM-dd')
    }
    return ([datetime]::Parse([string]$Value,[Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
}
function Set-DeliveryCell([string]$Range,$Value) {
    Register-FinanceRangeBeforeImage $transaction 'source' $token '订单汇总' $Range $lark
    $matrix=@();$matrix+=,@(@{value=$Value})
    $cells=ConvertTo-Json -InputObject $matrix -Depth 5 -Compress
    $null=Invoke-DeliveryLark { $cells | & $lark sheets +cells-set --as user --spreadsheet-token $token --sheet-name '订单汇总' --range $Range --cells - --format json }
}
try {
    if ($Apply) {
        $transaction=Join-FinanceTransaction $stateRoot
        if ($null -eq $transaction) { throw '配送费自动写入必须由同步主事务调用；单独运行请使用只读预演' }
    }
    $rows=Get-DeliveryRows $token '订单汇总' 'A2:Q2001'
    $interventions=@{}
    foreach($r in (Get-DeliveryRows $token '干预订单' 'A2:M1001')) {
        $id=[string](Get-DeliveryCell $r 3)
        if(Test-ActiveInterventionOrderNo $id){$interventions[$id]=$true}
    }
    $seedPath=Join-Path $ProjectRoot 'work\intervention_orders.json'
    if(Test-Path -LiteralPath $seedPath){foreach($id in @((Read-DeliveryJson $seedPath '干预订单种子').intervention_order_ids)){$interventions[[string]$id]=$true}}
    $overrides=@{};foreach($o in @(Get-OrderActualOverrides $ProjectRoot)){$overrides[[string]$o.order_no]=$o}
    $pending=@()
    if(Test-Path -LiteralPath $pendingPath){$pending=@((Read-DeliveryJson $pendingPath '配送费待查队列').orders)}
    $pendingIds=@{};foreach($p in $pending){$pendingIds[[string]$p.order_no]=$true}
    $orders=[Collections.Generic.List[object]]::new();$byId=@{};$monthKeys=@{}
    for($i=0;$i -lt $rows.Count;$i++){
        $r=$rows[$i];$id=[string](Get-DeliveryCell $r 2)
        if(-not $id -or (Get-DeliveryCell $r 1) -ne '美团'){continue}
        $day=Convert-DeliveryDay (Get-DeliveryCell $r 0)
        # Yesterday must be rediscovered even if the prior sync failed before
        # it could persist a pending queue. The pure policy filters resolved fees.
        if($day -ne $Date -and $day -ne $previousDay -and -not $pendingIds.ContainsKey($id)){continue}
        if($byId.ContainsKey($id)){throw '配送费订单号不唯一，停止费用写入'}
        $note=[string](Get-DeliveryCell $r 15);$fee=Get-DeliveryCell $r 9
        $auto=$null
        if($note -match '自动配送费\[(?:maiyatian|cyanclouds):\d+\]=(\d+(?:\.\d+)?)元'){$auto=[double]::Parse($Matches[1],[Globalization.CultureInfo]::InvariantCulture)}
        $protected=$note -match '人工确认配送费' -or $overrides.ContainsKey($id) -or
          ($null -ne $fee -and [string]$fee -ne '' -and [double]$fee -gt 0 -and ($null -eq $auto -or [math]::Abs([double]$fee-$auto) -gt 0.0049))
        $record=[pscustomobject]@{order_no=$id;date=$day;platform='美团';status=[string](Get-DeliveryCell $r 3);intervention=$interventions.ContainsKey($id);current_fee=$(if([string]$fee -eq ''){$null}else{[double]$fee});protected=[bool]$protected;confirmed_self=($note -match '人工确认配送费：?\s*0(?:\.0+)?元');row=($i+2);note=$note;auto_fee=$auto;manual_migration=$false}
        if($overrides.ContainsKey($id)){$record.current_fee=[double]$overrides[$id].delivery_fee;if($record.current_fee -eq 0){$record.confirmed_self=$true}}
        $orders.Add($record);$byId[$id]=$record
        if(-not $record.intervention -and $record.status -notmatch '取消'){$monthKeys[$day.Substring(0,7)]=$true}
    }
    # Income-sheet direct edits have priority even before the old manual-fee
    # migration has copied them into the authoritative summary.
    $ledgerMeta=Read-DeliveryJson (Join-Path $ProjectRoot 'work\daily_ledger_meta.json') '每日经营简账绑定'
    $ledgerToken=([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
    $book=Invoke-DeliveryLark { & $lark sheets +workbook-info --as user --spreadsheet-token $ledgerToken --format json }
    $sheetNames=@($book.data.sheets|ForEach-Object {$_.sheet_name})
    foreach($key in $monthKeys.Keys){
        $name=('{0}月收入' -f [int]$key.Substring(5,2))
        if($name -notin $sheetNames){continue}
        $seenIncome=@{}
        foreach($r in (Get-DeliveryRows $ledgerToken $name 'A2:K1001')){
            $id=[string](Get-DeliveryCell $r 1);if(-not $byId.ContainsKey($id)){continue}
            if($seenIncome.ContainsKey($id)){throw '月份收入的订单号不唯一，不写配送费'};$seenIncome[$id]=$true
            $o=$byId[$id];$v=Get-DeliveryCell $r 5;$n=[string](Get-DeliveryCell $r 10)
            if($n -match '(?:人工确认|店主确认).*(?:自配送|配送费.{0,3}0元)'){$o.protected=$true;$o.confirmed_self=$true;$o.current_fee=0;if([double](Get-DeliveryCell $rows[$o.row-2] 9) -ne 0){$o.manual_migration=$true}}
            if([string]$v -ne '' -and [double]$v -gt 0 -and ($null -eq $o.auto_fee -or [math]::Abs([double]$v-[double]$o.auto_fee) -gt 0.0049)){
                $o.protected=$true;$o.current_fee=[double]$v
                if([string](Get-DeliveryCell $rows[$o.row-2] 9) -eq '' -or [math]::Abs([double](Get-DeliveryCell $rows[$o.row-2] 9)-[double]$v) -gt 0.0049){$o.manual_migration=$true}
            }
        }
    }
    $payload=ConvertTo-Json -InputObject @{date=$Date;orders=@($orders.ToArray());pending=@($pending)} -Depth 10 -Compress
    $raw=$payload | & node (Join-Path $ProjectRoot 'work\delivery_fee_extract.mjs') | Out-String
    if($LASTEXITCODE -ne 0){throw '配送费浏览器提取未完成，保留全部原费用'}
    $result=$raw|ConvertFrom-Json
    if($result.ok -ne $true){throw '配送费返回不完整，保留全部原费用'}
    $changes=[Collections.Generic.List[object]]::new()
    foreach($o in $orders){
        if(-not $o.manual_migration -or $o.intervention -or $o.status -match '取消'){continue}
        $marker='人工确认配送费：{0}元' -f ([double]$o.current_fee).ToString('0.00',[Globalization.CultureInfo]::InvariantCulture)
        $note=([regex]::Replace($o.note,'(?:^|[；;])(?:自动配送费\[(?:maiyatian|cyanclouds):\d+\]=\d+(?:\.\d+)?元|人工确认配送费：?\s*\d+(?:\.\d+)?元)','')).Trim('；',';',' ')
        $note=(@($note,$marker)|Where-Object {$_}) -join '；'
        if($Apply){$writesStarted=$true;Set-DeliveryCell "J$($o.row):J$($o.row)" ([double]$o.current_fee);Set-DeliveryCell "P$($o.row):P$($o.row)" $note
            $back=Get-DeliveryRows $token '订单汇总' "A$($o.row):Q$($o.row)"
            if($back.Count -ne 1 -or (Get-DeliveryCell $back[0] 2) -ne $o.order_no -or [math]::Abs([double](Get-DeliveryCell $back[0] 9)-[double]$o.current_fee) -gt 0.0049 -or (Get-DeliveryCell $back[0] 15) -ne $note){throw '人工配送费归一回读失败'}}
        $changes.Add([pscustomobject]@{order_no=$o.order_no;date=$o.date;amount=$o.current_fee;provider='manual'})
    }
    foreach($d in @($result.decisions)){
        if($d.state -ne 'confirmed' -or -not $d.changed){continue}
        $o=$byId[[string]$d.order_no]
        if($null -eq $o -or $o.intervention -or $o.protected -or $o.confirmed_self){throw '配送费写入保护触发'}
        $marker='自动配送费[{0}:{1}]={2}元' -f $d.provider,$d.delivery_id,([double]$d.amount).ToString('0.00',[Globalization.CultureInfo]::InvariantCulture)
        $note=([regex]::Replace($o.note,'(?:^|[；;])自动配送费\[(?:maiyatian|cyanclouds):\d+\]=\d+(?:\.\d+)?元','')).Trim('；',';',' ')
        $note=(@($note,$marker)|Where-Object {$_}) -join '；'
        if($Apply){
            $writesStarted=$true
            Set-DeliveryCell "J$($o.row):J$($o.row)" ([double]$d.amount)
            Set-DeliveryCell "P$($o.row):P$($o.row)" $note
            $back=Get-DeliveryRows $token '订单汇总' "A$($o.row):Q$($o.row)"
            if($back.Count -ne 1 -or (Get-DeliveryCell $back[0] 2) -ne $d.order_no -or [math]::Abs([double](Get-DeliveryCell $back[0] 9)-[double]$d.amount) -gt 0.0049 -or (Get-DeliveryCell $back[0] 15) -ne $note){throw '配送费金额或来源回读失败'}
        }
        $changes.Add([pscustomobject]@{order_no=$d.order_no;date=$o.date;amount=$d.amount;provider=$d.provider})
    }
    if($Apply -and $changes.Count -gt 0){
        $verify=Invoke-DeliveryLark { & $lark sheets +formula-verify --as user --spreadsheet-token $token --sheet-name '订单汇总' --range 'K2:L2001' --exit-on-error --format json }
        if($verify.data.status -ne 'success'){throw '配送费更新后利润公式校验失败'}
    }
    if($Apply){
        # Keep freshly changed orders for one more run: the parent transaction
        # could still roll back after this child returns. This is a scheduling
        # queue only; no applied-money/idempotency fact is stored here.
        $carry=@($pending|Where-Object { [string]$_.date -gt $Date -or -not $byId.ContainsKey([string]$_.order_no) })
        $next=@($result.pending)+@($changes|ForEach-Object {[pscustomobject]@{order_no=$_.order_no;date=$_.date}})+$carry
        $next=@($next|Sort-Object order_no -Unique)
        $temp="$pendingPath.$PID.tmp"
        try{[IO.File]::WriteAllText($temp,(@{orders=$next}|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false));Move-Item -LiteralPath $temp -Destination $pendingPath -Force}
        finally{if(Test-Path -LiteralPath $temp){Remove-Item -LiteralPath $temp -Force}}
    }
    [ordered]@{ok=$true;applied=[bool]$Apply;candidate_count=$result.candidate_count;queried_count=$result.queried_count;providers=$result.providers;decisions=@($result.decisions);pending_count=@($result.pending).Count;changed_orders=@($changes);changed_order_nos=@($changes|ForEach-Object {$_.order_no});affected_months=@($changes|ForEach-Object {$_.date.Substring(0,7)}|Sort-Object -Unique)}|ConvertTo-Json -Depth 10 -Compress
} catch {
    if($writesStarted){throw "DELIVERY_WRITE_FAILED: $($_.Exception.Message)"}
    throw
} finally {
    if($null -ne $transaction){Close-FinanceTransaction $transaction}
}
