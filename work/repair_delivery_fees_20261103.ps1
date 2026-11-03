# One-off repair for 2026-11-03 only. Do not blindly rerun on another date.
# Default is read-only. -Apply joins fee writes and the affected month refresh
# in one global transaction; never calls cost rebuild or platform mutations.
param([string]$ProjectRoot,[switch]$Apply)
$ErrorActionPreference='Stop'
$OutputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
if(-not $ProjectRoot){$ProjectRoot=Split-Path $PSScriptRoot -Parent}
$date='2026-11-03'
. (Join-Path $ProjectRoot 'agent\finance_transaction.ps1')
$cli=Join-Path $ProjectRoot 'agent\lark_cli_noexit.ps1'
$source=Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$ledger=Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\daily_ledger_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$tokens=@{source=([uri]$source.workbook_url).Segments[-1].Trim('/');ledger=([uri]$ledger.workbook_url).Segments[-1].Trim('/')}
function Invoke-RepairChild([string]$Path,[hashtable]$Arguments){
    $global:LASTEXITCODE=0;$raw=@(& $Path @Arguments 2>&1|ForEach-Object {[string]$_})
    if($LASTEXITCODE -ne 0){throw "补同步子步骤失败：$(Split-Path $Path -Leaf)；$($raw -join ' ')"}
    $line=@($raw|Where-Object {$_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}')}|Select-Object -Last 1)
    if(-not $line.Count){throw '补同步子步骤没有返回有效结果'}
    $r=$line[0]|ConvertFrom-Json;if(-not $r.ok){throw '补同步子步骤未成功'};return $r
}
function Get-RepairRows([string]$Token,[string]$Sheet,[string]$Range){
    $global:LASTEXITCODE=0;$raw=& $cli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json|Out-String
    if($LASTEXITCODE -ne 0){throw "补同步回读失败：$Sheet"}
    $r=$raw|ConvertFrom-Json
    if(-not $r.ok -or $r.data.has_more -or $r.data.ranges[0].truncated){throw "补同步读取不完整：$Sheet"}
    return ,@($r.data.ranges[0].cells)
}
function Cell($Row,[int]$Index){if($null -eq $Row -or $Index -ge @($Row).Count){return ''};$p=$Row[$Index].PSObject.Properties['value'];if($null -eq $p){return ''};return [string]$p.Value}
$feeScript=Join-Path $ProjectRoot 'agent\sync_delivery_fees.ps1'
if(-not $Apply){Invoke-RepairChild $feeScript @{Date=$date;ProjectRoot=$ProjectRoot}|ConvertTo-Json -Depth 10 -Compress;return}
$tx=$null;$closed=$false
try{
    $tx=New-FinanceTransaction (Join-Path $ProjectRoot 'agent\state') '2026-11-03配送费定向补同步，不改商品成本' $null -AcquireLock -WorkbookTokens $tokens -LarkCli $cli -LockTimeoutSeconds 15
    $before=@{};foreach($row in (Get-RepairRows $tokens.source '订单汇总' 'A2:Q2001')){$id=Cell $row 2;if($id){if($before.ContainsKey($id)){throw '订单号重复，停止补同步'};$before[$id]=@{cost=(Cell $row 10);fee=(Cell $row 9)}}}
    # Compare like-for-like: summary cost is displayed to cents while income
    # retains fractional-cent historical item snapshots. Never round either.
    $incomeBefore=@{};foreach($row in (Get-RepairRows $tokens.ledger '11月收入' 'A2:K1001')){$id=Cell $row 1;if($id){if($incomeBefore.ContainsKey($id)){throw '收入订单号重复，停止补同步'};$incomeBefore[$id]=(Cell $row 4)}}
    $fees=Invoke-RepairChild $feeScript @{Date=$date;ProjectRoot=$ProjectRoot;Apply=$true}
    foreach($c in @($fees.changed_orders)){if($c.date -ne $date){throw '专项脚本不允许修改其他日期配送费'}}
    $ledgerResult=$null
    if(@($fees.changed_order_nos).Count -gt 0){
        if(@($fees.affected_months).Count -ne 1 -or $fees.affected_months[0] -ne '2026-11'){throw '月份范围异常'}
        $ledgerResult=Invoke-RepairChild (Join-Path $ProjectRoot 'agent\sync_daily_ledger.ps1') @{ThroughDate=[datetime]$date;MonthFilter='2026-11';ProjectRoot=$ProjectRoot}
    }
    $after=@{};foreach($row in (Get-RepairRows $tokens.source '订单汇总' 'A2:Q2001')){$id=Cell $row 2;if($id){$after[$id]=$row}}
    if($after.Count -ne $before.Count){throw '补同步期间订单数量发生变化'}
    foreach($id in $before.Keys){if(-not $after.ContainsKey($id) -or (Cell $after[$id] 10) -ne $before[$id].cost){throw '商品成本发生变化，停止并回滚'}}
    $receipts=@()
    if(@($fees.changed_order_nos).Count -gt 0){
        $income=Get-RepairRows $tokens.ledger '11月收入' 'A2:K1001'
        foreach($c in @($fees.changed_orders)){
            $matches=@($income|Where-Object {(Cell $_ 1) -eq $c.order_no})
            if($matches.Count -ne 1){throw '补同步收入订单不能唯一定位'}
            $row=$matches[0];$sourceRow=$after[$c.order_no]
            if([math]::Abs([double](Cell $sourceRow 9)-[double]$c.amount) -gt 0.0049 -or [math]::Abs([double](Cell $row 5)-[double]$c.amount) -gt 0.0049){throw '订单与收入表配送费不一致'}
            if(-not $incomeBefore.ContainsKey($c.order_no) -or (Cell $row 4) -ne $incomeBefore[$c.order_no]){throw '收入表商品成本变动，停止并回滚'}
            $expected=[double](Cell $row 3)-[double](Cell $row 4)-[double](Cell $row 5)
            if([math]::Abs([double](Cell $row 6)-$expected) -gt 0.02){throw '配送费补同步后利润校验失败'}
            $receipts+=,[pscustomobject]@{order_no=$c.order_no;daily_order_no=(Cell $row 0);old_delivery_fee=$before[$c.order_no].fee;delivery_fee=[double]$c.amount;product_cost=[double](Cell $row 4);profit=[double](Cell $row 6);provider=$c.provider}
        }
    }
    Complete-FinanceTransaction $tx;Close-FinanceTransaction $tx;$closed=$true
    [pscustomobject]@{ok=$true;date=$date;applied=$true;candidate_count=$fees.candidate_count;providers=$fees.providers;pending_count=$fees.pending_count;decisions=$fees.decisions;updated_orders=$receipts;product_costs_unchanged=$true;affected_months=$fees.affected_months;transaction_id=$tx.id}|ConvertTo-Json -Depth 10 -Compress
}catch{
    if($null -ne $tx -and -not $closed){$r=Invoke-FinanceTransactionRollback $tx $tokens $cli $_.Exception.Message;if(-not $r.ok){throw "补同步回滚待恢复，事务号：$($tx.id)"}}
    throw
}finally{if($null -ne $tx -and -not $closed){Close-FinanceTransaction $tx}}
