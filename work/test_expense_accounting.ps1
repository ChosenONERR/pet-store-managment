$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '..\agent\expense_accounting.ps1')
function Entry($Project,$Amount,$Order='',$Source='手工') {
    return [pscustomobject]@{project=$Project;fee=$Amount;order_no=$Order;source=$Source;note=''}
}
$entries=@(
    (Entry '投流推广费' 20), (Entry '补单平台扣款' 8 '123456789012'),
    (Entry '补单配送费' 0 '123456789012' '自配送确认'),
    (Entry '售后追加配送费' 5 '123456789013'), (Entry '维修' 10),
    (Entry '营销赠品支出' 2 '123456789012')
)
$s=Get-PetstarExpenseSummary $entries
if($s.total -ne 45 -or $s.promotion -ne 20 -or $s.intervention -ne 8 -or $s.after_sales -ne 5 -or $s.other -ne 12 -or $s.status -ne '已核对'){throw '分类或费用合计错误'}
# Gross 100 contains no promotion; deduct the expense source once, not twice.
if((100-$s.total) -ne 55){throw '重复扣费'}
$s=Get-PetstarExpenseSummary @((Entry '补单平台扣款' $null '123456789012'),(Entry '补单配送费' 0 '123456789012' '待补录'))
if($s.pending.Count -ne 2 -or $s.status -notmatch '利润暂估'){throw '未知费用未提示'}
$s=Get-PetstarExpenseSummary @((Entry '补货付款' 200),(Entry '老板取款' 100),(Entry '首次配送费' 6),(Entry '维修' 2))
if($s.total -ne 2 -or $s.excluded -ne 306 -or $s.pending.Count -ne 3){throw '非经营支出未隔离'}
$duplicateBlocked=$false
try{$null=Get-PetstarExpenseSummary @((Entry '补单平台扣款' 8 '123456789012'),(Entry '补单平台扣款' 8 '123456789012'))}catch{$duplicateBlocked=$true}
if(-not $duplicateBlocked){throw '重复补单费用未拦截'}
$empty=Get-PetstarExpenseSummary @()
if($empty.total -ne 0){throw '空月份处理错误'}
$f=Get-PetstarDashboardExpenseFormulas 'DATE(2026,11,1)' '32.93'
if($f.gross -ne '=C4-C5-(32.93)' -or $f.profit -match '干预订单' -or $f.profit -notmatch '\$G\$2' -or $f.profit_detail -notmatch '售后追加配送费'){throw '驾驶舱公式口径错误'}
# Simulate the real month-summary writer, including 2-D payloads and readback.
$script:expenseTestWrites=[Collections.Generic.List[object]]::new()
$script:expenseTestSnapshots=[Collections.Generic.List[string]]::new()
function Register-FinanceRangeBeforeImage($Tx,$Key,$Token,$Sheet,$Range,$Cli){$script:expenseTestSnapshots.Add($Range)}
function Invoke-PetstarExpenseCli($Cli,$Arguments,$Payload){
    $range=$Arguments[[array]::IndexOf($Arguments,'--range')+1]
    if($range -notin $script:expenseTestSnapshots){throw '写前缺少事务快照'}
    $cells=ConvertFrom-Json $Payload
    $script:expenseTestWrites.Add([pscustomobject]@{range=$range;cells=$cells;json=$Payload})
    return [pscustomobject]@{ok=$true}
}
function Get-PetstarExpenseRows($Cli,$Token,$Sheet,$Range){
    $grid=@()
    if($Range -eq 'A2:F1001'){
        $grid+=,@(@{value='2026-11-02'},@{value='售后追加配送费'},@{value=5},@{value='123456789013'},@{},@{} | ForEach-Object{[pscustomobject]$_})
        return [pscustomobject]@{cells=$grid;row_indices=@(2)}
    }
    if($Range -eq 'A1:R25'){
        $header=@();$monthRow=@()
        for($j=0;$j-lt18;$j++){$header+=,[pscustomobject]@{};$monthRow+=,[pscustomobject]@{}}
        $monthRow[0]=[pscustomobject]@{value='2026-11'}
        $grid+=,$header;$grid+=,$monthRow
        return [pscustomobject]@{cells=$grid;row_indices=@(1,5)}
    }
    $values=@(5,0,0,0,0,0,0,0,0,5,0,'已核对')
    $grid+=,@($values|ForEach-Object{[pscustomobject]@{value=$_}})
    return [pscustomobject]@{cells=$grid;row_indices=@(5)}
}
Set-StrictMode -Version 2
$s=Update-PetstarMonthExpenses -Month ([datetime]'2026-11-01') -SourceToken test -LedgerToken test -Cli mock -Transaction test
if($s.total-ne5 -or $s.after_sales-ne5 -or $script:expenseTestWrites.Count-ne3){throw '费用录入与汇总集成错误'}
$headerPayload=$script:expenseTestWrites[0].json
if(-not $headerPayload.StartsWith('[[') -or -not $headerPayload.EndsWith(']]')){throw "汇总列头不是二维矩阵：$headerPayload"}
[pscustomobject]@{ok=$true;cases=7;external_writes=0}|ConvertTo-Json
