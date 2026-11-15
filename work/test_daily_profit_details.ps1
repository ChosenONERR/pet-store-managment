$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'agent/expense_accounting.ps1')
$script:cases = 0
function Check([bool]$Pass, [string]$Label) {
    if (-not $Pass) { throw $Label }
    $script:cases++
}
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'agent/run_sync_and_notify.ps1'), [ref]$tokens, [ref]$errors)
Check (@($errors).Count -eq 0) 'Main script must parse in Windows PowerShell'
foreach ($name in @('Get-CellValue','Convert-ToNumber','Test-RowDate','Get-DailyNetProfit','Format-DailyProfitDetails','Format-CrossDayCorrection','Format-Money','Limit-Text','Build-Card','Build-Message')) {
    $f = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if ($null -eq $f) { throw "Missing production function $name" }
    $body = $f.Body.Extent.Text.Substring(1, $f.Body.Extent.Text.Length-2)
    if ($null -ne $f.Parameters -and @($f.Parameters).Count) { $body = 'param('+(@($f.Parameters | ForEach-Object { $_.Extent.Text }) -join ',')+")`n"+$body }
    $body = $body.Replace("Join-Path `$PSScriptRoot 'expense_accounting.ps1'", "Join-Path `$root 'agent/expense_accounting.ps1'")
    Set-Item "function:script:$name" ([scriptblock]::Create($body))
}
function Row([object[]]$Values) { return ,@($Values | ForEach-Object { [pscustomobject]@{value=$_} }) }
$script:income = @()
foreach ($numbers in @(
    @(35.08,20.05,4.80,10.23), @(72.43,33.53,6.17,32.73),
    @(51.79,19.55,5.48,26.76), @(112.83,100,3.69,9.14),
    @(19.12,9.40,3.58,6.14), @(31.49,25.72,7.70,-1.93), @(39.87,20.34,4.65,14.88)
)) { $script:income += ,(Row @('美团单','fixture','2026-11-14',$numbers[0],$numbers[1],$numbers[2],$numbers[3])) }
$script:income += ,(Row @('其他日期','other','2026-11-15',1000,200,100,700))
$script:expenses = @()
foreach ($fee in @(1.50,2,2,1.50)) { $script:expenses += ,(Row @('2026-11-14','骑手红包',$fee,'','店主确认')) }
$script:expenses += ,(Row @('2026-11-14','投流推广费',30.09,'','美团推广'))
$script:expenses += ,(Row @('2026-11-14','补单平台扣款',4.18,'fixture-a','平台实付'))
$script:expenses += ,(Row @('2026-11-14','补单配送费',0,'fixture-a','待补录'))
$script:expenses += ,(Row @('2026-11-14','补单平台扣款',6.29,'fixture-b','平台实付'))
$script:expenses += ,(Row @('2026-11-14','补单配送费',0,'fixture-b','自配送'))
$script:expenses += ,(Row @('2026-11-15','推广费',1000,'','其他日期'))
$script:reads = [Collections.Generic.List[string]]::new()
function Get-LedgerRows($Name,$Range) {
    $script:reads.Add("$Name!$Range")
    if ($Name -eq '11月收入') { return ,$script:income }
    if ($Name -eq '11月支出') { return ,$script:expenses }
    throw "Unexpected ledger read: $Name"
}
$script:Date = [datetime]'2026-11-14'
$profit = Get-DailyNetProfit
foreach ($expected in @(
    @('income_total',362.61), @('product_cost_total',228.59), @('first_delivery_total',36.07),
    @('gross_profit',97.95), @('promotion_expense',30.09), @('intervention_expense',10.47),
    @('rider_bonus_expense',7), @('after_sales_expense',0), @('other_expense',0),
    @('expense_total',47.56), @('net_profit',50.39), @('order_count',7)
)) { Check ($profit.($expected[0]) -eq $expected[1]) "Wrong breakdown: $($expected[0])" }
Check ($script:reads.Count -eq 2) 'Details must reuse exactly two ledger reads'
Check ($profit.pending_expenses.Count -eq 0) 'Default self delivery is not a pending expense'
Check ([math]::Round(($profit.gross_profit-$profit.promotion_expense-$profit.intervention_expense-$profit.rider_bonus_expense-$profit.after_sales_expense-$profit.other_expense),2) -eq $profit.net_profit) 'Expense breakdown deducts a fee twice'
$details = Format-DailyProfitDetails $profit
Check ($details.Contains('＝362.61－228.59－36.07＝97.95元')) 'Gross equation missing concrete amounts'
Check ($details.Contains('＝97.95－30.09－10.47－7.00－0.00－0.00＝50.39元')) 'Operating equation missing concrete amounts'
Check ($details.Contains('经营毛利＝97.95元') -and $details.Contains('经营利润＝50.39元')) 'Headings omit result'
$profit | Add-Member -NotePropertyName promotion_pending -NotePropertyValue $true
Check ((Format-DailyProfitDetails $profit).Contains('暂不列完整数值等式')) 'Uncaptured promotion shown as confirmed zero'
$profit.promotion_pending = $false

$script:Mode = 'night'
$script:meta = [pscustomobject]@{workbook_url='https://example.invalid/sheets/test'}
$snapshot = [pscustomobject]@{
    sync_warnings=@(); sync_degraded=$false; today_net_profit=$profit.net_profit
    today_promotion_expense=$profit.promotion_expense; daily_profit_details=$profit
    today_total=9; today_natural=7; today_intervention=2; today_offline=0; today_cancelled=0
    pending_cost=@(); pending_delivery=@(); urgent_restock=@(); restock_excluded_samples=@()
    restore_inventory=@(); expiring_products=@(); expiry_cutoff='2027-07-14'
    metrics=@{'当前缺货SKU数'=0;'当前低库存SKU数'=0;'待核对数据数'=0;'本月经营毛利'=97.95}
}
$card = Build-Card $snapshot $true
$detailBlocks = @($card.body.elements | Where-Object { $_.tag -eq 'markdown' -and $_.content -eq $details })
Check ($detailBlocks.Count -eq 1) 'Card must show the details once'
$message = Build-Message $snapshot $true
Check ($message.Contains($details)) 'Text fallback differs from card'
Check ((Format-CrossDayCorrection '2026-11-14' $profit).Contains($details)) 'Cross-day correction differs from card'
$snapshot.PSObject.Properties.Remove('daily_profit_details')
Check ((Build-Card $snapshot $true).schema -eq '2.0') 'Legacy snapshot no longer renders'
Check ((Build-Message $snapshot $true) -notmatch '首次配送费') 'Legacy snapshot fabricates missing components'

$script:expenses += ,(Row @('2026-11-14','售后追加配送费',5,'','手工'))
$script:expenses += ,(Row @('2026-11-14','店铺维修',3,'','手工'))
$p = Get-DailyNetProfit
Check ($p.after_sales_expense -eq 5 -and $p.other_expense -eq 3 -and $p.rider_bonus_expense -eq 7 -and $p.net_profit -eq 42.39) 'After-sales and other fees not partitioned'
$summary = Get-PetstarExpenseSummary @(
    [pscustomobject]@{project='老板取款';fee=100;order_no='';source='手工'},
    [pscustomobject]@{project='骑手红包';fee=2;order_no='';source='手工'}
)
Check ($summary.total -eq 2 -and $summary.excluded -eq 100 -and $summary.other -eq 2 -and $summary.rider_bonus -eq 2 -and $summary.other_without_rider_bonus -eq 0) 'Display subset changed existing expense/exclusion totals'
$summary = Get-PetstarExpenseSummary @([pscustomobject]@{project='干预单骑手红包';fee=2;order_no='';source='手工'})
Check ($summary.intervention -eq 2 -and $summary.rider_bonus -eq 0 -and $summary.total -eq 2) 'Intervention bonus double counted'
$script:expenses += ,(Row @('2026-11-14','补单平台扣款',$null,'missing','待核对'))
$p = Get-DailyNetProfit
$pendingDetails = Format-DailyProfitDetails $p
Check ($p.pending_expenses.Count -gt 0 -and $pendingDetails.Contains('暂估') -and $pendingDetails.Contains('暂不列完整数值等式')) 'Unknown fees fabricated as final zero'

$script:expenses = @()
$script:income = @(); $script:income += ,(Row @('线下1号','offline','2026-11-14',10,0,0,10))
$p = Get-DailyNetProfit
Check ($p.gross_profit -eq 10 -and $p.product_cost_total -eq 0 -and $p.order_count -eq 1) 'Known zero-cost offline order lost'
$script:income[0][4].value = 12; $script:income[0][6].value = -2
$p = Get-DailyNetProfit
Check ($p.net_profit -eq -2 -and (Format-DailyProfitDetails $p).Contains('经营利润＝-2.00元')) 'Negative profits broken'
$script:income[0][4].value = $null; $script:income[0][6].value = $null
$p = Get-DailyNetProfit
Check ($null -eq $p.gross_profit -and $null -eq $p.product_cost_total -and (Format-DailyProfitDetails $p).Contains('待核对')) 'Missing cost shown as confirmed zero'
$script:income[0][4].value = 2; $script:income[0][6].value = 8; $script:income[0][3].value = $null
$p = Get-DailyNetProfit
Check ($null -eq $p.income_total -and $null -eq $p.gross_profit) 'Missing income shown as known gross profit'
$script:income = @()
$p = Get-DailyNetProfit
Check ($p.order_count -eq 0 -and $p.net_profit -eq 0 -and (Format-DailyProfitDetails $p).Contains('＝0.00－0.00－0.00＝0.00元')) 'Empty day fails to render'
[ordered]@{ok=$true;cases=$script:cases;external_writes=0;credential_access=0;sample=$details} | ConvertTo-Json -Compress
