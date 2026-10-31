param([string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent))

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$agentRoot = Join-Path $ProjectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$sourceMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\daily_ledger_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $agentRoot 'finance_transaction.ps1')

function Invoke-LarkJson([scriptblock]$Action,[string]$Message) {
    $global:LASTEXITCODE = 0
    $raw = @(& $Action 2>&1 | ForEach-Object { [string]$_ }) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "$Message：$raw" }
    try { $result = $raw | ConvertFrom-Json } catch { throw "$Message：返回内容无法解析" }
    if (-not $result.ok) { throw "$Message：$raw" }
    return $result
}

function Get-CellValue($Row,[int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return $cell.value
}

function Convert-ToNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) { return $number }
    return $null
}

function Convert-ToDate($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $serial = Convert-ToNumber $Value
    if ($null -ne $serial -and $serial -gt 30000 -and $serial -lt 100000) { return [datetime]::FromOADate($serial) }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value,[ref]$date)) { return $date }
    return $null
}

function Set-Cells([string]$Sheet,[string]$Range,$Cells,[string]$Message) {
    $payload = ConvertTo-Json -InputObject $Cells -Depth 12 -Compress
    $null = Invoke-LarkJson { $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $sourceToken --sheet-name $Sheet --range $Range --cells - --format json } $Message
}

function Copy-FormulaOrValueCell($Cell) {
    if ($null -eq $Cell) { return @{value=''} }
    $formulaProperty = $Cell.PSObject.Properties['formula']
    if ($null -ne $formulaProperty -and $formulaProperty.Value) { return @{formula=[string]$formulaProperty.Value} }
    $valueProperty = $Cell.PSObject.Properties['value']
    return @{value=$(if($null-ne$valueProperty){$valueProperty.Value}else{''})}
}

function Get-Rows([string]$Token,[string]$Sheet,[string]$Range,[string]$Message) {
    $result = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include 'value,formula,style' --max-chars 1000000 --format json } $Message
    if ($result.data.has_more) { throw "$Message：读取结果不完整" }
    return [pscustomobject]@{rows=[object[]]@($result.data.ranges[0].cells);row_indices=[int[]]@($result.data.ranges[0].row_indices)}
}

$now = [datetime]::Today
$monthPrefix = if ($now.Year -eq 2026) { "$($now.Month)月" } else { "$($now.Year)年$($now.Month)月" }
$expenseSheet = "${monthPrefix}支出"
$expenseRows = (Get-Rows $ledgerToken $expenseSheet 'A2:C1001' '读取当月支出失败').rows
$promotionExpense = 0.0
$promotionRows = 0
foreach ($row in @($expenseRows)) {
    if ([string](Get-CellValue $row 1) -ne '投流推广费') { continue }
    $amount = Convert-ToNumber (Get-CellValue $row 2)
    if ($null -eq $amount) { continue }
    $promotionExpense += [double]$amount
    $promotionRows++
}
$promotionExpense = [math]::Round($promotionExpense,2)

$monthRowsResult = Get-Rows $sourceToken '月度结算' 'A2:N25' '读取月度结算失败'
$monthRowNumber = $null
for ($index = 0; $index -lt $monthRowsResult.rows.Count; $index++) {
    $monthValue = Convert-ToDate (Get-CellValue $monthRowsResult.rows[$index] 0)
    if ($null -eq $monthValue -or $monthValue.Year -ne $now.Year -or $monthValue.Month -ne $now.Month) { continue }
    $monthRowNumber = if ($monthRowsResult.row_indices.Count -eq $monthRowsResult.rows.Count) { [int]$monthRowsResult.row_indices[$index] } else { 2 + $index }
    break
}
if ($null -eq $monthRowNumber) { throw "月度结算中找不到当前月份：$($now.ToString('yyyy-MM'))" }

$dashboardLayout = Get-Rows $sourceToken '驾驶舱' 'A1:G19' '读取驾驶舱布局失败'
$dashboardHeaders = if ($dashboardLayout.rows.Count -gt 0) { @($dashboardLayout.rows[0] | ForEach-Object { if($_.value){[string]$_.value}else{''} }) } else { @() }
$layoutMode = if ($dashboardHeaders.Count -ge 7 -and $dashboardHeaders[3] -eq '计算明细' -and $dashboardHeaders[4] -eq '解释') {
    'current'
} elseif ($dashboardHeaders.Count -ge 7 -and $dashboardHeaders[3] -eq '解释' -and $dashboardHeaders[6] -eq '计算明细') {
    'legacy'
} else {
    throw "驾驶舱列布局无法安全识别：$($dashboardHeaders -join '｜')"
}
$settlementHeaderRows = (Get-Rows $sourceToken '月度结算' 'N1:N1' '读取月度结算推广费列失败').rows
$currentSettlementHeader = if ($settlementHeaderRows.Count -gt 0) { [string](Get-CellValue $settlementHeaderRows[0] 0) } else { '' }
if ($currentSettlementHeader -and $currentSettlementHeader -ne '平台推广费用') { throw "月度结算 N1 已有其他内容：$currentSettlementHeader" }

$deliveryExpenseExpression = 'SUMIFS(''订单汇总''!$J$2:$J$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$E$2:$E$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))'
$promotionExpenseExpression = 'SUMIFS(''月度结算''!$N$2:$N$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$allOperatingExpenseExpression = 'SUMIFS(''月度结算''!$G$2:$G$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$otherOperatingExpenseExpression = '(' + $allOperatingExpenseExpression + ')-(' + $promotionExpenseExpression + ')'
$cashReceiptExpression = 'SUMIFS(''月度结算''!$B$2:$B$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$C$2:$C$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$D$2:$D$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$E$2:$E$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$otherCashExpenseExpression = 'SUMIFS(''月度结算''!$F$2:$F$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$ownerWithdrawalExpression = 'SUMIFS(''月度结算''!$H$2:$H$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$beginningCashExpression = 'SUMIFS(''月度结算''!$K$2:$K$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$grossProfitFormula = '=C4-C5-(' + $deliveryExpenseExpression + ')-(' + $promotionExpenseExpression + ')'
$operatingProfitFormula = '=C6-(' + $otherOperatingExpenseExpression + ')'
$detailFormulas = @{
    6 = '="=线上与线下收入-商品成本-配送费-平台推广费用="&TEXT(C4,"#,##0.00")&"-"&TEXT(C5,"#,##0.00")&"-"&TEXT((' + $deliveryExpenseExpression + '),"#,##0.00")&"-"&TEXT((' + $promotionExpenseExpression + '),"#,##0.00")'
    7 = '="=经营毛利-其他经营费用（不含平台推广费用）="&TEXT(C6,"#,##0.00")&"-"&TEXT((' + $otherOperatingExpenseExpression + '),"#,##0.00")'
    9 = '="=平台与线下到账-补货付款-其他现金支出-老板取款="&TEXT((' + $cashReceiptExpression + '),"#,##0.00")&"-"&TEXT(C8,"#,##0.00")&"-"&TEXT((' + $otherCashExpenseExpression + '),"#,##0.00")&"-"&TEXT((' + $ownerWithdrawalExpression + '),"#,##0.00")'
    10 = '="=期初现金+本月现金净流入="&TEXT((' + $beginningCashExpression + '),"#,##0.00")&IF(C9>=0,"+","")&TEXT(C9,"#,##0.00")'
}

$transaction = New-FinanceTransaction $stateRoot '调整驾驶舱经营毛利并增加计算明细' ('TX-DASHBOARD-DETAILS-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')) -AcquireLock -WorkbookTokens @{source=$sourceToken} -LarkCli $larkCli
$closed = $false
try {
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '月度结算' 'N1:N25' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '驾驶舱' 'A1:G19' $larkCli

    Set-Cells '月度结算' 'N1:N1' @(@(@{value='平台推广费用'})) '设置月度结算推广费表头失败'
    Set-Cells '月度结算' "N${monthRowNumber}:N${monthRowNumber}" @(@(@{value=$promotionExpense;cell_styles=@{number_format='#,##0.00'}})) '回填当月平台推广费用失败'
    if ($layoutMode -eq 'legacy') {
        $layoutMatrix = @()
        foreach ($row in @($dashboardLayout.rows)) {
            $cells = @($row)
            $layoutMatrix += ,@(
                (Copy-FormulaOrValueCell $cells[6]),
                (Copy-FormulaOrValueCell $cells[3]),
                (Copy-FormulaOrValueCell $cells[4]),
                (Copy-FormulaOrValueCell $cells[5])
            )
        }
        Set-Cells '驾驶舱' 'D1:G19' $layoutMatrix '把计算明细移动到数值右侧失败'
    }
    Set-Cells '驾驶舱' 'C6:C7' @(
        @(@{formula=$grossProfitFormula;cell_styles=@{number_format='#,##0.00'}}),
        @(@{formula=$operatingProfitFormula;cell_styles=@{number_format='#,##0.00'}})
    ) '更新经营毛利和经营利润失败'
    Set-Cells '驾驶舱' 'E6:E7' @(
        @(@{value='线上与线下收入-商品成本-配送费-平台推广费用'}),
        @(@{value='经营毛利-其他经营费用（不含平台推广费用）；不再扣老板人工、摊销折旧和市场租金机会成本'})
    ) '更新经营指标解释失败'
    $detailCells = @()
    for ($rowNumber = 2; $rowNumber -le 19; $rowNumber++) {
        if ($detailFormulas.ContainsKey($rowNumber)) { $detailCells += ,@(@{formula=$detailFormulas[$rowNumber]}) }
        else { $detailCells += ,@(@{value=''}) }
    }
    Set-Cells '驾驶舱' 'D2:D19' $detailCells '写入计算明细失败'

    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'A1:G1' --background-color '#173A2D' --font-color '#FFFFFF' --font-weight bold --horizontal-alignment center --vertical-alignment middle --format json } '设置驾驶舱表头样式失败'
    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'D2:D19' --background-color '#EEF5F0' --word-wrap auto-wrap --vertical-alignment middle --format json } '设置计算明细样式失败'
    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'E2:E19' --background-color '#F4F7F5' --word-wrap auto-wrap --vertical-alignment middle --format json } '设置解释列样式失败'
    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'F2:F19' --background-color '#FFF8E1' --format json } '设置状态列样式失败'
    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'G2:G19' --background-color '#F4F7F5' --format json } '设置行动列样式失败'
    foreach ($widthSpec in @(@('D','520'),@('E','260'),@('F','120'),@('G','240'))) {
        $null = Invoke-LarkJson { & $larkCli sheets +cols-resize --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range $widthSpec[0] --width $widthSpec[1] --format json } "设置驾驶舱 $($widthSpec[0]) 列宽失败"
    }
    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $sourceToken --sheet-name '月度结算' --range 'N1:N1' --background-color '#173A2D' --font-color '#FFFFFF' --font-weight bold --horizontal-alignment center --format json } '设置推广费用表头样式失败'
    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $sourceToken --sheet-name '月度结算' --range 'N2:N25' --background-color '#FFF8E1' --number-format '#,##0.00' --format json } '设置推广费用列样式失败'
    $null = Invoke-LarkJson { & $larkCli sheets +cols-resize --as user --spreadsheet-token $sourceToken --sheet-name '月度结算' --range 'N' --width 130 --format json } '设置推广费用列宽失败'

    foreach ($range in @('C6:C7','D6:D7','D9:D10')) {
        $verify = Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range $range --format json --exit-on-error } "驾驶舱公式校验失败：$range"
        if ($verify.data.status -ne 'success') { throw "驾驶舱公式未通过：$range" }
    }

    $dashboardReadback = Get-Rows $sourceToken '驾驶舱' 'A1:G10' '驾驶舱回读失败'
    $dashboardRows = $dashboardReadback.rows
    if ([string](Get-CellValue $dashboardRows[0] 3) -ne '计算明细' -or [string](Get-CellValue $dashboardRows[0] 4) -ne '解释') { throw '驾驶舱列顺序回读不一致' }
    $grossProfit = Convert-ToNumber (Get-CellValue $dashboardRows[5] 2)
    $operatingProfit = Convert-ToNumber (Get-CellValue $dashboardRows[6] 2)
    $grossDetail = [string](Get-CellValue $dashboardRows[5] 3)
    $operatingDetail = [string](Get-CellValue $dashboardRows[6] 3)
    $cashFlowDetail = [string](Get-CellValue $dashboardRows[8] 3)
    $endingCashDetail = [string](Get-CellValue $dashboardRows[9] 3)
    if ($null -eq $grossProfit -or $null -eq $operatingProfit) { throw '经营毛利或经营利润未计算出数值' }
    if ($grossDetail -notmatch '^=线上与线下收入-商品成本-配送费-平台推广费用=') { throw '经营毛利计算明细回读不正确' }
    if ($operatingDetail -notmatch '^=经营毛利-其他经营费用') { throw '经营利润计算明细回读不正确' }
    if ($cashFlowDetail -notmatch '^=平台与线下到账-补货付款-其他现金支出-老板取款=') { throw '现金净流入计算明细回读不正确' }
    if ($endingCashDetail -notmatch '^=期初现金\+本月现金净流入=') { throw '期末现金计算明细回读不正确' }
    $settlementReadback = (Get-Rows $sourceToken '月度结算' "N${monthRowNumber}:N${monthRowNumber}" '推广费用回读失败').rows
    $actualPromotionExpense = if ($settlementReadback.Count -gt 0) { Convert-ToNumber (Get-CellValue $settlementReadback[0] 0) } else { $null }
    if ($null -eq $actualPromotionExpense -or [math]::Abs($actualPromotionExpense-$promotionExpense) -gt 0.005) { throw "推广费用回读不一致：预期=$promotionExpense，实际=$actualPromotionExpense" }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok = $true
        month = $now.ToString('yyyy-MM')
        promotion_rows = $promotionRows
        promotion_expense = $promotionExpense
        gross_profit = [math]::Round($grossProfit,2)
        operating_profit = [math]::Round($operatingProfit,2)
        detail_rows = @(6,7,9,10)
    } | ConvertTo-Json -Depth 6
} catch {
    $failure = $_.Exception.Message
    if (-not $closed) {
        $rollback = Invoke-FinanceTransactionRollback $transaction @{source=$sourceToken} $larkCli $failure
        Close-FinanceTransaction $transaction
        $closed = $true
        if (-not $rollback.ok) { throw "$failure；自动回滚失败：$($rollback.error)" }
    }
    throw $failure
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
