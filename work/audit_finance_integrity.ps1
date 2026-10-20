param([string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent))

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$sourceMeta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$ledgerMeta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\daily_ledger_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $ProjectRoot 'agent\intervention_record.ps1')

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    $valueProperty = $cell.PSObject.Properties['value']
    if ($null -eq $valueProperty -or $null -eq $valueProperty.Value) { return $null }
    return [string]$valueProperty.Value
}
function To-Number($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}
function To-Date($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.Date }
    $number = To-Number $Value
    if ($null -ne $number) { return ([datetime]'1899-12-30').AddDays($number).Date }
    return $null
}
function Get-Rows([string]$Token,[string]$Sheet,[string]$Range) {
    $raw = lark-cli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include value,formula --max-chars 5000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取失败：$Sheet" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "读取不完整：$Sheet" }
    return [pscustomobject]@{rows=[object[]]@($result.data.ranges[0].cells)}
}

function Get-LedgerMonth([string]$SheetName) {
    if ($SheetName -match '^(\d{4})年(\d{1,2})月(?:收入|支出)$') { return [datetime]::new([int]$Matches[1],[int]$Matches[2],1) }
    if ($SheetName -match '^(\d{1,2})月(?:收入|支出)$') { return [datetime]::new(2026,[int]$Matches[1],1) }
    return $null
}

$sourceRows = (Get-Rows $sourceToken '订单汇总' 'A2:Q2001').rows
$orderItemRows = (Get-Rows $sourceToken '订单商品' 'A2:Q4001').rows
$interventionRows = (Get-Rows $sourceToken '干预订单' 'A2:M1001').rows
$infoRaw = lark-cli sheets +workbook-info --as user --spreadsheet-token $ledgerToken --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '读取每日经营简账结构失败' }
$info = $infoRaw | ConvertFrom-Json
$monthSheets = @($info.data.sheets | ForEach-Object {
    $month = Get-LedgerMonth ([string]$_.sheet_name)
    if ($null -ne $month -and $month -ge [datetime]'2026-08-01' -and $month -le [datetime]::Today.Date) {
        [pscustomobject]@{name=[string]$_.sheet_name;month=$month}
    }
} | Sort-Object month,name)
$incomeSheets = @($monthSheets | Where-Object { $_.name.EndsWith('收入') })
$expenseSheets = @($monthSheets | Where-Object { $_.name.EndsWith('支出') })
if ($incomeSheets.Count -eq 0 -or $expenseSheets.Count -eq 0) { throw '未发现可审计的月度收入/支出页签' }
$incomeRows = [Collections.Generic.List[object]]::new()
$expenseRows = [Collections.Generic.List[object]]::new()
foreach ($sheet in $incomeSheets) { foreach ($row in (Get-Rows $ledgerToken $sheet.name 'A2:K1001').rows) { $incomeRows.Add($row) } }
foreach ($sheet in $expenseSheets) { foreach ($row in (Get-Rows $ledgerToken $sheet.name 'A2:F1001').rows) { $expenseRows.Add($row) } }

$sourceByOrder = @{}
$duplicateSourceOrders = @()
foreach ($row in $sourceRows) {
    $orderNo = Get-CellValue $row 2
    if (-not $orderNo) { continue }
    if ($sourceByOrder.ContainsKey($orderNo)) { $duplicateSourceOrders += $orderNo; continue }
    $sourceByOrder[$orderNo] = $row
}
$itemQuantityByOrder = @{}
foreach ($row in $orderItemRows) {
    $orderNo = Get-CellValue $row 2
    if (-not $orderNo) { continue }
    $quantity = To-Number (Get-CellValue $row 8)
    if ($null -eq $quantity) { $quantity = 0.0 }
    if (-not $itemQuantityByOrder.ContainsKey($orderNo)) { $itemQuantityByOrder[$orderNo] = 0.0 }
    $itemQuantityByOrder[$orderNo] = [double]$itemQuantityByOrder[$orderNo] + $quantity
}
$orderItemQuantityMismatches = @()
foreach ($orderNo in $sourceByOrder.Keys) {
    $summaryQuantity = To-Number (Get-CellValue $sourceByOrder[$orderNo] 12)
    if ($null -eq $summaryQuantity) { $summaryQuantity = 0.0 }
    $itemQuantity = if ($itemQuantityByOrder.ContainsKey($orderNo)) { [double]$itemQuantityByOrder[$orderNo] } else { 0.0 }
    if ([math]::Abs($summaryQuantity - $itemQuantity) -gt 0.000001) {
        $orderItemQuantityMismatches += "$orderNo|summary=$summaryQuantity|items=$itemQuantity"
    }
}
$interventionSet = @{}
foreach ($row in $interventionRows) {
    $orderNo = Get-CellValue $row 3
    if (Test-ActiveInterventionOrderNo $orderNo) { $interventionSet[$orderNo] = $true }
}

$incomeByOrder = @{}
$duplicateIncomeOrders = @()
$incomeSourceMismatch = @()
$unmigratedManualDelivery = @()
$interventionLeak = @()
$arithmeticErrors = @()
$sortErrors = @()
$lastSortKey = ''
foreach ($row in $incomeRows) {
    $orderNo = Get-CellValue $row 1
    if (-not $orderNo) { continue }
    if ($incomeByOrder.ContainsKey($orderNo)) { $duplicateIncomeOrders += $orderNo } else { $incomeByOrder[$orderNo] = $row }
    if ($interventionSet.ContainsKey($orderNo)) { $interventionLeak += $orderNo }
    if ($sourceByOrder.ContainsKey($orderNo)) {
        $source = $sourceByOrder[$orderNo]
        foreach ($pair in @(@('income',3,7),@('cost',4,10),@('delivery',5,9))) {
            $ledgerValue = To-Number (Get-CellValue $row $pair[1])
            $sourceValue = To-Number (Get-CellValue $source $pair[2])
            if ($null -eq $ledgerValue) { $ledgerValue = 0.0 }
            if ($null -eq $sourceValue) { $sourceValue = 0.0 }
            # 人工配送费必须回写订单汇总并带来源标记。若只存在月度收入表，
            # 驾驶舱会漏扣该费用，因此它是阻断性数据分叉，不再视为合法覆盖。
            if ($pair[0] -eq 'delivery' -and [math]::Abs($sourceValue) -lt 0.000001 -and $ledgerValue -gt 0) {
                $unmigratedManualDelivery += "$orderNo|$ledgerValue"
                continue
            }
            if ([math]::Abs($ledgerValue-$sourceValue) -gt 0.02) { $incomeSourceMismatch += "$orderNo|$($pair[0])|$ledgerValue|$sourceValue" }
        }
    } elseif ($orderNo -notmatch '^OFF-') {
        $incomeSourceMismatch += "$orderNo|missing_source"
    }
    $income = To-Number (Get-CellValue $row 3)
    $cost = To-Number (Get-CellValue $row 4)
    $delivery = To-Number (Get-CellValue $row 5)
    $profit = To-Number (Get-CellValue $row 6)
    if ($null -ne $cost -and $null -ne $profit) {
        $expected = $(if($null-eq$income){0}else{$income})-$cost-$(if($null-eq$delivery){0}else{$delivery})
        if ([math]::Abs($expected-$profit) -gt 0.02) { $arithmeticErrors += $orderNo }
    }
    $date = To-Date (Get-CellValue $row 2)
    $dailyNo = [string](Get-CellValue $row 0)
    $platformRank = if($dailyNo -like '美团*'){10}elseif($dailyNo -like '饿了么*'){20}elseif($dailyNo -like '京东*'){30}elseif($dailyNo -like '线下*'){50}else{90}
    $sequence = if($dailyNo -match '(\d+)\s*号单'){[int]$Matches[1]}else{999999}
    $sortKey = '{0}|{1:D2}|{2:D6}|{3}' -f $date.ToString('yyyy-MM-dd'),$platformRank,$sequence,$orderNo
    if ($lastSortKey -and $sortKey -lt $lastSortKey) { $sortErrors += "$lastSortKey -> $sortKey" }
    $lastSortKey = $sortKey
}

$missingNaturalOrders = @()
foreach ($orderNo in $sourceByOrder.Keys) {
    $row = $sourceByOrder[$orderNo]
    $date = To-Date (Get-CellValue $row 0)
    if ($null -eq $date -or $date -lt [datetime]'2026-08-24' -or $date -gt [datetime]::Today) { continue }
    if ((Get-CellValue $row 3) -ne '已取消' -and (Get-CellValue $row 4) -eq '自然' -and -not $incomeByOrder.ContainsKey($orderNo)) { $missingNaturalOrders += $orderNo }
}

$expensePairCounts = @{}
foreach ($row in $expenseRows) {
    $project = Get-CellValue $row 1
    $orderNo = Get-CellValue $row 3
    if (-not $orderNo -or $project -notin @('补单平台扣款','补单配送费')) { continue }
    $key = "$orderNo|$project"
    if (-not $expensePairCounts.ContainsKey($key)) { $expensePairCounts[$key] = 0 }
    $expensePairCounts[$key]++
}
$expensePairErrors = @()
foreach ($orderNo in $interventionSet.Keys) {
    foreach ($project in @('补单平台扣款','补单配送费')) {
        $key = "$orderNo|$project"
        $count = if($expensePairCounts.ContainsKey($key)){$expensePairCounts[$key]}else{0}
        if ($count -ne 1) { $expensePairErrors += "$key|count=$count" }
    }
}

[ordered]@{
    ok = ($duplicateSourceOrders.Count -eq 0 -and $duplicateIncomeOrders.Count -eq 0 -and $incomeSourceMismatch.Count -eq 0 -and $unmigratedManualDelivery.Count -eq 0 -and $interventionLeak.Count -eq 0 -and $arithmeticErrors.Count -eq 0 -and $sortErrors.Count -eq 0 -and $missingNaturalOrders.Count -eq 0 -and $expensePairErrors.Count -eq 0 -and $orderItemQuantityMismatches.Count -eq 0)
    source_orders = $sourceByOrder.Count
    income_orders = $incomeByOrder.Count
    intervention_orders = $interventionSet.Count
    audited_income_sheets = @($incomeSheets.name)
    audited_expense_sheets = @($expenseSheets.name)
    duplicate_source_orders = $duplicateSourceOrders
    duplicate_income_orders = $duplicateIncomeOrders
    income_source_mismatches = $incomeSourceMismatch
    unmigrated_manual_delivery_fees = $unmigratedManualDelivery
    intervention_orders_in_income = $interventionLeak
    profit_arithmetic_errors = $arithmeticErrors
    income_sort_errors = $sortErrors
    missing_natural_orders = $missingNaturalOrders
    intervention_expense_pair_errors = $expensePairErrors
    order_item_quantity_mismatches = $orderItemQuantityMismatches
} | ConvertTo-Json -Depth 8
