param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent),
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$sourceMeta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$ledgerMeta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\daily_ledger_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$larkCli = Join-Path $ProjectRoot 'agent\lark_cli_noexit.ps1'
. (Join-Path $ProjectRoot 'agent\sku_master_schema.ps1')
. (Join-Path $ProjectRoot 'agent\intervention_record.ps1')
$skuSchema = Get-SkuMasterSchema -SpreadsheetToken $sourceToken

if (-not $OutputPath) {
    $OutputPath = Join-Path $ProjectRoot ('outputs\order_cost_audit_{0}.json' -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}

function Convert-ToNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}

function Convert-ToDate($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.Date }
    $number = Convert-ToNumber $Value
    if ($null -ne $number) { return ([datetime]'1899-12-30').AddDays($number).Date }
    return $null
}

function Test-ZeroCostTrialItem([string]$Title, [string]$Spec) {
    $text = (([string]$Title) + ' ' + ([string]$Spec)).Trim()
    return [bool]($text -match '试吃|尝鲜|试用|体验装')
}

function Get-Rows([string]$Token, [string]$Sheet, [string]$Range) {
    $raw = & $larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取失败：$Sheet" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "读取不完整：$Sheet" }
    return @($result.data.ranges[0].cells)
}

function Round-Money($Value) {
    if ($null -eq $Value) { return $null }
    return [math]::Round([double]$Value, 2, [MidpointRounding]::AwayFromZero)
}

function Add-MapList([hashtable]$Map, [string]$Key, $Value) {
    if (-not $Key) { return }
    if (-not $Map.ContainsKey($Key)) { $Map[$Key] = [Collections.Generic.List[object]]::new() }
    $Map[$Key].Add($Value)
}

$referenceRows = Get-Rows $sourceToken '商品成本' 'A2:E2000'
$masterRows = Get-Rows $sourceToken 'SKU成本主表' 'A2:Z3000'
$itemRows = Get-Rows $sourceToken '订单商品' 'A2:Q4001'
$summaryRows = Get-Rows $sourceToken '订单汇总' 'A2:Q2001'
$offlineRows = Get-Rows $sourceToken '线下订单' 'A2:J1001'
$giftRows = Get-Rows $sourceToken '赠品记录' 'A2:L2001'
$replenishmentRows = Get-Rows $sourceToken '补货记录' 'A2:T1001'
$interventionRows = Get-Rows $sourceToken '干预订单' 'A2:M1001'
$augIncomeRows = Get-Rows $ledgerToken '8月收入' 'A2:K1001'
$sepIncomeRows = Get-Rows $ledgerToken '9月收入' 'A2:K1001'

$referenceByCode = @{}
foreach ($row in $referenceRows) {
    $code = ([string](Get-CellValue $row 0)).Trim().ToLowerInvariant()
    $cost = Convert-ToNumber (Get-CellValue $row 3)
    $status = [string](Get-CellValue $row 4)
    if ($code -and $null -ne $cost -and $status -notin @('待人工核对','缺少成本')) {
        $referenceByCode[$code] = [double]$cost
    }
}

$masterByProductId = @{}
$masterByAlias = @{}
$states = @{}
$duplicateProductIds = [Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $masterRows.Count; $i++) {
    $row = $masterRows[$i]
    $rowNumber = 2 + $i
    $kw = ([string](Get-CellValue $row $skuSchema.CurrentKw.Index)).Trim().ToLowerInvariant()
    $productId = ([string](Get-CellValue $row $skuSchema.ProductId.Index)).Trim()
    $transitionCodes = @([regex]::Split([string](Get-CellValue $row $skuSchema.TransitionCode.Index), '[/,，、;；\s]+') | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique)
    if (-not $kw -and -not $productId -and $transitionCodes.Count -eq 0) { continue }
    $canonical = if ($productId) { "sku:$productId" } else { "legacy:$rowNumber" }
    $initDate = Convert-ToDate (Get-CellValue $row $skuSchema.InitializationDate.Index)
    if ($null -eq $initDate) { $initDate = [datetime]'2026-08-24' }
    $fallbackCandidates = @($transitionCodes | Where-Object { $referenceByCode.ContainsKey($_) } | ForEach-Object { [double]$referenceByCode[$_] } | Sort-Object -Unique)
    $state = [pscustomobject]@{
        canonical = $canonical
        row = $rowNumber
        product_id = $productId
        kw = $kw
        transition_codes = $transitionCodes
        title = Get-CellValue $row $skuSchema.PlatformTitle.Index
        spec = Get-CellValue $row $skuSchema.CurrentSpec.Index
        status = Get-CellValue $row $skuSchema.CostStatus.Index
        baseline = Convert-ToNumber (Get-CellValue $row $skuSchema.BaselineUnitCost.Index)
        opening_qty = $(if ($null -eq (Convert-ToNumber (Get-CellValue $row $skuSchema.InitialInventory.Index))) { 0.0 } else { [double](Convert-ToNumber (Get-CellValue $row $skuSchema.InitialInventory.Index)) })
        qty = $(if ($null -eq (Convert-ToNumber (Get-CellValue $row $skuSchema.InitialInventory.Index))) { 0.0 } else { [double](Convert-ToNumber (Get-CellValue $row $skuSchema.InitialInventory.Index)) })
        avg = Convert-ToNumber (Get-CellValue $row $skuSchema.BaselineUnitCost.Index)
        units_per_pack = $(if ($null -eq (Convert-ToNumber (Get-CellValue $row $skuSchema.UnitsPerPack.Index))) { 1.0 } else { [double](Convert-ToNumber (Get-CellValue $row $skuSchema.UnitsPerPack.Index)) })
        init_date = $initDate.Date
        fallback = $(if ($fallbackCandidates.Count -eq 1) { [double]$fallbackCandidates[0] } else { $null })
    }
    $states[$canonical] = $state
    if ($productId) {
        if ($masterByProductId.ContainsKey($productId)) {
            $duplicateProductIds.Add([pscustomobject]@{product_id=$productId;rows=@($masterByProductId[$productId].row,$rowNumber)})
        } else {
            $masterByProductId[$productId] = $state
        }
    }
    foreach ($alias in @($kw) + $transitionCodes) { Add-MapList $masterByAlias $alias $state }
}

$uniqueAlias = @{}
foreach ($key in $masterByAlias.Keys) {
    $uniqueStates = @($masterByAlias[$key] | Sort-Object canonical -Unique)
    if ($uniqueStates.Count -eq 1) { $uniqueAlias[$key] = $uniqueStates[0] }
}

$interventionSet = @{}
foreach ($row in $interventionRows) {
    $orderNo = Get-CellValue $row 3
    if (Test-ActiveInterventionOrderNo $orderNo) { $interventionSet[$orderNo] = $true }
}

$items = [Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $itemRows.Count; $i++) {
    $row = $itemRows[$i]
    $orderNo = Get-CellValue $row 2
    $date = Convert-ToDate (Get-CellValue $row 0)
    if (-not $orderNo -or $null -eq $date) { continue }
    $rawCode = ([string](Get-CellValue $row 5)).Trim().ToLowerInvariant()
    $productId = ([string](Get-CellValue $row 16)).Trim()
    $state = $null
    $sourceMethod = 'unresolved'
    $consistentAliasCost = $null
    if ($productId -and $masterByProductId.ContainsKey($productId)) {
        $state = $masterByProductId[$productId]
        $sourceMethod = 'exact_product_id'
    } elseif ($rawCode -and $uniqueAlias.ContainsKey($rawCode)) {
        $state = $uniqueAlias[$rawCode]
        $sourceMethod = 'unique_alias'
    } elseif ($rawCode -and $masterByAlias.ContainsKey($rawCode)) {
        $aliasCosts = @($masterByAlias[$rawCode] | ForEach-Object { $_.baseline } | Where-Object { $null -ne $_ } | ForEach-Object { [double]$_ } | Sort-Object -Unique)
        if ($aliasCosts.Count -eq 1) {
            # 商品身份可能因货号复用而不唯一，但候选行成本完全一致时，成本金额本身仍可核验。
            $sourceMethod = 'consistent_alias_cost'
            $consistentAliasCost = [double]$aliasCosts[0]
        } else {
            $sourceMethod = 'ambiguous_alias'
        }
    }
    $items.Add([pscustomobject]@{
        row = 2 + $i
        date = $date.Date
        platform = Get-CellValue $row 1
        order_no = $orderNo
        order_status = Get-CellValue $row 3
        order_type = Get-CellValue $row 4
        raw_code = $rawCode
        title = Get-CellValue $row 6
        spec = Get-CellValue $row 7
        product_id = $productId
        quantity = $(if ($null -eq (Convert-ToNumber (Get-CellValue $row 8))) { 0.0 } else { [double](Convert-ToNumber (Get-CellValue $row 8)) })
        sales_amount = $(if ($null -eq (Convert-ToNumber (Get-CellValue $row 10))) { 0.0 } else { [double](Convert-ToNumber (Get-CellValue $row 10)) })
        actual_unit_cost = Convert-ToNumber (Get-CellValue $row 11)
        actual_item_cost = Convert-ToNumber (Get-CellValue $row 12)
        actual_margin = Convert-ToNumber (Get-CellValue $row 13)
        intervention = $interventionSet.ContainsKey($orderNo)
        cancelled = (Get-CellValue $row 3) -eq '已取消'
        trial_zero_cost = Test-ZeroCostTrialItem (Get-CellValue $row 6) (Get-CellValue $row 7)
        state = $state
        source_method = $sourceMethod
        consistent_alias_cost = $consistentAliasCost
        expected_unit_cost = $null
        expected_item_cost = $null
        expected_source = $null
    })
}

$events = [Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $replenishmentRows.Count; $i++) {
    $row = $replenishmentRows[$i]
    $date = Convert-ToDate (Get-CellValue $row 0)
    $rawCode = ([string](Get-CellValue $row 1)).Trim().ToLowerInvariant()
    $productId = ([string](Get-CellValue $row 14)).Trim()
    $state = if ($productId -and $masterByProductId.ContainsKey($productId)) { $masterByProductId[$productId] } elseif ($rawCode -and $uniqueAlias.ContainsKey($rawCode)) { $uniqueAlias[$rawCode] } else { $null }
    $packageQty = Convert-ToNumber (Get-CellValue $row 2)
    $amount = Convert-ToNumber (Get-CellValue $row 5)
    if ($null -eq $state -or $null -eq $date -or $null -eq $packageQty -or $packageQty -le 0 -or $null -eq $amount -or $amount -lt 0) { continue }
    if ($date.Date -lt $state.init_date) { continue }
    $units = [double]$packageQty * [double]$state.units_per_pack
    $unitCost = if($units -gt 0){[double]$amount / $units}else{$null}
    $isOpeningDuplicate = (
        $units -gt 0 -and
        $date.Date -eq $state.init_date -and
        [math]::Abs($units - [double]$state.opening_qty) -lt 0.000001 -and
        $null -ne $state.baseline -and
        [math]::Abs([double]$unitCost - [double]$state.baseline) -lt 0.000001
    )
    if($isOpeningDuplicate){continue}
    $events.Add([pscustomobject]@{date=$date.Date;sort=1;row=2+$i;kind='purchase';state=$state;package_qty=[double]$packageQty;amount=[double]$amount})
}

foreach ($item in $items) {
    if ($null -eq $item.state -or $item.intervention -or $item.cancelled -or $item.date -lt $item.state.init_date) { continue }
    $events.Add([pscustomobject]@{date=$item.date;sort=2;row=$item.row;kind='sale';state=$item.state;item=$item})
}

foreach ($event in @($events | Sort-Object date,sort,row)) {
    $state = $event.state
    if ($event.kind -eq 'purchase') {
        $units = [double]$event.package_qty * [double]$state.units_per_pack
        if ($units -le 0) { continue }
        $afterQty = [double]$state.qty + $units
        if ($null -eq $state.avg -and $state.qty -gt 0) {
            $afterAvg = $null
        } elseif ($afterQty -le 0) {
            $afterAvg = [double]$event.amount / $units
        } else {
            $oldValue = if ($null -eq $state.avg) { 0.0 } else { [double]$state.avg * [double]$state.qty }
            $afterAvg = ($oldValue + [double]$event.amount) / $afterQty
        }
        $state.qty = $afterQty
        $state.avg = $afterAvg
    } else {
        $item = $event.item
        $item.expected_unit_cost = $state.avg
        if ($null -ne $state.avg) {
            $item.expected_item_cost = [double]$state.avg * [double]$item.quantity
            $item.expected_source = 'moving_average_by_product_id'
        }
        $state.qty = [math]::Max(0.0, [double]$state.qty - [double]$item.quantity)
    }
}

foreach ($item in $items) {
    if ($item.intervention -or $item.cancelled) { continue }
    if ($item.trial_zero_cost) {
        $item.expected_unit_cost = 0.0
        $item.expected_item_cost = 0.0
        $item.expected_source = 'business_rule_trial_zero'
        continue
    }
    if ($item.source_method -eq 'consistent_alias_cost' -and $null -ne $item.consistent_alias_cost) {
        $item.expected_unit_cost = [double]$item.consistent_alias_cost
        $item.expected_item_cost = [double]$item.consistent_alias_cost * [double]$item.quantity
        $item.expected_source = 'consistent_cost_across_reused_alias'
        continue
    }
    if ($null -ne $item.expected_unit_cost) { continue }
    if ($null -ne $item.state -and $item.date -lt $item.state.init_date -and $null -ne $item.state.baseline) {
        $item.expected_unit_cost = [double]$item.state.baseline
        $item.expected_item_cost = [double]$item.state.baseline * [double]$item.quantity
        $item.expected_source = 'baseline_by_product_id'
        continue
    }
    if ($item.source_method -eq 'exact_product_id') {
        $item.expected_source = 'exact_product_id_missing_cost'
        continue
    }
    if ($item.source_method -eq 'unique_alias' -and $null -ne $item.state.baseline) {
        $item.expected_unit_cost = [double]$item.state.baseline
        $item.expected_item_cost = [double]$item.state.baseline * [double]$item.quantity
        $item.expected_source = 'baseline_by_unique_alias'
        continue
    }
    if ($referenceByCode.ContainsKey($item.raw_code)) {
        $item.expected_unit_cost = [double]$referenceByCode[$item.raw_code]
        $item.expected_item_cost = [double]$referenceByCode[$item.raw_code] * [double]$item.quantity
        $item.expected_source = 'legacy_reference_by_raw_code'
    } else {
        $item.expected_source = 'unresolved'
    }
}

$auditItems = @($items | Where-Object { $_.date -ge [datetime]'2026-08-01' -and $_.date -le [datetime]'2026-09-30' -and -not $_.intervention -and -not $_.cancelled })
$itemArithmeticErrors = [Collections.Generic.List[object]]::new()
$expectedCostMismatches = [Collections.Generic.List[object]]::new()
$unreliableCostItems = [Collections.Generic.List[object]]::new()
$orderItemCosts = @{}
$orderExpectedCosts = @{}
$orderHasMissingExpected = @{}
$orderItemsByOrder = @{}

foreach ($item in $auditItems) {
    Add-MapList $orderItemsByOrder $item.order_no $item
    if (-not $orderItemCosts.ContainsKey($item.order_no)) { $orderItemCosts[$item.order_no] = 0.0 }
    if ($null -ne $item.actual_item_cost) { $orderItemCosts[$item.order_no] += [double]$item.actual_item_cost }
    if (-not $orderExpectedCosts.ContainsKey($item.order_no)) { $orderExpectedCosts[$item.order_no] = 0.0 }
    if ($null -ne $item.expected_item_cost) { $orderExpectedCosts[$item.order_no] += [double]$item.expected_item_cost } else { $orderHasMissingExpected[$item.order_no] = $true }

    if ($null -ne $item.actual_unit_cost -and $null -ne $item.actual_item_cost) {
        $calculated = [double]$item.actual_unit_cost * [double]$item.quantity
        if ([math]::Abs($calculated - [double]$item.actual_item_cost) -gt 0.02) {
            $itemArithmeticErrors.Add([pscustomobject]@{order_no=$item.order_no;row=$item.row;product_id=$item.product_id;code=$item.raw_code;quantity=$item.quantity;unit_cost=$item.actual_unit_cost;item_cost=$item.actual_item_cost;calculated=(Round-Money $calculated)})
        }
    }
    if ($null -ne $item.expected_unit_cost -and $null -ne $item.actual_unit_cost -and [math]::Abs([double]$item.expected_unit_cost - [double]$item.actual_unit_cost) -gt 0.02) {
        $expectedCostMismatches.Add([pscustomobject]@{order_no=$item.order_no;date=$item.date.ToString('yyyy-MM-dd');row=$item.row;product_id=$item.product_id;code=$item.raw_code;title=$item.title;spec=$item.spec;quantity=$item.quantity;actual_unit_cost=(Round-Money $item.actual_unit_cost);expected_unit_cost=(Round-Money $item.expected_unit_cost);expected_source=$item.expected_source;master_row=$(if($item.state){$item.state.row}else{$null});master_status=$(if($item.state){$item.state.status}else{$null})})
    }
    if (-not $item.trial_zero_cost -and ($item.expected_source -in @('exact_product_id_missing_cost','legacy_reference_by_raw_code','unresolved') -or $item.source_method -in @('ambiguous_alias','unresolved'))) {
        $unreliableCostItems.Add([pscustomobject]@{order_no=$item.order_no;date=$item.date.ToString('yyyy-MM-dd');row=$item.row;product_id=$item.product_id;code=$item.raw_code;title=$item.title;spec=$item.spec;quantity=$item.quantity;actual_unit_cost=(Round-Money $item.actual_unit_cost);actual_item_cost=(Round-Money $item.actual_item_cost);source_method=$item.source_method;expected_source=$item.expected_source;master_row=$(if($item.state){$item.state.row}else{$null});master_status=$(if($item.state){$item.state.status}else{$null})})
    }
}

$giftCostByOrder = @{}
foreach ($row in $giftRows) {
    $orderNo = Get-CellValue $row 3
    $dailyNo = Get-CellValue $row 2
    $cost = Convert-ToNumber (Get-CellValue $row 7)
    if (-not $orderNo -or $dailyNo -in @('待同步','订单已取消待处理') -or $null -eq $cost) { continue }
    if (-not $giftCostByOrder.ContainsKey($orderNo)) { $giftCostByOrder[$orderNo] = 0.0 }
    $giftCostByOrder[$orderNo] += [double]$cost
}

$summaryByOrder = @{}
foreach ($row in $summaryRows) {
    $orderNo = Get-CellValue $row 2
    $date = Convert-ToDate (Get-CellValue $row 0)
    if (-not $orderNo -or $null -eq $date -or $date -lt [datetime]'2026-08-01' -or $date -gt [datetime]'2026-09-30') { continue }
    if ((Get-CellValue $row 3) -eq '已取消' -or (Get-CellValue $row 4) -ne '自然') { continue }
    $summaryByOrder[$orderNo] = [pscustomobject]@{
        date = $date.Date
        platform = Get-CellValue $row 1
        order_no = $orderNo
        daily_no = Get-CellValue $row 16
        income = Convert-ToNumber (Get-CellValue $row 7)
        delivery = Convert-ToNumber (Get-CellValue $row 9)
        product_cost = Convert-ToNumber (Get-CellValue $row 10)
        profit = Convert-ToNumber (Get-CellValue $row 11)
    }
}

$ledgerByOrder = @{}
foreach ($entry in @(@{month='2026-08';rows=$augIncomeRows},@{month='2026-09';rows=$sepIncomeRows})) {
    for ($i = 0; $i -lt $entry.rows.Count; $i++) {
        $row = $entry.rows[$i]
        $orderNo = Get-CellValue $row 1
        if (-not $orderNo) { continue }
        $ledgerByOrder[$orderNo] = [pscustomobject]@{
            month=$entry.month;row=2+$i;daily_no=Get-CellValue $row 0;date=Convert-ToDate (Get-CellValue $row 2)
            income=Convert-ToNumber (Get-CellValue $row 3);product_cost=Convert-ToNumber (Get-CellValue $row 4)
            delivery=Convert-ToNumber (Get-CellValue $row 5);profit=Convert-ToNumber (Get-CellValue $row 6)
        }
    }
}

$offlineByOrder = @{}
foreach ($row in $offlineRows) {
    $orderNo = Get-CellValue $row 2
    $date = Convert-ToDate (Get-CellValue $row 0)
    if (-not $orderNo -or $null -eq $date -or $date -lt [datetime]'2026-08-01' -or $date -gt [datetime]'2026-09-30') { continue }
    $offlineByOrder[$orderNo] = [pscustomobject]@{
        date=$date.Date;daily_no=Get-CellValue $row 1;order_no=$orderNo
        income=Convert-ToNumber (Get-CellValue $row 3);delivery=Convert-ToNumber (Get-CellValue $row 4)
        product_cost=Convert-ToNumber (Get-CellValue $row 5);profit=Convert-ToNumber (Get-CellValue $row 6)
        sold_goods=Get-CellValue $row 7
    }
}

$orderReconciliationErrors = [Collections.Generic.List[object]]::new()
$incomeSheetErrors = [Collections.Generic.List[object]]::new()
$largeLossOrders = [Collections.Generic.List[object]]::new()
$preservedManualDeliveryFees = [Collections.Generic.List[object]]::new()
$offlineReconciliationErrors = [Collections.Generic.List[object]]::new()

foreach ($orderNo in $summaryByOrder.Keys) {
    $summary = $summaryByOrder[$orderNo]
    $effectiveDelivery = $summary.delivery
    $effectiveProfit = $summary.profit
    $giftCost = if ($giftCostByOrder.ContainsKey($orderNo)) { [double]$giftCostByOrder[$orderNo] } else { 0.0 }
    $itemCost = if ($orderItemCosts.ContainsKey($orderNo)) { [double]$orderItemCosts[$orderNo] } else { 0.0 }
    $calculatedOrderCost = $itemCost + $giftCost
    if ($null -eq $summary.product_cost -or [math]::Abs($calculatedOrderCost - [double]$summary.product_cost) -gt 0.02) {
        $orderReconciliationErrors.Add([pscustomobject]@{order_no=$orderNo;date=$summary.date.ToString('yyyy-MM-dd');summary_cost=(Round-Money $summary.product_cost);item_cost=(Round-Money $itemCost);gift_cost=(Round-Money $giftCost);calculated=(Round-Money $calculatedOrderCost)})
    }
    $expectedProfit = if ($null -eq $summary.product_cost) { $null } else { [double]$summary.income - [double]$summary.product_cost - $(if($null-eq$summary.delivery){0.0}else{[double]$summary.delivery}) }
    if ($null -ne $expectedProfit -and ($null -eq $summary.profit -or [math]::Abs($expectedProfit - [double]$summary.profit) -gt 0.02)) {
        $orderReconciliationErrors.Add([pscustomobject]@{order_no=$orderNo;date=$summary.date.ToString('yyyy-MM-dd');profit=$summary.profit;calculated_profit=(Round-Money $expectedProfit);issue='profit_arithmetic'})
    }
    if (-not $ledgerByOrder.ContainsKey($orderNo)) {
        $incomeSheetErrors.Add([pscustomobject]@{order_no=$orderNo;date=$summary.date.ToString('yyyy-MM-dd');issue='missing_income_row'})
    } else {
        $ledger = $ledgerByOrder[$orderNo]
        $manualDeliveryPreserved = $null -ne $ledger.delivery -and [double]$ledger.delivery -gt 0 -and ($null -eq $summary.delivery -or [math]::Abs([double]$summary.delivery) -lt 0.000001)
        if ($manualDeliveryPreserved) {
            $preservedManualDeliveryFees.Add([pscustomobject]@{order_no=$orderNo;date=$summary.date.ToString('yyyy-MM-dd');delivery=(Round-Money $ledger.delivery);income_sheet=$ledger.month;row=$ledger.row})
        }
        # The monthly income sheet is authoritative for manually entered delivery fees.
        # Use its final delivery/profit values when reporting large-loss orders so the
        # audit does not understate losses after preserving a manual delivery charge.
        $effectiveDelivery = $ledger.delivery
        $effectiveProfit = $ledger.profit
        foreach ($field in @('income','product_cost','delivery','profit')) {
            $left = $summary.$field
            $right = $ledger.$field
            if ($manualDeliveryPreserved -and $field -eq 'delivery') { continue }
            if ($manualDeliveryPreserved -and $field -eq 'profit') {
                $manualProfit = [double]$summary.income - [double]$summary.product_cost - [double]$ledger.delivery
                if ($null -ne $right -and [math]::Abs($manualProfit - [double]$right) -le 0.02) { continue }
            }
            if ($null -eq $left -and $null -eq $right) { continue }
            if ($null -eq $left -or $null -eq $right -or [math]::Abs([double]$left - [double]$right) -gt 0.02) {
                $incomeSheetErrors.Add([pscustomobject]@{order_no=$orderNo;date=$summary.date.ToString('yyyy-MM-dd');field=$field;source=(Round-Money $left);income_sheet=(Round-Money $right);month=$ledger.month;row=$ledger.row})
            }
        }
    }
    if ($null -ne $effectiveProfit -and [double]$effectiveProfit -le -10) {
        $riskItems = @($unreliableCostItems | Where-Object order_no -eq $orderNo)
        $largeLossOrders.Add([pscustomobject]@{
            order_no=$orderNo;date=$summary.date.ToString('yyyy-MM-dd');daily_no=$summary.daily_no
            income=(Round-Money $summary.income);product_cost=(Round-Money $summary.product_cost);delivery=(Round-Money $effectiveDelivery);profit=(Round-Money $effectiveProfit)
            cost_source_risk=($riskItems.Count -gt 0);risk_items=$riskItems
            items=@($orderItemsByOrder[$orderNo] | ForEach-Object { [pscustomobject]@{code=$_.raw_code;product_id=$_.product_id;title=$_.title;spec=$_.spec;quantity=$_.quantity;unit_cost=(Round-Money $_.actual_unit_cost);item_cost=(Round-Money $_.actual_item_cost);expected_source=$_.expected_source;master_row=$(if($_.state){$_.state.row}else{$null});master_status=$(if($_.state){$_.state.status}else{$null})} })
        })
    }
}


foreach ($orderNo in $offlineByOrder.Keys) {
    $offline = $offlineByOrder[$orderNo]
    if (-not $ledgerByOrder.ContainsKey($orderNo)) {
        $offlineReconciliationErrors.Add([pscustomobject]@{order_no=$orderNo;date=$offline.date.ToString('yyyy-MM-dd');issue='missing_income_row'})
        continue
    }
    $ledger = $ledgerByOrder[$orderNo]
    foreach ($field in @('income','product_cost','delivery','profit')) {
        $left = $offline.$field
        $right = $ledger.$field
        if ($null -eq $left -and $null -eq $right) { continue }
        if ($null -eq $left -or $null -eq $right -or [math]::Abs([double]$left - [double]$right) -gt 0.02) {
            $offlineReconciliationErrors.Add([pscustomobject]@{order_no=$orderNo;date=$offline.date.ToString('yyyy-MM-dd');field=$field;offline_source=(Round-Money $left);income_sheet=(Round-Money $right);month=$ledger.month;row=$ledger.row})
        }
    }
}

$extraIncomeOrders = @($ledgerByOrder.Keys | Where-Object { -not $summaryByOrder.ContainsKey($_) -and -not $offlineByOrder.ContainsKey($_) } | Sort-Object)
$unreliableOrders = @($unreliableCostItems | Select-Object -ExpandProperty order_no -Unique | Sort-Object)
$verifiedPlatformOrders = @($summaryByOrder.Keys | Where-Object { $_ -notin $unreliableOrders } | Sort-Object)
$allIssues = $itemArithmeticErrors.Count + $expectedCostMismatches.Count + $orderReconciliationErrors.Count + $incomeSheetErrors.Count + $offlineReconciliationErrors.Count + $extraIncomeOrders.Count + $duplicateProductIds.Count

$report = [ordered]@{
    ok = ($allIssues -eq 0 -and $unreliableCostItems.Count -eq 0)
    generated_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    scope = '2026-08-01..2026-09-30 natural non-cancelled orders'
    counts = [ordered]@{
        natural_orders = $summaryByOrder.Count + $offlineByOrder.Count
        platform_natural_orders = $summaryByOrder.Count
        offline_orders = $offlineByOrder.Count
        order_item_rows = $auditItems.Count
        verified_orders = $verifiedPlatformOrders.Count + $(if($offlineReconciliationErrors.Count -eq 0){$offlineByOrder.Count}else{0})
        unreliable_orders = $unreliableOrders.Count
        unreliable_item_rows = $unreliableCostItems.Count
        large_loss_orders_profit_le_minus_10 = $largeLossOrders.Count
        item_arithmetic_errors = $itemArithmeticErrors.Count
        expected_cost_mismatches = $expectedCostMismatches.Count
        order_reconciliation_errors = $orderReconciliationErrors.Count
        income_sheet_errors = $incomeSheetErrors.Count
        preserved_manual_delivery_fees = $preservedManualDeliveryFees.Count
        offline_reconciliation_errors = $offlineReconciliationErrors.Count
        extra_income_orders = $extraIncomeOrders.Count
        duplicate_product_ids = $duplicateProductIds.Count
    }
    item_arithmetic_errors = @($itemArithmeticErrors)
    expected_cost_mismatches = @($expectedCostMismatches)
    order_reconciliation_errors = @($orderReconciliationErrors)
    income_sheet_errors = @($incomeSheetErrors)
    preserved_manual_delivery_fees = @($preservedManualDeliveryFees)
    offline_reconciliation_errors = @($offlineReconciliationErrors)
    extra_income_orders = $extraIncomeOrders
    duplicate_product_ids = @($duplicateProductIds)
    unreliable_cost_items = @($unreliableCostItems)
    large_loss_orders = @($largeLossOrders | Sort-Object profit)
}

$outputDirectory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) { $null = New-Item -ItemType Directory -Path $outputDirectory }
$report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
[ordered]@{ok=$report.ok;output_path=$OutputPath;counts=$report.counts;unreliable_orders=$unreliableOrders;large_loss_orders=@($largeLossOrders | Sort-Object profit | Select-Object order_no,date,daily_no,income,product_cost,delivery,profit,cost_source_risk)} | ConvertTo-Json -Depth 8
