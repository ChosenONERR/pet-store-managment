param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('overview','orders','order_items','pending_costs','interventions','sku','inventory','replenishments','income','expenses','profit','status')]
    [string]$Dataset,
    [string]$StartDate,
    [string]$EndDate,
    [string]$OrderNo,
    [string]$Code,
    [string]$Search,
    [ValidateRange(1,500)]
    [int]$Limit = 100
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$sourceMetaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
$ledgerMetaPath = Join-Path $ProjectRoot 'work\daily_ledger_meta.json'
$lastRunPath = Join-Path $PSScriptRoot 'state\last_run.json'
$transactionRoot = Join-Path $PSScriptRoot 'state\transactions'
. (Join-Path $PSScriptRoot 'intervention_record.ps1')

function Get-WorkbookToken([string]$MetaPath) {
    if (-not (Test-Path -LiteralPath $MetaPath)) { throw "缺少工作簿元数据：$MetaPath" }
    $meta = Get-Content -LiteralPath $MetaPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $meta.workbook_url) { throw "工作簿元数据缺少链接：$MetaPath" }
    return ([uri][string]$meta.workbook_url).Segments[-1].Trim('/')
}

function Invoke-LarkJson([scriptblock]$Action, [string]$FailureMessage) {
    $output = @(& $Action 2>&1 | ForEach-Object { [string]$_ })
    $exitCode = $LASTEXITCODE
    $text = $output -join "`n"
    try {
        $result = $text | ConvertFrom-Json
    } catch {
        $start = $text.IndexOf('{')
        $end = $text.LastIndexOf('}')
        if ($start -lt 0 -or $end -le $start) { throw "$FailureMessage：返回内容无法解析" }
        try { $result = $text.Substring($start, $end - $start + 1) | ConvertFrom-Json }
        catch { throw "$FailureMessage：返回内容无法解析" }
    }
    if ($exitCode -ne 0 -or -not $result.ok) {
        $detail = if ($result.error.message) { [string]$result.error.message } else { $FailureMessage }
        throw "$FailureMessage：$detail"
    }
    return $result
}

function Get-CellValue($Cell) {
    if ($null -eq $Cell -or $null -eq $Cell.value) { return $null }
    return $Cell.value
}

function Convert-DateValue($Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $number = 0.0
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and $number -gt 30000 -and $number -lt 90000) {
        return [datetime]::FromOADate($number).ToString('yyyy-MM-dd')
    }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [ref]$date)) { return $date.ToString('yyyy-MM-dd') }
    return [string]$Value
}

function Convert-Number($Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return 0.0 }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return 0.0
}

function Get-SheetTable([string]$Token, [string]$SheetName, [string]$Range) {
    $result = Invoke-LarkJson {
        lark-cli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $SheetName --range $Range --include value --max-chars 5000000 --format json
    } "读取飞书数据失败：$SheetName"
    if ($result.data.has_more) { throw "读取飞书数据不完整：$SheetName" }
    $rangeResult = @($result.data.ranges) | Select-Object -First 1
    if ($null -eq $rangeResult -or $rangeResult.truncated) { throw "读取飞书数据被截断：$SheetName" }
    $matrix = @($rangeResult.cells)
    if ($matrix.Count -eq 0) { return @() }
    $headers = @($matrix[0] | ForEach-Object { [string](Get-CellValue $_) })
    $rowIndices = @($rangeResult.row_indices)
    $records = [Collections.Generic.List[object]]::new()
    for ($rowOffset = 1; $rowOffset -lt $matrix.Count; $rowOffset++) {
        $cells = @($matrix[$rowOffset])
        $record = [ordered]@{ _row = if ($rowIndices.Count -gt $rowOffset) { [int]$rowIndices[$rowOffset] } else { $rowOffset + 1 } }
        $hasValue = $false
        for ($column = 0; $column -lt $headers.Count; $column++) {
            $header = $headers[$column].Trim()
            if (-not $header) { continue }
            $value = if ($column -lt $cells.Count) { Get-CellValue $cells[$column] } else { $null }
            if ($header -match '日期|时间' -and $null -ne $value) { $value = Convert-DateValue $value }
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { $hasValue = $true }
            $record[$header] = $value
        }
        if ($hasValue) { $records.Add([pscustomobject]$record) }
    }
    return @($records)
}

function Get-PropertyValue($Record, [string]$Name) {
    $property = $Record.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Test-DateRange($Record, [string]$DateColumn) {
    $value = [string](Get-PropertyValue $Record $DateColumn)
    if (-not $value) { return -not ($StartDate -or $EndDate) }
    if ($StartDate -and $value -lt $StartDate) { return $false }
    if ($EndDate -and $value -gt $EndDate) { return $false }
    return $true
}

function Test-Search($Record) {
    if (-not $Search) { return $true }
    $haystack = ($Record.PSObject.Properties | Where-Object { $_.Name -ne '_row' } | ForEach-Object { [string]$_.Value }) -join ' '
    return $haystack.IndexOf($Search, [StringComparison]::OrdinalIgnoreCase) -ge 0
}

function Limit-Rows([object[]]$Rows) {
    return @($Rows | Select-Object -First $Limit)
}

function Get-MonthKeys([datetime]$Start, [datetime]$End) {
    $months = [Collections.Generic.List[string]]::new()
    $cursor = [datetime]::new($Start.Year, $Start.Month, 1)
    $last = [datetime]::new($End.Year, $End.Month, 1)
    while ($cursor -le $last) {
        $months.Add($cursor.ToString('yyyy-MM'))
        $cursor = $cursor.AddMonths(1)
    }
    return @($months)
}

$sourceToken = Get-WorkbookToken $sourceMetaPath
$effectiveStart = if ($StartDate) { $StartDate } else { '2026-08-24' }
$effectiveEnd = if ($EndDate) { $EndDate } else { [datetime]::Today.ToString('yyyy-MM-dd') }
$payload = $null

switch ($Dataset) {
    'overview' {
        $orders = @(Get-SheetTable $sourceToken '订单汇总' 'A1:Q2001' | Where-Object { Test-DateRange $_ '下单日期' })
        $interventions = @(Get-SheetTable $sourceToken '干预订单' 'A1:M1001' | Where-Object {
            (Test-ActiveInterventionOrderNo ([string](Get-PropertyValue $_ '订单编号'))) -and (Test-DateRange $_ '日期')
        })
        $validOrders = @($orders | Where-Object { [string](Get-PropertyValue $_ '订单状态') -notmatch '取消' })
        $naturalOrders = @($validOrders | Where-Object { [string](Get-PropertyValue $_ '订单类型') -ne '干预单' })
        $interventionOrders = @($validOrders | Where-Object { [string](Get-PropertyValue $_ '订单类型') -eq '干预单' })
        $payload = [ordered]@{
            date_range = [ordered]@{ start=$effectiveStart; end=$effectiveEnd }
            platform_valid_orders = $validOrders.Count
            natural_orders = $naturalOrders.Count
            intervention_orders_in_order_book = $interventionOrders.Count
            intervention_records = $interventions.Count
            natural_estimated_income = [math]::Round((($naturalOrders | ForEach-Object { Convert-Number (Get-PropertyValue $_ '商家预计收入') } | Measure-Object -Sum).Sum), 2)
            natural_product_cost = [math]::Round((($naturalOrders | ForEach-Object { Convert-Number (Get-PropertyValue $_ '商品成本') } | Measure-Object -Sum).Sum), 2)
            natural_delivery_fee = [math]::Round((($naturalOrders | ForEach-Object { Convert-Number (Get-PropertyValue $_ '配送费') } | Measure-Object -Sum).Sum), 2)
            natural_order_profit = [math]::Round((($naturalOrders | ForEach-Object { Convert-Number (Get-PropertyValue $_ '订单毛利') } | Measure-Object -Sum).Sum), 2)
            intervention_platform_fee = [math]::Round((($interventions | ForEach-Object { Convert-Number (Get-PropertyValue $_ '平台费用（自动）') } | Measure-Object -Sum).Sum), 2)
            intervention_delivery_fee = [math]::Round((($interventions | ForEach-Object { Convert-Number (Get-PropertyValue $_ '配送费') } | Measure-Object -Sum).Sum), 2)
            intervention_total_cost = [math]::Round((($interventions | ForEach-Object { Convert-Number (Get-PropertyValue $_ '总成本') } | Measure-Object -Sum).Sum), 2)
            pending_inventory_restore = @($interventions | Where-Object { [string](Get-PropertyValue $_ '是否已恢复库存') -ne '已恢复' }).Count
        }
    }
    'orders' {
        $rows = @(Get-SheetTable $sourceToken '订单汇总' 'A1:Q2001' | Where-Object {
            (Test-DateRange $_ '下单日期') -and
            (-not $OrderNo -or [string](Get-PropertyValue $_ '平台订单号') -eq $OrderNo) -and
            (Test-Search $_)
        })
        $payload = [ordered]@{ count=$rows.Count; rows=(Limit-Rows $rows) }
    }
    'order_items' {
        $rows = @(Get-SheetTable $sourceToken '订单商品' 'A1:Q4001' | Where-Object {
            (Test-DateRange $_ '下单日期') -and
            (-not $OrderNo -or [string](Get-PropertyValue $_ '平台订单号') -eq $OrderNo) -and
            (-not $Code -or [string](Get-PropertyValue $_ '货号') -ieq $Code -or [string](Get-PropertyValue $_ '商品ID（美团SKU ID）') -ieq $Code) -and
            (Test-Search $_)
        })
        $payload = [ordered]@{ count=$rows.Count; rows=(Limit-Rows $rows) }
    }
    'pending_costs' {
        $itemRows = @(Get-SheetTable $sourceToken '订单商品' 'A1:Q4001' | Where-Object {
            (Test-DateRange $_ '下单日期') -and
            (-not $OrderNo -or [string](Get-PropertyValue $_ '平台订单号') -eq $OrderNo) -and
            [string](Get-PropertyValue $_ '订单类型') -notmatch '干预' -and
            [string](Get-PropertyValue $_ '订单状态') -notmatch '取消' -and
            (Test-Search $_)
        })
        $skuRows = @(Get-SheetTable $sourceToken 'SKU成本主表' 'A1:Z3000')
        $skuByProductId = @{}
        $skuByCode = @{}
        foreach ($sku in $skuRows) {
            $productId = ([string](Get-PropertyValue $sku '商品ID（美团SKU ID）')).Trim()
            $kw = ([string](Get-PropertyValue $sku '当前KW')).Trim().ToLowerInvariant()
            $csx = ([string](Get-PropertyValue $sku 'CSX/过渡货号')).Trim().ToLowerInvariant()
            if ($productId) {
                if (-not $skuByProductId.ContainsKey($productId)) { $skuByProductId[$productId] = [Collections.Generic.List[object]]::new() }
                $skuByProductId[$productId].Add($sku)
            }
            $codes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            if ($kw) { $null = $codes.Add($kw) }
            foreach ($candidate in @([regex]::Split($csx, '[/,，、;；\s]+') | Where-Object { $_ })) { $null = $codes.Add($candidate.Trim()) }
            foreach ($candidate in $codes) {
                if (-not $skuByCode.ContainsKey($candidate)) { $skuByCode[$candidate] = [Collections.Generic.List[object]]::new() }
                $skuByCode[$candidate].Add($sku)
            }
        }
        $pending = [Collections.Generic.List[object]]::new()
        foreach ($item in $itemRows) {
            $unitCostRaw = Get-PropertyValue $item '单位成本'
            $itemCode = ([string](Get-PropertyValue $item '货号')).Trim().ToLowerInvariant()
            $productId = ([string](Get-PropertyValue $item '商品ID（美团SKU ID）')).Trim()
            $master = $null
            # Wrap the entire conditional in @(...). Windows PowerShell unwraps
            # one-item branch output into a scalar; then `.Count -eq 1` is false
            # and a uniquely matched SKU is incorrectly reported as missing.
            $productMatches = @(if ($productId -and $skuByProductId.ContainsKey($productId)) { $skuByProductId[$productId] })
            $codeMatches = @(if ($itemCode -and $skuByCode.ContainsKey($itemCode)) { $skuByCode[$itemCode] })
            if ($productMatches.Count -eq 1) { $master = $productMatches[0] }
            elseif ($productMatches.Count -eq 0 -and $codeMatches.Count -eq 1) { $master = $codeMatches[0] }
            $currentKw = if ($master) { ([string](Get-PropertyValue $master '当前KW')).Trim().ToLowerInvariant() } else { '' }
            $transitionCode = if ($master) { ([string](Get-PropertyValue $master 'CSX/过渡货号')).Trim().ToLowerInvariant() } else { '' }
            $actionCode = if ($master -and $productId -and $productMatches.Count -eq 1) { "sku_id:$productId" } elseif ($currentKw) { $currentKw } elseif ($transitionCode -and $codeMatches.Count -eq 1) { $transitionCode } elseif ($itemCode -match '^(?:kw|csx)\d+$' -and $codeMatches.Count -le 1) { $itemCode } else { $null }
            $costStatus = if ($master) { [string](Get-PropertyValue $master '成本核对状态') } elseif ($productMatches.Count -gt 1) { 'SKU主表商品ID重复待核对' } elseif ($codeMatches.Count -gt 1) { 'SKU主表货号重复待核对' } else { 'SKU主表未匹配' }
            $snapshotMissing = $null -eq $unitCostRaw -or [string]::IsNullOrWhiteSpace([string]$unitCostRaw)
            $masterPending = -not $master -or $costStatus -match '缺少|待|未确认|未匹配'
            if (-not $snapshotMissing -and -not $masterPending) { continue }
            $pendingReason = if ($snapshotMissing -and $masterPending) { '订单成本为空，SKU主表成本也未确认' } elseif ($snapshotMissing) { '订单成本快照为空' } else { '订单已有成本快照，但SKU主表成本仍未确认，需核对快照来源' }
            $pending.Add([pscustomobject][ordered]@{
                order_no = [string](Get-PropertyValue $item '平台订单号')
                order_date = [string](Get-PropertyValue $item '下单日期')
                platform = [string](Get-PropertyValue $item '平台')
                order_type = [string](Get-PropertyValue $item '订单类型')
                item_code = $itemCode
                current_kw = $currentKw
                transition_code = $transitionCode
                action_code = $actionCode
                product_id = $productId
                product_match_count = $productMatches.Count
                code_match_count = $codeMatches.Count
                title = [string](Get-PropertyValue $item '商品名称')
                specification = [string](Get-PropertyValue $item '规格')
                quantity = Convert-Number (Get-PropertyValue $item '数量')
                order_unit_cost = if ($snapshotMissing) { $null } else { Convert-Number $unitCostRaw }
                pending_reason = $pendingReason
                sku_master_row = if ($master) { [int](Get-PropertyValue $master '_row') } else { $null }
                baseline_unit_cost = if ($master) { Get-PropertyValue $master '基准单位成本' } else { $null }
                moving_average_cost = if ($master) { Get-PropertyValue $master '当前移动平均成本' } else { $null }
                cost_status = $costStatus
            })
        }
        $pendingRows = @($pending)
        $payload = [ordered]@{
            order_count = @($pendingRows | ForEach-Object { $_.order_no } | Select-Object -Unique).Count
            pending_item_count = $pendingRows.Count
            items = @(Limit-Rows $pendingRows)
        }
    }
    'interventions' {
        $rows = @(Get-SheetTable $sourceToken '干预订单' 'A1:M1001' | Where-Object {
            (Test-ActiveInterventionOrderNo ([string](Get-PropertyValue $_ '订单编号'))) -and
            (Test-DateRange $_ '日期') -and
            (-not $OrderNo -or [string](Get-PropertyValue $_ '订单编号') -eq $OrderNo) -and
            (Test-Search $_)
        })
        $payload = [ordered]@{
            count=$rows.Count
            platform_fee=[math]::Round((($rows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '平台费用（自动）') } | Measure-Object -Sum).Sum),2)
            delivery_fee=[math]::Round((($rows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '配送费') } | Measure-Object -Sum).Sum),2)
            total_cost=[math]::Round((($rows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '总成本') } | Measure-Object -Sum).Sum),2)
            rows=(Limit-Rows $rows)
        }
    }
    'sku' {
        $rows = @(Get-SheetTable $sourceToken 'SKU成本主表' 'A1:Z3000' | Where-Object {
            $codeMatch = -not $Code -or [string](Get-PropertyValue $_ '当前KW') -ieq $Code -or [string](Get-PropertyValue $_ 'CSX/过渡货号') -ieq $Code
            $codeMatch -and (Test-Search $_)
        })
        foreach ($row in $rows) {
            $rowNumber = [int](Get-PropertyValue $row '_row')
            $productId = ([string](Get-PropertyValue $row '商品ID（美团SKU ID）')).Trim().ToLowerInvariant()
            $actionCode = if ($rowNumber -ge 2 -and $productId) { "sku_ref:${rowNumber}:$productId" } else { $null }
            $row | Add-Member -NotePropertyName 'action_code' -NotePropertyValue $actionCode -Force
        }
        $payload = [ordered]@{ count=$rows.Count; rows=(Limit-Rows $rows) }
    }
    'inventory' {
        $rows = @(Get-SheetTable $sourceToken '库存快照' 'A1:O2000' | Where-Object {
            (-not $Code -or [string](Get-PropertyValue $_ '货号') -ieq $Code -or [string](Get-PropertyValue $_ '货架码') -ieq $Code) -and (Test-Search $_)
        })
        $payload = [ordered]@{ count=$rows.Count; rows=(Limit-Rows $rows) }
    }
    'replenishments' {
        $rows = @(Get-SheetTable $sourceToken '补货记录' 'A1:T1001' | Where-Object {
            (Test-DateRange $_ '补货日期') -and (-not $Code -or [string](Get-PropertyValue $_ '货号') -ieq $Code) -and (Test-Search $_)
        })
        $payload = [ordered]@{
            count=$rows.Count
            purchase_amount=[math]::Round((($rows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '实际采购金额') } | Measure-Object -Sum).Sum),2)
            payment_amount=[math]::Round((($rows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '付款金额') } | Measure-Object -Sum).Sum),2)
            rows=(Limit-Rows $rows)
        }
    }
    { $_ -in @('income','expenses') } {
        $ledgerToken = Get-WorkbookToken $ledgerMetaPath
        $start = [datetime]::ParseExact($effectiveStart, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $end = [datetime]::ParseExact($effectiveEnd, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $rows = [Collections.Generic.List[object]]::new()
        foreach ($month in (Get-MonthKeys $start $end)) {
            $monthNumber = [int]$month.Substring(5,2)
            $sheetName = if ($Dataset -eq 'income') { "${monthNumber}月收入" } else { "${monthNumber}月支出" }
            $range = if ($Dataset -eq 'income') { 'A1:J1001' } else { 'A1:F1001' }
            foreach ($row in (Get-SheetTable $ledgerToken $sheetName $range)) {
                if ((Test-DateRange $row '日期') -and (-not $OrderNo -or [string](Get-PropertyValue $row '订单编号') -eq $OrderNo) -and (Test-Search $row)) { $rows.Add($row) }
            }
        }
        $allRows = @($rows)
        if ($Dataset -eq 'income') {
            $payload = [ordered]@{
                count=$allRows.Count
                income=[math]::Round((($allRows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '扣佣后收入') } | Measure-Object -Sum).Sum),2)
                product_cost=[math]::Round((($allRows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '货物成本') } | Measure-Object -Sum).Sum),2)
                delivery_fee=[math]::Round((($allRows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '配送费') } | Measure-Object -Sum).Sum),2)
                profit=[math]::Round((($allRows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '这一单赚/亏') } | Measure-Object -Sum).Sum),2)
                rows=(Limit-Rows $allRows)
            }
        } else {
            $payload = [ordered]@{
                count=$allRows.Count
                total=[math]::Round((($allRows | ForEach-Object { Convert-Number (Get-PropertyValue $_ '费用') } | Measure-Object -Sum).Sum),2)
                rows=(Limit-Rows $allRows)
            }
        }
    }
    'profit' {
        $ledgerToken = Get-WorkbookToken $ledgerMetaPath
        $start = [datetime]::ParseExact($effectiveStart, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $end = [datetime]::ParseExact($effectiveEnd, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $incomeRows = [Collections.Generic.List[object]]::new()
        $expenseRows = [Collections.Generic.List[object]]::new()
        foreach ($month in (Get-MonthKeys $start $end)) {
            $monthNumber = [int]$month.Substring(5,2)
            foreach ($row in (Get-SheetTable $ledgerToken "${monthNumber}月收入" 'A1:J1001')) {
                if (Test-DateRange $row '日期') { $incomeRows.Add($row) }
            }
            foreach ($row in (Get-SheetTable $ledgerToken "${monthNumber}月支出" 'A1:F1001')) {
                if (Test-DateRange $row '日期') { $expenseRows.Add($row) }
            }
        }
        $income = @($incomeRows)
        $expenses = @($expenseRows)
        $interventionExpenses = @($expenses | Where-Object { [string](Get-PropertyValue $_ '项目') -match '补单|干预' })
        $promotionExpenses = @($expenses | Where-Object { [string](Get-PropertyValue $_ '项目') -match '推广|广告|竞价' })
        $excludedExpenses = @($expenses | Where-Object { [string](Get-PropertyValue $_ '项目') -match '采购|补货付款|进货|老板取款|老板人工|市场租金|折旧|摊销' })
        $otherOperatingExpenses = @($expenses | Where-Object {
            $project = [string](Get-PropertyValue $_ '项目')
            $project -notmatch '补单|干预|推广|广告|竞价|采购|补货付款|进货|老板取款|老板人工|市场租金|折旧|摊销'
        })
        $netIncome = [math]::Round((($income | ForEach-Object { Convert-Number (Get-PropertyValue $_ '扣佣后收入') } | Measure-Object -Sum).Sum),2)
        $productCost = [math]::Round((($income | ForEach-Object { Convert-Number (Get-PropertyValue $_ '货物成本') } | Measure-Object -Sum).Sum),2)
        $deliveryFee = [math]::Round((($income | ForEach-Object { Convert-Number (Get-PropertyValue $_ '配送费') } | Measure-Object -Sum).Sum),2)
        $interventionCost = [math]::Round((($interventionExpenses | ForEach-Object { Convert-Number (Get-PropertyValue $_ '费用') } | Measure-Object -Sum).Sum),2)
        $promotionCost = [math]::Round((($promotionExpenses | ForEach-Object { Convert-Number (Get-PropertyValue $_ '费用') } | Measure-Object -Sum).Sum),2)
        $otherCost = [math]::Round((($otherOperatingExpenses | ForEach-Object { Convert-Number (Get-PropertyValue $_ '费用') } | Measure-Object -Sum).Sum),2)
        $payload = [ordered]@{
            date_range = [ordered]@{ start=$effectiveStart; end=$effectiveEnd }
            completed_order_count = $income.Count
            net_order_income = $netIncome
            product_and_gift_cost = $productCost
            order_delivery_fee = $deliveryFee
            intervention_cost = $interventionCost
            promotion_cost = $promotionCost
            other_operating_expense = $otherCost
            excluded_cash_or_management_expense = [math]::Round((($excludedExpenses | ForEach-Object { Convert-Number (Get-PropertyValue $_ '费用') } | Measure-Object -Sum).Sum),2)
            operating_net_profit = [math]::Round($netIncome-$productCost-$deliveryFee-$interventionCost-$promotionCost-$otherCost,2)
            formula = '真实订单扣佣后收入-货物及赠品成本-真实订单配送费-干预单成本-推广费用-其他经营费用'
            note = '不扣老板人工、折旧摊销、市场租金机会成本、采购付款或老板取款；货物成本已含自然/线下订单赠品成本。'
            intervention_breakdown = @(Limit-Rows $interventionExpenses)
            promotion_breakdown = @(Limit-Rows $promotionExpenses)
            other_expense_breakdown = @(Limit-Rows $otherOperatingExpenses)
        }
    }
    'status' {
        $lastRun = if (Test-Path -LiteralPath $lastRunPath) { Get-Content -LiteralPath $lastRunPath -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }
        $transactions = @()
        if (Test-Path -LiteralPath $transactionRoot) {
            $transactions = @(Get-ChildItem -LiteralPath $transactionRoot -Filter 'TX-*.json' -File | Sort-Object LastWriteTime -Descending | Select-Object -First 5 | ForEach-Object {
                $tx = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                [ordered]@{ id=[string]$tx.transaction_id; status=[string]$tx.status; started_at=$tx.started_at; error=$tx.error }
            })
        }
        $payload = [ordered]@{
            last_run = if ($lastRun) { [ordered]@{ ok=$lastRun.ok; degraded=$lastRun.degraded; date=$lastRun.date; mode=$lastRun.mode; completed_at=$lastRun.completed_at; today_orders=$lastRun.today_orders; today_profit=$lastRun.today_profit; pending_cost_count=$lastRun.pending_cost_count; pending_delivery_count=$lastRun.pending_delivery_count; formula_verification=$lastRun.formula_verification; audit_ok=$lastRun.integrity_audit.ok } } else { $null }
            recent_transactions = $transactions
        }
    }
}

[ordered]@{
    ok = $true
    dataset = $Dataset
    filters = [ordered]@{ start_date=$StartDate; end_date=$EndDate; order_no=$OrderNo; code=$Code; search=$Search; limit=$Limit }
    data = $payload
    queried_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
} | ConvertTo-Json -Depth 12 -Compress
