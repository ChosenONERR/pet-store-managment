param(
    [Parameter(Mandatory=$true)][string]$StartDate,
    [Parameter(Mandatory=$true)][string]$EndDate,
    [switch]$Initialize,
    [string]$NaturalOrdersPath,
    [string]$InterventionOrdersPath,
    [string]$WorkRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $WorkRoot) { $WorkRoot = $PSScriptRoot }
$projectRoot = Split-Path $WorkRoot -Parent
if (-not $NaturalOrdersPath) { $NaturalOrdersPath = Join-Path $WorkRoot 'initial_natural_orders.json' }
if (-not $InterventionOrdersPath) { $InterventionOrdersPath = Join-Path $WorkRoot 'intervention_orders.json' }
$metaPath = Join-Path $WorkRoot 'simple_finance_meta.json'
$extractorPath = Join-Path $WorkRoot 'meituan_order_extract.mjs'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少简版工作簿元数据' }
if (-not (Test-Path -LiteralPath $extractorPath)) { throw '缺少美团订单提取脚本' }

$meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
if (-not $spreadsheetToken) { throw '无法解析飞书工作簿引用' }
. (Join-Path $projectRoot 'agent\finance_transaction.ps1')
. (Join-Path $projectRoot 'agent\intervention_record.ps1')
. (Join-Path $projectRoot 'agent\refund_adjustments.ps1')
. (Join-Path $projectRoot 'agent\order_actual_overrides.ps1')
$stateRoot = Join-Path $projectRoot 'agent\state'
$script:financeTransaction = Join-FinanceTransaction $stateRoot
$script:financeTransactionOwner = $false
$script:financeTransactionClosed = $false
if ($null -eq $script:financeTransaction) {
    $script:financeTransaction = New-FinanceTransaction $stateRoot "同步美团订单 $StartDate~$EndDate" $null -AcquireLock -WorkbookTokens @{source=$spreadsheetToken} -LarkCli 'lark-cli'
    $script:financeTransactionOwner = $true
}
function Close-OrderSyncTransaction([bool]$Succeeded, [string]$Reason = '') {
    if ($script:financeTransactionClosed) { return }
    if ($script:financeTransactionOwner) {
        if ($Succeeded) { Complete-FinanceTransaction $script:financeTransaction }
        else { $null = Invoke-FinanceTransactionRollback $script:financeTransaction @{source=$spreadsheetToken} 'lark-cli' $Reason }
    }
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
}
function Convert-ToExcelSerial([string]$DateText, [string]$TimeText = '00:00') {
    $local = [datetime]::ParseExact("$DateText $TimeText", 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    $base = [datetime]::SpecifyKind([datetime]'1899-12-30', [DateTimeKind]::Unspecified)
    return ($local - $base).TotalDays
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    $valueProperty = $cell.PSObject.Properties['value']
    if ($null -eq $valueProperty -or $null -eq $valueProperty.Value) { return $null }
    return [string]$valueProperty.Value
}

function Get-CellNumber($Row, [int]$Index) {
    $value = Get-CellValue $Row $Index
    if ($null -eq $value -or $value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse($value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $null
}

function Invoke-LarkJson([scriptblock]$Command, [string]$Label, [int]$MaxAttempts = 3) {
    $errors = [Collections.Generic.List[string]]::new()
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $global:LASTEXITCODE = 0
        try {
            $raw = & $Command | Out-String
            $exitCode = $LASTEXITCODE
            $result = $raw | ConvertFrom-Json
            if ($exitCode -eq 0 -and $result.ok -ne $false) { return $result }
            $errors.Add("第${attempt}次：退出码=$exitCode")
        } catch {
            $errors.Add("第${attempt}次：$($_.Exception.Message)")
        }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds ([math]::Min(8, $attempt * 2)) }
    }
    throw "$Label；重试${MaxAttempts}次仍失败：$($errors -join '；')"
}

function Get-SheetRows([string]$SheetName, [string]$Range) {
    $result = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --include value --max-chars 5000000 --format json } "读取飞书失败：$SheetName $Range"
    if (-not $result.ok -or $result.data.has_more) { throw "飞书读取结果不完整：$SheetName $Range" }
    return @($result.data.ranges[0].cells)
}

function Set-CellMatrix([string]$SheetName, [string]$Range, $Rows) {
    if (@($Rows).Count -eq 0) { return }
    Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $spreadsheetToken $SheetName $Range 'lark-cli'
    $payload = ConvertTo-Json -InputObject $Rows -Depth 12 -Compress
    $null = Invoke-LarkJson { $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --cells - --format json } "写入飞书失败：$SheetName $Range"
}

function Clear-CellRange([string]$SheetName, [string]$Range) {
    Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $spreadsheetToken $SheetName $Range 'lark-cli'
    $ranges = ConvertTo-Json -InputObject @("$SheetName!$Range") -Compress
    $null = Invoke-LarkJson {
        $ranges | lark-cli sheets +cells-batch-clear --as user --spreadsheet-token $spreadsheetToken --ranges - --scope content --yes --format json
    } "清除飞书旧尾行失败：$SheetName $Range"
}

function New-ValueCell($Value) {
    if ($null -eq $Value -or "$Value" -eq '') { return @{ value = '' } }
    return @{ value = $Value }
}

function New-NumberCell($Value) {
    if ($null -eq $Value -or "$Value" -eq '') { return @{ value = 0 } }
    return @{ value = $Value }
}

function Convert-DateDisplayToSerial($Value) {
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    $match = [regex]::Match([string]$Value, '^(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}))?')
    if (-not $match.Success) { return $Value }
    $timeText = if ($match.Groups[2].Success) { $match.Groups[2].Value } else { '00:00' }
    return Convert-ToExcelSerial $match.Groups[1].Value $timeText
}

try {
$global:LASTEXITCODE = 0
$extractLines = @(& node $extractorPath --start $StartDate --end $EndDate 2>&1 | ForEach-Object { [string]$_ })
if ($LASTEXITCODE -ne 0) {
    $detail = ($extractLines -join ' ').Trim()
    if ($detail.Length -gt 1500) { $detail = $detail.Substring(0, 1500) }
    throw "美团订单提取失败：$detail"
}
$extract = (($extractLines -join "`n").Trim()) | ConvertFrom-Json
if ($extract.total -ne @($extract.orders).Count) { throw '订单提取数量不一致' }
$incompleteOrders = @($extract.orders | Where-Object {
    $expectedIncome = if ($null -eq $_.expectedIncome) { 0.0 } else { [double]$_.expectedIncome }
    $expectedItemCount = if ($null -eq $_.expectedItemCount) { 0.0 } else { [double]$_.expectedItemCount }
    $actualItemCount = (@($_.items) | Measure-Object -Property quantity -Sum).Sum
    $expectedIncome -gt 0 -and (
        @($_.items).Count -eq 0 -or
        $null -eq $_.productSales -or
        $null -eq $_.customerPay -or
        ($expectedItemCount -gt 0 -and [math]::Abs([double]$actualItemCount - $expectedItemCount) -gt 0.000001)
    )
})
if ($incompleteOrders.Count -gt 0) {
    $ids = @($incompleteOrders | ForEach-Object { [string]$_.orderNo }) -join '、'
    throw "美团订单明细未完整展开，已在写入飞书前停止：$ids"
}
$extractedOrderNos = @($extract.orders | ForEach-Object { [string]$_.orderNo } | Where-Object { $_ })
$interventionIds = @()
if (Test-Path -LiteralPath $InterventionOrdersPath) {
    $interventionIds = @((Get-Content -LiteralPath $InterventionOrdersPath -Raw | ConvertFrom-Json).intervention_order_ids | ForEach-Object { [string]$_ } | Select-Object -Unique)
}
# Feishu is the live source of truth for newly reported intervention orders.
# The local JSON remains the historical seed used during rebuilds.
$expectedInterventionHeaders = @('日期','单号','补单人','订单编号','补单金额','配送方式','平台费用（自动）','配送费','总成本','是否已恢复库存','是否已好评','购买内容','备注')
$actualInterventionHeaderRow = @(Get-SheetRows '干预订单' 'A1:M1')[0]
for ($i = 0; $i -lt $expectedInterventionHeaders.Count; $i++) {
    if ((Get-CellValue $actualInterventionHeaderRow $i) -ne $expectedInterventionHeaders[$i]) {
        throw ('干预订单表头与自动同步约定不一致：第{0}列应为[{1}]' -f ($i + 1), $expectedInterventionHeaders[$i])
    }
}
$registeredInterventionRows = Get-SheetRows '干预订单' 'A2:M1001'
$interventionIds += @($registeredInterventionRows | ForEach-Object { Get-CellValue $_ 3 } | Where-Object { Test-ActiveInterventionOrderNo $_ })
$interventionIds = @($interventionIds | Select-Object -Unique)
$interventionSet = [Collections.Generic.HashSet[string]]::new([string[]]$interventionIds)
$syncTime = [datetime]::Now.ToString('yyyy-MM-dd HH:mm')
$syncSerial = Convert-ToExcelSerial ([datetime]::Now.ToString('yyyy-MM-dd')) ([datetime]::Now.ToString('HH:mm'))

# SKU成本主表只维护当前 KW 与当前 CSX 过渡码；旧 KW 不作为历史别名。
$skuCodeCandidates = @{}
$skuRows = Get-SheetRows 'SKU成本主表' 'A2:B3000'
foreach ($row in $skuRows) {
    $kw = Get-CellValue $row 0
    if (-not $kw -or $kw -notmatch '(?i)^kw\d+$') { continue }
    $kw = $kw.ToLowerInvariant()
    foreach ($code in @($kw) + @([regex]::Split([string](Get-CellValue $row 1), '[/,，、;；\s]+') | Where-Object { $_ -match '^(?i)csx\d+$' } | ForEach-Object { $_.ToLowerInvariant() })) {
        if (-not $skuCodeCandidates.ContainsKey($code)) { $skuCodeCandidates[$code] = @() }
        $skuCodeCandidates[$code] += $kw
    }
}
$skuCodeToKw = @{}
foreach ($code in $skuCodeCandidates.Keys) {
    $unique = @($skuCodeCandidates[$code] | Select-Object -Unique)
    if ($unique.Count -eq 1) { $skuCodeToKw[$code] = $unique[0] }
}

function Normalize-ProductText([string]$Value) {
    if (-not $Value) { return '' }
    $text = [regex]::Replace($Value.ToLowerInvariant(), '\s*[\(（].*$', '')
    return ([regex]::Replace($text, '[\s\p{P}\p{S}]+', ''))
}

$inventoryCandidates = @{}
foreach ($row in (Get-SheetRows '库存快照' 'B2:O2000')) {
    $productId = Get-CellValue $row 12
    if (-not $productId) { continue }
    $codes = @((Get-CellValue $row 0),(Get-CellValue $row 5)) | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique
    $candidate = [pscustomobject]@{product_id=$productId;title=(Normalize-ProductText (Get-CellValue $row 1));spec=(Normalize-ProductText (Get-CellValue $row 13))}
    foreach ($code in $codes) {
        if (-not $inventoryCandidates.ContainsKey($code)) { $inventoryCandidates[$code] = @() }
        $inventoryCandidates[$code] += $candidate
    }
}
function Resolve-ProductId([string]$GoodsCode,[string]$Name,[string]$Spec) {
    if (-not $GoodsCode) { return $null }
    $code = $GoodsCode.Trim().ToLowerInvariant()
    $candidates = @($inventoryCandidates[$code])
    if ($candidates.Count -eq 0) { return $null }
    $title = Normalize-ProductText $Name; $specKey = Normalize-ProductText $Spec
    if ($candidates.Count -eq 1) {
        $candidate = $candidates[0]
        if ($title -and $candidate.title -and $candidate.title -ne $title) { return $null }
        if ($specKey -and $candidate.spec -and $candidate.spec -ne $specKey) { return $null }
        return [string]$candidate.product_id
    }
    $matches = @($candidates | Where-Object { (!$title -or $_.title -eq $title) -and (!$specKey -or $_.spec -eq $specKey) })
    if ($matches.Count -eq 1) { return [string]$matches[0].product_id }
    if ($title) {
        $matches = @($candidates | Where-Object { $_.title -eq $title })
        if ($matches.Count -eq 1) { return [string]$matches[0].product_id }
    }
    return $null
}

function Resolve-KwCode([string]$GoodsCode) {
    if (-not $GoodsCode) { return $null }
    $normalized = $GoodsCode.Trim().ToLowerInvariant()
    if ($normalized -match '^kw\d+$') { return $normalized }
    if ($skuCodeToKw.ContainsKey($normalized)) { return $skuCodeToKw[$normalized] }
    # 无法唯一映射时保留订单页原货号，不把不同商品强行合并。
    return $normalized
}

function Format-InterventionContents($Items) {
    $quantityByKw = [ordered]@{}
    foreach ($item in @($Items)) {
        $kw = Resolve-KwCode ([string]$item.goodsCode)
        if (-not $kw) { continue }
        $quantity = if ($null -eq $item.quantity) { 0.0 } else { [double]$item.quantity }
        if (-not $quantityByKw.Contains($kw)) { $quantityByKw[$kw] = 0.0 }
        $quantityByKw[$kw] += $quantity
    }
    return (@($quantityByKw.GetEnumerator() | ForEach-Object {
        $qty = if ([math]::Abs($_.Value - [math]::Round($_.Value)) -lt 0.000001) { [string][int][math]::Round($_.Value) } else { $_.Value.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) }
        "$($_.Key)×$qty"
    }) -join '、')
}

$existingSummaryRows = Get-SheetRows '订单汇总' 'A2:Q2001'
$summaryById = [ordered]@{}
$existingSummaryLastRow = 1
for ($rowIndex = 0; $rowIndex -lt $existingSummaryRows.Count; $rowIndex++) {
    $row = $existingSummaryRows[$rowIndex]
    $orderNo = Get-CellValue $row 2
    if (-not $orderNo) { continue }
    $existingSummaryLastRow = $rowIndex + 2
    $summaryById[$orderNo] = [pscustomobject]@{
        OrderDate = Get-CellValue $row 0
        Platform = Get-CellValue $row 1
        OrderNo = $orderNo
        Status = Get-CellValue $row 3
        ProductSales = Get-CellNumber $row 5
        CustomerPay = Get-CellNumber $row 6
        ExpectedIncome = Get-CellNumber $row 7
        PlatformFees = Get-CellNumber $row 8
        DeliveryFee = Get-CellNumber $row 9
        ItemCount = Get-CellNumber $row 12
        Source = Get-CellValue $row 13
        SyncTime = Get-CellValue $row 14
        Note = Get-CellValue $row 15
        DailyOrderNo = Get-CellValue $row 16
        SortKey = "$(Get-CellValue $row 0)|$orderNo"
    }
}

# 退款信息来自美团订单卡片里的独立售后区。先以售后记录号持久化，后续每次同步都按绝对值重建，
# 而不是在现有金额上反复相减，因此重试、跨小时同步和事务回滚都不会重复扣账。
$refundMerge = Merge-RefundDiscoveries $projectRoot @($extract.orders)
$refundAccountingByOrder = @{}
foreach ($order in @($extract.orders)) {
    $refundAccountingByOrder[[string]$order.orderNo] = Get-RefundAccountingSummary $projectRoot ([string]$order.orderNo)
}
$orderActualByOrderNo = @{}
foreach ($override in @(Get-OrderActualOverrides $projectRoot)) {
    $orderActualByOrderNo[[string]$override.order_no] = $override
}

foreach ($order in @($extract.orders)) {
    $actualOverride = if ($orderActualByOrderNo.ContainsKey([string]$order.orderNo)) { $orderActualByOrderNo[[string]$order.orderNo] } else { $null }
    $effectiveItems = if ($null -ne $actualOverride) { @($actualOverride.items) } else { @($order.items) }
    $itemCount = ($effectiveItems | Measure-Object -Property quantity -Sum).Sum
    $existingSummary = if ($summaryById.Contains([string]$order.orderNo)) { $summaryById[[string]$order.orderNo] } else { $null }
    $manualDelivery = $null
    $manualDeliveryMarker = $null
    if ($existingSummary -and $existingSummary.Note -match '人工确认配送费：?\s*(\d+(?:\.\d+)?)\s*元?') {
        $manualDelivery = [double]::Parse($Matches[1], [Globalization.CultureInfo]::InvariantCulture)
        $manualDeliveryMarker = "人工确认配送费：$($manualDelivery.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture))元"
    }
    $refundAccounting = $refundAccountingByOrder[[string]$order.orderNo]
    $refundEvents = @($refundAccounting.events)
    $effectiveExpectedIncome = $order.expectedIncome
    $refundStatus = [string]$order.status
    $refundNoteParts = @()
    if ($refundEvents.Count -gt 0) {
        $baseExpectedIncome = @($refundEvents | Where-Object { $null -ne $_.base_expected_income } | Select-Object -First 1)
        if ($baseExpectedIncome.Count -gt 0) {
            $effectiveExpectedIncome = [math]::Max([double]0.0, [double][math]::Round(([double]$baseExpectedIncome[0].base_expected_income - [double]$refundAccounting.income_adjustment),2))
        } elseif ($null -ne $effectiveExpectedIncome) {
            $effectiveExpectedIncome = [math]::Max([double]0.0, [double][math]::Round(([double]$effectiveExpectedIncome - [double]$refundAccounting.income_adjustment),2))
        }
        $refundStatus = if ($refundEvents.Count -gt 0) { '已完成（部分退款）' } else { [string]$order.status }
        foreach ($event in $refundEvents) {
            $itemText = @($event.items | ForEach-Object {
                $dispositionText = switch ([string]$_.disposition) {
                    'returned_restocked' { '已退回入库' }
                    'refund_no_return' { '未退货，保留成本' }
                    'returned_damaged' { '退回报损，保留成本' }
                    default { '退货状态待确认' }
                }
                "$($_.goods_code)×$($_.quantity) $dispositionText"
            }) -join '、'
            $refundNoteParts += "退款补偿[$($event.record_id)]：顾客退款¥$([double]$event.customer_refund_amount)；$itemText"
        }
    }
    $effectiveDeliveryFee = if ($null -ne $manualDelivery) { $manualDelivery } else { $order.deliveryFee }
    $actualOverrideNote = $null
    if ($null -ne $actualOverride) {
        $effectiveExpectedIncome = [double]$actualOverride.actual_income
        $effectiveDeliveryFee = [double]$actualOverride.delivery_fee
        $refundStatus = '已完成（实际履约已调整）'
        $actualItemsText = @($effectiveItems | ForEach-Object {
            $qty = [double]$_.quantity
            $qtyText = if ([math]::Abs($qty - [math]::Round($qty)) -lt 0.000001) { [string][int][math]::Round($qty) } else { $qty.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture) }
            "$($_.goods_code)×$qtyText"
        }) -join '、'
        $actualOverrideNote = "实际履约覆盖：实收¥$(([double]$actualOverride.actual_income).ToString('0.00',[Globalization.CultureInfo]::InvariantCulture))；实际售出$actualItemsText；平台库存不由本系统修改"
        if ($actualOverride.note) { $actualOverrideNote += "；$($actualOverride.note)" }
    }
    $summaryById[[string]$order.orderNo] = [pscustomobject]@{
        OrderDate = [string]$order.orderDate
        Platform = '美团'
        OrderNo = [string]$order.orderNo
        Status = $refundStatus
        ProductSales = $order.productSales
        CustomerPay = $order.customerPay
        ExpectedIncome = $effectiveExpectedIncome
        PlatformFees = $order.platformFees
        DeliveryFee = $effectiveDeliveryFee
        ItemCount = $itemCount
        Source = '美团订单页'
        SyncTime = $syncTime
        Note = ((@("原始状态：$($order.rawStatus)",$manualDeliveryMarker,$actualOverrideNote) + $refundNoteParts | Where-Object { $_ }) -join '；')
        DailyOrderNo = if ($order.dailyOrderNo) { "美团$($order.dailyOrderNo)号单" } else { $null }
        SortKey = "$($order.orderDate) $($order.orderTime)|$($order.orderNo)"
    }
}

$summaryRecords = @($summaryById.Values | Sort-Object SortKey)
if ($summaryRecords.Count -gt 2000) { throw '订单汇总超过2000行容量' }
$summaryAD = @(); $summaryFJ = @(); $summaryMP = @(); $summaryQ = @()
foreach ($record in $summaryRecords) {
    $dateValue = Convert-DateDisplayToSerial $record.OrderDate
    $summaryAD += ,@((New-ValueCell $dateValue),(New-ValueCell $record.Platform),(New-ValueCell $record.OrderNo),(New-ValueCell $record.Status))
    $summaryFJ += ,@((New-NumberCell $record.ProductSales),(New-NumberCell $record.CustomerPay),(New-NumberCell $record.ExpectedIncome),(New-NumberCell $record.PlatformFees),(New-NumberCell $record.DeliveryFee))
    $syncValue = if ([string]$record.SyncTime -eq $syncTime) { $syncSerial } else { Convert-DateDisplayToSerial $record.SyncTime }
    $summaryMP += ,@((New-NumberCell $record.ItemCount),(New-ValueCell $record.Source),(New-ValueCell $syncValue),(New-ValueCell $record.Note))
    $summaryQ += ,@((New-ValueCell $record.DailyOrderNo))
}
$summaryLastRow = $summaryRecords.Count + 1
Set-CellMatrix '订单汇总' "A2:D$summaryLastRow" $summaryAD
Set-CellMatrix '订单汇总' "F2:J$summaryLastRow" $summaryFJ
Set-CellMatrix '订单汇总' "M2:P$summaryLastRow" $summaryMP
Set-CellMatrix '订单汇总' "Q2:Q$summaryLastRow" $summaryQ
if ($existingSummaryLastRow -gt $summaryLastRow) {
    $tailStart = $summaryLastRow + 1
    Clear-CellRange '订单汇总' "A${tailStart}:D$existingSummaryLastRow"
    Clear-CellRange '订单汇总' "F${tailStart}:J$existingSummaryLastRow"
    Clear-CellRange '订单汇总' "M${tailStart}:Q$existingSummaryLastRow"
}

$extractedIds = [Collections.Generic.HashSet[string]]::new([string[]]$extractedOrderNos)
$existingItemRows = Get-SheetRows '订单商品' 'A2:Q4001'
$itemRecords = @()
$existingSequence = @{}
$existingItemsLastRow = 1
for ($rowIndex = 0; $rowIndex -lt $existingItemRows.Count; $rowIndex++) {
    $row = $existingItemRows[$rowIndex]
    $orderNo = Get-CellValue $row 2
    if (-not $orderNo) { continue }
    $existingItemsLastRow = $rowIndex + 2
    if ($extractedIds.Contains($orderNo)) { continue }
    if (-not $existingSequence.ContainsKey($orderNo)) { $existingSequence[$orderNo] = 0 }
    $existingSequence[$orderNo]++
    $itemRecords += [pscustomobject]@{
        OrderDate = Get-CellValue $row 0; Platform = Get-CellValue $row 1; OrderNo = $orderNo; Status = Get-CellValue $row 3
        GoodsCode = Get-CellValue $row 5; Name = Get-CellValue $row 6; Spec = Get-CellValue $row 7; Quantity = Get-CellNumber $row 8
        UnitPrice = Get-CellNumber $row 9; SyncTime = Get-CellValue $row 14; ProductId = Get-CellValue $row 16; SortKey = "$(Get-CellValue $row 0)|$orderNo|$($existingSequence[$orderNo])"
    }
}
foreach ($order in @($extract.orders)) {
    $actualOverride = if ($orderActualByOrderNo.ContainsKey([string]$order.orderNo)) { $orderActualByOrderNo[[string]$order.orderNo] } else { $null }
    $effectiveItems = if ($null -ne $actualOverride) { @($actualOverride.items) } else { @($order.items) }
    $refundAccounting = $refundAccountingByOrder[[string]$order.orderNo]
    $remainingReturned = @{}
    foreach ($entry in $refundAccounting.returned_quantity_by_code.GetEnumerator()) { $remainingReturned[[string]$entry.Key] = [double]$entry.Value }
    foreach ($item in $effectiveItems) {
        $normalizedCode = $(if ($null -ne $actualOverride) { ([string]$item.goods_code).Trim().ToLowerInvariant() } else { ([string]$item.goodsCode).Trim().ToLowerInvariant() })
        $itemStatus = if ($null -ne $actualOverride) { '已完成（实际履约已调整）' } else { [string]$order.status }
        if ($null -eq $actualOverride -and $remainingReturned.ContainsKey($normalizedCode) -and [double]$remainingReturned[$normalizedCode] -gt 0) {
            $returnedQuantity = [math]::Min([double]$item.quantity, [double]$remainingReturned[$normalizedCode])
            $remainingReturned[$normalizedCode] = [double]$remainingReturned[$normalizedCode] - $returnedQuantity
            $itemStatus = if ([math]::Abs($returnedQuantity - [double]$item.quantity) -lt 0.000001) { '已退款退货' } else { '已完成（部分退货）' }
        } elseif ($null -eq $actualOverride -and @($refundAccounting.events).Count -gt 0) {
            $itemStatus = '已完成（部分退款）'
        }
        $itemRecords += [pscustomobject]@{
            OrderDate = [string]$order.orderDate; Platform = '美团'; OrderNo = [string]$order.orderNo; Status = $itemStatus
            GoodsCode = $(if ($null -ne $actualOverride) { [string]$item.goods_code } else { [string]$item.goodsCode }); Name = $(if ($null -ne $actualOverride) { [string]$item.title } else { [string]$item.name }); Spec = [string]$item.spec; Quantity = $item.quantity
            UnitPrice = $(if ($null -ne $actualOverride) { [double]$item.unit_price } else { $item.unitPrice }); SyncTime = $syncTime; ProductId = $(if ($null -ne $actualOverride) { [string]$item.product_id } else { Resolve-ProductId ([string]$item.goodsCode) ([string]$item.name) ([string]$item.spec) }); SortKey = "$($order.orderDate) $($order.orderTime)|$($order.orderNo)|$(if ($null -ne $actualOverride) {$item.line_index} else {$item.lineIndex})"
        }
    }
}
$itemRecords = @($itemRecords | Sort-Object SortKey)
if ($itemRecords.Count -gt 4000) { throw '订单商品超过4000行容量' }
$itemsAD = @(); $itemsFJ = @(); $itemsO = @(); $itemsQ = @()
foreach ($item in $itemRecords) {
    $itemsAD += ,@((New-ValueCell (Convert-DateDisplayToSerial $item.OrderDate)),(New-ValueCell $item.Platform),(New-ValueCell $item.OrderNo),(New-ValueCell $item.Status))
    $itemsFJ += ,@((New-ValueCell $item.GoodsCode),(New-ValueCell $item.Name),(New-ValueCell $item.Spec),(New-NumberCell $item.Quantity),(New-NumberCell $item.UnitPrice))
    $itemSync = if ([string]$item.SyncTime -eq $syncTime) { $syncSerial } else { Convert-DateDisplayToSerial $item.SyncTime }
    $itemsO += ,@((New-ValueCell $itemSync))
    $itemsQ += ,@((New-ValueCell $item.ProductId))
}
$itemsLastRow = $itemRecords.Count + 1
Set-CellMatrix '订单商品' "A2:D$itemsLastRow" $itemsAD
Set-CellMatrix '订单商品' "F2:J$itemsLastRow" $itemsFJ
Set-CellMatrix '订单商品' "O2:O$itemsLastRow" $itemsO
Set-CellMatrix '订单商品' "Q2:Q$itemsLastRow" $itemsQ
if ($existingItemsLastRow -gt $itemsLastRow) {
    $tailStart = $itemsLastRow + 1
    Clear-CellRange '订单商品' "A${tailStart}:D$existingItemsLastRow"
    Clear-CellRange '订单商品' "F${tailStart}:J$existingItemsLastRow"
    Clear-CellRange '订单商品' "O${tailStart}:O$existingItemsLastRow"
    Clear-CellRange '订单商品' "Q${tailStart}:Q$existingItemsLastRow"
}

$existingInterventionRows = Get-SheetRows '干预订单' 'A2:M1001'
$interventionRecords = [ordered]@{}
foreach ($row in $existingInterventionRows) {
    $id = Get-CellValue $row 3
    if (-not $id) { continue }
    $interventionRecords[$id] = [ordered]@{
        DateTime = Get-CellValue $row 0; DailyOrderNo = Get-CellValue $row 1
        Person = Get-CellValue $row 2; OrderNo = $id; Amount = Get-CellNumber $row 4
        DeliveryType = Get-CellValue $row 5; PlatformFee = Get-CellNumber $row 6
        DeliveryFee = Get-CellNumber $row 7; RestoreState = Get-CellValue $row 9
        Reviewed = Get-CellValue $row 10; Contents = Get-CellValue $row 11; Note = Get-CellValue $row 12
    }
}
foreach ($id in $interventionIds) {
    if (-not $interventionRecords.Contains($id)) {
        $interventionRecords[$id] = [ordered]@{DateTime=$null;DailyOrderNo=$null;Person=$null;OrderNo=$id;Amount=$null;DeliveryType=$null;PlatformFee=$null;DeliveryFee=$null;RestoreState='待恢复';Reviewed=$null;Contents=$null;Note='用户确认的历史干预订单'}
    }
}
$interventionUpdated = 0
foreach ($order in @($extract.orders)) {
    $id = [string]$order.orderNo
    if (-not $interventionSet.Contains($id)) { continue }
    if (-not $interventionRecords.Contains($id)) {
        $interventionRecords[$id] = [ordered]@{DateTime=$null;DailyOrderNo=$null;Person=$null;OrderNo=$id;Amount=$null;DeliveryType=$null;PlatformFee=$null;DeliveryFee=$null;RestoreState='待恢复';Reviewed=$null;Contents=$null;Note='用户确认的历史干预订单'}
    }
    $record = $interventionRecords[$id]
    $record.DateTime = Convert-ToExcelSerial ([string]$order.orderDate) ([string]$order.orderTime)
    $record.Contents = Format-InterventionContents $order.items
    if ($order.dailyOrderNo) { $record.DailyOrderNo = "美团$($order.dailyOrderNo)号单" }
    $record.Amount = $order.customerPay
    $platformFee = if ($null -eq $order.platformFees) { 0.0 } else { [double]$order.platformFees }
    $platformDeliveryFee = if ($null -eq $order.deliveryFee) { 0.0 } else { [double]$order.deliveryFee }
    $existingDeliveryFee = if ($null -eq $record.DeliveryFee) { 0.0 } else { [double]$record.DeliveryFee }
    $deliveryFee = if ($platformDeliveryFee -gt 0) { $platformDeliveryFee } else { $existingDeliveryFee }
    $record.PlatformFee = $platformFee
    $record.DeliveryFee = $deliveryFee
    if ($deliveryFee -gt 0) { $record.DeliveryType = '骑手配送' }
    elseif ($record.DeliveryType -in @('美团骑手','外部配送')) { $record.DeliveryType = '骑手配送' }
    elseif ($record.DeliveryType -notin @('自配送','骑手配送')) { $record.DeliveryType = $null }
    if (-not $record.RestoreState) { $record.RestoreState = if ($order.status -eq '已取消') { '已恢复' } else { '待恢复' } }
    if ($order.status -eq '已取消' -and $record.RestoreState -eq '待恢复') { $record.RestoreState = '已恢复' }
    $interventionUpdated++
}
$interventionCells = @()
foreach ($record in $interventionRecords.Values) {
    $rowNumber = $interventionCells.Count + 2
    $interventionCells += ,@(
        (New-ValueCell $record.DateTime),(New-ValueCell $record.DailyOrderNo),(New-ValueCell $record.Person),(New-ValueCell $record.OrderNo),
        (New-NumberCell $record.Amount),(New-ValueCell $record.DeliveryType),(New-NumberCell $record.PlatformFee),(New-NumberCell $record.DeliveryFee),
        @{formula=('=IF(D{0}="","",IFERROR(G{0}*1,0)+IFERROR(H{0}*1,0))' -f $rowNumber);cell_styles=@{number_format='#,##0.00'}},
        (New-ValueCell $record.RestoreState),(New-ValueCell $record.Reviewed),(New-ValueCell $record.Contents),(New-ValueCell $record.Note)
    )
}
if ($interventionCells.Count -gt 1000) { throw '干预订单表容量不足' }
if ($interventionCells.Count -gt 0) {
    $interventionLastRow = $interventionCells.Count + 1
    Set-CellMatrix '干预订单' "A2:M$interventionLastRow" $interventionCells
}
$interventionCount = @($extract.orders | Where-Object { $interventionSet.Contains([string]$_.orderNo) }).Count
$naturalCount = $extract.total - $interventionCount

$verifyTargets = @(
    @{Sheet='干预订单';Range='I2:I1001'},
    @{Sheet='订单汇总';Range='E2:E2001,K2:L2001'},
    @{Sheet='订单商品';Range='E2:E4001,K2:N4001,P2:P4001'},
    @{Sheet='库存快照';Range='M2:M2000'},
    @{Sheet='驾驶舱';Range='C2:C19'}
)
foreach ($target in $verifyTargets) {
    $verify = Invoke-LarkJson { lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name $target.Sheet --range $target.Range --format json --exit-on-error } "公式校验调用失败：$($target.Sheet)"
    if ($verify.data.status -ne 'success') { throw "公式存在错误：$($target.Sheet)" }
}

$finalResult = [pscustomobject]@{
    start_date = $StartDate
    end_date = $EndDate
    extracted_orders = $extract.total
    extracted_items = @($extract.orders | ForEach-Object { @($_.items).Count } | Measure-Object -Sum).Sum
    natural_orders = $naturalCount
    intervention_orders = $interventionCount
    intervention_updated = $interventionUpdated
    total_orders_in_workbook = $summaryRecords.Count
    total_items_in_workbook = $itemRecords.Count
    completed_refunds = @($refundMerge.state.events | Where-Object { [string]$_.order_date -ge $StartDate -and [string]$_.order_date -le $EndDate }).Count
    newly_detected_refunds = @($refundMerge.newly_detected).Count
    pending_refund_items = @($refundMerge.state.events | ForEach-Object { @($_.items) } | Where-Object { $_.disposition -eq 'pending' }).Count
    refund_orders = @($extract.orders | Where-Object { @($_.refunds).Count -gt 0 } | ForEach-Object { [string]$_.orderNo })
    pending_refund_details = @($refundMerge.state.events | Where-Object { [string]$_.order_no -in $extractedOrderNos } | ForEach-Object {
        $event = $_
        @($event.items | Where-Object { $_.disposition -eq 'pending' } | ForEach-Object {
            [pscustomobject]@{order_no=$event.order_no;record_id=$event.record_id;refund_amount=$event.customer_refund_amount;goods_code=$_.goods_code;title=$_.title;quantity=$_.quantity}
        })
    })
    applied_refund_details = @($refundMerge.state.events | Where-Object { [string]$_.order_no -in $extractedOrderNos } | ForEach-Object {
        $event = $_
        @($event.items | Where-Object { $_.disposition -ne 'pending' } | ForEach-Object {
            [pscustomobject]@{order_no=$event.order_no;record_id=$event.record_id;refund_amount=$event.customer_refund_amount;goods_code=$_.goods_code;title=$_.title;quantity=$_.quantity;disposition=$_.disposition}
        })
    })
    synced_at = $syncTime
}
Close-OrderSyncTransaction $true
$finalResult | ConvertTo-Json -Compress
} catch {
    $originalError = $_
    $reason = [string]$originalError.Exception.Message
    try {
        Close-OrderSyncTransaction $false $reason
    } catch {
        throw "$reason；订单同步事务回滚失败：$($_.Exception.Message)"
    }
    # 不使用 exit：让异常原样冒泡给调度器，调度器才能在飞书卡片中展示真实失败阶段与原因。
    throw $originalError
}
