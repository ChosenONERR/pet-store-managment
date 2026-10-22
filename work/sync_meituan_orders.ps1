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

function Get-OrderItemSnapshotKey([string]$OrderNo, [string]$ProductId, [string]$GoodsCode, [string]$Name, [string]$Spec) {
    $stableProductId = ([string]$ProductId).Trim().ToLowerInvariant()
    if ($stableProductId) { return "$OrderNo|sku:$stableProductId" }
    $normalizedCode = ([string]$GoodsCode).Trim().ToLowerInvariant()
    $normalizedName = [regex]::Replace(([string]$Name).Trim().ToLowerInvariant(), '\s+', '')
    $normalizedSpec = [regex]::Replace(([string]$Spec).Trim().ToLowerInvariant(), '\s+', '')
    return "$OrderNo|legacy:$normalizedCode|$normalizedName|$normalizedSpec"
}

function Get-OrderItemFallbackKey([string]$OrderNo, [string]$GoodsCode, [string]$Name, [string]$Spec) {
    $normalizedCode = ([string]$GoodsCode).Trim().ToLowerInvariant()
    $normalizedName = [regex]::Replace(([string]$Name).Trim().ToLowerInvariant(), '\s+', '')
    $normalizedSpec = [regex]::Replace(([string]$Spec).Trim().ToLowerInvariant(), '\s+', '')
    return "$OrderNo|fallback:$normalizedCode|$normalizedName|$normalizedSpec"
}

function Add-CostSnapshotToQueue($Queues, [string]$Key, $Snapshot) {
    if (-not $Queues.ContainsKey($Key)) { $Queues[$Key] = [Collections.Queue]::new() }
    $Queues[$Key].Enqueue($Snapshot)
}

function Take-CostSnapshotFromQueue($Queues, [string[]]$Keys) {
    foreach ($key in $Keys) {
        if (-not $key -or -not $Queues.ContainsKey($key)) { continue }
        $queue = $Queues[$key]
        while ($queue.Count -gt 0) {
            $candidate = $queue.Dequeue()
            if (-not $candidate.Consumed) {
                $candidate.Consumed = $true
                return $candidate
            }
        }
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

function New-NullableNumberCell($Value) {
    if ($null -eq $Value -or "$Value" -eq '') { return @{ value = '' } }
    return @{ value = $Value }
}

function Convert-ToComparableText($Value) {
    if ($null -eq $Value) { return '' }
    return [string]$Value
}

function Convert-ToComparableNumber($Value, [bool]$NullAsZero = $true) {
    if ($null -eq $Value -or [string]$Value -eq '') {
        if ($NullAsZero) { return '0' }
        return ''
    }
    $number = 0.0
    if (-not [double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        throw "无法比较非数值单元格：$Value"
    }
    return $number.ToString('R', [Globalization.CultureInfo]::InvariantCulture)
}

function Convert-ToComparableDate($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return '' }
    $serial = Convert-DateDisplayToSerial $Value
    return Convert-ToComparableNumber $serial $false
}

function Get-RecordsFingerprint($Records) {
    $lines = @($Records | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 6 -Compress } | Sort-Object)
    $payload = $lines -join "`n"
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload))) -replace '-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function New-SummaryBusinessRecord($Record) {
    return [ordered]@{
        order_date = Convert-ToComparableDate $Record.OrderDate
        platform = Convert-ToComparableText $Record.Platform
        order_no = Convert-ToComparableText $Record.OrderNo
        status = Convert-ToComparableText $Record.Status
        product_sales = Convert-ToComparableNumber $Record.ProductSales
        customer_pay = Convert-ToComparableNumber $Record.CustomerPay
        expected_income = Convert-ToComparableNumber $Record.ExpectedIncome
        platform_fees = Convert-ToComparableNumber $Record.PlatformFees
        delivery_fee = Convert-ToComparableNumber $Record.DeliveryFee
        item_count = Convert-ToComparableNumber $Record.ItemCount
        source = Convert-ToComparableText $Record.Source
        note = Convert-ToComparableText $Record.Note
        daily_order_no = Convert-ToComparableText $Record.DailyOrderNo
    }
}

function New-ItemBusinessRecord($Record) {
    return [ordered]@{
        order_date = Convert-ToComparableDate $Record.OrderDate
        platform = Convert-ToComparableText $Record.Platform
        order_no = Convert-ToComparableText $Record.OrderNo
        status = Convert-ToComparableText $Record.Status
        goods_code = Convert-ToComparableText $Record.GoodsCode
        name = Convert-ToComparableText $Record.Name
        spec = Convert-ToComparableText $Record.Spec
        quantity = Convert-ToComparableNumber $Record.Quantity
        unit_price = Convert-ToComparableNumber $Record.UnitPrice
        unit_cost = Convert-ToComparableNumber $Record.UnitCost $false
        item_cost = Convert-ToComparableNumber $Record.ItemCost $false
        margin = Convert-ToComparableNumber $Record.Margin $false
        product_id = Convert-ToComparableText $Record.ProductId
    }
}

function New-InterventionBusinessRecord($Record) {
    return [ordered]@{
        date_time = Convert-ToComparableDate $Record.DateTime
        daily_order_no = Convert-ToComparableText $Record.DailyOrderNo
        person = Convert-ToComparableText $Record.Person
        order_no = Convert-ToComparableText $Record.OrderNo
        amount = Convert-ToComparableNumber $Record.Amount $false
        delivery_type = Convert-ToComparableText $Record.DeliveryType
        platform_fee = Convert-ToComparableNumber $Record.PlatformFee $false
        delivery_fee = Convert-ToComparableNumber $Record.DeliveryFee $false
        restore_state = Convert-ToComparableText $Record.RestoreState
        reviewed = Convert-ToComparableText $Record.Reviewed
        contents = Convert-ToComparableText $Record.Contents
        note = Convert-ToComparableText $Record.Note
    }
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
$skuRowsWithHeader = @(Get-SheetRows 'SKU成本主表' 'A1:Q3000')
$skuHeaderRow = if ($skuRowsWithHeader.Count -gt 0) { $skuRowsWithHeader[0] } else { $null }
if ((Get-CellValue $skuHeaderRow 0) -ne '当前KW' -or (Get-CellValue $skuHeaderRow 16) -ne '商品ID（美团SKU ID）') {
    throw 'SKU成本主表结构不符合订单同步约定：A列应为[当前KW]，Q列应为[商品ID（美团SKU ID）]'
}
$skuRows = @($skuRowsWithHeader | Select-Object -Skip 1)
$skuProductIdCandidates = @{}
foreach ($row in $skuRows) {
    $kw = Get-CellValue $row 0
    if (-not $kw -or $kw -notmatch '(?i)^kw\d+$') { continue }
    $kw = $kw.ToLowerInvariant()
    $productId = Get-CellValue $row 16
    $rowCodes = @($kw) + @([regex]::Split([string](Get-CellValue $row 1), '[/,，、;；\s]+') | Where-Object { $_ -match '^(?i)csx\d+$' } | ForEach-Object { $_.ToLowerInvariant() })
    foreach ($code in $rowCodes) {
        if (-not $skuCodeCandidates.ContainsKey($code)) { $skuCodeCandidates[$code] = @() }
        $skuCodeCandidates[$code] += $kw
        if ($productId) {
            if (-not $skuProductIdCandidates.ContainsKey($code)) { $skuProductIdCandidates[$code] = @() }
            $skuProductIdCandidates[$code] += [string]$productId
        }
    }
}
$skuCodeToKw = @{}
foreach ($code in $skuCodeCandidates.Keys) {
    $unique = @($skuCodeCandidates[$code] | Select-Object -Unique)
    if ($unique.Count -eq 1) { $skuCodeToKw[$code] = $unique[0] }
}
$skuCodeToProductId = @{}
foreach ($code in $skuProductIdCandidates.Keys) {
    $unique = @($skuProductIdCandidates[$code] | Where-Object { $_ } | Select-Object -Unique)
    if ($unique.Count -eq 1) { $skuCodeToProductId[$code] = $unique[0] }
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
    # Windows PowerShell 5.1 会把 @($hashtable['missing']) 视为一个包含 $null 的元素，
    # 因而不能仅依赖 Count 判断是否找到候选项，否则读取 $null.title 会中断整批同步。
    $candidates = @()
    if ($inventoryCandidates.ContainsKey($code)) {
        $candidates = @($inventoryCandidates[$code] | Where-Object { $null -ne $_ })
    }
    if ($candidates.Count -eq 0) {
        # 新商品可能已进入 SKU 主表，但尚未出现在下一次库存快照中。只有商品 ID 唯一时才兜底，
        # 无法唯一确认则留空等待后续核对，绝不因单个 SKU 阻断其他订单。
        if ($skuCodeToProductId.ContainsKey($code)) { return [string]$skuCodeToProductId[$code] }
        return $null
    }
    $title = Normalize-ProductText $Name; $specKey = Normalize-ProductText $Spec
    if ($candidates.Count -eq 1) {
        $candidate = $candidates[0]
        $candidateTitle = if ($candidate.PSObject.Properties['title']) { [string]$candidate.title } else { '' }
        $candidateSpec = if ($candidate.PSObject.Properties['spec']) { [string]$candidate.spec } else { '' }
        $candidateProductId = if ($candidate.PSObject.Properties['product_id']) { [string]$candidate.product_id } else { '' }
        if (-not $candidateProductId) { return $null }
        if ($title -and $candidateTitle -and $candidateTitle -ne $title) { return $null }
        if ($specKey -and $candidateSpec -and $candidateSpec -ne $specKey) { return $null }
        return $candidateProductId
    }
    $matches = @($candidates | Where-Object {
        $candidateTitle = if ($_.PSObject.Properties['title']) { [string]$_.title } else { '' }
        $candidateSpec = if ($_.PSObject.Properties['spec']) { [string]$_.spec } else { '' }
        (!$title -or $candidateTitle -eq $title) -and (!$specKey -or $candidateSpec -eq $specKey)
    })
    if ($matches.Count -eq 1 -and $matches[0].PSObject.Properties['product_id']) { return [string]$matches[0].product_id }
    if ($title) {
        $matches = @($candidates | Where-Object {
            $_.PSObject.Properties['title'] -and [string]$_.title -eq $title
        })
        if ($matches.Count -eq 1 -and $matches[0].PSObject.Properties['product_id']) { return [string]$matches[0].product_id }
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
$existingSummaryBusinessRecords = [Collections.Generic.List[object]]::new()
$existingSummaryLastRow = 1
for ($rowIndex = 0; $rowIndex -lt $existingSummaryRows.Count; $rowIndex++) {
    $row = $existingSummaryRows[$rowIndex]
    $orderNo = Get-CellValue $row 2
    if (-not $orderNo) { continue }
    $existingSummaryLastRow = $rowIndex + 2
    $existingSummaryRecord = [pscustomobject]@{
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
    $summaryById[$orderNo] = $existingSummaryRecord
    $existingSummaryBusinessRecords.Add((New-SummaryBusinessRecord $existingSummaryRecord))
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
$changedOrderNumbers = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

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
    $desiredSummary = [pscustomobject]@{
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
    $summaryById[[string]$order.orderNo] = $desiredSummary
    $existingComparable = if ($null -ne $existingSummary) { ConvertTo-Json -InputObject (New-SummaryBusinessRecord $existingSummary) -Depth 6 -Compress } else { '' }
    $desiredComparable = ConvertTo-Json -InputObject (New-SummaryBusinessRecord $desiredSummary) -Depth 6 -Compress
    if ($existingComparable -ne $desiredComparable) { $null = $changedOrderNumbers.Add([string]$order.orderNo) }
}

$summaryRecords = @($summaryById.Values | Sort-Object SortKey)
if ($summaryRecords.Count -gt 2000) { throw '订单汇总超过2000行容量' }
$desiredSummaryBusinessRecords = @($summaryRecords | ForEach-Object { New-SummaryBusinessRecord $_ })
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

$extractedIds = [Collections.Generic.HashSet[string]]::new([string[]]$extractedOrderNos)
$existingItemRows = Get-SheetRows '订单商品' 'A2:Q4001'
$itemRecords = @()
$existingItemBusinessRecords = [Collections.Generic.List[object]]::new()
$existingItemBusinessByOrder = @{}
$existingSequence = @{}
$existingCostSnapshots = @{}
$existingItemsLastRow = 1
for ($rowIndex = 0; $rowIndex -lt $existingItemRows.Count; $rowIndex++) {
    $row = $existingItemRows[$rowIndex]
    $orderNo = Get-CellValue $row 2
    if (-not $orderNo) { continue }
    $existingItemsLastRow = $rowIndex + 2
    $snapshotKey = Get-OrderItemSnapshotKey $orderNo (Get-CellValue $row 16) (Get-CellValue $row 5) (Get-CellValue $row 6) (Get-CellValue $row 7)
    $fallbackSnapshotKey = Get-OrderItemFallbackKey $orderNo (Get-CellValue $row 5) (Get-CellValue $row 6) (Get-CellValue $row 7)
    $costSnapshot = [pscustomobject]@{
        UnitCost = Get-CellNumber $row 11
        ItemCost = Get-CellNumber $row 12
        Margin = Get-CellNumber $row 13
        Consumed = $false
    }
    Add-CostSnapshotToQueue $existingCostSnapshots $snapshotKey $costSnapshot
    if ($fallbackSnapshotKey -ne $snapshotKey) { Add-CostSnapshotToQueue $existingCostSnapshots $fallbackSnapshotKey $costSnapshot }
    $existingItemBusinessRecord = New-ItemBusinessRecord ([pscustomobject]@{
        OrderDate=Get-CellValue $row 0;Platform=Get-CellValue $row 1;OrderNo=$orderNo;Status=Get-CellValue $row 3
        GoodsCode=Get-CellValue $row 5;Name=Get-CellValue $row 6;Spec=Get-CellValue $row 7;Quantity=Get-CellNumber $row 8
        UnitPrice=Get-CellNumber $row 9;UnitCost=Get-CellNumber $row 11;ItemCost=Get-CellNumber $row 12;Margin=Get-CellNumber $row 13
        ProductId=Get-CellValue $row 16
    })
    $existingItemBusinessRecords.Add($existingItemBusinessRecord)
    if (-not $existingItemBusinessByOrder.ContainsKey($orderNo)) { $existingItemBusinessByOrder[$orderNo] = @() }
    $existingItemBusinessByOrder[$orderNo] += $existingItemBusinessRecord
    if ($extractedIds.Contains($orderNo)) { continue }
    if (-not $existingSequence.ContainsKey($orderNo)) { $existingSequence[$orderNo] = 0 }
    $existingSequence[$orderNo]++
    $itemRecords += [pscustomobject]@{
        OrderDate = Get-CellValue $row 0; Platform = Get-CellValue $row 1; OrderNo = $orderNo; Status = Get-CellValue $row 3
        GoodsCode = Get-CellValue $row 5; Name = Get-CellValue $row 6; Spec = Get-CellValue $row 7; Quantity = Get-CellNumber $row 8
        UnitPrice = Get-CellNumber $row 9; UnitCost = Get-CellNumber $row 11; ItemCost = Get-CellNumber $row 12; Margin = Get-CellNumber $row 13
        SyncTime = Get-CellValue $row 14; ProductId = Get-CellValue $row 16; SortKey = "$(Get-CellValue $row 0)|$orderNo|$($existingSequence[$orderNo])"
    }
}
$newItemOrderNumbers = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($order in @($extract.orders)) {
    $actualOverride = if ($orderActualByOrderNo.ContainsKey([string]$order.orderNo)) { $orderActualByOrderNo[[string]$order.orderNo] } else { $null }
    $effectiveItems = if ($null -ne $actualOverride) { @($actualOverride.items) } else { @($order.items) }
    $refundAccounting = $refundAccountingByOrder[[string]$order.orderNo]
    $remainingReturned = @{}
    foreach ($entry in $refundAccounting.returned_quantity_by_code.GetEnumerator()) { $remainingReturned[[string]$entry.Key] = [double]$entry.Value }
    foreach ($item in $effectiveItems) {
        $normalizedCode = $(if ($null -ne $actualOverride) { ([string]$item.goods_code).Trim().ToLowerInvariant() } else { ([string]$item.goodsCode).Trim().ToLowerInvariant() })
        $itemName = $(if ($null -ne $actualOverride) { [string]$item.title } else { [string]$item.name })
        $itemSpec = [string]$item.spec
        $itemProductId = $(if ($null -ne $actualOverride) { [string]$item.product_id } else { Resolve-ProductId ([string]$item.goodsCode) ([string]$item.name) ([string]$item.spec) })
        $snapshotKey = Get-OrderItemSnapshotKey ([string]$order.orderNo) $itemProductId $normalizedCode $itemName $itemSpec
        $fallbackSnapshotKey = Get-OrderItemFallbackKey ([string]$order.orderNo) $normalizedCode $itemName $itemSpec
        $priorSnapshot = Take-CostSnapshotFromQueue $existingCostSnapshots @($snapshotKey,$fallbackSnapshotKey)
        if ($null -eq $priorSnapshot) {
            $null = $newItemOrderNumbers.Add([string]$order.orderNo)
        }
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
            GoodsCode = $(if ($null -ne $actualOverride) { [string]$item.goods_code } else { [string]$item.goodsCode }); Name = $itemName; Spec = $itemSpec; Quantity = $item.quantity
            UnitPrice = $(if ($null -ne $actualOverride) { [double]$item.unit_price } else { $item.unitPrice })
            UnitCost = $(if ($null -ne $priorSnapshot) { $priorSnapshot.UnitCost } else { $null })
            ItemCost = $(if ($null -ne $priorSnapshot) { $priorSnapshot.ItemCost } else { $null })
            Margin = $(if ($null -ne $priorSnapshot) { $priorSnapshot.Margin } else { $null })
            SyncTime = $syncTime; ProductId = $itemProductId; SortKey = "$($order.orderDate) $($order.orderTime)|$($order.orderNo)|$(if ($null -ne $actualOverride) {$item.line_index} else {$item.lineIndex})"
        }
    }
}
$itemRecords = @($itemRecords | Sort-Object SortKey)
if ($itemRecords.Count -gt 4000) { throw '订单商品超过4000行容量' }
$desiredItemBusinessRecords = @($itemRecords | ForEach-Object { New-ItemBusinessRecord $_ })
$desiredItemBusinessByOrder = @{}
foreach ($record in $desiredItemBusinessRecords) {
    $orderNo = [string]$record.order_no
    if (-not $desiredItemBusinessByOrder.ContainsKey($orderNo)) { $desiredItemBusinessByOrder[$orderNo] = @() }
    $desiredItemBusinessByOrder[$orderNo] += $record
}
foreach ($orderNo in $extractedOrderNos) {
    $existingFingerprint = Get-RecordsFingerprint $(if ($existingItemBusinessByOrder.ContainsKey($orderNo)) { @($existingItemBusinessByOrder[$orderNo]) } else { @() })
    $desiredFingerprint = Get-RecordsFingerprint $(if ($desiredItemBusinessByOrder.ContainsKey($orderNo)) { @($desiredItemBusinessByOrder[$orderNo]) } else { @() })
    if ($existingFingerprint -ne $desiredFingerprint) { $null = $changedOrderNumbers.Add([string]$orderNo) }
}
$itemsAD = @(); $itemsFJ = @(); $itemsLN = @(); $itemsO = @(); $itemsQ = @()
foreach ($item in $itemRecords) {
    $itemsAD += ,@((New-ValueCell (Convert-DateDisplayToSerial $item.OrderDate)),(New-ValueCell $item.Platform),(New-ValueCell $item.OrderNo),(New-ValueCell $item.Status))
    $itemsFJ += ,@((New-ValueCell $item.GoodsCode),(New-ValueCell $item.Name),(New-ValueCell $item.Spec),(New-NumberCell $item.Quantity),(New-NumberCell $item.UnitPrice))
    $itemsLN += ,@((New-NullableNumberCell $item.UnitCost),(New-NullableNumberCell $item.ItemCost),(New-NullableNumberCell $item.Margin))
    $itemSync = if ([string]$item.SyncTime -eq $syncTime) { $syncSerial } else { Convert-DateDisplayToSerial $item.SyncTime }
    $itemsO += ,@((New-ValueCell $itemSync))
    $itemsQ += ,@((New-ValueCell $item.ProductId))
}
$itemsLastRow = $itemRecords.Count + 1

$existingInterventionRows = Get-SheetRows '干预订单' 'A2:M1001'
$interventionRecords = [ordered]@{}
$existingInterventionBusinessRecords = [Collections.Generic.List[object]]::new()
foreach ($row in $existingInterventionRows) {
    $id = Get-CellValue $row 3
    if (-not $id) { continue }
    $existingInterventionRecord = [pscustomobject][ordered]@{
        DateTime = Get-CellValue $row 0; DailyOrderNo = Get-CellValue $row 1
        Person = Get-CellValue $row 2; OrderNo = $id; Amount = Get-CellNumber $row 4
        DeliveryType = Get-CellValue $row 5; PlatformFee = Get-CellNumber $row 6
        DeliveryFee = Get-CellNumber $row 7; RestoreState = Get-CellValue $row 9
        Reviewed = Get-CellValue $row 10; Contents = Get-CellValue $row 11; Note = Get-CellValue $row 12
    }
    $interventionRecords[$id] = $existingInterventionRecord
    $existingInterventionBusinessRecords.Add((New-InterventionBusinessRecord $existingInterventionRecord))
}
foreach ($id in $interventionIds) {
    if (-not $interventionRecords.Contains($id)) {
        $interventionRecords[$id] = [ordered]@{DateTime=$null;DailyOrderNo=$null;Person=$null;OrderNo=$id;Amount=$null;DeliveryType=$null;PlatformFee=$null;DeliveryFee=$null;RestoreState='待恢复';Reviewed=$null;Contents=$null;Note='用户确认的历史干预订单'}
    }
}
$interventionUpdated = 0
$interventionRepairs = [Collections.Generic.List[object]]::new()
$interventionIssues = [Collections.Generic.List[object]]::new()
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
    $platformFee = if ($null -eq $order.platformFees) { $null } else { [double]$order.platformFees }
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

# A reported intervention may be registered after its original order fell out of
# the current extraction date. Reconcile every intervention against the durable
# order summary/items already in Feishu, instead of leaving known values blank or
# converting them to a misleading zero in the monthly expense sheet.
$itemsByOrder = @{}
foreach ($itemRecord in $itemRecords) {
    $itemOrderNo = [string]$itemRecord.OrderNo
    if (-not $itemOrderNo) { continue }
    if (-not $itemsByOrder.ContainsKey($itemOrderNo)) { $itemsByOrder[$itemOrderNo] = @() }
    $itemsByOrder[$itemOrderNo] += $itemRecord
}
foreach ($record in $interventionRecords.Values) {
    $orderNo = [string]$record.OrderNo
    if (-not (Test-ActiveInterventionOrderNo $orderNo)) { continue }
    $repairedFields = [Collections.Generic.List[string]]::new()
    if ($summaryById.Contains($orderNo)) {
        $summary = $summaryById[$orderNo]
        if ($summary.OrderDate -and (Convert-ToComparableDate $record.DateTime) -ne (Convert-ToComparableDate $summary.OrderDate)) {
            $record.DateTime = $summary.OrderDate
            $repairedFields.Add('日期')
        }
        if ($summary.DailyOrderNo -and [string]$record.DailyOrderNo -ne [string]$summary.DailyOrderNo) {
            $record.DailyOrderNo = $summary.DailyOrderNo
            $repairedFields.Add('当日单号')
        }
        if ($null -ne $summary.CustomerPay -and ($null -eq $record.Amount -or [math]::Abs([double]$record.Amount - [double]$summary.CustomerPay) -gt 0.005)) {
            $record.Amount = [double]$summary.CustomerPay
            $repairedFields.Add('补单金额')
        }
        if ($null -ne $summary.PlatformFees -and ($null -eq $record.PlatformFee -or [math]::Abs([double]$record.PlatformFee - [double]$summary.PlatformFees) -gt 0.005)) {
            $record.PlatformFee = [double]$summary.PlatformFees
            $repairedFields.Add('平台费用')
        }
        if (-not $record.Contents -and $itemsByOrder.ContainsKey($orderNo)) {
            $contents = Format-InterventionContents @($itemsByOrder[$orderNo])
            if ($contents) {
                $record.Contents = $contents
                $repairedFields.Add('购买内容')
            }
        }
    }
    if ($repairedFields.Count -gt 0) {
        $interventionRepairs.Add([pscustomobject]@{order_no=$orderNo;fields=@($repairedFields)})
        $null = $changedOrderNumbers.Add($orderNo)
    }

    $missingFields = [Collections.Generic.List[string]]::new()
    if (-not $summaryById.Contains($orderNo)) {
        $missingFields.Add('订单汇总中找不到原订单')
    } else {
        $summary = $summaryById[$orderNo]
        if (-not $record.DateTime) { $missingFields.Add('日期') }
        if (-not $record.DailyOrderNo) { $missingFields.Add('当日单号') }
        if ($null -eq $record.Amount) { $missingFields.Add('补单金额') }
        if ($null -eq $record.PlatformFee) { $missingFields.Add('平台费用') }
        if ([double]$summary.ItemCount -gt 0 -and -not $record.Contents) { $missingFields.Add('购买内容') }
    }
    if (-not $record.DeliveryType) { $missingFields.Add('配送方式/配送费确认') }
    if ($missingFields.Count -gt 0) {
        $interventionIssues.Add([pscustomobject]@{order_no=$orderNo;missing_fields=@($missingFields)})
    }
}
$interventionCells = @()
foreach ($record in $interventionRecords.Values) {
    $rowNumber = $interventionCells.Count + 2
    $interventionCells += ,@(
        (New-ValueCell $record.DateTime),(New-ValueCell $record.DailyOrderNo),(New-ValueCell $record.Person),(New-ValueCell $record.OrderNo),
        (New-NullableNumberCell $record.Amount),(New-ValueCell $record.DeliveryType),(New-NullableNumberCell $record.PlatformFee),(New-NullableNumberCell $record.DeliveryFee),
        @{formula=('=IF(D{0}="","",IF(OR(G{0}="",H{0}=""),"",G{0}*1+H{0}*1))' -f $rowNumber);cell_styles=@{number_format='#,##0.00'}},
        (New-ValueCell $record.RestoreState),(New-ValueCell $record.Reviewed),(New-ValueCell $record.Contents),(New-ValueCell $record.Note)
    )
}
if ($interventionCells.Count -gt 1000) { throw '干预订单表容量不足' }
$desiredInterventionBusinessRecords = @($interventionRecords.Values | ForEach-Object { New-InterventionBusinessRecord $_ })
$summaryChanged = (Get-RecordsFingerprint $existingSummaryBusinessRecords) -ne (Get-RecordsFingerprint $desiredSummaryBusinessRecords)
$itemsChanged = (Get-RecordsFingerprint $existingItemBusinessRecords) -ne (Get-RecordsFingerprint $desiredItemBusinessRecords)
$interventionsChanged = (Get-RecordsFingerprint $existingInterventionBusinessRecords) -ne (Get-RecordsFingerprint $desiredInterventionBusinessRecords)
$dataChanged = [bool]$Initialize -or $summaryChanged -or $itemsChanged -or $interventionsChanged
$affectedMonthKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($orderNo in $changedOrderNumbers) {
    if (-not $summaryById.Contains($orderNo)) { continue }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$summaryById[$orderNo].OrderDate,[ref]$date)) { $null = $affectedMonthKeys.Add($date.ToString('yyyy-MM')) }
}

if ($dataChanged) {
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

    Set-CellMatrix '订单商品' "A2:D$itemsLastRow" $itemsAD
    Set-CellMatrix '订单商品' "F2:J$itemsLastRow" $itemsFJ
    Set-CellMatrix '订单商品' "L2:N$itemsLastRow" $itemsLN
    Set-CellMatrix '订单商品' "O2:O$itemsLastRow" $itemsO
    Set-CellMatrix '订单商品' "Q2:Q$itemsLastRow" $itemsQ
    if ($existingItemsLastRow -gt $itemsLastRow) {
        $tailStart = $itemsLastRow + 1
        Clear-CellRange '订单商品' "A${tailStart}:D$existingItemsLastRow"
        Clear-CellRange '订单商品' "F${tailStart}:J$existingItemsLastRow"
        Clear-CellRange '订单商品' "L${tailStart}:N$existingItemsLastRow"
        Clear-CellRange '订单商品' "O${tailStart}:O$existingItemsLastRow"
        Clear-CellRange '订单商品' "Q${tailStart}:Q$existingItemsLastRow"
    }

    if ($interventionCells.Count -gt 0) {
        $interventionLastRow = $interventionCells.Count + 1
        Set-CellMatrix '干预订单' "A2:M$interventionLastRow" $interventionCells
    }
}
$interventionCount = @($extract.orders | Where-Object { $interventionSet.Contains([string]$_.orderNo) }).Count
$naturalCount = $extract.total - $interventionCount

$verifyTargets = if ($dataChanged) { @(
    @{Sheet='干预订单';Range='I2:I1001'},
    @{Sheet='订单汇总';Range='E2:E2001,K2:L2001'},
    @{Sheet='订单商品';Range='E2:E4001,K2:N4001,P2:P4001'},
    @{Sheet='库存快照';Range='M2:M2000'},
    @{Sheet='驾驶舱';Range='C2:C19'}
) } else { @() }
foreach ($target in $verifyTargets) {
    $verify = Invoke-LarkJson { lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name $target.Sheet --range $target.Range --format json --exit-on-error } "公式校验调用失败：$($target.Sheet)"
    if ($verify.data.status -ne 'success') { throw "公式存在错误：$($target.Sheet)" }
}

$finalResult = [pscustomobject]@{
    ok = $true
    start_date = $StartDate
    end_date = $EndDate
    extracted_orders = $extract.total
    extracted_items = @($extract.orders | ForEach-Object { @($_.items).Count } | Measure-Object -Sum).Sum
    natural_orders = $naturalCount
    intervention_orders = $interventionCount
    intervention_updated = $interventionUpdated
    intervention_repair_count = $interventionRepairs.Count
    intervention_repairs = @($interventionRepairs)
    intervention_issue_count = $interventionIssues.Count
    intervention_issues = @($interventionIssues)
    total_orders_in_workbook = $summaryRecords.Count
    total_items_in_workbook = $itemRecords.Count
    completed_refunds = @($refundMerge.state.events | Where-Object { [string]$_.order_date -ge $StartDate -and [string]$_.order_date -le $EndDate }).Count
    newly_detected_refunds = @($refundMerge.newly_detected).Count
    pending_refund_items = @($refundMerge.state.events | ForEach-Object { @($_.items) } | Where-Object { $_.disposition -eq 'pending' }).Count
    data_changed = [bool]$dataChanged
    writes_skipped = -not [bool]$dataChanged
    change_details = [ordered]@{summary=[bool]$summaryChanged;items=[bool]$itemsChanged;interventions=[bool]$interventionsChanged;initialize_forced=[bool]$Initialize}
    changed_order_nos = @($changedOrderNumbers | Sort-Object)
    affected_months = @($affectedMonthKeys | Sort-Object)
    new_item_order_nos = @($newItemOrderNumbers | Sort-Object)
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
