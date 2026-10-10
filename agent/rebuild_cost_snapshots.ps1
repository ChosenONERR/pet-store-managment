param(
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少经营驾驶舱元数据' }
$meta = [System.IO.File]::ReadAllText($metaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$larkCli = Join-Path $PSScriptRoot 'lark_cli_noexit.ps1'
. (Join-Path $PSScriptRoot 'finance_transaction.ps1')
. (Join-Path $PSScriptRoot 'sku_master_schema.ps1')
. (Join-Path $PSScriptRoot 'intervention_record.ps1')
. (Join-Path $PSScriptRoot 'refund_adjustments.ps1')
$skuSchema = Get-SkuMasterSchema -SpreadsheetToken $token
$stateRoot = Join-Path $ProjectRoot 'agent\state'
$script:financeTransaction = Join-FinanceTransaction $stateRoot
$script:financeTransactionOwner = $false
$script:financeTransactionClosed = $false
if ($null -eq $script:financeTransaction) {
    $script:financeTransaction = New-FinanceTransaction $stateRoot '重建成本快照' $null -AcquireLock -WorkbookTokens @{source=$token} -LarkCli $larkCli
    $script:financeTransactionOwner = $true
}

function Close-CostTransaction([bool]$Succeeded, [string]$Reason = '') {
    if ($script:financeTransactionClosed -or $null -eq $script:financeTransaction) { return }
    if ($script:financeTransactionOwner) {
        if ($Succeeded) { Complete-FinanceTransaction $script:financeTransaction }
        else {
            $rollback = Invoke-FinanceTransactionRollback $script:financeTransaction @{source=$token} $larkCli $Reason
            if (-not $rollback.ok) { [Console]::Error.WriteLine("成本快照回滚待恢复：事务 $($rollback.transaction_id)；$($rollback.error)") }
        }
    }
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
}

trap {
    $reason = [string]$_.Exception.Message
    Close-CostTransaction $false $reason
    [Console]::Error.WriteLine($reason)
    exit 1
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    $valueProperty = $cell.PSObject.Properties['value']
    if ($null -eq $valueProperty -or $null -eq $valueProperty.Value) { return $null }
    return [string]$valueProperty.Value
}
function Convert-ToNullableNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}
function Convert-ToDate($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed }
    $number = Convert-ToNullableNumber $Value
    if ($null -ne $number) { return ([datetime]'1899-12-30').AddDays($number) }
    return $null
}
function Normalize-ProductTitle($Value) {
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value).Trim().ToLowerInvariant()
    # 美团订单标题末尾会附加规格，例如“(50g/包)”。成本主表只保留商品标题。
    $text = [regex]::Replace($text, '\s*[\(（].*$', '')
    $text = [regex]::Replace($text, '\s+', '')
    return $text
}
function Test-ZeroCostTrialItem([string]$Title, [string]$Spec) {
    $text = (([string]$Title) + ' ' + ([string]$Spec)).Trim()
    return [bool]($text -match '试吃|尝鲜|试用|体验装')
}
function Convert-ToExcelSerial([datetime]$Value) { return ($Value - [datetime]'1899-12-30').TotalDays }
function New-OverwriteCell($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return @{value=''} }
    return @{value=$Value}
}
function New-UnitCostCell($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return @{value='';cell_styles=@{number_format='#,##0.####'}} }
    return @{value=$Value;cell_styles=@{number_format='#,##0.####'}}
}
function New-FormulaCell([string]$Formula) { return @{formula=$Formula} }
function Get-SheetRows([string]$SheetName, [string]$Range) {
    $raw = & $larkCli sheets +cells-get --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --include 'value,formula' --max-chars 3000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取飞书失败：$SheetName" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "飞书读取结果不完整：$SheetName" }
    return @($result.data.ranges[0].cells)
}
function Set-CellMatrix([string]$SheetName, [string]$Range, $Rows) {
    if (@($Rows).Count -eq 0) { return }
    Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $token $SheetName $Range $larkCli
    $payload = ConvertTo-Json -InputObject $Rows -Depth 12 -Compress
    $raw = $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not (($raw | ConvertFrom-Json).ok)) { throw "写入飞书失败：$SheetName $Range" }
}
function Clear-Range([string]$Range) {
    if ($Range -notmatch '^([^!]+)!(.+)$') { throw "无法识别清理范围：$Range" }
    Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $token $Matches[1] $Matches[2] $larkCli
    $ranges = ConvertTo-Json -InputObject @($Range) -Compress
    $raw = $ranges | & $larkCli sheets +cells-batch-clear --as user --spreadsheet-token $token --ranges - --scope content --yes --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not (($raw | ConvertFrom-Json).ok)) { throw "清理范围失败：$Range" }
}
function Assert-FormulaRange([string]$SheetName, [string]$Range) {
    $raw = & $larkCli sheets +formula-verify --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --format json --exit-on-error | Out-String
    if ($LASTEXITCODE -ne 0) { throw "公式校验失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if ($result.data.status -ne 'success') { throw "公式存在错误：$SheetName $Range" }
}

$masterRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
$states = @{}
$aliasCandidates = @{}
$masterRowByCanonical = @{}
for ($i=0; $i -lt $masterRows.Count; $i++) {
    $row = $masterRows[$i]
    $kw = ([string](Get-CellValue $row $skuSchema.CurrentKw.Index)).Trim().ToLowerInvariant()
    $productId = ([string](Get-CellValue $row $skuSchema.ProductId.Index)).Trim()
    $canonical = if ($productId) { "sku:$productId" } else { "legacy:$($i+2)" }
    $titleKey = Normalize-ProductTitle (Get-CellValue $row $skuSchema.PlatformTitle.Index)
    $csxAliases = @([regex]::Split([string](Get-CellValue $row $skuSchema.TransitionCode.Index), '[/,，、;；\s]+') | Where-Object { $_ -match '^(?i)csx\d+$' } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique)
    if (-not $kw -and -not $productId -and -not $titleKey -and $csxAliases.Count -eq 0) { continue }
    $costStatus = [string](Get-CellValue $row $skuSchema.CostStatus.Index)
    $rawBaselineCost = Convert-ToNullableNumber (Get-CellValue $row $skuSchema.BaselineUnitCost.Index)
    $confirmedBaselineCost = if ($null -ne $rawBaselineCost) { [double]$rawBaselineCost } else { $null }
    $baselineDate = Convert-ToDate (Get-CellValue $row $skuSchema.InitializationDate.Index)
    if ($null -eq $baselineDate) { $baselineDate = [datetime]'2026-08-24' }
    $state = [pscustomobject]@{
        canonical=$canonical
        product_id=$productId
        current_kw=$kw
        master_row=2+$i
        opening_qty=$(if ($null -eq (Convert-ToNullableNumber (Get-CellValue $row $skuSchema.InitialInventory.Index))) {0.0}else{[double](Convert-ToNullableNumber (Get-CellValue $row $skuSchema.InitialInventory.Index))})
        baseline_cost=$confirmedBaselineCost
        qty=$(if ($null -eq (Convert-ToNullableNumber (Get-CellValue $row $skuSchema.InitialInventory.Index))) {0.0}else{[double](Convert-ToNullableNumber (Get-CellValue $row $skuSchema.InitialInventory.Index))})
        avg=$confirmedBaselineCost
        units_per_pack=$(if ($null -eq (Convert-ToNullableNumber (Get-CellValue $row $skuSchema.UnitsPerPack.Index))) {1.0}else{[double](Convert-ToNullableNumber (Get-CellValue $row $skuSchema.UnitsPerPack.Index))})
        baseline_date=$baselineDate.Date
        status=$costStatus
        title_key=$titleKey
    }
    $states[$canonical] = $state
    $masterRowByCanonical[$canonical] = 2+$i
    foreach ($alias in @($kw) + $csxAliases) {
        if (-not $alias) { continue }
        if (-not $aliasCandidates.ContainsKey($alias)) { $aliasCandidates[$alias] = @() }
        $aliasCandidates[$alias] += $canonical
    }
}
$aliasToCanonical = @{}
foreach ($alias in $aliasCandidates.Keys) {
    $unique = @($aliasCandidates[$alias] | Select-Object -Unique)
    if ($unique.Count -eq 1) { $aliasToCanonical[$alias] = $unique[0] }
}

$interventionRows = @(Get-SheetRows '干预订单' 'A2:M1001' | Where-Object { Test-ActiveInterventionOrderNo (Get-CellValue $_ 3) })
$interventionSet = @{}
foreach ($row in $interventionRows) {
    $orderNo = Get-CellValue $row 3
    if ($orderNo) { $interventionSet[$orderNo] = $true }
}

$giftRows = Get-SheetRows '赠品记录' 'A2:L2001'
$giftEvents = @()
$giftCostByOrder = @{}
$pendingGiftCostByOrder = @{}
for ($i=0; $i -lt $giftRows.Count; $i++) {
    $row = $giftRows[$i]
    $dailyOrderNo = Get-CellValue $row 2
    $orderNo = Get-CellValue $row 3
    $rawCode = ([string](Get-CellValue $row 4)).Trim().ToLowerInvariant()
    $quantity = Convert-ToNullableNumber (Get-CellValue $row 5)
    $giftUnitCost = Convert-ToNullableNumber (Get-CellValue $row 6)
    $giftCost = Convert-ToNullableNumber (Get-CellValue $row 7)
    $costStatus = Get-CellValue $row 8
    $giftNote = Get-CellValue $row 9
    $giftDate = Convert-ToDate (Get-CellValue $row 0)
    if ($dailyOrderNo -in @('待同步','订单已取消待处理')) { continue }
    if (-not $orderNo -or -not $rawCode -or $null -eq $quantity -or $quantity -le 0 -or $null -eq $giftDate) { continue }
    $canonical = if ($aliasToCanonical.ContainsKey($rawCode)) { $aliasToCanonical[$rawCode] } else { $null }
    $giftEvents += [pscustomobject]@{
        row=2+$i;order_no=$orderNo;raw_code=$rawCode;canonical=$canonical;quantity=[double]$quantity;date=$giftDate.Date
        unit_cost=$giftUnitCost;gift_cost=$giftCost;cost_status=$costStatus;note=$giftNote;resolved_from_master=$false
    }
    if ($null -eq $giftCost -or $costStatus -eq '成本待核对') {
        $pendingGiftCostByOrder[$orderNo] = $true
    } else {
        if (-not $giftCostByOrder.ContainsKey($orderNo)) { $giftCostByOrder[$orderNo] = 0.0 }
        $giftCostByOrder[$orderNo] += [double]$giftCost
    }
}

$itemRows = Get-SheetRows '订单商品' 'A2:Q4001'
$itemOutputs = @()
$saleEvents = @()
$usedItemCount = 0
for ($i=0; $i -lt $itemRows.Count; $i++) {
    $row = $itemRows[$i]
    $orderNo = Get-CellValue $row 2
    $rawCode = ([string](Get-CellValue $row 5)).ToLowerInvariant()
    $titleKey = Normalize-ProductTitle (Get-CellValue $row 6)
    $productId = ([string](Get-CellValue $row 16)).Trim()
    if (-not $orderNo) { continue }
    $usedItemCount++
    $identityKey = if ($productId) { "sku:$productId" } else { $null }
    $canonical = if ($identityKey -and $states.ContainsKey($identityKey)) { $identityKey } elseif ($aliasToCanonical.ContainsKey($rawCode)) { $aliasToCanonical[$rawCode] } else { $null }
    $isIntervention = $interventionSet.ContainsKey($orderNo)
    $isCancelled = (Get-CellValue $row 3) -eq '已取消'
    $quantity = Convert-ToNullableNumber (Get-CellValue $row 8)
    if ($null -eq $quantity) { $quantity = 0.0 }
    $salesAmount = Convert-ToNullableNumber (Get-CellValue $row 10)
    if ($null -eq $salesAmount) {
        $unitPrice = Convert-ToNullableNumber (Get-CellValue $row 9)
        $salesAmount = $(if ($null -eq $unitPrice) {0.0}else{$unitPrice}) * $quantity
    }
    $itemOutputs += [pscustomobject]@{
        row=2+$i; order_no=$orderNo; canonical=$canonical; product_id=$productId; raw_code=$rawCode; title_key=$titleKey; quantity=[double]$quantity; sales_amount=[double]$salesAmount
        returned_quantity=0.0; effective_quantity=[double]$quantity
        date=Convert-ToDate (Get-CellValue $row 0); intervention=$isIntervention; cancelled=$isCancelled
        trial_zero_cost=(Test-ZeroCostTrialItem (Get-CellValue $row 6) (Get-CellValue $row 7))
        previous_unit_cost=Convert-ToNullableNumber (Get-CellValue $row 11); previous_item_cost=Convert-ToNullableNumber (Get-CellValue $row 12); previous_margin=Convert-ToNullableNumber (Get-CellValue $row 13)
        unit_cost=$null; item_cost=$null; margin=$null
    }
}

# 已退回并重新入库的退款商品不再属于已售货物成本。退款事件以售后记录号幂等保存；
# 这里只按“净售出数量”重放移动平均成本，不改写订单原始数量。
$returnedQuantityByOrderCode = @{}
foreach ($refundEvent in @(Get-RefundEvents $ProjectRoot)) {
    foreach ($refundItem in @($refundEvent.items | Where-Object { $_.disposition -eq 'returned_restocked' })) {
        $code = ([string]$refundItem.goods_code).Trim().ToLowerInvariant()
        if (-not $code) { throw "退款事件 $($refundEvent.event_id) 缺少商品货号，无法安全撤销成本" }
        $key = "$($refundEvent.order_no)|$code"
        if (-not $returnedQuantityByOrderCode.ContainsKey($key)) { $returnedQuantityByOrderCode[$key] = 0.0 }
        $returnedQuantityByOrderCode[$key] += [double]$refundItem.quantity
    }
}
foreach ($item in @($itemOutputs | Sort-Object row)) {
    $key = "$($item.order_no)|$($item.raw_code)"
    if (-not $returnedQuantityByOrderCode.ContainsKey($key) -or [double]$returnedQuantityByOrderCode[$key] -le 0) { continue }
    $returned = [math]::Min([double]$item.quantity, [double]$returnedQuantityByOrderCode[$key])
    $item.returned_quantity = $returned
    $item.effective_quantity = [math]::Max([double]0.0, [double]$item.quantity - $returned)
    $returnedQuantityByOrderCode[$key] = [double]$returnedQuantityByOrderCode[$key] - $returned
}
$unmatchedReturns = @($returnedQuantityByOrderCode.GetEnumerator() | Where-Object { [double]$_.Value -gt 0.000001 })
if ($unmatchedReturns.Count -gt 0) {
    $sample = @($unmatchedReturns | Select-Object -First 3 | ForEach-Object { "$($_.Key) 尚有 $($_.Value) 件" }) -join '、'
    throw "退款退货数量超过订单商品数量，已停止成本重算：$sample"
}

$replenishmentRows = Get-SheetRows '补货记录' 'A2:T1001'
$replenishments = @()
$unresolvedReplenishments = [Collections.Generic.List[string]]::new()
for ($i=0; $i -lt $replenishmentRows.Count; $i++) {
    $row = $replenishmentRows[$i]
    $rawCode = ([string](Get-CellValue $row 1)).ToLowerInvariant()
    if (-not $rawCode) { continue }
    $productId = ([string](Get-CellValue $row 14)).Trim()
    $identityKey = if ($productId) { "sku:$productId" } else { $null }
    $canonical = if ($identityKey -and $states.ContainsKey($identityKey)) { $identityKey } elseif ($aliasToCanonical.ContainsKey($rawCode)) { $aliasToCanonical[$rawCode] } else { $null }
    if (-not $canonical) { $unresolvedReplenishments.Add("第$(2+$i)行 $rawCode") ; continue }
    $packageQty = Convert-ToNullableNumber (Get-CellValue $row 2)
    $amount = Convert-ToNullableNumber (Get-CellValue $row 5)
    $date = Convert-ToDate (Get-CellValue $row 0)
    if ($null -eq $date -or $null -eq $packageQty -or $packageQty -le 0 -or $null -eq $amount -or $amount -lt 0) { throw "补货记录第$(2+$i)行数据不完整" }
    $unitsPerPack = [double]$states[$canonical].units_per_pack
    $units = [double]$packageQty * $unitsPerPack
    $unitCost = [double]$amount / $units
    $note = [string](Get-CellValue $row 13)
    $batchId = [string](Get-CellValue $row 19)
    # A new SKU is often created after its first stock has already been entered on
    # Meituan.  In that case the opening inventory and the first replenishment row
    # describe the same physical stock.  Counting both doubles quantity/value.
    $isOpeningDuplicate = (
        $date.Date -eq $states[$canonical].baseline_date -and
        [math]::Abs($units - [double]$states[$canonical].opening_qty) -lt 0.000001 -and
        $null -ne $states[$canonical].baseline_cost -and
        [math]::Abs($unitCost - [double]$states[$canonical].baseline_cost) -lt 0.000001
    )
    # 历史漏登的补货若早于 SKU 初始化快照，数量已包含在“初始化库存”里，
    # 不能再次增加库存；但它必须修正该批次之后的历史销售成本，以及初始化
    # 快照开始时的成本基准。只识别带专用批次标记和备注的人工补录，避免把
    # 普通旧记录误判成成本基准调整。
    $isHistoricalBaselineAdjustment = (
        $date.Date -lt $states[$canonical].baseline_date -and
        $batchId -match '^MANUAL-' -and
        $note -match '历史漏登补货'
    )
    $replenishments += [pscustomobject]@{
        row=2+$i; canonical=$canonical; product_id=$productId; date=$date.Date; package_qty=[double]$packageQty; units_per_pack=$unitsPerPack
        units=$units; amount=[double]$amount; unit_cost=$unitCost; opening_duplicate=$isOpeningDuplicate
        historical_baseline_adjustment=$isHistoricalBaselineAdjustment; batch_id=$batchId; note=$note
        before_qty=$(if($isOpeningDuplicate){[double]$states[$canonical].opening_qty}elseif($isHistoricalBaselineAdjustment){[double]$states[$canonical].opening_qty}else{$null})
        before_avg=$(if($isOpeningDuplicate){$states[$canonical].baseline_cost}elseif($isHistoricalBaselineAdjustment){$states[$canonical].baseline_cost}else{$null})
        after_qty=$(if($isOpeningDuplicate){[double]$states[$canonical].opening_qty}elseif($isHistoricalBaselineAdjustment){[double]$states[$canonical].opening_qty}else{$null})
        after_avg=$(if($isOpeningDuplicate){$states[$canonical].baseline_cost}elseif($isHistoricalBaselineAdjustment){$unitCost}else{$null})
    }
}

$historicalAdjustmentsByCanonical = @{}
foreach ($purchase in @($replenishments | Where-Object { $_.historical_baseline_adjustment } | Sort-Object date,row)) {
    if (-not $historicalAdjustmentsByCanonical.ContainsKey($purchase.canonical)) { $historicalAdjustmentsByCanonical[$purchase.canonical] = @() }
    $history = @($historicalAdjustmentsByCanonical[$purchase.canonical])
    $priorCost = if ($history.Count -gt 0) { [double]$history[-1].unit_cost } else { $states[$purchase.canonical].baseline_cost }
    $purchase.before_qty = [double]$states[$purchase.canonical].opening_qty
    $purchase.before_avg = $priorCost
    $purchase.after_qty = [double]$states[$purchase.canonical].opening_qty
    $purchase.after_avg = [double]$purchase.unit_cost
    $historicalAdjustmentsByCanonical[$purchase.canonical] = @($history + $purchase)
    # 初始化库存是快照数量，不因历史补录重复增加；当前成本从最近一笔
    # 初始化日前补货开始重放。
    $states[$purchase.canonical].avg = [double]$purchase.unit_cost
}

$events = @()
foreach ($purchase in $replenishments) {
    $state = $states[$purchase.canonical]
    if (-not $purchase.opening_duplicate -and -not $purchase.historical_baseline_adjustment -and $purchase.date -ge $state.baseline_date) {
        $events += [pscustomobject]@{date=$purchase.date;sort=1;row=$purchase.row;kind='purchase';canonical=$purchase.canonical;payload=$purchase}
    }
}
foreach ($item in $itemOutputs) {
    if (-not $item.canonical) { continue }
    $state = $states[$item.canonical]
    if ($null -ne $item.date -and $item.date.Date -ge $state.baseline_date -and -not $item.intervention -and -not $item.cancelled) {
        $events += [pscustomobject]@{date=$item.date.Date;sort=2;row=$item.row;kind='sale';canonical=$item.canonical;payload=$item}
    }
}
foreach ($gift in $giftEvents) {
    if (-not $gift.canonical) { continue }
    $state = $states[$gift.canonical]
    if ($null -ne $state -and $gift.date -ge $state.baseline_date) {
        $events += [pscustomobject]@{date=$gift.date;sort=3;row=$gift.row;kind='gift';canonical=$gift.canonical;payload=$gift}
    }
}
$events = @($events | Sort-Object date,sort,row)
foreach ($event in $events) {
    $state = $states[$event.canonical]
    if ($event.kind -eq 'purchase') {
        $purchase = $event.payload
        $purchase.before_qty = [double]$state.qty
        $purchase.before_avg = $state.avg
        $afterQty = [double]$state.qty + [double]$purchase.units
        if ($null -eq $state.avg -and $state.qty -gt 0) {
            $afterAvg = $null
        } elseif ($afterQty -le 0) {
            $afterAvg = $purchase.unit_cost
        } else {
            $oldValue = $(if ($null -eq $state.avg) {0.0}else{[double]$state.avg * [double]$state.qty})
            $afterAvg = ($oldValue + [double]$purchase.amount) / $afterQty
        }
        $state.qty = $afterQty
        $state.avg = $afterAvg
        $purchase.after_qty = $afterQty
        $purchase.after_avg = $afterAvg
    } elseif ($event.kind -eq 'sale') {
        $item = $event.payload
        $item.unit_cost = $state.avg
        if ($null -ne $state.avg) {
            $item.item_cost = [double]$state.avg * [double]$item.effective_quantity
            $effectiveSalesAmount = if ([double]$item.quantity -le 0) { 0.0 } else { [double]$item.sales_amount * ([double]$item.effective_quantity / [double]$item.quantity) }
            $item.margin = $effectiveSalesAmount - [double]$item.item_cost
        }
        $state.qty = [math]::Max(0.0, [double]$state.qty - [double]$item.effective_quantity)
    } else {
        $gift = $event.payload
        if (($null -eq $gift.gift_cost -or $gift.cost_status -eq '成本待核对') -and $null -ne $state.avg) {
            # 赠品登记时成本未知，后来在 SKU 成本主表确认后，应按赠送当时的移动平均成本补齐。
            # 数值 0 是有效成本，不能再按“没有成本”处理。
            $gift.unit_cost = [double]$state.avg
            $gift.gift_cost = [math]::Round(([double]$state.avg * [double]$gift.quantity),2)
            $gift.cost_status = '已计成本'
            $gift.resolved_from_master = $true
        }
        $state.qty = [math]::Max(0.0, [double]$state.qty - [double]$gift.quantity)
    }
}

# 补齐未进入移动平均事件流的历史赠品，只允许使用 SKU 成本主表中已确认的基准成本。
# 旧“商品成本”表只作为人工核对参考，不能绕过待确认状态参与利润计算。
foreach ($gift in $giftEvents) {
    if ($null -ne $gift.gift_cost -and $gift.cost_status -ne '成本待核对') { continue }
    $fallbackCost = $null
    if ($gift.canonical -and $states.ContainsKey($gift.canonical)) {
        $giftState = $states[$gift.canonical]
        if ($null -ne $giftState.baseline_cost) { $fallbackCost = [double]$giftState.baseline_cost }
    }
    if ($null -eq $fallbackCost) { continue }
    $gift.unit_cost = $fallbackCost
    $gift.gift_cost = [math]::Round(($fallbackCost * [double]$gift.quantity),2)
    $gift.cost_status = '已计成本'
    $gift.resolved_from_master = $true
}

$normalizedGiftCostRows = @($giftEvents | Where-Object { $null -ne $_.gift_cost -and $_.cost_status -ne '成本待核对' })
$resolvedGiftCostRows = @($giftEvents | Where-Object { $_.resolved_from_master })
foreach ($gift in $normalizedGiftCostRows) {
    $cleanNote = [string]$gift.note
    if($gift.resolved_from_master){
        $cleanNoteParts = @([regex]::Split([string]$gift.note, '[；;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notmatch '^赠品成本待核对(?:[:：]|$)' })
        $cleanNote = if ($cleanNoteParts.Count -gt 0) { $cleanNoteParts -join '；' } else { '成本已按SKU成本主表重新计算' }
    }
    $giftUpdate = @(); $giftUpdate += ,@(
        (New-UnitCostCell $(if($null-eq$gift.unit_cost){$null}else{[double]$gift.unit_cost})),(New-OverwriteCell ([double]$gift.gift_cost)),
        (New-OverwriteCell '已计成本'),(New-OverwriteCell $cleanNote)
    )
    Set-CellMatrix '赠品记录' "G$($gift.row):J$($gift.row)" $giftUpdate
}

# 重新统计，确保刚补齐的 0 元赠品不再被计入待核对。
$giftCostByOrder = @{}
$pendingGiftCostByOrder = @{}
foreach ($gift in $giftEvents) {
    if ($null -eq $gift.gift_cost -or $gift.cost_status -eq '成本待核对') {
        $pendingGiftCostByOrder[$gift.order_no] = $true
    } else {
        if (-not $giftCostByOrder.ContainsKey($gift.order_no)) { $giftCostByOrder[$gift.order_no] = 0.0 }
        $giftCostByOrder[$gift.order_no] += [double]$gift.gift_cost
    }
}

foreach ($item in $itemOutputs) {
    if ($item.intervention) {
        $item.unit_cost = 0.0; $item.item_cost = 0.0; $item.margin = [double]$item.sales_amount
        continue
    }
    if ($item.cancelled) {
        $item.unit_cost = 0.0; $item.item_cost = 0.0; $item.margin = 0.0
        continue
    }
    if ($item.trial_zero_cost) {
        # 店主确认的统一规则：试吃、尝鲜、试用和体验装不计货物成本。
        # 该规则按订单商品标题/规格识别，覆盖没有稳定商品ID的历史订单。
        $item.unit_cost = 0.0; $item.item_cost = 0.0; $item.margin = [double]$item.sales_amount
        continue
    }
    if ($null -ne $item.unit_cost) { continue }
    if ($item.canonical) {
        $state = $states[$item.canonical]
        if ($null -ne $item.date -and $item.date.Date -lt $state.baseline_date) {
            $effectiveHistoricalCost = $state.baseline_cost
            if ($historicalAdjustmentsByCanonical.ContainsKey($item.canonical)) {
                $eligibleAdjustments = @($historicalAdjustmentsByCanonical[$item.canonical] | Where-Object { $_.date -le $item.date.Date } | Sort-Object date,row)
                if ($eligibleAdjustments.Count -gt 0) { $effectiveHistoricalCost = [double]$eligibleAdjustments[-1].unit_cost }
            }
            $item.unit_cost = $effectiveHistoricalCost
            if ($null -ne $effectiveHistoricalCost) {
                $item.item_cost = [double]$effectiveHistoricalCost * [double]$item.effective_quantity
                $effectiveSalesAmount = if ([double]$item.quantity -le 0) { 0.0 } else { [double]$item.sales_amount * ([double]$item.effective_quantity / [double]$item.quantity) }
                $item.margin = $effectiveSalesAmount - [double]$item.item_cost
            }
        }
    }
    if ($null -eq $item.unit_cost) {
        if (-not $item.canonical -and $null -ne $item.previous_unit_cost -and $null -ne $item.previous_item_cost) {
            # 无法确定稳定商品身份时，冻结既有订单快照，绝不因当前 KW 被复用而改写历史成本。
            $item.unit_cost = [double]$item.previous_unit_cost
            $item.item_cost = [double]$item.previous_item_cost
            $item.margin = if ($null -ne $item.previous_margin) { [double]$item.previous_margin } else { [double]$item.sales_amount - [double]$item.previous_item_cost }
        }
    }
}

Clear-Range '订单商品!L2:N4001'
$itemCostMatrix = @()
$itemByRow = @{}
foreach ($item in $itemOutputs) { $itemByRow[[int]$item.row] = $item }
$lastItemRow = if ($itemOutputs.Count -gt 0) { [int](($itemOutputs | Measure-Object -Property row -Maximum).Maximum) } else { 1 }
for ($sheetRow=2; $sheetRow -le $lastItemRow; $sheetRow++) {
    if ($itemByRow.ContainsKey($sheetRow)) {
        $item = $itemByRow[$sheetRow]
        $itemCostMatrix += ,@((New-UnitCostCell $item.unit_cost),(New-OverwriteCell $item.item_cost),(New-OverwriteCell $item.margin))
    } else {
        $itemCostMatrix += ,@((New-OverwriteCell $null),(New-OverwriteCell $null),(New-OverwriteCell $null))
    }
}
if ($itemCostMatrix.Count -gt 0) { Set-CellMatrix '订单商品' "L2:N$lastItemRow" $itemCostMatrix }

$masterUpdate = @()
$lastMasterIndex = -1
for ($i=0; $i -lt $masterRows.Count; $i++) {
    $row = $masterRows[$i]
    if ((Get-CellValue $row 0) -or (Get-CellValue $row 1)) { $lastMasterIndex = $i }
}
$lastMasterRow = 1 + $lastMasterIndex + 1
for ($i=0; $i -le $lastMasterIndex; $i++) {
    $row = $masterRows[$i]
    $productId = ([string](Get-CellValue $row $skuSchema.ProductId.Index)).Trim()
    $canonical = if ($productId) { "sku:$productId" } else { "legacy:$($i+2)" }
    if ($states.ContainsKey($canonical)) {
        $state = $states[$canonical]
        $masterUpdate += ,@((New-UnitCostCell $state.avg),(New-OverwriteCell $state.qty),(New-OverwriteCell (Convert-ToExcelSerial ([datetime]::Now)))
        )
    } else {
        $masterUpdate += ,@(
            (New-OverwriteCell (Get-CellValue $row $skuSchema.MovingAverageCost.Index)),
            (New-OverwriteCell (Get-CellValue $row $skuSchema.CostBookQuantity.Index)),
            (New-OverwriteCell (Get-CellValue $row $skuSchema.LastCostUpdate.Index))
        )
    }
}
if ($masterUpdate.Count -gt 0) {
    $fg = @(); $l = @()
    foreach ($row in $masterUpdate) { $fg += ,@($row[0],$row[1]); $l += ,@($row[2]) }
    Set-CellMatrix 'SKU成本主表' "$($skuSchema.MovingAverageCost.Column)2:$($skuSchema.CostBookQuantity.Column)$lastMasterRow" $fg
    Set-CellMatrix 'SKU成本主表' "$($skuSchema.LastCostUpdate.Column)2:$($skuSchema.LastCostUpdate.Column)$lastMasterRow" $l
}

if ($replenishments.Count -gt 0) {
    foreach ($purchase in $replenishments) {
        $replenishmentSnapshot = @(); $replenishmentSnapshot += ,@(
            (New-OverwriteCell $purchase.before_qty),(New-OverwriteCell $purchase.before_avg),
            (New-OverwriteCell $purchase.after_qty),(New-OverwriteCell $purchase.after_avg)
        )
        Set-CellMatrix '补货记录' "H$($purchase.row):K$($purchase.row)" $replenishmentSnapshot
    }
}

$summaryRows = Get-SheetRows '订单汇总' 'A2:P2001'
$summaryCount = @($summaryRows | Where-Object { Get-CellValue $_ 2 }).Count
$summaryK=@();$summaryL=@()
for ($row=2; $row -le 2001; $row++) {
    $summaryK += ,@((New-FormulaCell "=IF(C$row=`"`",`"`",IF(COUNTIFS('订单商品'!`$C`$2:`$C`$4001,C$row,'订单商品'!`$E`$2:`$E`$4001,`"自然`",'订单商品'!`$L`$2:`$L`$4001,`"`")+IF(E$row=`"自然`",COUNTIFS('赠品记录'!`$D`$2:`$D`$2001,C$row,'赠品记录'!`$I`$2:`$I`$2001,`"成本待核对`",'赠品记录'!`$C`$2:`$C`$2001,`"<>待同步`",'赠品记录'!`$C`$2:`$C`$2001,`"<>订单已取消待处理`"),0)>0,`"`",SUMIFS('订单商品'!`$M`$2:`$M`$4001,'订单商品'!`$C`$2:`$C`$4001,C$row)+IF(E$row=`"自然`",SUMIFS('赠品记录'!`$H`$2:`$H`$2001,'赠品记录'!`$D`$2:`$D`$2001,C$row,'赠品记录'!`$C`$2:`$C`$2001,`"<>待同步`",'赠品记录'!`$C`$2:`$C`$2001,`"<>订单已取消待处理`"),0)))"))
    $summaryL += ,@((New-FormulaCell "=IF(C$row=`"`",`"`",IF(K$row=`"`",`"`",IFERROR(H$row*1,0)-IFERROR(K$row*1,0)-IFERROR(J$row*1,0)))"))
}
Set-CellMatrix '订单汇总' 'K2:K2001' $summaryK
Set-CellMatrix '订单汇总' 'L2:L2001' $summaryL

Assert-FormulaRange '订单汇总' 'E2:E2001,K2:L2001'
Assert-FormulaRange '订单商品' 'E2:E4001,K2:K4001,P2:P4001'
Assert-FormulaRange '补货记录' 'D2:E1001,G2:G1001'

$costArithmeticErrors = @($itemOutputs | Where-Object {
    -not $_.intervention -and -not $_.cancelled -and $null -ne $_.unit_cost -and $null -ne $_.item_cost -and
    [math]::Abs([double]$_.item_cost - ([double]$_.unit_cost * [double]$_.effective_quantity)) -gt 0.02
})
if ($costArithmeticErrors.Count -gt 0) {
    $sample = @($costArithmeticErrors | Select-Object -First 3 | ForEach-Object { "$($_.order_no)/$($_.raw_code)" }) -join '、'
    throw "订单商品成本算术校验失败：$sample"
}

$finalResult = [ordered]@{
    ok=$true
    item_rows=$itemOutputs.Count
    natural_costed=@($itemOutputs | Where-Object { -not $_.intervention -and -not $_.cancelled -and $null -ne $_.item_cost }).Count
    natural_pending_cost=@($itemOutputs | Where-Object { -not $_.intervention -and -not $_.cancelled -and $null -eq $_.item_cost }).Count
    reference_fallback_costed=0
    trial_zero_cost=@($itemOutputs | Where-Object { -not $_.intervention -and -not $_.cancelled -and $_.trial_zero_cost -and $_.item_cost -eq 0 }).Count
    intervention_zero_cost=@($itemOutputs | Where-Object { $_.intervention -and $_.item_cost -eq 0 }).Count
    replenishment_rows=$replenishments.Count
    unresolved_replenishments=@($unresolvedReplenishments)
    historical_baseline_adjustments=@($replenishments | Where-Object { $_.historical_baseline_adjustment }).Count
    cost_arithmetic_errors=$costArithmeticErrors.Count
    returned_refund_items=@($itemOutputs | Where-Object { $_.returned_quantity -gt 0 }).Count
    returned_refund_quantity=[math]::Round((@($itemOutputs | Measure-Object -Property returned_quantity -Sum).Sum),2)
    summary_orders=$summaryCount
    gift_rows=$giftEvents.Count
    gift_inventory_deductions=@($giftEvents | Where-Object { $_.canonical }).Count
    gift_cost_rows_resolved=$resolvedGiftCostRows.Count
    gift_pending_cost_orders=$pendingGiftCostByOrder.Count
}
Close-CostTransaction $true
$finalResult | ConvertTo-Json -Compress
