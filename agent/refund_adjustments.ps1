Set-StrictMode -Version 2

function Get-RefundAdjustmentStatePath([string]$ProjectRoot) {
    return (Join-Path $ProjectRoot 'agent\state\refund_adjustments.json')
}

function New-RefundAdjustmentState {
    return [pscustomobject]@{
        version = 1
        updated_at = $null
        events = @()
    }
}

function Read-RefundAdjustmentState([string]$ProjectRoot) {
    $path = Get-RefundAdjustmentStatePath $ProjectRoot
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return (New-RefundAdjustmentState) }
    try {
        $state = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
    } catch {
        throw "退款补偿状态文件损坏，已停止算账：$path；$($_.Exception.Message)"
    }
    if ($null -eq $state -or [int]$state.version -ne 1) { throw "退款补偿状态版本无法识别：$path" }
    if ($null -eq $state.events) { $state | Add-Member -NotePropertyName events -NotePropertyValue @() -Force }
    return $state
}

function Write-RefundAdjustmentState([string]$ProjectRoot, $State) {
    $path = Get-RefundAdjustmentStatePath $ProjectRoot
    $directory = Split-Path $path -Parent
    if (-not (Test-Path -LiteralPath $directory)) { $null = New-Item -ItemType Directory -Path $directory }
    $State.updated_at = [datetime]::Now.ToString('o')
    $json = $State | ConvertTo-Json -Depth 30
    $temporary = "$path.$PID.tmp"
    try {
        [IO.File]::WriteAllText($temporary, $json, [Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $temporary -Destination $path -Force
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue }
    }
}

function Convert-ToRefundNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}

function Merge-RefundDiscoveries([string]$ProjectRoot, $Orders) {
    $state = Read-RefundAdjustmentState $ProjectRoot
    $events = [Collections.Generic.List[object]]::new()
    foreach ($event in @($state.events)) { $events.Add($event) }
    $changed = $false
    $detected = [Collections.Generic.List[object]]::new()

    foreach ($order in @($Orders)) {
        foreach ($refund in @($order.refunds)) {
            if (-not $refund.completed -or -not $refund.recordId) { continue }
            $eventId = "meituan:$($order.orderNo):$($refund.recordId)"
            $existing = @($events | Where-Object { [string]$_.event_id -eq $eventId } | Select-Object -First 1)
            if ($existing.Count -eq 0) {
                $priorOrderEvent = @($events | Where-Object { [string]$_.order_no -eq [string]$order.orderNo } | Select-Object -First 1)
                $baseExpectedIncome = if ($priorOrderEvent.Count -and $null -ne $priorOrderEvent[0].base_expected_income) {
                    [double]$priorOrderEvent[0].base_expected_income
                } elseif ($null -ne $order.expectedIncome) {
                    [double]$order.expectedIncome
                } else {
                    $null
                }
                $items = @($refund.items | ForEach-Object {
                    [pscustomobject]@{
                        goods_code = ([string]$_.goodsCode).Trim().ToLowerInvariant()
                        product_id = [string]$_.productId
                        title = [string]$_.title
                        spec = [string]$_.spec
                        quantity = [double]$_.quantity
                        refund_amount = [double]$_.refundAmount
                        disposition = 'pending'
                        disposition_note = $null
                        disposition_updated_at = $null
                    }
                })
                $newEvent = [pscustomobject]@{
                    event_id = $eventId
                    platform = '美团'
                    order_no = [string]$order.orderNo
                    order_date = [string]$order.orderDate
                    record_id = [string]$refund.recordId
                    refund_type = [string]$refund.refundType
                    completed_at = [string]$refund.completedAt
                    requested_at = [string]$refund.requestedAt
                    reason = [string]$refund.reason
                    customer_refund_amount = [double]$refund.amount
                    income_adjustment = [double]$refund.amount
                    income_adjustment_basis = '美团已完成的顾客退款金额'
                    base_expected_income = $baseExpectedIncome
                    first_detected_at = [datetime]::Now.ToString('o')
                    last_seen_at = [datetime]::Now.ToString('o')
                    items = $items
                }
                $events.Add($newEvent)
                $detected.Add($newEvent)
                $changed = $true
                continue
            }

            $current = $existing[0]
            $current.last_seen_at = [datetime]::Now.ToString('o')
            $current.completed_at = [string]$refund.completedAt
            $current.requested_at = [string]$refund.requestedAt
            $current.reason = [string]$refund.reason
            $current.customer_refund_amount = [double]$refund.amount
            $current.income_adjustment = [double]$refund.amount
            $knownItems = @($current.items)
            $mergedItems = @()
            foreach ($item in @($refund.items)) {
                $code = ([string]$item.goodsCode).Trim().ToLowerInvariant()
                $known = @($knownItems | Where-Object { ([string]$_.goods_code).Trim().ToLowerInvariant() -eq $code } | Select-Object -First 1)
                $mergedItems += [pscustomobject]@{
                    goods_code = $code
                    product_id = [string]$item.productId
                    title = [string]$item.title
                    spec = [string]$item.spec
                    quantity = [double]$item.quantity
                    refund_amount = [double]$item.refundAmount
                    disposition = $(if ($known.Count) { [string]$known[0].disposition } else { 'pending' })
                    disposition_note = $(if ($known.Count) { [string]$known[0].disposition_note } else { $null })
                    disposition_updated_at = $(if ($known.Count) { [string]$known[0].disposition_updated_at } else { $null })
                }
            }
            $current.items = $mergedItems
            $changed = $true
        }
    }
    $state.events = @($events)
    if ($changed) { Write-RefundAdjustmentState $ProjectRoot $state }
    return [pscustomobject]@{ state=$state; newly_detected=@($detected); changed=$changed }
}

function Get-RefundEvents([string]$ProjectRoot) {
    return @((Read-RefundAdjustmentState $ProjectRoot).events)
}

function Get-RefundEventsForOrder([string]$ProjectRoot, [string]$OrderNo) {
    return @(Get-RefundEvents $ProjectRoot | Where-Object { [string]$_.order_no -eq $OrderNo })
}

function Set-RefundItemDisposition(
    [string]$ProjectRoot,
    [string]$OrderNo,
    [string]$GoodsCode,
    [ValidateSet('returned_restocked','refund_no_return','returned_damaged')][string]$Disposition,
    [string]$Note
) {
    $state = Read-RefundAdjustmentState $ProjectRoot
    $normalizedCode = ([string]$GoodsCode).Trim().ToLowerInvariant()
    $matched = [Collections.Generic.List[object]]::new()
    foreach ($event in @($state.events | Where-Object { [string]$_.order_no -eq $OrderNo })) {
        foreach ($item in @($event.items)) {
            if ($normalizedCode -and ([string]$item.goods_code).Trim().ToLowerInvariant() -ne $normalizedCode) { continue }
            $item.disposition = $Disposition
            $item.disposition_note = $Note
            $item.disposition_updated_at = [datetime]::Now.ToString('o')
            $matched.Add([pscustomobject]@{event_id=$event.event_id;order_no=$OrderNo;goods_code=$item.goods_code;quantity=$item.quantity;disposition=$Disposition;title=$item.title})
        }
    }
    if ($matched.Count -eq 0) {
        $suffix = if ($normalizedCode) { "，商品 $normalizedCode" } else { '' }
        throw "退款记录中找不到订单 $OrderNo$suffix；请先同步美团订单"
    }
    Write-RefundAdjustmentState $ProjectRoot $state
    return @($matched)
}

function Get-RefundAccountingSummary([string]$ProjectRoot, [string]$OrderNo) {
    $events = @(Get-RefundEventsForOrder $ProjectRoot $OrderNo)
    $incomeAdjustment = 0.0
    $returned = @{}
    $pending = [Collections.Generic.List[object]]::new()
    $damaged = [Collections.Generic.List[object]]::new()
    foreach ($event in $events) {
        $incomeAdjustment += [double]$event.income_adjustment
        foreach ($item in @($event.items)) {
            $code = ([string]$item.goods_code).Trim().ToLowerInvariant()
            if ($item.disposition -eq 'returned_restocked') {
                if (-not $returned.ContainsKey($code)) { $returned[$code] = 0.0 }
                $returned[$code] += [double]$item.quantity
            } elseif ($item.disposition -eq 'pending') {
                $pending.Add($item)
            } elseif ($item.disposition -eq 'returned_damaged') {
                $damaged.Add($item)
            }
        }
    }
    return [pscustomobject]@{
        events = $events
        income_adjustment = [math]::Round($incomeAdjustment,2)
        returned_quantity_by_code = $returned
        pending_items = @($pending)
        damaged_items = @($damaged)
    }
}
