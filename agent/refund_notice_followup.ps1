# Notification metadata only. Never reads or writes a business workbook.

function Get-RefundRejectionNoticeKey($Observation) {
    if ($null -eq $Observation -or [string]$Observation.state -ne 'unknown') { return $null }
    if ([string]$Observation.status_text -notmatch '^\s*商家(?:已)?(?:驳回|拒绝)(?:全部)?退款(?:请求|申请)\s*$') { return $null }
    if ([string]$Observation.order_no -notmatch '^\d{18,20}$' -or [string]$Observation.record_id -notmatch '^[A-Za-z0-9_.:-]{1,120}$') { return $null }
    if ($null -eq $Observation.refund_amount -or $null -eq $Observation.item_count) { return $null }
    $amount = 0.0; $items = 0
    if (-not [double]::TryParse([string]$Observation.refund_amount,[ref]$amount) -or
        -not [int]::TryParse([string]$Observation.item_count,[ref]$items) -or $amount -ne 0 -or $items -ne 0) { return $null }
    return "$($Observation.order_no)|$($Observation.record_id)"
}

function Read-RefundNoticeAcknowledgements([string]$StateRoot) {
    $path = Join-Path $StateRoot 'refund_notice_acknowledgements.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return @() }
    $stored = Get-Content -Raw -LiteralPath $path -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $stored -or $stored.version -ne 1 -or -not $stored.PSObject.Properties['records']) { throw '退款提醒回执文件格式无效' }
    return @($stored.records | Where-Object { $_ -and $_.key })
}

function Test-AcknowledgedRefundRejection($Observation, $Acknowledgements) {
    $key = Get-RefundRejectionNoticeKey $Observation
    return [bool]($key -and @($Acknowledgements | Where-Object { $_.key -ceq $key }).Count)
}

function Register-RefundRejectionNoRefund([string]$StateRoot, $Observation) {
    $key = Get-RefundRejectionNoticeKey $Observation
    if (-not $key) { throw '只允许确认已观察到的零退款驳回记录；未记回执' }
    $path = Join-Path $StateRoot 'refund_notice_acknowledgements.json'
    $mutex = [Threading.Mutex]::new($false, 'Local\PetstarRefundNoticeAcknowledgements')
    $locked = $false; $temp = $null
    try {
        try { $locked = $mutex.WaitOne([TimeSpan]::FromSeconds(10)) }
        catch [Threading.AbandonedMutexException] { $locked = $true }
        if (-not $locked) { throw '退款提醒回执正忙，请稍后再试' }
        $records = @(Read-RefundNoticeAcknowledgements $StateRoot)
        $prior = $records | Where-Object { $_.key -ceq $key } | Select-Object -First 1
        if ($prior) { return $prior }
        $entry = [pscustomobject]@{
            key=$key; order_no=[string]$Observation.order_no; record_id=[string]$Observation.record_id
            resolution='owner_confirmed_rejected_no_refund'; acknowledged_at=[datetimeoffset]::Now.ToString('o')
        }
        $temp = Join-Path $StateRoot ("refund_notice_ack_{0}.tmp" -f [guid]::NewGuid().ToString('N'))
        $payload = [ordered]@{version=1;records=($records + @($entry))} | ConvertTo-Json -Depth 8
        [IO.File]::WriteAllText($temp,$payload,[Text.UTF8Encoding]::new($false))
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            [IO.File]::Replace($temp,$path,[NullString]::Value)
        } else {
            [IO.File]::Move($temp,$path)
        }
        return $entry
    } finally {
        if ($temp -and [IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
        if ($locked) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}

function Get-RefundNoticeAcknowledgementIntent([string]$Text) {
    if (-not $Text) { return $null }
    $answer = $Text.Trim()
    $numbers = @([regex]::Matches($answer,'(?<!\d)\d{18,20}(?!\d)') | ForEach-Object Value | Select-Object -Unique)
    if ($numbers.Count -gt 1) { return $null }
    if ($numbers.Count -eq 1) {
        $answer = $answer -replace '^\s*(?:美团)?(?:订单(?:号)?\s*[:：]?)?\s*\d{18,20}\s*(?:[,，:：]\s*)?',''
    }
    if ($answer -notmatch '^(?:(?:这(?:单|笔)|退款(?:请求)?)[，,\s]*)?(?:已处理(?:好|完毕)?|处理好了|已协商好|已拒绝退款)[，,。\s]*(?:没有退款|未退款|没退款)[。！!\s]*$') { return $null }
    return [pscustomobject]@{order_no=if ($numbers.Count -eq 1) { $numbers[0] } else { '' }}
}

function Find-RecentRefundRejectionNotice([string]$StateRoot, [string]$OrderNo = '', [datetimeoffset]$Now = [datetimeoffset]::Now) {
    # Reuse durable sync records of messages that were actually sent. Do
    # not assume a refund observation was mentioned in a chat reply.
    $runs = [Collections.Generic.List[object]]::new()
    $lastPath = Join-Path $StateRoot 'last_run.json'
    if (Test-Path -LiteralPath $lastPath -PathType Leaf) {
        try { $run = Get-Content -LiteralPath $lastPath -Encoding UTF8 -Raw | ConvertFrom-Json; if ($run) { $runs.Add($run) } } catch {}
    }
    $history = Join-Path $StateRoot 'sync_runs.ndjson'
    if (Test-Path -LiteralPath $history -PathType Leaf) {
        foreach ($line in @(Get-Content -LiteralPath $history -Encoding UTF8 -Tail 64)) {
            try { $run = $line | ConvertFrom-Json; if ($run) { $runs.Add($run) } } catch {}
        }
    }
    $notified = @{}
    $candidates = @{}
    $changesSeen = @{}
    $hours = if ($OrderNo) { 24 } else { 2 }
    foreach ($run in @($runs | Sort-Object completed_at -Descending)) {
        if ($run.ok -ne $true -or $run.dry_run -or $run.orders_skipped -or $run.sync_skipped) { continue }
        $observedAt = [datetimeoffset]::MinValue
        if (-not [datetimeoffset]::TryParse([string]$run.completed_at,[ref]$observedAt) -or $observedAt -lt $Now.AddHours(-$hours) -or $observedAt -gt $Now) { continue }
        $observations = @($run.order_sync.refund_observation_details | Where-Object { $_ })
        $realRefunds = @($run.order_sync.pending_refund_details) + @($run.order_sync.applied_refund_details)
        foreach ($real in $realRefunds) {
            if ($real.order_no -and -not $changesSeen.ContainsKey([string]$real.order_no)) { $changesSeen[[string]$real.order_no] = $true }
        }
        foreach ($observation in $observations) {
            $number = [string]$observation.order_no
            if ($number -and -not (Get-RefundRejectionNoticeKey $observation)) { $changesSeen[$number] = $true }
        }
        # A daily report can contain an observation without sending a
        # separate alert. In that case require the owner to name the order.
        if ($run.refund_observation_alert.sent -ne $true -and -not ($OrderNo -and $run.sent -eq $true)) { continue }
        foreach ($warning in @($run.warnings | Where-Object { $_ -match '^退款(?:驳回待确认|识别待核对|状态观察|待确认)：' })) {
            foreach ($match in [regex]::Matches([string]$warning,'(?<!\d)\d{18,20}(?!\d)')) {
                $number = $match.Value
                if ($OrderNo -and $number -ne $OrderNo) { continue }
                $notified[$number] = $true
                foreach ($observation in @($observations | Where-Object { $_.order_no -ceq $number })) {
                    $key = Get-RefundRejectionNoticeKey $observation
                    if ($key -and -not $changesSeen.ContainsKey($number) -and -not $candidates.ContainsKey($key)) {
                        $candidates[$key] = $observation
                    }
                }
            }
        }
    }
    $orderNumbers = @($notified.Keys | Sort-Object)
    $choices = @($candidates.Values)
    $selected = if ($orderNumbers.Count -eq 1 -and $choices.Count -eq 1) { $choices[0] } else { $null }
    return [pscustomobject]@{observation=$selected;order_numbers=$orderNumbers;candidate_count=$choices.Count}
}
