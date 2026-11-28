$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'agent\refund_notice_followup.ps1')

$base = [IO.Path]::GetTempPath()
$owned = Join-Path $base ("petstar-refund-notice-test-" + [guid]::NewGuid().ToString('N'))
if (-not [IO.Path]::GetFullPath($owned).StartsWith($base,[StringComparison]::OrdinalIgnoreCase)) { throw 'Invalid temp path' }
$null = New-Item -ItemType Directory -Path $owned
$last = Join-Path $owned 'last_run.json'
$history = Join-Path $owned 'sync_runs.ndjson'

function Write-Case($Run, [switch]$KeepHistory) {
    $json = $Run | ConvertTo-Json -Depth 8 -Compress
    [IO.File]::WriteAllText($last,$json)
    if ($KeepHistory) { [IO.File]::AppendAllText($history,$json + [Environment]::NewLine) }
}
function Check($Passed, [string]$Reason) {
    if (-not $Passed) { throw $Reason }
    $script:cases++
}
$script:cases = 0
$now = [datetimeoffset]::Now
$first = [pscustomobject]@{order_no='3400000000000000001';record_id='case-01';state='unknown';status_text='商家驳回全部退款请求';refund_amount=0;item_count=0}
$second = $first.PSObject.Copy(); $second.record_id='case-02'
$thirdOrder = $first.PSObject.Copy(); $thirdOrder.order_no='3400000000000000002'
$run = [pscustomobject]@{
    ok=$true;dry_run=$false;orders_skipped=$false;sync_skipped=$false;completed_at=$now.AddMinutes(-25).ToString('o')
    refund_observation_alert=[pscustomobject]@{sent=$true}
    warnings=@("退款识别待核对：美团1号单（$($first.order_no)）平台显示商家驳回。")
    order_sync=[pscustomobject]@{refund_observation_details=@($first);pending_refund_details=@();applied_refund_details=@()}
}

try {
    Write-Case $run -KeepHistory
    $intent = Get-RefundNoticeAcknowledgementIntent '处理好了，没有退款'
    Check ($intent -and -not $intent.order_no) 'Short reply was not identified'
    $explicit = Get-RefundNoticeAcknowledgementIntent "$($first.order_no) 处理好了，没有退款"
    Check ($explicit.order_no -eq $first.order_no) 'Explicit order was lost'
    Check (-not (Get-RefundNoticeAcknowledgementIntent '这笔退了多少钱？')) 'Question was treated as a confirmation'
    Check (-not (Get-RefundNoticeAcknowledgementIntent '处理好了，退款了')) 'Actual refund was treated as no refund'
    $context = Find-RecentRefundRejectionNotice $owned '' $now
    Check ($context.observation.record_id -eq $first.record_id) 'Unique earlier delivered notice did not resume'
    $actualReplyTime = $now.AddMinutes(-30)
    Check (-not (Find-RecentRefundRejectionNotice $owned '' $actualReplyTime).observation) 'Notice sent after user reply was matched'
    Check (-not (Find-RecentRefundRejectionNotice $owned '' $now.AddHours(3)).observation) 'Old bare reply should not be guessed'

    # Exercise the actual Feishu reply handler without loading the listener,
    # or permitting any write to a workbook or real message.
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'agent\lark_event_bridge.ps1'),[ref]$null,[ref]$null)
    $fn = $ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-RefundNoticeReply'},$true)
    Invoke-Expression $fn.Extent.Text
    $stateRoot = $owned
    function Send-Acknowledgement($OpenId,$Text,$MessageId,$Purpose) { $script:replyText = $Text }
    $event = [pscustomobject]@{message_id='test-event';create_time=[string]$now.ToUnixTimeMilliseconds()}
    Check (Invoke-RefundNoticeReply $event 'test-owner' '处理好了，没有退款') 'Short reply was sent to generic model'
    Check ($script:replyText -match '收入、商品成本和库存不变') 'Reply did not specify what was unchanged'
    $acks = @(Read-RefundNoticeAcknowledgements $owned)
    Check (Test-AcknowledgedRefundRejection $first $acks) 'Acknowledgement not saved'
    $null = Register-RefundRejectionNoRefund $owned $first
    Check (@(Read-RefundNoticeAcknowledgements $owned).Count -eq 1) 'Repeated acknowledgement duplicated'
    Check (-not (Test-AcknowledgedRefundRejection $second $acks)) 'New case with same order was hidden'
    foreach ($variation in @(@{state='pending'},@{state='completed'},@{refund_amount=4.2},@{item_count=1},@{status_text='平台介入处理中'})) {
        $reopened = $first.PSObject.Copy()
        foreach ($k in $variation.Keys) { $reopened.$k = $variation[$k] }
        Check (-not (Test-AcknowledgedRefundRejection $reopened $acks)) 'New refund evidence was suppressed'
    }

    $run.completed_at = $now.AddMinutes(-8).ToString('o')
    $run.warnings += "退款驳回待确认：美团2号单（$($thirdOrder.order_no)）"
    $run.order_sync.refund_observation_details = @($first,$thirdOrder)
    Write-Case $run -KeepHistory
    $ambiguous = Find-RecentRefundRejectionNotice $owned '' $now
    Check (-not $ambiguous.observation -and @($ambiguous.order_numbers).Count -eq 2) 'Multiple recent orders were guessed'
    Check ((Find-RecentRefundRejectionNotice $owned $thirdOrder.order_no $now).observation.order_no -eq $thirdOrder.order_no) 'Explicit order did not disambiguate'
    $null = Register-RefundRejectionNoRefund $owned $thirdOrder
    Check (@(Read-RefundNoticeAcknowledgements $owned).Count -eq 2) 'Atomic replacement lost prior record'
    $run.completed_at = $now.AddMinutes(-2).ToString('o')
    $run.refund_observation_alert.sent = $false
    $run.order_sync.applied_refund_details = @([pscustomobject]@{order_no=$first.order_no;refund_amount=6})
    Write-Case $run
    Check (-not (Find-RecentRefundRejectionNotice $owned $first.order_no $now).observation) 'Newer actual refund fell back to old rejection'

    $run | Add-Member -NotePropertyName sent -NotePropertyValue $true
    $dailyOnly = $thirdOrder.PSObject.Copy(); $dailyOnly.order_no = '3400000000000000003'
    $run.warnings = @("退款驳回待确认：订单 $($dailyOnly.order_no)")
    $run.order_sync.refund_observation_details = @($dailyOnly)
    Write-Case $run
    Check ((Find-RecentRefundRejectionNotice $owned $dailyOnly.order_no $now).observation.order_no -eq $dailyOnly.order_no) 'Explicit daily-report reply was not found'
    Check ($dailyOnly.order_no -notin (Find-RecentRefundRejectionNotice $owned '' $now).order_numbers) 'Daily report was bound to a bare reply'

    [pscustomobject]@{ok=$true;cases=$script:cases;business_writes=0;real_messages=0} | ConvertTo-Json -Compress
} finally {
    foreach ($name in @('last_run.json','sync_runs.ndjson','refund_notice_acknowledgements.json')) {
        $path = Join-Path $owned $name
        if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) }
    }
    [IO.Directory]::Delete($owned)
}
