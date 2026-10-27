$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$sourcePath = Join-Path $projectRoot 'agent\lark_event_bridge.ps1'
$source = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8
if ($source -notmatch "lark_cli_noexit\.ps1" -or $source -notmatch "不会.*结束父进程") {
    throw '飞书 CLI 长驻监听保护未启用'
}
if ($source -notmatch '经营事件执行失败：\$detail') {
    throw '经营事件底层错误透传未启用'
}
$codexReplyStart = $source.IndexOf('function Invoke-CodexAndReply')
$codexReplyEnd = $source.IndexOf('function Convert-PetstarArgumentsToMap',$codexReplyStart)
if ($codexReplyStart -lt 0 -or $codexReplyEnd -le $codexReplyStart) { throw '无法定位 Codex 回执函数' }
$codexReplyBlock = $source.Substring($codexReplyStart,$codexReplyEnd-$codexReplyStart)
if ($codexReplyBlock -match '(?m)^\s*Send-Acknowledgement\s') {
    throw 'Codex 回执结果未被抑制，可能污染布尔处理状态并把失败消息误记为完成'
}
$financeSourcePath = Join-Path $projectRoot 'agent\record_finance_event.ps1'
$financeSource = Get-Content -LiteralPath $financeSourcePath -Raw -Encoding UTF8
$resolverStart = $financeSource.IndexOf('function Resolve-UniqueSkuReference')
$resolverEnd = $financeSource.IndexOf('function Resolve-OrderForGift',$resolverStart)
if ($resolverStart -lt 0 -or $resolverEnd -le $resolverStart) { throw '无法定位 SKU 安全引用解析器' }
function Get-CellValue($Row, [int]$Index) { if ($null -eq $Row -or $Index -ge $Row.Count) { return $null }; return $Row[$Index] }
$skuSchema = [pscustomobject]@{ProductId=[pscustomobject]@{Index=2}}
Invoke-Expression $financeSource.Substring($resolverStart,$resolverEnd-$resolverStart)
$resolverRows = @(@('kw1113','csx-old','pid-shared'),@('kw1113','csx-new','pid-shared'))
$safeReference = Resolve-UniqueSkuReference $resolverRows 'sku_ref:3:pid-shared'
if ($safeReference.row_number -ne 3 -or $safeReference.cells[1] -ne 'csx-new') { throw 'SKU安全引用没有精确定位指定主表行' }
$staleReferenceRejected = $false
try { $null = Resolve-UniqueSkuReference $resolverRows 'sku_ref:3:wrong-id' } catch { $staleReferenceRejected = [string]$_.Exception.Message -match '安全定位已失效' }
if (-not $staleReferenceRejected) { throw 'SKU安全引用没有在商品ID变化时安全停止' }
if ($financeSource -notmatch 'sku_ref:\\d\+:\[a-z0-9_-\]\+') { throw '线下单解析器未接受 SKU 安全引用' }
$querySource = Get-Content -LiteralPath (Join-Path $projectRoot 'agent\query_business_data.ps1') -Raw -Encoding UTF8
if ($querySource -notmatch 'sku_ref:\$\{rowNumber\}:\$productId') { throw 'SKU查询没有返回安全 action_code' }
$codexSource = Get-Content -LiteralPath (Join-Path $projectRoot 'agent\codex_chat_bridge.py') -Raw -Encoding UTF8
if ($codexSource -notmatch '必须逐字使用它' -or $codexSource -notmatch '禁止自行改写成裸商品ID') { throw 'Codex 未被约束使用 SKU 查询返回的安全 action_code' }
$start = $source.IndexOf('$ErrorActionPreference')
$marker = "`n`$script:handledMessageHashes = Get-HandledMessageHashes"
$cut = $source.LastIndexOf($marker)
if ($start -lt 0 -or $cut -le $start) { throw '无法定位监听器函数区' }
$escapedRoot = $projectRoot.Replace("'", "''")
Invoke-Expression ("`$ProjectRoot='$escapedRoot';" + $source.Substring($start, $cut - $start))

$prettySuccess = @(
    'diagnostic line'
    '{'
    '  "ok": true,'
    '  "result": {'
    '    "order_no": "3402300350051537057"'
    '  }'
    '}'
)
$parsedPrettySuccess = Convert-CommandOutputToJson $prettySuccess
if (-not $parsedPrettySuccess.ok -or $parsedPrettySuccess.result.order_no -ne '3402300350051537057') {
    throw '多行成功 JSON 解析失败'
}

$prettyFailure = @(
    '{'
    '  "ok": false,'
    '  "error": {'
    '    "message": "库存恢复动作失败"'
    '  }'
    '}'
)
$parsedPrettyFailure = Convert-CommandOutputToJson $prettyFailure
$prettyFailureDetail = Get-CommandFailureDetail $parsedPrettyFailure $prettyFailure
if ($prettyFailureDetail -ne '库存恢复动作失败') { throw "多行失败 JSON 原因提取失败：$prettyFailureDetail" }

$argparseFailure = @(
    'usage: petstar_cli.py event [-h] --action ACTION'
    'petstar_cli.py event: error: argument --platform: invalid choice: 京东外卖'
)
$argparseFailureDetail = Get-CommandFailureDetail $null $argparseFailure
if ($argparseFailureDetail -notmatch 'argument --platform: invalid choice') { throw "CLI 参数错误没有提取具体原因：$argparseFailureDetail" }

$longPurposeKey = New-AcknowledgementIdempotencyKey 'confirmation-error' 'om_abcdefghijklmnopqrstuvwxyz0123456789abcdefghijklmnopqrstuvwxyz'
if ($longPurposeKey.Length -gt 50 -or -not $longPurposeKey.StartsWith('petstar-confirmation-error-')) {
    throw "飞书回执幂等键长度不合规：$($longPurposeKey.Length)"
}

$testStateRoot = Join-Path $env:TEMP ("petstar-confirmation-test-" + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $testStateRoot
$stateRoot = $testStateRoot
$pendingConfirmationsPath = Join-Path $testStateRoot 'pending.json'
$pendingAcknowledgementsPath = Join-Path $testStateRoot 'pending_ack.json'
$handledMessagesLog = Join-Path $testStateRoot 'handled.ndjson'
$runtimeLog = Join-Path $testStateRoot 'bridge.log'
$interactivePriorityPath = Join-Path $testStateRoot 'interactive_priority.json'
$script:testAckNowShouldFail = $true
$script:testAckNowCalls = 0
function Send-AcknowledgementNow([string]$OpenId, [string]$Text, [string]$IdempotencyKey) {
    $script:testAckNowCalls++
    if ($script:testAckNowShouldFail) { throw '模拟飞书临时不可用' }
}
$queuedAck = Send-Acknowledgement 'user-1' '测试持久回执' 'outbox-message-1' 'ack'
if (-not $queuedAck.queued -or @(Read-PendingAcknowledgementsUnsafe).Count -ne 1) { throw '飞书发送失败后没有进入持久回执队列' }
$script:testAckNowShouldFail = $false
$flushedAck = Flush-PendingAcknowledgements
if ($flushedAck.sent -ne 1 -or $flushedAck.failed -ne 0 -or @(Read-PendingAcknowledgementsUnsafe).Count -ne 0) { throw '持久回执队列恢复发送失败' }
$script:testAcks = [Collections.Generic.List[string]]::new()
$script:testCalls = [Collections.Generic.List[object]]::new()
$script:testCodexCalls = [Collections.Generic.List[object]]::new()
$script:testBatchCalls = [Collections.Generic.List[object]]::new()
$script:failNextAction = ''
$script:testCostRebuildCalls = 0
$script:previewMode = $false
$script:confirmedExecution = $false
$script:handledMessageHashes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

function Send-Acknowledgement([string]$OpenId, [string]$Text, [string]$MessageId, [string]$Purpose = 'ack') {
    if ($script:previewMode) { $script:previewAcknowledgement = $Text; return }
    $script:testAcks.Add($Text)
}
function Write-BridgeLog([string]$Message) {}
function Add-ProcessingReaction([string]$MessageId) { return $null }
function Remove-ProcessingReaction([string]$MessageId, [string]$ReactionId) {}
function Invoke-DailyLedgerSync([datetime]$ThroughDate) {
    return [pscustomobject]@{
        ok=$true
        formula_verification='success'
        months=@([pscustomobject]@{month='2026-08';expense_rows=60},[pscustomobject]@{month='2026-09';expense_rows=68})
        requested_date_summary=[pscustomobject]@{date=$ThroughDate.ToString('yyyy-MM-dd');order_count=6;income=295.4;product_cost=138.2;delivery_fee=43.08;profit=114.12;pending_cost_count=0;pending_orders=@()}
    }
}
function Invoke-CostSnapshotRebuild {
    $script:testCostRebuildCalls++
    return [pscustomobject]@{ok=$true;item_rows=12;natural_costed=8;natural_pending_cost=0}
}
function New-TestActionArguments([hashtable]$Values) {
    $arguments = [ordered]@{
        date=$null;order_no=$null;daily_order_no=$null;person=$null;items=$null;goods_code=$null;transition_code=$null;new_goods_code=$null
        retire_existing_new_code=$null;expiry_date=$null;clear_expiry=$null;production_date=$null;shelf_life_value=$null
        shelf_life_unit=$null;quantity=$null;purchase_amount=$null;payment_amount=$null;payment_date=$null;amount=$null
        delivery_fee=$null;receipt_status=$null;platform=$null;note=$null;mode=$null
    }
    foreach ($key in $Values.Keys) { $arguments[$key] = $Values[$key] }
    return [pscustomobject]$arguments
}
function New-TestCodexAction([string]$Action,[hashtable]$Values) {
    return [pscustomobject]@{action=$Action;arguments=(New-TestActionArguments $Values)}
}
function Invoke-PetstarCodex([string]$Text, [string]$MessageId, [string]$Mode = 'normal') {
    $script:testCodexCalls.Add([pscustomobject]@{text=$Text;mode=$Mode})
    $actions = @()
    $reply = '本机 Codex 正常回答。'
    $needsInformation = $false
    if ($Text -match '模拟预览限流') {
        return [pscustomobject]@{ok=$false;cached=$false;retryable=$true;retry_after_seconds=120;error='Codex 当前请求较多，经营数据汇总已延后；系统会自动重试，无需重发。';reply='';actions=@()}
    } elseif ($Text -match '组合线下成本测试') {
        $actions = @(
            (New-TestCodexAction 'SetSkuCost' @{goods_code='kw1109';amount=0;note='组合预演成本'}),
            (New-TestCodexAction 'OfflineSale' @{date='2026-09-13';items='kw1109×4';amount=8;delivery_fee=0;receipt_status='已到账';platform='线下';note='组合预演线下单'})
        )
        $reply = '📝 待确认：先把成本改为0，再登记线下销售。'
    } elseif ($Text -match '成本改成0元，但先核对') {
        $reply = '成本证据仍需核对；原待确认修改已保留。'
        $needsInformation = $true
    } elseif ($Text -match '同步一下八九月支出') {
        $actions = @(New-TestCodexAction 'RefreshLedger' @{date='2026-09-13'})
        $reply = '刷新八九月支出。'
    } elseif ($Text -match '重新更新一遍9月收入表中今天的订单情况') {
        $actions = @(New-TestCodexAction 'RefreshLedger' @{date='2026-09-13'})
        $reply = '按修改后的 SKU 成本刷新今天的订单核算。'
    } elseif ($Text -match '3402294620217435837') {
        $actions = @(
            (New-TestCodexAction 'AddIntervention' @{date='2026-09-12';order_no='3402294620217435837';person='爸爸'}),
            (New-TestCodexAction 'ConfirmSelfDelivery' @{order_no='3402294620217435837'}),
            (New-TestCodexAction 'RestoreInventory' @{order_no='3402294620217435837'})
        )
        $reply = '📝 待确认：登记干预单、确认自配送并恢复库存。'
    } elseif ($Text -match '3302310843791845601' -and $Text -match '3202310773147918261') {
        $actions = @(
            (New-TestCodexAction 'RemoveIntervention' @{order_no='3302310843791845601';note='纠正误登记的干预单'}),
            (New-TestCodexAction 'AddIntervention' @{date='2026-09-13';order_no='3202310773147918261';person='wikar';note='纠正后的真实干预单'}),
            (New-TestCodexAction 'ConfirmSelfDelivery' @{order_no='3202310773147918261'}),
            (New-TestCodexAction 'RestoreInventory' @{order_no='3202310773147918261'})
        )
        $reply = '📝 待确认：撤销错误干预，并把正确订单登记为干预。'
    } elseif ($Text -match '没有动作的修改测试') {
        $reply = '📝 待确认：执行层暂不支持这个修改。'
    } elseif ($Text -match '3402294850565247821') {
        $fee = if ($Text -match '改成8') { 8 } else { 6.5 }
        $actions = @(New-TestCodexAction 'SetOrderDeliveryFee' @{date='2026-09-12';order_no='3402294850565247821';amount=$fee})
        $reply = "📝 待确认：配送费改为 $fee 元。"
    } elseif ($Text -match '批量SKU维护测试') {
        $actions = @(
            (New-TestCodexAction 'Replenishment' @{date='2026-09-13';goods_code='kw1104';quantity=2;purchase_amount=20;note='批量测试补货'}),
            (New-TestCodexAction 'SetSkuCost' @{date='2026-09-13';goods_code='kw1109';amount=1.25;note='批量测试成本'}),
            (New-TestCodexAction 'Replenishment' @{date='2026-09-13';goods_code='kw1113';quantity=3;purchase_amount=0;note='批量测试零成本补货'})
        )
        $reply = '📝 待确认：批量登记两笔补货和一笔普通成本修改。'
    } elseif ($Text -match 'kw1113') {
        $actions = @(New-TestCodexAction 'Replenishment' @{date='2026-09-13';goods_code='kw1113';quantity=12;purchase_amount=0;production_date='2026-08-26';shelf_life_value=18;shelf_life_unit='月';note='试吃装，成本为0'})
        $reply = '📝 待确认：登记 kw1113 零成本试吃装补货。'
    } elseif ($Text -match '模糊补货') {
        $actions = @(New-TestCodexAction 'Replenishment' @{date='2026-09-12';goods_code='kw1104';quantity=6;purchase_amount=330;payment_amount=330;payment_date='2026-09-12';production_date='2026-08-28';shelf_life_value=18;shelf_life_unit='个月';note='补货测试'})
        $reply = '📝 待确认：登记昨天 kw1104 补货 6 包，采购并付款 330 元。'
    } elseif ($Text -match 'kw1048.*生产日期') {
        $actions = @(New-TestCodexAction 'SetSkuBatchDate' @{goods_code='kw1048';production_date='2026-08-13';shelf_life_value=12;shelf_life_unit='月'})
        $reply = '📝 待确认：更新 kw1048 批次日期。'
    } elseif ($Text -match '历史订单成本修正') {
        $actions = @(New-TestCodexAction 'SetSkuCost' @{goods_code='kw1109';amount=0;order_no='888888888888888888';note='明确修正指定历史订单成本'})
        $reply = '📝 待确认：修改 kw1109 成本，并只联动指定历史订单。'
    } elseif ($Text -match 'kw1109.*成本') {
        $actions = @(New-TestCodexAction 'SetSkuCost' @{goods_code='kw1109';amount=0;note='试吃装成本为0'})
        $reply = '📝 待确认：把 kw1109 每个销售单位成本改为0元。'
    } elseif ($Text -match '999999999999999999') {
        $actions = @(New-TestCodexAction 'AddOrderGift' @{date='2026-09-13';order_no='999999999999999999';items='kw4189×1'})
        $reply = '📝 待确认：提前登记订单赠品。'
    }
    return [pscustomobject]@{ok=$true;cached=$false;needs_information=$needsInformation;reply=$reply;actions=$actions}
}
function Invoke-PetstarEvent([string[]]$Arguments) {
    $map = Convert-PetstarArgumentsToMap $Arguments
    $action = [string]$map.action
    if (-not $script:previewMode -and $script:failNextAction -eq $action) {
        $script:failNextAction = ''
        throw "模拟执行失败：$action"
    }
    $amount = if ($map.amount) { [double]$map.amount } else { 0.0 }
    $result = switch ($action) {
        'AddIntervention' { [pscustomobject]@{order_no=$map.order_no;person=$map.person} }
        'RemoveIntervention' { [pscustomobject]@{row=86;order_no=$map.order_no;person='wikar';delivery_type='自配送';delivery_fee=0;inventory_state='已恢复';new_order_type='自然'} }
        'ConfirmSelfDelivery' { [pscustomobject]@{order_no=$map.order_no;delivery_type='自配送';delivery_fee=0;total_cost=0} }
        'SetExternalDeliveryFee' { [pscustomobject]@{order_no=$map.order_no;delivery_type='骑手配送';delivery_fee=$amount;total_cost=$amount} }
        'RestoreInventory' { [pscustomobject]@{order_no=$map.order_no;state='已恢复'} }
        'OfflineSale' { [pscustomobject]@{daily_order_no='线下1号单';sold_goods='kw4189×1';income=5;product_cost=2.1;delivery_fee=0} }
        'AddOrderGift' {
            $isPendingGift = [string]$map.order_no -eq '999999999999999999'
            [pscustomobject]@{status=$(if($isPendingGift){'pending_order'}else{'matched'});order_no=$map.order_no;daily_order_no=$(if($isPendingGift){'待同步'}else{'美团1号单'});gift_items='kw4189×1';gift_cost=2.1;pending_cost_codes=@();accounting=$(if($isPendingGift){'待匹配'}else{'并入货物成本'})}
        }
        'Replenishment' { [pscustomobject]@{goods_code=$map.goods_code;package_quantity=[double]$map.quantity;purchase_amount=[double]$map.purchase_amount;payment_amount=$(if($null-ne$map.payment_amount){[double]$map.payment_amount}else{$null});after_average_cost=$(if([double]$map.purchase_amount -eq 0){0}else{55});production_date=$map.production_date;expiry_date=$(if($map.goods_code -eq 'kw1113'){'2028-02-26'}else{'2028-02-28'});sku_status=$(if($map.goods_code -eq 'kw1113'){$(if($script:previewMode){'will_create_pending_profile'}else{'created_pending_profile'})}else{'matched'})} }
        'ChangeSkuCode' { [pscustomobject]@{status='ready';old_code=$map.goods_code;new_code=$map.new_goods_code} }
        'SetOrderDeliveryFee' { [pscustomobject]@{order_no=$map.order_no;order_type='自然单';date='2026-09-12';daily_order_no='美团2号单';old_delivery_fee=0;delivery_fee=$amount} }
        'SetSkuBatchDate' { [pscustomobject]@{goods_code=$map.goods_code;production_date=$map.production_date;shelf_life="$($map.shelf_life_value)$($map.shelf_life_unit)";expiry_date='2027-08-13'} }
        'SetSkuCost' { [pscustomobject]@{sheet='SKU成本主表';row=19;range='F19:G19,K19,M19:O19';goods_code=$map.goods_code;title='测试试吃装';old_baseline_cost=3.5;old_moving_average_cost=3.5;new_unit_cost=[double]$map.amount;new_moving_average_cost=[double]$map.amount;old_status='待人工核对';new_status='用户确认';historical_order_no=$map.order_no;affected_tables=$(if($map.order_no){@('SKU成本主表','指定订单的订单商品','订单汇总','9月收入','驾驶舱')}else{@('SKU成本主表')})} }
        default { [pscustomobject]@{} }
    }
    if ($script:previewMode) {
        $script:previewOperations.Add([pscustomobject]@{action=$action;arguments=$map;result=$result})
    } else {
        $script:testCalls.Add([pscustomobject]@{action=$action;arguments=$map})
    }
    return [pscustomobject]@{ok=$true;duplicate=$false;result=$result}
}

function Invoke-PetstarBatchSkuMaintenance($StoredOperations, $Event, [int]$StartIndex, [switch]$Preview) {
    $items = @($StoredOperations)
    $script:testBatchCalls.Add([pscustomobject]@{count=$items.Count;start_index=$StartIndex;actions=@($items.action);preview=[bool]$Preview})
    $executed = [Collections.Generic.List[object]]::new()
    foreach ($stored in $items) {
        $args = $stored.arguments
        $details = if ([string]$stored.action -eq 'Replenishment') {
            [pscustomobject]@{row=80+$executed.Count;goods_code=$args.goods_code;package_quantity=[double]$args.quantity;purchase_amount=[double]$args.purchase_amount;after_average_cost=$(if([double]$args.purchase_amount-eq 0){0}else{10});sku_status=$(if($Preview -and $args.goods_code -eq 'kw1113'){'will_create_pending_profile'}else{'matched'})}
        } else {
            [pscustomobject]@{sheet='SKU成本主表';row=19;goods_code=$args.goods_code;old_baseline_cost=3.5;new_unit_cost=[double]$args.amount;new_moving_average_cost=[double]$args.amount;historical_order_no=$null}
        }
        $executed.Add([pscustomobject]@{action=[string]$stored.action;arguments=$args;result=$details;formula_verification='success';duplicate=$false})
    }
    return $executed.ToArray()
}

$script:previewMode = $true
$script:previewOperations = [Collections.Generic.List[object]]::new()
$aliasAction = New-TestCodexAction 'OfflineSale' @{
    date='2026-09-24';items='KW5011×1、KW2089×1';amount=34.92;delivery_fee=0
    receipt_status='当天到账';platform='京东外卖';note='平台枚举归一化测试'
}
$aliasOperation = Invoke-CodexAction $aliasAction ([pscustomobject]@{message_id='alias-normalization';content='';create_time=''}) 0 'preview'
if ([string]$aliasOperation.arguments.receipt_status -ne '已到账' -or [string]$aliasOperation.arguments.platform -ne '京东到家') {
    throw "自然语言枚举没有在进入 CLI 前归一化：$($aliasOperation.arguments | ConvertTo-Json -Compress)"
}
$script:previewMode = $false
$script:previewOperations = [Collections.Generic.List[object]]::new()

$testCreateTime = [DateTimeOffset]::new(2026,9,13,12,0,0,[timespan]::FromHours(8)).ToUnixTimeMilliseconds().ToString()
$baseEvent = @{chat_id='chat-1';chat_type='p2p';message_type='text';sender_type='user';sender_id='user-1';create_time=$testCreateTime}
function Invoke-TestMessage([string]$Id,[string]$Text) {
    $event=$baseEvent.Clone();$event.message_id=$Id;$event.content=$Text
    $script:testAcks.Clear()
    $ok=Process-Event ($event|ConvertTo-Json -Compress) 'user-1'
    return [pscustomobject]@{ok=$ok;ack=($script:testAcks -join "`n")}
}

try {
    $manual = Invoke-TestMessage 'manual-sync' '麻烦同步一下八九月支出'
    if (-not $manual.ok -or $manual.ack -notmatch '已重新计算 2026-09-13' -or $manual.ack -notmatch '当天赚/亏 ¥114.12' -or $manual.ack -notmatch '没有修改人工录入字段') { throw '只读同步应返回指定日期的实际重算结果' }
    if ($script:testCostRebuildCalls -ne 1) { throw '只读同步未先重建成本快照' }

    $script:testCalls.Clear()
    $draft = Invoke-TestMessage 'draft-intervention' '昨天3402294620217435837是干预单，补单人爸爸，我自己送，没有配送费，库存已恢复'
    if ($draft.ack -notmatch '待确认' -or $draft.ack -notmatch '2026-09-12' -or $script:testCalls.Count -ne 0) { throw "干预单预览失败或预览阶段发生写入：ack=$($draft.ack) calls=$($script:testCalls.Count)" }
    $pendingBeforeRefresh = Read-PendingConfirmation
    $costCallsBeforeRefresh = $script:testCostRebuildCalls
    $refreshWhilePending = Invoke-TestMessage 'refresh-while-pending' '你重新更新一遍9月收入表中今天的订单情况，我刚刚发现，有一些 SKU 的成本填错了，所以改了一下成本。想让你重新更新一下今天的订单情况。'
    $pendingAfterRefresh = Read-PendingConfirmation
    if ($refreshWhilePending.ack -notmatch '已重新计算 2026-09-13' -or $script:testCostRebuildCalls -ne ($costCallsBeforeRefresh + 1)) { throw "存在待确认项时，独立刷新请求没有直接完成：$($refreshWhilePending.ack)" }
    if (-not $pendingAfterRefresh -or $pendingAfterRefresh.change_id -ne $pendingBeforeRefresh.change_id -or @($pendingAfterRefresh.operations).Count -ne @($pendingBeforeRefresh.operations).Count) { throw '独立刷新请求错误修改或替换了原待确认项' }
    $codexCallsBeforeConfirm = $script:testCodexCalls.Count
    $confirm = Invoke-TestMessage 'confirm-intervention' '确认'
    $actions = @($script:testCalls | ForEach-Object {$_.action})
    if ($confirm.ack -notmatch '已按确认内容完成修改' -or $confirm.ack -notmatch '干预订单第' -or $actions.Count -ne 3 -or $actions -notcontains 'AddIntervention' -or $actions -notcontains 'ConfirmSelfDelivery' -or $actions -notcontains 'RestoreInventory') { throw "干预单确认执行失败：ack=$($confirm.ack) actions=$($actions -join ',')" }
    if ($script:testCodexCalls.Count -ne $codexCallsBeforeConfirm) { throw '确认阶段不应再次调用 Codex 重解释原消息' }
    if ((Read-PendingConfirmation)) { throw '确认后未清空待确认状态' }

    $script:testCalls.Clear()
    $correctionDraft = Invoke-TestMessage 'draft-intervention-correction' '3302310843791845601不是干预单，真正的wikar干预单是3202310773147918261，还是自配送且库存已恢复'
    if ($correctionDraft.ack -notmatch '撤销错误干预' -or $correctionDraft.ack -notmatch '3302310843791845601' -or $correctionDraft.ack -notmatch '3202310773147918261' -or $script:testCalls.Count -ne 0) { throw "干预单纠错预览不完整：$($correctionDraft.ack)" }
    $correctionConfirm = Invoke-TestMessage 'confirm-intervention-correction' '确认'
    $correctionActions = @($script:testCalls | ForEach-Object action)
    if ($correctionConfirm.ack -notmatch '已撤销错误干预' -or $correctionActions.Count -ne 4 -or $correctionActions[0] -ne 'RemoveIntervention' -or $correctionActions[1] -ne 'AddIntervention' -or $correctionActions -notcontains 'ConfirmSelfDelivery' -or $correctionActions -notcontains 'RestoreInventory') { throw "干预单纠错没有按同一确认完整执行：ack=$($correctionConfirm.ack) actions=$($correctionActions -join ',')" }

    $noActionReply = Invoke-TestMessage 'no-action-fake-preview' '请做一个没有动作的修改测试'
    if ($noActionReply.ack -match '待确认' -or $noActionReply.ack -notmatch '没有生成可执行修改' -or (Read-PendingConfirmation)) { throw "无执行动作时仍发送了误导性的待确认：$($noActionReply.ack)" }

    $combinedOfflineDraft = Invoke-TestMessage 'combined-offline-cost-preview' '组合线下成本测试'
    $combinedPending = Read-PendingConfirmation
    if ($combinedOfflineDraft.ack -notmatch '成本 ¥0' -or @($combinedPending.operations).Count -ne 2 -or [double]@($combinedPending.operations)[1].result.product_cost -ne 0) { throw "组合操作预演没有按前序成本修改计算线下单：$($combinedOfflineDraft.ack)" }
    $null = Clear-PendingConfirmation

    $script:testCalls.Clear()
    $feeDraft = Invoke-TestMessage 'draft-fee' '昨天订单3402294850565247821配送费是6.5元'
    if ($feeDraft.ack -notmatch '¥0 → ¥6.5') { throw '自然单配送费预览失败' }
    $feePendingBeforeClarification = Read-PendingConfirmation
    $feeClarification = Invoke-TestMessage 'clarify-fee-without-action' '成本改成0元，但先核对'
    $feePendingAfterClarification = Read-PendingConfirmation
    if ($feeClarification.ack -notmatch '原待确认修改已保留' -or -not $feePendingAfterClarification -or $feePendingAfterClarification.change_id -ne $feePendingBeforeClarification.change_id -or $feePendingAfterClarification.operation_signature -ne $feePendingBeforeClarification.operation_signature -or @($feePendingAfterClarification.operations).Count -ne 1) {
        throw "无动作补充说明覆盖了原待确认项：ack=$($feeClarification.ack)"
    }
    $feeAmend = Invoke-TestMessage 'amend-fee' '配送费改成8元'
    if ($feeAmend.ack -notmatch '¥0 → ¥8') { throw "待确认内容修改失败：$($feeAmend.ack)" }
    $feeConfirm = Invoke-TestMessage 'confirm-fee' '确认'
    $feeCall = @($script:testCalls | Where-Object action -eq 'SetOrderDeliveryFee')[-1]
    if (-not $feeCall -or [double]$feeCall.arguments.amount -ne 8 -or $feeConfirm.ack -notmatch '¥8') { throw '自然单配送费确认执行失败' }

    $batchDraft = Invoke-TestMessage 'draft-batch' 'kw1048的生产日期是2026年8月13日，保质期12个月'
    if ($batchDraft.ack -notmatch '生产 2026-08-13' -or $batchDraft.ack -notmatch '到期 2027-08-13') { throw '生产日期与到期日期预览失败' }
    $cancel = Invoke-TestMessage 'cancel-batch' '取消'
    if ($cancel.ack -notmatch '没有修改表格' -or (Read-PendingConfirmation)) { throw '取消待确认失败' }

    $script:testCalls.Clear()
    $codexDraft = Invoke-TestMessage 'draft-codex-replenishment' '把昨天那批模糊补货也记一下'
    if ($codexDraft.ack -notmatch '待确认' -or $codexDraft.ack -notmatch 'kw1104' -or $codexDraft.ack -notmatch '2026-09-12' -or $script:testCalls.Count -ne 0) { throw "Codex 结构化补货预览失败：$($codexDraft.ack)" }
    $codexConfirm = Invoke-TestMessage 'confirm-codex-replenishment' '确认'
    $replenishmentCall = @($script:testCalls | Where-Object action -eq 'Replenishment')[-1]
    if (-not $replenishmentCall -or [double]$replenishmentCall.arguments.quantity -ne 6 -or [double]$replenishmentCall.arguments.purchase_amount -ne 330 -or [string]$replenishmentCall.arguments.shelf_life_unit -ne '月' -or $codexConfirm.ack -notmatch '补货记录第.*kw1104') { throw "Codex 结构化补货确认执行失败或保质期单位未归一化：$($codexConfirm.ack)" }

    $script:testCalls.Clear()
    $zeroCostDraft = Invoke-TestMessage 'draft-zero-cost-replenishment' '今天补货 KW1113，都是试吃装，成本为0，生产日期为2026年8月26日，保质期为18个月，数量为12个'
    if ($zeroCostDraft.ack -notmatch '补货：kw1113' -or $zeroCostDraft.ack -notmatch '采购 ¥0' -or $zeroCostDraft.ack -notmatch '2028-02-26' -or $zeroCostDraft.ack -notmatch '待平台同步档案' -or $script:testCalls.Count -ne 0) { throw "零成本新SKU补货预览失败：$($zeroCostDraft.ack)" }
    $zeroCostConfirm = Invoke-TestMessage 'confirm-zero-cost-replenishment' '确认'
    $zeroCostCall = @($script:testCalls | Where-Object action -eq 'Replenishment')[-1]
    if (-not $zeroCostCall -or [double]$zeroCostCall.arguments.purchase_amount -ne 0 -or [double]$zeroCostCall.arguments.quantity -ne 12 -or $zeroCostConfirm.ack -notmatch '待平台同步') { throw "零成本新SKU补货确认执行失败：ack=$($zeroCostConfirm.ack) call=$($zeroCostCall|ConvertTo-Json -Compress -Depth 6)" }

    $script:testCalls.Clear()
    $script:testBatchCalls.Clear()
    $batchSkuDraft = Invoke-TestMessage 'draft-batch-sku-maintenance' '批量SKU维护测试'
    if ($batchSkuDraft.ack -notmatch '待确认' -or $batchSkuDraft.ack -notmatch '不会重算历史订单' -or $script:testBatchCalls.Count -ne 1 -or -not [bool]$script:testBatchCalls[0].preview -or [int]$script:testBatchCalls[0].count -ne 3 -or $script:testCalls.Count -ne 0) {
        throw "批量 SKU 维护预览没有一次处理全部兼容动作：ack=$($batchSkuDraft.ack) batch=$($script:testBatchCalls|ConvertTo-Json -Compress -Depth 8) single=$($script:testCalls.Count)"
    }
    $batchSkuConfirm = Invoke-TestMessage 'confirm-batch-sku-maintenance' '确认'
    $executionBatchCalls = @($script:testBatchCalls | Where-Object { -not [bool]$_.preview })
    if ($script:testBatchCalls.Count -ne 2 -or $executionBatchCalls.Count -ne 1 -or [int]$executionBatchCalls[0].count -ne 3 -or $script:testCalls.Count -ne 0 -or $batchSkuConfirm.ack -notmatch '补货记录第' -or $batchSkuConfirm.ack -notmatch '未修改历史订单或经营简账') {
        throw "批量 SKU 维护没有一次执行全部兼容动作：ack=$($batchSkuConfirm.ack) batch=$($script:testBatchCalls|ConvertTo-Json -Compress -Depth 8) single=$($script:testCalls.Count)"
    }

    $codexQuestion = Invoke-TestMessage 'codex-question' '你现在用什么方式理解我的话？'
    if (-not $codexQuestion.ok -or $codexQuestion.ack -notmatch '本机 Codex 正常回答') { throw 'Codex 普通对话回退失败' }

    $script:testCalls.Clear()
    $pendingGiftDraft = Invoke-TestMessage 'draft-pending-gift' '999999999999999999这单送了kw4189×1作为赠品'
    if ($pendingGiftDraft.ack -notmatch '订单尚未同步' -or $script:testCalls.Count -ne 0) { throw "未同步订单赠品预览失败：$($pendingGiftDraft.ack)" }
    $pendingGiftConfirm = Invoke-TestMessage 'confirm-pending-gift' '确认'
    $giftCall = @($script:testCalls | Where-Object action -eq 'AddOrderGift')[-1]
    if (-not $giftCall -or $pendingGiftConfirm.ack -notmatch '赠品记录' -or $pendingGiftConfirm.ack -notmatch '等待订单同步') { throw "未同步订单赠品确认失败：$($pendingGiftConfirm.ack)" }

    $script:testCalls.Clear()
    $costDraft = Invoke-TestMessage 'draft-sku-cost' '把kw1109的销售单位成本改成0元，这是试吃装'
    if ($costDraft.ack -notmatch 'SKU成本主表第19行' -or $costDraft.ack -notmatch '¥3.5 → ¥0' -or $costDraft.ack -match '订单商品' -or $costDraft.ack -notmatch '不会自动重算历史订单' -or $script:testCalls.Count -ne 0) { throw "普通SKU成本修改错误联动历史订单：$($costDraft.ack)" }
    $costConfirm = Invoke-TestMessage 'confirm-sku-cost' '确认'
    $costCall = @($script:testCalls | Where-Object action -eq 'SetSkuCost')[-1]
    if (-not $costCall -or [double]$costCall.arguments.amount -ne 0 -or $costCall.arguments.order_no -or $costConfirm.ack -notmatch 'SKU成本主表第19行' -or $costConfirm.ack -notmatch '未修改历史订单') { throw "普通SKU成本确认错误联动历史订单：$($costConfirm.ack)" }

    $script:testCalls.Clear()
    $targetedCostDraft = Invoke-TestMessage 'draft-targeted-order-cost' '历史订单成本修正：订单888888888888888888的kw1109成本改成0元'
    if ($targetedCostDraft.ack -notmatch '明确联动历史订单：888888888888888888' -or $targetedCostDraft.ack -notmatch '指定订单的订单商品' -or $script:testCalls.Count -ne 0) { throw "指定历史订单成本预览失败：$($targetedCostDraft.ack)" }
    $targetedCostConfirm = Invoke-TestMessage 'confirm-targeted-order-cost' '确认'
    $targetedCostCall = @($script:testCalls | Where-Object action -eq 'SetSkuCost')[-1]
    if (-not $targetedCostCall -or $targetedCostCall.arguments.order_no -ne '888888888888888888' -or $targetedCostConfirm.ack -notmatch '已按明确要求更新订单 888888888888888888') { throw "指定历史订单成本确认失败：$($targetedCostConfirm.ack)" }

    $script:testCalls.Clear()
    $failureDraft = Invoke-TestMessage 'draft-confirmation-failure' '昨天3402294620217435837是干预单，补单人爸爸，我自己送，没有配送费，库存已恢复'
    $failureChangeId = (Read-PendingConfirmation).change_id
    $script:failNextAction = 'ConfirmSelfDelivery'
    $failureConfirm = Invoke-TestMessage 'confirm-with-explicit-failure' '确认'
    if ($failureConfirm.ack -notmatch [regex]::Escape($failureChangeId) -or $failureConfirm.ack -notmatch '3402294620217435837' -or $failureConfirm.ack -notmatch '失败阶段：执行已确认修改' -or $failureConfirm.ack -notmatch '仍保留在待确认队列' -or -not (Read-PendingConfirmation)) {
        throw "确认失败提示不明确或错误清空了待确认项：$($failureConfirm.ack)"
    }
    $pendingAfterFailure = Read-PendingConfirmation
    if (@($pendingAfterFailure.completed_operations | Where-Object { $null -ne $_ }).Count -ne 0 -or $failureConfirm.ack -notmatch '全部回滚') { throw "确认失败后不应保留半成品进度：$($pendingAfterFailure | ConvertTo-Json -Compress -Depth 20)" }
    $failureRetry = Invoke-TestMessage 'retry-confirmation-failure' '确认'
    if ($failureRetry.ack -notmatch '已按确认内容完成修改' -or (Read-PendingConfirmation)) { throw '失败后的确认重试没有成功完成' }
    $failureActionNames = @($script:testCalls | ForEach-Object { $_.action })
    if (@($failureActionNames | Where-Object { $_ -eq 'AddIntervention' }).Count -ne 2 -or @($failureActionNames | Where-Object { $_ -eq 'ConfirmSelfDelivery' }).Count -ne 1 -or @($failureActionNames | Where-Object { $_ -eq 'RestoreInventory' }).Count -ne 1) { throw "原子回滚后的整条重试范围不正确：$($failureActionNames -join ',')" }

    $script:testCalls.Clear()
    $firstQueued = Invoke-TestMessage 'queue-first' '昨天3402294620217435837是干预单，补单人爸爸，我自己送，没有配送费，库存已恢复'
    $secondQueued = Invoke-TestMessage 'queue-second' '把kw1109的销售单位成本改成0元，这是试吃装'
    $queueState = Read-PendingConfirmationState
    if ($secondQueued.ack -notmatch '第 2 项待确认' -or @($queueState.queue).Count -ne 1 -or $script:testCalls.Count -ne 0) { throw "待确认队列没有保存第二项：$($secondQueued.ack)" }
    $firstConfirm = Invoke-TestMessage 'queue-confirm-first' '确认'
    $afterFirstQueueState = Read-PendingConfirmationState
    if ($firstConfirm.ack -notmatch '下一项待确认' -or $afterFirstQueueState.active.change_id -ne @($queueState.queue)[0].change_id -or @($afterFirstQueueState.queue).Count -ne 0) { throw '第一项完成后没有正确展示下一项，或残留了空队列项' }
    $secondConfirm = Invoke-TestMessage 'queue-confirm-second' '确认'
    if ((Read-PendingConfirmation) -or @($script:testCalls | Where-Object action -eq 'SetSkuCost').Count -ne 1) { throw '第二项待确认没有按队列执行' }

    $stalePending = [pscustomobject]@{change_id='DEADBEEF';conversation_hash=(Get-ConversationHash ([pscustomobject]$baseEvent));original_message='旧说明';effective_message='旧说明';original_create_time=$testCreateTime;created_at=[datetimeoffset]::Now.ToString('o');expires_at=[datetimeoffset]::Now.AddHours(1).ToString('o');preview='只是一段说明';needs_information=$true;mutation_count=0;operations=@();operation_signature='[]'}
    Write-PendingConfirmation $stalePending
    $staleRecovery = Invoke-TestMessage 'recover-empty-pending' '把kw1109的销售单位成本改成0元，这是试吃装'
    $recoveredPending = Read-PendingConfirmation
    if ($staleRecovery.ack -notmatch 'SKU成本主表第19行' -or -not $recoveredPending -or $recoveredPending.change_id -eq 'DEADBEEF' -or @($recoveredPending.operations).Count -ne 1) { throw "零操作旧待确认项没有被安全恢复：$($staleRecovery.ack)" }
    $null = Clear-PendingConfirmation

    $rateLimitEvent = $baseEvent.Clone(); $rateLimitEvent.message_id='preview-rate-limit'; $rateLimitEvent.content='模拟预览限流并修改成本'
    $script:testAcks.Clear()
    $rateLimitResult = Process-EventWithLedger ($rateLimitEvent | ConvertTo-Json -Compress) 'user-1' 'test'
    $rateLimitAck = $script:testAcks -join "`n"
    if ($rateLimitResult -ne 'failed_retryable' -or $rateLimitAck -notmatch '约 2 分钟后自动重试' -or $rateLimitAck -match '无法生成待确认预览') { throw "预览阶段限流没有转入安全重试：result=$rateLimitResult ack=$rateLimitAck" }

    if (@($script:testCodexCalls | Where-Object {$_.text -match '3402294620217435837'}).Count -eq 0) { throw '干预单自然语言没有经过 Codex' }
    [ordered]@{ok=$true;command_json='multiline-supported';ack_idempotency='max-50';ack_outbox='queue-and-retry';routing='codex-first';manual_sync='codex-dispatched';intervention='preview-confirm';intervention_correction='remove-add-atomic';non_executable_reply='not-pending';confirmation_execution='stored-plan-no-reparse';confirmation_failure='atomic-rollback-and-full-retry';delivery_fee='preview-amend-confirm';pending_amendment='zero-action-retains-original';stale_pending='auto-recovered';preview_rate_limit='deferred-retry';batch_date='calculated-preview-cancel';zero_cost_new_sku='preview-confirm';codex_question='direct-reply';pending_gift='preview-confirm-auto-match';sku_cost='explicit-preview-confirm';batch_sku_maintenance='preview-and-execute-single-call';pending_queue='fifo-preserved'} | ConvertTo-Json -Depth 5
} finally {
    if (Test-Path -LiteralPath $testStateRoot) { Remove-Item -LiteralPath $testStateRoot -Recurse -Force }
}
