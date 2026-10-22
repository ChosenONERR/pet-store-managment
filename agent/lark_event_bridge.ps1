param(
    [string]$ProjectRoot,
    [ValidateSet('Live', 'Initialize', 'Inspect', 'Replay', 'FlushAcks', 'RetryPending')]
    [string]$Mode = 'Live',
    [int]$ReplayHours = 24,
    [string]$ChatName = '宠速星经营管家',
    [string]$RetryChangeId = ''
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$petstarCli = Join-Path $ProjectRoot 'agent\petstar_cli.py'
$financeEventScript = Join-Path $ProjectRoot 'agent\record_finance_event.ps1'
$codexBridge = Join-Path $ProjectRoot 'agent\codex_chat_bridge.py'
$ledgerScript = Join-Path $ProjectRoot 'agent\sync_daily_ledger.ps1'
$costSnapshotScript = Join-Path $ProjectRoot 'agent\rebuild_cost_snapshots.ps1'
$stateRoot = Join-Path $ProjectRoot 'agent\state'
$runtimeLog = Join-Path $stateRoot 'lark_event_bridge.log'
$consumerLog = Join-Path $stateRoot 'lark_event_consumer.log'
$handledMessagesLog = Join-Path $stateRoot 'handled_inbound_messages.ndjson'
$pendingConfirmationsPath = Join-Path $stateRoot 'pending_confirmations.json'
$pendingAcknowledgementsPath = Join-Path $stateRoot 'pending_acknowledgements.json'
$interactivePriorityPath = Join-Path $stateRoot 'interactive_priority.json'
$script:codexFirstRouting = $true
$larkCli = Join-Path $ProjectRoot 'agent\lark_cli_noexit.ps1'
. (Join-Path $ProjectRoot 'agent\finance_transaction.ps1')
. (Join-Path $ProjectRoot 'agent\runtime_paths.ps1')
$pythonExe = Get-PetstarPython
# npm 原生 lark-cli.ps1 末尾会执行 exit，长驻监听器直接调用时会结束父进程。
# 本地薄包装调用同一个官方 Node 入口，不会结束父进程，同时保留多行文本参数。
$powershellExe = Get-PetstarWindowsPowerShell
if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
if (-not (Test-Path -LiteralPath $larkCli)) { throw "找不到飞书 CLI：$larkCli" }

function Get-FinanceWorkbookTokens {
    $tokens = @{}
    $sourceMetaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
    $ledgerMetaPath = Join-Path $ProjectRoot 'work\daily_ledger_meta.json'
    if (Test-Path -LiteralPath $sourceMetaPath) {
        $sourceMeta = [System.IO.File]::ReadAllText($sourceMetaPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $tokens['source'] = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
    }
    if (Test-Path -LiteralPath $ledgerMetaPath) {
        $ledgerMeta = [System.IO.File]::ReadAllText($ledgerMetaPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $tokens['ledger'] = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
    }
    return $tokens
}

function Write-BridgeLog([string]$Message) {
    Add-Content -LiteralPath $runtimeLog -Value "$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')) $Message" -Encoding UTF8
}

function Set-InteractivePriority([string]$ChangeId) {
    $payload = [ordered]@{change_id=$ChangeId;process_id=$PID;started_at=[datetime]::Now.ToString('o')} | ConvertTo-Json -Compress
    $temporaryPath = "$interactivePriorityPath.$([guid]::NewGuid().ToString('N')).tmp"
    [IO.File]::WriteAllText($temporaryPath,$payload,[Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporaryPath -Destination $interactivePriorityPath -Force
}

function Clear-InteractivePriority([string]$ChangeId) {
    try {
        if (-not (Test-Path -LiteralPath $interactivePriorityPath)) { return }
        $state = Get-Content -LiteralPath $interactivePriorityPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]$state.change_id -eq $ChangeId -and [int]$state.process_id -eq $PID) {
            Remove-Item -LiteralPath $interactivePriorityPath -Force
        }
    } catch {
        Write-BridgeLog "interactive priority cleanup failed change_id=$ChangeId"
    }
}

function Start-PendingConfirmationRetryWorker([string]$ChangeId) {
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath),
        '-ProjectRoot', ('"{0}"' -f $ProjectRoot), '-Mode', 'RetryPending',
        '-RetryChangeId', $ChangeId, '-ChatName', ('"{0}"' -f $ChatName)
    )
    $process = Start-Process -FilePath $powershellExe -ArgumentList $arguments -WorkingDirectory $ProjectRoot -WindowStyle Hidden -PassThru
    Write-BridgeLog "pending confirmation retry worker launched change_id=$ChangeId pid=$($process.Id)"
    return $process.Id
}

function Get-MessageHash([string]$MessageId) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString(
            $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($MessageId))
        ).Replace('-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Get-CurrentUserOpenId {
    $statusRaw = & $larkCli auth status --json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '无法读取飞书授权状态' }
    $status = $statusRaw | ConvertFrom-Json
    $openId = $status.identities.user.openId
    if (-not $openId) { $openId = $status.identities.user.open_id }
    if (-not $openId) { throw '飞书授权中缺少当前用户标识' }
    if ($status.identities.bot.status -ne 'ready') { throw '飞书机器人授权未就绪' }
    return [string]$openId
}

function Convert-CommandOutputToJson([string[]]$OutputLines) {
    if ($null -eq $OutputLines -or $OutputLines.Count -eq 0) { return $null }
    $nonEmptyLines = @($OutputLines | Where-Object { $_ -and $_.Trim() })
    if ($nonEmptyLines.Count -eq 0) { return $null }

    $wholeOutput = $nonEmptyLines -join "`n"
    try { return ($wholeOutput | ConvertFrom-Json) } catch {}

    for ($start = 0; $start -lt $nonEmptyLines.Count; $start++) {
        if (-not $nonEmptyLines[$start].TrimStart().StartsWith('{')) { continue }
        for ($end = $nonEmptyLines.Count - 1; $end -ge $start; $end--) {
            if (-not $nonEmptyLines[$end].TrimEnd().EndsWith('}')) { continue }
            $candidate = $nonEmptyLines[$start..$end] -join "`n"
            try { return ($candidate | ConvertFrom-Json) } catch {}
        }
    }
    return $null
}

function Get-CommandFailureDetail($Result, [string[]]$OutputLines) {
    if ($null -ne $Result) {
        if ($Result.error -is [string] -and ([string]$Result.error).Trim()) { return ([string]$Result.error).Trim() }
        if ($Result.error -and $Result.error.message) { return ([string]$Result.error.message).Trim() }
        if ($Result.message) { return ([string]$Result.message).Trim() }
    }
    $meaningfulLines = @($OutputLines | Where-Object {
        $trimmed = $_.Trim()
        $trimmed -and $trimmed -notin @('{', '}', '[', ']')
    })
    $explicitError = @($meaningfulLines | Where-Object { $_ -match '(?i)(?:^|:\s*)error:|错误|失败' } | Select-Object -Last 1)
    if ($explicitError.Count -gt 0) { return ([string]$explicitError[0]).Trim() }
    if ($meaningfulLines.Count -gt 0) { return ([string]$meaningfulLines[-1]).Trim() }
    return '执行层没有返回具体原因'
}

function New-AcknowledgementIdempotencyKey([string]$Purpose, [string]$MessageId) {
    $safePurpose = if ($Purpose -match '^[a-z0-9-]+$') { $Purpose } else { 'ack' }
    $prefix = "petstar-$safePurpose-"
    if ($prefix.Length -ge 50) { $prefix = $prefix.Substring(0,49) + '-' }
    $hash = Get-MessageHash $MessageId
    $hashLength = [Math]::Min($hash.Length, 50 - $prefix.Length)
    return $prefix + $hash.Substring(0, $hashLength)
}

function Send-AcknowledgementNow([string]$OpenId, [string]$Text, [string]$IdempotencyKey) {
    # Windows PowerShell 5.1 may split native arguments when a value contains
    # literal JSON quotes. Curly quotes keep the user-visible meaning while
    # guaranteeing that the whole acknowledgement remains one --text value.
    $transportText = $Text.Replace('"', '“')
    $raw = & $larkCli im +messages-send --as bot --user-id $OpenId --text $transportText --idempotency-key $idempotencyKey --format json 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "确认消息发送失败：$($raw.Trim())" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "确认消息发送失败：$($result.error.message)" }
}

function Invoke-WithAcknowledgementLock([scriptblock]$Action) {
    $ackMutex = [Threading.Mutex]::new($false, 'Local\PetstarLarkAcknowledgementOutbox')
    $ackLockTaken = $false
    try {
        try { $ackLockTaken = $ackMutex.WaitOne([timespan]::FromSeconds(15)) }
        catch [Threading.AbandonedMutexException] { $ackLockTaken = $true }
        if (-not $ackLockTaken) { throw '飞书回执队列正忙，未能在15秒内取得锁' }
        return & $Action
    } finally {
        if ($ackLockTaken) { try { $ackMutex.ReleaseMutex() } catch {} }
        $ackMutex.Dispose()
    }
}

function Read-PendingAcknowledgementsUnsafe {
    if (-not (Test-Path -LiteralPath $pendingAcknowledgementsPath)) { return @() }
    try { return @((Get-Content -LiteralPath $pendingAcknowledgementsPath -Raw -Encoding UTF8 | ConvertFrom-Json) | Where-Object { $null -ne $_ }) }
    catch {
        $backupPath = "$pendingAcknowledgementsPath.corrupt.$([datetime]::Now.ToString('yyyyMMdd_HHmmss')).json"
        try { Copy-Item -LiteralPath $pendingAcknowledgementsPath -Destination $backupPath -Force } catch {}
        Write-BridgeLog 'acknowledgement outbox could not be read; copied invalid file aside before recovery'
        return @()
    }
}

function Save-PendingAcknowledgementsUnsafe([object[]]$Entries) {
    $payload = @($Entries | Where-Object { $null -ne $_ }) | ConvertTo-Json -Depth 12
    if (-not $payload) { $payload = '[]' }
    $tempPath = "$pendingAcknowledgementsPath.$([guid]::NewGuid().ToString('N')).tmp"
    [IO.File]::WriteAllText($tempPath,$payload,[Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tempPath -Destination $pendingAcknowledgementsPath -Force
}

function Add-PendingAcknowledgement($Entry) {
    Invoke-WithAcknowledgementLock {
        $entries = @(Read-PendingAcknowledgementsUnsafe)
        if (-not @($entries | Where-Object { [string]$_.idempotency_key -eq [string]$Entry.idempotency_key }).Count) {
            $entries += $Entry
            Save-PendingAcknowledgementsUnsafe $entries
        }
    }
}

function Remove-PendingAcknowledgement([string]$IdempotencyKey) {
    Invoke-WithAcknowledgementLock {
        $entries = @(Read-PendingAcknowledgementsUnsafe | Where-Object { [string]$_.idempotency_key -ne $IdempotencyKey })
        Save-PendingAcknowledgementsUnsafe $entries
    }
}

function Flush-PendingAcknowledgements {
    $entries = Invoke-WithAcknowledgementLock { return @(Read-PendingAcknowledgementsUnsafe) }
    $sent = 0
    $failed = 0
    foreach ($entry in @($entries)) {
        try {
            Send-AcknowledgementNow ([string]$entry.open_id) ([string]$entry.text) ([string]$entry.idempotency_key)
            Remove-PendingAcknowledgement ([string]$entry.idempotency_key)
            $sent++
        } catch {
            $failed++
            Write-BridgeLog "acknowledgement retry failed key=$([string]$entry.idempotency_key) error=$($_.Exception.Message)"
        }
    }
    return [pscustomobject]@{queued=@($entries).Count;sent=$sent;failed=$failed}
}

function Send-Acknowledgement([string]$OpenId, [string]$Text, [string]$MessageId, [string]$Purpose = 'ack') {
    if ($script:previewMode) {
        $script:previewAcknowledgement = $Text
        return [pscustomobject]@{sent=$false;preview=$true;queued=$false}
    }
    if ($Text.Length -gt 3500) { $Text = $Text.Substring(0,3460) + "`n`n（内容较长，已截断；详细数据请在表格查看。）" }
    $idempotencyKey = New-AcknowledgementIdempotencyKey $Purpose $MessageId
    $entry = [pscustomobject]@{
        idempotency_key = $idempotencyKey
        open_id = $OpenId
        text = $Text
        purpose = $Purpose
        created_at = [datetime]::Now.ToString('o')
        attempts = 0
    }
    Add-PendingAcknowledgement $entry
    try {
        Send-AcknowledgementNow $OpenId $Text $idempotencyKey
        Remove-PendingAcknowledgement $idempotencyKey
        return [pscustomobject]@{sent=$true;preview=$false;queued=$false}
    } catch {
        Write-BridgeLog "acknowledgement queued key=$idempotencyKey error=$($_.Exception.Message)"
        return [pscustomobject]@{sent=$false;preview=$false;queued=$true;error=$_.Exception.Message}
    }
}

function Add-ProcessingReaction([string]$MessageId) {
    if (-not $MessageId) { return $null }
    try {
        $payload = '{"reaction_type":{"emoji_type":"Typing"}}'
        # lark-cli.cmd 会吞掉内联 JSON 的双引号；通过 stdin 传入，避免 payload 被 cmd.exe 改写。
        $raw = $payload | & $larkCli im reactions create --as bot --message-id $MessageId --data - --format json 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            Write-BridgeLog "processing reaction add failed message_hash=$((Get-MessageHash $MessageId).Substring(0,12))"
            return $null
        }
        $response = $raw | ConvertFrom-Json
        if ($response.ok -eq $false) {
            Write-BridgeLog "processing reaction add rejected message_hash=$((Get-MessageHash $MessageId).Substring(0,12))"
            return $null
        }
        $reactionId = if ($response.data.reaction_id) { [string]$response.data.reaction_id } else { [string]$response.reaction_id }
        if (-not $reactionId) {
            Write-BridgeLog "processing reaction add returned no id message_hash=$((Get-MessageHash $MessageId).Substring(0,12))"
            return $null
        }
        return $reactionId
    } catch {
        Write-BridgeLog "processing reaction add error message_hash=$((Get-MessageHash $MessageId).Substring(0,12))"
        return $null
    }
}

function Remove-ProcessingReaction([string]$MessageId, [string]$ReactionId) {
    if (-not $MessageId -or -not $ReactionId) { return }
    try {
        $raw = & $larkCli im reactions delete --as bot --message-id $MessageId --reaction-id $ReactionId --format json 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            Write-BridgeLog "processing reaction remove failed message_hash=$((Get-MessageHash $MessageId).Substring(0,12))"
            return
        }
        $response = $raw | ConvertFrom-Json
        if ($response.ok -eq $false) {
            Write-BridgeLog "processing reaction remove rejected message_hash=$((Get-MessageHash $MessageId).Substring(0,12))"
        }
    } catch {
        Write-BridgeLog "processing reaction remove error message_hash=$((Get-MessageHash $MessageId).Substring(0,12))"
    }
}

function Invoke-PetstarCodex([string]$Text, [string]$MessageId, [string]$Mode = 'normal') {
    $startedAt = [datetime]::UtcNow
    $encoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text))
    $output = & $pythonExe $codexBridge --prompt-base64 $encoded --event-key $MessageId --mode $Mode | Out-String
    try {
        $result = $output | ConvertFrom-Json
    } catch {
        throw '本机 Codex 桥接返回了无法解析的结果'
    }
    $timings = $result.PSObject.Properties['_timings']
    if ($null -ne $timings) {
        $t = $timings.Value
        Write-BridgeLog "codex timing mode=$Mode total=$($t.total_seconds)s model=$($t.model_seconds)s query=$($t.query_seconds)s lock=$($t.lock_wait_seconds)s queries=$($t.query_count) cache_hits=$($t.query_cache_hits)"
    } else {
        Write-BridgeLog "codex timing mode=$Mode total=$([math]::Round(([datetime]::UtcNow-$startedAt).TotalSeconds,3))s"
    }
    return $result
}

function Clear-PetstarCodexQueryCache {
    try {
        $null = & $pythonExe $codexBridge --invalidate-cache 2>$null | Out-String
    } catch {
        Write-BridgeLog 'codex query cache invalidation failed; cache expires automatically within 30 seconds'
    }
}

function Resolve-MessageDate([string]$Text, [string]$CreateTime) {
    $baseDate = [datetime]::Today
    $milliseconds = 0L
    if ([long]::TryParse($CreateTime, [ref]$milliseconds) -and $milliseconds -gt 0) {
        $baseDate = [DateTimeOffset]::FromUnixTimeMilliseconds($milliseconds).LocalDateTime.Date
    }
    $explicit = [regex]::Match($Text, '(20\d{2})[-/.年](\d{1,2})[-/.月](\d{1,2})日?')
    if ($explicit.Success) {
        return [datetime]::new([int]$explicit.Groups[1].Value, [int]$explicit.Groups[2].Value, [int]$explicit.Groups[3].Value)
    }
    if ($Text -match '昨天|昨日') { return $baseDate.AddDays(-1) }
    if ($Text -match '前天') { return $baseDate.AddDays(-2) }
    return $baseDate
}

function Get-ExplicitExpiryDate([string]$Text) {
    $match = [regex]::Match($Text, '(20\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?')
    if (-not $match.Success) { return $null }
    try {
        return [datetime]::new([int]$match.Groups[1].Value, [int]$match.Groups[2].Value, [int]$match.Groups[3].Value)
    } catch {
        throw "过期日期无效：$($match.Value)"
    }
}

function Get-BatchDateFacts([string]$Text) {
    $skuCodes = @(Get-SkuCodes $Text)
    $productionMatch = [regex]::Match($Text, '(?:生产日期|生产日)\s*(?:是|为|：|:)?\s*(20\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?')
    $shelfMatch = [regex]::Match($Text, '保质期\s*(?:是|为|：|:)?\s*(\d+(?:\.\d+)?)\s*(天|个?月|年)')
    $productionDate = $null
    if ($productionMatch.Success) {
        try {
            $productionDate = [datetime]::new([int]$productionMatch.Groups[1].Value,[int]$productionMatch.Groups[2].Value,[int]$productionMatch.Groups[3].Value)
        } catch {
            throw "生产日期无效：$($productionMatch.Value)"
        }
    }
    $shelfValue = $null
    $shelfUnit = $null
    if ($shelfMatch.Success) {
        $shelfValue = [double]::Parse($shelfMatch.Groups[1].Value,[Globalization.CultureInfo]::InvariantCulture)
        $shelfUnit = if ($shelfMatch.Groups[2].Value -match '月') { '月' } else { [string]$shelfMatch.Groups[2].Value }
    }
    $expiryDate = $null
    if ($productionDate -and $null -ne $shelfValue) {
        switch ($shelfUnit) {
            '天' { $expiryDate = $productionDate.AddDays($shelfValue) }
            '月' { if ($shelfValue -ne [math]::Round($shelfValue)) { throw '按月填写保质期时必须是整数' }; $expiryDate = $productionDate.AddMonths([int]$shelfValue) }
            '年' { if ($shelfValue -ne [math]::Round($shelfValue)) { throw '按年填写保质期时必须是整数' }; $expiryDate = $productionDate.AddYears([int]$shelfValue) }
        }
    }
    return [pscustomobject]@{
        complete = $skuCodes.Count -eq 1 -and $productionDate -and $null -ne $shelfValue
        goods_code = $(if ($skuCodes.Count -eq 1) { $skuCodes[0] } else { $null })
        production_date = $productionDate
        shelf_life_value = $shelfValue
        shelf_life_unit = $shelfUnit
        expiry_date = $expiryDate
    }
}

function Get-SkuCodes([string]$Text) {
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $codes = [Collections.Generic.List[string]]::new()
    foreach ($match in @([regex]::Matches($Text, '(?i)(?<![A-Za-z0-9])(?:kw|csx)\s*\d+(?!\d)'))) {
        $code = ([string]$match.Value).Replace(' ', '').ToLowerInvariant()
        if ($seen.Add($code)) { $codes.Add($code) }
    }
    return $codes.ToArray()
}

function Get-SkuCodeChangeFacts([string]$Text) {
    $match = [regex]::Match($Text, '(?i)(?:将|把)?\s*((?:kw|csx)\s*\d+)(?:\s*的?\s*(?:kw\s*号|货号))?\s*[，,、]?\s*(?:修改|更改|改|调整|更换|换|登记|设置|设定|绑定|对应)(?:成|为|到)?\s*(kw\s*\d+)')
    if (-not $match.Success) { return $null }
    $oldCode = ([string]$match.Groups[1].Value).Replace(' ','').ToLowerInvariant()
    $newCode = ([string]$match.Groups[2].Value).Replace(' ','').ToLowerInvariant()
    return [pscustomobject]@{old_code=$oldCode;new_code=$newCode}
}

function Test-ClearExpiry([string]$Text) {
    return [bool]($Text -match '(?:清空|删除|去掉|取消|没有|无).{0,8}(?:最早)?(?:过期|到期|有效期)|(?:最早)?(?:过期|到期|有效期).{0,8}(?:清空|删除|去掉|取消|没有|无)')
}

function Format-NonExecutableCodexReply([string]$Text) {
    $reply = [string]$Text
    if ($reply -match '^\s*(?:📝\s*)?待确认(?:\s*[｜|:]\s*[^\r\n]*)?') {
        $body = ([regex]::Replace($reply, '^\s*(?:📝\s*)?待确认(?:\s*[｜|:]\s*)?', '')).Trim()
        if ($body) { return "没有生成可执行修改，也没有改表。`n$body" }
        return '没有生成可执行修改，也没有改表。请补充要改的对象和新值。'
    }
    return $reply
}

function Invoke-CodexAndReply($Event, [string]$AuthorizedOpenId, [string]$MessageText) {
    if ($Event.chat_type -ne 'p2p') { return $false }
    $codexMode = if ($script:previewMode) { 'preview' } elseif ($script:confirmedExecution) { 'execute' } else { 'normal' }
    $codexResult = Invoke-PetstarCodex $MessageText ([string]$Event.message_id) $codexMode
    if (-not $codexResult.ok) {
        $errorReply = '宠速星经营管家暂时无法处理这条消息。' + [Environment]::NewLine + [string]$codexResult.error
        Send-Acknowledgement $AuthorizedOpenId $errorReply ([string]$Event.message_id) 'error'
        Write-BridgeLog "codex failed error=$([string]$codexResult.error)"
        $script:lastProcessFailure = [string]$codexResult.error
        return $false
    }

    $actions = @($codexResult.actions)
    $mutatingActions = @($actions | Where-Object { [string]$_.action -notin @('RefreshLedger','SendDailySummary','SyncPlatformData') })

    # 如果正则快速路径没有识别为写入，但 Codex 识别出了写表意图，仍必须自动进入两阶段确认。
    if ($codexMode -eq 'normal' -and $mutatingActions.Count -gt 0) {
        # Codex already produced structured actions from current facts. Reuse them
        # for deterministic dry-run validation instead of asking the model to parse
        # the same message a second time.
        $draft = New-ConfirmationDraftFromCodexResult $Event $MessageText $codexResult
        if ([int]$draft.mutation_count -eq 0) {
            $reply = if ($draft.needs_information) { [string]$draft.preview } else { '我理解到你想修改数据，但这次没有生成明确的修改项，所以没有写表。请补充要改的对象和新值。' }
            Send-Acknowledgement $AuthorizedOpenId $reply ([string]$Event.message_id) 'clarify'
            return $true
        }
        $queued = Add-PendingConfirmation $draft
        if ($queued.duplicate) {
            $existing = $queued.existing
            $queueText = "相同修改已经在待确认队列中，没有重复新增。`n`n$([string]$existing.preview)"
            Send-Acknowledgement $AuthorizedOpenId $queueText ([string]$Event.message_id) 'preview-duplicate'
            Write-BridgeLog "codex mutation deduplicated requested_change_id=$($draft.change_id) existing_change_id=$($existing.change_id)"
        } else {
            $queueText = Format-QueuedConfirmationReply $draft $queued.position $queued.active
            Send-Acknowledgement $AuthorizedOpenId $queueText ([string]$Event.message_id) 'preview'
            Write-BridgeLog "codex promoted mutation to confirmation change_id=$($draft.change_id)"
        }
        return $true
    }

    $executedOperations = [Collections.Generic.List[object]]::new()
    foreach ($operation in @(Invoke-CodexActionsInOrder $actions $Event $codexMode)) { $executedOperations.Add($operation) }

    if ($script:previewMode) {
        # 只为真正的数据修改生成确认预览。RefreshLedger 等维护动作无需确认，
        # 在预览阶段只保留 Codex 的说明，随后由正常路径直接执行。
        $previewMutations = @($executedOperations | Where-Object { [string]$_.action -notin @('RefreshLedger','SendDailySummary','SyncPlatformData') })
        $script:previewNeedsInformation = [bool]$codexResult.needs_information -or $executedOperations.Count -eq 0
        $reply = if ($previewMutations.Count -gt 0) {
            Format-ConfirmationPreview '' $previewMutations ([string]$codexResult.reply)
        } else {
            Format-NonExecutableCodexReply ([string]$codexResult.reply)
        }
        Send-Acknowledgement $AuthorizedOpenId $reply ([string]$Event.message_id)
    } elseif ($executedOperations.Count -gt 0) {
        Send-Acknowledgement $AuthorizedOpenId (Format-CodexExecutionReply @($executedOperations)) ([string]$Event.message_id)
    } else {
        Send-Acknowledgement $AuthorizedOpenId (Format-NonExecutableCodexReply ([string]$codexResult.reply)) ([string]$Event.message_id)
    }
    Write-BridgeLog "codex processed message_hash=$((Get-MessageHash ([string]$Event.message_id)).Substring(0,12)) cached=$($codexResult.cached) actions=$($actions.Count)"
    return $true
}

function Convert-PetstarArgumentsToMap([string[]]$Arguments) {
    $map = [ordered]@{}
    for ($index = 0; $index -lt $Arguments.Count; $index++) {
        $name = [string]$Arguments[$index]
        if (-not $name.StartsWith('--')) { continue }
        $key = $name.Substring(2).Replace('-','_')
        if ($index + 1 -lt $Arguments.Count -and -not ([string]$Arguments[$index + 1]).StartsWith('--')) {
            $map[$key] = [string]$Arguments[$index + 1]
            $index++
        } else {
            $map[$key] = $true
        }
    }
    return [pscustomobject]$map
}

function Convert-PetstarReceiptStatus([string]$Value) {
    $text = $Value.Trim()
    if ($text -in @('已到账','当天到账','当日到账','已经到账','已收款','当天收款')) { return '已到账' }
    if ($text -in @('未到账','尚未到账','待到账','未收款','赊账')) { return '未到账' }
    return $text
}

function Convert-PetstarPlatform([string]$Value) {
    $text = $Value.Trim()
    switch ($text) {
        '美团外卖' { return '美团' }
        '京东外卖' { return '京东到家' }
        '京东秒送' { return '京东到家' }
        '饿了么' { return '淘宝闪购' }
        default { return $text }
    }
}

function Invoke-PetstarEvent([string[]]$Arguments) {
    $argumentMap = Convert-PetstarArgumentsToMap $Arguments
    $action = [string]$argumentMap.action
    if ($script:previewMode) {
        $previewResult = $null
        if ($action -in @('RemoveIntervention','OfflineSale','AddOrderGift','Replenishment','SetSkuExpiry','SetSkuBatchDate','SetSkuCost','ChangeSkuCode','SetOrderDeliveryFee','SetOrderActuals','SetRefundDisposition')) {
            $previewOutputLines = @(& $pythonExe $petstarCli event @Arguments --dry-run 2>&1 | ForEach-Object { [string]$_ })
            $previewExitCode = $LASTEXITCODE
            $previewResult = Convert-CommandOutputToJson $previewOutputLines
            if ($previewExitCode -ne 0 -or $null -eq $previewResult -or -not $previewResult.ok) {
                $previewError = Get-CommandFailureDetail $previewResult $previewOutputLines
                throw "经营事件预览失败：$previewError"
            }
        } else {
            $amount = if ($argumentMap.amount) { [double]::Parse([string]$argumentMap.amount,[Globalization.CultureInfo]::InvariantCulture) } else { 0.0 }
            switch ($action) {
                'AddIntervention' { $previewResult = [pscustomobject]@{ok=$true;result=[pscustomobject]@{order_no=$argumentMap.order_no;person=$argumentMap.person}} }
                'RestoreInventory' { $previewResult = [pscustomobject]@{ok=$true;result=[pscustomobject]@{order_no=$argumentMap.order_no;state='已恢复'}} }
                'ConfirmSelfDelivery' { $previewResult = [pscustomobject]@{ok=$true;result=[pscustomobject]@{order_no=$argumentMap.order_no;delivery_type='自配送';delivery_fee=0;total_cost=0}} }
                'SetExternalDeliveryFee' { $previewResult = [pscustomobject]@{ok=$true;result=[pscustomobject]@{order_no=$argumentMap.order_no;delivery_type='骑手配送';delivery_fee=$amount;total_cost=$amount}} }
                default { $previewResult = [pscustomobject]@{ok=$true;result=[pscustomobject]@{}} }
            }
        }
        $script:previewOperations.Add([pscustomobject]@{action=$action;arguments=$argumentMap;result=$previewResult.result})
        return $previewResult
    }
    $outputLines = @(& $pythonExe $petstarCli event @Arguments 2>&1 | ForEach-Object { [string]$_ })
    $exitCode = $LASTEXITCODE
    $result = Convert-CommandOutputToJson $outputLines
    if ($exitCode -ne 0 -or $null -eq $result -or -not $result.ok) {
        $detail = Get-CommandFailureDetail $result $outputLines
        throw "经营事件执行失败：$detail"
    }
    return $result
}

function Test-BatchCompatibleSkuOperation($Operation) {
    $action = [string]$Operation.action
    if ($action -eq 'Replenishment') { return $true }
    if ($action -ne 'SetSkuCost') { return $false }
    return -not (Get-CodexActionValue $Operation.arguments 'order_no')
}

function Invoke-PetstarBatchSkuMaintenance($StoredOperations, $Event, [int]$StartIndex, [switch]$Preview) {
    $operationList = @($StoredOperations)
    if ($operationList.Count -lt 2) { throw '批量 SKU 维护至少需要两项兼容操作' }
    $payloadOperations = [Collections.Generic.List[object]]::new()
    $normalizedArguments = [Collections.Generic.List[object]]::new()
    for ($offset = 0; $offset -lt $operationList.Count; $offset++) {
        $stored = $operationList[$offset]
        $arguments = [ordered]@{}
        foreach ($property in @($stored.arguments.PSObject.Properties)) { $arguments[$property.Name] = $property.Value }
        if ([string]$stored.action -eq 'Replenishment' -and -not $arguments['date']) {
            $arguments['date'] = (Resolve-MessageDate ([string]$Event.content) ([string]$Event.create_time)).ToString('yyyy-MM-dd')
        }
        if (-not $arguments['note']) { $arguments['note'] = '本机 Codex 通过飞书自然语言解析' }
        $arguments['event_id'] = "codex-$($Event.message_id)-$($StartIndex+$offset)-$([string]$stored.action)"
        $payloadOperations.Add([ordered]@{action=[string]$stored.action;arguments=$arguments})
        $normalizedArguments.Add([pscustomobject]$arguments)
    }
    $requestPath = Join-Path $stateRoot ("batch_sku_maintenance_{0}.json" -f [guid]::NewGuid().ToString('N'))
    $batchEventId = "codex-$($Event.message_id)-batch-sku-$StartIndex-$($operationList.Count)"
    try {
        [IO.File]::WriteAllText($requestPath,(ConvertTo-Json -InputObject $payloadOperations.ToArray() -Depth 20 -Compress),[Text.UTF8Encoding]::new($false))
        $commandArguments = @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$financeEventScript,'-Action','BatchSkuMaintenance','-OperationsPath',$requestPath,'-EventId',$batchEventId,'-ProjectRoot',$ProjectRoot)
        if ($Preview) { $commandArguments += '-DryRun' }
        $outputLines = @(& $powershellExe @commandArguments 2>&1 | ForEach-Object { [string]$_ })
        $exitCode = $LASTEXITCODE
        $result = Convert-CommandOutputToJson $outputLines
        if ($exitCode -ne 0 -or $null -eq $result -or -not $result.ok) {
            $detail = Get-CommandFailureDetail $result $outputLines
            throw "批量经营事件执行失败：$detail"
        }
        $executed = [Collections.Generic.List[object]]::new()
        if ([bool]$result.duplicate) {
            foreach ($stored in $operationList) {
                $executed.Add([pscustomobject]@{action=[string]$stored.action;arguments=$stored.arguments;result=$null;formula_verification='success';duplicate=$true})
            }
            return $executed.ToArray()
        }
        $details = @($result.result.operations)
        if ($details.Count -ne $operationList.Count) { throw "批量经营事件结果不完整：计划 $($operationList.Count) 项，实际 $($details.Count) 项" }
        for ($offset = 0; $offset -lt $operationList.Count; $offset++) {
            $stored = $operationList[$offset]
            $executed.Add([pscustomobject]@{action=[string]$stored.action;arguments=$normalizedArguments[$offset];result=$details[$offset];formula_verification=$result.formula_verification;duplicate=$false})
        }
        return $executed.ToArray()
    } finally {
        if (Test-Path -LiteralPath $requestPath) { Remove-Item -LiteralPath $requestPath -Force }
    }
}

function Invoke-DailyLedgerSync([datetime]$ThroughDate) {
    $output = & $powershellExe -NoProfile -ExecutionPolicy Bypass -File $ledgerScript -ProjectRoot $ProjectRoot -ThroughDate $ThroughDate.Date | Out-String
    try { $result = $output | ConvertFrom-Json } catch { throw '每日经营简账同步返回了无法解析的结果' }
    if (-not $result.ok -or $result.formula_verification -ne 'success') { throw '每日经营简账同步或公式校验失败' }
    return $result
}

function Invoke-CostSnapshotRebuild {
    if (-not (Test-Path -LiteralPath $costSnapshotScript)) { throw '缺少成本快照刷新脚本' }
    $outputLines = @(& $powershellExe -NoProfile -ExecutionPolicy Bypass -File $costSnapshotScript -ProjectRoot $ProjectRoot 2>&1 | ForEach-Object { [string]$_ })
    $exitCode = $LASTEXITCODE
    $result = Convert-CommandOutputToJson $outputLines
    if ($exitCode -ne 0 -or $null -eq $result -or -not $result.ok) {
        $detail = Get-CommandFailureDetail $result $outputLines
        throw "成本快照刷新失败：$detail"
    }
    return $result
}

function Get-CodexActionValue($Arguments, [string]$Name) {
    if ($null -eq $Arguments) { return $null }
    $property = $Arguments.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $null }
    if ($property.Value -is [string] -and -not ([string]$property.Value).Trim()) { return $null }
    return $property.Value
}

function Invoke-PetstarSyncAction([string]$Action, $Arguments, $Event) {
    $dateValue = Get-CodexActionValue $Arguments 'date'
    $resolvedDate = if ($dateValue) { [datetime]::Parse([string]$dateValue) } else { Resolve-MessageDate ([string]$Event.content) ([string]$Event.create_time) }
    if ($script:previewMode) {
        return [pscustomobject]@{ok=$true;result=[pscustomobject]@{date=$resolvedDate.ToString('yyyy-MM-dd');status='待执行'}}
    }
    if ($Action -eq 'RefreshLedger') {
        # 用户修改 SKU 成本后，必须先把新成本回写到订单商品/订单汇总，
        # 再刷新月份收入表；只跑简账会继续沿用旧订单成本。
        $costResult = Invoke-CostSnapshotRebuild
        $ledgerResult = Invoke-DailyLedgerSync $resolvedDate
        $ledgerResult | Add-Member -NotePropertyName cost_snapshot_rebuild -NotePropertyValue $costResult -Force
        # 统一成其他动作的 { ok, result } 结构，避免上层错误读取不存在的 .result 而丢失重算明细。
        return [pscustomobject]@{ok=$true;result=$ledgerResult}
    }

    $syncArguments = @('sync','--date',$resolvedDate.ToString('yyyy-MM-dd'))
    if ($Action -eq 'SendDailySummary') {
        $syncArguments += @('--mode','night','--skip-sync','--force-resend')
    } else {
        $syncArguments += @('--mode','manual','--silent')
    }
    $output = & $pythonExe $petstarCli @syncArguments | Out-String
    try { $result = $output | ConvertFrom-Json } catch { throw '经营同步返回了无法解析的结果' }
    if (-not $result.ok) { throw '经营同步没有成功完成' }
    return $result
}

function Invoke-CodexAction($CodexAction, $Event, [int]$Index, [string]$Mode) {
    $actionName = [string]$CodexAction.action
    $arguments = $CodexAction.arguments
    $allowed = @(
        'AddIntervention','RemoveIntervention','RestoreInventory','ConfirmSelfDelivery','SetExternalDeliveryFee','SetOrderDeliveryFee','SetOrderActuals',
        'OfflineSale','AddOrderGift','Replenishment','CashExpense','OperatingExpense','OwnerWithdrawal',
        'OwnerLabor','MarketRent','BeginningCash','ActualReceipt','AssetAmortization','SetSkuExpiry',
        'SetSkuBatchDate','SetSkuCost','ChangeSkuCode','SetRefundDisposition','RefreshLedger','SendDailySummary','SyncPlatformData'
    )
    if ($actionName -notin $allowed) { throw "Codex 返回了不允许的经营动作：$actionName" }

    if ($actionName -in @('RefreshLedger','SendDailySummary','SyncPlatformData')) {
        $result = Invoke-PetstarSyncAction $actionName $arguments $Event
        return [pscustomobject]@{action=$actionName;arguments=$arguments;result=$result.result}
    }

    $eventArguments = @('--action',$actionName)
    $propertyMap = [ordered]@{
        date='date'; order_no='order-no'; daily_order_no='daily-order-no'; person='person'; items='items';
        goods_code='goods-code'; transition_code='transition-code'; new_goods_code='new-goods-code'; expiry_date='expiry-date';
        production_date='production-date'; shelf_life_value='shelf-life-value'; shelf_life_unit='shelf-life-unit';
        quantity='quantity'; purchase_amount='purchase-amount'; payment_amount='payment-amount'; payment_date='payment-date';
        amount='amount'; delivery_fee='delivery-fee'; receipt_status='receipt-status'; platform='platform'; refund_disposition='refund-disposition'; note='note'
    }
    foreach ($entry in $propertyMap.GetEnumerator()) {
        $value = Get-CodexActionValue $arguments $entry.Key
        if ($actionName -eq 'SetOrderActuals' -and $entry.Key -eq 'delivery_fee' -and
            ([string]$Event.content) -notmatch '配送费|运费|骑手|自配送|自己送|商家自配|自提|到店(?:取|拿)|没有.{0,6}配送') {
            # 用户没有谈配送时，不能把模型查询到的旧值固化进待确认项。
            # 执行器会在确认时读取该订单的当前配送费，避免等待期间同步更新后被旧值覆盖。
            $value = $null
        }
        if ($entry.Key -eq 'shelf_life_unit' -and $null -ne $value) {
            $unitText = ([string]$value).Trim()
            $value = if ($unitText -match '月') { '月' } elseif ($unitText -match '年') { '年' } elseif ($unitText -match '天|日') { '天' } else { $unitText }
        }
        if ($entry.Key -eq 'receipt_status' -and $null -ne $value) { $value = Convert-PetstarReceiptStatus ([string]$value) }
        if ($entry.Key -eq 'platform' -and $null -ne $value) { $value = Convert-PetstarPlatform ([string]$value) }
        if ($null -ne $value) { $eventArguments += @("--$($entry.Value)",[string]$value) }
    }
    if ($actionName -in @('SetExternalDeliveryFee','SetOrderDeliveryFee') -and $null -eq (Get-CodexActionValue $arguments 'amount')) {
        $deliveryValue = Get-CodexActionValue $arguments 'delivery_fee'
        if ($null -ne $deliveryValue) { $eventArguments += @('--amount',[string]$deliveryValue) }
    }
    if ([bool](Get-CodexActionValue $arguments 'retire_existing_new_code')) { $eventArguments += '--retire-existing-new-code' }
    if ([bool](Get-CodexActionValue $arguments 'clear_expiry')) { $eventArguments += '--clear-expiry' }

    $dateActions = @('AddIntervention','OfflineSale','AddOrderGift','Replenishment','CashExpense','OperatingExpense','OwnerWithdrawal','OwnerLabor','MarketRent','BeginningCash','ActualReceipt','AssetAmortization')
    if ($actionName -in $dateActions -and $null -eq (Get-CodexActionValue $arguments 'date')) {
        $resolvedDate = Resolve-MessageDate ([string]$Event.content) ([string]$Event.create_time)
        $eventArguments += @('--date',$resolvedDate.ToString('yyyy-MM-dd'))
    }
    if ($null -eq (Get-CodexActionValue $arguments 'note')) {
        $eventArguments += @('--note','本机 Codex 通过飞书自然语言解析')
    }
    $eventArguments += @('--event-id',"codex-$($Event.message_id)-$Index-$actionName")

    $result = Invoke-PetstarEvent $eventArguments
    if ($script:previewMode) {
        return @($script:previewOperations)[-1]
    }
    return [pscustomobject]@{
        action = $actionName
        arguments = Convert-PetstarArgumentsToMap $eventArguments
        result = $result.result
        formula_verification = $result.formula_verification
        duplicate = [bool]$result.duplicate
    }
}

function Invoke-CodexActionsInOrder($Actions, $Event, [string]$Mode) {
    $actionList = @($Actions)
    $operations = [Collections.Generic.List[object]]::new()
    $index = 0
    while ($index -lt $actionList.Count) {
        if ($Mode -eq 'preview' -and (Test-BatchCompatibleSkuOperation $actionList[$index])) {
            $compatibleRun = [Collections.Generic.List[object]]::new()
            $scanIndex = $index
            while ($scanIndex -lt $actionList.Count -and (Test-BatchCompatibleSkuOperation $actionList[$scanIndex])) {
                $compatibleRun.Add($actionList[$scanIndex])
                $scanIndex++
            }
            if ($compatibleRun.Count -ge 2) {
                foreach ($batchOperation in @(Invoke-PetstarBatchSkuMaintenance $compatibleRun.ToArray() $Event $index -Preview)) {
                    $script:previewOperations.Add($batchOperation)
                    $operations.Add($batchOperation)
                }
                $index = $scanIndex
                continue
            }
        }
        $operation = Invoke-CodexAction $actionList[$index] $Event $index $Mode
        if ($null -ne $operation) { $operations.Add($operation) }
        $index++
    }
    return $operations.ToArray()
}

function Format-CodexExecutionReply($Operations) {
    $lines = [Collections.Generic.List[string]]::new()
    $mutations = @($Operations | Where-Object { [string]$_.action -notin @('RefreshLedger','SendDailySummary','SyncPlatformData') })
    if ($mutations.Count -gt 0) { $lines.Add('✅ 已按确认内容完成修改：') }
    foreach ($operation in @($Operations)) {
        $args = $operation.arguments
        $details = $operation.result
        if ([bool]$operation.duplicate) {
            $lines.Add("无需重复写入：$($operation.action) 已处理过。")
            continue
        }
        switch ([string]$operation.action) {
            'AddIntervention' { $lines.Add("干预订单第$($details.row)行：$($args.order_no)$(if($details.person){'｜补单人 '+$details.person}else{''})") }
            'RemoveIntervention' { $lines.Add("已撤销错误干预：干预订单第$($details.row)行｜$($args.order_no)；订单重新按自然单核算") }
            'RestoreInventory' { $lines.Add("干预订单第$($details.row)行：库存状态改为已恢复") }
            'ConfirmSelfDelivery' { $lines.Add("干预订单第$($details.row)行：自配送，配送费 ¥0") }
            'SetExternalDeliveryFee' { $lines.Add("干预订单第$($details.row)行：骑手配送费 ¥$($args.amount)") }
            'SetOrderDeliveryFee' { $lines.Add("$($details.order_type) $($details.daily_order_no)：配送费 ¥$($details.old_delivery_fee) → ¥$($details.delivery_fee)") }
            'SetOrderActuals' { $lines.Add("$($details.daily_order_no)｜订单 $($details.order_no)：实收 ¥$($details.new_income)｜实际售出 $($details.new_items)｜货物成本 ¥$($details.expected_new_cost)｜赚/亏 ¥$($details.expected_new_profit)；未修改美团库存。") }
            'OfflineSale' { $lines.Add("线下订单第$($details.row)行：$($details.daily_order_no)｜实收 ¥$($details.income)") }
            'AddOrderGift' {
                if ($details.status -eq 'pending_order') {
                    $lines.Add("赠品记录第$($details.first_row)行：$($args.order_no)｜$($details.gift_items)｜等待订单同步")
                } else {
                    $lines.Add("赠品记录第$($details.first_row)-$($details.last_row)行：$($details.daily_order_no)｜$($details.gift_items)")
                }
            }
            'Replenishment' {
                $lines.Add("补货记录第$($details.row)行：$($details.goods_code)｜$($details.package_quantity) 个采购包装")
                $afterCostText = if ($null -eq $details.after_average_cost -or [string]$details.after_average_cost -eq '') { '待核对' } else { "¥$([math]::Round([double]$details.after_average_cost,4))" }
                $lines.Add("SKU成本主表当前移动平均成本：$afterCostText；未重算历史订单或经营简账。")
                if ($details.sku_status -eq 'created_pending_profile') { $lines.Add('已建立待平台同步的SKU档案。') }
                if ($details.sku_status -eq 'attached_kw_by_csx') { $lines.Add("已按 CSX 匹配原档案并绑定当前 KW：$($details.goods_code)。") }
            }
            'SetSkuBatchDate' { $lines.Add("$($details.goods_code)：到期 $($details.expiry_date)") }
            'SetSkuExpiry' { $lines.Add("$($details.goods_code)：最早到期 $($details.earliest_expiry)") }
            'SetSkuCost' {
                $oldCostText = if ($null -eq $details.old_baseline_cost -or [string]$details.old_baseline_cost -eq '') { '未填写' } else { "¥$($details.old_baseline_cost)" }
                $historyText = if ($details.historical_order_no) { "；已按明确要求更新订单 $($details.historical_order_no) 并刷新所属月份" } else { '；未修改历史订单或经营简账' }
                $lines.Add("SKU成本主表第$($details.row)行：$($details.goods_code) 销售单位成本 $oldCostText → ¥$($details.new_unit_cost)；当前成本现为 ¥$($details.new_moving_average_cost)$historyText")
            }
            'SetRefundDisposition' {
                $dispositionText = switch ([string]$details.new_disposition) { 'returned_restocked' {'已退回并恢复库存，已撤销货物成本'} 'refund_no_return' {'未退货，已保留货物成本'} 'returned_damaged' {'退回报损，已保留货物成本'} default {[string]$details.new_disposition} }
                $lines.Add("订单 $($details.order_no)：$($details.goods_code)×$($details.quantity)｜$dispositionText；收入、成本和月份简账已重算。")
            }
            'ChangeSkuCode' { $lines.Add("当前货号：$($details.old_code) → $($details.new_code)") }
            'RefreshLedger' {
                $dateSummary = $details.requested_date_summary
                if ($null -ne $dateSummary) {
                    $lines.Add("✅ 已重新计算 $($dateSummary.date)：")
                    $lines.Add("订单 $($dateSummary.order_count) 单｜扣佣后收入 ¥$($dateSummary.income)｜配送费 ¥$($dateSummary.delivery_fee)")
                    if ([int]$dateSummary.pending_cost_count -gt 0) {
                        $lines.Add("⚠️ 仍有 $($dateSummary.pending_cost_count) 单成本待核对，当天赚/亏暂不显示：")
                        foreach ($pendingOrder in @($dateSummary.pending_orders)) {
                            $pendingLabel = if ($pendingOrder.daily_order_no) { [string]$pendingOrder.daily_order_no } else { [string]$pendingOrder.order_no }
                            $pendingNote = if ($pendingOrder.note) { "｜$($pendingOrder.note)" } else { '' }
                            $lines.Add("- $pendingLabel$pendingNote")
                        }
                    } else {
                        $lines.Add("货物成本 ¥$($dateSummary.product_cost)｜当天赚/亏 ¥$($dateSummary.profit)")
                    }
                    $lines.Add('没有修改人工录入字段。')
                } else {
                    $lines.Add('已刷新每日经营简账；没有修改人工录入字段。')
                }
            }
            'SendDailySummary' { $lines.Add('已补发经营日结；没有修改经营数据。') }
            'SyncPlatformData' { $lines.Add('已同步订单、库存与经营表。') }
            default {
                $dateText = if ($args.date) { "｜$($args.date)" } else { '' }
                $amountText = if ($args.amount) { "｜¥$($args.amount)" } else { '' }
                $lines.Add("已执行 $($operation.action)$dateText$amountText")
            }
        }
    }
    if ($lines.Count -eq 0) { return '没有执行任何数据修改。' }
    return $lines -join "`n"
}

function Get-OfflineSaleFacts([string]$Text) {
    $offlineMatch = [regex]::Match($Text, '线下(?:单|订单|卖出|卖了|销售|成交)?')
    if (-not $offlineMatch.Success) { return $null }
    $segment = $Text.Substring($offlineMatch.Index)
    $itemMatches = @([regex]::Matches($segment, '(?i)\b((?:kw|csx)\s*\d+)\s*(?:[×xX*]\s*(\d+(?:\.\d+)?)|(?:卖出(?:了)?|卖了|售出(?:了)?)?\s*(\d+(?:\.\d+)?)\s*(?:个|件|包|袋|盒|罐|支|瓶))'))
    $parts = [Collections.Generic.List[string]]::new()
    foreach ($itemMatch in $itemMatches) {
        $code = ([string]$itemMatch.Groups[1].Value).Replace(' ','').ToLowerInvariant()
        $quantityText = if ($itemMatch.Groups[2].Success) { [string]$itemMatch.Groups[2].Value } else { [string]$itemMatch.Groups[3].Value }
        if ($code -and $quantityText) { $parts.Add("$code×$quantityText") }
    }
    $amountMatch = [regex]::Match($segment, '(?:销售总金额|合计收款|合计实收|实收|收款|收了|金额)\s*(?:为|是|：|:)?\s*[￥¥]?\s*(\d{1,7}(?:\.\d{1,2})?)\s*元?')
    if ($parts.Count -eq 0 -or -not $amountMatch.Success) {
        return [pscustomobject]@{ complete=$false; items=($parts -join '、'); amount=$null; segment=$segment }
    }
    $amount = [double]::Parse($amountMatch.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
    $fee = Get-DeliveryFee $segment
    if ($null -eq $fee) { $fee = 0.0 }
    $receiptStatus = if ($segment -match '赊账|未到账|还没收(?:到)?款|尚未收(?:到)?款') { '未到账' } else { '已到账' }
    return [pscustomobject]@{ complete=$true; items=($parts -join '、'); amount=$amount; delivery_fee=[double]$fee; receipt_status=$receiptStatus; segment=$segment }
}

function Invoke-OfflineSaleEvent($Event, [datetime]$Date, $Facts) {
    $args = @(
        '--action','OfflineSale','--date',$Date.ToString('yyyy-MM-dd'),
        '--amount',[string]$Facts.amount,'--items',[string]$Facts.items,
        '--delivery-fee',[string]$Facts.delivery_fee,'--receipt-status',[string]$Facts.receipt_status,
        '--note','用户通过自然语言登记线下销售','--event-id',"$($Event.message_id)-offline-sale"
    )
    return Invoke-PetstarEvent $args
}

function Get-DeliveryFee([string]$Text) {
    $feeMatch = [regex]::Match($Text, '(?:外部(?:配送|骑手|跑腿)?(?:费|费用)?|配送费(?:用)?|运费|跑腿费(?:用)?)\s*(?:是|为|：|:|共|花了|实付|支付|付了|改成|改为|调整为|更正为|设为)?\s*[￥¥]?\s*(\d{1,5}(?:\.\d{1,2})?)\s*元?')
    if (-not $feeMatch.Success) {
        $feeMatch = [regex]::Match($Text, '[￥¥]?\s*(\d{1,5}(?:\.\d{1,2})?)\s*元\s*(?:的)?(?:外部配送费(?:用)?|配送费(?:用)?|运费|跑腿费(?:用)?)')
    }
    if (-not $feeMatch.Success) { return $null }
    return [double]::Parse($feeMatch.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
}

function Get-InterventionPerson([string]$Text) {
    $match = [regex]::Match($Text, '补单人\s*(?:是|叫|为|：|:)?\s*([^，,。；;\r\n]{1,24})')
    if (-not $match.Success) { return $null }
    return $match.Groups[1].Value.Trim()
}

function Test-SelfDeliveryWithoutFee([string]$Text) {
    if ($Text -match '(?:没有|没|无需|不用|不需要|不产生|免)(?:出|付|收)?\s*(?:配送费|运费|跑腿费)') { return $true }
    if ($Text -match '(?:配送费|运费|跑腿费)\s*(?:是|为|：|:)?\s*(?:0(?:\.0+)?|零)\s*元?') { return $true }
    if ($Text -match '(?:我|本人|店主|商家)?\s*(?:自己送|自己配送|自行配送|自配送|商家自配)') { return $true }
    return $false
}

function Test-InventoryRestored([string]$Text) {
    return [bool]($Text -match '(?:库存|货品|商品).{0,8}(?:已|已经)?(?:恢复|改回|还原)|(?:已|已经)(?:恢复|改回|还原).{0,8}(?:库存|货品|商品)')
}

function Get-OrderSegments([string]$Text) {
    $matches = @([regex]::Matches($Text, '(?<!\d)\d{18,20}(?!\d)'))
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $segments = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $matches.Count; $index++) {
        $orderNo = [string]$matches[$index].Value
        if (-not $seen.Add($orderNo)) { continue }
        $start = $matches[$index].Index
        $end = if ($index + 1 -lt $matches.Count) { $matches[$index + 1].Index } else { $Text.Length }
        $segments.Add([pscustomobject]@{
            order_no = $orderNo
            text = $Text.Substring($start, $end - $start)
        })
    }
    return $segments.ToArray()
}

function Get-GiftFacts([string]$Text) {
    if (-not $Text -or $Text -notmatch '赠品|加送|赠送|送了|作为礼物|随单送') { return $null }
    $giftCue = [regex]::Match($Text, '赠品|加送(?:了)?|赠送(?:了)?|送了|随单送(?:了)?')
    $segment = if ($giftCue.Success) { $Text.Substring($giftCue.Index) } else { $Text }
    $itemMatches = @([regex]::Matches($segment, '(?i)(?<![A-Za-z0-9])((?:kw|csx)\s*\d+)\s*(?:[×xX*]\s*(\d+(?:\.\d+)?)|(?:送|赠送|加送)?(?:了)?\s*(\d+(?:\.\d+)?)\s*(?:个|件|包|袋|盒|罐|支|瓶|份))'))
    if ($itemMatches.Count -eq 0 -and $Text -match '作为(?:赠品|礼物)') {
        $itemMatches = @([regex]::Matches($Text, '(?i)(?<![A-Za-z0-9])((?:kw|csx)\s*\d+)\s*(?:[×xX*]\s*(\d+(?:\.\d+)?)|(?:送|赠送|加送)?(?:了)?\s*(\d+(?:\.\d+)?)\s*(?:个|件|包|袋|盒|罐|支|瓶|份))'))
    }
    $items = [Collections.Generic.List[string]]::new()
    foreach ($match in $itemMatches) {
        $code = ([string]$match.Groups[1].Value).Replace(' ','').ToLowerInvariant()
        $quantity = if ($match.Groups[2].Success) { [string]$match.Groups[2].Value } else { [string]$match.Groups[3].Value }
        if ($code -and $quantity) { $items.Add("$code×$quantity") }
    }
    $orderMatch = [regex]::Match($Text, '(?<!\d)\d{18,20}(?!\d)')
    $dailyMatch = [regex]::Match($Text, '(?i)(美团|线下|饿了么|京东到家|淘宝闪购)\s*(\d+)\s*号单')
    return [pscustomobject]@{
        complete = $items.Count -gt 0 -and ($orderMatch.Success -or $dailyMatch.Success)
        order_no = $(if ($orderMatch.Success) { [string]$orderMatch.Value } else { $null })
        daily_order_no = $(if ($dailyMatch.Success) { "$($dailyMatch.Groups[1].Value)$($dailyMatch.Groups[2].Value)号单" } else { $null })
        items = $items -join '、'
    }
}

function Invoke-GiftEvent($Event, [datetime]$Date, $Facts) {
    $args = @('--action','AddOrderGift','--date',$Date.ToString('yyyy-MM-dd'),'--items',[string]$Facts.items,'--note','用户通过自然语言登记订单赠品','--event-id',"$($Event.message_id)-order-gift")
    if ($Facts.order_no) { $args += @('--order-no',[string]$Facts.order_no) }
    if ($Facts.daily_order_no) { $args += @('--daily-order-no',[string]$Facts.daily_order_no) }
    return Invoke-PetstarEvent $args
}

function Test-ActionableBusinessMessage([string]$Text) {
    if (-not $Text) { return $false }
    if ($Text -match '(?<!\d)\d{18,20}(?!\d)') { return $true }
    return [bool]($Text -match '补单|干预(?:订)?单|刷单|配送费|运费|线下(?:单|卖|销售)|补货|进货|库存|支出|简账|费用|收款|到账|取款|生产日期|保质期|过期|到期|有效期|赠品|加送|赠送|随单送|SKU|货号|(?:kw\s*\d+).{0,12}(?:修改|更改|改|调整|更换|换)')
}

function Test-ReadOnlyOperationalMessage([string]$Text) {
    if (-not $Text) { return $false }
    $requestsRefresh = $Text -match '(?:重新)?(?:同步|刷新|重算|重新计算|更新|汇总)'
    $targetsReport = $Text -match '(?:收入表|支出表|每日经营简账|经营简账|简账|订单情况|经营情况|当天.{0,6}(?:赚|亏)|经营日结)'
    if (-not ($requestsRefresh -and $targetsReport)) { return $false }

    # “我已经改了一下成本，所以请刷新”是在陈述已发生的事实；
    # 只有同时给出明确的新值时，才把它当成新的数据修改请求。
    $containsExplicitNewValue = $Text -match '(?i)(?:kw|csx)\s*\d+.{0,24}(?:成本|货号|生产日期|保质期|过期|到期).{0,12}(?:改成|改为|设为|调整为|是|为)\s*(?:\d|kw|csx)' -or
        $Text -match '(?<!\d)\d{18,20}(?!\d).{0,30}(?:配送费|运费).{0,12}(?:改成|改为|设为|是|为)?\s*\d'
    return -not $containsExplicitNewValue
}

function Test-DataMutationMessage([string]$Text) {
    if (-not $Text) { return $false }
    if (Test-ReadOnlyOperationalMessage $Text) { return $false }
    if ($Text -match '(?:同步|刷新|重建).{0,10}(?:支出|简账)|(?:支出|简账).{0,10}(?:同步|刷新|重建)') { return $false }
    if ($Text -match '(?:同步|刷新|拉取|更新).{0,12}(?:美团数据|订单|库存快照|SKU成本主表|库存)' -and $Text -notmatch '配送费|运费|干预(?:订)?单|补单|线下|补货|进货|采购|支出|费用|付款|到账|取款|生产日期|生产日|保质期|过期|到期|赠品|加送|货号.{0,8}(?:改|换|绑定)|(?<!\d)\d{18,20}(?!\d)') { return $false }
    $explicitWrite = $Text -match '登记|记录|记下|补录|填入|填上|补上|修改|更改|改成|改为|调整|设置|设为|绑定|更新|恢复库存|改回库存|补货|进货|采购|线下(?:卖|成交|销售)|卖出|售出|干预(?:订)?单|补单|刷单|加送|赠送|赠品|支出|花了|付款|已付款|到账|取款|生产日期|生产日|保质期|最早过期|最早到期'
    if ($explicitWrite) { return $true }
    if ($Text -match '(?<!\d)\d{18,20}(?!\d).{0,30}(?:配送费|运费|自配送|自己送|骑手配送)') { return $true }
    if ($Text -match '(?:配送费|运费|自配送|自己送|骑手配送).{0,30}(?<!\d)\d{18,20}(?!\d)') { return $true }
    return $false
}

function Test-ConfirmationText([string]$Text) {
    return [bool]($Text -match '^\s*(?:确认|确认执行|确定|执行|可以|没问题|就这样|好的?[,，]?\s*(?:确认|执行))(?:\s*[:：#-]?\s*([A-Fa-f0-9]{8}))?\s*[。！!]*\s*$')
}

function Get-RequestedConfirmationId([string]$Text) {
    if ($Text -match '(?i)([a-f0-9]{8})') { return $Matches[1].ToUpperInvariant() }
    return ''
}

function Test-CancellationText([string]$Text) {
    return [bool]($Text -match '^\s*(?:取消|算了|不用了|不改了|放弃|取消这次)\s*[。！!]*\s*$')
}

function Get-ConversationHash($Event) {
    return Get-MessageHash ("$([string]$Event.sender_id)|$([string]$Event.chat_id)")
}

function Invoke-WithPendingConfirmationLock([scriptblock]$Action) {
    $pendingMutex = [Threading.Mutex]::new($false, 'Local\PetstarLarkPendingConfirmations')
    $pendingLockTaken = $false
    try {
        try { $pendingLockTaken = $pendingMutex.WaitOne([timespan]::FromSeconds(15)) }
        catch [Threading.AbandonedMutexException] { $pendingLockTaken = $true }
        if (-not $pendingLockTaken) { throw '待确认队列正忙，未能在15秒内取得锁' }
        return & $Action
    } finally {
        if ($pendingLockTaken) { try { $pendingMutex.ReleaseMutex() } catch {} }
        $pendingMutex.Dispose()
    }
}

function Read-PendingConfirmationStateUnsafe {
    $emptyState = [pscustomobject]@{version=2;active=$null;queue=@();updated_at=[datetime]::Now.ToString('o')}
    if (-not (Test-Path -LiteralPath $pendingConfirmationsPath)) { return $emptyState }
    try {
        $state = Get-Content -LiteralPath $pendingConfirmationsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $queue = if ($state.PSObject.Properties['queue']) { @($state.queue | Where-Object { $null -ne $_ }) } else { @() }
        return [pscustomobject]@{version=2;active=$state.active;queue=$queue;updated_at=$state.updated_at}
    } catch {
        Write-BridgeLog 'pending confirmation state could not be read; ignoring invalid state'
        return $emptyState
    }
}

function Save-PendingConfirmationStateUnsafe($State) {
    $cleanQueue = @($State.queue | Where-Object { $null -ne $_ })
    $payload = [ordered]@{version=3;active=$State.active;queue=$cleanQueue;updated_at=[datetime]::Now.ToString('o')} | ConvertTo-Json -Depth 40
    $tempPath = "$pendingConfirmationsPath.$([guid]::NewGuid().ToString('N')).tmp"
    [IO.File]::WriteAllText($tempPath,$payload,[Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tempPath -Destination $pendingConfirmationsPath -Force
}

function Read-PendingConfirmationState {
    return Invoke-WithPendingConfirmationLock { return Read-PendingConfirmationStateUnsafe }
}

function Save-PendingConfirmationState($State) {
    Invoke-WithPendingConfirmationLock { Save-PendingConfirmationStateUnsafe $State }
}

function Read-PendingConfirmation {
    return (Read-PendingConfirmationState).active
}

function Write-PendingConfirmation($Pending) {
    Invoke-WithPendingConfirmationLock {
        $state = Read-PendingConfirmationStateUnsafe
        $state.active = $Pending
        Save-PendingConfirmationStateUnsafe $state
    }
}

function Add-PendingConfirmation($Pending) {
    return Invoke-WithPendingConfirmationLock {
        $state = Read-PendingConfirmationStateUnsafe
        $candidates = @()
        if ($state.active) { $candidates += [pscustomobject]@{item=$state.active;position=1} }
        for ($index = 0; $index -lt @($state.queue).Count; $index++) {
            $candidates += [pscustomobject]@{item=@($state.queue)[$index];position=2+$index}
        }
        $duplicate = @($candidates | Where-Object {
            [string]$_.item.conversation_hash -eq [string]$Pending.conversation_hash -and
            [string]$_.item.operation_signature -eq [string]$Pending.operation_signature
        } | Select-Object -First 1)
        if ($duplicate.Count -gt 0) {
            return [pscustomobject]@{
                position = [int]$duplicate[0].position
                active = $state.active
                queued = @($state.queue).Count
                duplicate = $true
                existing = $duplicate[0].item
            }
        }
        if (-not $state.active) {
            $state.active = $Pending
            $position = 1
        } else {
            $state.queue = @($state.queue) + @($Pending)
            $position = 1 + @($state.queue).Count
        }
        Save-PendingConfirmationStateUnsafe $state
        return [pscustomobject]@{position=$position;active=$state.active;queued=@($state.queue).Count;duplicate=$false;existing=$null}
    }
}

function Clear-PendingConfirmation {
    return Invoke-WithPendingConfirmationLock {
        $state = Read-PendingConfirmationStateUnsafe
        $next = $null
        if (@($state.queue).Count -gt 0) {
            $next = @($state.queue)[0]
            $state.active = $next
            $state.queue = if (@($state.queue).Count -gt 1) { @($state.queue)[1..(@($state.queue).Count-1)] } else { @() }
        } else {
            $state.active = $null
            $state.queue = @()
        }
        Save-PendingConfirmationStateUnsafe $state
        return $next
    }
}

function Save-PendingOperationProgress([string]$ChangeId, [int]$Index, $Result) {
    Invoke-WithPendingConfirmationLock {
        $state = Read-PendingConfirmationStateUnsafe
        if (-not $state.active -or [string]$state.active.change_id -ne $ChangeId) {
            throw "待确认项 $ChangeId 已不再是当前项目，拒绝写入执行进度"
        }
        [object[]]$completed = @($state.active.completed_operations | Where-Object { $null -ne $_ })
        if (-not @($completed | Where-Object { [int]$_.index -eq $Index }).Count) {
            $completed += [pscustomobject]@{index=$Index;result=$Result;completed_at=[datetime]::Now.ToString('o')}
        }
        $state.active | Add-Member -NotePropertyName completed_operations -NotePropertyValue @($completed) -Force
        Save-PendingConfirmationStateUnsafe $state
    }
}

function Complete-PendingConfirmation([string]$ChangeId) {
    return Invoke-WithPendingConfirmationLock {
        $state = Read-PendingConfirmationStateUnsafe
        if (-not $state.active -or [string]$state.active.change_id -ne $ChangeId) {
            throw "待确认项 $ChangeId 状态已变化，拒绝重复完成"
        }
        $next = $null
        if (@($state.queue).Count -gt 0) {
            $next = @($state.queue)[0]
            $state.active = $next
            $state.queue = if (@($state.queue).Count -gt 1) { @($state.queue)[1..(@($state.queue).Count-1)] } else { @() }
        } else {
            $state.active = $null
            $state.queue = @()
        }
        Save-PendingConfirmationStateUnsafe $state
        return $next
    }
}

function Format-QueuedConfirmationReply($Draft, [int]$Position, $Active) {
    if ($Position -le 1) { return [string]$Draft.preview }
    return "已保存为第 $Position 项待确认，没有丢失。`n`n本条准备修改：`n$($Draft.preview)`n`n当前先处理第1项 $($Active.change_id)。回复确认或取消后，我会自动展示下一项。"
}

function Get-ActivePendingConfirmation($Event) {
    while ($true) {
        $pending = Read-PendingConfirmation
        if (-not $pending) { return $null }
        if ([string]$pending.conversation_hash -ne (Get-ConversationHash $Event)) { return $null }
        $expiresAt = [datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse([string]$pending.expires_at,[ref]$expiresAt) -and $expiresAt -ge [datetimeoffset]::Now) { return $pending }
        $null = Clear-PendingConfirmation
    }
}

function Test-PendingAmendment([string]$Text, $Pending) {
    if (-not $Pending -or -not $Text) { return $false }
    # 刷新报表是独立维护动作，不得因为包含“改了/更新”而串到上一项待确认变更。
    if (Test-ReadOnlyOperationalMessage $Text) { return $false }
    if ($Pending.needs_information -and $Text -match '(?<!\d)\d{18,20}(?!\d)|(?i)(?:kw|csx)\s*\d+|\d+(?:\.\d+)?\s*(?:元|天|个?月|年|件|包|袋|盒|罐)|补单人|自配送|自己送|骑手|已恢复|未恢复') { return $true }
    $newSkuCodes = @([regex]::Matches($Text,'(?i)(?<![a-z0-9])(?:kw|csx)\s*\d+') | ForEach-Object { $_.Value.Replace(' ','').ToLowerInvariant() } | Select-Object -Unique)
    if ($newSkuCodes.Count -gt 0) {
        $pendingHasSkuAction = @($Pending.operations | Where-Object { [string]$_.action -in @('Replenishment','SetSkuExpiry','SetSkuBatchDate','SetSkuCost','ChangeSkuCode','AddOrderGift','OfflineSale') }).Count -gt 0
        if (-not $pendingHasSkuAction) { return $false }
    }
    if ($Text -notmatch '改|更正|纠正|补充|不是|应该|日期|配送费|运费|补单人|数量|金额|生产日期|保质期|已恢复|没恢复|自配送|自己送|骑手') { return $false }
    $newOrders = @([regex]::Matches($Text,'(?<!\d)\d{18,20}(?!\d)') | ForEach-Object { $_.Value } | Select-Object -Unique)
    foreach ($orderNo in $newOrders) {
        if ([string]$Pending.effective_message -notmatch [regex]::Escape([string]$orderNo) -and $Text -notmatch '订单号.{0,6}(?:改|更正)') { return $false }
    }
    return $true
}

function Get-OperationSignature($Operations) {
    $normalized = @()
    foreach ($operation in @($Operations)) {
        $arguments = [ordered]@{}
        foreach ($property in @($operation.arguments.PSObject.Properties | Sort-Object Name)) {
            if ($property.Name -eq 'event_id') { continue }
            $arguments[$property.Name] = $property.Value
        }
        $normalized += [ordered]@{action=$operation.action;arguments=$arguments}
    }
    return ($normalized | ConvertTo-Json -Depth 20 -Compress)
}

function Format-ConfirmationPreview([string]$ChangeId, $Operations, [string]$FallbackText) {
    if (@($Operations).Count -eq 0) { return $FallbackText }
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add("📝 待确认｜$ChangeId")
    $lines.Add('准备修改：')
    $impacts = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($operation in @($Operations)) {
        $args = $operation.arguments
        $details = $operation.result
        switch ([string]$operation.action) {
            'AddIntervention' {
                $personText = if ($args.person) { "｜补单人：$($args.person)" } else { '' }
                $lines.Add("干预订单｜按订单号定位或写入首个空行：$($args.order_no)｜日期：$($args.date)$personText")
                $null = $impacts.Add('干预订单')
                $null = $impacts.Add('对应月份收入/支出')
                $null = $impacts.Add('当天赚/亏')
            }
            'RemoveIntervention' {
                $lines.Add("撤销错误干预：干预订单第$($details.row)行｜$($args.order_no)｜补单人：$($details.person)｜配送：$($details.delivery_type) ¥$($details.delivery_fee)｜库存：$($details.inventory_state)")
                $null = $impacts.Add('干预订单')
                $null = $impacts.Add('订单汇总及订单商品成本')
                $null = $impacts.Add('对应月份收入/支出及当天赚/亏')
            }
            'ConfirmSelfDelivery' { $lines.Add("干预订单｜按订单号定位：$($args.order_no)｜配送方式改为自配送、配送费改为 ¥0"); $null = $impacts.Add('干预订单') }
            'SetExternalDeliveryFee' { $lines.Add("干预订单｜按订单号定位：$($args.order_no)｜配送方式改为骑手配送、配送费改为 ¥$($args.amount)"); $null = $impacts.Add('干预订单') }
            'SetOrderDeliveryFee' {
                $orderName = if ($details.daily_order_no) { [string]$details.daily_order_no } else { [string]$args.order_no }
                $lines.Add("配送费：$orderName｜¥$([double]$details.old_delivery_fee) → ¥$([double]$details.delivery_fee)｜$($details.order_type)")
                $null = $impacts.Add($(if ($details.order_type -eq '干预单') { '干预订单及月份支出' } else { '订单汇总及月份收入' }))
                $null = $impacts.Add('单笔及当天赚/亏')
            }
            'SetOrderActuals' {
                $oldItemsText = @($details.old_items) -join '、'
                $actualItemText = @($details.actual_items | ForEach-Object {
                    $lineAmount = [math]::Round([double]$_.unit_price * [double]$_.quantity, 2)
                    "$($_.goods_code)×$($_.quantity)（成交 ¥$lineAmount）"
                }) -join '、'
                $costSourceText = @($details.cost_details | ForEach-Object {
                    "$($_.goods_code)：SKU成本主表第$($_.sku_master_row)行，成本 ¥$($_.unit_cost)"
                }) -join '；'
                $lines.Add("实际履约：$($details.daily_order_no)｜订单 $($details.order_no)｜订单汇总第$($details.order_summary_row)行")
                $lines.Add("实收：¥$($details.old_income) → ¥$($details.new_income)；配送费：¥$($details.old_delivery_fee) → ¥$($details.new_delivery_fee)")
                $lines.Add("实际商品：$oldItemsText → $actualItemText")
                $lines.Add("成本依据：$costSourceText")
                $lines.Add("预计货物成本 ¥$($details.expected_new_cost)；预计赚/亏 ¥$($details.expected_new_profit)；美团库存不修改。")
                foreach ($sheet in @($details.affected_tables)) { if ($sheet) { $null = $impacts.Add([string]$sheet) } }
            }
            'RestoreInventory' { $lines.Add("干预订单｜按订单号定位：$($args.order_no)｜库存状态改为已恢复"); $null = $impacts.Add('干预订单') }
            'OfflineSale' {
                $costText = if ($null -eq $details.product_cost) { '成本待核对' } else { "成本 ¥$([double]$details.product_cost)" }
                $lines.Add("线下单：$($args.date)｜$($details.sold_goods)｜实收 ¥$([double]$details.income)｜$costText")
                $null = $impacts.Add('线下订单及对应月份收入')
                $null = $impacts.Add('月度到账及当天赚/亏')
            }
            'AddOrderGift' {
                $lines.Add("赠品：$($details.daily_order_no)｜$($details.gift_items)｜成本 $(if($null -eq $details.gift_cost){'待核对'}else{'¥'+[string][double]$details.gift_cost})")
                if ($details.status -eq 'pending_order') { $lines.Add('订单尚未同步；确认后先暂存，后续自动匹配。') }
                $null = $impacts.Add('赠品记录及订单成本')
                $null = $impacts.Add('对应月份收入/支出')
            }
            'Replenishment' {
                $lines.Add("补货：$($details.goods_code)｜日期：$($args.date)｜$($details.package_quantity)个采购包装｜采购 ¥$($details.purchase_amount)")
                if ($details.production_date) { $lines.Add("批次：生产 $($details.production_date)｜到期 $($details.expiry_date)") }
                $afterCostText = if ($null -eq $details.after_average_cost -or [string]$details.after_average_cost -eq '') { '仍待核对' } else { "¥$([math]::Round([double]$details.after_average_cost,4))" }
                $lines.Add("SKU当前移动平均成本：$afterCostText；不会重算历史订单、收入、利润或经营简账。")
                if ($details.sku_status -eq 'will_create_pending_profile') { $lines.Add('SKU主表尚无该货号；确认后会先建立待平台同步档案。') }
                if ($details.sku_status -eq 'will_attach_kw_by_csx') { $lines.Add("已通过 $($details.transition_code) 唯一找到 SKU主表第$($details.sku_master_row)行；确认后绑定当前KW为 $($details.goods_code)。") }
                $null = $impacts.Add('补货记录及SKU成本主表')
            }
            'SetSkuBatchDate' {
                $lines.Add("批次：$($details.goods_code)｜生产 $($details.production_date)｜保质期 $($details.shelf_life)｜到期 $($details.expiry_date)")
                $null = $impacts.Add('库存批次及SKU成本主表')
            }
            'SetSkuExpiry' {
                $lines.Add("最早到期：$($details.goods_code)｜$($details.old_expiry) → $($details.new_expiry)")
                $null = $impacts.Add('SKU成本主表')
            }
            'SetSkuCost' {
                $titleText = if ($details.title) { "｜$($details.title)" } else { '' }
                $oldCostText = if ($null -eq $details.old_baseline_cost -or [string]$details.old_baseline_cost -eq '') { '未填写' } else { "¥$($details.old_baseline_cost)" }
                $lines.Add("修改：SKU成本主表第$($details.row)行｜$($details.goods_code)$titleText")
                $lines.Add("成本：$oldCostText → ¥$($details.new_unit_cost)（每个销售单位）")
                $lines.Add("位置：$($details.range)；只更新SKU成本主表，不会自动重算历史订单。")
                if ($details.historical_order_no) {
                    $lines.Add("明确联动历史订单：$($details.historical_order_no)；只修改该订单中此SKU的成本，并刷新该订单所属月份。")
                }
                foreach ($sheet in @($details.affected_tables)) { if ($sheet) { $null = $impacts.Add([string]$sheet) } }
            }
            'SetRefundDisposition' {
                $dispositionText = switch ([string]$details.new_disposition) { 'returned_restocked' {'已退回并恢复库存'} 'refund_no_return' {'未退货，保留成本'} 'returned_damaged' {'退回报损，保留成本'} default {[string]$details.new_disposition} }
                $lines.Add("退款商品：订单 $($details.order_no)｜$($details.goods_code)×$($details.quantity)｜顾客退款 ¥$($details.refund_amount)")
                $lines.Add("处理：$dispositionText")
                foreach ($sheet in @($details.affected_tables)) { if ($sheet) { $null = $impacts.Add([string]$sheet) } }
            }
            'ChangeSkuCode' {
                $lines.Add("当前货号：$($details.old_code) → $($details.new_code)")
                $null = $impacts.Add('SKU成本主表及成本快照')
            }
            { $_ -in @('CashExpense','OperatingExpense','OwnerWithdrawal','OwnerLabor','MarketRent','BeginningCash','ActualReceipt','AssetAmortization') } {
                $lines.Add("月度结算｜$($args.date)｜$($operation.action)｜金额 ¥$($args.amount)$(if($args.note){'｜'+$args.note}else{''})")
                $null = $impacts.Add('月度结算及驾驶舱')
            }
            default {
                throw "无法为待确认动作生成明确预览：$($operation.action)"
            }
        }
    }
    if ($impacts.Count -gt 0) { $lines.Add('将更新：' + (@($impacts) -join '、')) }
    $lines.Add('回复“确认”执行；直接说“改成……”可以调整；回复“取消”放弃。')
    return $lines -join "`n"
}

function Get-HandledMessageHashes {
    $hashes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path -LiteralPath $handledMessagesLog)) { return ,$hashes }
    foreach ($line in Get-Content -LiteralPath $handledMessagesLog -Encoding UTF8) {
        if (-not $line.Trim().StartsWith('{')) { continue }
        try { $record = $line | ConvertFrom-Json } catch { continue }
        if ($record.message_hash -and $record.status -in @('completed', 'baseline', 'failed_replied', 'needs_clarification')) {
            $null = $hashes.Add([string]$record.message_hash)
        }
    }
    return ,$hashes
}

function Invoke-WithHandledLedgerLock([scriptblock]$Action) {
    $ledgerMutex = [Threading.Mutex]::new($false, 'Local\PetstarLarkHandledMessageLedger')
    $ledgerLockTaken = $false
    try {
        try { $ledgerLockTaken = $ledgerMutex.WaitOne([timespan]::FromSeconds(15)) }
        catch [Threading.AbandonedMutexException] { $ledgerLockTaken = $true }
        if (-not $ledgerLockTaken) { throw '消息处理账本正忙，未能在15秒内取得锁' }
        return & $Action
    } finally {
        if ($ledgerLockTaken) { try { $ledgerMutex.ReleaseMutex() } catch {} }
        $ledgerMutex.Dispose()
    }
}

function Add-HandledMessageRecord([string]$MessageId, [string]$Status, [string]$Source, [string]$CreateTime) {
    $hash = Get-MessageHash $MessageId
    if ($Status -in @('completed','baseline','failed_replied','needs_clarification') -and $script:handledMessageHashes.Contains($hash)) { return }
    $record = [ordered]@{
        message_hash = $hash
        status = $Status
        source = $Source
        create_time = $CreateTime
        recorded_at = [datetime]::Now.ToString('o')
    }
    Invoke-WithHandledLedgerLock {
        Add-Content -LiteralPath $handledMessagesLog -Value ($record | ConvertTo-Json -Compress) -Encoding UTF8
    }
    if ($Status -in @('completed','baseline','failed_replied','needs_clarification')) { $null = $script:handledMessageHashes.Add($hash) }
}

function Test-MessageHandled([string]$Hash) {
    if ($script:handledMessageHashes.Contains($Hash)) { return $true }
    $diskHashes = Get-HandledMessageHashes
    if ($diskHashes.Contains($Hash)) {
        $null = $script:handledMessageHashes.Add($Hash)
        return $true
    }
    return $false
}

function Get-ReplayChat {
    $pageToken = ''
    for ($page = 0; $page -lt 20; $page++) {
        $arguments = @('im', '+chat-list', '--as', 'user', '--types', 'p2p', '--sort', 'active_time', '--page-size', '100', '--format', 'json')
        if ($pageToken) { $arguments += @('--page-token', $pageToken) }
        $raw = & $larkCli @arguments 2>$null | Out-String
        if ($LASTEXITCODE -ne 0) { throw '无法读取飞书私聊列表' }
        $response = $raw | ConvertFrom-Json
        if (-not $response.ok) { throw "无法读取飞书私聊列表：$($response.error.message)" }
        $matches = @($response.data.chats | Where-Object { $_.chat_mode -eq 'p2p' -and $_.name -eq $ChatName })
        if ($matches.Count -gt 1) { throw "发现多个同名私聊：$ChatName" }
        if ($matches.Count -eq 1) { return $matches[0] }
        if (-not $response.data.has_more -or -not $response.data.page_token) { break }
        $pageToken = [string]$response.data.page_token
    }
    throw "找不到飞书私聊：$ChatName"
}

function Get-RecentReplayMessages {
    $chat = Get-ReplayChat
    $start = [DateTimeOffset]::Now.AddHours(-1 * $ReplayHours).ToString('o')
    $end = [DateTimeOffset]::Now.ToString('o')
    $pageToken = ''
    $messages = [Collections.Generic.List[object]]::new()
    for ($page = 0; $page -lt 20; $page++) {
        $arguments = @(
            'im', '+chat-messages-list', '--as', 'user', '--chat-id', [string]$chat.chat_id,
            '--start', $start, '--end', $end, '--order', 'asc', '--page-size', '50',
            '--no-reactions', '--format', 'json'
        )
        if ($pageToken) { $arguments += @('--page-token', $pageToken) }
        $raw = & $larkCli @arguments 2>$null | Out-String
        if ($LASTEXITCODE -ne 0) { throw '无法读取飞书最近消息' }
        $response = $raw | ConvertFrom-Json
        if (-not $response.ok) { throw "无法读取飞书最近消息：$($response.error.message)" }
        foreach ($message in @($response.data.messages)) { $messages.Add($message) }
        if (-not $response.data.has_more -or -not $response.data.page_token) { break }
        $pageToken = [string]$response.data.page_token
    }
    return @($messages)
}

function Get-EligibleReplayMessages([object[]]$Messages, [string]$AuthorizedOpenId) {
    return @($Messages | Where-Object {
        -not $_.deleted -and
        $_.msg_type -eq 'text' -and
        $_.sender.sender_type -eq 'user' -and
        $_.sender.id -eq $AuthorizedOpenId -and
        $_.message_id
    })
}

function Convert-ReplayMessageToEvent($Message) {
    $createTime = [string]$Message.create_time
    $parsedTime = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($createTime, [ref]$parsedTime)) {
        $createTime = [string]$parsedTime.ToUnixTimeMilliseconds()
    }
    return [ordered]@{
        chat_id = [string]$Message.chat_id
        chat_type = 'p2p'
        content = [string]$Message.content
        create_time = $createTime
        message_id = [string]$Message.message_id
        message_type = [string]$Message.msg_type
        sender_id = [string]$Message.sender.id
        sender_type = [string]$Message.sender.sender_type
    }
}

function New-ConfirmationDraft($Event, [string]$AuthorizedOpenId, [string]$EffectiveMessage, [string]$ChangeId = '') {
    if (-not $ChangeId) { $ChangeId = ([guid]::NewGuid().ToString('N').Substring(0,8)).ToUpperInvariant() }
    $previewEvent = $Event | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $previewEvent.content = $EffectiveMessage
    $oldPreviewMode = [bool]$script:previewMode
    $oldConfirmedExecution = [bool]$script:confirmedExecution
    $script:previewMode = $true
    $script:confirmedExecution = $false
    $script:previewAcknowledgement = ''
    $script:previewNeedsInformation = $false
    $script:previewOperations = [Collections.Generic.List[object]]::new()
    try {
        $handled = Process-Event ($previewEvent | ConvertTo-Json -Compress -Depth 20) $AuthorizedOpenId -BypassConfirmation
        if (-not $handled -or -not $script:previewAcknowledgement) { throw '无法生成待确认预览' }
        $operations = @($script:previewOperations | Where-Object { [string]$_.action -notin @('RefreshLedger','SendDailySummary','SyncPlatformData') })
        $previewText = if ($script:previewNeedsInformation) {
            [string]$script:previewAcknowledgement
        } elseif ($operations.Count -eq 0) {
            [string]$script:previewAcknowledgement
        } else {
            Format-ConfirmationPreview $ChangeId $operations ([string]$script:previewAcknowledgement)
        }
        return [pscustomobject]@{
            change_id = $ChangeId
            conversation_hash = Get-ConversationHash $Event
            original_message = [string]$Event.content
            effective_message = $EffectiveMessage
            original_create_time = [string]$Event.create_time
            created_at = [datetimeoffset]::Now.ToString('o')
            expires_at = [datetimeoffset]::Now.AddHours(48).ToString('o')
            preview = $previewText
            needs_information = [bool]$script:previewNeedsInformation
            mutation_count = $operations.Count
            operations = $operations
            operation_signature = Get-OperationSignature $operations
        }
    } finally {
        $script:previewMode = $oldPreviewMode
        $script:confirmedExecution = $oldConfirmedExecution
        $script:previewAcknowledgement = ''
        $script:previewNeedsInformation = $false
        $script:previewOperations = $null
    }
}

function New-ConfirmationDraftFromCodexResult($Event, [string]$EffectiveMessage, $CodexResult, [string]$ChangeId = '') {
    if (-not $ChangeId) { $ChangeId = ([guid]::NewGuid().ToString('N').Substring(0,8)).ToUpperInvariant() }
    $oldPreviewMode = [bool]$script:previewMode
    $oldConfirmedExecution = [bool]$script:confirmedExecution
    $script:previewMode = $true
    $script:confirmedExecution = $false
    $script:previewOperations = [Collections.Generic.List[object]]::new()
    try {
        $actions = @($CodexResult.actions | Where-Object { [string]$_.action -notin @('RefreshLedger','SendDailySummary','SyncPlatformData') })
        $null = Invoke-CodexActionsInOrder $actions $Event 'preview'
        $operations = @($script:previewOperations)
        $needsInformation = [bool]$CodexResult.needs_information -or $operations.Count -eq 0
        $fallback = Format-NonExecutableCodexReply ([string]$CodexResult.reply)
        $previewText = if ($operations.Count -gt 0) { Format-ConfirmationPreview $ChangeId $operations ([string]$CodexResult.reply) } else { $fallback }
        return [pscustomobject]@{
            change_id = $ChangeId
            conversation_hash = Get-ConversationHash $Event
            original_message = [string]$Event.content
            effective_message = $EffectiveMessage
            original_create_time = [string]$Event.create_time
            created_at = [datetimeoffset]::Now.ToString('o')
            expires_at = [datetimeoffset]::Now.AddHours(48).ToString('o')
            preview = $previewText
            needs_information = $needsInformation
            mutation_count = $operations.Count
            operations = $operations
            operation_signature = Get-OperationSignature $operations
        }
    } finally {
        $script:previewMode = $oldPreviewMode
        $script:confirmedExecution = $oldConfirmedExecution
        $script:previewOperations = $null
    }
}

function Get-PendingRequestSummary($Pending) {
    $summary = [string]$Pending.original_message
    if (-not $summary) { $summary = [string]$Pending.effective_message }
    $summary = ([regex]::Replace($summary, '\s+', ' ')).Trim()
    if ($summary.Length -gt 180) { $summary = $summary.Substring(0,177) + '...' }
    if (-not $summary) { $summary = '未能读取原请求文本' }
    return $summary
}

function Invoke-StoredConfirmationOperations($Pending, $ExecuteEvent, [int]$LockTimeoutSeconds = 2700) {
    $executedOperations = [Collections.Generic.List[object]]::new()
    $storedOperations = @($Pending.operations)
    $tokens = Get-FinanceWorkbookTokens
    if (-not $tokens.ContainsKey('source') -or -not $tokens.ContainsKey('ledger')) { throw '经营事务缺少工作簿映射，未开始写入' }
    $transactionId = "confirm-$($Pending.change_id)-$([datetime]::Now.ToString('yyyyMMddHHmmss'))"
    $transaction = New-FinanceTransaction $stateRoot "飞书确认 $($Pending.change_id)" $transactionId -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli -LockTimeoutSeconds $LockTimeoutSeconds
    $closed = $false
    try {
        $index = 0
        while ($index -lt $storedOperations.Count) {
            $storedOperation = $storedOperations[$index]
            if (Test-BatchCompatibleSkuOperation $storedOperation) {
                $compatibleRun = [Collections.Generic.List[object]]::new()
                $scanIndex = $index
                while ($scanIndex -lt $storedOperations.Count -and (Test-BatchCompatibleSkuOperation $storedOperations[$scanIndex])) {
                    $compatibleRun.Add($storedOperations[$scanIndex])
                    $scanIndex++
                }
                if ($compatibleRun.Count -ge 2) {
                    foreach ($batchResult in @(Invoke-PetstarBatchSkuMaintenance $compatibleRun.ToArray() $ExecuteEvent $index)) {
                        if ($null -eq $batchResult) { throw "批量 SKU 动作从第 $index 项开始没有返回完整结果" }
                        $executedOperations.Add($batchResult)
                    }
                    $index = $scanIndex
                    continue
                }
            }
            $storedAction = [pscustomobject]@{
                action = [string]$storedOperation.action
                arguments = $storedOperation.arguments
            }
            $executedOperation = Invoke-CodexAction $storedAction $ExecuteEvent $index 'execute'
            if ($null -eq $executedOperation) { throw "动作 $($storedAction.action) 没有返回执行结果" }
            $executedOperations.Add($executedOperation)
            $index++
        }
        Complete-FinanceTransaction $transaction
        Close-FinanceTransaction $transaction
        $closed = $true
        return $executedOperations.ToArray()
    } catch {
        $reason = [string]$_.Exception.Message
        $rollback = Invoke-FinanceTransactionRollback $transaction $tokens $larkCli $reason
        if ($rollback.ok) { throw "$reason；本条消息产生的改动已全部回滚" }
        throw "$reason；自动回滚尚未完成，事务 $($rollback.transaction_id) 已阻止后续写入：$($rollback.error)"
    } finally {
        if (-not $closed) { Close-FinanceTransaction $transaction }
    }
}

function Invoke-ConfirmationFlow($Event, [string]$AuthorizedOpenId, [string]$MessageText) {
    $pending = Get-ActivePendingConfirmation $Event
    if (Test-CancellationText $MessageText) {
        if ($pending) {
            $next = Clear-PendingConfirmation
            $reply = "已取消待确认变更：$($pending.change_id)。没有修改表格。"
            if ($next) { $reply += "`n`n下一项：`n$($next.preview)" }
            Send-Acknowledgement $AuthorizedOpenId $reply ([string]$Event.message_id) 'cancel'
        } else {
            Send-Acknowledgement $AuthorizedOpenId '当前没有待确认的经营数据变更。' ([string]$Event.message_id) 'cancel'
        }
        return [pscustomobject]@{handled=$true;result=$true}
    }

    if (Test-ConfirmationText $MessageText) {
        if (-not $pending) {
            Send-Acknowledgement $AuthorizedOpenId '当前没有待确认的经营数据变更。' ([string]$Event.message_id) 'confirm'
            return [pscustomobject]@{handled=$true;result=$true}
        }
        $requestedChangeId = Get-RequestedConfirmationId $MessageText
        if ($requestedChangeId -and $requestedChangeId -ne ([string]$pending.change_id).ToUpperInvariant()) {
            Send-Acknowledgement $AuthorizedOpenId "你指定的是待确认项 $requestedChangeId，但当前待确认项是 $($pending.change_id)。为避免执行错账，本次没有修改表格；请回复“确认 $($pending.change_id)”执行当前项。" ([string]$Event.message_id) 'confirm-mismatch'
            return [pscustomobject]@{handled=$true;result=$true}
        }
        if ($pending.needs_information) {
            Send-Acknowledgement $AuthorizedOpenId "这项还缺关键信息，暂时不能执行。`n请直接补充缺少的内容，我会重新给你确认。" ([string]$Event.message_id) 'confirm'
            return [pscustomobject]@{handled=$true;result=$true}
        }

        $executeEvent = $Event | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        $executeEvent.content = [string]$pending.effective_message
        $executeEvent.create_time = [string]$pending.original_create_time
        $executeEvent.message_id = "confirmed-$($pending.change_id)"
        $oldConfirmedExecution = [bool]$script:confirmedExecution
        $script:confirmedExecution = $true
        $executedOperations = @()
        try {
            Set-InteractivePriority ([string]$pending.change_id)
            # 用户确认的是已经展示过的明确预览。这里直接执行已保存动作，禁止再次调用 Codex 重解释原话。
            # Keep the live message consumer responsive. If a scheduled write is
            # active, hand this durable confirmation to the hidden retry worker.
            $executedOperations = @(Invoke-StoredConfirmationOperations $pending $executeEvent 5)
            if ($executedOperations.Count -ne @($pending.operations).Count) {
                throw "计划执行 $(@($pending.operations).Count) 项，实际仅返回 $($executedOperations.Count) 项结果"
            }
            $successReply = Format-CodexExecutionReply $executedOperations
            # 先准备回执文本，再提交状态并发送持久化回执。飞书暂时不可用也不会导致经营动作重复执行。
            $next = Complete-PendingConfirmation ([string]$pending.change_id)
            Clear-PetstarCodexQueryCache
            if ($next) { $successReply += "`n`n下一项待确认：`n" + [string]$next.preview }
            $null = Send-Acknowledgement $AuthorizedOpenId $successReply ([string]$Event.message_id) 'confirmed'
            Write-BridgeLog "pending confirmation executed change_id=$($pending.change_id)"
            return [pscustomobject]@{handled=$true;result=$true}
        } catch {
            $failureReason = [string]$_.Exception.Message
            if ($failureReason -eq '另一个经营数据事务仍在执行，本次未写入') {
                $null = Start-PendingConfirmationRetryWorker ([string]$pending.change_id)
                $queuedReply = "⏳ 已进入自动排队｜$($pending.change_id)`n当前同步结束后会自动执行，无需再次确认。"
                $null = Send-Acknowledgement $AuthorizedOpenId $queuedReply ([string]$Event.message_id) 'confirmation-queued'
                Write-BridgeLog "pending confirmation queued change_id=$($pending.change_id)"
                return [pscustomobject]@{handled=$true;result=$true}
            }
            $requestSummary = Get-PendingRequestSummary $pending
            $failureReply = @(
                "❌ 待确认项 $($pending.change_id) 未执行成功"
                "原请求：$requestSummary"
                '失败阶段：执行已确认修改'
                "原因：$failureReason"
                '记账状态：本条消息不会保留半成品；成功回滚后所有表格均恢复到执行前。'
                '当前状态：该项仍保留在待确认队列。回复“确认”可整条重试，回复“取消”可放弃。'
            ) -join "`n"
            $null = Send-Acknowledgement $AuthorizedOpenId $failureReply ([string]$Event.message_id) 'confirmation-error'
            Write-BridgeLog "pending confirmation execution failed change_id=$($pending.change_id) atomic_rollback=true error=$failureReason"
            return [pscustomobject]@{handled=$true;result=$true}
        } finally {
            Clear-InteractivePriority ([string]$pending.change_id)
            $script:confirmedExecution = $oldConfirmedExecution
        }
    }

    if ($pending -and (Test-PendingAmendment $MessageText $pending)) {
        $effectiveMessage = "$MessageText；原待确认请求：$($pending.effective_message)"
        $amendEvent = $Event | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        $amendEvent.content = $effectiveMessage
        $amendEvent.create_time = [string]$pending.original_create_time
        $refreshed = New-ConfirmationDraft $amendEvent $AuthorizedOpenId $effectiveMessage ([string]$pending.change_id)
        $refreshed.original_message = [string]$pending.original_message
        Write-PendingConfirmation $refreshed
        Send-Acknowledgement $AuthorizedOpenId ([string]$refreshed.preview) ([string]$Event.message_id) 'amend'
        return [pscustomobject]@{handled=$true;result=$true}
    }

    if (Test-DataMutationMessage $MessageText) {
        $draft = New-ConfirmationDraft $Event $AuthorizedOpenId $MessageText
        if ([int]$draft.mutation_count -eq 0) {
            if ($draft.needs_information) {
                Send-Acknowledgement $AuthorizedOpenId ([string]$draft.preview) ([string]$Event.message_id) 'clarify'
                return [pscustomobject]@{handled=$true;result=$true}
            }
            return [pscustomobject]@{handled=$true;result=(Invoke-CodexAndReply $Event $AuthorizedOpenId $MessageText)}
        }
        $queued = Add-PendingConfirmation $draft
        if ($queued.duplicate) {
            $existing = $queued.existing
            $duplicateReply = "相同修改已经在待确认队列中，没有重复新增。`n`n$([string]$existing.preview)"
            Send-Acknowledgement $AuthorizedOpenId $duplicateReply ([string]$Event.message_id) 'preview-duplicate'
            Write-BridgeLog "pending confirmation deduplicated requested_change_id=$($draft.change_id) existing_change_id=$($existing.change_id)"
        } else {
            Send-Acknowledgement $AuthorizedOpenId (Format-QueuedConfirmationReply $draft $queued.position $queued.active) ([string]$Event.message_id) 'preview'
            Write-BridgeLog "pending confirmation created change_id=$($draft.change_id)"
        }
        return [pscustomobject]@{handled=$true;result=$true}
    }

    return [pscustomobject]@{handled=$false;result=$false}
}

function Invoke-PendingConfirmationRetry([string]$AuthorizedOpenId, [string]$ChangeId) {
    $retryMutex = [Threading.Mutex]::new($false, 'Local\PetstarPendingConfirmationRetry')
    $retryLockTaken = $false
    try {
        try { $retryLockTaken = $retryMutex.WaitOne(0) }
        catch [Threading.AbandonedMutexException] { $retryLockTaken = $true }
        if (-not $retryLockTaken) { return [ordered]@{ok=$true;skipped='worker_running';change_id=$ChangeId} }

        $pending = Read-PendingConfirmation
        if (-not $pending) { return [ordered]@{ok=$true;skipped='no_pending_confirmation';change_id=$ChangeId} }
        if ($ChangeId -and [string]$pending.change_id -ne $ChangeId) {
            return [ordered]@{ok=$true;skipped='pending_item_changed';change_id=$ChangeId;active_change_id=[string]$pending.change_id}
        }

        $executeEvent = [pscustomobject]@{
            content = [string]$pending.effective_message
            create_time = [string]$pending.original_create_time
            message_id = "confirmed-retry-$($pending.change_id)"
            sender_id = $AuthorizedOpenId
            sender_type = 'user'
            chat_type = 'p2p'
            message_type = 'text'
        }
        $oldConfirmedExecution = [bool]$script:confirmedExecution
        $script:confirmedExecution = $true
        try {
            Set-InteractivePriority ([string]$pending.change_id)
            $executedOperations = @(Invoke-StoredConfirmationOperations $pending $executeEvent 2700)
            if ($executedOperations.Count -ne @($pending.operations).Count) {
                throw "计划执行 $(@($pending.operations).Count) 项，实际仅返回 $($executedOperations.Count) 项结果"
            }
            $successReply = Format-CodexExecutionReply $executedOperations
            $next = Complete-PendingConfirmation ([string]$pending.change_id)
            Clear-PetstarCodexQueryCache
            if ($next) { $successReply += "`n`n下一项待确认：`n" + [string]$next.preview }
            $null = Send-Acknowledgement $AuthorizedOpenId $successReply ([string]$executeEvent.message_id) 'confirmed-retry'
            Write-BridgeLog "pending confirmation retry executed change_id=$($pending.change_id)"
            return [ordered]@{ok=$true;change_id=[string]$pending.change_id;operation_count=$executedOperations.Count}
        } catch {
            $reason = [string]$_.Exception.Message
            $requestSummary = Get-PendingRequestSummary $pending
            $failureReply = @(
                "❌ 排队执行仍未成功｜$($pending.change_id)"
                "原请求：$requestSummary"
                "原因：$reason"
                '记账状态：没有保留半成品；该项仍在待确认队列。'
            ) -join "`n"
            $null = Send-Acknowledgement $AuthorizedOpenId $failureReply ([string]$executeEvent.message_id) 'confirmation-retry-error'
            Write-BridgeLog "pending confirmation retry failed change_id=$($pending.change_id) error=$reason"
            return [ordered]@{ok=$false;change_id=[string]$pending.change_id;error=$reason}
        } finally {
            Clear-InteractivePriority ([string]$pending.change_id)
            $script:confirmedExecution = $oldConfirmedExecution
        }
    } finally {
        if ($retryLockTaken) { try { $retryMutex.ReleaseMutex() } catch {} }
        $retryMutex.Dispose()
    }
}

function Process-Event([string]$Line, [string]$AuthorizedOpenId, [switch]$BypassConfirmation) {
    $script:lastProcessFailure = $null
    if (-not $Line.Trim().StartsWith('{')) { return $false }
    try { $event = $Line | ConvertFrom-Json } catch { return $false }
    if ($event.sender_type -ne 'user' -or $event.sender_id -ne $AuthorizedOpenId) { return $false }
    if ($event.message_type -ne 'text' -or $event.chat_type -ne 'p2p') { return $false }

    $messageText = [string]$event.content
    if (-not $BypassConfirmation -and -not $script:previewMode -and -not $script:confirmedExecution) {
        $confirmationResult = Invoke-ConfirmationFlow $event $AuthorizedOpenId $messageText
        if ($confirmationResult.handled) { return [bool]$confirmationResult.result }
    }

    # 普通自然语言一律先交给本机 Codex。下方旧规则仅保留为可回滚实现，默认不再参与意图判断。
    # 确认、取消和待确认内容修改仍由上面的确定性确认层处理；Codex 只生成白名单动作，不能直接写表。
    if ($script:codexFirstRouting) {
        return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText)
    }

    $orderSegments = @(Get-OrderSegments $messageText)
    $orderNumbers = @($orderSegments | ForEach-Object { $_.order_no })
    $isIntervention = $messageText -match '补单|干预(?:订)?单|刷单'
    $mentionsDelivery = $messageText -match '配送费|运费|外部配送|外部骑手|跑腿费|跑腿|骑手|闪送|顺丰同城|达达|UU跑腿|自己送|自己配送|自行配送|自配送|商家自配'
    $isSelfDelivery = Test-SelfDeliveryWithoutFee $messageText
    $deliveryFee = Get-DeliveryFee $messageText
    $interventionPerson = Get-InterventionPerson $messageText
    $inventoryRestored = Test-InventoryRestored $messageText
    $isBatchDateMessage = $messageText -match '生产日期|生产日|保质期'
    $batchDateFacts = if ($isBatchDateMessage) { Get-BatchDateFacts $messageText } else { $null }
    $isExpiryMessage = -not $isBatchDateMessage -and $messageText -match '过期|到期|有效期'
    $isLedgerSyncRequest = $messageText -match '(?:同步|刷新|更新|重建).{0,10}(?:支出|简账)|(?:支出|简账).{0,10}(?:同步|刷新|更新|重建)'
    $offlineFacts = Get-OfflineSaleFacts $messageText
    $giftFacts = Get-GiftFacts $messageText
    $skuCodeChange = Get-SkuCodeChangeFacts $messageText

    if ($isLedgerSyncRequest) {
        $ledgerResult = Invoke-DailyLedgerSync (Resolve-MessageDate $messageText ([string]$event.create_time))
        $monthLines = @($ledgerResult.months | ForEach-Object { "$([int]([string]$_.month).Substring(5,2))月支出 $($_.expense_rows) 行" })
        $reply = "支出已同步。`n" + ($monthLines -join '｜')
        Send-Acknowledgement $AuthorizedOpenId $reply ([string]$event.message_id)
        Write-BridgeLog "ledger sync processed message_hash=$((Get-MessageHash ([string]$event.message_id)).Substring(0,12))"
        return $true
    }

    if ($null -ne $skuCodeChange) {
        $changeArgs = @('--action','ChangeSkuCode','--goods-code',[string]$skuCodeChange.old_code,'--new-goods-code',[string]$skuCodeChange.new_code,'--note','用户通过自然语言调整SKU货号','--event-id',"$($event.message_id)-sku-code-change")
        if ($messageText -match '(?:原|之前|以前).{0,12}(?:商品|试吃装|尝鲜装)?.{0,12}(?:已|已经)?(?:下架|停用|不再使用|不用了)|(?:下架|停用).{0,12}(?:原|旧).{0,8}(?:商品|货号)') {
            $changeArgs += '--retire-existing-new-code'
        }
        $result = Invoke-PetstarEvent $changeArgs
        $details = $result.result
        if ($details.status -eq 'conflict') {
            $oldTitle = [string]$details.old_title
            $newTitle = [string]$details.new_title
            $reply = "暂未修改：$($details.new_code) 当前仍分配给另一件商品。`n$($details.old_code)：$oldTitle`n$($details.new_code)：$newTitle`n如果原商品已下架或移走，直接告诉我即可完成改绑。"
        } elseif ($details.status -eq 'not_found') {
            $reply = "没有修改：SKU 成本主表中找不到原货号 $($details.old_code)。`n请核对原 KW，或先等下一次美团库存同步完成建档。"
        } elseif ($details.status -eq 'unchanged') {
            $reply = "无需重复修改：$($details.old_code) 已经对应到 $($details.new_code)。"
        } else {
            $retiredText = if ($details.retired_title) { "`n原先占用 $($details.new_code) 的商品已解除该货号。" } else { '' }
            $reply = "已更新当前 KW：$($details.old_code) → $($details.new_code)$retiredText`n旧 KW 没有保存成别名；历史订单仍保留成交时快照。"
        }
        Send-Acknowledgement $AuthorizedOpenId $reply ([string]$event.message_id)
        Write-BridgeLog "sku code change processed message_hash=$((Get-MessageHash ([string]$event.message_id)).Substring(0,12)) status=$($details.status)"
        return $true
    }

    if ($isBatchDateMessage) {
        if (-not $batchDateFacts.complete) { return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText) }
        $batchArgs = @(
            '--action','SetSkuBatchDate','--goods-code',[string]$batchDateFacts.goods_code,
            '--production-date',$batchDateFacts.production_date.ToString('yyyy-MM-dd'),
            '--shelf-life-value',[string]$batchDateFacts.shelf_life_value,
            '--shelf-life-unit',[string]$batchDateFacts.shelf_life_unit,
            '--event-id',"$($event.message_id)-sku-batch-date-$($batchDateFacts.goods_code)"
        )
        $result = Invoke-PetstarEvent $batchArgs
        $details = $result.result
        $reply = "已更新：$($details.goods_code)`n生产日期：$($details.production_date)｜保质期：$($details.shelf_life)`n推算到期：$($details.expiry_date)"
        Send-Acknowledgement $AuthorizedOpenId $reply ([string]$event.message_id)
        Write-BridgeLog "sku batch date processed message_hash=$((Get-MessageHash ([string]$event.message_id)).Substring(0,12)) code=$($details.goods_code)"
        return $true
    }

    if ($isExpiryMessage) {
        $skuCodes = @(Get-SkuCodes $messageText)
        $clearExpiry = Test-ClearExpiry $messageText
        $expiryDate = Get-ExplicitExpiryDate $messageText
        if ($skuCodes.Count -ne 1 -or (-not $clearExpiry -and $null -eq $expiryDate)) {
            return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText)
        }
        $eventArgs = @('--action','SetSkuExpiry','--goods-code',$skuCodes[0],'--event-id',"$($event.message_id)-sku-expiry-$($skuCodes[0])")
        if ($clearExpiry) {
            $eventArgs += '--clear-expiry'
        } else {
            $eventArgs += @('--expiry-date',$expiryDate.ToString('yyyy-MM-dd'))
        }
        $result = Invoke-PetstarEvent $eventArgs
        $displayCode = [string]$result.result.goods_code
        $expiryText = [string]$result.result.earliest_expiry
        $reply = if ($clearExpiry) {
            "已更新：$displayCode`n最早过期日期已标记为无。"
        } else {
            "已更新：$displayCode`n当前库存最早过期：$expiryText"
        }
        Send-Acknowledgement $AuthorizedOpenId $reply ([string]$event.message_id)
        Write-BridgeLog "sku expiry processed message_hash=$((Get-MessageHash ([string]$event.message_id)).Substring(0,12)) code=$displayCode"
        return $true
    }

    # 完整、确定的经营事实走快速路径；有歧义或需要上下文的表达交给本机 Codex 理解。
    if ($isIntervention) {
        if ($orderNumbers.Count -eq 0) {
            return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText)
        }

        # 多笔订单必须逐段提取配送方式和费用，不能把整段文字交给模型重新抄写订单号。
        $collectiveSelfDelivery = $messageText -match '(?:都|全部|全都|这几单|这些单|以上订单).{0,10}(?:自己送|自己配送|自行配送|自配送|商家自配)'
        $collectiveInventoryRestored = $messageText -match '(?:都|全部|全都|这几单|这些单|以上订单).{0,10}(?:已|已经)?(?:恢复|改回|还原).{0,8}(?:库存|货品|商品)'
        $segmentFacts = [Collections.Generic.List[object]]::new()
        foreach ($segment in $orderSegments) {
            $segmentText = [string]$segment.text
            $segmentMentionsDelivery = $segmentText -match '配送费|运费|外部配送|外部骑手|跑腿费|跑腿|骑手|闪送|顺丰同城|达达|UU跑腿|自己送|自己配送|自行配送|自配送|商家自配'
            $segmentSelfDelivery = $collectiveSelfDelivery -or (Test-SelfDeliveryWithoutFee $segmentText)
            $segmentDeliveryFee = Get-DeliveryFee $segmentText
            if ($segmentMentionsDelivery -and -not $segmentSelfDelivery -and $null -eq $segmentDeliveryFee) {
                return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText)
            }
            $segmentPerson = Get-InterventionPerson $segmentText
            if (-not $segmentPerson -and $orderNumbers.Count -eq 1) { $segmentPerson = $interventionPerson }
            $segmentFacts.Add([pscustomobject]@{
                order_no = [string]$segment.order_no
                person = $segmentPerson
                self_delivery = $segmentSelfDelivery
                delivery_fee = $segmentDeliveryFee
                inventory_restored = $collectiveInventoryRestored -or (Test-InventoryRestored $segmentText)
            })
        }

        $eventDate = Resolve-MessageDate $messageText ([string]$event.create_time)
        $registered = @()
        $duplicates = @()
        $processed = [Collections.Generic.List[object]]::new()
        foreach ($fact in $segmentFacts) {
            $orderNo = [string]$fact.order_no
            $addArgs = @('--action','AddIntervention','--date',$eventDate.ToString('yyyy-MM-dd'),'--order-no',$orderNo,'--note','用户通过自然语言登记为干预单','--event-id',"$($event.message_id)-intervention-$orderNo")
            if ($fact.person) { $addArgs += @('--person',[string]$fact.person) }
            $addResult = Invoke-PetstarEvent $addArgs
            if ($addResult.duplicate) { $duplicates += $orderNo } else { $registered += $orderNo }

            if ($fact.self_delivery) {
                $null = Invoke-PetstarEvent @('--action','ConfirmSelfDelivery','--order-no',$orderNo,'--note','用户确认自己配送且配送费为0','--event-id',"$($event.message_id)-self-delivery-$orderNo")
            } elseif ($null -ne $fact.delivery_fee) {
                $null = Invoke-PetstarEvent @('--action','SetExternalDeliveryFee','--order-no',$orderNo,'--amount',[string]$fact.delivery_fee,'--note','用户通过自然语言记录配送费','--event-id',"$($event.message_id)-delivery-$orderNo")
            }
            if ($fact.inventory_restored) {
                $null = Invoke-PetstarEvent @('--action','RestoreInventory','--order-no',$orderNo,'--note','用户通过自然语言确认库存已恢复','--event-id',"$($event.message_id)-restore-$orderNo")
            }
            $processed.Add($fact)
        }
        if ($processed.Count -ne $orderNumbers.Count) { throw "多订单登记数量不一致：识别 $($orderNumbers.Count) 单，实际处理 $($processed.Count) 单" }

        $offlineResult = $null
        if ($null -ne $offlineFacts) {
            if (-not $offlineFacts.complete) { return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText) }
            $offlineResult = Invoke-OfflineSaleEvent $event $eventDate $offlineFacts
        }

        $giftResult = $null
        if ($null -ne $giftFacts) {
            if (-not $giftFacts.complete -or $orderNumbers.Count -ne 1) { return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText) }
            $giftResult = Invoke-GiftEvent $event $eventDate $giftFacts
        }

        $lines = [Collections.Generic.List[string]]::new()
        if ($null -ne $offlineResult) {
            $lines.Add("已经记好了：$($processed.Count) 笔干预单、1 笔线下单。")
        } else {
            $lines.Add("已经记好了：$($processed.Count) 笔干预单。")
        }
        if ($processed.Count -eq 1 -and $registered.Count -gt 0) { $lines.Add("订单号：$($registered[0])") }
        if ($duplicates.Count -gt 0) { $lines.Add("已登记过：$($duplicates -join '、')（没有重复新增）") }
        foreach ($fact in $processed) {
            $orderText = [string]$fact.order_no
            $shortOrderNo = $orderText.Substring($orderText.Length - 4)
            $deliverySummary = if ($fact.self_delivery) { '自配送 ¥0' } elseif ($null -ne $fact.delivery_fee) { ('骑手配送 ¥{0:N2}' -f [double]$fact.delivery_fee) } else { '配送费待确认' }
            $restoreSummary = if ($fact.inventory_restored) { '库存已恢复' } else { '库存待恢复' }
            $lines.Add("尾号 $shortOrderNo｜$deliverySummary｜$restoreSummary")
        }
        if ($null -ne $giftResult) {
            $giftDetails = $giftResult.result
            $giftCostText = if ($null -ne $giftDetails.gift_cost) { ('¥{0:N2}' -f [double]$giftDetails.gift_cost) } else { '待核对' }
            $lines.Add("赠品：$($giftDetails.gift_items)｜成本 $giftCostText（营销赠品支出）")
            $lines.Add('请同步扣减赠品库存。')
        }
        if ($null -ne $offlineResult) {
            $offlineDetails = $offlineResult.result
            $profitText = if ($null -ne $offlineDetails.product_cost) { '｜赚 ¥{0:N2}' -f ([double]$offlineDetails.income - [double]$offlineDetails.product_cost - [double]$offlineDetails.delivery_fee) } else { '｜成本待核对' }
            $lines.Add("线下：$($offlineDetails.sold_goods)｜实收 ¥$([double]$offlineDetails.income)$profitText")
        }
        $lines.Add('以上订单已从自然单收入中排除。')
        Send-Acknowledgement $AuthorizedOpenId ($lines -join "`n") ([string]$event.message_id)
        Write-BridgeLog "intervention processed message_hash=$((Get-MessageHash ([string]$event.message_id)).Substring(0,12)) orders=$($orderNumbers.Count)"
        return $true
    }

    if ($null -ne $giftFacts) {
        if (-not $giftFacts.complete) { return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText) }
        $eventDate = Resolve-MessageDate $messageText ([string]$event.create_time)
        $giftResult = Invoke-GiftEvent $event $eventDate $giftFacts
        $details = $giftResult.result
        $costText = if ($null -ne $details.gift_cost) { ('¥{0:N2}' -f [double]$details.gift_cost) } else { '待核对' }
        $accountingText = if ($details.accounting -eq '营销赠品支出') { '已记入营销赠品支出' } else { '已计入本单货物成本' }
        $reply = if ($details.status -eq 'pending_order') {
            "已暂存：订单 $($details.order_no)赠品 $($details.gift_items)`n订单同步后会自动匹配并计入成本。请手动扣减对应库存。"
        } else {
            "已记好：$($details.daily_order_no)赠品 $($details.gift_items)`n赠品成本 $costText，$accountingText。`n请手动扣减对应库存。"
        }
        if (@($details.pending_cost_codes).Count -gt 0) {
            $reply += "`n成本待核对：$(@($details.pending_cost_codes) -join '、')"
        }
        Send-Acknowledgement $AuthorizedOpenId $reply ([string]$event.message_id)
        Write-BridgeLog "gift processed message_hash=$((Get-MessageHash ([string]$event.message_id)).Substring(0,12)) order=$($details.order_no)"
        return $true
    }

    if ($null -ne $offlineFacts) {
        if (-not $offlineFacts.complete) { return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText) }
        $eventDate = Resolve-MessageDate $messageText ([string]$event.create_time)
        $offlineResult = Invoke-OfflineSaleEvent $event $eventDate $offlineFacts
        $details = $offlineResult.result
        $profitText = if ($null -ne $details.product_cost) { '赚 ¥{0:N2}' -f ([double]$details.income - [double]$details.product_cost - [double]$details.delivery_fee) } else { '成本待核对' }
        $reply = "已记好：$($details.daily_order_no)`n实收 ¥$([double]$details.income)｜$profitText`n商品：$($details.sold_goods)`n请手动调整美团库存。"
        Send-Acknowledgement $AuthorizedOpenId $reply ([string]$event.message_id)
        Write-BridgeLog "offline sale processed message_hash=$((Get-MessageHash ([string]$event.message_id)).Substring(0,12))"
        return $true
    }

    if ($mentionsDelivery) {
        if ($orderNumbers.Count -ne 1 -or (-not $isSelfDelivery -and $null -eq $deliveryFee)) {
            return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText)
        }
        $deliveryOrderNo = $orderNumbers[0]
        try {
            $resolvedFee = if ($isSelfDelivery) { 0.0 } else { [double]$deliveryFee }
            $result = Invoke-PetstarEvent @('--action','SetOrderDeliveryFee','--order-no',$deliveryOrderNo,'--amount',[string]$resolvedFee,'--note','用户通过自然语言确认订单配送费','--event-id',"$($event.message_id)-order-delivery-$deliveryOrderNo")
        } catch {
            return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText)
        }
        $details = $result.result
        $deliveryAmount = [double]$details.delivery_fee
        $deliveryText = if ($isSelfDelivery) { '自配送，配送费 0 元' } else { ('骑手配送，配送费 {0:N2} 元' -f $deliveryAmount) }
        $profitText = if ($details.order_type -eq '自然单') { '对应收入页和当天赚/亏已经重算。' } else { '对应月份的补单配送费支出已经同步。' }
        $reply = @(
            '已经记好了。',
            "订单号：$deliveryOrderNo",
            "配送：$deliveryText",
            $profitText
        ) -join [Environment]::NewLine
        Send-Acknowledgement $AuthorizedOpenId $reply ([string]$event.message_id)
        Write-BridgeLog "delivery processed message_hash=$((Get-MessageHash ([string]$event.message_id)).Substring(0,12))"
        return $true
    }

    return (Invoke-CodexAndReply $event $AuthorizedOpenId $messageText)
}

function Process-EventWithLedger([string]$Line, [string]$AuthorizedOpenId, [string]$Source) {
    $eventStartedAt = [datetime]::UtcNow
    if (-not $Line.Trim().StartsWith('{')) { return 'ignored' }
    try { $event = $Line | ConvertFrom-Json } catch { return 'ignored' }
    $messageId = [string]$event.message_id
    if (-not $messageId) { return 'ignored' }
    $hash = Get-MessageHash $messageId
    $messageMutex = [Threading.Mutex]::new($false, "Local\PetstarLarkMessage_$($hash.Substring(0,32))")
    $lockTaken = $false
    $processingReactionId = $null
    try {
        try {
            $lockTaken = $messageMutex.WaitOne(0)
        } catch [Threading.AbandonedMutexException] {
            $lockTaken = $true
        }
        if (-not $lockTaken) { return 'locked' }
        if (Test-MessageHandled $hash) { return 'duplicate' }
        if ($event.sender_type -eq 'user' -and $event.sender_id -eq $AuthorizedOpenId -and $event.message_type -eq 'text' -and $event.chat_type -eq 'p2p') {
            $processingReactionId = Add-ProcessingReaction $messageId
        }
        $handled = Process-Event $Line $AuthorizedOpenId
        if (-not $handled) {
            if ($script:lastProcessFailure) {
                # Codex/网络临时失败不能永久吞掉原消息；回执使用固定幂等键，后台重放不会刷屏。
                Add-HandledMessageRecord $messageId 'failed_retryable' $Source ([string]$event.create_time)
                return 'failed_retryable'
            }
            return 'ignored'
        }
        Add-HandledMessageRecord $messageId 'completed' $Source ([string]$event.create_time)
        return 'completed'
    } catch {
        $errorText = [string]$_.Exception.Message
        Write-BridgeLog "message processing failed message_hash=$($hash.Substring(0,12)) error=$errorText"
        try {
            $messageSummary = ([regex]::Replace([string]$event.content, '\s+', ' ')).Trim()
            if ($messageSummary.Length -gt 160) { $messageSummary = $messageSummary.Substring(0,157) + '...' }
            $failureReply = @(
                '❌ 消息未处理成功'
                "原消息：$messageSummary"
                '失败阶段：消息解析或经营事件预览'
                "原因：$errorText"
                '处理状态：没有标记为完成；请按原因补充信息，或直接重发原话。'
            ) -join "`n"
            Send-Acknowledgement $AuthorizedOpenId $failureReply $messageId 'error'
            Add-HandledMessageRecord $messageId 'failed_retryable' $Source ([string]$event.create_time)
            return 'failed_retryable'
        } catch {
            Write-BridgeLog "failure reply could not be sent message_hash=$($hash.Substring(0,12))"
            Add-HandledMessageRecord $messageId 'failed_unreplied' $Source ([string]$event.create_time)
            return 'failed_unreplied'
        }
    } finally {
        Write-BridgeLog "message timing source=$Source total=$([math]::Round(([datetime]::UtcNow-$eventStartedAt).TotalSeconds,3))s message_hash=$($hash.Substring(0,12))"
        Remove-ProcessingReaction $messageId $processingReactionId
        if ($lockTaken) { try { $messageMutex.ReleaseMutex() } catch {} }
        $messageMutex.Dispose()
    }
}

function Initialize-ReplayLedger([string]$AuthorizedOpenId) {
    $messages = Get-EligibleReplayMessages (Get-RecentReplayMessages) $AuthorizedOpenId
    $added = 0
    $pending = 0
    foreach ($message in $messages) {
        $hash = Get-MessageHash ([string]$message.message_id)
        if ($script:handledMessageHashes.Contains($hash)) { continue }
        # 首次建立账本时，普通聊天可作为基线；经营指令必须留给 Replay 处理，不能静默吞掉。
        if (Test-ActionableBusinessMessage ([string]$message.content)) {
            $pending++
            continue
        }
        Add-HandledMessageRecord ([string]$message.message_id) 'baseline' 'initialization' ([string]$message.create_time)
        $added++
    }
    return [ordered]@{ scanned = $messages.Count; added = $added; pending = $pending }
}

function Inspect-ReplayLedger([string]$AuthorizedOpenId) {
    $messages = Get-EligibleReplayMessages (Get-RecentReplayMessages) $AuthorizedOpenId
    $pending = @($messages | Where-Object {
        -not $script:handledMessageHashes.Contains((Get-MessageHash ([string]$_.message_id)))
    })
    return [ordered]@{ scanned = $messages.Count; pending = $pending.Count; handled = $messages.Count - $pending.Count }
}

function Invoke-RecentMessageReplay([string]$AuthorizedOpenId) {
    $messages = Get-EligibleReplayMessages (Get-RecentReplayMessages) $AuthorizedOpenId
    $completed = 0
    $duplicates = 0
    foreach ($message in $messages) {
        $event = Convert-ReplayMessageToEvent $message
        $result = Process-EventWithLedger ($event | ConvertTo-Json -Compress) $AuthorizedOpenId 'replay'
        if ($result -eq 'completed') { $completed++ }
        if ($result -eq 'duplicate') { $duplicates++ }
    }
    return [ordered]@{ scanned = $messages.Count; completed = $completed; duplicates = $duplicates }
}

$script:handledMessageHashes = Get-HandledMessageHashes
$authorizedOpenId = Get-CurrentUserOpenId

if ($Mode -eq 'RetryPending') {
    $result = Invoke-PendingConfirmationRetry $authorizedOpenId $RetryChangeId
    [ordered]@{ok=[bool]$result.ok;mode=$Mode;result=$result} | ConvertTo-Json -Depth 8
    if (-not $result.ok) { exit 1 }
    exit 0
}

if ($Mode -eq 'FlushAcks') {
    try {
        $result = Flush-PendingAcknowledgements
        [ordered]@{ok=($result.failed -eq 0);mode=$Mode;result=$result} | ConvertTo-Json -Compress
        if ($result.failed -gt 0) { exit 1 }
        exit 0
    } catch {
        [ordered]@{ok=$false;mode=$Mode;error=$_.Exception.Message} | ConvertTo-Json -Compress
        exit 1
    }
}

if ($Mode -eq 'Initialize') {
    $result = Initialize-ReplayLedger $authorizedOpenId
    [ordered]@{ ok = $true; mode = $Mode; result = $result } | ConvertTo-Json -Compress
    exit 0
}

if ($Mode -eq 'Inspect') {
    $result = Inspect-ReplayLedger $authorizedOpenId
    [ordered]@{ ok = $true; mode = $Mode; result = $result } | ConvertTo-Json -Compress
    exit 0
}

if ($Mode -eq 'Replay') {
    $replayMutex = [Threading.Mutex]::new($false, 'Local\PetstarLarkReplayWorker')
    $replayLockTaken = $false
    try {
        try {
            $replayLockTaken = $replayMutex.WaitOne(0)
        } catch [Threading.AbandonedMutexException] {
            $replayLockTaken = $true
        }
        if (-not $replayLockTaken) {
            [ordered]@{ ok = $true; mode = $Mode; skipped = 'worker_running' } | ConvertTo-Json -Compress
            exit 0
        }
        $result = Invoke-RecentMessageReplay $authorizedOpenId
        Write-BridgeLog "replay completed scanned=$($result.scanned) completed=$($result.completed) duplicates=$($result.duplicates)"
        [ordered]@{ ok = $true; mode = $Mode; result = $result } | ConvertTo-Json -Compress
        exit 0
    } catch {
        Write-BridgeLog "replay failed error=$($_.Exception.Message)"
        [ordered]@{ ok = $false; mode = $Mode; error = $_.Exception.Message } | ConvertTo-Json -Compress
        exit 1
    } finally {
        if ($replayLockTaken) { try { $replayMutex.ReleaseMutex() } catch {} }
        $replayMutex.Dispose()
    }
}

$mutex = [Threading.Mutex]::new($false, 'Local\PetstarLarkEventBridge')
if (-not $mutex.WaitOne(0)) { exit 0 }

try {
    Write-BridgeLog "bridge started pid=$PID"
    while ($true) {
        try {
            $authorizedOpenId = Get-CurrentUserOpenId
            if (-not (Test-Path -LiteralPath $handledMessagesLog)) {
                $baseline = Initialize-ReplayLedger $authorizedOpenId
                Write-BridgeLog "replay ledger initialized scanned=$($baseline.scanned) added=$($baseline.added)"
            }
            $ackFlush = Flush-PendingAcknowledgements
            if ($ackFlush.queued -gt 0) {
                Write-BridgeLog "acknowledgement outbox flushed queued=$($ackFlush.queued) sent=$($ackFlush.sent) failed=$($ackFlush.failed)"
            }
            $replayArguments = @(
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath),
                '-ProjectRoot', ('"{0}"' -f $ProjectRoot), '-Mode', 'Replay',
                '-ReplayHours', [string]$ReplayHours, '-ChatName', ('"{0}"' -f $ChatName)
            )
            $replayProcess = Start-Process -FilePath $powershellExe -ArgumentList $replayArguments -WorkingDirectory $ProjectRoot -WindowStyle Hidden -PassThru
            Write-BridgeLog "replay worker launched pid=$($replayProcess.Id)"
            $previousErrorAction = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            # 无限长连接；只在真实断线时重连，消除每10分钟主动退出造成的接收空窗。
            & $larkCli event consume im.message.receive_v1 --as bot --timeout 0 --quiet 2>> $consumerLog | ForEach-Object {
                $null = Process-EventWithLedger ([string]$_) $authorizedOpenId 'live'
            }
            $ErrorActionPreference = $previousErrorAction
            Write-BridgeLog "event consumer exited code=$LASTEXITCODE; restarting"
        } catch {
            $ErrorActionPreference = 'Stop'
            Write-BridgeLog "event consumer error=$($_.Exception.Message)"
        }
        Start-Sleep -Seconds 5
    }
} finally {
    Write-BridgeLog 'bridge stopped'
    try { $mutex.ReleaseMutex() } catch {}
    $mutex.Dispose()
}
