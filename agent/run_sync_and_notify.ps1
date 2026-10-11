param(
    [datetime]$Date = [datetime]::Today,
    [ValidateSet('midday','reminder','night','manual')][string]$Mode = 'manual',
    [switch]$SkipSync,
    [switch]$SkipOrders,
    [switch]$SkipInventory,
    [switch]$Silent,
    [switch]$DryRun,
    [switch]$ForceResend,
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
# 定时任务和 Hermes 子进程必须使用同一个 CLI 与凭据目录，避免 PATH 命中不同安装。
$larkCliDirectory = Join-Path $env:APPDATA 'npm'
$larkCliCommand = Join-Path $larkCliDirectory 'lark-cli.cmd'
if (-not (Test-Path -LiteralPath $larkCliCommand -PathType Leaf)) { throw "找不到飞书 CLI：$larkCliCommand" }
$env:PATH = "$larkCliDirectory;$env:PATH"
# Hermes Gateway 会继承系统代理；飞书 CLI 的 OAuth 刷新不经过该代理，降低本地代理波动导致的瞬时失败。
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$projectRoot = $ProjectRoot
$workRoot = Join-Path $projectRoot 'work'
$stateRoot = Join-Path $projectRoot 'agent\state'
$metaPath = Join-Path $workRoot 'simple_finance_meta.json'
$orderScript = Join-Path $workRoot 'sync_meituan_orders.ps1'
$inventoryScript = Join-Path $workRoot 'sync_meituan_inventory.ps1'
$ledgerScript = Join-Path $projectRoot 'agent\sync_daily_ledger.ps1'
$costScript = Join-Path $projectRoot 'agent\rebuild_cost_snapshots.ps1'
$giftReconcileScript = Join-Path $projectRoot 'agent\reconcile_pending_gifts.ps1'
$integrityAuditScript = Join-Path $workRoot 'audit_finance_integrity.ps1'
$ledgerMetaPath = Join-Path $workRoot 'daily_ledger_meta.json'
$larkIdentityCachePath = Join-Path $stateRoot 'lark_user_identity.json'
$failureNotificationRoot = Join-Path $stateRoot 'failure_notifications'
$interactivePriorityPath = Join-Path $stateRoot 'interactive_priority.json'
$promotionCaptureScript = Join-Path $projectRoot 'agent\sync_promotion_expense.ps1'
$promotionAutomationStartDate = [datetime]'2026-09-23'
. (Join-Path $PSScriptRoot 'intervention_record.ps1')

$syncMutex = [Threading.Mutex]::new($false, 'Local\PetstarSyncAndNotify')
if (-not $syncMutex.WaitOne(0)) {
    [ordered]@{ok=$true;skipped=$true;reason='已有同步任务正在运行';date=$Date.ToString('yyyy-MM-dd');mode=$Mode;completed_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')} | ConvertTo-Json -Compress
    $syncMutex.Dispose()
    exit 0
}

# Scheduled syncs yield before taking the global finance lock when a confirmed
# Feishu write is starting. This never interrupts an in-flight transaction.
if ($Mode -ne 'manual' -and (Test-Path -LiteralPath $interactivePriorityPath)) {
    $priorityIsActive = $false
    try {
        $priorityState = Get-Content -LiteralPath $interactivePriorityPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $priorityStartedAt = [datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse([string]$priorityState.started_at,[ref]$priorityStartedAt)) {
            # A queued confirmation may wait behind one full rebuild and then run
            # its own atomic recalculation. Keep later scheduled jobs out of the
            # way for the whole bounded retry window.
            $priorityIsActive = ([datetimeoffset]::Now - $priorityStartedAt).TotalMinutes -lt 90
        }
    } catch {}
    if ($priorityIsActive) {
        [ordered]@{ok=$true;skipped=$true;reason='人工确认写账优先，定时同步已让路；下一次计划任务会自动补跑';date=$Date.ToString('yyyy-MM-dd');mode=$Mode;completed_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')} | ConvertTo-Json -Compress
        try { $syncMutex.ReleaseMutex() } catch {}
        $syncMutex.Dispose()
        exit 0
    }
}

if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少飞书经营驾驶舱元数据' }
$meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $PSScriptRoot 'finance_transaction.ps1')
. (Join-Path $PSScriptRoot 'sku_master_schema.ps1')
. (Join-Path $PSScriptRoot 'runtime_paths.ps1')
$webAccessCheck = Get-PetstarWebAccessCheck
$skuSchema = Get-SkuMasterSchema -SpreadsheetToken $spreadsheetToken
$dateText = $Date.ToString('yyyy-MM-dd')
$runStartedAt = [datetime]::Now
$stageTelemetry = [ordered]@{}
$script:financeTransaction = $null
$script:financeTransactionClosed = $false
$script:financeDataCommitted = $false

function Get-FinanceWorkbookTokens {
    if (-not (Test-Path -LiteralPath $ledgerMetaPath)) { throw '缺少每日经营简账元数据' }
    $ledgerMeta = [System.IO.File]::ReadAllText($ledgerMetaPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    return @{source=$spreadsheetToken;ledger=([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')}
}
function Start-SyncFinanceTransaction {
    if ($null -ne $script:financeTransaction) { return }
    $tokens = Get-FinanceWorkbookTokens
    $script:financeTransaction = New-FinanceTransaction $stateRoot "定时经营同步 $dateText $Mode" $null -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCliCommand
}
function Complete-SyncFinanceTransaction {
    if ($script:financeTransactionClosed -or $null -eq $script:financeTransaction) { return }
    Complete-FinanceTransaction $script:financeTransaction
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
    $script:financeDataCommitted = $true
}
function Rollback-SyncFinanceTransaction([string]$Reason) {
    if ($script:financeTransactionClosed -or $null -eq $script:financeTransaction) { return $null }
    $rollback = Invoke-FinanceTransactionRollback $script:financeTransaction (Get-FinanceWorkbookTokens) $larkCliCommand $Reason
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
    return $rollback
}

function Invoke-FinanceIntegrityAudit {
    if (-not (Test-Path -LiteralPath $integrityAuditScript -PathType Leaf)) { throw '缺少经营账目一致性审计脚本' }
    $powershellExe = Get-PetstarWindowsPowerShell
    $global:LASTEXITCODE = 0
    $raw = & $powershellExe -NoProfile -ExecutionPolicy Bypass -File $integrityAuditScript -ProjectRoot $projectRoot | Out-String
    if ($LASTEXITCODE -ne 0) { throw "经营账目一致性审计执行失败：$($raw.Trim())" }
    try { $result = $raw | ConvertFrom-Json }
    catch { throw "经营账目一致性审计没有返回有效结果：$($_.Exception.Message)" }
    if (-not $result.ok) {
        $issues = [Collections.Generic.List[string]]::new()
        foreach ($spec in @(
            @('重复源订单','duplicate_source_orders'),@('重复收入订单','duplicate_income_orders'),
            @('收入与源表不一致','income_source_mismatches'),@('干预单混入收入','intervention_orders_in_income'),
            @('利润算术错误','profit_arithmetic_errors'),@('收入排序错误','income_sort_errors'),
            @('自然单缺失','missing_natural_orders'),@('干预支出配对错误','intervention_expense_pair_errors')
        )) {
            $property = $result.PSObject.Properties[$spec[1]]
            $values = if ($null -ne $property) { @($property.Value) } else { @() }
            if ($values.Count -gt 0) { $issues.Add("$($spec[0])：$((@($values | Select-Object -First 5)) -join '、')") }
        }
        $detail = if ($issues.Count -gt 0) { $issues -join '；' } else { '审计结果未提供具体异常' }
        throw "经营账目一致性审计未通过：$detail"
    }
    return $result
}

function Get-PromotionCaptureState {
    if ($Mode -ne 'night' -or $Date.Date -lt $promotionAutomationStartDate.Date) { return $null }
    $path = Join-Path $stateRoot "promotion_expense\$dateText.json"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf) -and
        $Date.Date -eq [datetime]::Today.AddDays(-1)) {
        if (-not (Test-Path -LiteralPath $promotionCaptureScript -PathType Leaf)) {
            throw "缺少推广费补抓脚本：$promotionCaptureScript"
        }
        $powershellExe = Get-PetstarWindowsPowerShell
        $global:LASTEXITCODE = 0
        $recoveryOutput = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $promotionCaptureScript -Date $dateText -ProjectRoot $projectRoot 2>&1 | ForEach-Object { [string]$_ })
        if ($LASTEXITCODE -ne 0) {
            $recoveryDetail = (($recoveryOutput | Where-Object { $_ -and $_.Trim() }) -join ' ').Trim()
            throw "当天推广费凌晨补抓未成功：$dateText；$recoveryDetail；已停止日结，避免漏扣推广费"
        }
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "当天推广费自动读取未成功：$dateText；已停止日结，避免漏扣推广费"
    }
    try { $state = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json }
    catch { throw "当天推广费状态文件无法读取：$dateText；已停止日结，避免错账" }
    if (-not $state.ok -or [string]$state.date -ne $dateText -or $null -eq $state.amount) {
        throw "当天推广费状态未通过校验：$dateText；已停止日结，避免漏扣推广费"
    }
    return $state
}

function Get-TextFingerprint([string]$Value) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $normalized = ([string]$Value).Trim()
        $normalized = $normalized -replace '\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(?::\d{2})?', '<timestamp>'
        $normalized = $normalized -replace '"failed_at"\s*:\s*"[^"]+"', '"failed_at":"<timestamp>"'
        $normalized = $normalized -replace '\.tmp\.\d+', '.tmp.<pid>'
        $bytes = [Text.Encoding]::UTF8.GetBytes($normalized)
        return (([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').Substring(0, 12)).ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Get-FailureNotificationCategory([string]$Reason) {
    if ($Reason -match '推广费.*(?:读取|补抓|状态|抓取)|当天推广费') { return 'promotion-capture' }
    if ($Reason -match 'MEITUAN_STORE_CONTEXT_REQUIRED|MEITUAN_BOUND_PROFILE_UNAVAILABLE|MEITUAN_STORE_IDENTITY_UNVERIFIED') { return 'meituan-context' }
    if ($Reason -match 'MEITUAN_AUTH_REQUIRED|login_gw|美团商家后台登录态已失效') { return 'meituan-auth' }
    return (Get-TextFingerprint $Reason)
}

function Get-FailureNotificationReceiptPath([string]$Category) {
    if (-not (Test-Path -LiteralPath $failureNotificationRoot)) {
        $null = New-Item -ItemType Directory -Path $failureNotificationRoot
    }
    $safeCategory = $Category -replace '[^a-zA-Z0-9_-]', '_'
    if ($Category -in @('meituan-auth','meituan-context')) {
        return (Join-Path $failureNotificationRoot "active-$safeCategory.json")
    }
    return (Join-Path $failureNotificationRoot "$dateText-$Mode-$safeCategory.json")
}

function Test-FailureNotificationReceipt([string]$Path, [string]$Category) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    if ($Category -notin @('meituan-auth','meituan-context')) { return $true }
    try {
        $receipt = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
        return ($receipt.active -eq $true)
    } catch {
        # A corrupt receipt must not hide an incident that still needs attention.
        return $false
    }
}

function Save-FailureNotificationReceipt([string]$Path, [string]$Category, [string]$IdempotencyKey) {
    $payload = [ordered]@{
        active = $true
        date = $dateText
        mode = $Mode
        category = $Category
        idempotency_key = $IdempotencyKey
        notified_at = [datetime]::Now.ToString('o')
    } | ConvertTo-Json -Compress
    $temporary = "$Path.$PID.tmp"
    [IO.File]::WriteAllText($temporary, $payload, [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Resolve-FailureNotification([string]$Category) {
    if ($Category -notin @('meituan-auth','meituan-context')) { return }
    $path = Get-FailureNotificationReceiptPath $Category
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return }
    $payload = [ordered]@{
        active = $false
        category = $Category
        resolved_at = [datetime]::Now.ToString('o')
        resolution = 'authenticated_order_read_succeeded'
    } | ConvertTo-Json -Compress
    $temporary = "$path.$PID.tmp"
    [IO.File]::WriteAllText($temporary, $payload, [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $path -Force
}

function Invoke-WithRetry([string]$Name, [scriptblock]$Operation, [int]$MaxAttempts = 3) {
    $started = [datetime]::Now
    $errors = [Collections.Generic.List[string]]::new()
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $value = & $Operation
            $stageTelemetry[$Name] = [ordered]@{
                ok = $true
                attempts = $attempt
                duration_seconds = [math]::Round(([datetime]::Now - $started).TotalSeconds, 2)
                errors = @($errors)
            }
            return [pscustomobject]@{ value = $value; attempts = $attempt }
        } catch {
            $message = $_.Exception.Message
            $errors.Add($message)
            $nonRetryable = $message -match 'MEITUAN_STORE_CONTEXT_REQUIRED|MEITUAN_BOUND_PROFILE_UNAVAILABLE|MEITUAN_STORE_IDENTITY_UNVERIFIED|MEITUAN_AUTH_REQUIRED|美团商家后台登录态已失效|表头与自动同步约定不一致|缺少必要字段|超过\d+行容量|公式存在错误|数量不一致|不是完整XLSX|干预订单匹配校验失败|"type"\s*:\s*"validation"|invalid JSON|invalid_argument'
            if ($attempt -ge $MaxAttempts -or $nonRetryable) {
                $stageTelemetry[$Name] = [ordered]@{
                    ok = $false
                    attempts = $attempt
                    duration_seconds = [math]::Round(([datetime]::Now - $started).TotalSeconds, 2)
                    errors = @($errors)
                }
                throw
            }
            Start-Sleep -Seconds ([math]::Min(12, 2 + ($attempt * 3)))
        }
    }
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
            $detail = $null
            if ($null -ne $result.error -and $result.error.message) {
                $detail = [string]$result.error.message
            } elseif ($null -ne $result.identities -and $null -ne $result.identities.user -and $result.identities.user.message) {
                $detail = [string]$result.identities.user.message
            }
            $detailSuffix = if ($detail) { "，$detail" } else { '' }
            $errors.Add("第${attempt}次：退出码=$exitCode$detailSuffix")
        } catch {
            $errors.Add("第${attempt}次：$($_.Exception.Message)")
        }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds ([math]::Min(8, $attempt * 2)) }
    }
    throw "$Label；重试${MaxAttempts}次仍失败：$($errors -join '；')"
}

function Get-LarkAuthStatus([switch]$Verify) {
    $commandArguments = @('auth','status','--json')
    if ($Verify) { $commandArguments += '--verify' }
    return Invoke-LarkJson { & $larkCliCommand @commandArguments } '无法读取飞书授权状态' 2
}

function Save-LarkUserIdentity($AuthStatus) {
    if ($null -eq $AuthStatus -or $null -eq $AuthStatus.identities -or $null -eq $AuthStatus.identities.user) { return }
    $openId = $AuthStatus.identities.user.openId
    if (-not $openId) { $openId = $AuthStatus.identities.user.open_id }
    if (-not $openId) { return }
    if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
    $identity = [ordered]@{
        open_id = [string]$openId
        app_id = [string]$AuthStatus.appId
        updated_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    }
    $tempPath = "$larkIdentityCachePath.tmp"
    [System.IO.File]::WriteAllText($tempPath, ($identity | ConvertTo-Json), [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tempPath -Destination $larkIdentityCachePath -Force
}

function Get-CachedLarkUserOpenId {
    if (-not (Test-Path -LiteralPath $larkIdentityCachePath -PathType Leaf)) { return $null }
    try {
        $identity = [System.IO.File]::ReadAllText($larkIdentityCachePath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($identity.open_id) { return [string]$identity.open_id }
    } catch {}
    return $null
}

function Convert-ToSafeState($Value) {
    if ($null -eq $Value) { return $null }
    $json = $Value | ConvertTo-Json -Depth 30 -Compress
    $json = [regex]::Replace($json, 'https://[^\s)"'']+/sheets/[A-Za-z0-9_-]+', '[飞书表格链接已省略]')
    $json = [regex]::Replace($json, '(?i)("(?:access_?token|refresh_?token|authorization|cookie|secret)"\s*:\s*)"[^"]*"', '$1"[已省略]"')
    return ($json | ConvertFrom-Json)
}

function Save-RunState($Value) {
    if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
    $json = $Value | ConvertTo-Json -Depth 16
    $stateFileName = if ($Value.sync_skipped -eq $true) { 'last_health_check.json' } else { 'last_run.json' }
    $tempPath = Join-Path $stateRoot "$stateFileName.tmp"
    [System.IO.File]::WriteAllText($tempPath, $json, [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tempPath -Destination (Join-Path $stateRoot $stateFileName) -Force
    $historyPath = Join-Path $stateRoot 'sync_runs.ndjson'
    if (Test-Path -LiteralPath $historyPath) {
        $existingHistory = [System.IO.File]::ReadAllText($historyPath, [System.Text.Encoding]::UTF8)
        $sanitizedHistory = [regex]::Replace($existingHistory, 'https://[^\s)"'']+/sheets/[A-Za-z0-9_-]+', '[飞书表格链接已省略]')
        if ($sanitizedHistory -ne $existingHistory) {
            [System.IO.File]::WriteAllText($historyPath, $sanitizedHistory, [System.Text.UTF8Encoding]::new($false))
        }
    }
    $historyLine = ($Value | ConvertTo-Json -Depth 16 -Compress) + [Environment]::NewLine
    [System.IO.File]::AppendAllText($historyPath, $historyLine, [System.Text.UTF8Encoding]::new($false))
    if ($Value.mode -eq 'night' -and $Value.ok -eq $true -and $Value.degraded -ne $true) {
        $healthyNightPath = Join-Path $stateRoot 'last_healthy_night.json'
        $healthyNightTemp = "$healthyNightPath.tmp"
        [System.IO.File]::WriteAllText($healthyNightTemp, $json, [System.Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $healthyNightTemp -Destination $healthyNightPath -Force
    }
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    $valueProperty = $cell.PSObject.Properties['value']
    if ($null -eq $valueProperty -or $null -eq $valueProperty.Value) { return $null }
    return [string]$valueProperty.Value
}

function Convert-ToNumber($Value) {
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $null
}

function Convert-ToDateValue($Value) {
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if (-not $text -or $text -eq '-') { return $null }

    # 飞书日期单元格回读时可能是 Excel 序列值，也可能是已格式化的日期文本。
    $serial = Convert-ToNumber $text
    if ($null -ne $serial -and $serial -ge 20000 -and $serial -le 80000) {
        return ([datetime]'1899-12-30').AddDays($serial).Date
    }

    $parsed = [datetime]::MinValue
    $formats = @('yyyy-MM-dd','yyyy/M/d','yyyy年M月d日','yyyy.MM.dd','M/d/yyyy')
    foreach ($format in $formats) {
        if ([datetime]::TryParseExact($text, $format, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed.Date
        }
    }
    if ([datetime]::TryParse($text, [Globalization.CultureInfo]::CurrentCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.Date
    }
    return $null
}

function Get-ExpiryRemainingText([int]$DaysRemaining) {
    if ($DaysRemaining -lt 0) { return "已过期$([math]::Abs($DaysRemaining))天" }
    if ($DaysRemaining -eq 0) { return '今天到期' }
    return "剩$DaysRemaining天"
}

function Get-ExpiringProductRecords($MasterRows, $InventoryRecords, [datetime]$AsOfDate) {
    $inventoryByCode = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($record in @($InventoryRecords)) {
        foreach ($candidateCode in @($record.code, $record.shelf_code)) {
            if ($candidateCode -and -not $inventoryByCode.ContainsKey([string]$candidateCode)) {
                $inventoryByCode.Add([string]$candidateCode, $record)
            }
        }
    }

    $cutoff = $AsOfDate.Date.AddMonths(8)
    $items = [Collections.Generic.List[object]]::new()
    foreach ($row in @($MasterRows)) {
        $expiryDate = Convert-ToDateValue (Get-CellValue $row $skuSchema.EarliestExpiry.Index)
        if ($null -eq $expiryDate -or $expiryDate -gt $cutoff) { continue }

        $currentKw = Get-CellValue $row 0
        $transitionCode = Get-CellValue $row 1
        $inventoryRecord = $null
        foreach ($candidateCode in @($currentKw, $transitionCode)) {
            if ($candidateCode -and $inventoryByCode.ContainsKey($candidateCode)) {
                $inventoryRecord = $inventoryByCode[$candidateCode]
                break
            }
        }

        $stock = if ($null -ne $inventoryRecord) { Convert-ToNumber $inventoryRecord.stock } else { Convert-ToNumber (Get-CellValue $row $skuSchema.CostBookQuantity.Index) }
        # 已经没有库存的商品不再提醒；无法取得数量但主表仍有到期日时保守提醒并标记待核对。
        if ($null -ne $stock -and $stock -le 0) { continue }
        $daysRemaining = [int]($expiryDate - $AsOfDate.Date).TotalDays
        $items.Add([pscustomobject]@{
            code = $(if ($currentKw) { $currentKw } elseif ($transitionCode) { $transitionCode } else { '缺货号' })
            title = $(if ($null -ne $inventoryRecord -and $inventoryRecord.name) { $inventoryRecord.name } else { Get-CellValue $row 2 })
            stock = $stock
            expiry_date = $expiryDate.ToString('yyyy-MM-dd')
            days_remaining = $daysRemaining
            remaining_text = Get-ExpiryRemainingText $daysRemaining
        })
    }
    return @($items | Sort-Object days_remaining, code)
}

function Get-SheetRows([string]$SheetName, [string]$Range) {
    $readRange = $null
    $readRange = {
        param([string]$CurrentRange)

        $result = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $CurrentRange --include value --max-chars 500000 --format json } "读取飞书失败：$SheetName $CurrentRange"
        if (-not $result.ok) { throw "飞书读取失败：$SheetName $CurrentRange" }
        if (-not $result.data.has_more) { return @($result.data.ranges[0].cells) }

        # lark-cli 会在大范围结果超过 max-chars 时返回 has_more=true。这个标志表示
        # 当前 JSON 已被截断，不能继续拿部分数据计算日结；按行二分读取，直到每段完整。
        if ($CurrentRange -notmatch '^([A-Z]+)(\d+):([A-Z]+)(\d+)$') {
            throw "飞书读取结果不完整且范围无法拆分：$SheetName $CurrentRange"
        }
        $startColumn = $Matches[1]
        $startRow = [int]$Matches[2]
        $endColumn = $Matches[3]
        $endRow = [int]$Matches[4]
        if ($startRow -ge $endRow) {
            throw "飞书单行读取仍不完整：$SheetName $CurrentRange"
        }

        $middleRow = [int][math]::Floor(($startRow + $endRow) / 2)
        $leftRange = "${startColumn}${startRow}:${endColumn}${middleRow}"
        $rightStart = $middleRow + 1
        $rightRange = "${startColumn}${rightStart}:${endColumn}${endRow}"
        $rows = [Collections.Generic.List[object]]::new()
        foreach ($row in @(& $readRange $leftRange)) { $rows.Add($row) }
        foreach ($row in @(& $readRange $rightRange)) { $rows.Add($row) }
        return $rows.ToArray()
    }

    return @(& $readRange $Range)
}

function Get-LedgerRows([string]$SheetName, [string]$Range) {
    if (-not (Test-Path -LiteralPath $ledgerMetaPath)) { throw '缺少每日经营简账元数据' }
    $ledgerMeta = [System.IO.File]::ReadAllText($ledgerMetaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    $ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
    $result = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $ledgerToken --sheet-name $SheetName --range $Range --include value --max-chars 1000000 --format json } "读取每日经营简账失败：$SheetName"
    if (-not $result.ok -or $result.data.has_more) { throw "每日经营简账读取结果不完整：$SheetName" }
    return @($result.data.ranges[0].cells)
}

function Test-RowDate($Value, [datetime]$ExpectedDate) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $false }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.Date -eq $ExpectedDate.Date
    }
    $serial = Convert-ToNumber $Value
    if ($null -ne $serial) {
        return ([datetime]'1899-12-30').AddDays($serial).Date -eq $ExpectedDate.Date
    }
    return $false
}

function Get-OrderNumberDistance([string]$Left, [string]$Right) {
    if (-not $Left -or -not $Right -or $Left.Length -ne $Right.Length) { return [int]::MaxValue }
    $distance = 0
    for ($index = 0; $index -lt $Left.Length; $index++) {
        if ($Left[$index] -ne $Right[$index]) { $distance++ }
    }
    return $distance
}

function Assert-TargetDateInterventionCoverage {
    $orderRows = Get-SheetRows '订单汇总' 'A2:E2001'
    $targetOrders = @($orderRows | Where-Object { Test-RowDate (Get-CellValue $_ 0) $Date })
    $targetOrderNumbers = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($row in $targetOrders) {
        $orderNo = Get-CellValue $row 2
        if ($orderNo) { $null = $targetOrderNumbers.Add($orderNo) }
    }

    $interventionRows = Get-SheetRows '干预订单' 'A2:M1001'
    $targetInterventions = @($interventionRows | Where-Object {
        (Test-ActiveInterventionOrderNo (Get-CellValue $_ 3)) -and (Test-RowDate (Get-CellValue $_ 0) $Date)
    })
    $unmatched = [Collections.Generic.List[string]]::new()
    foreach ($row in $targetInterventions) {
        $interventionOrderNo = Get-CellValue $row 3
        if (-not $interventionOrderNo -or $targetOrderNumbers.Contains($interventionOrderNo)) { continue }
        $nearby = @($targetOrders | ForEach-Object { Get-CellValue $_ 2 } | Where-Object {
            $_ -and (Get-OrderNumberDistance $interventionOrderNo $_) -eq 1
        } | Select-Object -Unique)
        if ($nearby.Count -eq 1) {
            $unmatched.Add("$interventionOrderNo（疑似 $($nearby[0])）")
        } else {
            $unmatched.Add($interventionOrderNo)
        }
    }
    if ($unmatched.Count -gt 0) {
        throw "干预订单匹配校验失败：$($unmatched -join '、')。系统已停止生成日结，避免把干预单误算成自然单"
    }
    return [pscustomobject]@{
        order_count = $targetOrderNumbers.Count
        intervention_count = $targetInterventions.Count
        unmatched_count = 0
    }
}

function Get-DailyNetProfit {
    $incomeSheet = "$($Date.Month)月收入"
    $expenseSheet = "$($Date.Month)月支出"
    $incomeRows = Get-LedgerRows $incomeSheet 'A2:K1001'
    $expenseRows = Get-LedgerRows $expenseSheet 'A2:F1001'
    $orderProfit = 0.0
    foreach ($row in $incomeRows) {
        if (-not (Test-RowDate (Get-CellValue $row 2) $Date)) { continue }
        $value = Convert-ToNumber (Get-CellValue $row 6)
        if ($null -ne $value) { $orderProfit += $value }
    }
    $expenseTotal = 0.0
    $promotionExpense = 0.0
    foreach ($row in $expenseRows) {
        if (-not (Test-RowDate (Get-CellValue $row 0) $Date)) { continue }
        $value = Convert-ToNumber (Get-CellValue $row 2)
        if ($null -ne $value) {
            $expenseTotal += $value
            $project = [string](Get-CellValue $row 1)
            if ($project -match '推广|广告|竞价') { $promotionExpense += $value }
        }
    }
    return [pscustomobject]@{
        order_profit = [math]::Round($orderProfit, 2)
        expense_total = [math]::Round($expenseTotal, 2)
        promotion_expense = [math]::Round($promotionExpense, 2)
        net_profit = [math]::Round($orderProfit - $expenseTotal, 2)
    }
}

function Assert-FormulaRange([string]$SheetName, [string]$Range) {
    $result = Invoke-LarkJson { lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --format json --exit-on-error } "公式校验调用失败：$SheetName"
    if ($result.data.status -ne 'success') { throw "公式存在错误：$SheetName" }
}

function Invoke-JsonPowerShellScript([string]$Path, [hashtable]$Arguments) {
    $invokeArguments = @{ WorkRoot = $workRoot }
    foreach ($entry in $Arguments.GetEnumerator()) {
        if ($entry.Value -is [switch] -and $entry.Value.IsPresent) {
            $invokeArguments[$entry.Key] = $true
        } elseif ($null -ne $entry.Value -and [string]$entry.Value -ne '') {
            $invokeArguments[$entry.Key] = $entry.Value
        }
    }
    $global:LASTEXITCODE = 0
    # 按真实文件路径执行，保留子脚本的 $PSScriptRoot；动态脚本块会令其为空并破坏相对依赖路径。
    $output = @(& $Path @invokeArguments 2>&1 | ForEach-Object { [string]$_ })
    if ($LASTEXITCODE -ne 0) {
        $detail = (($output | Where-Object { $_ -and $_.Trim() }) -join ' ').Trim()
        $jsonErrorLine = @($output | Where-Object { $_ -and $_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}') } | Select-Object -Last 1)
        if ($jsonErrorLine) {
            try {
                $parsedError = ([string]$jsonErrorLine[-1]).Trim() | ConvertFrom-Json
                if ($parsedError.error) { $detail = [string]$parsedError.error }
                elseif ($parsedError.message) { $detail = [string]$parsedError.message }
            } catch {}
        }
        if (-not $detail) { $detail = '子脚本未返回错误详情' }
        if ($detail.Length -gt 1800) { $detail = $detail.Substring(0, 1800) }
        throw "执行失败：$(Split-Path $Path -Leaf)：$detail"
    }
    $jsonLine = @($output | ForEach-Object { [string]$_ } | Where-Object { $_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}') } | Select-Object -Last 1)
    if (-not $jsonLine) { throw "脚本未返回有效结果：$(Split-Path $Path -Leaf)" }
    return ($jsonLine | ConvertFrom-Json)
}

function Invoke-DailyLedgerSync {
    $global:LASTEXITCODE = 0
    $output = @(& $ledgerScript -ThroughDate $Date -ProjectRoot $projectRoot 2>&1 | ForEach-Object { [string]$_ })
    if ($LASTEXITCODE -ne 0) {
        $detail = (($output | Where-Object { $_ -and $_.Trim() }) -join ' ').Trim()
        if (-not $detail) { $detail = '子脚本未返回错误详情' }
        if ($detail.Length -gt 1800) { $detail = $detail.Substring(0, 1800) }
        throw "每日经营简账刷新失败：$detail"
    }
    $jsonLine = @($output | ForEach-Object { [string]$_ } | Where-Object { $_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}') } | Select-Object -Last 1)
    if (-not $jsonLine) { throw '每日经营简账未返回有效结果' }
    return ($jsonLine | ConvertFrom-Json)
}
function Invoke-CostSnapshotRebuild([string[]]$EligibleOrderNumbers = @()) {
    if (-not (Test-Path -LiteralPath $costScript)) { throw '缺少成本快照刷新脚本' }
    $global:LASTEXITCODE = 0
    $output = @(& $costScript -ProjectRoot $projectRoot -EligibleOrderNumbers $EligibleOrderNumbers)
    if ($LASTEXITCODE -ne 0) { throw '成本快照刷新失败' }
    $jsonLine = @($output | ForEach-Object { [string]$_ } | Where-Object { $_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}') } | Select-Object -Last 1)
    if (-not $jsonLine) { throw '成本快照未返回有效结果' }
    return ($jsonLine | ConvertFrom-Json)
}

function Test-UseNoChangeFastPath(
    [string]$SyncMode,
    $OrderSyncResult,
    $GiftSyncResult,
    [bool]$WasSyncSkipped,
    [bool]$WereOrdersSkipped
) {
    if ($SyncMode -notin @('midday','reminder') -or $WasSyncSkipped -or $WereOrdersSkipped) { return $false }
    if ($null -eq $OrderSyncResult -or $null -eq $GiftSyncResult) { return $false }
    $orderOkProperty = $OrderSyncResult.PSObject.Properties['ok']
    $orderChangedProperty = $OrderSyncResult.PSObject.Properties['data_changed']
    $giftUpdatedProperty = $GiftSyncResult.PSObject.Properties['updated_cells']
    $giftOkProperty = $GiftSyncResult.PSObject.Properties['ok']
    if ($null -eq $orderOkProperty -or $null -eq $orderChangedProperty -or $null -eq $giftUpdatedProperty -or $null -eq $giftOkProperty) { return $false }
    if ($OrderSyncResult.ok -ne $true -or $GiftSyncResult.ok -ne $true) { return $false }
    return ($OrderSyncResult.data_changed -eq $false -and [int]$GiftSyncResult.updated_cells -eq 0)
}

function Get-CurrentUserOpenId {
    try {
        $status = Get-LarkAuthStatus -Verify
        $openId = $status.identities.user.openId
        if (-not $openId) { $openId = $status.identities.user.open_id }
        if ($openId) {
            Save-LarkUserIdentity $status
            return [string]$openId
        }
    } catch {}
    $cachedOpenId = Get-CachedLarkUserOpenId
    if ($cachedOpenId) { return $cachedOpenId }
    throw '无法取得飞书通知收件人；当前授权和本地身份缓存均不可用'
}

function Test-SyncPreflight {
    foreach ($requiredCommand in @('node','curl.exe')) {
        if (-not (Get-Command $requiredCommand -ErrorAction SilentlyContinue)) { throw "同步依赖缺失：$requiredCommand" }
    }
    foreach ($requiredFile in @($orderScript,$inventoryScript,$ledgerScript,$costScript,$giftReconcileScript,$webAccessCheck,$metaPath,$ledgerMetaPath)) {
        if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) { throw "同步文件缺失：$(Split-Path $requiredFile -Leaf)" }
    }
    $inventoryDrive = [System.IO.DriveInfo]::new('D')
    if ($inventoryDrive.IsReady -and $inventoryDrive.AvailableFreeSpace -lt 200MB) { throw 'D盘可用空间不足200MB，暂停库存下载以避免损坏文件' }
    $stateStoragePath = $stateRoot
    $stateItem = Get-Item -LiteralPath $stateRoot -Force -ErrorAction SilentlyContinue
    if ($null -ne $stateItem -and $stateItem.LinkType -in @('Junction','SymbolicLink') -and $stateItem.Target) {
        $stateStoragePath = [string]@($stateItem.Target)[0]
    }
    $stateDriveRoot = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($stateStoragePath))
    $stateDrive = [System.IO.DriveInfo]::new($stateDriveRoot)
    if (-not $stateDrive.IsReady -or $stateDrive.AvailableFreeSpace -lt 1GB) {
        throw "经营事务状态盘 $stateDriveRoot 可用空间不足1GB，已在写入飞书前停止同步"
    }
    $auth = $null
    $authError = $null
    try {
        $auth = Get-LarkAuthStatus -Verify
    } catch {
        $authError = $_.Exception.Message
    }

    # 本地 auth status 偶尔会在刷新瞬间误报。以最小只读请求作为最终可用性判据。
    try {
        $null = Invoke-LarkJson {
            & $larkCliCommand sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name '驾驶舱' --range 'A1:A1' --include value --max-chars 2000 --format json
        } '飞书最小只读探针失败' 2
    } catch {
        $userStatus = if ($null -ne $auth) { [string]$auth.identities.user.status } else { 'unknown' }
        $tokenStatus = if ($null -ne $auth) { [string]$auth.identities.user.tokenStatus } else { 'unknown' }
        $expiresAt = if ($null -ne $auth) { [string]$auth.identities.user.expiresAt } else { 'unknown' }
        $authDetail = if ($authError) { $authError } else { 'none' }
        throw "飞书用户凭据暂时不可用：user_status=$userStatus；token_status=$tokenStatus；expires_at=$expiresAt；read_probe=$($_.Exception.Message)；auth_status=$authDetail"
    }
    # 只读请求会在需要时刷新用户凭据；再次读取，记录刷新后的真实状态。
    try { $auth = Get-LarkAuthStatus -Verify } catch {
        if ($null -eq $auth) { try { $auth = Get-LarkAuthStatus } catch {} }
    }
    if ($null -ne $auth) { Save-LarkUserIdentity $auth }
    return [pscustomobject]@{
        node = 'ready'
        curl = 'ready'
        lark_user = 'verified_by_sheet_read'
        lark_bot = if ($null -ne $auth) { [string]$auth.identities.bot.status } else { 'unknown' }
        lark_cli = $larkCliCommand
        token_status = if ($null -ne $auth) { [string]$auth.identities.user.tokenStatus } else { 'unknown' }
        token_expires_at = if ($null -ne $auth) { [string]$auth.identities.user.expiresAt } else { 'unknown' }
        free_space_gb = [math]::Round($stateDrive.AvailableFreeSpace / 1GB, 2)
        state_drive = $stateDriveRoot
        state_free_space_gb = [math]::Round($stateDrive.AvailableFreeSpace / 1GB, 2)
        inventory_drive = $inventoryDrive.Name
        inventory_free_space_gb = [math]::Round($inventoryDrive.AvailableFreeSpace / 1GB, 2)
    }
}

function Convert-ToWindowsCommandLineArgument([string]$Value) {
    if ($null -eq $Value) { return '""' }
    if ($Value.Length -gt 0 -and $Value -notmatch '[\s"]') { return $Value }
    $builder = [System.Text.StringBuilder]::new()
    $null = $builder.Append('"')
    $backslashes = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq '\') {
            $backslashes++
            continue
        }
        if ($character -eq '"') {
            $null = $builder.Append((('\' * ($backslashes * 2 + 1)) -join ''))
            $null = $builder.Append('"')
            $backslashes = 0
            continue
        }
        if ($backslashes -gt 0) { $null = $builder.Append((('\' * $backslashes) -join '')); $backslashes = 0 }
        $null = $builder.Append($character)
    }
    if ($backslashes -gt 0) { $null = $builder.Append((('\' * ($backslashes * 2)) -join '')) }
    $null = $builder.Append('"')
    return $builder.ToString()
}

function Invoke-LarkCliCaptured([string[]]$Arguments, [int]$MaxAttempts = 3) {
    $nodeCommand = (Get-Command node -ErrorAction Stop).Source
    $larkEntry = Join-Path $larkCliDirectory 'node_modules\@larksuite\cli\scripts\run.js'
    if (-not (Test-Path -LiteralPath $larkEntry)) { throw '找不到飞书 CLI Node 入口' }
    $lastResult = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
        $startInfo.FileName = $nodeCommand
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.StandardOutputEncoding = [System.Text.UTF8Encoding]::new($false)
        $startInfo.StandardErrorEncoding = [System.Text.UTF8Encoding]::new($false)
        $allArguments = @($larkEntry) + @($Arguments)
        if ($startInfo.PSObject.Properties.Name -contains 'ArgumentList') {
            foreach ($argument in $allArguments) { $startInfo.ArgumentList.Add([string]$argument) }
        } else {
            $startInfo.Arguments = (@($allArguments | ForEach-Object { Convert-ToWindowsCommandLineArgument ([string]$_) }) -join ' ')
        }
        $process = [System.Diagnostics.Process]::new()
        $process.StartInfo = $startInfo
        if (-not $process.Start()) { throw '无法启动飞书 CLI' }
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $lastResult = [pscustomobject]@{ exit_code=$process.ExitCode; stdout=$stdout; stderr=$stderr; attempts=$attempt }
        $validSuccess = $false
        if ($process.ExitCode -eq 0) {
            try { $validSuccess = [bool](($stdout | ConvertFrom-Json).ok) } catch { $validSuccess = $false }
        }
        if ($validSuccess) { return $lastResult }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds ([math]::Min(8, $attempt * 2)) }
    }
    return $lastResult
}

function Send-FeishuMessage([string]$Text, $Card, [string]$IdempotencyKey) {
    if ($DryRun) {
        $previewFormat = if ($null -eq $Card) { 'text' } else { 'card' }
        return [pscustomobject]@{ sent = $false; dry_run = $true; format = $previewFormat }
    }
    $openId = Get-CurrentUserOpenId
    $cardError = $null
    if ($null -ne $Card) {
        $cardJson = $Card | ConvertTo-Json -Depth 30 -Compress
        $invocation = Invoke-LarkCliCaptured @('im','+messages-send','--as','bot','--user-id',$openId,'--msg-type','interactive','--content',$cardJson,'--idempotency-key',$IdempotencyKey,'--format','json')
        $raw = $invocation.stdout
        $cardExitCode = $invocation.exit_code
        try { $result = $raw | ConvertFrom-Json } catch { $result = $null }
        if ($cardExitCode -eq 0 -and $result.ok) {
            return [pscustomobject]@{ sent = $true; dry_run = $false; format = 'card' }
        }
        $cardError = if ($null -ne $result -and $result.error.message) { [string]$result.error.message } elseif ($invocation.stderr) { $invocation.stderr.Trim() } else { ($raw.Trim() -split "`r?`n" | Select-Object -First 3) -join ' ' }
    }

    $fallbackKey = if ($IdempotencyKey.Length -le 45) { "$IdempotencyKey-text" } else { $IdempotencyKey.Substring(0,45) + '-text' }
    $invocation = Invoke-LarkCliCaptured @('im','+messages-send','--as','bot','--user-id',$openId,'--text',$Text,'--idempotency-key',$fallbackKey,'--format','json')
    $raw = $invocation.stdout
    $fallbackExitCode = $invocation.exit_code
    if ($fallbackExitCode -ne 0) { throw "飞书私聊发送失败：$($raw.Trim())" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "飞书私聊发送失败：$($result.error.message)" }
    $messageFormat = if ($null -eq $Card) { 'text' } else { 'text-fallback' }
    return [pscustomobject]@{ sent = $true; dry_run = $false; format = $messageFormat; card_error = $cardError }
}

function Format-Money($Value) {
    $number = Convert-ToNumber $Value
    if ($null -eq $number) { return '待补' }
    return ('¥{0:N2}' -f $number)
}

function Build-Snapshot {
    $dashboardRows = Get-SheetRows '驾驶舱' 'A2:C19'
    $metrics = [ordered]@{}
    foreach ($row in $dashboardRows) {
        $name = Get-CellValue $row 1
        if ($name) { $metrics[$name] = Get-CellValue $row 2 }
    }

    $orderRows = Get-SheetRows '订单汇总' 'A2:E2001'
    $todayRows = @($orderRows | Where-Object {
        $orderDate = Get-CellValue $_ 0
        $orderDate -and $orderDate.StartsWith($dateText)
    })
    $activeRows = @($todayRows | Where-Object { (Get-CellValue $_ 3) -ne '已取消' })
    $platformNaturalCount = @($activeRows | Where-Object { (Get-CellValue $_ 4) -eq '自然' }).Count
    $interventionCount = @($activeRows | Where-Object { (Get-CellValue $_ 4) -eq '干预' }).Count
    $cancelledCount = @($todayRows | Where-Object { (Get-CellValue $_ 3) -eq '已取消' }).Count
    $offlineRows = Get-SheetRows '线下订单' 'A2:J1001'
    $offlineCount = @($offlineRows | Where-Object {
        (Test-RowDate (Get-CellValue $_ 0) $Date) -and (Get-CellValue $_ 2)
    }).Count
    $naturalCount = $platformNaturalCount + $offlineCount

    $pendingDelivery = [Collections.Generic.List[object]]::new()
    foreach ($row in (Get-SheetRows '干预订单' 'A2:M1001')) {
        if (-not (Test-RowDate (Get-CellValue $row 0) $Date)) { continue }
        $orderNo = Get-CellValue $row 3
        if (-not (Test-ActiveInterventionOrderNo $orderNo)) { continue }
        $deliveryType = Get-CellValue $row 5
        $deliveryFee = Convert-ToNumber (Get-CellValue $row 7)
        $needsConfirmation = (-not $deliveryType) -or ($deliveryType -eq '骑手配送' -and ($null -eq $deliveryFee -or $deliveryFee -eq 0))
        if ($needsConfirmation) {
            $pendingDelivery.Add([pscustomobject]@{
                daily_order_no = Get-CellValue $row 1
                order_no = $orderNo
                delivery_type = $deliveryType
            })
        }
    }

    $pendingCost = [Collections.Generic.List[object]]::new()
    $incomeSheet = "$($Date.Month)月收入"
    foreach ($row in (Get-LedgerRows $incomeSheet 'A2:K1001')) {
        if (-not (Test-RowDate (Get-CellValue $row 2) $Date)) { continue }
        $orderNo = Get-CellValue $row 1
        if (-not $orderNo) { continue }
        $productCost = Convert-ToNumber (Get-CellValue $row 4)
        if ($null -eq $productCost) {
            $pendingCost.Add([pscustomobject]@{
                daily_order_no = Get-CellValue $row 0
                order_no = $orderNo
                sold_goods = Get-CellValue $row 7
                note = Get-CellValue $row 10
            })
        }
    }

    $interventionSales = @{}
    $windowStart = $Date.Date.AddDays(-29)
    $itemRows = Get-SheetRows '订单商品' 'A2:I4001'
    foreach ($row in $itemRows) {
        $rowDateText = Get-CellValue $row 0
        $rowStatus = Get-CellValue $row 3
        $rowType = Get-CellValue $row 4
        $goodsCode = Get-CellValue $row 5
        if (-not $rowDateText -or -not $goodsCode -or $rowType -ne '干预' -or $rowStatus -eq '已取消') { continue }
        $rowDate = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($rowDateText.Substring(0,10), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$rowDate)) { continue }
        if ($rowDate.Date -lt $windowStart -or $rowDate.Date -gt $Date.Date) { continue }
        $quantity = Convert-ToNumber (Get-CellValue $row 8)
        if ($null -eq $quantity) { continue }
        if (-not $interventionSales.ContainsKey($goodsCode)) { $interventionSales[$goodsCode] = 0.0 }
        $interventionSales[$goodsCode] += $quantity
    }

    $inventoryRows = Get-SheetRows '库存快照' 'B2:M2000'
    $inventoryRecords = @()
    foreach ($row in $inventoryRows) {
        $code = Get-CellValue $row 0
        $name = Get-CellValue $row 1
        if (-not $code -and -not $name) { continue }
        $shelfCode = Get-CellValue $row 5
        $platformMonthlySales = Convert-ToNumber (Get-CellValue $row 3)
        $interventionQuantity = 0.0
        if ($code -and $interventionSales.ContainsKey($code)) { $interventionQuantity += $interventionSales[$code] }
        if ($shelfCode -and $shelfCode -ne $code -and $interventionSales.ContainsKey($shelfCode)) { $interventionQuantity += $interventionSales[$shelfCode] }
        $naturalMonthlySales = if ($null -eq $platformMonthlySales) { $null } else { [math]::Max([double]0.0, [double]$platformMonthlySales - [double]$interventionQuantity) }
        $inventoryRecords += [pscustomobject]@{
            code = $code
            name = $name
            shelf_code = $shelfCode
            stock = Convert-ToNumber (Get-CellValue $row 2)
            platform_monthly_sales = $platformMonthlySales
            intervention_sales_30d = $interventionQuantity
            natural_monthly_sales = $naturalMonthlySales
            status = Get-CellValue $row 4
            reminder = Get-CellValue $row 9
            restore_quantity = Convert-ToNumber (Get-CellValue $row 11)
            is_sample = [bool]($name -match '(?i)(试吃装|尝鲜装|试吃尝鲜|试用装|体验装)')
        }
    }
    $urgent = @($inventoryRecords | Where-Object {
        -not $_.is_sample -and $null -ne $_.stock -and $_.stock -le 1 -and $null -ne $_.natural_monthly_sales -and $_.natural_monthly_sales -ge 1
    } | Sort-Object @{Expression='natural_monthly_sales';Descending=$true}, @{Expression='stock';Descending=$false}, code)
    $excludedSamples = @($inventoryRecords | Where-Object {
        $_.is_sample -and $null -ne $_.stock -and $_.stock -le 1 -and $null -ne $_.natural_monthly_sales -and $_.natural_monthly_sales -ge 1
    })
    $restore = @($inventoryRecords | Where-Object { $null -ne $_.restore_quantity -and $_.restore_quantity -gt 0 } | Sort-Object @{Expression='restore_quantity';Descending=$true}, code)
    $expiringProducts = @()
    $expiryCutoff = $null
    if ($Mode -eq 'night') {
        $skuMasterRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $expiringProducts = Get-ExpiringProductRecords $skuMasterRows $inventoryRecords $Date
        $expiryCutoff = $Date.Date.AddMonths(8).ToString('yyyy-MM-dd')
    }

    return [pscustomobject]@{
        metrics = $metrics
        today_total = $activeRows.Count + $offlineCount
        today_natural = $naturalCount
        today_platform_natural = $platformNaturalCount
        today_intervention = $interventionCount
        today_cancelled = $cancelledCount
        today_offline = $offlineCount
        pending_delivery = $pendingDelivery.ToArray()
        pending_cost = $pendingCost.ToArray()
        urgent_restock = $urgent
        restock_excluded_samples = $excludedSamples
        restore_inventory = $restore
        expiring_products = $expiringProducts
        expiry_cutoff = $expiryCutoff
    }
}

function Limit-Text([string]$Value, [int]$Length) {
    if (-not $Value) { return '' }
    if ($Value.Length -le $Length) { return $Value }
    return $Value.Substring(0, $Length - 1) + '…'
}

function Get-FailureGuidance([string]$Reason) {
    if ($Reason -match 'MEITUAN_STORE_CONTEXT_REQUIRED|MEITUAN_BOUND_PROFILE_UNAVAILABLE|MEITUAN_STORE_IDENTITY_UNVERIFIED') { return '系统只使用已绑定的 Chrome 资料“人员 1”。请打开该资料并确认进入“宠速星宠物用品店”后台；系统不会检查其他 Chrome 账户。' }
    if ($Reason -match 'CDP|remote debugging|远程调试') { return '请保持 Chrome 打开，并确认 chrome://inspect/#remote-debugging 中已允许远程调试。' }
    if ($Reason -match '登录|login|扫码') { return '请在 Chrome 中重新登录美团商家后台；登录后系统会在下一次计划任务自动重试。' }
    if ($Reason -match '库存导出|inventory export|create-export|导出弹窗|下载文件') { return '系统已完成多轮页面重开与重试；零点失败后还会在 00:20、00:40 自动补跑。' }
    if ($Reason -match '下载|XLSX|D盘|空间') { return '请检查网络和 D 盘空间；系统不会用损坏或未完成的文件覆盖上一份库存。' }
    if ($Reason -match 'refresh.*expired|refresh.*missing|重新.*授权|not_authorized') { return '飞书刷新凭据已经失效，需要重新完成一次飞书 CLI 用户授权。' }
    if ($Reason -match '飞书.*凭据|飞书.*授权|授权未就绪|auth|token') { return '系统会自动执行只读探针并阶梯重试；若连续到下一整点仍失败，再检查飞书 CLI 授权状态。' }
    if ($Reason -match '公式|表头|必要字段|容量') { return '这是数据结构或公式校验问题，系统已停止继续写入，避免带着错误数据生成日结。' }
    return '系统已保留分阶段运行记录；下一次计划任务会重新执行，连续失败时请查看失败卡片中的具体阶段。'
}

function Build-Card($Snapshot, [bool]$Succeeded, [string]$FailureReason = '') {
    $modeName = switch ($Mode) { 'midday' { '定时' }; 'reminder' { '补录提醒' }; 'night' { '日结' }; default { '手动' } }
    $timestamp = [datetime]::Now.ToString('yyyy-MM-dd HH:mm')
    if (-not $Succeeded) {
        return [ordered]@{
            schema = '2.0'
            config = [ordered]@{ width_mode = 'fill'; summary = [ordered]@{ content = '宠速星同步失败，需要检查' } }
            header = [ordered]@{
                title = [ordered]@{ tag = 'plain_text'; content = '宠速星｜同步失败' }
                subtitle = [ordered]@{ tag = 'plain_text'; content = $timestamp }
                template = 'red'
                icon = [ordered]@{ tag = 'standard_icon'; token = 'warning_outlined'; color = 'red' }
                text_tag_list = @([ordered]@{ tag = 'text_tag'; text = [ordered]@{ tag = 'plain_text'; content = '需要处理' }; color = 'red' })
            }
            body = [ordered]@{
                direction = 'vertical'
                padding = '12px 12px 20px 12px'
                vertical_spacing = '12px'
                elements = @(
                    [ordered]@{
                        tag = 'column_set'; flex_mode = 'none'
                        columns = @([ordered]@{
                            tag = 'column'; width = 'weighted'; weight = 1; background_style = 'red-50'; padding = '12px'; vertical_spacing = '4px'
                            elements = @(
                                [ordered]@{ tag = 'markdown'; content = '**同步未完成**' },
                                [ordered]@{ tag = 'markdown'; content = (Limit-Text $FailureReason 260); text_size = 'normal' }
                            )
                        })
                    },
                    [ordered]@{ tag = 'markdown'; content = "$(Get-FailureGuidance $FailureReason)`n系统没有修改任何美团订单或库存。`n[打开经营驾驶舱]($($meta.workbook_url))"; text_size = 'notation' }
                )
            }
        }
    }

    $isNight = $Mode -eq 'night'
    $syncWarnings = @($Snapshot.sync_warnings)
    $hasWarnings = $syncWarnings.Count -gt 0
    $isOperationallyDegraded = [bool]$Snapshot.sync_degraded
    $focusValue = if ($isNight) { Format-Money $Snapshot.today_net_profit } else { Format-Money $Snapshot.metrics['本月经营毛利'] }
    $focusLabel = if ($isNight) { '当天赚/亏（暂估）' } else { '本月订单毛利（暂估）' }
    $profitNumber = if ($isNight) { Convert-ToNumber $Snapshot.today_net_profit } else { Convert-ToNumber $Snapshot.metrics['本月经营毛利'] }
    $profitColor = if ($null -ne $profitNumber -and $profitNumber -lt 0) { 'red' } else { 'green' }
    $profitBackground = if ($profitColor -eq 'red') { 'red-50' } else { 'green-50' }
    $elements = [Collections.Generic.List[object]]::new()
    $elements.Add([ordered]@{
        tag = 'column_set'; flex_mode = 'none'; horizontal_spacing = '12px'
        columns = @([ordered]@{
            tag = 'column'; width = 'weighted'; weight = 1; background_style = $profitBackground; padding = '12px'; vertical_spacing = '2px'
            elements = @(
                [ordered]@{ tag = 'markdown'; content = "## <font color='$profitColor'>$focusValue</font>"; text_align = 'center' },
                [ordered]@{ tag = 'markdown'; content = "<font color='grey'>$focusLabel</font>"; text_align = 'center'; text_size = 'notation' }
            )
        })
    })

    if ($isNight) {
        $elements.Add([ordered]@{ tag = 'markdown'; content = "**推广费**　$(Format-Money $Snapshot.today_promotion_expense)　<font color='grey'>23:55 自动读取</font>" })
    }

    if ($hasWarnings) {
        $warningTitle = if ($isOperationallyDegraded) { '**本次同步已自动降级处理**' } else { '**本次同步完成，但有经营事项需要核对**' }
        $warningContent = "$warningTitle`n" + (($syncWarnings | ForEach-Object { "- $_" }) -join "`n")
        $elements.Add([ordered]@{
            tag = 'column_set'; flex_mode = 'none'
            columns = @([ordered]@{
                tag = 'column'; width = 'weighted'; weight = 1; background_style = 'orange-50'; padding = '12px'; vertical_spacing = '4px'
                elements = @([ordered]@{ tag = 'markdown'; content = $warningContent })
            })
        })
    }
    $elements.Add([ordered]@{
        tag = 'column_set'; flex_mode = 'none'; horizontal_spacing = '8px'
        columns = @(
            [ordered]@{ tag = 'column'; width = 'weighted'; weight = 1; background_style = 'grey-50'; padding = '8px'; vertical_spacing = '2px'; elements = @([ordered]@{ tag = 'markdown'; content = "## $($Snapshot.today_total)"; text_align = 'center' }, [ordered]@{ tag = 'markdown'; content = "<font color='grey'>有效订单</font>"; text_align = 'center'; text_size = 'notation' }) },
            [ordered]@{ tag = 'column'; width = 'weighted'; weight = 1; background_style = 'grey-50'; padding = '8px'; vertical_spacing = '2px'; elements = @([ordered]@{ tag = 'markdown'; content = "## $($Snapshot.today_natural)"; text_align = 'center' }, [ordered]@{ tag = 'markdown'; content = "<font color='grey'>真实单</font>"; text_align = 'center'; text_size = 'notation' }) },
            [ordered]@{ tag = 'column'; width = 'weighted'; weight = 1; background_style = 'grey-50'; padding = '8px'; vertical_spacing = '2px'; elements = @([ordered]@{ tag = 'markdown'; content = "## $($Snapshot.today_intervention)"; text_align = 'center' }, [ordered]@{ tag = 'markdown'; content = "<font color='grey'>干预单</font>"; text_align = 'center'; text_size = 'notation' }) },
            [ordered]@{ tag = 'column'; width = 'weighted'; weight = 1; background_style = 'grey-50'; padding = '8px'; vertical_spacing = '2px'; elements = @([ordered]@{ tag = 'markdown'; content = "## $($Snapshot.today_offline)"; text_align = 'center' }, [ordered]@{ tag = 'markdown'; content = "<font color='grey'>线下单</font>"; text_align = 'center'; text_size = 'notation' }) }
        )
    })

    if ($isNight -and (@($Snapshot.pending_cost).Count -gt 0 -or @($Snapshot.pending_delivery).Count -gt 0)) {
        $warningText = "**数据仍待补齐**"
        if (@($Snapshot.pending_cost).Count -gt 0) { $warningText += "`n货物成本待核对：$(@($Snapshot.pending_cost).Count) 单" }
        if (@($Snapshot.pending_delivery).Count -gt 0) { $warningText += "`n干预单配送费待确认：$(@($Snapshot.pending_delivery).Count) 单" }
        $warningText += "`n<font color='grey'>当天赚/亏目前是暂估值，补录后系统会重算原日期。</font>"
        $elements.Add([ordered]@{
            tag = 'column_set'; flex_mode = 'none'
            columns = @([ordered]@{
                tag = 'column'; width = 'weighted'; weight = 1; background_style = 'orange-50'; padding = '12px'; vertical_spacing = '4px'
                elements = @([ordered]@{ tag = 'markdown'; content = $warningText })
            })
        })
    }

    $urgentItems = @($Snapshot.urgent_restock | Select-Object -First 10)
    $elements.Add([ordered]@{ tag = 'markdown'; content = "**重点补货**　共 $(@($Snapshot.urgent_restock).Count) 个商品" })
    if ($urgentItems.Count -gt 0) {
        $tableRows = @()
        foreach ($item in $urgentItems) {
            $tableRows += [ordered]@{
                code = $(if ($item.code) { $item.code } else { '缺货号' })
                title = Limit-Text $item.name 42
                stock = [double]$item.stock
                natural_sales = [double]$item.natural_monthly_sales
            }
        }
        $elements.Add([ordered]@{
            tag = 'table'; page_size = [math]::Min(10, $tableRows.Count); row_height = 'auto'; row_max_height = '72px'; freeze_first_column = $true
            header_style = [ordered]@{ background_style = 'grey'; bold = $true; lines = 1 }
            columns = @(
                [ordered]@{ name = 'code'; display_name = '货号'; data_type = 'text'; width = '18%' },
                [ordered]@{ name = 'title'; display_name = '商品标题'; data_type = 'text'; width = '52%' },
                [ordered]@{ name = 'stock'; display_name = '库存'; data_type = 'number'; width = '12%'; horizontal_align = 'right'; format = [ordered]@{ precision = 0 } },
                [ordered]@{ name = 'natural_sales'; display_name = '自然月售'; data_type = 'number'; width = '18%'; horizontal_align = 'right'; format = [ordered]@{ precision = 0 } }
            )
            rows = $tableRows
        })
    } else {
        $elements.Add([ordered]@{ tag = 'markdown'; content = '<font color=''grey''>暂无符合规则的重点补货商品。</font>' })
    }

    if ($isNight) {
        $expiringItems = @($Snapshot.expiring_products | Select-Object -First 10)
        $elements.Add([ordered]@{ tag = 'markdown'; content = "**临期商品**　8个月内到期，共 $(@($Snapshot.expiring_products).Count) 个商品" })
        if ($expiringItems.Count -gt 0) {
            $expiryRows = @()
            foreach ($item in $expiringItems) {
                $expiryRows += [ordered]@{
                    code = $item.code
                    title = Limit-Text $item.title 34
                    stock = $(if ($null -eq $item.stock) { '待核对' } else { ('{0:0.##}' -f [double]$item.stock) })
                    expiry = $item.expiry_date
                    remaining = $item.remaining_text
                }
            }
            $elements.Add([ordered]@{
                tag = 'table'; page_size = [math]::Min(10, $expiryRows.Count); row_height = 'auto'; row_max_height = '72px'; freeze_first_column = $true
                header_style = [ordered]@{ background_style = 'grey'; bold = $true; lines = 1 }
                columns = @(
                    [ordered]@{ name = 'code'; display_name = '货号'; data_type = 'text'; width = '18%' },
                    [ordered]@{ name = 'title'; display_name = '商品标题'; data_type = 'text'; width = '35%' },
                    [ordered]@{ name = 'stock'; display_name = '库存'; data_type = 'text'; width = '12%'; horizontal_align = 'right' },
                    [ordered]@{ name = 'expiry'; display_name = '到期日'; data_type = 'text'; width = '20%' },
                    [ordered]@{ name = 'remaining'; display_name = '剩余'; data_type = 'text'; width = '15%' }
                )
                rows = $expiryRows
            })
        } else {
            $elements.Add([ordered]@{ tag = 'markdown'; content = '<font color=''grey''>未来8个月内暂无已登记的临期商品；未填写到期日期的商品不会进入提醒。</font>' })
        }
    }

    $footer = "**补货规则**　库存≤1，且自然月售估算≥1；试吃/尝鲜/试用/体验装不提醒`n<font color='grey'>自然月售估算 = 平台月售 − 近30天未取消干预销量，最低按0计。</font>"
    if (@($Snapshot.urgent_restock).Count -gt 10) { $footer += "`n其余 $(@($Snapshot.urgent_restock).Count - 10) 个请在驾驶舱查看。" }
    if (@($Snapshot.restock_excluded_samples).Count -gt 0) { $footer += "`n已自动排除 $(@($Snapshot.restock_excluded_samples).Count) 个试吃/尝鲜类商品。" }
    if (@($Snapshot.restore_inventory).Count -gt 0) { $footer += "`n**待恢复库存：$(@($Snapshot.restore_inventory).Count) 个货号**" }
    if ($isNight) {
        $footer += "`n**临期规则**　当前库存最早到期日≤$($Snapshot.expiry_cutoff)（日结日起8个自然月）；未登记日期的不提醒。"
        if (@($Snapshot.expiring_products).Count -gt 10) { $footer += "`n其余 $(@($Snapshot.expiring_products).Count - 10) 个临期商品请在 SKU 成本主表查看。" }
    }
    $footer += "`n取消单 $($Snapshot.today_cancelled)｜缺货 $($Snapshot.metrics['当前缺货SKU数'])｜低库存 $($Snapshot.metrics['当前低库存SKU数'])｜待核对 $($Snapshot.metrics['待核对数据数'])"
    $footer += "`n[打开经营驾驶舱]($($meta.workbook_url))"
    $elements.Add([ordered]@{ tag = 'markdown'; content = $footer; text_size = 'notation' })

    return [ordered]@{
        schema = '2.0'
        config = [ordered]@{
            width_mode = 'fill'; update_multi = $true; summary = [ordered]@{ content = $(if ($isNight) { "宠速星经营日结：当天赚/亏（暂估）$focusValue" } else { "宠速星${modeName}同步：本月订单毛利（暂估）$focusValue" }) }
            style = [ordered]@{
                color = [ordered]@{
                    'cus-primary' = [ordered]@{ light_mode = 'rgba(30,120,255,1)'; dark_mode = 'rgba(80,150,255,1)' }
                    'cus-muted' = [ordered]@{ light_mode = 'rgba(100,106,115,1)'; dark_mode = 'rgba(150,155,163,1)' }
                }
            }
        }
        header = [ordered]@{
            title = [ordered]@{ tag = 'plain_text'; content = $(if ($isNight) { "宠速星｜$($Date.Month)月$($Date.Day)日经营日结" } else { "宠速星｜${modeName}同步" }) }
            subtitle = [ordered]@{ tag = 'plain_text'; content = "$timestamp · 平台订单、线下订单与库存" }
            template = $(if ($hasWarnings) { 'orange' } else { $profitColor })
            icon = [ordered]@{ tag = 'standard_icon'; token = 'wiki-bitable_colorful' }
            text_tag_list = @([ordered]@{ tag = 'text_tag'; text = [ordered]@{ tag = 'plain_text'; content = $(if ($isOperationallyDegraded) { '已完成·有降级' } elseif ($hasWarnings) { '已完成·需核对' } else { '已完成' }) }; color = $(if ($hasWarnings) { 'orange' } else { $profitColor }) })
        }
        body = [ordered]@{ direction = 'vertical'; padding = '12px 12px 20px 12px'; vertical_spacing = '12px'; elements = $elements.ToArray() }
    }
}

function Build-Message($Snapshot, [bool]$Succeeded, [string]$FailureReason = '') {
    $modeName = switch ($Mode) { 'midday' { '定时' }; 'reminder' { '补录提醒' }; 'night' { '日结' }; default { '手动' } }
    if (-not $Succeeded) {
        return "【宠速星经营助手｜同步失败】`n时间：$([datetime]::Now.ToString('yyyy-MM-dd HH:mm'))`n原因：$FailureReason`n处理：$(Get-FailureGuidance $FailureReason)`n系统没有修改任何美团订单或库存。"
    }
    if ($Mode -eq 'reminder') {
        $lines = [Collections.Generic.List[string]]::new()
        $lines.Add("⏰ 23:00 日结前补录提醒")
        $lines.Add('')
        $lines.Add('距离 00:00 日结还有 1 小时，请检查今天是否都已告诉经营管家：')
        $lines.Add('')
        $lines.Add('1. 🚚 每一单配送费')
        $lines.Add('逐单确认配送费；自己送也要确认，按 0 元记录。')
        $lines.Add('')
        $lines.Add('2. 💸 其他支出')
        $lines.Add('补录采购付款、店铺杂费、平台外配送等当天实际支出。')
        $lines.Add('')
        $lines.Add('3. 📦 补货')
        $lines.Add('补录货号、数量和金额；录入后记得把平台库存调高。')
        $lines.Add('')
        $lines.Add('4. 🧾 干预单')
        $lines.Add('补录订单号、补单人和配送费；完成后记得恢复平台库存。')
        $lines.Add('')
        $lines.Add('5. 🛍️ 线下成单')
        $lines.Add('补录实收金额、售出 KW 和数量；完成后记得扣减平台库存。')
        $pendingDelivery = @($Snapshot.pending_delivery)
        $pendingCost = @($Snapshot.pending_cost)
        $lines.Add('')
        $lines.Add("🔎 系统检测到的待确认项")
        $lines.Add("配送费待确认：$($pendingDelivery.Count) 单")
        if ($pendingDelivery.Count -eq 0) {
            $lines.Add('✅ 暂未发现干预单配送费缺失。')
        } else {
            foreach ($item in @($pendingDelivery | Select-Object -First 12)) {
                $label = if ($item.daily_order_no) { $item.daily_order_no } else { '单号待同步' }
                $lines.Add("- $label｜$($item.order_no)")
            }
            $lines.Add('你可以直接说“这单我自己送”或“这单配送费 6.5 元”。')
        }
        $lines.Add('')
        $lines.Add("货物成本待核对：$($pendingCost.Count) 单")
        if ($pendingCost.Count -eq 0) {
            $lines.Add('✅ 今天自然单货物成本已齐。')
        } else {
            foreach ($item in @($pendingCost | Select-Object -First 12)) {
                $label = if ($item.daily_order_no) { $item.daily_order_no } else { '单号待同步' }
                $detail = if ($item.note) { $item.note } else { $item.sold_goods }
                $lines.Add("- $label｜$($item.order_no)｜$((Limit-Text $detail 70))")
            }
            $lines.Add('⚠️ 这些订单不补齐成本，00:00 的赚/亏只能是暂估值。')
        }
        return ($lines -join "`n")
    }
    if ($Mode -eq 'night') {
        $dayLabel = "$($Date.Month)月$($Date.Day)日"
        $profitText = if ($null -eq $Snapshot.today_net_profit) { '待核对' } else { ('¥{0:N2}' -f [double]$Snapshot.today_net_profit) }
        $allUrgentItems = @($Snapshot.urgent_restock)
        $urgentItems = @($allUrgentItems | Select-Object -First 8)
        $lines = [Collections.Generic.List[string]]::new()
        $lines.Add("📊 $dayLabel 经营日结")
        $lines.Add('')
        $lines.Add("🧾 有效订单：$($Snapshot.today_total) 单")
        $lines.Add("真实单 $($Snapshot.today_natural)｜干预单 $($Snapshot.today_intervention)｜取消单 $($Snapshot.today_cancelled)｜线下单 $($Snapshot.today_offline)")
        $lines.Add('')
        $lines.Add("💰 当天赚/亏（暂估）：$profitText")
        $lines.Add("📣 推广费：$(Format-Money $Snapshot.today_promotion_expense)（23:55自动读取）")
        if (@($Snapshot.sync_warnings).Count -gt 0) {
            $lines.Add('')
            $warningHeading = if ([bool]$Snapshot.sync_degraded) { '⚠️ 本次同步已自动降级处理：' } else { '⚠️ 经营事项提醒：' }
            $lines.Add($warningHeading)
            foreach ($warning in @($Snapshot.sync_warnings)) { $lines.Add("- $warning") }
        }
        $pendingCost = @($Snapshot.pending_cost)
        $pendingDelivery = @($Snapshot.pending_delivery)
        if ($pendingCost.Count -gt 0) {
            $lines.Add("🚨 强烈提醒：仍有 $($pendingCost.Count) 单货物成本待核对，上述金额不是最终结果！")
            foreach ($item in @($pendingCost | Select-Object -First 8)) {
                $detail = if ($item.note) { $item.note } else { '' }
                $lines.Add("- $($item.daily_order_no)｜$($item.order_no)$(if ($detail) { "｜$((Limit-Text $detail 60))" } else { '' })")
            }
        }
        if ($pendingDelivery.Count -gt 0) {
            $lines.Add("⚠️ 仍有 $($pendingDelivery.Count) 笔干预单配送费待确认，暂估赚/亏可能偏高。")
            foreach ($item in @($pendingDelivery | Select-Object -First 8)) { $lines.Add("- $($item.daily_order_no)｜$($item.order_no)") }
        }
        $lines.Add('')
        $lines.Add("🚨 紧急补货：$($allUrgentItems.Count) 款")
        if ($urgentItems.Count -eq 0) {
            $lines.Add('✅ 目前没有符合规则的紧急补货商品。')
        } else {
            foreach ($item in $urgentItems) {
                $code = if ($item.code) { $item.code } elseif ($item.shelf_code) { $item.shelf_code } else { '缺货号' }
                $title = Limit-Text $item.name 24
                $stock = if ($null -eq $item.stock) { '待核对' } else { ('{0:N0}' -f [double]$item.stock) }
                $lines.Add("- $code｜$title｜库存 $stock")
            }
            if ($allUrgentItems.Count -gt $urgentItems.Count) { $lines.Add("另有 $($allUrgentItems.Count - $urgentItems.Count) 款请在驾驶舱查看。") }
        }
        $allExpiringItems = @($Snapshot.expiring_products)
        $expiringItems = @($allExpiringItems | Select-Object -First 8)
        $lines.Add('')
        $lines.Add("⏳ 临期商品：$($allExpiringItems.Count) 款（未来8个月内到期）")
        if ($expiringItems.Count -eq 0) {
            $lines.Add('✅ 暂无已登记的临期商品；未填写到期日期的商品不会进入提醒。')
        } else {
            foreach ($item in $expiringItems) {
                $title = Limit-Text $item.title 22
                $stock = if ($null -eq $item.stock) { '待核对' } else { ('{0:0.##}' -f [double]$item.stock) }
                $lines.Add("- $($item.code)｜$title｜库存 $stock｜到期 $($item.expiry_date)｜$($item.remaining_text)")
            }
            if ($allExpiringItems.Count -gt $expiringItems.Count) { $lines.Add("另有 $($allExpiringItems.Count - $expiringItems.Count) 款请在 SKU 成本主表查看。") }
        }
        return ($lines -join "`n")
    }
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add("【宠速星经营助手｜${modeName}同步完成】")
    $lines.Add("本月订单毛利（暂估）：$(Format-Money $Snapshot.metrics['本月经营毛利'])")
    $lines.Add("今日订单：自然 $($Snapshot.today_natural) / 干预 $($Snapshot.today_intervention) / 取消 $($Snapshot.today_cancelled)")
    $lines.Add("重点补货：$(@($Snapshot.urgent_restock).Count) 个（库存≤1且自然月售估算≥1）")
    foreach ($item in @($Snapshot.urgent_restock | Select-Object -First 10)) {
        $lines.Add("- $($item.code)｜$((Limit-Text $item.name 30))｜库存 $($item.stock)｜自然月售估算 $($item.natural_monthly_sales)")
    }
    if (@($Snapshot.restore_inventory).Count -gt 0) { $lines.Add("待恢复库存：$(@($Snapshot.restore_inventory).Count) 个货号") }
    $lines.Add("驾驶舱：$($meta.workbook_url)")
    return ($lines -join "`n")
}

$orderResult = $null
$inventoryResult = $null
$ledgerResult = $null
$costResult = $null
$giftReconcileResult = $null
$preflightResult = $null
$promotionCapture = $null
$integrityAuditResult = $null
$useNoChangeFastPath = $false
$syncWarnings = [Collections.Generic.List[string]]::new()
$operationalDegraded = $false
try {
    $stage = Invoke-WithRetry 'preflight' { Test-SyncPreflight } 5
    $preflightResult = $stage.value
    $promotionCapture = Get-PromotionCaptureState
    if (-not $SkipSync) {
        $dependencyCheck = Invoke-WithRetry 'chrome_cdp' {
            $global:LASTEXITCODE = 0
            $output = @(& node $webAccessCheck 2>&1 | ForEach-Object { [string]$_ })
            if ($LASTEXITCODE -ne 0 -or -not ($output -match 'chrome: ok') -or -not ($output -match 'proxy: ready')) {
                throw "Chrome CDP 未连接：$(($output -join ' ').Trim())"
            }
            return ($output -join '; ')
        } 3
        # Browser availability is read-only. Never take the global finance lock
        # until Chrome is usable; otherwise an unavailable browser blocks owner-
        # confirmed writes even though this run cannot mutate valid data.
        Start-SyncFinanceTransaction
        if (-not $SkipOrders) {
            $stage = Invoke-WithRetry 'order_sync' { Invoke-JsonPowerShellScript $orderScript @{ StartDate = $dateText; EndDate = $dateText } } 3
            $orderResult = $stage.value
            # A successful authenticated order read resolves the persistent login incident.
            Resolve-FailureNotification 'meituan-auth'
            Resolve-FailureNotification 'meituan-context'
            foreach ($pendingRefund in @($orderResult.pending_refund_details)) {
                $syncWarnings.Add("退款待确认：订单 $($pendingRefund.order_no)｜$($pendingRefund.goods_code)×$($pendingRefund.quantity)｜顾客退款 ¥$($pendingRefund.refund_amount)；收入已扣减，货物成本暂保留。请告诉经营管家商品是‘已退回入库’、‘未退货’还是‘退回报损’。")
            }
        }
        if (-not $SkipInventory) {
            try {
                $stage = Invoke-WithRetry 'inventory_sync' { Invoke-JsonPowerShellScript $inventoryScript @{ SnapshotDate = $Date } } 2
                $inventoryResult = $stage.value
                if ([int]$inventoryResult.new_kw_count -gt 0) {
                    $syncWarnings.Add("发现 $([int]$inventoryResult.new_kw_count) 个新KW，已加入SKU成本主表并标记为缺少成本，请补充单位成本。")
                }
                if ($inventoryResult.degraded) {
                    $operationalDegraded = $true
                    $syncWarnings.Add("库存导出未取得新文件，已使用最近一次有效文件；补货清单可能不是零点实时库存。")
                }
            } catch {
                # 无法证明失败发生在任何写入之前，因此禁止带着可能的半成品继续提交日结。
                # 准确性优先于“勉强发出一份旧库存日结”。外层事务会恢复全部已写范围。
                throw
            }
        }
    } else {
        Start-SyncFinanceTransaction
    }

    if ($Mode -eq 'night') {
        $null = Invoke-WithRetry 'intervention_reconciliation' { Assert-TargetDateInterventionCoverage } 1
    }
    $stage = Invoke-WithRetry 'pending_gift_reconciliation' {
        Invoke-JsonPowerShellScript $giftReconcileScript @{ ProjectRoot=$projectRoot; GraceHours=24; DryRun=[bool]$DryRun }
    } 3
    $giftReconcileResult = $stage.value
    foreach ($orderNo in @($giftReconcileResult.cancelled_orders)) {
        $syncWarnings.Add("赠品对应订单已取消：$orderNo；赠品成本未计入，请确认礼物是否实际送出。")
    }
    foreach ($orderNo in @($giftReconcileResult.overdue_orders)) {
        $syncWarnings.Add("赠品等待订单同步已超过24小时：$orderNo；请核对订单编号。")
    }
    foreach ($orderNo in @($giftReconcileResult.ambiguous_orders)) {
        $syncWarnings.Add("赠品订单匹配不唯一：$orderNo；系统未自动绑定。")
    }
    $useNoChangeFastPath = Test-UseNoChangeFastPath $Mode $orderResult $giftReconcileResult ([bool]$SkipSync) ([bool]$SkipOrders)
    if ($useNoChangeFastPath) {
        foreach ($skippedStage in @('cost_snapshot','formula_verification','daily_ledger','integrity_audit')) {
            $stageTelemetry[$skippedStage] = [ordered]@{seconds=0.0;attempts=0;skipped=$true;reason='订单与赠品数据无变化'}
        }
    } else {
        $eligibleCostOrders = if ($null -ne $orderResult) { @($orderResult.new_item_order_nos) } else { @() }
        $stage = Invoke-WithRetry 'cost_snapshot' { Invoke-CostSnapshotRebuild $eligibleCostOrders } 3
        $costResult = $stage.value
        $null = Invoke-WithRetry 'formula_verification' {
            Assert-FormulaRange '订单汇总' 'E2:E2001,K2:L2001'
            Assert-FormulaRange '订单商品' 'E2:E4001,K2:K4001,P2:P4001'
            Assert-FormulaRange '库存快照' 'H2:I2000,K2:K2000,M2:M2000'
            Assert-FormulaRange '补货记录' 'D2:E1001,G2:G1001'
            Assert-FormulaRange '驾驶舱' 'C2:C19'
            return 'success'
        } 3

        $stage = Invoke-WithRetry 'daily_ledger' { Invoke-DailyLedgerSync } 3
        $ledgerResult = $stage.value
        # 飞书不提供跨工作簿原生事务，因此在提交前做一次全账交叉核对。
        # 任一重复、漏单、干预单泄漏、配送费丢失或利润算术错误都会触发整批回滚。
        $stage = Invoke-WithRetry 'integrity_audit' { Invoke-FinanceIntegrityAudit } 2
        $integrityAuditResult = $stage.value
    }
    # 当日利润与推广费是账目提交条件；通知卡片中的库存、临期和提醒信息不是。
    # 先完成关键财务回读，避免纯通知读取故障把已通过完整性审计的整套账目回滚。
    $stage = Invoke-WithRetry 'daily_profit_readback' { Get-DailyNetProfit } 3
    $dailyProfit = $stage.value

    if ($null -ne $promotionCapture) {
        $capturedAmount = [math]::Round([double]$promotionCapture.amount, 2)
        if ([math]::Abs([double]$dailyProfit.promotion_expense - $capturedAmount) -gt 0.01) {
            throw "推广费账表与抓取结果不一致：抓取 ¥$capturedAmount，$($Date.Month)月支出合计 ¥$($dailyProfit.promotion_expense)；已停止日结"
        }
    }

    # 所有财务写入、公式校验、完整性审计和关键金额回读均成功后立即提交。
    # 后续摘要或通知失败只报告后置异常，不得反向破坏已提交账目。
    Complete-SyncFinanceTransaction
    $stage = Invoke-WithRetry 'summary_readback' { Build-Snapshot } 3
    $snapshot = $stage.value
    $snapshot | Add-Member -NotePropertyName today_order_profit -NotePropertyValue $dailyProfit.order_profit
    $snapshot | Add-Member -NotePropertyName today_expense_total -NotePropertyValue $dailyProfit.expense_total
    $snapshot | Add-Member -NotePropertyName today_promotion_expense -NotePropertyValue $dailyProfit.promotion_expense
    $snapshot | Add-Member -NotePropertyName today_net_profit -NotePropertyValue $dailyProfit.net_profit
    $snapshot | Add-Member -NotePropertyName sync_warnings -NotePropertyValue @($syncWarnings)
    $snapshot | Add-Member -NotePropertyName sync_degraded -NotePropertyValue ([bool]$operationalDegraded)
    $message = Build-Message $snapshot $true
    $card = if ($Mode -eq 'reminder') { $null } else { Build-Card $snapshot $true }
    $key = if ($ForceResend) {
        "petstar-$($Date.ToString('yyyyMMdd'))-$Mode-resend-$([datetime]::Now.ToString('yyyyMMddHHmmss'))"
    } elseif ($Mode -eq 'manual') {
        "petstar-$($Date.ToString('yyyyMMdd'))-manual-$([datetime]::Now.ToString('HHmmss'))"
    } else {
        "petstar-$($Date.ToString('yyyyMMdd'))-$Mode"
    }
    $giftAlertResult = $null
    $giftAlerts = @($syncWarnings | Where-Object { $_ -match '^赠品' })
    if ($Silent -and $giftAlerts.Count -gt 0) {
        $alertText = "🎁 赠品记录需要处理`n" + ($giftAlerts -join "`n")
        $alertKey = "petstar-gift-alert-$($Date.ToString('yyyyMMdd'))"
        $alertStage = Invoke-WithRetry 'pending_gift_alert' { Send-FeishuMessage $alertText $null $alertKey } 2
        $giftAlertResult = $alertStage.value
    }
    $sendResult = if ($Silent) {
        [pscustomobject]@{ sent = $false; dry_run = [bool]$DryRun; format = 'silent' }
    } else {
        $stage = Invoke-WithRetry 'notification' { Send-FeishuMessage $message $card $key } 2
        $stage.value
    }
    $result = [ordered]@{
        ok = $true
        degraded = [bool]$operationalDegraded
        sync_skipped = [bool]$SkipSync
        date = $dateText
        mode = $Mode
        fast_path = [ordered]@{
            used = [bool]$useNoChangeFastPath
            reason = if ($useNoChangeFastPath) { '订单与赠品数据无变化，已跳过成本重建、全量公式校验、每日简账刷新和完整性全量审计' } else { $null }
            skipped_stages = if ($useNoChangeFastPath) { @('cost_snapshot','formula_verification','daily_ledger','integrity_audit') } else { @() }
        }
        sent = [bool]$sendResult.sent
        dry_run = [bool]$sendResult.dry_run
        preflight = $preflightResult
        order_sync = $orderResult
        inventory_sync = $inventoryResult
        pending_gift_reconciliation = $giftReconcileResult
        pending_gift_alert = $giftAlertResult
        cost_snapshot_rebuild = $costResult
        daily_ledger_sync = $ledgerResult
        integrity_audit = $integrityAuditResult
        today_orders = [ordered]@{ total=$snapshot.today_total; natural=$snapshot.today_natural; platform_natural=$snapshot.today_platform_natural; intervention=$snapshot.today_intervention; cancelled=$snapshot.today_cancelled; offline=$snapshot.today_offline }
        today_profit = [ordered]@{ order_profit=$snapshot.today_order_profit; expenses=$snapshot.today_expense_total; promotion_expense=$snapshot.today_promotion_expense; net_profit=$snapshot.today_net_profit }
        estimated_order_gross_profit = Convert-ToNumber $snapshot.metrics['本月经营毛利']
        urgent_restock_count = @($snapshot.urgent_restock).Count
        expiring_product_count = @($snapshot.expiring_products).Count
        expiry_cutoff = $snapshot.expiry_cutoff
        restore_inventory_count = @($snapshot.restore_inventory).Count
        pending_delivery_count = @($snapshot.pending_delivery).Count
        pending_cost_count = @($snapshot.pending_cost).Count
        formula_verification = if ($useNoChangeFastPath) { 'skipped_no_change' } else { 'success' }
        message_format = $sendResult.format
        card_error = $sendResult.card_error
        warnings = @($syncWarnings)
        stages = $stageTelemetry
        duration_seconds = [math]::Round(([datetime]::Now - $runStartedAt).TotalSeconds, 2)
        message_preview = if ($DryRun) { [ordered]@{ text = $message; card = $card } } else { $null }
        completed_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    }
    $safeResult = Convert-ToSafeState $result
    Save-RunState $safeResult
    try { $syncMutex.ReleaseMutex() } catch {}
    $syncMutex.Dispose()
    $safeResult | ConvertTo-Json -Depth 12
} catch {
    $failureReason = $_.Exception.Message
    $rollback = Rollback-SyncFinanceTransaction $failureReason
    $postCommitFailure = [bool]$script:financeDataCommitted
    if ($null -ne $rollback) {
        if ($rollback.ok) { $failureReason = "$failureReason；本次数据写入已全部回滚" }
        else { $failureReason = "$failureReason；自动回滚尚未完成，事务 $($rollback.transaction_id) 已阻止后续写入：$($rollback.error)" }
    } elseif ($postCommitFailure) {
        $failureReason = "$failureReason；经营数据此前已校验并提交，本次仅后续通知或状态记录失败"
    }
    if ($postCommitFailure) {
        # 账目已经提交时，绝不能再发“同步失败/未完成”误导店主重复操作。
        $failureMessage = "【宠速星经营助手｜账目已同步】`n时间：$([datetime]::Now.ToString('yyyy-MM-dd HH:mm'))`n账目状态：已校验并提交，无需重复录入。`n后续异常：$failureReason"
        $failureCard = $null
    } else {
        $failureMessage = Build-Message $null $false $failureReason
        $failureCard = Build-Card $null $false $failureReason
    }
    $failureSent = $false
    $failureNotificationSuppressed = $false
    try {
        $failureCategory = Get-FailureNotificationCategory $failureReason
        $failureKey = "petstar-failed-$($Date.ToString('yyyyMMdd'))-$Mode-$failureCategory"
        $failureReceiptPath = Get-FailureNotificationReceiptPath $failureCategory
        if (-not $ForceResend -and (Test-FailureNotificationReceipt $failureReceiptPath $failureCategory)) {
            $failureNotificationSuppressed = $true
        } else {
            $failureSend = Send-FeishuMessage $failureMessage $failureCard $failureKey
            $failureSent = [bool]$failureSend.sent
            if ($failureSent -and -not $DryRun) {
                Save-FailureNotificationReceipt $failureReceiptPath $failureCategory $failureKey
            }
        }
    } catch {
        $notificationError = $_.Exception.Message
        $failureReason = "$failureReason；飞书失败通知也未送达：$notificationError"
    }
    $failure = [ordered]@{
        ok = $false
        sync_skipped = [bool]$SkipSync
        date = $dateText
        mode = $Mode
        sent = $failureSent
        notification_suppressed = $failureNotificationSuppressed
        data_committed = $postCommitFailure
        error = $failureReason
        stages = $stageTelemetry
        duration_seconds = [math]::Round(([datetime]::Now - $runStartedAt).TotalSeconds, 2)
        completed_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    }
    $safeFailure = Convert-ToSafeState $failure
    Save-RunState $safeFailure
    try { $syncMutex.ReleaseMutex() } catch {}
    $syncMutex.Dispose()
    $safeFailure | ConvertTo-Json -Depth 8
    exit 1
}
