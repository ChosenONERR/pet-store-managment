param(
    [string]$ProjectRoot,
    [int]$ReadyTimeoutSeconds = 30
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path

$bridgePath = Join-Path $ProjectRoot 'agent\lark_event_bridge.ps1'
$codexBridgePath = Join-Path $ProjectRoot 'agent\codex_chat_bridge.py'
$stateRoot = Join-Path $ProjectRoot 'agent\state'
$watchdogLog = Join-Path $stateRoot 'lark_event_watchdog.log'
$hermesProfileRoot = Join-Path $env:LOCALAPPDATA 'hermes\profiles\viral-video-analyzer'
$gatewayPidPath = Join-Path $hermesProfileRoot 'gateway.pid'
$gatewayHeartbeatPath = Join-Path $hermesProfileRoot 'state\gateway.heartbeat'
$gatewayCommandPath = Join-Path $hermesProfileRoot 'gateway-service\Hermes_Gateway_viral-video-analyzer.cmd'
$nightRecoveryPath = Join-Path $hermesProfileRoot 'scripts\petstar_night_recovery.py'
$healthyNightPath = Join-Path $stateRoot 'last_healthy_night.json'
$nightRecoveryAttemptPath = Join-Path $stateRoot 'night_recovery_attempt.json'
$pendingAcknowledgementsPath = Join-Path $stateRoot 'pending_acknowledgements.json'
. (Join-Path $ProjectRoot 'agent\runtime_paths.ps1')
$powershellExe = Get-PetstarWindowsPowerShell
$codexPython = Get-PetstarPython
$hermesPython = $null
try { $hermesPython = Get-PetstarHermesPython } catch {}
$larkCli = Get-PetstarLarkCli

if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
if (-not (Test-Path -LiteralPath $bridgePath)) { throw "找不到监听脚本：$bridgePath" }
if (-not (Test-Path -LiteralPath $codexBridgePath)) { throw "找不到本机 Codex 桥接器：$codexBridgePath" }
if (-not (Test-Path -LiteralPath $codexPython)) { throw "找不到本机 Codex 桥接器 Python：$codexPython" }
if (-not (Test-Path -LiteralPath $larkCli)) { throw "找不到飞书 CLI：$larkCli" }

function Write-WatchdogLog([string]$Message) {
    Add-Content -LiteralPath $watchdogLog -Value "$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')) $Message" -Encoding UTF8
}

function Save-NightRecoveryAttempt([string]$TargetDate, [int]$ProcessId = 0) {
    $payload = [ordered]@{target_date=$TargetDate;attempted_at=[datetime]::Now.ToString('o');pid=$ProcessId} | ConvertTo-Json -Compress
    $tempPath = "$nightRecoveryAttemptPath.$([guid]::NewGuid().ToString('N')).tmp"
    [IO.File]::WriteAllText($tempPath,$payload,[Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tempPath -Destination $nightRecoveryAttemptPath -Force
}

function Test-NightRecoveryCooldown([string]$TargetDate, [int]$Minutes = 30) {
    if (-not (Test-Path -LiteralPath $nightRecoveryAttemptPath)) { return $false }
    try {
        $state = Get-Content -LiteralPath $nightRecoveryAttemptPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $attemptedAt = [datetimeoffset]::MinValue
        if ([string]$state.target_date -ne $TargetDate) { return $false }
        if (-not [datetimeoffset]::TryParse([string]$state.attempted_at,[ref]$attemptedAt)) { return $false }
        return ([datetimeoffset]::Now - $attemptedAt).TotalMinutes -lt $Minutes
    } catch { return $false }
}

function Flush-AcknowledgementOutbox {
    if (-not (Test-Path -LiteralPath $pendingAcknowledgementsPath)) { return 'empty' }
    $content = (Get-Content -LiteralPath $pendingAcknowledgementsPath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue).Trim()
    if (-not $content -or $content -eq '[]') { return 'empty' }
    try {
        $raw = & $powershellExe -NoProfile -ExecutionPolicy Bypass -File $bridgePath -ProjectRoot $ProjectRoot -Mode FlushAcks 2>$null | Out-String
        if ($LASTEXITCODE -eq 0 -and $raw) { return 'flushed' }
        Write-WatchdogLog 'acknowledgement outbox flush returned an error; queued replies remain for retry'
        return 'failed'
    } catch {
        Write-WatchdogLog 'acknowledgement outbox flush failed; queued replies remain for retry'
        return 'failed'
    }
}

function Start-ReplayWorker {
    $workers = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
        $_.CommandLine -and $_.CommandLine.IndexOf($bridgePath,[StringComparison]::OrdinalIgnoreCase) -ge 0 -and $_.CommandLine -match '(?i)-Mode\s+Replay'
    })
    if ($workers.Count -gt 0) { return 'running' }
    try {
        $arguments = @(
            '-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $bridgePath),
            '-ProjectRoot',('"{0}"' -f $ProjectRoot),'-Mode','Replay','-ReplayHours','48'
        )
        $null = Start-Process -FilePath $powershellExe -ArgumentList $arguments -WorkingDirectory $ProjectRoot -WindowStyle Hidden -PassThru
        return 'started'
    } catch {
        Write-WatchdogLog 'message replay worker could not be started'
        return 'failed'
    }
}

function Get-BridgeProcesses {
    return @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
        $_.CommandLine -and
        $_.CommandLine.IndexOf($bridgePath, [StringComparison]::OrdinalIgnoreCase) -ge 0 -and
        $_.CommandLine -notmatch '(?i)-Mode\s+(Replay|Inspect|Initialize|FlushAcks)'
    })
}

function Get-DescendantProcessIds([int[]]$RootProcessIds) {
    $allProcesses = @(Get-CimInstance Win32_Process)
    $known = [Collections.Generic.HashSet[int]]::new()
    $frontier = [Collections.Generic.Queue[int]]::new()
    foreach ($rootId in @($RootProcessIds)) {
        if ($rootId -gt 0 -and $known.Add($rootId)) { $frontier.Enqueue($rootId) }
    }
    while ($frontier.Count -gt 0) {
        $parentId = $frontier.Dequeue()
        foreach ($child in @($allProcesses | Where-Object { [int]$_.ParentProcessId -eq $parentId })) {
            if ($known.Add([int]$child.ProcessId)) { $frontier.Enqueue([int]$child.ProcessId) }
        }
    }
    return @($known)
}

function Repair-OrphanEventConsumers($BridgeProcesses, $EventHealth) {
    if ($null -eq $EventHealth) { return 0 }
    $bridgeIds = @($BridgeProcesses | ForEach-Object { [int]$_.ProcessId })
    $ownedIds = @(Get-DescendantProcessIds $bridgeIds)
    $orphanIds = @($EventHealth.consumer_pids | Where-Object { $_ -and [int]$_ -notin $ownedIds })
    if ($orphanIds.Count -eq 0) { return 0 }

    $allProcesses = @(Get-CimInstance Win32_Process)
    $stopped = [Collections.Generic.HashSet[int]]::new()
    foreach ($consumerId in $orphanIds) {
        $consumer = $allProcesses | Where-Object { [int]$_.ProcessId -eq [int]$consumerId } | Select-Object -First 1
        if ($null -eq $consumer -or $consumer.Name -ne 'lark-cli.exe' -or
            $consumer.CommandLine -notmatch '(?i)event\s+consume\s+im\.message\.receive_v1') { continue }

        $parent = $allProcesses | Where-Object { [int]$_.ProcessId -eq [int]$consumer.ParentProcessId } | Select-Object -First 1
        if ($null -ne $parent -and $parent.Name -eq 'node.exe' -and
            $parent.CommandLine -match '(?i)run\.js\s+event\s+consume\s+im\.message\.receive_v1') {
            try { Stop-Process -Id ([int]$parent.ProcessId) -Force -ErrorAction Stop; $null = $stopped.Add([int]$parent.ProcessId) } catch {}
        }
        try { Stop-Process -Id ([int]$consumer.ProcessId) -Force -ErrorAction Stop; $null = $stopped.Add([int]$consumer.ProcessId) } catch {}
    }
    if ($stopped.Count -gt 0) {
        Write-WatchdogLog "removed orphan event consumers process_ids=$(@($stopped) -join ',')"
        Start-Sleep -Seconds 2
    }
    return $stopped.Count
}

function Get-EventHealth {
    $authRaw = & $larkCli auth status --json 2>$null | Out-String
    if ($LASTEXITCODE -ne 0) { return $null }
    $auth = $authRaw | ConvertFrom-Json
    $statusRaw = & $larkCli event status --json 2>$null | Out-String
    if ($LASTEXITCODE -ne 0) { return $null }
    $status = $statusRaw | ConvertFrom-Json
    $app = @($status.apps | Where-Object { $_.app_id -eq $auth.appId } | Select-Object -First 1)
    if (-not $app) { return $null }
    return [pscustomobject]@{
        running = [bool]$app.running
        app_status = [string]$app.status
        active_consumers = [int]$app.active_consumers
        consumer_pids = @($app.consumers | ForEach-Object { [int]$_.pid })
        user_status = [string]$auth.identities.user.status
        bot_status = [string]$auth.identities.bot.status
    }
}

function Get-GatewayHealth {
    $pidValue = $null
    if (Test-Path -LiteralPath $gatewayPidPath) {
        $rawPid = (Get-Content -LiteralPath $gatewayPidPath -Raw -ErrorAction SilentlyContinue).Trim()
        if ($rawPid -match '^\d+$') {
            $pidValue = [int]$rawPid
        } else {
            try { $pidValue = [int](($rawPid | ConvertFrom-Json).pid) } catch {}
        }
    }
    $processAlive = $false
    if ($pidValue) { $processAlive = $null -ne (Get-Process -Id $pidValue -ErrorAction SilentlyContinue) }
    $heartbeatAge = [double]::PositiveInfinity
    if (Test-Path -LiteralPath $gatewayHeartbeatPath) {
        $heartbeatAge = ([datetime]::Now - (Get-Item -LiteralPath $gatewayHeartbeatPath).LastWriteTime).TotalSeconds
    }
    return [pscustomobject]@{
        pid = $pidValue
        process_alive = $processAlive
        heartbeat_age_seconds = [math]::Round($heartbeatAge, 1)
        healthy = $processAlive -and $heartbeatAge -le 180
    }
}

function Get-CodexHealth {
    try {
        $runtimeRaw = & $codexPython $codexBridgePath --ensure-runtime 2>$null | Out-String
        if ($LASTEXITCODE -ne 0) { return [pscustomobject]@{healthy=$false;runtime='';session_ready=$false;persistent_runtime=$false} }
        $raw = & $codexPython $codexBridgePath --check 2>$null | Out-String
        if ($LASTEXITCODE -ne 0) { return [pscustomobject]@{healthy=$false;runtime='';session_ready=$false;persistent_runtime=$false} }
        $status = $raw | ConvertFrom-Json
        return [pscustomobject]@{
            healthy = [bool]$status.ok
            runtime = [string]$status.runtime
            session_ready = [bool]$status.session_ready
            persistent_runtime = [bool]$status.persistent_runtime.ok
        }
    } catch {
        return [pscustomobject]@{healthy=$false;runtime='';session_ready=$false;persistent_runtime=$false}
    }
}

$gatewayAction = 'none'
$gatewayHealth = Get-GatewayHealth
if (-not $gatewayHealth.healthy) {
    if ($gatewayHealth.process_alive) {
        Write-WatchdogLog "gateway process alive but heartbeat stale pid=$($gatewayHealth.pid) age=$($gatewayHealth.heartbeat_age_seconds)s"
        $gatewayAction = 'stale-warning'
    } elseif (Test-Path -LiteralPath $gatewayCommandPath) {
        try {
            $started = Start-Process -FilePath 'cmd.exe' -ArgumentList @('/c', ('"{0}"' -f $gatewayCommandPath)) -WorkingDirectory $hermesProfileRoot -WindowStyle Hidden -PassThru
            $gatewayAction = 'start-requested'
            Write-WatchdogLog "legacy scheduler launch requested wrapper_pid=$($started.Id)"
            $gatewayDeadline = [datetime]::Now.AddSeconds($ReadyTimeoutSeconds)
            do {
                Start-Sleep -Seconds 1
                $gatewayHealth = Get-GatewayHealth
                if ($gatewayHealth.healthy) { break }
            } while ([datetime]::Now -lt $gatewayDeadline)
            if ($gatewayHealth.healthy) {
                $gatewayAction = 'started'
                Write-WatchdogLog "legacy scheduler ready pid=$($gatewayHealth.pid) heartbeat_age=$($gatewayHealth.heartbeat_age_seconds)s"
            } else {
                $gatewayAction = 'start-failed-warning'
                Write-WatchdogLog "legacy scheduler health check failed; Feishu/Codex bridge will continue pid=$($gatewayHealth.pid) alive=$($gatewayHealth.process_alive)"
            }
        } catch {
            $gatewayAction = 'start-error-warning'
            Write-WatchdogLog "legacy scheduler start error; Feishu/Codex bridge will continue"
        }
    } else {
        $gatewayAction = 'missing-warning'
        Write-WatchdogLog 'legacy scheduler command missing; Feishu/Codex bridge will continue'
    }
}

$codexHealth = Get-CodexHealth
if (-not $codexHealth.healthy) {
    Write-WatchdogLog 'local Codex health check failed; deterministic Feishu commands remain available and natural-language fallback will return an explicit error'
}

$catchUpAction = 'none'
$now = [datetime]::Now
if ($now.Hour -lt 9 -and ($now.Hour -gt 0 -or $now.Minute -ge 5)) {
    $targetNightDate = $now.Date.AddDays(-1).ToString('yyyy-MM-dd')
    $healthyNightDate = $null
    if (Test-Path -LiteralPath $healthyNightPath) {
        try { $healthyNightDate = [string]((Get-Content -LiteralPath $healthyNightPath -Raw | ConvertFrom-Json).date) } catch {}
    }
    if ($healthyNightDate -ne $targetNightDate) {
        if (Test-NightRecoveryCooldown $targetNightDate 30) {
            $catchUpAction = 'cooldown'
        }
        $syncProcesses = @(Get-CimInstance Win32_Process | Where-Object {
            $_.CommandLine -and $_.CommandLine -match '(?i)petstar_(night_sync|night_recovery)\.py|petstar_cli\.py\s+sync'
        })
        if ($catchUpAction -ne 'cooldown' -and $syncProcesses.Count -eq 0 -and (Test-Path -LiteralPath $nightRecoveryPath) -and (Test-Path -LiteralPath $hermesPython)) {
            Save-NightRecoveryAttempt $targetNightDate
            $started = Start-Process -FilePath $hermesPython -ArgumentList @(('"{0}"' -f $nightRecoveryPath)) -WorkingDirectory (Split-Path $nightRecoveryPath -Parent) -WindowStyle Hidden -PassThru
            Save-NightRecoveryAttempt $targetNightDate $started.Id
            $catchUpAction = 'started'
            Write-WatchdogLog "night catch-up launch requested pid=$($started.Id) target=$targetNightDate"
        }
    }
}

$processes = @(Get-BridgeProcesses)
$health = Get-EventHealth
$orphanConsumersRemoved = Repair-OrphanEventConsumers $processes $health
if ($orphanConsumersRemoved -gt 0) { $health = Get-EventHealth }
$ackOutboxAction = Flush-AcknowledgementOutbox
$replayAction = if ($null -ne $health -and $health.running) { Start-ReplayWorker } else { 'deferred' }
if ($processes.Count -gt 0 -and $null -ne $health -and $health.running) {
    [ordered]@{
        ok = $true
        action = 'none'
        process_count = $processes.Count
        active_consumers = $health.active_consumers
        orphan_consumers_removed = $orphanConsumersRemoved
        event_status = $health.app_status
        user_status = $health.user_status
        codex_ready = $codexHealth.healthy
        codex_runtime = $codexHealth.runtime
        codex_session_ready = $codexHealth.session_ready
        gateway_action = $gatewayAction
        gateway_pid = $gatewayHealth.pid
        gateway_heartbeat_age_seconds = $gatewayHealth.heartbeat_age_seconds
        night_catch_up_action = $catchUpAction
        acknowledgement_outbox = $ackOutboxAction
        message_replay = $replayAction
        checked_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    } | ConvertTo-Json -Compress
    exit 0
}

if ($processes.Count -eq 0) {
    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $bridgePath),
        '-ProjectRoot', ('"{0}"' -f $ProjectRoot)
    )
    $started = Start-Process -FilePath $powershellExe -ArgumentList $arguments -WorkingDirectory $ProjectRoot -WindowStyle Hidden -PassThru
    Write-WatchdogLog "bridge launch requested pid=$($started.Id)"
} else {
    Write-WatchdogLog "bridge process exists but event bus is not ready process_count=$($processes.Count)"
}

$deadline = [datetime]::Now.AddSeconds($ReadyTimeoutSeconds)
do {
    Start-Sleep -Seconds 1
    $health = Get-EventHealth
    if ($null -ne $health -and $health.running) { break }
} while ([datetime]::Now -lt $deadline)

$processes = @(Get-BridgeProcesses)
if ($null -eq $health -or -not $health.running -or $processes.Count -eq 0) {
    Write-WatchdogLog "bridge health check failed process_count=$($processes.Count)"
    throw '飞书事件监听器未在限定时间内就绪'
}

Write-WatchdogLog "bridge ready process_count=$($processes.Count) user_status=$($health.user_status)"
[ordered]@{
    ok = $true
    action = 'started'
    process_count = $processes.Count
    active_consumers = $health.active_consumers
    orphan_consumers_removed = $orphanConsumersRemoved
    event_status = $health.app_status
    user_status = $health.user_status
    bot_status = $health.bot_status
    codex_ready = $codexHealth.healthy
    codex_runtime = $codexHealth.runtime
    codex_session_ready = $codexHealth.session_ready
    gateway_action = $gatewayAction
    gateway_pid = $gatewayHealth.pid
    gateway_heartbeat_age_seconds = $gatewayHealth.heartbeat_age_seconds
    night_catch_up_action = $catchUpAction
    acknowledgement_outbox = $ackOutboxAction
    message_replay = $replayAction
    checked_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
} | ConvertTo-Json -Compress
