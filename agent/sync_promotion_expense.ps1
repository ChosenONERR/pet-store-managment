param(
    [datetime]$Date = [datetime]::Today,
    [switch]$DryRun,
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$projectRoot = $ProjectRoot
$stateRoot = Join-Path $projectRoot 'agent\state'
$promotionStateRoot = Join-Path $stateRoot 'promotion_expense'
$extractor = Join-Path $projectRoot 'work\meituan_promotion_extract.mjs'
$recordScript = Join-Path $projectRoot 'agent\record_finance_event.ps1'
$runtimePaths = Join-Path $projectRoot 'agent\runtime_paths.ps1'
. $runtimePaths
$powershellExe = Get-PetstarWindowsPowerShell
$dependencyCheck = Get-PetstarWebAccessCheck
$dateText = $Date.ToString('yyyy-MM-dd')

function Write-AtomicJson([string]$Path, $Value) {
    $directory = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $directory)) { $null = New-Item -ItemType Directory -Path $directory }
    $temporary = "$Path.tmp.$PID"
    [System.IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json -Depth 10), [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Invoke-JsonProcess([string]$Executable, [string[]]$Arguments, [string]$FailureLabel) {
    $global:LASTEXITCODE = 0
    $output = @(& $Executable @Arguments 2>&1 | ForEach-Object { [string]$_ })
    if ($LASTEXITCODE -ne 0) {
        $detail = (($output | Where-Object { $_ -and $_.Trim() }) -join ' ').Trim()
        throw "$FailureLabel：$detail"
    }
    $raw = ($output -join "`n").Trim()
    try { return ($raw | ConvertFrom-Json) }
    catch { throw "$FailureLabel：没有返回有效 JSON" }
}

try {
    $isRecoveryDate = $Date.Date -eq [datetime]::Today.AddDays(-1)
    if ($Date.Date -ne [datetime]::Today -and -not $isRecoveryDate) {
        throw "推广后台只能读取当天消费或前一天已结算消费：$dateText"
    }
    if (-not (Test-Path -LiteralPath $extractor -PathType Leaf)) { throw "缺少推广费读取脚本：$extractor" }
    $dependency = @(& node $dependencyCheck 2>&1 | ForEach-Object { [string]$_ })
    if ($LASTEXITCODE -ne 0 -or -not ($dependency -match 'chrome: ok') -or -not ($dependency -match 'proxy: ready')) {
        throw "Chrome CDP 未连接：$(($dependency -join ' ').Trim())"
    }

    $extraction = $null
    $errors = [Collections.Generic.List[string]]::new()
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $extraction = Invoke-JsonProcess 'node' @($extractor, '--date', $dateText) '美团推广费读取失败'
            break
        } catch {
            $attemptError = [string]$_.Exception.Message
            $errors.Add($attemptError)
            # A missing/stale DevTools endpoint does not improve by reopening the
            # promotion page three times. Stop immediately and let the scheduled
            # recovery try again after Chrome becomes reachable.
            if ($attemptError -match 'Could not connect to Chrome DevTools|Chrome DevToolsActivePort is invalid|CDP command timed out: Target\.attachToTarget|WebSocket 未连接|CDP proxy request failed') { break }
            if ($attempt -lt 3) { Start-Sleep -Seconds (2 * $attempt) }
        }
    }
    if ($null -eq $extraction) { throw ($errors -join '；') }
    if (-not $extraction.ok) { throw '美团推广费读取结果未标记成功' }
    $amount = [math]::Round([double]$extraction.amount, 2)
    if ($amount -lt 0) { throw "推广费不能为负数：$amount" }

    $writeResult = [pscustomobject]@{ status = 'zero'; changed = $false; amount = 0 }
    if ($amount -gt 0) {
        $expenseMemo = if ([bool]$extraction.recovery_capture) { '次日自动补抓（已结算推广消费）' } else { '23:55自动读取（当日累计消费）' }
        $arguments = @(
            '-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$recordScript,
            '-Action','OperatingExpense','-Date',$dateText,'-Amount',([string]$amount),'-Note','投流推广费',
            '-ReplaceMatchingProject','-ExpenseSource','美团推广','-ExpenseMemo',$expenseMemo,
            '-ProjectRoot',$projectRoot
        )
        if ($DryRun) { $arguments += '-DryRun' }
        $write = Invoke-JsonProcess $powershellExe $arguments '推广费写入失败'
        if (-not $write.ok) { throw '推广费经营事件未返回成功状态' }
        $writeResult = $write.result
    }

    $result = [ordered]@{
        ok = $true
        date = $dateText
        amount = $amount
        provisional = [bool]$extraction.provisional
        recovery_capture = [bool]$extraction.recovery_capture
        dry_run = [bool]$DryRun
        write_status = [string]$writeResult.status
        expense_sheet = $writeResult.expense_sheet
        expense_row = $writeResult.expense_row
        captured_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
        evidence = @($extraction.evidence)
    }
    if (-not $DryRun) {
        Write-AtomicJson (Join-Path $promotionStateRoot "$dateText.json") $result
        Write-AtomicJson (Join-Path $stateRoot 'promotion_expense_last_run.json') $result
    }
    $result | ConvertTo-Json -Depth 8
} catch {
    $reason = $_.Exception.Message
    $awaitingSettlement = [bool]($reason -match 'Promotion history row.+not settled yet|推广流水.+尚未结算')
    $failure = [ordered]@{
        ok=$false
        date=$dateText
        error=$reason
        status=$(if($awaitingSettlement){'awaiting_settlement'}else{'capture_failed'})
        notification=$(if($awaitingSettlement){'retry_after_settlement'}else{'deferred_to_night_sync'})
        failed_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    }
    if (-not $DryRun) {
        Write-AtomicJson (Join-Path $stateRoot 'promotion_expense_last_run.json') $failure
    }
    $failure | ConvertTo-Json -Depth 6
    # 23:55 时美团尚未结算当天流水属于可预期状态，不把计划任务标成故障。
    # 00:00 日结和 00:20/00:40 补偿任务仍会继续尝试；真正的抓取错误保持非零退出码。
    if ($awaitingSettlement) { exit 0 }
    exit 1
}
