param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Hourly','TrafficSuggestion','Reminder','Promotion','Night','Weekly','NightRecovery')]
    [string]$Job,
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path

# Lark CLI 1.0.98 isolates credentials by host workspace. A scheduled task can
# inherit HERMES_HOME from the desktop agent even though this job must use the
# owner's normal user authorization. Remove every host-routing marker in this
# child only, then pin the local config root; never copy or expose OAuth tokens.
. (Join-Path $PSScriptRoot 'runtime_paths.ps1')
Initialize-PetstarLarkUserEnvironment
$powershellExe = Get-PetstarWindowsPowerShell
$syncScript = Join-Path $PSScriptRoot 'run_sync_and_notify.ps1'
$promotionScript = Join-Path $PSScriptRoot 'sync_promotion_expense.ps1'
$weeklyScript = Join-Path $PSScriptRoot 'run_weekly_maintenance.ps1'
$trafficSuggestionScript = Join-Path $PSScriptRoot 'send_daily_traffic_suggestion.ps1'
$dashboardMonthScript = Join-Path $PSScriptRoot 'update_dashboard_month_selector.ps1'
$statePath = Join-Path $ProjectRoot 'agent\state\last_healthy_night.json'

function Invoke-Script([string]$Path, [string[]]$Arguments) {
    $global:LASTEXITCODE = 0
    & $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Path @Arguments
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

function Invoke-CurrentDayHourlySync {
    Invoke-Script $syncScript @('-Date',$today.ToString('yyyy-MM-dd'),'-Mode','midday','-SkipInventory','-Silent','-ProjectRoot',$ProjectRoot)
}

$today = [datetime]::Today
switch ($Job) {
    'Hourly' {
        Invoke-CurrentDayHourlySync
    }
    'TrafficSuggestion' {
        Invoke-Script $trafficSuggestionScript @('-Date',$today.ToString('yyyy-MM-dd'),'-ProjectRoot',$ProjectRoot)
    }
    'Reminder' {
        Invoke-Script $syncScript @('-Date',$today.ToString('yyyy-MM-dd'),'-Mode','reminder','-SkipInventory','-ProjectRoot',$ProjectRoot)
    }
    'Promotion' {
        Invoke-Script $promotionScript @('-Date',$today.ToString('yyyy-MM-dd'),'-ProjectRoot',$ProjectRoot)
    }
    'Night' {
        $targetDate = $today.AddDays(-1).ToString('yyyy-MM-dd')
        Invoke-Script $syncScript @('-Date',$targetDate,'-Mode','night','-ProjectRoot',$ProjectRoot)
        Invoke-CurrentDayHourlySync
        Invoke-Script $dashboardMonthScript @('-ProjectRoot',$ProjectRoot)
    }
    'Weekly' {
        # 周一 00:00 即周日晚结束后：先完成周日常规日结，再做全历史只读审计。
        # 只有审计失败时，weekly 脚本才调用全量重建。
        $targetDate = $today.AddDays(-1).ToString('yyyy-MM-dd')
        Invoke-Script $syncScript @('-Date',$targetDate,'-Mode','night','-ProjectRoot',$ProjectRoot)
        Invoke-CurrentDayHourlySync
        Invoke-Script $weeklyScript @('-Date',$targetDate,'-ProjectRoot',$ProjectRoot)
        Invoke-Script $dashboardMonthScript @('-ProjectRoot',$ProjectRoot)
    }
    'NightRecovery' {
        $targetDate = $today.AddDays(-1).ToString('yyyy-MM-dd')
        $healthy = -not (Test-PetstarNightRecoveryRequired (Split-Path $statePath -Parent) $targetDate)
        if ($healthy) {
            [ordered]@{ok=$true;skipped=$true;reason='night sync already healthy';date=$targetDate} | ConvertTo-Json -Compress
            break
        }
        Invoke-Script $syncScript @('-Date',$targetDate,'-Mode','night','-ProjectRoot',$ProjectRoot)
        Invoke-CurrentDayHourlySync
    }
}
