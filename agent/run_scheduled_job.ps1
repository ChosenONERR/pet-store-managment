param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Hourly','Reminder','Promotion','Night','Weekly','NightRecovery')]
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
foreach ($variableName in @(
    'HERMES_HOME','HERMES_QUIET','HERMES_EXEC_ASK','HERMES_GATEWAY_TOKEN','HERMES_SESSION_KEY',
    'OPENCLAW_CLI','OPENCLAW_HOME','OPENCLAW_STATE_DIR','OPENCLAW_CONFIG_PATH',
    'OPENCLAW_SERVICE_MARKER','OPENCLAW_SERVICE_VERSION','OPENCLAW_GATEWAY_PORT','OPENCLAW_SHELL',
    'LARK_CHANNEL'
)) {
    Remove-Item -LiteralPath "Env:$variableName" -ErrorAction SilentlyContinue
}
$env:LARKSUITE_CLI_CONFIG_DIR = Join-Path $env:USERPROFILE '.lark-cli'

. (Join-Path $PSScriptRoot 'runtime_paths.ps1')
$powershellExe = Get-PetstarWindowsPowerShell
$syncScript = Join-Path $PSScriptRoot 'run_sync_and_notify.ps1'
$promotionScript = Join-Path $PSScriptRoot 'sync_promotion_expense.ps1'
$weeklyScript = Join-Path $PSScriptRoot 'run_weekly_maintenance.ps1'
$statePath = Join-Path $ProjectRoot 'agent\state\last_run.json'

function Invoke-Script([string]$Path, [string[]]$Arguments) {
    $global:LASTEXITCODE = 0
    & $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Path @Arguments
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

$today = [datetime]::Today
switch ($Job) {
    'Hourly' {
        Invoke-Script $syncScript @('-Date',$today.ToString('yyyy-MM-dd'),'-Mode','midday','-SkipInventory','-Silent','-ProjectRoot',$ProjectRoot)
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
    }
    'Weekly' {
        # 周一 00:00 即周日晚结束后：先完成周日常规日结，再做全历史只读审计。
        # 只有审计失败时，weekly 脚本才调用全量重建。
        $targetDate = $today.AddDays(-1).ToString('yyyy-MM-dd')
        Invoke-Script $syncScript @('-Date',$targetDate,'-Mode','night','-ProjectRoot',$ProjectRoot)
        Invoke-Script $weeklyScript @('-Date',$targetDate,'-ProjectRoot',$ProjectRoot)
    }
    'NightRecovery' {
        $targetDate = $today.AddDays(-1).ToString('yyyy-MM-dd')
        $healthy = $false
        if (Test-Path -LiteralPath $statePath -PathType Leaf) {
            try {
                $lastRun = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
                $healthy = ([string]$lastRun.date -eq $targetDate -and [string]$lastRun.mode -eq 'night' -and
                    [bool]$lastRun.ok -and -not [bool]$lastRun.degraded)
            } catch {}
        }
        if ($healthy) {
            [ordered]@{ok=$true;skipped=$true;reason='night sync already healthy';date=$targetDate} | ConvertTo-Json -Compress
            exit 0
        }
        Invoke-Script $syncScript @('-Date',$targetDate,'-Mode','night','-ProjectRoot',$ProjectRoot)
    }
}
