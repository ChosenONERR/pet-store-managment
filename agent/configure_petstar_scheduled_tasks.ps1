param(
    [string]$ProjectRoot,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$launcher = Join-Path $ProjectRoot 'agent\run_hidden_scheduled_job.vbs'
$watchdogLauncher = Join-Path $ProjectRoot 'agent\run_hidden_watchdog.vbs'
$wscript = Join-Path $env:SystemRoot 'System32\wscript.exe'
if (-not (Test-Path -LiteralPath $launcher -PathType Leaf)) { throw "Scheduled job launcher not found: $launcher" }
if (-not (Test-Path -LiteralPath $watchdogLauncher -PathType Leaf)) { throw "Watchdog launcher not found: $watchdogLauncher" }
if (-not (Test-Path -LiteralPath $wscript -PathType Leaf)) { throw "Windows Script Host not found: $wscript" }

$specs = @(
    [ordered]@{Name='Petstar-Hourly-Sync';Job='Hourly';At='09:00-22:00 hourly'},
    [ordered]@{Name='Petstar-Input-Reminder';Job='Reminder';At='23:00'},
    [ordered]@{Name='Petstar-Promotion-Expense';Job='Promotion';At='23:55'},
    [ordered]@{Name='Petstar-Night-Sync';Job='Night';At='00:00'},
    [ordered]@{Name='Petstar-Night-Recovery-0020';Job='NightRecovery';At='00:20'},
    [ordered]@{Name='Petstar-Night-Recovery-0040';Job='NightRecovery';At='00:40'}
)

if (-not $Apply) {
    [ordered]@{ok=$true;dry_run=$true;message='No system configuration changed. Re-run with -Apply after explicit confirmation.';tasks=$specs} | ConvertTo-Json -Depth 5
    exit 0
}

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 60) -MultipleInstances IgnoreNew
$installed = [Collections.Generic.List[object]]::new()
$watchdogAction = New-ScheduledTaskAction -Execute $wscript -Argument ('//B //Nologo "{0}"' -f $watchdogLauncher)
$watchdogTrigger = New-ScheduledTaskTrigger -Once -At ([datetime]::Now.AddMinutes(1)) -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days 3650)
$null = Register-ScheduledTask -TaskName 'Petstar-Lark-Event-Watchdog' -Action $watchdogAction -Trigger $watchdogTrigger -Settings $settings -Description 'Petstar Lark listener watchdog (hidden window)' -Force
$watchdogTask = Get-ScheduledTask -TaskName 'Petstar-Lark-Event-Watchdog' -ErrorAction Stop
$installed.Add([ordered]@{name='Petstar-Lark-Event-Watchdog';state=[string]$watchdogTask.State;job='Watchdog';at='every 5 minutes'})
foreach ($spec in $specs) {
    $action = New-ScheduledTaskAction -Execute $wscript -Argument ('//B //Nologo "{0}" {1}' -f $launcher,$spec.Job)
    if ($spec.Job -eq 'Hourly') {
        $trigger = @(9..22 | ForEach-Object { New-ScheduledTaskTrigger -Daily -At ('{0:D2}:00' -f $_) })
    } else {
        $trigger = New-ScheduledTaskTrigger -Daily -At $spec.At
    }
    $null = Register-ScheduledTask -TaskName $spec.Name -Action $action -Trigger $trigger -Settings $settings -Description 'Petstar deterministic finance job (hidden window)' -Force
    $actual = Get-ScheduledTask -TaskName $spec.Name -ErrorAction Stop
    $installed.Add([ordered]@{name=$spec.Name;state=[string]$actual.State;job=$spec.Job;at=$spec.At})
}

$legacyHermesDisabled = $false
$legacyJobsPath = Join-Path $env:LOCALAPPDATA 'hermes\profiles\viral-video-analyzer\cron\jobs.json'
if (Test-Path -LiteralPath $legacyJobsPath -PathType Leaf) {
    try {
        $legacyJobNames = @('petstar-hourly-sync','petstar-night-sync','petstar-input-reminder','petstar-night-recovery','petstar-promotion-expense')
        $legacyJobs = @(([IO.File]::ReadAllText($legacyJobsPath,[Text.Encoding]::UTF8) | ConvertFrom-Json).jobs | Where-Object { $_.name -in $legacyJobNames })
        $legacyHermesDisabled = ($legacyJobs.Count -eq $legacyJobNames.Count -and @($legacyJobs | Where-Object enabled).Count -eq 0)
    } catch {}
}

[ordered]@{ok=$true;dry_run=$false;launcher=$launcher;watchdog_launcher=$watchdogLauncher;tasks=@($installed);legacy_hermes_scheduler_disabled=$legacyHermesDisabled} | ConvertTo-Json -Depth 5
