$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$installer=Join-Path $root 'agent\configure_petstar_scheduled_tasks.ps1'
$runner=Join-Path $root 'agent\run_scheduled_job.ps1'
$fakeStore=[pscustomobject]@{writes=0;triggers=@()};$script:cases=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:cases++}
# Shadow every task mutation. This test must never register a real task.
function Register-ScheduledTask {throw 'Unexpected unrelated task registration'}
function New-ScheduledTaskTrigger {
    param([switch]$Daily,[string]$At)
    [pscustomobject]@{StartBoundary=('2026-11-05T'+$At+':00+08:00');Enabled=$true}
}
function Get-ScheduledTask {
    param([string]$TaskName)
    Check ($TaskName -eq 'Petstar-Hourly-Sync') 'Installer touched an unrelated task'
    [pscustomobject]@{TaskName=$TaskName;Triggers=$fakeStore.triggers;Actions='preserve';Principal='preserve';Settings='preserve'}
}
function Set-ScheduledTask {
    param([string]$TaskName,$Trigger)
    $fakeStore.writes++;$fakeStore.triggers=@($Trigger)
    Check ($TaskName -eq 'Petstar-Hourly-Sync') 'Wrong task updated'
}
$preview=(& $installer -ProjectRoot $root -HourlyOnly | Out-String)|ConvertFrom-Json
Check ($preview.dry_run -and @($preview.tasks).Count -eq 1 -and $fakeStore.writes -eq 0) 'Preview changed system state'
$applied=(& $installer -ProjectRoot $root -Apply -HourlyOnly | Out-String)|ConvertFrom-Json
Check ($applied.ok -and $fakeStore.writes -eq 1 -and -not $applied.other_tasks_changed) 'Installer did not preserve other tasks'
$hours=@($fakeStore.triggers|ForEach-Object{([datetimeoffset]$_.StartBoundary).Hour})
Check (($hours -join ',') -eq ((1..22) -join ',')) 'Independent hour coverage is wrong'
$covered=@(@(0,23)+$hours|Sort-Object -Unique)
Check ($covered.Count -eq 24 -and ($covered -join ',') -eq ((0..23)-join ',')) 'Not all 24 hours are covered'
Check (0 -notin $hours -and 23 -notin $hours) 'Hourly task collides with midnight/reminder'

$ast=[Management.Automation.Language.Parser]::ParseFile($runner,[ref]$null,[ref]$null)
$hourly=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-CurrentDayHourlySync'},$true)
Invoke-Expression $hourly.Extent.Text
$dispatch=$ast.Find({param($n)$n -is [Management.Automation.Language.SwitchStatementAst]},$true)
$script:recordedCalls=[Collections.Generic.List[object]]::new();$script:failNight=$false
function Invoke-Script([string]$Path,[string[]]$Arguments){
    if($script:failNight -and 'night' -in $Arguments){throw 'offline-night-failure'}
    $script:recordedCalls.Add([pscustomobject]@{path=$Path;arguments=$Arguments})
}
function Scenario([string]$Job,[datetime]$AsOf,[string]$HealthPath=''){
    $today=$AsOf;$syncScript='sync';$weeklyScript='weekly';$dashboardMonthScript='dashboard'
    $trafficSuggestionScript='traffic';$promotionScript='promotion';$ProjectRoot='offline-root';$statePath=$HealthPath
    $script:recordedCalls.Clear()
    & ([scriptblock]::Create($dispatch.Extent.Text)) | Out-Null
    return ,$script:recordedCalls.ToArray()
}
foreach($day in @([datetime]'2026-11-05',[datetime]'2026-12-01',[datetime]'2027-01-01')){
    $calls=Scenario Night $day
    Check ($calls.Count -eq 3) 'Night did not run night/hourly/dashboard serially'
    Check ($calls[0].path -eq 'sync' -and 'night' -in $calls[0].arguments -and $day.AddDays(-1).ToString('yyyy-MM-dd') -in $calls[0].arguments) 'Wrong prior-day date'
    Check ($calls[1].path -eq 'sync' -and $day.ToString('yyyy-MM-dd') -in $calls[1].arguments -and 'midday' -in $calls[1].arguments -and '-SkipInventory' -in $calls[1].arguments -and '-Silent' -in $calls[1].arguments) 'New-day run is not lightweight/current-day'
    Check ($calls[2].path -eq 'dashboard') 'Dashboard did not run after new-day sync'
}
$calls=Scenario Weekly ([datetime]'2026-11-02')
Check (($calls.path -join ',') -eq 'sync,sync,weekly,dashboard') 'Weekly audit delayed midnight new-day sync'
$calls=Scenario Hourly ([datetime]'2026-11-05')
Check ($calls.Count -eq 1 -and 'midday' -in $calls[0].arguments -and '-SkipInventory' -in $calls[0].arguments) 'Ordinary hour became heavy'
$calls=Scenario Reminder ([datetime]'2026-11-05')
Check ($calls.Count -eq 1 -and 'reminder' -in $calls[0].arguments -and '-Silent' -notin $calls[0].arguments) '23:00 reminder changed'
$script:failNight=$true;$failed=$false
try{$null=Scenario Weekly ([datetime]'2026-11-02')}catch{$failed=$_.Exception.Message -match 'offline-night-failure'}
Check ($failed -and $script:recordedCalls.Count -eq 0) 'Night failure continued into hourly or audit'
$script:failNight=$false
$health=Join-Path ([IO.Path]::GetTempPath()) ('petstar-health-'+[guid]::NewGuid().ToString('N')+'.json')
try{
    [IO.File]::WriteAllText($health,'{"date":"2026-11-04","mode":"night","ok":true,"degraded":false}',[Text.UTF8Encoding]::new($false))
    $calls=Scenario NightRecovery ([datetime]'2026-11-05') $health
    Check ($calls.Count -eq 0) 'Healthy prior-day close repeated after hourly overwrite'
    [IO.File]::WriteAllText($health,'{"date":"2026-11-03","mode":"night","ok":true}',[Text.UTF8Encoding]::new($false))
    $calls=Scenario NightRecovery ([datetime]'2026-11-05') $health
    Check ($calls.Count -eq 2 -and 'night' -in $calls[0].arguments -and 'midday' -in $calls[1].arguments) 'Stale close did not recover then sync current day'
}finally{if(Test-Path -LiteralPath $health){Remove-Item -LiteralPath $health}}
$source=[IO.File]::ReadAllText($runner)
Check ($source -match 'agent\\state\\last_healthy_night.json' -and $source -notmatch 'agent\\state\\last_run.json') 'Recovery depends on mutable last-run record'
[ordered]@{ok=$true;cases=$script:cases;external_writes=0;actual_tasks_changed=0;hours_covered=24}|ConvertTo-Json -Compress
