$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'agent\runtime_paths.ps1')
$script:cases = 0
function Check([bool]$Condition,[string]$Label) {
    if (-not $Condition) { throw $Label }; $script:cases++
}
$date = '2026-11-05'
$run = [pscustomobject]@{date=$date; mode='night'; ok=$true; degraded=$true; recovery_required=$false; dry_run=$false; sync_skipped=$false}
Check (Test-PetstarNightCompleted $run $date) 'Business review caused retry'
$run.recovery_required=$true
Check (-not (Test-PetstarNightCompleted $run $date)) 'Technical degradation skipped retry'
$run.recovery_required=$false; $run.dry_run=$true
Check (-not (Test-PetstarNightCompleted $run $date)) 'Dry run treated as complete'
$run.dry_run=$false; $run.sync_skipped=$true
Check (-not (Test-PetstarNightCompleted $run $date)) 'Read-only resend treated as complete'
$run.sync_skipped=$false; $run.ok=$false
Check (-not (Test-PetstarNightCompleted $run $date)) 'Failed write treated as complete'
$run.ok=$true
Check (-not (Test-PetstarNightCompleted $run '2026-11-06')) 'Wrong date skipped retry'
$run.mode='midday'
Check (-not (Test-PetstarNightCompleted $run $date)) 'Hourly run treated as night'
$run.mode='night'
$run | Add-Member orders_skipped $true
Check (-not (Test-PetstarNightCompleted $run $date)) 'Skipped order extraction treated as full close'
$run.orders_skipped=$false
$run | Add-Member inventory_skipped $true
Check (-not (Test-PetstarNightCompleted $run $date)) 'Skipped inventory treated as full close'
$run.inventory_skipped=$false
$legacy=[pscustomobject]@{date=$date;mode='night';ok=$true;degraded=$true}
Check (-not (Test-PetstarNightCompleted $legacy $date)) 'Unclassified old warning trusted'
$legacy.degraded=$false
Check (Test-PetstarNightCompleted $legacy $date) 'Healthy legacy record rejected'
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('petstar-night-policy-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixture
$completed=Join-Path $fixture 'last_completed_night.json'
try {
    Check (Test-PetstarNightRecoveryRequired $fixture $date) 'Missing record skipped retry'
    [IO.File]::WriteAllText($completed,($run|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
    Check (-not (Test-PetstarNightRecoveryRequired $fixture $date)) 'Completed review triggered retry'
    [IO.File]::WriteAllText($completed,'broken-json',[Text.UTF8Encoding]::new($false))
    Check (Test-PetstarNightRecoveryRequired $fixture $date) 'Corrupt completion accepted'
    $saveAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'agent\run_sync_and_notify.ps1'),[ref]$null,[ref]$null)
    $saveFunction=$saveAst.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Save-RunState'},$true)
    Invoke-Expression $saveFunction.Extent.Text
    $stateRoot=$fixture
    Save-RunState $run
    Check (-not (Test-PetstarNightRecoveryRequired $fixture $date)) 'Successful business-review night did not persist completion'
    $snapshot=[IO.File]::ReadAllText($completed)
    $run.mode='midday'; Save-RunState $run; $run.mode='night'
    Check ([IO.File]::ReadAllText($completed) -eq $snapshot) 'Hourly persistence overwrote independent completion'
} finally {
    if(Test-Path -LiteralPath $completed){Remove-Item -LiteralPath $completed}
    foreach($name in @('last_run.json','sync_runs.ndjson')) {
        $path=Join-Path $fixture $name
        if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path}
    }
    Remove-Item -LiteralPath $fixture
}
$syncPath=Join-Path $root 'agent\run_sync_and_notify.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($syncPath,[ref]$null,[ref]$null)
$preflight=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-SyncPreflight'},$true).Extent.Text
Check ($preflight -notmatch 'Get-LarkAuthStatus -Verify') 'Preflight still performs extra network verification'
Check ($preflight.IndexOf('sheets +cells-get') -lt $preflight.IndexOf('Get-LarkAuthStatus')) 'Status ran before read'
# Execute the actual preflight auth block with no CLI/network/credential access.
$authStart=$preflight.IndexOf('$auth = $null')
$authEnd=$preflight.IndexOf('    return [pscustomobject]', $authStart)
$authBlock=[scriptblock]::Create($preflight.Substring($authStart,$authEnd-$authStart))
$script:calls=[Collections.Generic.List[string]]::new()
$script:probeFails=$false; $script:metadataFails=$false
function Invoke-LarkJson {param([scriptblock]$Command,[string]$Label,[int]$MaxAttempts)
    $script:calls.Add('read'); if($script:probeFails){throw 'offline probe failure'}
    [pscustomobject]@{ok=$true}
}
function Get-LarkAuthStatus {
    $script:calls.Add('metadata'); if($script:metadataFails){throw 'offline status unavailable'}
    [pscustomobject]@{identities=[pscustomobject]@{user=[pscustomobject]@{status='ready';tokenStatus='valid';expiresAt='fixture'}}}
}
function Save-LarkUserIdentity($AuthStatus) {$script:calls.Add('identity')}
& $authBlock
Check (($script:calls -join ',') -eq 'read,metadata,identity') 'Read-first execution order changed'
$script:calls.Clear(); $script:metadataFails=$true
& $authBlock
Check (($script:calls -join ',') -eq 'read,metadata') 'Successful read blocked by metadata outage'
$script:calls.Clear();$script:probeFails=$true;$script:metadataFails=$false;$failed=$false
try{& $authBlock}catch{$failed=$_.Exception.Message -match 'offline probe failure'}
Check ($failed -and ($script:calls -join ',') -eq 'read,metadata') 'Read failure allowed continuation or lost cause'
$recipient=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-CurrentUserOpenId'},$true).Extent.Text
Check ($recipient -notmatch '-Verify') 'Notification lookup refreshes again'
$source=[IO.File]::ReadAllText($syncPath)
Check ($source -match 'recovery_required = \[bool\]\$nightRecoveryRequired' -and $source -match 'last_completed_night.json') 'Night outcome not persisted'
foreach($name in @('run_scheduled_job.ps1','ensure_lark_event_bridge.ps1')) {
    Check ([IO.File]::ReadAllText((Join-Path $root "agent\$name")) -match 'Test-PetstarNightRecoveryRequired') 'Recovery entry bypasses common policy'
}
$probe=[IO.File]::ReadAllText((Join-Path $root 'agent\check_lark_user_access.ps1'))
Check ($probe -match '\$OutputEncoding = \[Text.UTF8Encoding\]::new\(\$false\)' -and $probe -match '\[Console\]::OutputEncoding = \[Text.UTF8Encoding\]::new\(\$false\)') 'Background probe lost UTF8 native output encoding'
Check ($probe -match '--range A1:A1' -and $probe -notmatch 'auth login|cells-set|cells-clear|--verify') 'Diagnostic is not strictly read-only'
[ordered]@{ok=$true;cases=$script:cases;external_writes=0;credential_access=0}|ConvertTo-Json -Compress
