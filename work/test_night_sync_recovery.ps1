$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'agent\runtime_paths.ps1')
$saved=@{}
foreach($key in @('HERMES_HOME','OPENCLAW_HOME','LARK_CHANNEL','LARKSUITE_CLI_CONFIG_DIR')){$saved[$key]=[Environment]::GetEnvironmentVariable($key,'Process')}
$temp=Join-Path ([IO.Path]::GetTempPath()) ('petstar-child-'+[guid]::NewGuid().ToString('N')+'.ps1')
$nativeFixture=[IO.Path]::ChangeExtension($temp,'.cmd')
$cases=0
try {
    $env:HERMES_HOME='offline-fixture';$env:OPENCLAW_HOME='offline-fixture';$env:LARK_CHANNEL='offline-fixture'
    Initialize-PetstarLarkUserEnvironment
    if($env:HERMES_HOME -or $env:OPENCLAW_HOME -or $env:LARK_CHANNEL){throw 'Host routing markers survived'};$cases++
    if($env:LARKSUITE_CLI_CONFIG_DIR -ne (Join-Path (Get-PetstarUserProfile) '.lark-cli')){throw 'User credential directory was not pinned'};$cases++
    $run=Join-Path $root 'agent\run_sync_and_notify.ps1'
    $ast=[Management.Automation.Language.Parser]::ParseFile($run,[ref]$null,[ref]$null)
    $f=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-JsonPowerShellScript'},$true)
    Invoke-Expression $f.Extent.Text
    $workRoot=$root
    [IO.File]::WriteAllText($temp,"param(`$WorkRoot)`nthrow 'inventory nested sentinel'",[Text.UTF8Encoding]::new($true))
    $message='';try{$null=Invoke-JsonPowerShellScript $temp @{}}catch{$message=$_.Exception.Message}
    if($message -notmatch 'inventory nested sentinel'){throw 'Nested inventory reason was lost'};$cases++
    [IO.File]::WriteAllText($temp,'param($WorkRoot) [ordered]@{ok=$true;value=7}|ConvertTo-Json -Compress',[Text.UTF8Encoding]::new($true))
    if((Invoke-JsonPowerShellScript $temp @{}).value -ne 7){throw 'Successful child JSON changed'};$cases++
    foreach($path in @('work\sync_meituan_inventory.ps1','agent\sync_sku_master_by_identity.ps1')) {
        $text=[IO.File]::ReadAllText((Join-Path $root $path),[Text.Encoding]::UTF8)
        $err=$null;$null=[Management.Automation.Language.Parser]::ParseInput($text,[ref]$null,[ref]$err)
        if(@($err).Count -or $text -match '(?m)^\s*trap\s*\{' -or $text -notmatch 'throw \$originalError'){throw "Nested failure must propagate: $path"};$cases++
        if($text -notmatch 'Get-Content -LiteralPath \$metaPath -Raw -Encoding UTF8'){throw 'Inventory binding must read UTF8'};$cases++
    }
    $watchdog=[IO.File]::ReadAllText((Join-Path $root 'agent\ensure_lark_event_bridge.ps1'))
    if($watchdog -match 'Start-Process -FilePath \$hermesPython -ArgumentList' -or $watchdog -notmatch "'-Job','NightRecovery'"){throw 'Watchdog still uses legacy recovery workspace'};$cases++
    $delivery=[IO.File]::ReadAllText((Join-Path $root 'agent\sync_delivery_fees.ps1'))
    if($delivery -notmatch '\$day -ne \$previousDay'){throw 'Queue-free previous day recovery missing'};$cases++
    $recovery=Join-Path $root 'work\recover_night_transaction_20261105.ps1'
    $recoveryAst=[Management.Automation.Language.Parser]::ParseFile($recovery,[ref]$null,[ref]$null)
    $signature=$recoveryAst.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Recovery-CellSignature'},$true)
    Invoke-Expression $signature.Extent.Text
    if((Recovery-CellSignature ([ordered]@{value=12.5})) -ne (Recovery-CellSignature ([pscustomobject]@{value='12.5'}))){throw 'Rollback verification dictionary mismatch'};$cases++
    if((Recovery-CellSignature ([ordered]@{formula='=1+1'})) -ne (Recovery-CellSignature ([pscustomobject]@{formula='=1+1';value=2}))){throw 'Rollback verification formula mismatch'};$cases++
    $correctionLoop=$ast.Find({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Variable.VariablePath.UserPath -eq 'change' -and $n.Condition.Extent.Text -match 'deliveryFeeResult.changed_orders'},$true)
    $deliveryFeeResult=$null;$dateText='2026-11-05';$syncWarnings=[Collections.Generic.List[string]]::new()
    Invoke-Expression $correctionLoop.Extent.Text
    if($syncWarnings.Count -ne 0){throw 'Read-only resend fabricated a blank-order fee correction'};$cases++
    $deliveryFeeResult=[pscustomobject]@{changed_orders=@([pscustomobject]@{order_no='3402374350443155017';date='2026-11-04';amount=5.6})}
    Invoke-Expression $correctionLoop.Extent.Text
    if($syncWarnings.Count -ne 1 -or $syncWarnings[0] -notmatch '3402374350443155017'){throw 'Actual cross-day correction was lost'};$cases++
    $authAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'agent\authorize_lark_desktop.ps1'),[ref]$null,[ref]$null)
    $authFunction=$authAst.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-DesktopLarkJson'},$true)
    Invoke-Expression $authFunction.Extent.Text
    [IO.File]::WriteAllText($nativeFixture,"@echo off`r`necho Waiting for approval 1>&2`r`necho {`"ok`":true,`"fixture`":7}`r`nexit /b 0`r`n",[Text.Encoding]::ASCII)
    $larkCli=$nativeFixture
    if((Invoke-DesktopLarkJson @('auth','status')).value.fixture -ne 7){throw 'Native authorization progress interrupted JSON handling'};$cases++
    [ordered]@{ok=$true;cases=$cases;external_writes=0}|ConvertTo-Json -Compress
}finally {
    foreach($key in $saved.Keys){[Environment]::SetEnvironmentVariable($key,$saved[$key],'Process')}
    if(Test-Path -LiteralPath $temp){Remove-Item -LiteralPath $temp}
    if(Test-Path -LiteralPath $nativeFixture){Remove-Item -LiteralPath $nativeFixture}
}
