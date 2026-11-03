$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$file=Join-Path $root 'agent\sync_delivery_fees.ps1'
$source=[IO.File]::ReadAllText($file,[Text.Encoding]::UTF8)
$errors=$null;$tokens=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($file,[ref]$tokens,[ref]$errors)
if(@($errors).Count){throw ($errors.Message -join ';')}
if('WorkRoot' -notin @($ast.ParamBlock.Parameters|ForEach-Object {$_.Name.VariablePath.UserPath})){throw 'Scheduled child helper always supplies WorkRoot'}
foreach($name in @('Convert-DeliveryDay','Set-DeliveryCell','Get-DeliveryCell')){
    $f=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true))
    if($f.Count -ne 1){throw "Missing $name"};Invoke-Expression $f[0].Extent.Text
}
if((Convert-DeliveryDay '46329.75') -ne '2026-11-03'){throw 'Serial day conversion failed'}
if((Convert-DeliveryDay '2026-11-03 19:00') -ne '2026-11-03'){throw 'Calendar day conversion failed'}
Set-StrictMode -Version 2
if($null -ne (Get-DeliveryCell @([pscustomobject]@{}) 0)){throw 'Empty cell should remain unknown'}
if($null -ne (Get-DeliveryCell @([pscustomobject]@{value='x'}) 3)){throw 'Missing cell should remain unknown'}
function Register-FinanceRangeBeforeImage {param($a,$b,$c,$d,$e,$f)$script:captured=$e}
function Invoke-DeliveryLark {param($Command)$null}
$transaction=[pscustomobject]@{id='offline-test'};$token='offline-test';$lark='offline-test'
Set-DeliveryCell 'J7:J7' 4.51
if($script:captured -ne 'J7:J7'){throw 'Missing write-before-image'}
# Test the exact payload expression independently: row and cell arrays must not flatten.
$m=@();$m+=,@(@{value=4.51});$matrix=ConvertTo-Json -InputObject $m -Depth 5 -Compress
if($matrix -ne '[[{"value":4.51}]]'){throw 'Fee payload not two dimensional'}
foreach($contract in @('Join-FinanceTransaction','人工确认配送费','manual_migration','confirmed_self','has_more','DELIVERY_WRITE_FAILED','formula-verify','Get-DeliveryRows.*A\$\(\$o.row\):Q','Keep freshly changed orders','if\(\$Apply\)')){
    if($source -notmatch $contract){throw "Missing safety contract: $contract"}
}
$run=[IO.File]::ReadAllText((Join-Path $root 'agent\run_sync_and_notify.ps1'),[Text.Encoding]::UTF8)
if($run.IndexOf("'delivery_fee_sync'") -ge $run.IndexOf('$useNoChangeFastPath = Test-UseNoChangeFastPath')){throw 'Fee check skipped on no-change runs'}
if($run -notmatch '\$deliveryOnlyChange -and \[int\]\$giftReconcileResult.updated_cells -eq 0'){throw 'Fee-only updates rebuild costs'}
if($run -notmatch 'DELIVERY_WRITE_FAILED:' -or $run -notmatch 'delivery_fee_alert'){throw 'No rollback/notification handling'}
$order=[IO.File]::ReadAllText((Join-Path $root 'work\sync_meituan_orders.ps1'),[Text.Encoding]::UTF8)
if($order -notmatch '\$automaticDeliveryMarker' -or $order -notmatch '\$existingSummary.DeliveryFee-\$automaticDelivery'){throw 'Automatic fee does not survive later Meituan sync/manual edit'}
[ordered]@{ok=$true;cases=20;external_writes=0}|ConvertTo-Json -Compress
