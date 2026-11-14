$ErrorActionPreference='Stop'
Set-StrictMode -Version 2
$root=Split-Path $PSScriptRoot -Parent
$script:cases=0
function Check([bool]$Pass,[string]$Label) {
    if(-not $Pass){throw $Label};$script:cases++
}
function Parse([string]$Relative) {
    $errors=$null;$tokens=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root $Relative),[ref]$tokens,[ref]$errors)
    if(@($errors).Count){throw ($errors.Message -join ';')}
    return $ast
}
function Import-Function($Ast,[string]$Name) {
    $f=$Ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true)
    if(-not $f){throw "Missing $Name"}
    $body=$f.Body.Extent.Text.Substring(1,$f.Body.Extent.Text.Length-2)
    if(@($f.Parameters).Count){$body='param('+(@($f.Parameters|ForEach-Object{$_.Extent.Text})-join ',')+")`n"+$body}
    Set-Item "function:script:$Name" ([scriptblock]::Create($body))
}
$run=Parse 'agent/run_sync_and_notify.ps1'
$alertLoop=$run.Find({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Variable.VariablePath.UserPath -eq 'decision'},$true)
Check ($null -ne $alertLoop) 'Missing real delivery warning path'
$syncWarnings=[Collections.Generic.List[string]]::new()
$deliveryFeeResult=[pscustomobject]@{decisions=@(
    [pscustomobject]@{state='default_self';order_no='1234567890123456789';reason='no fee'},
    [pscustomobject]@{state='deferred';order_no='1234567890123456789';reason='provider unavailable'},
    [pscustomobject]@{state='pending';order_no='1234567890123456789';reason='legacy pending'}
)}
Invoke-Expression $alertLoop.Extent.Text
Check ($syncWarnings.Count -eq 0) 'No-fee/read failures still create per-order notifications'
$deliveryFeeResult.decisions=@([pscustomobject]@{state='conflict';order_no='1234567890123456789';reason='different amounts'})
Invoke-Expression $alertLoop.Extent.Text
Check ($syncWarnings.Count -eq 1 -and $syncWarnings[0] -match '金额冲突') 'Real amount conflict was suppressed'
Import-Function $run 'Get-DeliveryFeeFailureWarning'
Check ((Get-DeliveryFeeFailureWarning 'Could not connect to Chrome DevTools') -notmatch '^配送费待核对：') 'Technical failure still triggers per-order alert channel'
$snapshot=$run.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Extent.Text -match '\$pendingDelivery ='},$true)
Check ($snapshot.Extent.Text -notmatch 'needsConfirmation') 'Snapshot still flags zero delivery charges'

# Exercise the actual write/readback loop entirely against an in-memory sheet.
$child=Parse 'agent/sync_delivery_fees.ps1'
Import-Function $child 'Get-DeliveryCell'
$writeLoop=$child.Find({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Variable.VariablePath.UserPath -eq 'd'},$true)
Check ($null -ne $writeLoop) 'Missing actual delivery writer'
$script:grid=[object[]]::new(17)
for($i=0;$i -lt 17;$i++){$script:grid[$i]=[pscustomobject]@{value=$null}}
$id='1234567890123456789';$script:grid[2].value=$id
$script:writes=[Collections.Generic.List[string]]::new()
function Set-DeliveryCell([string]$Range,$Value) {
    $script:writes.Add($Range)
    if($Range -eq 'J2:J2'){$script:grid[9].value=$Value}
    elseif($Range -eq 'P2:P2'){$script:grid[15].value=$Value}
    else{throw 'Unexpected write target'}
}
function Get-DeliveryRows($Token,$Sheet,$Range) {
    if($Sheet -ne '订单汇总' -or $Range -ne 'A2:Q2'){throw 'Unexpected readback target'}
    $matrix=@();$matrix+=,$script:grid;return ,$matrix
}
$Apply=$true;$token='offline';$writesStarted=$false
$byId=@{};$byId[$id]=[pscustomobject]@{order_no=$id;date='2026-11-14';row=2;note='original note';intervention=$false;protected=$false;confirmed_self=$false}
$changes=[Collections.Generic.List[object]]::new()
$result=[pscustomobject]@{decisions=@([pscustomobject]@{order_no=$id;state='default_self';amount=0;changed=$true})}
Invoke-Expression $writeLoop.Extent.Text
Check ($script:grid[9].value -eq 0 -and $script:grid[15].value -match '默认自配送') 'Default zero not written with provisional source'
Check ($script:grid[15].value -notmatch '人工确认' -and $changes[0].provider -eq 'default_self') 'Default zero became a manual lock'
$byId[$id].note=$script:grid[15].value
$result.decisions=@([pscustomobject]@{order_no=$id;state='confirmed';amount=4.51;changed=$true;provider='cyanclouds';delivery_id='123'})
Invoke-Expression $writeLoop.Extent.Text
Check ($script:grid[9].value -eq 4.51 -and $script:grid[15].value -match '自动配送费\[cyanclouds:123\]=4.51元' -and $script:grid[15].value -notmatch '默认自配送') 'Later real fee did not replace provisional zero/source'
Check (($script:writes -join ',') -eq 'J2:J2,P2:P2,J2:J2,P2:P2') 'Writer touched non-delivery business fields'
$byId[$id].protected=$true
$blocked=$false;try{Invoke-Expression $writeLoop.Extent.Text}catch{$blocked=$_.Exception.Message -match '保护触发'}
Check ($blocked -and $script:writes.Count -eq 4) 'Manual amount protection bypassed'

# A newly registered intervention defaults only its delivery fields, never
# inventing a known platform fee or changing a natural order on its own.
$finance=Parse 'agent/record_finance_event.ps1'
foreach($name in @('Get-CellValue','Convert-ToDateValue','Convert-ToNumber')){Import-Function $finance $name}
$dispatch=$finance.FindAll({param($n)$n -is [Management.Automation.Language.SwitchStatementAst] -and $n.Condition.Extent.Text -eq '$Action'},$true)|Select-Object -Last 1
$case=@($dispatch.Clauses|Where-Object {$_.Item1.Extent.Text -eq "'AddIntervention'"})[0].Item2.Extent.Text
$summary=[object[]]::new(17);for($i=0;$i -lt 17;$i++){$summary[$i]=[pscustomobject]@{value=$null}}
$summary[0].value='2026-11-13';$summary[2].value=$id;$summary[16].value='美团4号单'
function Get-SheetRows($Sheet,$Range){if($Sheet -eq '干预订单'){return ,@()};if($Sheet -ne '订单汇总'){throw 'Unexpected sheet read'};$m=@();$m+=,$summary;return ,$m}
function Get-FirstEmptyRow {return 2}
function New-ValueCell($Value){return @{value=$Value}}
function Convert-ToExcelSerial($Value){return 46339}
function Set-CellMatrix($Sheet,$Range,$Matrix){if($Sheet -ne '干预订单' -or $Range -ne 'A2:M2'){throw 'Unexpected intervention write'};$script:newRecord=$Matrix[0]}
$DryRun=$false;$OrderNo=$id;$Date=[datetime]'2026-11-14';$Person='fixture';$Note='fixture'
Invoke-Expression ($case.Substring(1,$case.Length-2))
Check ($script:newRecord[5].value -eq '自配送' -and $script:newRecord[7].value -eq 0) 'New intervention still has unknown delivery'
Check ($null -eq $script:newRecord[6].value) 'Delivery default silently zeroed unknown platform deduction'
Check ($script:newRecord[3].value -eq $id -and $script:newRecord[1].value -eq '美团4号单') 'Order identity changed'
$summary[9].value=7.55
Invoke-Expression ($case.Substring(1,$case.Length-2))
Check ($script:newRecord[5].value -eq '骑手配送' -and $script:newRecord[7].value -eq 7.55) 'Registering intervention discarded an existing positive delivery charge'

# Use the real ledger expense builder without loading any cloud/CLI entry code.
$ledger=[IO.File]::ReadAllText((Join-Path $root 'agent/sync_daily_ledger.ps1'),[Text.Encoding]::UTF8)
$start=$ledger.IndexOf('$deliveryTotal = if ($null -eq $deliveryFee)')
$end=$ledger.IndexOf('$expenseEntries +=', $start)
Check ($start -ge 0 -and $end -gt $start) 'Missing actual delivery expense builder'
$expenseBody=$ledger.Substring($start,$end-$start)
$orderNo=$id
foreach($fixture in @(
    @{fee=$null;kind='';old=$null;expected=0;source='默认自配送'},
    @{fee=0;kind='骑手配送';old=$null;expected=0;source='默认自配送'},
    @{fee=$null;kind='';old=[pscustomobject]@{fee=6.78;source='人工补录'};expected=6.78;source='人工补录'},
    @{fee=4.51;kind='骑手配送';old=[pscustomobject]@{fee=6.78;source='人工补录'};expected=4.51;source='人工补录'}
)) {
    $deliveryFee=$fixture.fee;$deliveryType=$fixture.kind;$oldManagedExpenseByKey=@{}
    if($null -ne $fixture.old){$oldManagedExpenseByKey["$orderNo|补单配送费"]=$fixture.old}
    Invoke-Expression $expenseBody
    Check ($deliveryTotal -eq $fixture.expected -and $deliverySource -eq $fixture.source) 'Ledger default zero or positive amount preservation failed'
}
[ordered]@{ok=$true;cases=$script:cases;external_writes=0;credential_access=0}|ConvertTo-Json -Compress
