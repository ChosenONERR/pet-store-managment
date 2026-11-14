$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$finance=Join-Path $root 'agent/record_finance_event.ps1'
$bridge=Join-Path $root 'agent/lark_event_bridge.ps1'
function Import-TestFunction([string]$Path,[string]$Name) {
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$null,[ref]$null)
    $f=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true)
    if(-not $f){throw "Missing test function $Name"}
    $body=$f.Body.Extent.Text.Substring(1,$f.Body.Extent.Text.Length-2)
    if(@($f.Parameters).Count){$body='param('+(@($f.Parameters|ForEach-Object{$_.Extent.Text})-join ',')+")`n"+$body}
    Set-Item -Path "function:script:$Name" -Value ([scriptblock]::Create($body))
}
foreach($n in @('Get-CellValue','Convert-ToNumber','Convert-ToDateValue','Assert-NonNegative')){Import-TestFunction $finance $n}
foreach($n in @('Convert-PetstarArgumentsToMap','Convert-CommandOutputToJson','Get-CommandFailureDetail','Test-PreviewNeedsClarification','Invoke-PetstarEvent','Format-ConfirmationPreview','Format-CodexExecutionReply')){Import-TestFunction $bridge $n}
$ast=[Management.Automation.Language.Parser]::ParseFile($finance,[ref]$null,[ref]$null)
$dispatch=$ast.FindAll({param($n)$n -is [Management.Automation.Language.SwitchStatementAst] -and $n.Condition.Extent.Text -eq '$Action'},$true)|Select-Object -Last 1
$case=@($dispatch.Clauses|Where-Object {$_.Item1.Extent.Text -match "'ConfirmSelfDelivery','SetOrderDeliveryFee'"})
if($case.Count -ne 1){throw 'Self-delivery does not share the general order-fee branch'}
$text=$case[0].Item2.Extent.Text
$body=[scriptblock]::Create($text.Substring(1,$text.Length-2).Replace('exit 0','return')+"`nif(-not `$DryRun){`$writeResult}")
$script:count=0
function Check([bool]$Pass,[string]$Label){if(-not $Pass){throw $Label};$script:count++}
function Row([int]$Length){$r=[object[]]::new($Length);for($i=0;$i -lt $Length;$i++){$r[$i]=@{}};return ,$r}
function Cell($Value){return @{value=$Value}}
$id='3402378304070746079'
$natural=Row 17
$natural[0]=Cell '2026-11-07';$natural[2]=Cell $id;$natural[3]=Cell '已完成'
$natural[9]=Cell 4.51;$natural[10]=Cell 14.9;$natural[15]=Cell '原备注';$natural[16]=Cell '美团2号单'
$intervention=Row 13
$intervention[0]=Cell '2026-11-07';$intervention[1]=Cell '美团2号单';$intervention[3]=Cell $id
$intervention[5]=Cell '骑手配送';$intervention[6]=Cell 6.42;$intervention[7]=Cell 5;$intervention[12]=Cell '原补单备注'
$script:orderRows=,$natural;$script:interventionRows=@()
$script:writes=[Collections.Generic.List[object]]::new()
function Get-SheetRows([string]$Sheet,[string]$Range){
    if($Sheet -eq '订单汇总'){return ,$script:orderRows}
    if($Sheet -eq '干预订单'){return ,$script:interventionRows}
    throw 'Unexpected financial read'
}
function Set-OneRow([string]$Sheet,[string]$Range,$Values){$script:writes.Add([pscustomobject]@{sheet=$Sheet;range=$Range;values=$Values})}
function Complete-CurrentFinanceTransaction {}
function Scenario([string]$Action,[bool]$DryRun,$Amount=9.99){
    $OrderNo=$id;$Note='店主确认自配送，没有骑手费';$script:writes.Clear()
    $result=& $body
    if($DryRun){return (($result|Out-String)|ConvertFrom-Json).result}
    return $result
}
$before=$natural|ConvertTo-Json -Depth 5 -Compress
$r=Scenario ConfirmSelfDelivery $true
Check ($r.order_type -eq '自然单' -and $r.delivery_fee -eq 0 -and $r.date -eq '2026-11-07') 'Legacy natural preview has wrong type/fee/date'
Check ($script:writes.Count -eq 0) 'Preview mutated a sheet'
$r=Scenario ConfirmSelfDelivery $false
Check ($r.order_type -eq '自然单' -and $r.delivery_fee -eq 0 -and $r.old_delivery_fee -eq 4.51) 'Natural self delivery was not zeroed'
Check (($script:writes.range -join ',') -eq 'J2:J2,P2:P2' -and @($script:writes|Where-Object {$_.sheet -ne '订单汇总'}).Count -eq 0) 'Natural action changed unrelated tables/costs'
$marker=[string]$script:writes[1].values[0]
Check ($marker -match '人工确认配送费：0.00元' -and $marker -match '原备注') 'Manual zero fee is not protected or note lost'
Check (($natural|ConvertTo-Json -Depth 5 -Compress) -eq $before) 'Planning mutated the source row'
$natural[15]=Cell $marker
$null=Scenario ConfirmSelfDelivery $false
Check ([regex]::Matches([string]$script:writes[1].values[0],'人工确认配送费').Count -eq 1) 'Manual marker duplicated on retry'
$r=Scenario SetOrderDeliveryFee $false 6.78
Check ($r.delivery_fee -eq 6.78 -and $r.delivery_type -eq '骑手配送') 'Existing positive fee action regressed'
$script:interventionRows=,$intervention
$r=Scenario ConfirmSelfDelivery $true
Check ($r.order_type -eq '干预单' -and $r.total_cost -eq 6.42 -and $script:writes.Count -eq 0) 'Intervention preview regressed'
$r=Scenario ConfirmSelfDelivery $false
Check ($r.total_cost -eq 6.42 -and ($script:writes.range -join ',') -eq 'F2:F2,H2:H2,M2:M2') 'Intervention platform cost was lost'
$script:interventionRows=@();$script:orderRows=@()
$rejected=$false;try{$null=Scenario ConfirmSelfDelivery $true}catch{$rejected=$_.Exception.Message -match '未找到订单'}
Check ($rejected -and $script:writes.Count -eq 0) 'Unknown order accepted in preview'
$script:orderRows=,$natural;$natural[3]=Cell '已取消'
$rejected=$false;try{$null=Scenario ConfirmSelfDelivery $true}catch{$rejected=$_.Exception.Message -match '已取消'}
Check ($rejected -and $script:writes.Count -eq 0) 'Cancelled order accepted in preview'
$pythonExe='Invoke-FakePython';$petstarCli='offline-cli';$script:fakeCode=0
function Invoke-FakePython {$script:fakeArgs=@($args);$global:LASTEXITCODE=$script:fakeCode;$script:fakeJson}
$details=[pscustomobject]@{order_no=$id;order_type='自然单';daily_order_no='美团2号单';old_delivery_fee=4.51;delivery_fee=0}
$script:fakeJson=([ordered]@{ok=$true;result=$details}|ConvertTo-Json -Depth 5)
$script:previewMode=$true;$script:previewOperations=[Collections.Generic.List[object]]::new()
$r=Invoke-PetstarEvent @('--action','ConfirmSelfDelivery','--order-no',$id)
Check ($r.ok -and '--dry-run' -in $script:fakeArgs -and $script:previewOperations.Count -eq 1) 'Bridge still synthesizes self-delivery preview'
$op=$script:previewOperations[0]
$preview=Format-ConfirmationPreview 'fixture' @($op) ''
Check ($preview -match '自然单' -and $preview -match '订单汇总及月份收入' -and $preview -notmatch '干预订单') 'Natural preview claims an intervention mutation'
$reply=Format-CodexExecutionReply @($op)
Check ($reply -match '自然单' -and $reply -notmatch '干预订单') 'Natural completion reply has wrong type'
$script:fakeCode=1;$script:fakeJson='{"ok":false,"error":"missing-order-fixture"}'
$script:previewOperations.Clear();$rejected=$false
try{$null=Invoke-PetstarEvent @('--action','ConfirmSelfDelivery','--order-no',$id)}catch{$rejected=$_.Exception.Message -match 'missing-order-fixture'}
Check ($rejected -and $script:previewOperations.Count -eq 0) 'Failed lookup became an executable pending operation'
[ordered]@{ok=$true;cases=$script:count;external_writes=0;credential_access=0}|ConvertTo-Json -Compress
