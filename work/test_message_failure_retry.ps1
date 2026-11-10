$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'agent/lark_event_bridge.ps1'),[ref]$null,[ref]$errors)
if($errors){throw 'Bridge parse failure'}
foreach($name in @('Test-PreviewNeedsClarification','Process-EventWithLedger','Get-CodexActionValue','Invoke-CodexAction')){
    $node=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name-eq$name},$true)
    $body=$node.Body.Extent.Text.Substring(1,$node.Body.Extent.Text.Length-2)
    $params=@($node.Parameters|ForEach-Object{$_.Extent.Text})-join','
    Set-Item "function:script:$name" ([scriptblock]::Create("param($params)`n"+$body))
}
$script:cases=0;$script:handled=@{};$script:acks=0;$script:events=0;$script:writes=0;$script:ackFails=$false
function Check([bool]$Ok,[string]$Why){if(-not$Ok){throw $Why};$script:cases++}
function Get-MessageHash([string]$MessageId){return ('a'*32)+$MessageId}
function Test-MessageHandled($Hash){return $script:handled.ContainsKey($Hash)}
function Add-HandledMessageRecord($Id,$Status,$Source,$CreateTime){if($Status-in@('completed','failed_replied','needs_clarification')){$script:handled[(Get-MessageHash $Id)]=$Status}}
function Add-ProcessingReaction {return $null}
function Remove-ProcessingReaction {}
function Write-BridgeLog {}
function Send-Acknowledgement {if($script:ackFails){throw 'simulated acknowledgement failure'};$script:acks++;return @{queued=$true}}
function Invoke-PetstarEvent {$script:writes++;throw 'Unexpected executor call'}
function Process-Event {
    $script:events++;$script:lastProcessFailure=$null;$script:lastProcessRetryable=$false
    switch($script:testMode){
        'missing' {Invoke-CodexAction ([pscustomobject]@{action='SetOrderActuals';arguments=[pscustomobject]@{amount=0;items=$null}}) ([pscustomobject]@{}) 0 'preview'}
        'validation' {$script:lastProcessNeedsClarification=Test-PreviewNeedsClarification '必须提供有效货号';throw '必须提供有效货号'}
        'network' {throw 'network timeout'}
        'rate' {$script:lastProcessFailure='429';$script:lastProcessRetryable=$true;$script:lastProcessRetryAfterSeconds=60;return $false}
        'nonretryable' {$script:lastProcessFailure='permanent model error';return $false}
        default {return $true}
    }
}
function RunCase([string]$Id){$e=@{message_id=$Id;sender_type='user';sender_id='user';message_type='text';chat_type='p2p';content='没有收入';create_time='1'};return Process-EventWithLedger ($e|ConvertTo-Json -Compress) 'user' 'test'}
foreach($reason in @('修正订单实际履约必须提供实际售出商品','没有识别到赠品','请明确具体商品','SKU成本主表安全定位已失效','未知成本价')){Check (Test-PreviewNeedsClarification $reason) 'Business validation not classified'}
foreach($reason in @('network timeout 必须提供','429 Too Many Requests','读取飞书失败：请补充','授权失败','另一个经营数据事务仍在执行','未知执行错误')){Check (-not(Test-PreviewNeedsClarification $reason)) 'Infrastructure failure swallowed'}
$script:testMode='missing';Check ((RunCase 'missing')-eq'needs_clarification') 'Missing items was retried'
Check ($script:writes-eq0) 'Missing income changed order'
$count=$script:events;$acks=$script:acks
Check ((RunCase 'missing')-eq'duplicate' -and $script:events-eq$count -and $script:acks-eq$acks) 'Old message was processed or notified again'
$script:testMode='validation';Check ((RunCase 'validation')-eq'needs_clarification') 'Validation not stopped'
$script:testMode='network';Check ((RunCase 'network')-eq'failed_retryable') 'Network retry lost'
$count=$script:events;$null=RunCase 'network';Check ($script:events-eq($count+1)) 'Network retry blocked'
$script:testMode='rate';Check ((RunCase 'rate')-eq'failed_retryable') 'Rate-limit retry lost'
$script:testMode='nonretryable';Check ((RunCase 'permanent')-eq'failed_replied') 'Permanent model error retried'
$script:testMode='success';Check ((RunCase 'new-followup')-eq'completed') 'New user message blocked'
$script:testMode='missing';$script:ackFails=$true;Check ((RunCase 'ack-failure')-eq'failed_unreplied') 'Failed clarification delivery swallowed'
$script:ackFails=$false;Check ((RunCase 'ack-failure')-eq'needs_clarification') 'Clarification delivery could not recover'
Check ($script:writes-eq0) 'Test attempted external writes'
[ordered]@{ok=$true;cases=$script:cases;external_writes=0}|ConvertTo-Json -Compress
