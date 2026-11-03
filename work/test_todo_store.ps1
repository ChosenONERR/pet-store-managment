# Entirely offline. Network and journal adapters below are in-memory fakes;
# no Feishu message, task, finance data or physical journal is created.
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2
$projectRoot=Split-Path $PSScriptRoot -Parent
. (Join-Path $projectRoot 'agent\todo_store.ps1')
$script:checks=0
function Assert($Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
function Throws([scriptblock]$Body,[string]$Message){$threw=$false;try{& $Body|Out-Null}catch{$threw=$true};Assert $threw $Message}
function Row($Values){return ,@($Values|ForEach-Object {[pscustomobject]@{value=$_}})}
function Action($Type,$Text,$Day=$null,$Ref=$null){return [pscustomobject]@{action=$Type;arguments=[pscustomobject]@{todo_text=$Text;todo_date=$Day;todo_ref=$Ref;note=$null}}}
function Snapshot($Rows){return [pscustomobject]@{rows=$Rows;indices=@();items=@(ConvertFrom-TodoRows $Rows)}}
$rows=@();$rows+=,(Row @('日期','待办事项','状态','备注'))
$rows+=,(Row @('2026-11-03','核对进货单','',''))
$rows+=,(Row @('2026-11-02','检查货架','未完成','逾期'))
$rows+=,(Row @('','打扫仓库','',''))
$rows+=,(Row @('2026-11-04','明日采购','未完成',''))
$rows+=,(Row @('2026-11-03','已经完成','已完成',''))
$rows+=,(Row @('日期写错','日期异常','未知状态',''))
$snapshot=Snapshot $rows
Assert ($snapshot.items.Count -eq 6) 'All named tasks are read'
Assert ($snapshot.items[0].status -eq '未完成') 'Blank status defaults pending'
$reminder=Get-TodoReminderText $snapshot.items '2026-11-03'
Assert ($reminder -match '未完成待办：4 项') 'Today, overdue, undated and invalid date remain due'
Assert ($reminder -notmatch '明日采购|已经完成') 'Future and completed tasks omitted'
Assert ($reminder -match '日期待核对，状态待核对') 'Malformed manual entry cannot silently disappear'
Assert ((ConvertTo-TodoDay '46329') -eq ([datetime]::FromOADate(46329)).ToString('yyyy-MM-dd')) 'Excel dates supported'
Assert ((Resolve-TodoDay 'today' '2026-11-02T16:01:00Z') -eq '2026-11-03') 'Original message date in China'
Assert ((Resolve-TodoDay 'tomorrow' '2026-11-02T16:01:00Z') -eq '2026-11-04') 'Relative dates anchored before delayed replay'
$epoch=[datetimeoffset]::Parse('2026-11-02T16:01:00Z').ToUnixTimeMilliseconds().ToString()
Assert ((Resolve-TodoDay 'today' $epoch) -eq '2026-11-03') 'Feishu millisecond timestamps supported'
Assert ((Resolve-TodoDay 'undated' $epoch) -eq '') 'Explicit undated supported'
Throws {Resolve-TodoDay 'not a date' $epoch} 'Bad date rejected'
$plan=@(New-TodoPlan $snapshot @((Action 'AddTodo' '盘点赠品' 'tomorrow')) $epoch)
Assert ($plan[0].row -eq 8 -and $plan[0].date -eq '2026-11-04') 'Append preserves all prior rows'
$draftRows=@($rows);$draftRows+=,(Row @('2026-11-03','','','手工草稿'))
$plan=@(New-TodoPlan (Snapshot $draftRows) @((Action 'AddTodo' '新增')) $epoch)
Assert ($plan[0].row -eq 9) 'Do not overwrite manual incomplete drafts'
$dupes=@($rows);$dupes+=,(Row @('2026-11-01','核对进货单','未完成',''))
Throws {New-TodoPlan (Snapshot $dupes) @((Action 'CompleteTodo' '核对进货单')) $epoch} 'Ambiguous matching must not write'
$plan=@(New-TodoPlan (Snapshot $dupes) @((Action 'CompleteTodo' '核对进货单' '2026-11-01')) $epoch)
Assert ($plan[0].row -eq 8) 'Explicit day disambiguates'
$ref=$snapshot.items[0].todo_ref
$plan=@(New-TodoPlan $snapshot @((Action 'CompleteTodo' '核对进货单' $null $ref)) $epoch)
Assert ($plan[0].row -eq 2 -and $plan[0].changed) 'Reference binds correct row and text'
Throws {New-TodoPlan $snapshot @((Action 'CompleteTodo' '检查货架' $null $ref)) $epoch} 'Reference/text mismatch rejected'
$plan=@(New-TodoPlan $snapshot @((Action 'CompleteTodo' '已经完成')) $epoch)
Assert (-not $plan[0].changed) 'Already completed is a no-op'
Throws {New-TodoPlan $snapshot @((Action 'SetSkuCost' '非法')) $epoch} 'No finance operations allowed'
Throws {New-TodoPlan $snapshot @((Action 'AddTodo' '有效'),(Action 'CompleteTodo' '不存在')) $epoch} 'Whole batch is planned before any write'

# Execute production orchestration with only dependency imports replaced.
$body=(Get-Command Invoke-TodoActions).Definition.Replace(". (Join-Path `$ProjectRoot 'agent\finance_transaction.ps1')",'# In-memory transaction adapters')
Set-Item -Path function:Invoke-TodoActions -Value ([scriptblock]::Create($body))
$script:liveRows=@($rows);$script:committed=@{};$script:writeCount=0;$script:rollbackCount=0;$script:failVerify=$false
function Get-TodoContext($ProjectRoot){return [pscustomobject]@{token='offline';cli='offline';url='offline'}}
function Read-TodoSheet($Context){return Snapshot $script:liveRows}
function Test-FinanceCommittedEvent($Root,$Id){return $script:committed.ContainsKey($Id)}
function Get-FinanceTransactionPath($Root,$Id){return $Id}
function Read-FinanceTransactionJournal($Path){return $script:committed[$Path.Substring(6)]}
function New-FinanceTransaction($Root,$Description,$Id,[switch]$AcquireLock,$WorkbookTokens,$LarkCli,$LockTimeoutSeconds){$script:backup=ConvertTo-Json -InputObject $script:liveRows -Depth 10 -Compress;return [pscustomobject]@{id=$Id;processed_events=@()}}
function Register-FinanceRangeBeforeImage($Tx,$Role,$Token,$Sheet,$Range,$Cli){Assert ($Sheet -eq '待办事项' -and $Range -match '^[AC]\d+:[CD]\d+$') 'Only task row/status snapshots'}
function Register-FinanceProcessedEvent($Tx,$Id,$Action,$Reference){$Tx.processed_events=@([pscustomobject]@{event_id=$Id;reference=$Reference})}
function Complete-FinanceTransaction($Tx){$script:committed[$Tx.id.Substring(6)]=$Tx}
function Close-FinanceTransaction($Tx){}
function Invoke-FinanceTransactionRollback($Tx,$Tokens,$Cli,$Reason){$script:rollbackCount++;$script:liveRows=@((ConvertFrom-Json -InputObject ('{"rows":'+$script:backup+'}')).rows);return [pscustomobject]@{ok=$true}}
function Invoke-TodoLark($Context,$Arguments,$Payload){
    if($Arguments -contains '+formula-verify'){if($script:failVerify){throw 'simulated readback failure'};return [pscustomobject]@{data=[pscustomobject]@{status='success'}}}
    Assert ($Arguments -contains '+cells-set') 'Only cell writes and verification'
    $range=$Arguments[[array]::IndexOf($Arguments,'--range')+1]
    $rowNumber=[int]([regex]::Match($range,'\d+').Value)
    $cells=(ConvertFrom-Json -InputObject ('{"rows":'+$Payload+'}')).rows
    if($range.StartsWith('A')){while($script:liveRows.Count -lt $rowNumber){$script:liveRows+=,(Row @('','','',''))};$script:liveRows[$rowNumber-1]=@($cells[0])}
    else{$script:liveRows[$rowNumber-1][2]=[pscustomobject]@{value='已完成'}}
    $script:writeCount++;return [pscustomobject]@{ok=$true}
}
$result=Invoke-TodoActions $projectRoot @((Action 'AddTodo' '测试新增')) 'offline-message-a' $epoch
Assert ($result.ok -and $script:writeCount -eq 1) 'Add writes exactly one row'
$result=Invoke-TodoActions $projectRoot @((Action 'AddTodo' '测试新增')) 'offline-message-a' $epoch
Assert ($result.ok -and $script:writeCount -eq 1) 'Duplicate message receipt never duplicates task'
$result=Invoke-TodoActions $projectRoot @((Action 'CompleteTodo' '测试新增')) 'offline-message-b' $epoch
Assert ($result.ok -and (Snapshot $script:liveRows).items[-1].status -eq '已完成') 'Complete writes only status'
$before=ConvertTo-Json -InputObject $script:liveRows -Depth 10 -Compress
$script:failVerify=$true
Throws {Invoke-TodoActions $projectRoot @((Action 'AddTodo' '不能留下半成品')) 'offline-message-c' $epoch} 'Verification failure must fail action'
Assert ($script:rollbackCount -eq 1 -and (ConvertTo-Json -InputObject $script:liveRows -Depth 10 -Compress) -eq $before) 'Failed write rolled back completely'
$script:failVerify=$false
$result=Invoke-TodoActions $projectRoot @((Action 'AddTodo' '不能留下半成品')) 'offline-message-c' $epoch
Assert ($result.ok -and $result.plans.Count -eq 1) 'Rolled-back message can safely retry'
Throws {Invoke-TodoActions $projectRoot @((Action 'AddTodo' '缺少消息编号')) '' $epoch} 'Missing id cannot collapse independent messages'
Write-Output "PASS: $script:checks todo checks; no remote writes or messages"
