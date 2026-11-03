# Todo-only helpers. Financial routines are used only for the shared lock,
# targeted before-images and message-id idempotency; never recalculate accounts.
function Get-TodoCell($Row,[int]$Index) {
    if($null -eq $Row -or $Index -ge @($Row).Count){return ''}
    $cell=@($Row)[$Index];if($null -eq $cell){return ''}
    $p=$cell.PSObject.Properties['value'];if($null -eq $p){return ''}
    return ([string]$p.Value).Trim()
}
function Get-TodoHash([string]$Value) {
    $hash=[Security.Cryptography.SHA256]::Create()
    try{return ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value)))).Replace('-','').ToLowerInvariant()}
    finally{$hash.Dispose()}
}
function ConvertTo-TodoDay([string]$Value) {
    if(-not $Value){return ''}
    $n=0.0
    if([double]::TryParse($Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){
        if($n -lt 30000 -or $n -gt 90000){throw '待办日期无效'}
        return ([datetime]::FromOADate($n)).ToString('yyyy-MM-dd')
    }
    return ([datetime]::Parse($Value,[Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
}
function Resolve-TodoDay([string]$Value,[string]$CreateTime) {
    $stamp=[datetimeoffset]::Now;$n=0L
    if([long]::TryParse($CreateTime,[ref]$n) -and $n -gt 0){
        $stamp=if($n -gt 100000000000){[datetimeoffset]::FromUnixTimeMilliseconds($n)}else{[datetimeoffset]::FromUnixTimeSeconds($n)}
    }elseif($CreateTime){$stamp=[datetimeoffset]::Parse($CreateTime,[Globalization.CultureInfo]::InvariantCulture)}
    $zone=[TimeZoneInfo]::FindSystemTimeZoneById('China Standard Time')
    $day=[TimeZoneInfo]::ConvertTime($stamp,$zone).Date
    switch($Value){
        '' {return $day.ToString('yyyy-MM-dd')}
        'today' {return $day.ToString('yyyy-MM-dd')}
        'tomorrow' {return $day.AddDays(1).ToString('yyyy-MM-dd')}
        'yesterday' {return $day.AddDays(-1).ToString('yyyy-MM-dd')}
        'undated' {return ''}
    }
    return ([datetime]::ParseExact($Value,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
}
function Get-TodoRef([int]$Row,[string]$Day,[string]$Text,[string]$Note) {
    $digest=Get-TodoHash (ConvertTo-Json -InputObject @($Day,$Text,$Note) -Compress)
    return "todo:$Row`:$($digest.Substring(0,20))"
}
function ConvertFrom-TodoRows($Rows,$RowIndices=@()) {
    $headers=@('日期','待办事项','状态','备注')
    if(@($Rows).Count -lt 1){throw '待办事项表为空，请先初始化'}
    for($i=0;$i -lt 4;$i++){if((Get-TodoCell $Rows[0] $i) -ne $headers[$i]){throw "待办事项第$($i+1)列表头应为：$($headers[$i])"}}
    $items=[Collections.Generic.List[object]]::new()
    for($i=1;$i -lt @($Rows).Count;$i++){
        $text=Get-TodoCell $Rows[$i] 1;if(-not $text){continue}
        $dayText=Get-TodoCell $Rows[$i] 0;$invalidDate=$false
        try{$day=ConvertTo-TodoDay $dayText}catch{$day='';$invalidDate=$true}
        $status=Get-TodoCell $Rows[$i] 2;if(-not $status){$status='未完成'}
        $row=if(@($RowIndices).Count -gt $i){[int]$RowIndices[$i]}else{$i+1}
        $note=Get-TodoCell $Rows[$i] 3
        $items.Add([pscustomobject]@{row=$row;date=$day;text=$text;status=$status;note=$note;date_error=$invalidDate;todo_ref=(Get-TodoRef $row $dayText $text $note)})
    }
    return $items.ToArray()
}
function Get-TodoReminderText($Items,[string]$Date) {
    $due=@($Items|Where-Object {$_.status -ne '已完成' -and (-not $_.date -or $_.date -le $Date)}|Sort-Object date,row)
    if(-not $due.Count){return '✅ 今日没有未完成待办。'}
    $lines=[Collections.Generic.List[string]]::new()
    $lines.Add("📋 未完成待办：$($due.Count) 项")
    foreach($item in @($due|Select-Object -First 20)){
        $label=if($item.date_error){'日期待核对'}elseif(-not $item.date){'未指定日期'}elseif($item.date -lt $Date){"遗留 $($item.date)"}else{'今天'}
        if($item.status -notin @('未完成','已完成')){$label+='，状态待核对'}
        $text=([regex]::Replace($item.text,'[\r\n]+',' '));if($text.Length -gt 150){$text=$text.Substring(0,150)+'…'}
        $lines.Add("- [$label] $text")
    }
    if($due.Count -gt 20){$lines.Add("另有 $($due.Count-20) 项，请到待办事项表查看。")}
    return $lines -join "`n"
}
function Get-TodoContext([string]$ProjectRoot) {
    $meta=[IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8)|ConvertFrom-Json
    return [pscustomobject]@{token=([uri]$meta.workbook_url).Segments[-1].Trim('/');cli=(Join-Path $ProjectRoot 'agent\lark_cli_noexit.ps1');url=$meta.workbook_url}
}
function Invoke-TodoLark($Context,[string[]]$Arguments,[string]$Payload=$null) {
    $global:LASTEXITCODE=0
    if($PSBoundParameters.ContainsKey('Payload')){$raw=$Payload|& $Context.cli @Arguments|Out-String}else{$raw=& $Context.cli @Arguments|Out-String}
    if($LASTEXITCODE -ne 0){throw '待办事项飞书操作失败，请检查授权；未确认完成写入'}
    $r=$raw|ConvertFrom-Json;if($r.ok -ne $true){throw '待办事项飞书操作未成功'}
    return $r
}
function Read-TodoSheet($Context) {
    $r=Invoke-TodoLark $Context @('sheets','+cells-get','--as','user','--spreadsheet-token',$Context.token,'--sheet-name','待办事项','--range','A1:D1001','--include','value','--max-chars','1500000','--format','json')
    if($r.data.has_more -or $r.data.ranges[0].truncated){throw '待办清单读取不完整，未按空清单处理'}
    $rows=@($r.data.ranges[0].cells)
    if($rows.Count -ge 1001 -and (Get-TodoCell $rows[1000] 1)){throw '待办清单已达读取上限，请扩展容量后继续'}
    return [pscustomobject]@{rows=$rows;indices=@($r.data.ranges[0].row_indices);items=@(ConvertFrom-TodoRows $rows @($r.data.ranges[0].row_indices))}
}
function Get-PetstarTodos([string]$ProjectRoot) {
    return @((Read-TodoSheet (Get-TodoContext $ProjectRoot)).items)
}
function New-TodoPlan($Snapshot,$Actions,[string]$CreateTime) {
    $items=[Collections.Generic.List[object]]::new()
    foreach($item in @($Snapshot.items)){$items.Add(($item|ConvertTo-Json|ConvertFrom-Json))}
    $maxRow=1
    for($i=1;$i -lt @($Snapshot.rows).Count;$i++){
        $row=if(@($Snapshot.indices).Count -gt $i){[int]$Snapshot.indices[$i]}else{$i+1}
        for($c=0;$c -lt 4;$c++){if(Get-TodoCell $Snapshot.rows[$i] $c){$maxRow=[math]::Max($maxRow,$row);break}}
    }
    $plans=[Collections.Generic.List[object]]::new()
    foreach($action in @($Actions)){
        if([string]$action.action -notin @('AddTodo','CompleteTodo')){throw '待办执行器只允许新增或完成待办'}
        $a=$action.arguments;$text=([string]$a.todo_text).Trim()
        if([string]$action.action -eq 'AddTodo'){
            if(-not $text -or $text.Length -gt 1000){throw '请提供明确的待办事项（不超过1000字）'}
            $day=Resolve-TodoDay ([string]$a.todo_date) $CreateTime
            $note=([string]$a.note).Trim();$row=$maxRow+1
            if($row -gt 1000){throw '待办清单容量不足，请扩展后再新增'}
            $maxRow=$row
            $plans.Add([pscustomobject]@{action='AddTodo';row=$row;date=$day;text=$text;note=$note;changed=$true})
            $dateValue=if($day){([datetime]::ParseExact($day,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)).ToOADate().ToString([Globalization.CultureInfo]::InvariantCulture)}else{''}
            $items.Add([pscustomobject]@{row=$row;date=$day;text=$text;status='未完成';note=$note;date_error=$false;todo_ref=(Get-TodoRef $row $dateValue $text $note)})
        }else{
            $ref=[string]$a.todo_ref
            if($ref){$matches=@($items|Where-Object {$_.todo_ref -eq $ref -and (-not $text -or $_.text -eq $text)})}
            elseif($text){$matches=@($items|Where-Object {$_.text -eq $text})}
            else{throw '请说明完成的是哪一项待办'}
            if($a.todo_date){$day=Resolve-TodoDay ([string]$a.todo_date) $CreateTime;$matches=@($matches|Where-Object {$_.date -eq $day})}
            if($matches.Count -ne 1){throw '待办未能唯一定位或原引用已变化，请说明日期和完整事项；本次没有修改'}
            $item=$matches[0];$plans.Add([pscustomobject]@{action='CompleteTodo';row=$item.row;date=$item.date;text=$item.text;note=$item.note;expected_ref=$item.todo_ref;changed=($item.status -ne '已完成')});$item.status='已完成'
        }
    }
    return $plans.ToArray()
}
function Invoke-TodoActions([string]$ProjectRoot,$Actions,[string]$MessageId,[string]$CreateTime,[switch]$DryRun) {
    if(-not $MessageId){throw '缺少消息编号，不能安全登记待办'}
    if(@($Actions).Count -lt 1 -or @($Actions).Count -gt 20){throw '待办动作数量无效'}
    foreach($action in @($Actions)){if($action.action -notin @('AddTodo','CompleteTodo')){throw '待办与其他操作不能混合执行'}}
    . (Join-Path $ProjectRoot 'agent\finance_transaction.ps1')
    $context=Get-TodoContext $ProjectRoot;$state=Join-Path $ProjectRoot 'agent\state'
    $eventId='todo-'+(Get-TodoHash $MessageId);$tx=$null;$closed=$false
    $mutex=[Threading.Mutex]::new($false,('Local\PetstarTodo_'+$eventId));$taken=$false
    try{
        try{$taken=$mutex.WaitOne([timespan]::FromSeconds(5))}catch [Threading.AbandonedMutexException]{$taken=$true}
        if(-not $taken){throw '待办正在处理中，请稍后重试'}
        if(Test-FinanceCommittedEvent $state $eventId){
            $journal=Read-FinanceTransactionJournal (Get-FinanceTransactionPath $state "event-$eventId")
            return (($journal.processed_events|Where-Object {$_.event_id -eq $eventId}|Select-Object -First 1).reference|ConvertFrom-Json)
        }
        if(-not $DryRun){$tx=New-FinanceTransaction $state '待办事项修改' "event-$eventId" -AcquireLock -WorkbookTokens @{source=$context.token} -LarkCli $context.cli -LockTimeoutSeconds 5}
        $snapshot=Read-TodoSheet $context
        try{$plans=@(New-TodoPlan $snapshot $Actions $CreateTime)}catch{
            if($null -ne $tx){$rollback=Invoke-FinanceTransactionRollback $tx @{source=$context.token} $context.cli $_.Exception.Message;if(-not $rollback.ok){throw '待办事务待恢复，请先处理后再重试'};Close-FinanceTransaction $tx;$closed=$true}
            return [pscustomobject]@{ok=$true;needs_information=$true;reply=$_.Exception.Message;plans=@()}
        }
        if($DryRun){return [pscustomobject]@{ok=$true;dry_run=$true;plans=$plans}}
        $lines=[Collections.Generic.List[string]]::new()
        foreach($p in $plans){
            if($p.changed){
                # Manual sheet edits do not hold our process lock: check the
                # current target immediately before recording its before-image.
                $current=Read-TodoSheet $context
                if($p.action -eq 'AddTodo'){
                    for($i=1;$i -lt @($current.rows).Count;$i++){
                        $row=if(@($current.indices).Count -gt $i){[int]$current.indices[$i]}else{$i+1}
                        if($row -eq $p.row){for($c=0;$c -lt 4;$c++){if(Get-TodoCell $current.rows[$i] $c){throw '待办目标行已被手动填写，请重试；没有覆盖该行'}}}
                    }
                }else{
                    $currentItem=@($current.items|Where-Object {$_.row -eq $p.row -and $_.todo_ref -eq $p.expected_ref})
                    if($currentItem.Count -ne 1){throw '待办内容或行号已变化，请重新说明要完成的事项'}
                }
                $range=if($p.action -eq 'AddTodo'){"A$($p.row):D$($p.row)"}else{"C$($p.row):C$($p.row)"}
                Register-FinanceRangeBeforeImage $tx 'source' $context.token '待办事项' $range $context.cli
                $cells=@()
                if($p.action -eq 'AddTodo'){$date=if($p.date){([datetime]::ParseExact($p.date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)).ToOADate()}else{''};$cells+=,@(@{value=$date},@{value=$p.text},@{value='未完成'},@{value=$p.note})}
                else{$cells+=,@(@{value='已完成'})}
                $null=Invoke-TodoLark $context @('sheets','+cells-set','--as','user','--spreadsheet-token',$context.token,'--sheet-name','待办事项','--range',$range,'--cells','-','--format','json') (ConvertTo-Json -InputObject $cells -Depth 6 -Compress)
                $back=Read-TodoSheet $context;$item=@($back.items|Where-Object {$_.row -eq $p.row})
                $status=if($p.action -eq 'AddTodo'){'未完成'}else{'已完成'}
                if($item.Count -ne 1 -or $item[0].text -ne $p.text -or $item[0].date -ne $p.date -or $item[0].status -ne $status -or $item[0].note -ne $p.note){throw '待办事项回读校验失败'}
            }
            $verb=if($p.action -eq 'AddTodo'){'已新增待办'}elseif(-not $p.changed){'该待办已完成，无需重复修改'}else{'已完成'}
            $lines.Add("$verb`：$($p.text)$(if($p.date){'（'+$p.date+'）'})")
        }
        $v=Invoke-TodoLark $context @('sheets','+formula-verify','--as','user','--spreadsheet-token',$context.token,'--sheet-name','待办事项','--range','A1:D1001','--exit-on-error','--format','json')
        if($v.data.status -ne 'success'){throw '待办事项单元格检查未通过'}
        $receipt=[pscustomobject]@{ok=$true;reply=($lines -join "`n");plans=$plans;needs_information=$false}
        Register-FinanceProcessedEvent $tx $eventId 'TodoActions' (ConvertTo-Json -InputObject $receipt -Depth 10 -Compress)
        Complete-FinanceTransaction $tx;Close-FinanceTransaction $tx;$closed=$true
        return $receipt
    }catch{
        if($null -ne $tx -and -not $closed){$rollback=Invoke-FinanceTransactionRollback $tx @{source=$context.token} $context.cli $_.Exception.Message;if(-not $rollback.ok){throw '待办回滚待恢复，请先处理事务后再重试'}}
        throw
    }finally{
        if($null -ne $tx -and -not $closed){Close-FinanceTransaction $tx}
        if($taken){try{$mutex.ReleaseMutex()}catch{}};$mutex.Dispose()
    }
}
