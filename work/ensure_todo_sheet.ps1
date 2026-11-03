param([string]$ProjectRoot,[switch]$Apply)
$ErrorActionPreference='Stop'
$OutputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
if(-not $ProjectRoot){$ProjectRoot=Split-Path $PSScriptRoot -Parent}
. (Join-Path $ProjectRoot 'agent\todo_store.ps1')
. (Join-Path $ProjectRoot 'agent\finance_transaction.ps1')
$context=Get-TodoContext $ProjectRoot
$baseArgs=@('--as','user','--spreadsheet-token',$context.token,'--sheet-name','待办事项')
$tx=$null;$closed=$false
try{
    if($Apply){$tx=New-FinanceTransaction (Join-Path $ProjectRoot 'agent\state') '初始化待办事项' $null -AcquireLock -WorkbookTokens @{source=$context.token} -LarkCli $context.cli}
    $book=Invoke-TodoLark $context @('sheets','+workbook-info','--as','user','--spreadsheet-token',$context.token,'--format','json')
    $found=@($book.data.sheets|Where-Object {$_.sheet_name -eq '待办事项'})
    if($found.Count -gt 1){throw '待办事项表不唯一，停止初始化'}
    $hasContent=$false
    if($found.Count){
        $r=Invoke-TodoLark $context (@('sheets','+cells-get')+$baseArgs+@('--range','A1:D1001','--include','value','--max-chars','1500000','--format','json'))
        if($r.data.has_more -or $r.data.ranges[0].truncated){throw '读取不完整，停止初始化'}
        foreach($row in @($r.data.ranges[0].cells)){for($c=0;$c -lt 4;$c++){if(Get-TodoCell $row $c){$hasContent=$true}}}
        if($hasContent){$null=ConvertFrom-TodoRows @($r.data.ranges[0].cells) @($r.data.ranges[0].row_indices)}
    }
    if(-not $Apply){[pscustomobject]@{ok=$true;dry_run=$true;exists=($found.Count -eq 1);preserve_existing=$hasContent;headers=@('日期','待办事项','状态','备注');workbook_url=$context.url}|ConvertTo-Json -Depth 4;return}
    if(-not $found.Count){$null=Invoke-TodoLark $context @('sheets','+sheet-create','--as','user','--spreadsheet-token',$context.token,'--title','待办事项','--row-count','1001','--col-count','4','--format','json')}
    if(-not $hasContent){
        Register-FinanceRangeBeforeImage $tx 'source' $context.token '待办事项' 'A1:D1' $context.cli
        $cells=@();$cells+=,@(@{value='日期'},@{value='待办事项'},@{value='状态'},@{value='备注'})
        $null=Invoke-TodoLark $context (@('sheets','+cells-set')+$baseArgs+@('--range','A1:D1','--cells','-','--allow-overwrite=false','--format','json')) (ConvertTo-Json -InputObject $cells -Depth 5 -Compress)
    }
    # Existing task values are preserved; these idempotent formats also let a
    # partially completed initial installation be safely resumed.
    $null=Invoke-TodoLark $context (@('sheets','+cells-set-style')+$baseArgs+@('--range','A1:D1001','--font-family','Microsoft YaHei','--font-size','11','--vertical-alignment','middle','--word-wrap','auto-wrap','--format','json'))
    $null=Invoke-TodoLark $context (@('sheets','+cells-set-style')+$baseArgs+@('--range','A1:D1','--background-color','#173A2D','--font-color','#FFFFFF','--font-weight','bold','--format','json'))
    $null=Invoke-TodoLark $context (@('sheets','+cells-set-style')+$baseArgs+@('--range','A2:A1001','--number-format','yyyy-mm-dd','--format','json'))
    $null=Invoke-TodoLark $context (@('sheets','+cells-set-style')+$baseArgs+@('--range','B2:D1001','--number-format','@','--format','json'))
    $null=Invoke-TodoLark $context (@('sheets','+dropdown-set')+$baseArgs+@('--range','C2:C1001','--options','-','--highlight=false','--format','json')) (ConvertTo-Json -InputObject @('未完成','已完成') -Compress)
    $null=Invoke-TodoLark $context (@('sheets','+cols-resize')+$baseArgs+@('--widths','-','--format','json')) (ConvertTo-Json -InputObject @{A=120;B=440;C=100;D=260} -Compress)
    $null=Invoke-TodoLark $context (@('sheets','+dim-freeze')+$baseArgs+@('--rows','1','--format','json'))
    $back=Read-TodoSheet $context
    $verify=Invoke-TodoLark $context (@('sheets','+formula-verify')+$baseArgs+@('--range','A1:D1001','--exit-on-error','--format','json'))
    if($verify.data.status -ne 'success'){throw '待办事项检查失败'}
    Complete-FinanceTransaction $tx;Close-FinanceTransaction $tx;$closed=$true
    [pscustomobject]@{ok=$true;applied=$true;preserved_tasks=@($back.items).Count;workbook_url=$context.url}|ConvertTo-Json
}catch{
    if($null -ne $tx -and -not $closed){$rollback=Invoke-FinanceTransactionRollback $tx @{source=$context.token} $context.cli $_.Exception.Message;if(-not $rollback.ok){throw '初始化失败，待办表快照尚待恢复'}}
    throw
}finally{if($null -ne $tx -and -not $closed){Close-FinanceTransaction $tx}}
