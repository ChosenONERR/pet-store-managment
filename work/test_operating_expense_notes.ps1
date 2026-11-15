$ErrorActionPreference='Stop'
Set-StrictMode -Version 2
$root=Split-Path $PSScriptRoot -Parent
$script:cases=0
function Check([bool]$Pass,[string]$Label){if(-not $Pass){throw $Label};$script:cases++}
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'agent/record_finance_event.ps1'),[ref]$tokens,[ref]$errors)
foreach($name in @('Get-CellValue','Convert-ToNumber','Convert-ToDateValue','Convert-ToExcelSerial','Normalize-BusinessText','Test-IsPromotionExpense','Get-OperatingExpenseState','Invoke-OperatingExpenseWrite','Get-FirstCompletelyEmptyRow')){
    $f=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    if($null -eq $f){throw "Missing $name"}
    Invoke-Expression $f.Extent.Text
}
function Row($Values){return ,@($Values|ForEach-Object{[pscustomobject]@{value=$_}})}
function Reset-Fixture {
    $script:grid=@();$script:grid+=,(Row @('2026-11-15','原有骑手红包',2,'','手工','已记账'))
    for($i=0;$i -lt 4;$i++){$script:grid+=,(Row @($null,$null,$null,$null,$null,$null))}
    $script:monthRow=Row @(45962,$null,$null,$null,$null,$null,2,$null,$null,$null,$null,'11-15 其他经营费用 2：原有骑手红包',$null,0)
    $script:writes=0
}
function Get-LedgerSheetRows($Name,$Range){return ,$script:grid}
function Get-MonthRow($TargetDate){return [pscustomobject]@{row_number=5;cells=$script:monthRow}}
function Set-LedgerOneRow($Name,$Range,$Values){
    if($Range -notmatch '^A(\d+):F\1$'){throw 'Unexpected ledger range'}
    $script:grid[[int]$Matches[1]-2]=Row $Values;$script:writes++
}
function Set-OneRow($Name,$Range,$Values){
    $index=switch($Range){'G5:G5'{6};'L5:L5'{11};'N5:N5'{13};default{throw 'Unexpected month range'}}
    $script:monthRow[$index].value=$Values[0];$script:writes++
}
$ledgerSpreadsheetToken='fixture';$OrderNo='';$ExpenseSource='手工';$ExpenseMemo='通过飞书确认录入';$ReplaceMatchingProject=$false
foreach($note in @(
    '骑手红包；刚刚又给骑手2元，新增一笔，不覆盖此前记录',
    '骑手红包;再次追加;不是覆盖',
    '骑手红包；含中文;英文分号和日期11-15说明',
    '普通杂费'
)){
    Reset-Fixture
    $preview=Invoke-OperatingExpenseWrite ([datetime]'2026-11-15') 2 $note -PreviewOnly
    Check ($preview.status -eq 'ready' -and $script:writes -eq 0) 'Preview wrote data'
    $result=Invoke-OperatingExpenseWrite ([datetime]'2026-11-15') 2 $note
    Check ($result.status -eq 'recorded' -and $result.changed) 'Actual write/readback failed'
    Check ($script:monthRow[6].value -eq 4) 'Existing bonus overwritten or incremented twice'
    Check ($script:grid[0][2].value -eq 2 -and $script:grid[1][1].value -eq $note) 'Original bonus or description changed'
    $state=Get-OperatingExpenseState ([datetime]'2026-11-15') 2 $note
    Check ($state.ledger_match_count -eq 1 -and $state.month_note_count -eq 1) 'Semicolon split the new expense entry'
    $writes=$script:writes
    $again=Invoke-OperatingExpenseWrite ([datetime]'2026-11-15') 2 $note
    Check ($again.status -eq 'duplicate' -and $script:writes -eq $writes -and $script:monthRow[6].value -eq 4) 'Replay duplicated expense'
    $script:monthRow[11].value+=';11-14 其他经营费用 3：旧杂费；11-15 其他经营费用 4：其他项目'
    $state=Get-OperatingExpenseState ([datetime]'2026-11-15') 2 $note
    Check ($state.month_note_count -eq 1 -and $state.month_note_parts.Count -eq 4) 'Legacy Chinese/ASCII record delimiters not recognized'
    $script:monthRow[11].value+='；11-15 其他现金支出 8：现金付款；11-15 老板取款 10'
    $state=Get-OperatingExpenseState ([datetime]'2026-11-15') 2 $note
    Check ($state.month_note_count -eq 1 -and $state.month_note_parts.Count -eq 6) 'Other monthly events merged into expense entry'
}
Reset-Fixture
$note='骑手红包；新增一笔'
$null=Invoke-OperatingExpenseWrite ([datetime]'2026-11-15') 2 $note
$script:grid[2]=Row @('2026-11-15',$note,2,'','手工','重复')
$blocked=$false;try{$null=Invoke-OperatingExpenseWrite ([datetime]'2026-11-15') 2 $note}catch{$blocked=$_.Exception.Message -match '重复'}
Check $blocked 'Duplicate ledger guard lost'
Reset-Fixture
$script:grid[1]=Row @('2026-11-15',$note,2,'','手工','缺月度条目')
$blocked=$false;try{$null=Invoke-OperatingExpenseWrite ([datetime]'2026-11-15') 2 $note}catch{$blocked=$_.Exception.Message -match '跨表状态不一致'}
Check ($blocked -and $script:writes -eq 0) 'Cross-table inconsistency guard lost'
Reset-Fixture
$OrderNo='3302389813279749611'
$result=Invoke-OperatingExpenseWrite ([datetime]'2026-11-15') 2 '取消派单赔付；用户已自配送'
$state=Get-OperatingExpenseState ([datetime]'2026-11-15') 2 '取消派单赔付；用户已自配送'
Check ($result.status -eq 'recorded' -and $state.month_note_count -eq 1 -and $script:grid[1][3].value -eq $OrderNo) 'Order-linked description broken'
[ordered]@{ok=$true;cases=$script:cases;external_writes=0;credential_access=0}|ConvertTo-Json -Compress
