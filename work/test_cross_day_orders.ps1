$ErrorActionPreference='Stop'
Set-StrictMode -Version 2
$root=Split-Path $PSScriptRoot -Parent
$script:cases=0
function Check([bool]$Pass,[string]$Label){if(-not $Pass){throw $Label};$script:cases++}
function Parse([string]$Path){
    $errors=$null;$tokens=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if(@($errors).Count){throw ($errors.Message -join ';')};return $ast
}
$runPath=Join-Path $root 'agent/run_sync_and_notify.ps1'
$run=Parse $runPath
$source=[IO.File]::ReadAllText($runPath,[Text.Encoding]::UTF8)
$window=$run.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$orderStartDate'},$true)
foreach($fixture in @(
    @{day='2026-11-15';mode='midday';expected='2026-11-14'},
    @{day='2026-11-01';mode='reminder';expected='2026-10-31'},
    @{day='2027-01-01';mode='midday';expected='2026-12-31'},
    @{day='2026-11-15';mode='night';expected='2026-11-08'}
)){
    $Date=[datetime]$fixture.day;$Mode=$fixture.mode;$dateText=$Date.ToString('yyyy-MM-dd')
    Invoke-Expression $window.Extent.Text
    Check ($orderStartDate -eq $fixture.expected) 'Cross-day read window is incorrect'
}
$orderPath=Join-Path $root 'work/sync_meituan_orders.ps1'
$orderSource=[IO.File]::ReadAllText($orderPath,[Text.Encoding]::UTF8)
$start=$orderSource.IndexOf('$priorSnapshot = Take-CostSnapshotFromQueue')
$end=$orderSource.IndexOf('$itemStatus =', $start)
Check ($start -ge 0 -and $end -gt $start) 'Missing real cost snapshot reactivation block'
$resetBlock=$orderSource.Substring($start,$end-$start)
function Take-CostSnapshotFromQueue {return $script:fixtureSnapshot}
foreach($fixture in @(
    @{old='已取消';new='已完成';cost=0;reset=$true},
    @{old='已取消';new='处理中';cost=0;reset=$true},
    @{old='已取消';new='已取消';cost=0;reset=$false},
    @{old='已完成';new='已完成';cost=0;reset=$false},
    @{old='已取消';new='已完成';cost=5;reset=$false}
)){
    $script:fixtureSnapshot=[pscustomobject]@{Status=$fixture.old;UnitCost=$fixture.cost;ItemCost=$fixture.cost}
    $order=[pscustomobject]@{orderNo='3402388953856956469';status=$fixture.new}
    $existingCostSnapshots=@{};$snapshotKey='fixture';$fallbackSnapshotKey='fallback'
    $newItemOrderNumbers=[Collections.Generic.HashSet[string]]::new()
    Invoke-Expression $resetBlock
    Check (($null -eq $priorSnapshot) -eq $fixture.reset) 'Wrong cancellation placeholder reset decision'
    Check ($newItemOrderNumbers.Count -eq [int]$fixture.reset) 'Reactivation did not scope cost initialization to one order'
}
Check ($orderSource -match 'Status = Get-CellValue \$row 3' -and $orderSource -match 'affected_dates =') 'Snapshot status or affected dates not transported'
foreach($name in @('Get-CellValue','Convert-ToNumber','Test-RowDate','Get-DailyNetProfit','Format-CrossDayCorrection')){
    $f=$run.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    Check ($null -ne $f) "Missing $name"
    $body=$f.Body.Extent.Text.Substring(1,$f.Body.Extent.Text.Length-2)
    if($null -ne $f.Parameters -and @($f.Parameters).Count){$body='param('+(@($f.Parameters|ForEach-Object{$_.Extent.Text})-join ',')+")`n"+$body}
    $body=$body.Replace("Join-Path `$PSScriptRoot 'expense_accounting.ps1'","Join-Path `$root 'agent/expense_accounting.ps1'")
    Set-Item "function:script:$name" ([scriptblock]::Create($body))
}
function Row($Values){return ,@($Values|ForEach-Object {[pscustomobject]@{value=$_}})}
$script:income=@();$script:income+=,(Row @('美团9号单','3402388953856956469','2026-11-14',20,6,4,10))
$script:income+=,(Row @('美团1号单','other','2026-11-15',200,60,40,100))
$script:expenses=@();$script:expenses+=,(Row @('2026-11-14','投流推广费',2.01,'','美团推广'))
function Get-LedgerRows($Name,$Range){if($Name -eq '11月收入'){return ,$script:income};if($Name -eq '11月支出'){return ,$script:expenses};throw 'Unexpected ledger read'}
$Date=[datetime]'2026-11-15'
$profit=Get-DailyNetProfit ([datetime]'2026-11-14')
Check ($profit.order_count -eq 1 -and $profit.net_profit -eq 7.99 -and $profit.order_profit -eq 10) 'Correction used today instead of the original business date'
$message=Format-CrossDayCorrection '2026-11-14' $profit
Check ($message -match '2026-11-14' -and $message -match '7.99' -and $message -match '替代此前' -and $message -notmatch '仍有成本或费用待核对') 'Correction omitted the original-date final amount'
$script:income[0][6].value=$null
$profit=Get-DailyNetProfit ([datetime]'2026-11-14')
Check ($profit.pending_cost_count -eq 1 -and (Format-CrossDayCorrection '2026-11-14' $profit) -match '暂估') 'Missing order cost presented as final zero'
Check ($source.IndexOf('Complete-SyncFinanceTransaction', $source.IndexOf('$dailyProfit =')) -lt $source.IndexOf("Invoke-WithRetry 'prior_day_correction'")) 'Correction sent before the finance commit'
Check ($source -match 'if \(-not \$DryRun -and -not \$SkipSync\)' -and $source -match 'foreach \(\$report in \$priorDayReports\)') 'Dry run/no-change resend can emit financial corrections'
[ordered]@{ok=$true;cases=$script:cases;external_writes=0;credential_access=0}|ConvertTo-Json -Compress
