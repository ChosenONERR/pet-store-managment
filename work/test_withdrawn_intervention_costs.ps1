param()
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'agent/intervention_record.ps1')
$script:cases=0
function Check($Pass,$Label){if(-not$Pass){throw $Label};$script:cases++}
function New-ItemRow([string]$Order='3402391430051537057',$Unit=0,$Cost=0,$Qty=1,[string]$Status='已完成'){
    $row=@();for($i=0;$i-lt17;$i++){$row+=[pscustomobject]@{}}
    $row[2]=[pscustomobject]@{value=$Order};$row[3]=[pscustomobject]@{value=$Status};$row[5]=[pscustomobject]@{value='kw4039'};$row[8]=[pscustomobject]@{value=$Qty};$row[11]=[pscustomobject]@{value=$Unit};$row[12]=[pscustomobject]@{value=$Cost};$row[16]=[pscustomobject]@{value='51186993240'}
    return ,$row
}
$id='3402391430051537057';$row=New-ItemRow
$before=$row|ConvertTo-Json -Depth 5 -Compress
$p=@(Get-WithdrawnInterventionCostResets (,$row) $id)
Check ($p.Count-eq1-and$p[0].row-eq2-and$p[0].product_id-eq'51186993240') 'Zero placeholder was not selected'
Check (($row|ConvertTo-Json -Depth 5 -Compress)-eq$before) 'Planner mutated input'
foreach($values in @(@{Unit='0.';Cost='0.00';Qty='1.'},@{Unit=0;Cost=0;Qty=3})){
    $p=@(Get-WithdrawnInterventionCostResets (,(New-ItemRow @values)) $id);Check ($p.Count-eq1) 'Valid zero placeholder missed'
}
foreach($values in @(@{Order='3202391603486436114'},@{Unit=3;Cost=3},@{Unit=0;Cost=2},@{Unit=$null;Cost=$null},@{Unit='bad';Cost=0},@{Qty=0},@{Qty=-1},@{Status='已取消'})){
    $p=@(Get-WithdrawnInterventionCostResets (,(New-ItemRow @values)) $id);Check ($p.Count-eq0) 'Unrelated/genuine snapshot was reset'
}
$reject=$false;try{Get-WithdrawnInterventionCostResets (,$row) 'WITHDRAWN:123'}catch{$reject=$true};Check $reject 'Invalid identity accepted'
$finance=Join-Path $root 'agent/record_finance_event.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($finance,[ref]$null,[ref]$null)
$switch=$ast.FindAll({param($n)$n-is[Management.Automation.Language.SwitchStatementAst]-and$n.Condition.Extent.Text-eq'$Action'},$true)|Select-Object -Last 1
$clause=@($switch.Clauses|Where-Object{$_.Item1.Extent.Text-eq"'RemoveIntervention'"})[0].Item2.Extent.Text
Check ($clause.IndexOf('Get-WithdrawnInterventionCostResets')-lt$clause.IndexOf('if ($DryRun)')) 'Preview does not expose resets'
Check ($clause.IndexOf('foreach($reset in $withdrawnCostResets)')-gt$clause.IndexOf('if ($DryRun)')) 'Preview writes costs'
Check ($clause.Contains("@{value=''},@{value=''},@{value=''}")) 'Explicit clearing of forced zero missing'
# Execute the real withdrawal branch with in-memory sheets and writes.
foreach($name in @('Get-CellValue','Convert-ToNumber','Convert-ToDateValue','New-ValueCell')){
    $f=$ast.Find({param($n)$n-is[Management.Automation.Language.FunctionDefinitionAst]-and$n.Name-eq$name},$true)
    Invoke-Expression $f.Extent.Text
}
$script:mockItemRows=@();$script:mockItemRows+=,(New-ItemRow);$script:mockItemRows+=,(New-ItemRow -Order '3202391603486436114')
$script:mockIntervention=@();for($i=0;$i-lt13;$i++){$script:mockIntervention+=[pscustomobject]@{}}
$script:mockIntervention[0]=[pscustomobject]@{value='2026-11-16'};$script:mockIntervention[3]=[pscustomobject]@{value=$id};$script:mockIntervention[6]=[pscustomobject]@{value=5.87};$script:mockIntervention[7]=[pscustomobject]@{value=0}
$script:mockOrder=@();for($i=0;$i-lt17;$i++){$script:mockOrder+=[pscustomobject]@{}}
$script:mockOrder[0]=[pscustomobject]@{value='2026-11-16'};$script:mockOrder[2]=[pscustomobject]@{value=$id}
$script:mockWrites=@()
function Get-SheetRows($Sheet,$Range){switch($Sheet){'订单汇总'{return ,(,$script:mockOrder)};'干预订单'{return ,(,$script:mockIntervention)};'订单商品'{return $script:mockItemRows};default{throw 'Unexpected read'}}}
function Get-SingleSheetRow($Sheet,$Range){return ,$script:mockIntervention}
function Set-CellMatrix($Sheet,$Range,$Rows){$script:mockWrites+=@{sheet=$Sheet;range=$Range;cells=$Rows};if($Sheet-eq'干预订单'){$script:mockIntervention=$Rows[0]}}
function Complete-CurrentFinanceTransaction {}
$body=[scriptblock]::Create($clause.Substring(1,$clause.Length-2).Replace('exit 0','return')+"`nif(-not `$DryRun){`$writeResult}")
$OrderNo=$id;$Note='Owner explicitly withdraws this order';$DryRun=$true
$preview=((& $body|Out-String)|ConvertFrom-Json)
Check ($preview.dry_run-and$script:mockWrites.Count-eq0) 'Real withdrawal preview wrote data'
Check (@($preview.result.zero_cost_rows_to_reinitialize).Count-eq1) 'Preview reset scope incorrect'
$DryRun=$false;$out=& $body
Check ((@($script:mockWrites|Where-Object{$_.sheet-eq'订单商品'})).Count-eq1) 'Withdrawal changed unrelated item rows'
$resetWrite=@($script:mockWrites|Where-Object{$_.sheet-eq'订单商品'})[0]
Check ($resetWrite.range-eq'L2:N2'-and$resetWrite.cells[0][0].value-eq''-and$resetWrite.cells[0][1].value-eq'') 'Withdrawal did not clear its forced zero'
Check ($script:mockItemRows[1][11].value-eq0) 'Other order zero cost was mutated'
# Execute the production frozen-cost selector, not a rewritten approximation.
$engine=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'agent/rebuild_cost_snapshots.ps1'),[ref]$null,[ref]$null)
$assignment=$engine.Find({param($n)$n-is[Management.Automation.Language.AssignmentStatementAst]-and$n.Left.Extent.Text-eq'$snapshotUnitCost'},$true).Extent.Text
foreach($scenario in @(@{old=$null;avg=14;expected=14},@{old=0;avg=14;expected=0},@{old=5;avg=14;expected=5},@{old=$null;avg=0;expected=0},@{old=$null;avg=$null;expected=$null})){
    $item=[pscustomobject]@{previous_unit_cost=$scenario.old;order_no=$id};$state=[pscustomobject]@{avg=$scenario.avg};$eligibleOrderSet=[Collections.Generic.HashSet[string]]::new();$null=$eligibleOrderSet.Add($id)
    Invoke-Expression $assignment
    Check ($snapshotUnitCost-eq$scenario.expected) 'Historical selector restored wrong cost'
}
$prompt=Get-Content -LiteralPath (Join-Path $root 'agent/codex_chat_bridge.py') -Raw -Encoding UTF8
Check ($prompt.Contains('仅说“是这一单”')-and$prompt.Contains('先询问新订单是新增还是替换')) 'Ambiguous replacement guard missing'
[ordered]@{ok=$true;cases=$script:cases;external_writes=0}|ConvertTo-Json -Compress
