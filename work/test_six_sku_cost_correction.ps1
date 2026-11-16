param()
$ErrorActionPreference='Stop'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'correct_six_sku_costs_20261116.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Correction script syntax invalid'}
$names=@('C','Date-Key','In-Scope','Code','Num','Near','Get-CorrectionPlan','Get-LedgerCostCorrection')
foreach($f in $ast.FindAll({param($n)$n-is[Management.Automation.Language.FunctionDefinitionAst]},$true)){if($f.Name-in$names){Invoke-Expression $f.Extent.Text}}
$script:cases=0
function Check([scriptblock]$Body){& $Body;$script:cases++}
function Equal($Actual,$Expected){if($Actual-ne$Expected){throw "Expected $Expected but got $Actual"}}
function Reject([scriptblock]$Body){$failed=$false;try{& $Body|Out-Null}catch{$failed=$true};if(-not$failed){throw 'Expected rejection'};$script:cases++}
$profiles=@()
$ids=[ordered]@{kw1090='51185971712';kw1107='51185971714';kw1108='51189679003';kw1091='51189679004';kw1048='51186449393';kw1114='52586891723'}
$costs=@{kw1090=17.78;kw1107=17.78;kw1108=22.23;kw1091=22.23;kw1048=6.12;kw1114=24.44}
foreach($code in $ids.Keys){$profiles+=[pscustomobject]@{code=$code;product_id=$ids[$code];new_cost=$costs[$code];title="Title-$code";spec="Spec-$code";old_moving=24.44;old_baseline=24.44;row=100+$profiles.Count}}
$profiles+=[pscustomobject]@{code='kw1107';product_id='51185971723';title='Unrelated dog food';spec='2kg';new_cost=17.78;old_moving=51;old_baseline=51;row=4}
function New-Sale([string]$Code='kw1107',[string]$Id='51185971714'){
    return [pscustomobject]@{row=10;date='2026-11-15';order_no='3202389680435550906';daily_no='美团1号单';code=$Code;product_id=$Id;title="Title-$Code (Spec-$Code)";spec="Spec-$Code";quantity='1.';old_unit=0;old_cost=0;type='自然';status='已完成'}
}
function New-Orders($Sale,[double]$Cost=9){
    $row=@();for($i=0;$i-lt17;$i++){$row+=@{}}
    $row[0]=@{value=$Sale.date};$row[2]=@{value=$Sale.order_no};$row[3]=@{value='已完成'};$row[4]=@{value='自然'};$row[7]=@{value=40};$row[9]=@{value=5};$row[10]=@{value=$Cost;formula='=SUMIF()'};$row[11]=@{value=35-$Cost;formula='=H2-J2-K2'}
    $cells=@();$cells+=,@();$cells+=,$row
    return ([pscustomobject]@{cells=$cells;row_indices=@(1,2)}|ConvertTo-Json -Depth 8|ConvertFrom-Json)
}
$sale=New-Sale;$orders=New-Orders $sale
$plan=Get-CorrectionPlan $profiles @($sale) @() $orders
Check {Equal $plan.profiles.Count 6}
Check {Equal (@($plan.profiles|Where-Object{$_.row-eq4}).Count) 0}
Check {Equal $plan.lines[0].new_unit 17.78}
Check {Equal $plan.lines[0].new_cost 17.78}
Check {Equal $plan.orders[0].new_cost 26.78}
Check {Near $plan.orders[0].new_profit 8.22 'profit'}
Check {Equal $plan.orders[0].income 40}
Check {Equal $plan.orders[0].delivery 5}
$sale.old_cost=17.78;$sale.old_unit=17.78;$orders=New-Orders $sale 26.78
$again=Get-CorrectionPlan $profiles @($sale) @() $orders
Check {Equal $again.lines[0].changed $false}
Check {Near $again.orders[0].cost_delta 0 'idempotence'}
$sale=New-Sale 'kw1048' '51186449393';$sale.old_cost=5.5;$sale.old_unit=5.5
$plan=Get-CorrectionPlan $profiles @($sale) @() (New-Orders $sale 10)
Check {Near $plan.orders[0].cost_delta 0.62 'increase'}
$sale.quantity=2;$sale.old_cost=11
$plan=Get-CorrectionPlan $profiles @($sale) @() (New-Orders $sale 15)
Check {Near $plan.lines[0].new_cost 12.24 'quantity'}
$sale=New-Sale 'kw1108' ''
$plan=Get-CorrectionPlan $profiles @($sale) @() (New-Orders $sale)
Check {Equal $plan.lines[0].new_cost 22.23}
$sale.title='Different product'
Reject {Get-CorrectionPlan $profiles @($sale) @() (New-Orders $sale)}
$sale=New-Sale;$sale.type='干预';$sale.old_cost=0
$plan=Get-CorrectionPlan $profiles @($sale) @() (New-Orders $sale)
Check {Equal $plan.lines.Count 0}
Check {Equal $plan.skipped.Count 1}
$sale=New-Sale;$sale.status='已取消'
$plan=Get-CorrectionPlan $profiles @($sale) @() (New-Orders $sale)
Check {Equal $plan.lines.Count 0}
$sale=New-Sale 'kw1107' '51185971723'
$plan=Get-CorrectionPlan $profiles @($sale) @() (New-Orders $sale)
Check {Equal $plan.lines.Count 0}
foreach($change in @(@{quantity=0},@{quantity=''},@{old_cost=$null},@{product_id='unknown'},@{code='kw1091'},@{type='unknown'},@{date='2026-09-30'},@{date='2026-12-01'})){
    $sale=New-Sale;foreach($key in $change.Keys){$sale.$key=$change[$key]}
    Reject {Get-CorrectionPlan $profiles @($sale) @() (New-Orders $sale)}
}
$sale=New-Sale
Reject {Get-CorrectionPlan @($profiles|Where-Object{$_.code-ne'kw1090'}) @($sale) @() (New-Orders $sale)}
Reject {Get-CorrectionPlan (@($profiles)+@($profiles[0])) @($sale) @() (New-Orders $sale)}
$orders=New-Orders $sale;$orders.cells[1][10]=[pscustomobject]@{value=9}
Reject {Get-CorrectionPlan $profiles @($sale) @() $orders}
$orders=New-Orders $sale;$orders.cells[1][3]=[pscustomobject]@{value='已取消'}
Reject {Get-CorrectionPlan $profiles @($sale) @() $orders}
$off=[pscustomobject]@{row=25;date='2026-11-15';order_no='OFF-20261115-X';items='kw1090×1';income=29;delivery=0;old_cost=24.44;old_profit=4.56;note='original'}
$plan=Get-CorrectionPlan $profiles @() @($off) (New-Orders $sale)
Check {Equal $plan.offline[0].new_cost 17.78}
Check {Equal $plan.offline[0].new_profit 11.22}
Check {Near $plan.offline[0].cost_delta -6.66 'offline delta'}
$off.items='kw1090×1、kw1001×1'
Reject {Get-CorrectionPlan $profiles @() @($off) (New-Orders $sale)}
$off.items='kw1090×0'
Reject {Get-CorrectionPlan $profiles @() @($off) (New-Orders $sale)}
Check {Equal (Date-Key 46341) '2026-11-15'}
$ledgerRow=@([pscustomobject]@{},[pscustomobject]@{},[pscustomobject]@{},[pscustomobject]@{value=30.1},[pscustomobject]@{value=20.976},[pscustomobject]@{value=4.08},[pscustomobject]@{formula='=D2-E2-F2';value=5.044})
$ledgerPlan=[pscustomobject]@{income=30.1;delivery=4.08;cost_delta=0.62}
$corrected=Get-LedgerCostCorrection $ledgerRow $ledgerPlan
Check {Near $corrected.cost 21.596 'keep fractional cents'}
Check {Near $corrected.profit 4.424 'fractional-cent profit'}
$ledgerPlan.cost_delta=0
Check {Near (Get-LedgerCostCorrection $ledgerRow $ledgerPlan).cost 20.976 'unchanged cost'}
$ledgerPlan.cost_delta=-6.66
Check {Near (Get-LedgerCostCorrection $ledgerRow $ledgerPlan).cost 14.316 'cost reduction'}
$ledgerPlan.income=31
Reject {Get-LedgerCostCorrection $ledgerRow $ledgerPlan}
$ledgerPlan.income=30.1;$ledgerPlan.delivery=0
Reject {Get-LedgerCostCorrection $ledgerRow $ledgerPlan}
$ledgerPlan.delivery=4.08;$ledgerRow[6]=[pscustomobject]@{value=5.044}
Reject {Get-LedgerCostCorrection $ledgerRow $ledgerPlan}
Check {if($ast.Extent.Text-match'sync_daily_ledger\.ps1|rebuild_cost_snapshots\.ps1'){throw 'One-off correction must not rebuild unrelated rows'}}
[ordered]@{ok=$true;cases=$script:cases;external_writes=0}|ConvertTo-Json -Compress
