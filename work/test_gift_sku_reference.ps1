$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
function Import-TestFunction([string]$Path,[string]$Name){
    $errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$null,[ref]$errors)
    if(@($errors).Count){throw 'Source parse failed'}
    $node=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true))
    if($node.Count -ne 1){throw "Missing function $Name"}
    $body=$node[0].Body.Extent.Text.Substring(1,$node[0].Body.Extent.Text.Length-2)
    if(@($node[0].Parameters).Count){$body='param('+(@($node[0].Parameters|ForEach-Object{$_.Extent.Text}) -join ',')+")`n"+$body}
    Set-Item -Path "function:script:$Name" -Value ([scriptblock]::Create($body))
}
$finance=Join-Path $root 'agent/record_finance_event.ps1'
foreach($n in @('Get-CellValue','Convert-ToNumber','New-ValueCell','Convert-ToExcelSerial','Get-CanonicalSkuMap','Resolve-UniqueSkuReference','New-OrderGiftPlan')){Import-TestFunction $finance $n}
Import-TestFunction (Join-Path $root 'agent/rebuild_cost_snapshots.ps1') 'Resolve-GiftCanonicalIdentity'
$fields=@{CurrentKw=0;PlatformTitle=2;CustomTitle=3;BaselineUnitCost=5;MovingAverageCost=6;CostStatus=10;ProductId=16;CurrentSpec=18}
$skuSchema=[pscustomobject]@{}
foreach($f in $fields.GetEnumerator()){$skuSchema|Add-Member NoteProperty $f.Key ([pscustomobject]@{Index=$f.Value})}
function New-Row([string]$Id,$Cost,[string]$Spec){
    $r=[object[]]::new(26);for($i=0;$i -lt 26;$i++){$r[$i]=@{}}
    $r[0]=@{value='kw2060'};$r[2]=@{value='羊乳猫罐'};$r[6]=New-ValueCell $Cost
    $r[16]=@{value=$Id};$r[18]=@{value=$Spec}
    return ,$r
}
$rows=@((New-Row 'pid-other' $null '多罐装'),(New-Row 'pid-random' 1.4 '随机口味一罐'),(New-Row 'pid-another' $null '另一口味'))
$order=[pscustomobject]@{date=[datetime]'2026-11-04';platform='美团';daily_order_no='美团2号单';order_no='3402373773748593952'}
$count=0
function Check([bool]$Pass,[string]$Label){if(-not $Pass){throw $Label};$script:count++}
function Reject([scriptblock]$Body,[string]$Label){$caught=$false;try{$null=&$Body}catch{$caught=$true};Check $caught $Label}
$before=$rows|ConvertTo-Json -Depth 8 -Compress
$plan=New-OrderGiftPlan 'sku_ref:3:pid-random×1' $rows $order $false '店主确认' 'test-event'
Check ($plan.gift_cost -eq 1.4 -and $plan.gift_items -eq 'kw2060×1') 'Chosen specification cost/display incorrect'
Check ($plan.gift_details[0].master_row -eq 3 -and $plan.gift_details[0].product_id -eq 'pid-random') 'Chosen row/ID not retained'
Check ((Get-CellValue $plan.rows[0] 9) -match '\[gift_sku_id=pid-random\]') 'Gift identity marker missing'
Check ((Get-CellValue $plan.rows[0] 11) -eq 'test-event') 'Idempotency event missing'
Check (($rows|ConvertTo-Json -Depth 8 -Compress) -eq $before) 'Pure planning changed source rows'
$ambiguous=New-OrderGiftPlan 'kw2060×1' $rows $order $false '' 'test'
Check ($null -eq $ambiguous.gift_cost -and @($ambiguous.pending_cost_codes).Count -eq 1) 'Duplicate KW guessed a specification'
Reject {New-OrderGiftPlan 'sku_ref:3:wrong-id×1' $rows $order $false '' 'test'} 'Stale ID accepted'
Reject {New-OrderGiftPlan 'sku_ref:99:pid-random×1' $rows $order $false '' 'test'} 'Missing row accepted'
Reject {New-OrderGiftPlan 'sku_ref:3:pid-random×0' $rows $order $false '' 'test'} 'Zero quantity accepted'
Reject {New-OrderGiftPlan 'sku_ref:bad×1、kw2060×1' $rows $order $false '' 'test'} 'Invalid reference silently dropped'
Reject {New-OrderGiftPlan 'sku_ref:3:pid-random×1' $rows $order $false '[gift_sku_id=pid-other]' 'test'} 'Caller could forge persisted identity'
$unknown=New-OrderGiftPlan 'sku_ref:2:pid-other×1' $rows $order $false '' 'test'
Check ($null -eq $unknown.gift_cost) 'Unknown cost treated as zero'
$zeroRows=@();$zeroRows+=,(New-Row 'zero' 0 '试吃')
$zero=New-OrderGiftPlan 'sku_ref:2:zero×4' $zeroRows $order $false '' 'test'
Check ($null -ne $zero.gift_cost -and $zero.gift_cost -eq 0) 'Explicit zero cost lost'
$many=New-OrderGiftPlan 'sku_ref:3:pid-random×1、sku_ref:3:pid-random×2' $rows $order $false '' 'test'
Check ($many.gift_cost -eq 4.2 -and $many.gift_details[0].quantity -eq 3) 'Repeated gift quantity not aggregated'
$pending=New-OrderGiftPlan 'sku_ref:3:pid-random×1' $rows $order $true '' 'test'
Check ((Get-CellValue $pending.rows[0] 9).StartsWith('等待订单同步')) 'Pending-order evidence missing'
$aliases=@{kw2060='sku:pid-other'};$states=@{'sku:pid-other'=@{};'sku:pid-random'=@{}}
Check ((Resolve-GiftCanonicalIdentity 'kw2060' '[gift_sku_id=pid-random]' $aliases $states) -eq 'sku:pid-random') 'KW reuse moved gift to wrong product'
Check ($null -eq (Resolve-GiftCanonicalIdentity 'kw2060' '[gift_sku_id=deleted-id]' $aliases $states)) 'Missing product fell back to reused KW'
Check ((Resolve-GiftCanonicalIdentity 'kw2060' 'legacy note' $aliases $states) -eq 'sku:pid-other') 'Legacy unambiguous gift changed'
Reject {Resolve-GiftCanonicalIdentity 'kw2060' '[gift_sku_id=pid-random];[gift_sku_id=pid-other]' $aliases $states} 'Conflicting identity markers accepted'
Reject {Resolve-GiftCanonicalIdentity 'kw2060' '[gift_sku_id=broken' $aliases $states} 'Broken identity marker ignored'
$source=[IO.File]::ReadAllText($finance,[Text.Encoding]::UTF8)
Check ($source -match 'if \(\$EventId -and -not \$DryRun\)') 'Duplicate gift dry-run can trigger rebuild'
Check ($source -match "orderScopedCostActions = @\([^\r\n]*'AddOrderGift'") 'Gift confirmation not order-scoped'
$rebuild=[IO.File]::ReadAllText((Join-Path $root 'agent/rebuild_cost_snapshots.ps1'),[Text.Encoding]::UTF8)
Check ($rebuild -match '\$giftEvents \| Where-Object \{ \$eligibleOrderSet.Contains') 'Incremental SKU selection omits gift product'
[pscustomobject]@{ok=$true;cases=$count;external_writes=0}|ConvertTo-Json -Compress
