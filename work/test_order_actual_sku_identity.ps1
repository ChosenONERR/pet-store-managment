$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$path=Join-Path $root 'agent/record_finance_event.ps1'
$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$errors)
if(@($errors).Count){throw 'Source parse failed'}
$functions=@{}
foreach($fn in @($ast.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst]},$true))){$functions[$fn.Name]=$fn}
foreach($name in @('Get-CellValue','Convert-ToNumber','Convert-ToDateValue','Resolve-UniqueSkuReference','Merge-OrderActualOriginalLines','Resolve-OrderActualSkuReference','Assert-NonNegative')) {
    $n=@($functions[$name])
    if($n.Count -ne 1){throw "Missing function $name"}
    $body=$n[0].Body.Extent.Text.Substring(1,$n[0].Body.Extent.Text.Length-2)
    if(@($n[0].Parameters).Count){$body='param('+(@($n[0].Parameters|ForEach-Object{$_.Extent.Text}) -join ',')+")`n"+$body}
    Set-Item "function:script:$name" ([scriptblock]::Create($body))
}
$fields=@{CurrentKw=0;PlatformTitle=2;CustomTitle=3;BaselineUnitCost=5;MovingAverageCost=6;ProductId=16;CurrentSpec=18}
$skuSchema=[pscustomobject]@{}
foreach($f in $fields.GetEnumerator()){$skuSchema|Add-Member NoteProperty $f.Key ([pscustomobject]@{Index=$f.Value})}
function CellRow($Values,[int]$Size=26){
    $r=[object[]]::new($Size);for($i=0;$i -lt $Size;$i++){$r[$i]=@{}}
    foreach($e in $Values.GetEnumerator()){$r[[int]$e.Key]=@{value=$e.Value}}
    return ,$r
}
$rows=@((CellRow @{0='kw4025';2='probiotic';5=1.188;6=1.188;16='pid-ten';18='ten-minimum'}),
    (CellRow @{0='kw4025';2='probiotic';5=99;6=99;16='pid-five';18='five-minimum'}),
    (CellRow @{0='kw4152';2='eye drops';5=3.5;6=3.5;16='pid-eye';18='bottle'}))
$original=@([pscustomobject]@{code='kw4025';product_id='pid-ten';quantity=10;unit_price=2.9},
    [pscustomobject]@{code='kw4152';product_id='pid-eye';quantity=1;unit_price=12})
$script:cases=0
function Check([bool]$Ok,[string]$Label){if(-not $Ok){throw $Label};$script:cases++}
function Reject([scriptblock]$Body,[string]$Label){$thrown=$false;try{$null=&$Body}catch{$thrown=$true};Check $thrown $Label}
$split=@([pscustomobject]@{code='kw4025';product_id='pid-ten';title='pad';spec='L';quantity=55;unit_price=0.7},
    [pscustomobject]@{code='kw4025';product_id='pid-ten';title='pad';spec='L';quantity=1;unit_price=0.7})
$splitBefore=$split|ConvertTo-Json -Compress
$merged=Resolve-OrderActualSkuReference $rows 'kw4025' $split
Check ($merged.original_item.quantity -eq 56 -and $merged.original_item.unit_price -eq 0.7) 'Split same product not merged safely'
Check (($split|ConvertTo-Json -Compress) -eq $splitBefore) 'Merge mutated source lines'
Check ((Resolve-OrderActualSkuReference $rows 'sku_ref:2:pid-ten' $split).original_item.quantity -eq 56) 'Explicit identity split-line resolution failed'
foreach($field in @('product_id','spec','unit_price','quantity')){
    $bad=$split|ConvertTo-Json|ConvertFrom-Json
    $bad[1].$field=switch($field){'product_id'{'pid-five'};'spec'{'XL'};'unit_price'{0.8};'quantity'{0}}
    Reject {Resolve-OrderActualSkuReference $rows 'kw4025' $bad} "Unsafe split $field accepted"
}
$bad=$split|ConvertTo-Json|ConvertFrom-Json;$bad[0].product_id='';$bad[1].product_id=''
Reject {Resolve-OrderActualSkuReference $rows 'kw4025' $bad} 'Missing identity split rows accepted'
Check ([math]::Round(80*0.3725,2) -eq 29.8 -and [math]::Round(56*0.3725+24*0.3725,2) -eq 29.8) 'Paid plus free delivery cost incorrect'
Reject {Resolve-UniqueSkuReference $rows 'kw4025'} 'Legacy duplicate-KW failure not reproduced'
$r=Resolve-OrderActualSkuReference $rows 'kw4025' $original
Check ($r.target.row_number -eq 2 -and $r.original_item.product_id -eq 'pid-ten') 'Original stable identity not used'
Check ((Get-CellValue $r.target.cells 6) -eq 1.188) 'Wrong specification cost chosen'
$explicit=Resolve-OrderActualSkuReference $rows 'sku_ref:2:pid-ten' $original
Check ($explicit.original_item.unit_price -eq 2.9) 'Explicit reference lost original sale price'
$other=Resolve-OrderActualSkuReference $rows 'sku_id:pid-five' $original
Check ($null -eq $other.original_item) 'Different product with same KW matched original'
Reject {Resolve-OrderActualSkuReference $rows 'kw4025' @($original[0],([pscustomobject]@{code='kw4025';product_id='pid-five'}))} 'Ambiguous original lines accepted'
Reject {Resolve-OrderActualSkuReference $rows 'kw4025' @([pscustomobject]@{code='kw4025';product_id='missing'})} 'Missing historical identity fell back to KW'
Reject {Resolve-OrderActualSkuReference @($rows[0],$rows[0]) 'kw4025' $original} 'Duplicate master product identity accepted'
Reject {Resolve-OrderActualSkuReference $rows 'sku_ref:2:wrong-id' $original} 'Stale explicit reference accepted'
$legacy=Resolve-OrderActualSkuReference $rows 'kw4152' @([pscustomobject]@{code='kw4152';product_id='';unit_price=12})
Check ($legacy.target.row_number -eq 4 -and $legacy.original_item.unit_price -eq 12) 'Legacy unique-code behavior changed'
$before=$rows|ConvertTo-Json -Depth 8 -Compress
$clauses=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.SwitchStatementAst]},$true)|ForEach-Object{$_.Clauses}|Where-Object{$_.Item1.Extent.Text -eq "'SetOrderActuals'"})
if($clauses.Count -ne 1){throw 'Actuals action block missing'}
$text=$clauses[0].Item2.Extent.Text
$block=[scriptblock]::Create(($text.Substring(1,$text.Length-2) -replace 'exit 0','return'))
function Get-SheetRows([string]$Sheet,[string]$Range){
    switch($Sheet){
        '订单汇总' {return @((CellRow @{0='2026-11-04';1='美团';2='3402373820497768164';7=20.91;9=5.6;10=15.38;16='美团3号单'}),(CellRow @{}))}
        '订单商品' {
            if($script:splitPreview){return @((CellRow @{2='3402373820497768164';5='kw4025';6='probiotic';7='ten';8=9;9=2.9;16='pid-ten'}),
                (CellRow @{2='3402373820497768164';5='kw4025';6='probiotic';7='ten';8=1;9=2.9;16='pid-ten'}),
                (CellRow @{2='3402373820497768164';5='kw4152';6='eye';7='bottle';8=1;9=12;16='pid-eye'}))}
            return @((CellRow @{2='3402373820497768164';5='kw4025';6='probiotic';7='ten';8=10;9=2.9;16='pid-ten'}),
                (CellRow @{2='3402373820497768164';5='kw4152';6='eye';7='bottle';8=1;9=12;16='pid-eye'}))
        }
        'SKU成本主表' {return $rows}
        default {throw 'Unexpected table read'}
    }
}
function Complete-CurrentFinanceTransaction {$script:completed++}
function Set-OrderActualOverride {throw 'Preview attempted a real write'}
function Invoke-OrderSyncForDate {throw 'Preview attempted a real sync'}
$Action='SetOrderActuals';$OrderNo='3402373820497768164';$Items='kw4025×5、kw4152×1';$Amount=11.96;$DeliveryFee=$null;$DryRun=$true;$script:completed=0
$script:splitPreview=$false
$preview=(& $block|Out-String)|ConvertFrom-Json
Check ($preview.ok -and $preview.dry_run) 'Preview failed'
Check ($preview.result.expected_new_cost -eq 9.44) 'Partial non-delivery cost incorrect'
Check ($preview.result.expected_new_profit -eq -3.08) 'Partial refund profit incorrect'
Check ($preview.result.new_delivery_fee -eq 5.6) 'Unspecified delivery fee changed'
Check ($preview.result.date -eq '2026-11-04') 'Actuals omitted month-local ledger date'
Check ($preview.result.actual_items[0].product_id -eq 'pid-ten' -and $preview.result.actual_items[0].unit_price -eq 2.9) 'Preview lost identity/price'
Check ($script:completed -eq 1) 'Preview did not close safely'
Check (($rows|ConvertTo-Json -Depth 8 -Compress) -eq $before) 'Preview mutated master'
$script:splitPreview=$true;$Items='kw4025×10、kw4152×1、sku_id:pid-five×1';$Amount=26.91
$addition=(&$block|Out-String|ConvertFrom-Json).result
Check ($addition.actual_items[2].unit_price -eq 6 -and $addition.actual_items[0].quantity -eq 10) 'Split paid lines blocked additional-income allocation'
$script:splitPreview=$false;$Amount=11.96
$Items='sku_ref:2:pid-ten×5、kw4152×1'
Check (((&$block|Out-String|ConvertFrom-Json).result.expected_new_profit) -eq -3.08) 'Explicit identity not parsed'
$Items='sku_ref:bad×5、kw4152×1'
Reject {&$block} 'Malformed reference partly accepted'
$Items='kw4025×5、kw4152×1';$rows[0][6]=@{value=0};$rows[0][5]=@{value=0}
Check (((&$block|Out-String|ConvertFrom-Json).result.expected_new_cost) -eq 3.5) 'Explicit zero cost lost'
$rows[0][6]=@{};$rows[0][5]=@{}
Reject {&$block} 'Missing cost treated as free'
$source=[IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
Check ($source -match '\[Console\]::Out.WriteLine\(\(\[ordered\]@\{ok=\$false;error=\$reason\}') 'Structured business error missing'
[pscustomobject]@{ok=$true;cases=$script:cases;external_writes=0}|ConvertTo-Json -Compress
