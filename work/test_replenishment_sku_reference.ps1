$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'agent/record_finance_event.ps1'),[ref]$null,[ref]$parseErrors)
if (@($parseErrors).Count) { throw 'Finance source failed to parse' }
# Load only real pure helpers; never execute the live entry point or a CLI.
$names = @('Get-CellValue','Convert-ToNumber','Convert-ToNullableNumber','Assert-Positive','Assert-NonNegative','New-ValueCell','New-UnitCostCell','Copy-CellRow','Set-WorkingRow','Convert-BatchDate','Get-BatchArgument','Get-SkuCodeCandidates','Resolve-UniqueSkuReference','Get-FirstEmptyRow','Get-StableEventKey','Convert-ToExcelSerial','Test-NullableNumberEqual','Invoke-BatchSkuMaintenance')
foreach ($name in $names) {
    $node = @($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true))
    if ($node.Count -ne 1) { throw "Missing helper: $name" }
    Invoke-Expression $node[0].Extent.Text
}
$indices = [ordered]@{CurrentKw=0;TransitionCode=1;PlatformTitle=2;CustomTitle=3;InitialInventory=4;BaselineUnitCost=5;MovingAverageCost=6;CostBookQuantity=7;UnitsPerPack=8;PurchaseUnit=9;CostStatus=10;InitializationDate=11;LastCostUpdate=12;ConversionBasis=13;Note=14;EarliestExpiry=15;ProductId=16;PlatformStatus=17;CurrentSpec=18}
$skuSchema = [pscustomobject]@{}
foreach ($item in $indices.GetEnumerator()) { $skuSchema | Add-Member NoteProperty $item.Key ([pscustomobject]@{Index=$item.Value;Column=[string][char](65+$item.Value)}) }
function New-MockRow([string]$Code,[string]$Id,$Cost,[double]$Quantity) {
    $row = Copy-CellRow $null 26
    $row[0]=New-ValueCell $Code; $row[16]=New-ValueCell $Id
    $row[6]=New-UnitCostCell $Cost; $row[7]=New-UnitCostCell $Quantity; $row[8]=New-ValueCell 1
    # ASCII-only test source remains valid in Windows PowerShell 5.1.
    $row[17]=New-ValueCell ([string]([char]0x552E)+[char]0x5356+[char]0x4E2D)
    return ,$row
}
function Reset-Mocks {
    $script:skuMock = [Collections.ArrayList]::new()
    $null=$script:skuMock.Add((New-MockRow 'kw4171' 'old-can' 7 1))
    $null=$script:skuMock.Add((New-MockRow 'kw4171' 'new-box' $null 3))
    $script:replenishmentMock=[Collections.ArrayList]::new()
    $script:writeCalls=0; $script:readCalls=0; $script:formulaChecks=0
}
function Get-SheetRows([string]$SheetName,[string]$Range) {
    $script:readCalls++
    if($Range -eq 'A2:Z3000'){$rows=$script:skuMock}else{$rows=$script:replenishmentMock}
    foreach($row in $rows){Write-Output -NoEnumerate $row}
}
function Convert-MockColumn([string]$Column) { $n=0;foreach($c in $Column.ToCharArray()){$n=$n*26+[int]$c-64};return $n-1 }
function Set-AtomicCellOperations([object[]]$Operations) {
    $script:writeCalls++
    foreach($op in $Operations){
        $write=$op.input
        if($write.range -notmatch '^([A-Z]+)(\d+):([A-Z]+)(\d+)$'){throw 'Unexpected test range'}
        $column=Convert-MockColumn $Matches[1];$rowNumber=[int]$Matches[2]
        if($write.sheet_name -eq ([string]([char]0x0053)+'KU'+[char]0x6210+[char]0x672C+[char]0x4E3B+[char]0x8868)){$rows=$script:skuMock}else{$rows=$script:replenishmentMock}
        while($rows.Count -le $rowNumber-2){$null=$rows.Add((Copy-CellRow $null 26))}
        $target=$rows[$rowNumber-2]
        for($c=0;$c -lt @($write.cells[0]).Count;$c++){$target[$column+$c]=$write.cells[0][$c]}
    }
}
function Get-SingleSheetRow([string]$SheetName,[string]$Range) {
    if($Range -notmatch '^A(\d+):'){throw 'Unexpected readback range'}
    if($Range.EndsWith(':Z'+$Matches[1])){$rows=$script:skuMock}else{$rows=$script:replenishmentMock}
    return ,$rows[[int]$Matches[1]-2]
}
function Assert-FormulaRange { $script:formulaChecks++ }
function New-Op([string]$Action,[string]$Code,[double]$Quantity) {
    return [pscustomobject]@{action=$Action;arguments=[pscustomobject]@{goods_code=$Code;quantity=$Quantity;purchase_amount=21.35;date='2026-11-04';event_id="test-$Action-$Code"}}
}
function Assert-Rejected($Operations,[string]$Pattern) {
    $rejected=$false
    try{$null=Invoke-BatchSkuMaintenance $Operations}catch{$rejected=$_.Exception.Message -match $Pattern}
    if(-not $rejected -or $script:writeCalls){throw 'Invalid reference did not stop before writes'}
}
$Date=[datetime]'2026-11-04';$EventId='offline-test';$DryRun=$false
Reset-Mocks
Assert-Rejected @((New-Op 'Replenishment' 'kw4171' 3)) 'kw4171'
Reset-Mocks
Assert-Rejected @((New-Op 'Replenishment' 'sku_ref:3:wrong-product' 3)) '3'
Reset-Mocks
Assert-Rejected @((New-Op 'Replenishment' 'sku_ref:99:new-box' 3)) '99'
Reset-Mocks
$op=New-Op 'Replenishment' 'sku_ref:3:new-box' 3
$op.arguments|Add-Member NoteProperty transition_code 'csx123'
Assert-Rejected @($op) '.+'
Reset-Mocks
$old=ConvertTo-Json -InputObject $script:skuMock[0] -Depth 6 -Compress
$ops=@((New-Op 'SetCostBookQuantity' 'sku_ref:3:new-box' 0),(New-Op 'Replenishment' 'sku_ref:3:new-box' 3))
$DryRun=$true;$preview=Invoke-BatchSkuMaintenance $ops
if($script:writeCalls -ne 0 -or $preview.operations[1].after_quantity -ne 3 -or [math]::Abs($preview.operations[1].after_average_cost-21.35/3) -gt 0.000001){throw 'Preview quantity or cost incorrect'}
$DryRun=$false;$result=Invoke-BatchSkuMaintenance $ops
if($script:writeCalls -ne 1 -or $script:formulaChecks -ne 1){throw 'Expected one atomic write and formula check'}
if((ConvertTo-Json -InputObject $script:skuMock[0] -Depth 6 -Compress) -ne $old){throw 'Old can profile changed'}
if((Get-CellValue $script:skuMock[1] 7) -ne '3' -or $result.operations[1].goods_code -ne 'kw4171'){throw 'Wrong SKU quantity or display code'}
if($script:replenishmentMock.Count -ne 1 -or (Get-CellValue $script:replenishmentMock[0] 14) -ne 'new-box'){throw 'Wrong product in replenishment record'}
Reset-Mocks
$unknown=Invoke-BatchSkuMaintenance @((New-Op 'Replenishment' 'sku_ref:3:new-box' 3))
if($null -ne $unknown.operations[0].after_average_cost -or $unknown.operations[0].after_quantity -ne 6){throw 'Unknown old cost was invented'}
[pscustomobject]@{ok=$true;cases=8;external_writes=0}|ConvertTo-Json -Compress
