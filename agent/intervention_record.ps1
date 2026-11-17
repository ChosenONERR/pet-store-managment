function Test-ActiveInterventionOrderNo([string]$OrderNo) {
    return -not [string]::IsNullOrWhiteSpace($OrderNo) -and $OrderNo -match '^\d{10,30}$'
}

function Get-InterventionTombstoneOrderNo([string]$OrderNo) {
    if (-not (Test-ActiveInterventionOrderNo $OrderNo)) { throw "Invalid intervention order number: $OrderNo" }
    return "WITHDRAWN:$OrderNo"
}

# Call only after verifying the order is an active intervention being withdrawn.
# Numeric zero is otherwise a valid frozen natural-sale cost and must survive.
function Get-WithdrawnInterventionCostResets($Rows,[string]$OrderNo) {
    if (-not(Test-ActiveInterventionOrderNo $OrderNo)) { throw 'Invalid withdrawn order identity' }
    $result=@()
    for($i=0;$i-lt@($Rows).Count;$i++){
        $row=$Rows[$i]
        if([string]$row[2].value-ne$OrderNo-or[string]$row[3].value-eq'已取消'){continue}
        $quantity=0.0;$unit=0.0;$cost=0.0
        if(-not[double]::TryParse([string]$row[8].value,[ref]$quantity)-or$quantity-le0){continue}
        if([string]::IsNullOrWhiteSpace([string]$row[11].value)-or[string]::IsNullOrWhiteSpace([string]$row[12].value)){continue}
        if([double]::TryParse([string]$row[11].value,[ref]$unit)-and[double]::TryParse([string]$row[12].value,[ref]$cost)-and$unit-eq0-and$cost-eq0){
            $result+=[pscustomobject]@{row=2+$i;order_no=$OrderNo;goods_code=[string]$row[5].value;product_id=[string]$row[16].value}
        }
    }
    return @($result)
}
