function Test-ActiveInterventionOrderNo([string]$OrderNo) {
    return -not [string]::IsNullOrWhiteSpace($OrderNo) -and $OrderNo -match '^\d{10,30}$'
}

function Get-InterventionTombstoneOrderNo([string]$OrderNo) {
    if (-not (Test-ActiveInterventionOrderNo $OrderNo)) { throw "Invalid intervention order number: $OrderNo" }
    return "WITHDRAWN:$OrderNo"
}
