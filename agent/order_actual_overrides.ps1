Set-StrictMode -Version 2

function Get-OrderActualOverrideStatePath([string]$ProjectRoot) {
    return (Join-Path $ProjectRoot 'agent\state\order_actual_overrides.json')
}

function New-OrderActualOverrideState {
    return [pscustomobject]@{
        version = 1
        updated_at = $null
        overrides = @()
    }
}

function Read-OrderActualOverrideState([string]$ProjectRoot) {
    $path = Get-OrderActualOverrideStatePath $ProjectRoot
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return (New-OrderActualOverrideState) }
    try {
        $state = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
    } catch {
        throw "订单实际履约覆盖文件损坏，已停止算账：$path；$($_.Exception.Message)"
    }
    if ($null -eq $state -or [int]$state.version -ne 1) { throw "订单实际履约覆盖版本无法识别：$path" }
    if ($null -eq $state.overrides) { $state | Add-Member -NotePropertyName overrides -NotePropertyValue @() -Force }
    return $state
}

function Write-OrderActualOverrideState([string]$ProjectRoot, $State) {
    $path = Get-OrderActualOverrideStatePath $ProjectRoot
    $directory = Split-Path $path -Parent
    if (-not (Test-Path -LiteralPath $directory)) { $null = New-Item -ItemType Directory -Path $directory }
    $State.updated_at = [datetime]::Now.ToString('o')
    $temporary = "$path.$PID.tmp"
    try {
        [IO.File]::WriteAllText($temporary, ($State | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $temporary -Destination $path -Force
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue }
    }
}

function Get-OrderActualOverride([string]$ProjectRoot, [string]$OrderNo) {
    return @((Read-OrderActualOverrideState $ProjectRoot).overrides | Where-Object { [string]$_.order_no -eq $OrderNo } | Select-Object -First 1)
}

function Get-OrderActualOverrides([string]$ProjectRoot) {
    return @((Read-OrderActualOverrideState $ProjectRoot).overrides)
}

function Set-OrderActualOverride(
    [string]$ProjectRoot,
    [string]$OrderNo,
    [string]$OrderDate,
    [double]$ActualIncome,
    [double]$DeliveryFee,
    [object[]]$Items,
    [string]$FulfillmentMode,
    [string]$Note
) {
    if (-not $OrderNo) { throw '订单实际履约覆盖缺少订单号' }
    if ($ActualIncome -lt 0 -or $DeliveryFee -lt 0) { throw '订单实际收入和配送费不能小于0' }
    if (@($Items).Count -eq 0) { throw '订单实际履约覆盖至少需要一个实际售出商品' }

    $state = Read-OrderActualOverrideState $ProjectRoot
    $remaining = @($state.overrides | Where-Object { [string]$_.order_no -ne $OrderNo })
    $override = [pscustomobject]@{
        order_no = $OrderNo
        order_date = $OrderDate
        actual_income = [double][math]::Round($ActualIncome, 2)
        delivery_fee = [double][math]::Round($DeliveryFee, 2)
        items = @($Items)
        fulfillment_mode = $FulfillmentMode
        note = $Note
        updated_at = [datetime]::Now.ToString('o')
    }
    $state.overrides = @($remaining) + @($override)
    Write-OrderActualOverrideState $ProjectRoot $state
    return $override
}
