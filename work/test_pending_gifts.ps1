$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $projectRoot 'agent\reconcile_pending_gifts.ps1'

function New-TestCell($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return @{} }
    return @{value=$Value}
}
function New-TestRow([object[]]$Values, [int]$Width) {
    $row = @()
    for ($index=0; $index -lt $Width; $index++) {
        $value = if ($index -lt $Values.Count) { $Values[$index] } else { $null }
        $row += ,(New-TestCell $value)
    }
    return ,$row
}

$todaySerial = ([datetime]::Now - [datetime]'1899-12-30').TotalDays
$oldSerial = ([datetime]::Now.AddHours(-25) - [datetime]'1899-12-30').TotalDays
$orderRows = @()
$orderRows += ,(New-TestRow @('2026-09-13','美团','111111111111111111','已完成','自然',$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,'美团1号单') 17)
$orderRows += ,(New-TestRow @('2026-09-13','美团','222222222222222222','已取消','自然',$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,'美团2号单') 17)
$giftRows = @()
$giftRows += ,(New-TestRow @('2026-09-13','美团','待同步','111111111111111111','kw1001',1,5,5,'已计成本','等待订单同步',$todaySerial,'event-1') 12)
$giftRows += ,(New-TestRow @('2026-09-13','美团','待同步','222222222222222222','kw1002',1,6,6,'已计成本','等待订单同步',$todaySerial,'event-2') 12)
$giftRows += ,(New-TestRow @('2026-09-12','美团','待同步','333333333333333333','kw1003',1,7,7,'已计成本','等待订单同步',$oldSerial,'event-3') 12)

function lark-cli {
    $arguments = @($args | ForEach-Object { [string]$_ })
    $sheetIndex = [Array]::IndexOf($arguments,'--sheet-name')
    $sheetName = if ($sheetIndex -ge 0) { $arguments[$sheetIndex+1] } else { '' }
    $cells = switch ($sheetName) {
        '订单汇总' { $orderRows }
        '线下订单' { @() }
        '赠品记录' { $giftRows }
        default { throw "测试桩收到未知工作表：$sheetName" }
    }
    [ordered]@{ok=$true;data=[ordered]@{has_more=$false;ranges=@([ordered]@{cells=$cells})}} | ConvertTo-Json -Depth 30 -Compress
    $global:LASTEXITCODE = 0
}

$source = [IO.File]::ReadAllText($scriptPath,[Text.Encoding]::UTF8)
$entry = [ScriptBlock]::Create($source)
$output = @(& $entry -ProjectRoot $projectRoot -GraceHours 24 -DryRun)
$jsonLine = @($output | ForEach-Object {[string]$_} | Where-Object {$_.Trim().StartsWith('{')} | Select-Object -Last 1)
if (-not $jsonLine) { throw '待匹配赠品测试没有返回结果' }
$result = $jsonLine | ConvertFrom-Json
if (-not $result.ok -or $result.matched_order_count -ne 1) { throw '有效订单自动匹配测试失败' }
if (@($result.cancelled_orders) -notcontains '222222222222222222') { throw '取消订单隔离测试失败' }
if (@($result.overdue_orders) -notcontains '333333333333333333') { throw '24小时未匹配告警测试失败' }
if ($result.updated_cells -ne 4) { throw "待匹配赠品更新操作数量异常：$($result.updated_cells)" }

[ordered]@{ok=$true;matched='valid-order';cancelled='excluded';overdue='alert-after-24h';external_writes=0} | ConvertTo-Json -Depth 5
