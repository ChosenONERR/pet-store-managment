$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $projectRoot 'agent\sync_new_skus.ps1'
$global:testSkuWrites = [Collections.Generic.List[object]]::new()
$global:testSkuLarkCalls = [Collections.Generic.List[string]]::new()

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

$global:testSkuHeaders = @(
    ,(New-TestRow -Values @('当前KW','CSX/过渡货号','商品标题','自定义标题','初始化库存','基准单位成本','当前移动平均成本','成本账数量','每采购包装含销售单位数','采购单位','成本核对状态','初始化日期','最近成本更新时间','换算依据','备注','当前库存最早到期日期','商品ID（美团SKU ID）','平台销售状态','当前规格') -Width 26)
)
$global:testSkuMasterRows = @(
    ,(New-TestRow -Values @('kw1113',$null,$null,$null,0,0,0,12,1,'个','零成本',46278,46278,'用户明确试吃装成本为0','新货号待平台库存同步补全商品资料',46799,$null,'待平台同步',$null) -Width 26)
    ,(New-TestRow -Values @($null) -Width 26)
)
$global:testSkuInventoryRows = @(
    ,(New-TestRow -Values @('kw1113','试吃装商品',12,0,'售卖中','csx1113',$null,$null,2,'正常','美团商品库存导出',0,'51199999999','50g') -Width 14)
    ,(New-TestRow -Values @('kw1113','试吃装商品',12,0,'售卖中','csx1113',$null,$null,2,'正常','美团商品库存导出',0,'51199999998','另一平台规格') -Width 14)
    ,(New-TestRow -Values @($null) -Width 14)
)

function global:Invoke-TestSetContent { process {} }
Set-Alias -Name Set-Content -Value Invoke-TestSetContent -Scope Global
function global:Invoke-TestLarkCli {
    $arguments = @($args | ForEach-Object {[string]$_})
    $global:testSkuLarkCalls.Add($arguments -join ' ')
    if ($arguments -contains '+cells-get') {
        $sheetIndex = [Array]::IndexOf($arguments,'--sheet-name')
        $sheet = $arguments[$sheetIndex+1]
        $rangeIndex = [Array]::IndexOf($arguments,'--range')
        $requestedRange = if ($rangeIndex -ge 0) { $arguments[$rangeIndex+1] } else { '' }
        $cells = if ($sheet -eq 'SKU成本主表' -and $requestedRange -eq 'A1:Z1') { $global:testSkuHeaders } elseif ($sheet -eq 'SKU成本主表') { $global:testSkuMasterRows } elseif ($sheet -eq '库存快照') { $global:testSkuInventoryRows } else { throw "未知工作表：$sheet" }
        $indices = if ($requestedRange -eq 'A1:Z1') { @(1) } else { @(2,3) }
        $cellsPayload = if ($requestedRange -eq 'A1:Z1') { ,$cells } else { $cells }
        [ordered]@{ok=$true;data=[ordered]@{ranges=@([ordered]@{cells=$cellsPayload;row_indices=$indices})}} | ConvertTo-Json -Depth 30 -Compress
        $global:LASTEXITCODE = 0
        return
    }
    if ($arguments -contains '+cells-set') {
        $rangeIndex = [Array]::IndexOf($arguments,'--range')
        $global:testSkuWrites.Add([pscustomobject]@{range=$arguments[$rangeIndex+1]})
        [ordered]@{ok=$true} | ConvertTo-Json -Compress
        $global:LASTEXITCODE = 0
        return
    }
    throw "未知 lark-cli 调用：$($arguments -join ' ')"
}
Set-Alias -Name lark-cli -Value Invoke-TestLarkCli -Scope Global

$output = @(. $scriptPath -ProjectRoot $projectRoot -SnapshotDate ([datetime]'2026-09-13'))
$jsonLine = @($output | ForEach-Object {[string]$_} | Where-Object {$_.Trim().StartsWith('{')} | Select-Object -Last 1)
if (-not $jsonLine) { throw '新SKU同步测试没有返回JSON' }
$result = $jsonLine | ConvertFrom-Json
$ranges = @($global:testSkuWrites | ForEach-Object {$_.range})
if (-not $result.ok -or $result.enriched_pending_count -ne 1 -or $result.new_kw_count -ne 0) {
    throw "待平台同步SKU补全结果不正确：result=$($result | ConvertTo-Json -Compress)；calls=$($global:testSkuLarkCalls.Count)"
}
if ($ranges -notcontains 'B2:C2' -or $ranges -notcontains 'Q2:Q2' -or $ranges -notcontains 'R2:R2' -or $ranges -notcontains 'S2:S2') { throw "待平台同步SKU没有补全商品资料：$($ranges -join ',')" }
if ($ranges.Count -ne 4) { throw "同一KW的多个平台规格不应造成重复写入：$($ranges -join ',')" }
if (@($ranges | Where-Object {$_ -match '^[DEFGH]2:'}).Count -gt 0) { throw '平台资料补全不应覆盖成本账字段' }

[ordered]@{ok=$true;enriched='title-csx-product-status-spec';cost_ledger_preserved=$true;external_writes=0} | ConvertTo-Json -Depth 5
