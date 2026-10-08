param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')

function Get-Value($Row,[int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count -or $null -eq @($Row)[$Index].value) { return $null }
    return [string]@($Row)[$Index].value
}
function Serial([datetime]$Date) { return ($Date.Date - [datetime]'1899-12-30').TotalDays }
function Cell($Value) { if ($null -eq $Value -or [string]$Value -eq '') { return @{value=''} }; return @{value=$Value} }
function Get-Rows([string]$Sheet,[string]$Range) {
    $raw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name $Sheet --range $Range --include value --max-chars 1500000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取失败：$Sheet" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "读取不完整：$Sheet" }
    return @($result.data.ranges[0].cells)
}
function Set-Row([string]$Sheet,[string]$Range,[object[]]$Values) {
    $row = [Collections.ArrayList]::new()
    foreach ($value in $Values) { $null = $row.Add((Cell $value)) }
    $matrix = [Collections.ArrayList]::new(); $null = $matrix.Add([object[]]$row.ToArray())
    $payload = ConvertTo-Json -InputObject $matrix -Depth 8 -Compress
    $raw = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $Sheet --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not (($raw | ConvertFrom-Json).ok)) { throw "写入失败：$Sheet $Range" }
}

$master = Get-Rows 'SKU成本主表' 'A2:R3000'
$byKw = @{}
for ($i=0;$i-lt $master.Count;$i++) {
    $kw = ([string](Get-Value $master[$i] 0)).Trim().ToLowerInvariant()
    if ($kw) {
        $unitsPerPack = Get-Value $master[$i] 7
        if (-not $unitsPerPack -or [double]$unitsPerPack -le 0) { $unitsPerPack = 1 }
        $byKw[$kw] = [pscustomobject]@{row=2+$i;product_id=Get-Value $master[$i] 15;units_per_pack=[double]$unitsPerPack}
    }
}
foreach ($required in @('kw1048','kw1106')) {
    if (-not $byKw.ContainsKey($required) -or -not $byKw[$required].product_id) { throw "缺少稳定商品身份：$required" }
}

$repairs = @(
    [pscustomobject]@{row=2;event_code='kw1003';current_kw='kw1106';product_id=$byKw['kw1106'].product_id;units_per_pack=$byKw['kw1106'].units_per_pack;production=[datetime]'2026-08-07';batch_id='BATCH-20260912-0002'},
    [pscustomobject]@{row=3;event_code='kw1048';current_kw='kw1048';product_id=$byKw['kw1048'].product_id;units_per_pack=$byKw['kw1048'].units_per_pack;production=[datetime]'2026-08-13';batch_id='BATCH-20260912-0003'}
)

$replenishmentRows = Get-Rows '补货记录' 'A2:T1001'

$batchRows = Get-Rows '库存批次' 'A2:M2000'
$sourceToRow = @{}
for ($i=0;$i-lt $batchRows.Count;$i++) {
    $source = Get-Value $batchRows[$i] 10
    if ($source) { $sourceToRow[$source] = 2+$i }
}
$nextBatchRow = 2 + @($batchRows | Where-Object { Get-Value $_ 0 }).Count

foreach ($repair in $repairs) {
    $replenishment = @($replenishmentRows)[$repair.row - 2]
    $packageQuantity = [double](Get-Value $replenishment 2)
    $salesUnitQuantity = $packageQuantity * [double]$repair.units_per_pack
    $remainingQuantity = Get-Value $replenishment 9
    if (-not $remainingQuantity) { $remainingQuantity = $salesUnitQuantity }
    # 修正旧模板把“单位成本”误当“每包装销售单位数”的历史显示错误。
    Set-Row '补货记录' "D$($repair.row):D$($repair.row)" @([double]$repair.units_per_pack)
    Set-Row '补货记录' "O$($repair.row):T$($repair.row)" @($repair.product_id,(Serial $repair.production),$null,$null,$null,$repair.batch_id)
    $source = "补货记录第$($repair.row)行"
    $targetBatchRow = if ($sourceToRow.ContainsKey($source)) { [int]$sourceToRow[$source] } else { $nextBatchRow; $nextBatchRow++ }
    Set-Row '库存批次' "A${targetBatchRow}:M${targetBatchRow}" @(
        $repair.batch_id,$repair.product_id,$repair.current_kw,(Serial ([datetime]'2026-09-12')),$salesUnitQuantity,(Serial $repair.production),
        $null,$null,$null,[double]$remainingQuantity,$source,'待补保质期',"历史对话误把生产日期当成到期日，现已纠正；事件发生时货号为 $($repair.event_code)"
    )
}

# kw1048 的 2026-08-13 是生产日期，不是到期日；未提供保质期前不得推算。
Set-Row 'SKU成本主表' "O$($byKw['kw1048'].row):O$($byKw['kw1048'].row)" @($null)

$checkRep = Get-Rows '补货记录' 'A2:T3'
if ((Get-Value $checkRep[0] 14) -ne $byKw['kw1106'].product_id -or (Get-Value $checkRep[1] 14) -ne $byKw['kw1048'].product_id) { throw '补货商品身份回读不一致' }
if (Get-Value (Get-Rows 'SKU成本主表' "O$($byKw['kw1048'].row):O$($byKw['kw1048'].row)")[0] 0) { throw 'kw1048错误到期日未清除' }

[ordered]@{ok=$true;repaired_rows=@(2,3);production_dates=@('2026-08-07','2026-08-13');pending_shelf_life=2} | ConvertTo-Json -Compress
