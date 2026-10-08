param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$ledgerScript = Join-Path $ProjectRoot 'agent\sync_daily_ledger.ps1'
$fees = [ordered]@{
    # Already present in the ledger, but missing the durable source marker.
    '3402291701683862688' = 8.38
    '3802291914174628333' = 6.02
    '2402292044289050098' = 9.08
    '2602292263475034171' = 10.75
    # Lost during the historical-cutoff regression and restored from history.
    '3402292991666683590' = 9.35
    '5402293121421730221' = 4.63
    '3402295941330966736' = 9.70
    '3402296210368044898' = 9.15
    '3202296332380539240' = 6.50
    '4002296543748879356' = 9.30
    '3402296663853116161' = 6.40
    '3302297223400792016' = 7.40
    '3302297563875232949' = 6.48
    '3302298081846081056' = 6.20
    '4102298960223685130' = 18.25
    '4102299151749346084' = 9.70
    '2002299302339243922' = 9.50
    '3402299431696905994' = 5.20
    '3302299491023380387' = 9.08
    '3402299673613017366' = 9.60
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}

function Get-SheetRows([string]$SheetName, [string]$Range) {
    $raw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --include value,formula --max-chars 5000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取飞书失败：$SheetName" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "飞书读取结果不完整：$SheetName" }
    return $result.data.ranges[0]
}

function Set-OneCell([string]$SheetName, [string]$Range, $Value) {
    $payload = ConvertTo-Json -InputObject @(,@(@{value=$Value})) -Depth 6 -Compress
    $raw = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not (($raw | ConvertFrom-Json).ok)) { throw "写入飞书失败：$SheetName $Range" }
}

$range = Get-SheetRows '订单汇总' 'A2:Q2001'
$rows = @($range.cells)
$rowNumbers = @($range.row_indices)
$located = @{}
for ($index = 0; $index -lt $rows.Count; $index++) {
    $row = @($rows[$index])
    $orderNo = Get-CellValue $row 2
    if (-not $orderNo -or -not $fees.Contains($orderNo)) { continue }
    if ($located.ContainsKey($orderNo)) { throw "订单汇总存在重复订单：$orderNo" }
    if ((Get-CellValue $row 3) -eq '已取消' -or (Get-CellValue $row 4) -ne '自然') { throw "目标订单不是有效自然单：$orderNo" }
    $located[$orderNo] = [pscustomobject]@{row=[int]$rowNumbers[$index];note=[string](Get-CellValue $row 15)}
}
if ($located.Count -ne $fees.Count) {
    $missing = @($fees.Keys | Where-Object { -not $located.ContainsKey($_) })
    throw "没有完整定位配送费保护订单：$($missing -join ',')"
}

foreach ($orderNo in $fees.Keys) {
    $target = $located[$orderNo]
    $amount = [double]$fees[$orderNo]
    $cleanNote = [regex]::Replace($target.note, '(?:^|[；;])\s*人工确认配送费：?\s*\d+(?:\.\d+)?\s*元?', '').Trim('；',';',' ')
    $marker = "人工确认配送费：$($amount.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture))元"
    $newNote = (@($cleanNote,$marker) | Where-Object { $_ }) -join '；'
    Set-OneCell '订单汇总' "J$($target.row):J$($target.row)" $amount
    Set-OneCell '订单汇总' "P$($target.row):P$($target.row)" $newNote
}

$syncOutput = & $ledgerScript -ProjectRoot $ProjectRoot -ThroughDate ([datetime]::Today) | Out-String
if ($LASTEXITCODE -ne 0) { throw "每日经营简账同步失败：$syncOutput" }

$verifySource = Get-SheetRows '订单汇总' 'A2:Q2001'
$verified = @{}
foreach ($row in @($verifySource.cells)) {
    $orderNo = Get-CellValue $row 2
    if (-not $orderNo -or -not $fees.Contains($orderNo)) { continue }
    $actual = [double](Get-CellValue $row 9)
    $note = [string](Get-CellValue $row 15)
    if ([math]::Abs($actual - [double]$fees[$orderNo]) -gt 0.001) { throw "订单汇总配送费回读不一致：$orderNo" }
    if ($note -notmatch '人工确认配送费') { throw "订单汇总缺少人工配送费保护标记：$orderNo" }
    $verified[$orderNo] = $actual
}
if ($verified.Count -ne $fees.Count) { throw '订单汇总配送费回读数量不足' }

[pscustomobject]@{
    ok = $true
    restored_orders = 16
    restored_total = 136.44
    protected_orders = $verified.Count
    protected_total = [math]::Round((($fees.Values | Measure-Object -Sum).Sum),2)
    source_protected = $true
    ledger_sync = 'success'
} | ConvertTo-Json -Compress
