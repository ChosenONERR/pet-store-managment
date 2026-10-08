param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$stateRoot = Join-Path $ProjectRoot 'agent\state'
if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }

function Invoke-LarkJson([scriptblock]$Command,[string]$Label) {
    $raw = & $Command | Out-String
    if ($LASTEXITCODE -ne 0) { throw "$Label（退出码 $LASTEXITCODE）" }
    $result = $raw | ConvertFrom-Json
    if ($result.ok -eq $false) { throw $Label }
    return $result
}
function Get-CellValue($Row,[int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}

$read = Invoke-LarkJson {
    lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'A6:T8' --include formula --max-chars 30000 --format json
} '读取 KW1119 补货记录失败'
$rows = @($read.data.ranges[0].cells)
if ($rows.Count -ne 3 -or (Get-CellValue $rows[0] 1).ToLowerInvariant() -ne 'kw1119') {
    throw '补货记录第6行已不是预期的 KW1119 保留记录，停止清理'
}
$duplicateStillPresent = (Get-CellValue $rows[2] 1) -ne $null
if ((Get-CellValue $rows[0] 2) -ne '2' -or (Get-CellValue $rows[0] 5) -ne '112.3') { throw 'KW1119 保留记录数量或金额已变化，停止清理' }
if ($duplicateStillPresent -and ((Get-CellValue $rows[2] 2) -ne '2' -or (Get-CellValue $rows[2] 5) -ne '112.3')) { throw 'KW1119 重复记录数量或金额已变化，停止清理' }
if ((Get-CellValue $rows[0] 19) -ne 'EVT-1B30CBCA235AB5373744' -or ($duplicateStillPresent -and (Get-CellValue $rows[2] 19) -ne 'EVT-F2FDA48FAE001F531485')) {
    throw 'KW1119 重复记录的幂等标记与预期不符，停止清理'
}

$skuRead = Invoke-LarkJson {
    lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1397:Z1397' --include value --max-chars 10000 --format json
} '读取 KW1119 SKU 档案失败'
$skuRows = @($skuRead.data.ranges[0].cells)
if ($skuRows.Count -ne 1 -or (Get-CellValue $skuRows[0] 0).ToLowerInvariant() -ne 'kw1119') { throw 'SKU成本主表第1397行已不是 KW1119，停止清理' }

$backupPath = Join-Path $stateRoot ("repair_kw1119_a884ceed_backup_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
[ordered]@{
    captured_at=[datetime]::Now.ToString('o');reason='修复A884CEED路径错误后清理诊断重试产生的KW1119重复补货'
    replenishment_range='补货记录!A6:T8';replenishment_rows=$rows
    sku_range='SKU成本主表!A1397:Z1397';sku_rows=$skuRows
} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $backupPath -Encoding UTF8

if ($duplicateStillPresent) {
    $clearPayload = ConvertTo-Json -InputObject @('补货记录!A8:T8') -Compress
    $null = Invoke-LarkJson {
        $clearPayload | lark-cli sheets +cells-batch-clear --as user --spreadsheet-token $token --ranges - --scope content --yes --format json
    } '清除 KW1119 重复补货行失败'
}

$ePayload = '[[{"formula":"=IF(OR(C6=\"\",D6=\"\"),\"\",C6*D6)"}]]'
$gPayload = '[[{"formula":"=IF(OR(E6=\"\",F6=\"\"),\"\",F6/E6)"}]]'
$null = Invoke-LarkJson { $ePayload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'E6:E6' --cells - --format json } '补回 E6 公式失败'
$null = Invoke-LarkJson { $gPayload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'G6:G6' --cells - --format json } '补回 G6 公式失败'

$powershellExe = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$rebuildScript = Join-Path $ProjectRoot 'agent\rebuild_cost_snapshots.ps1'
$ledgerScript = Join-Path $ProjectRoot 'agent\sync_daily_ledger.ps1'
$rebuildRaw = & $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $rebuildScript -ProjectRoot $ProjectRoot | Out-String
if ($LASTEXITCODE -ne 0) { throw "成本重算失败：$rebuildRaw" }
$ledgerRaw = & $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ledgerScript -ThroughDate '2026-09-15' -ProjectRoot $ProjectRoot | Out-String
if ($LASTEXITCODE -ne 0) { throw "每日经营简账刷新失败：$ledgerRaw" }

$verify = Invoke-LarkJson {
    lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'A6:T8' --include formula --max-chars 30000 --format json
} '回读 KW1119 补货记录失败'
$verifyRows = @($verify.data.ranges[0].cells)
if ((Get-CellValue $verifyRows[0] 4) -ne '2' -or (Get-CellValue $verifyRows[0] 6) -ne '56.15') { throw 'KW1119 保留行公式结果不正确' }
if (@($verifyRows[2] | Where-Object { $_.value -ne $null -and [string]$_.value -ne '' }).Count -ne 0) { throw 'KW1119 重复行未完全清空' }

$skuVerify = Invoke-LarkJson {
    lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1397:H1397' --include value --max-chars 10000 --format json
} '回读 KW1119 SKU 档案失败'
$skuVerifyRow = @($skuVerify.data.ranges[0].cells)[0]
if ((Get-CellValue $skuVerifyRow 6) -ne '56.15' -or (Get-CellValue $skuVerifyRow 7) -ne '2') { throw 'KW1119 移动平均成本或成本账数量未恢复为单次补货结果' }

$formula = Invoke-LarkJson {
    lark-cli sheets +formula-verify --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'D2:E1001,G2:G1001' --format json --exit-on-error
} '补货记录公式校验失败'
if ($formula.data.status -ne 'success') { throw '补货记录公式校验未通过' }

[ordered]@{
    ok=$true;backup_path=$backupPath;kept_row=6;cleared_row=8;sku_row=1397
    package_quantity=2;sales_unit_quantity=2;moving_average_cost=56.15;cost_book_quantity=2
    formula_verification=$formula.data.status
} | ConvertTo-Json -Depth 8
