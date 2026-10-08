param(
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
$stateRoot = Join-Path $ProjectRoot 'agent\state'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少经营驾驶舱元数据' }
if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
$meta = Get-Content -LiteralPath $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')

function Invoke-LarkJson([scriptblock]$Command, [string]$Label) {
    $raw = & $Command | Out-String
    if ($LASTEXITCODE -ne 0) { throw "$Label（退出码 $LASTEXITCODE）" }
    $result = $raw | ConvertFrom-Json
    if ($result.ok -eq $false) { throw $Label }
    return $result
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}

function Get-StableEventKey([string]$Value, [int]$Length = 20) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString(
            $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value))
        ).Replace('-', '').Substring(0,$Length).ToUpperInvariant()
    } finally {
        $sha.Dispose()
    }
}

$replenishment = Invoke-LarkJson {
    lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'A5:T6' --include value,formula --max-chars 30000 --format json
} '读取补货记录失败'
$sku = Invoke-LarkJson {
    lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1399:R1399' --include value,formula --max-chars 30000 --format json
} '读取SKU成本主表失败'

$replenishmentRows = @($replenishment.data.ranges[0].cells)
$skuRows = @($sku.data.ranges[0].cells)
if ($replenishmentRows.Count -ne 2) { throw '补货记录目标范围结构变化，停止清理' }
if ((Get-CellValue $replenishmentRows[0] 1).ToLowerInvariant() -ne 'kw2131' -or
    (Get-CellValue $replenishmentRows[1] 1).ToLowerInvariant() -ne 'kw2131') {
    throw '补货记录第5/6行已不是预期的KW2131重复记录，停止清理'
}
if ((Get-CellValue $replenishmentRows[0] 2) -ne '12' -or (Get-CellValue $replenishmentRows[1] 2) -ne '12' -or
    (Get-CellValue $replenishmentRows[0] 5) -ne '0' -or (Get-CellValue $replenishmentRows[1] 5) -ne '0') {
    throw 'KW2131重复记录数量或金额已变化，停止清理'
}
if ($skuRows.Count -ne 1 -or (Get-CellValue $skuRows[0] 0).ToLowerInvariant() -ne 'kw2131') {
    throw 'SKU成本主表第1399行已不是KW2131，停止清理'
}

$backupPath = Join-Path $stateRoot ("duplicate_replenishment_kw2131_backup_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
[ordered]@{
    captured_at = [datetime]::Now.ToString('o')
    reason = '清理同一飞书经营事件重放造成的KW2131重复补货'
    replenishment_range = '补货记录!A5:T6'
    replenishment_rows = $replenishmentRows
    sku_range = 'SKU成本主表!A1399:R1399'
    sku_rows = $skuRows
} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $backupPath -Encoding UTF8

$eventId = 'codex-confirmed-F7B6FA14-0-Replenishment'
$batchId = 'EVT-' + (Get-StableEventKey $eventId)
$markerMatrix = @()
$markerMatrix += ,@(@{value=$batchId})
$markerPayload = ConvertTo-Json -InputObject $markerMatrix -Depth 5 -Compress
$null = Invoke-LarkJson {
    $markerPayload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'T5:T5' --cells - --format json
} '写入补货远端幂等标记失败'

$clearPayload = ConvertTo-Json -InputObject @('补货记录!A6:T6') -Compress
$null = Invoke-LarkJson {
    $clearPayload | lark-cli sheets +cells-batch-clear --as user --spreadsheet-token $token --ranges - --scope content --yes --format json
} '清除重复补货行失败'

$rebuildScript = Join-Path $ProjectRoot 'agent\rebuild_cost_snapshots.ps1'
$ledgerScript = Join-Path $ProjectRoot 'agent\sync_daily_ledger.ps1'
$rebuildRaw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $rebuildScript -ProjectRoot $ProjectRoot | Out-String
if ($LASTEXITCODE -ne 0) { throw "成本重算失败：$rebuildRaw" }
$ledgerRaw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ledgerScript -ProjectRoot $ProjectRoot | Out-String
if ($LASTEXITCODE -ne 0) { throw "每日经营简账刷新失败：$ledgerRaw" }

$verifyReplenishment = Invoke-LarkJson {
    lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'A5:T6' --include value,formula --max-chars 30000 --format json
} '回读补货记录失败'
$verifySku = Invoke-LarkJson {
    lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1399:R1399' --include value,formula --max-chars 30000 --format json
} '回读SKU成本主表失败'
$verifyRows = @($verifyReplenishment.data.ranges[0].cells)
$verifySkuRows = @($verifySku.data.ranges[0].cells)
if ((Get-CellValue $verifyRows[0] 19) -ne $batchId) { throw '远端幂等标记回读不一致' }
if (@($verifyRows[1] | Where-Object { $null -ne $_.value -and [string]$_.value -ne '' }).Count -ne 0) { throw '重复补货行未完全清空' }
if ((Get-CellValue $verifySkuRows[0] 6) -ne '12') { throw "KW2131成本账数量应为12，实际为$(Get-CellValue $verifySkuRows[0] 6)" }

$formula = Invoke-LarkJson {
    lark-cli sheets +formula-verify --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'D2:E1001,G2:G1001' --format json --exit-on-error
} '补货记录公式校验失败'
if ($formula.data.status -ne 'success') { throw '补货记录公式校验未通过' }

[ordered]@{
    ok = $true
    backup_path = $backupPath
    kept_row = 5
    cleared_row = 6
    remote_idempotency_marker = $batchId
    sku_row = 1399
    corrected_cost_quantity = 12
    formula_verification = $formula.data.status
    cost_rebuild = ($rebuildRaw.Trim() | ConvertFrom-Json)
    ledger_refresh = ($ledgerRaw.Trim() | ConvertFrom-Json)
} | ConvertTo-Json -Depth 12
