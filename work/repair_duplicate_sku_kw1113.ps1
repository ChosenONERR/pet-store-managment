param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$stateRoot = Join-Path $ProjectRoot 'agent\state'

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
function Set-Row([string]$Range, $Values) {
    $row = @()
    foreach ($value in $Values) { $row += ,$(if ($null -eq $value -or [string]$value -eq '') { @{} } else { @{value=$value} }) }
    $matrix = @(); $matrix += ,$row
    $payload = ConvertTo-Json -InputObject $matrix -Depth 8 -Compress
    $null = Invoke-LarkJson { $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range $Range --cells - --format json } "写入SKU成本主表失败：$Range"
}

$platformResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A998:R998' --include value --max-chars 10000 --format json } '读取KW1113平台档案失败'
$pendingResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1397:R1397' --include value --max-chars 10000 --format json } '读取KW1113空壳档案失败'
$replenishmentResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'A4:T4' --include value,formula --max-chars 10000 --format json } '读取KW1113补货记录失败'
$platformRow = @($platformResult.data.ranges[0].cells)[0]
$pendingRow = @($pendingResult.data.ranges[0].cells)[0]
$replenishmentRow = @($replenishmentResult.data.ranges[0].cells)[0]
if ((Get-CellValue $platformRow 0).ToLowerInvariant() -ne 'kw1113' -or (Get-CellValue $platformRow 15) -ne '51180062720') { throw '第998行不再是预期的KW1113平台档案' }
if ((Get-CellValue $pendingRow 0).ToLowerInvariant() -ne 'kw1113' -or (Get-CellValue $pendingRow 16) -ne '待平台同步') { throw '第1397行不再是预期的KW1113空壳档案' }
if ((Get-CellValue $replenishmentRow 1).ToLowerInvariant() -ne 'kw1113' -or (Get-CellValue $replenishmentRow 5) -ne '0') { throw 'KW1113补货记录已变化，停止合并' }

$backupPath = Join-Path $stateRoot ("duplicate_sku_kw1113_backup_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
[ordered]@{captured_at=[datetime]::Now.ToString('o');platform_row=998;platform_cells=$platformRow;pending_row=1397;pending_cells=$pendingRow;replenishment_row=4;replenishment_cells=$replenishmentRow} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $backupPath -Encoding UTF8

$mergedNote = @((Get-CellValue $platformRow 13),'用户明确该试吃装成本为0；已合并重复空壳档案') | Where-Object { $_ } | Select-Object -Unique
Set-Row 'E998:F998' @(0,0)
Set-Row 'J998:J998' @('零成本')
Set-Row 'M998:O998' @('用户明确试吃装/赠品成本为0',($mergedNote -join '；'),'2028-02-26')
$clearPayload = ConvertTo-Json -InputObject @('SKU成本主表!A1397:R1397') -Compress
$null = Invoke-LarkJson { $clearPayload | lark-cli sheets +cells-batch-clear --as user --spreadsheet-token $token --ranges - --scope content --yes --format json } '清除KW1113重复空壳档案失败'

$rebuildRaw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ProjectRoot 'agent\rebuild_cost_snapshots.ps1') -ProjectRoot $ProjectRoot | Out-String
if ($LASTEXITCODE -ne 0) { throw "成本重算失败：$rebuildRaw" }
$rebuild = $rebuildRaw.Trim() | ConvertFrom-Json
if (@($rebuild.unresolved_replenishments).Count -ne 0) { throw "合并后仍有未解析补货：$($rebuild.unresolved_replenishments -join '、')" }
$ledgerRaw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ProjectRoot 'agent\sync_daily_ledger.ps1') -ProjectRoot $ProjectRoot | Out-String
if ($LASTEXITCODE -ne 0) { throw "每日经营简账刷新失败：$ledgerRaw" }

$verifyPlatform = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A998:R998' --include value --max-chars 10000 --format json } '回读KW1113平台档案失败'
$verifyPending = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1397:R1397' --include value --max-chars 10000 --format json } '回读KW1113空壳行失败'
$verifiedRow = @($verifyPlatform.data.ranges[0].cells)[0]
$clearedRow = @($verifyPending.data.ranges[0].cells)[0]
if ((Get-CellValue $verifiedRow 5) -ne '0' -or (Get-CellValue $verifiedRow 6) -ne '12' -or (Get-CellValue $verifiedRow 9) -ne '零成本') { throw 'KW1113合并后成本账回读不一致' }
if (@($clearedRow | Where-Object { $null -ne $_.value -and [string]$_.value -ne '' }).Count -ne 0) { throw 'KW1113重复空壳行未清空' }

[ordered]@{ok=$true;backup_path=$backupPath;kept_row=998;cleared_row=1397;cost_quantity=12;unit_cost=0;unresolved_replenishments=@($rebuild.unresolved_replenishments);formula_verification='success';ledger_refresh=($ledgerRaw.Trim()|ConvertFrom-Json)} | ConvertTo-Json -Depth 10
