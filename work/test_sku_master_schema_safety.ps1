param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$activeScripts = @(
    'agent\rebuild_cost_snapshots.ps1',
    'agent\record_finance_event.ps1',
    'agent\sync_sku_master_by_identity.ps1',
    'agent\sync_new_skus.ps1',
    'agent\run_sync_and_notify.ps1'
)
$errors = [Collections.Generic.List[string]]::new()
foreach ($relativePath in $activeScripts) {
    $path = Join-Path $ProjectRoot $relativePath
    $content = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    if ($content -notmatch 'sku_master_schema\.ps1' -or $content -notmatch 'Get-SkuMasterSchema') {
        $errors.Add("未使用表头映射：$relativePath")
    }
    if ($content -match "sheet-name 'SKU成本主表'.*--range 'A2:R") {
        $errors.Add("仍按旧18列读取：$relativePath")
    }
}
$rebuild = Get-Content -LiteralPath (Join-Path $ProjectRoot 'agent\rebuild_cost_snapshots.ps1') -Raw -Encoding UTF8
if ($rebuild -match 'kw2054历史订单成本抽查失败') { $errors.Add('仍存在kw2054硬编码断言') }
$orchestrator = Get-Content -LiteralPath (Join-Path $ProjectRoot 'agent\run_sync_and_notify.ps1') -Raw -Encoding UTF8
$inventorySync = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\sync_meituan_inventory.ps1') -Raw -Encoding UTF8
$identitySync = Get-Content -LiteralPath (Join-Path $ProjectRoot 'agent\sync_sku_master_by_identity.ps1') -Raw -Encoding UTF8
if ($orchestrator -match '\[ScriptBlock\]::Create') { $errors.Add('总调度器仍通过动态脚本块调用子脚本，会丢失PSScriptRoot') }
if ($inventorySync -match '\[ScriptBlock\]::Create') { $errors.Add('库存同步仍通过动态脚本块调用SKU同步，会丢失PSScriptRoot') }
if ($identitySync -notmatch "PlatformStatus\.Index\) -eq '待平台同步'") { $errors.Add('SKU身份同步缺少待平台占位档案复用规则') }

$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $ProjectRoot 'agent\sku_master_schema.ps1')
$schema = Get-SkuMasterSchema -SpreadsheetToken $token
$expected = [ordered]@{BaselineUnitCost='F';MovingAverageCost='G';CostBookQuantity='H';InitializationDate='L';LastCostUpdate='M';EarliestExpiry='P';ProductId='Q';PlatformStatus='R';CurrentSpec='S'}
foreach ($entry in $expected.GetEnumerator()) {
    if ($schema.($entry.Key).Column -ne $entry.Value) { $errors.Add("列映射异常：$($entry.Key) 应为 $($entry.Value)，实际为 $($schema.($entry.Key).Column)") }
}
if ($errors.Count -gt 0) { throw ($errors -join '；') }
[ordered]@{ok=$true;checked_scripts=$activeScripts.Count;schema_assertions=$expected.Count;hardcoded_kw2054_assertion=$false;file_invocation_preserves_psscriptroot=$true;pending_profile_reused=$true} | ConvertTo-Json -Compress
