$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$updaterPath = Join-Path $PSScriptRoot 'ensure_replenishment_product_name.ps1'
$migrationPath = Join-Path $PSScriptRoot 'migrate_product_identity_and_batches.ps1'
$initializerPath = Join-Path $PSScriptRoot 'initialize_dynamic_costs.ps1'
$updater = [IO.File]::ReadAllText($updaterPath,[Text.Encoding]::UTF8)
$migration = [IO.File]::ReadAllText($migrationPath,[Text.Encoding]::UTF8)
$initializer = [IO.File]::ReadAllText($initializerPath,[Text.Encoding]::UTF8)

$contracts = [ordered]@{
    updater_uses_stable_identity_then_kw = ($updater -match "MATCH\(O\$\{RowNumber\},'SKU成本主表'!.*Q.*3000" -and $updater -match "VLOOKUP\(B\$\{RowNumber\},'SKU成本主表'!.*A.*D.*3000")
    updater_prefers_custom_title = ($updater -match "SKU成本主表'!.*D.*3000" -and $updater -match "SKU成本主表'!.*C.*3000")
    updater_is_atomic_and_verified = ($updater -match 'New-FinanceTransaction' -and $updater -match 'Register-FinanceRangeBeforeImage' -and $updater -match 'Invoke-FinanceTransactionRollback' -and $updater -match 'formula-verify')
    updater_only_appends_after_existing_columns = ($updater -match '--position U --count 1 --inherit-style before' -and $updater -notmatch 'dim-move' -and $updater -match "range 'U1:U1'")
    migration_creates_product_name_column = ($migration -match "'补货记录' 'U1:U1'" -and $migration -match 'New-ReplenishmentProductNameFormula')
    initializer_creates_product_name_column = ($initializer -match "'补货记录' 'U1:U1'" -and $initializer -match 'New-ReplenishmentProductNameFormula')
}
$failed = @($contracts.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object { $_.Key })
if ($failed.Count -gt 0) { throw "Missing replenishment product-name contracts: $($failed -join ', ')" }

[ordered]@{ok=$true;contracts=$contracts;external_writes=0} | ConvertTo-Json -Depth 5
