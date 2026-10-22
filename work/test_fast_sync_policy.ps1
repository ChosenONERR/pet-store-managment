$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$syncPath = Join-Path $projectRoot 'agent\run_sync_and_notify.ps1'
$orderSyncPath = Join-Path $projectRoot 'work\sync_meituan_orders.ps1'
$costRebuildPath = Join-Path $projectRoot 'agent\rebuild_cost_snapshots.ps1'
$cliPath = Join-Path $projectRoot 'agent\petstar_cli.py'

$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($syncPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count -gt 0) { throw "run_sync_and_notify.ps1 parse failed: $(@($errors).Message -join '; ')" }
$functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-UseNoChangeFastPath' }, $true))
if ($functionAst.Count -ne 1) { throw 'Expected exactly one Test-UseNoChangeFastPath function' }
Invoke-Expression $functionAst[0].Extent.Text
$incrementalFunctionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-UseIncrementalChangePath' }, $true))
if ($incrementalFunctionAst.Count -ne 1) { throw 'Expected exactly one Test-UseIncrementalChangePath function' }
Invoke-Expression $incrementalFunctionAst[0].Extent.Text

$unchangedOrder = [pscustomobject]@{
    ok=$true;data_changed=$false;change_details=[pscustomobject]@{summary=$false;items=$false;interventions=$false}
    new_item_order_nos=@();changed_order_nos=@();affected_months=@();newly_detected_refunds=0;pending_refund_items=0;applied_refund_details=@()
}
$changedOrder = [pscustomobject]@{
    ok=$true;data_changed=$true
    change_details=[pscustomobject]@{summary=$true;items=$true;interventions=$false}
    new_item_order_nos=@('1001');changed_order_nos=@('1001');affected_months=@('2026-10');newly_detected_refunds=0;pending_refund_items=0;applied_refund_details=@()
}
$unchangedGift = [pscustomobject]@{ok=$true;updated_cells=0}
$changedGift = [pscustomobject]@{ok=$true;updated_cells=2}

if (-not (Test-UseNoChangeFastPath 'midday' $unchangedOrder $unchangedGift $false $false)) {
    throw 'midday unchanged data should use the fast path'
}
if (-not (Test-UseNoChangeFastPath 'reminder' $unchangedOrder $unchangedGift $false $false)) {
    throw 'reminder unchanged data should use the fast path'
}
if (-not (Test-UseNoChangeFastPath 'night' $unchangedOrder $unchangedGift $false $false)) {
    throw 'night sync without changes should use the light audit path'
}
if (-not (Test-UseNoChangeFastPath 'manual' $unchangedOrder $unchangedGift $false $false)) {
    throw 'manual unchanged data should use the fast path unless full rebuild is requested'
}
if (Test-UseNoChangeFastPath 'manual' $unchangedOrder $unchangedGift $false $false $true) {
    throw 'explicit full rebuild must retain the full flow'
}
if (Test-UseNoChangeFastPath 'midday' $changedOrder $unchangedGift $false $false) {
    throw 'changed order data must retain the full flow'
}
if (Test-UseNoChangeFastPath 'midday' $unchangedOrder $changedGift $false $false) {
    throw 'changed gift data must retain the full flow'
}
if (Test-UseNoChangeFastPath 'midday' ([pscustomobject]@{ok=$true}) $unchangedGift $false $false) {
    throw 'missing change evidence must not be treated as unchanged'
}
if (Test-UseNoChangeFastPath 'midday' ([pscustomobject]@{data_changed=$false}) $unchangedGift $false $false) {
    throw 'order result without explicit ok=true must not use the fast path'
}
if (Test-UseNoChangeFastPath 'midday' ([pscustomobject]@{ok=$false;data_changed=$false}) $unchangedGift $false $false) {
    throw 'order result with ok=false must not use the fast path'
}
if (Test-UseNoChangeFastPath 'midday' $unchangedOrder ([pscustomobject]@{updated_cells=0}) $false $false) {
    throw 'gift result without explicit ok=true must not use the fast path'
}
if (Test-UseNoChangeFastPath 'midday' $unchangedOrder ([pscustomobject]@{ok=$false;updated_cells=0}) $false $false) {
    throw 'gift result with ok=false must not use the fast path'
}

if (-not (Test-UseIncrementalChangePath 'midday' $changedOrder $unchangedGift $false $false)) {
    throw 'ordinary new orders should use the incremental path'
}
if (-not (Test-UseIncrementalChangePath 'manual' $changedOrder $unchangedGift $false $false)) {
    throw 'manual order recovery should use the incremental path'
}
if (-not (Test-UseIncrementalChangePath 'night' $changedOrder $unchangedGift $false $false)) {
    throw 'night changes should stay scoped to affected orders and months'
}
if (Test-UseIncrementalChangePath 'manual' $changedOrder $unchangedGift $false $false $true) {
    throw 'explicit full rebuild must not use the incremental path'
}
$refundOrder = $changedOrder.PSObject.Copy(); $refundOrder.newly_detected_refunds = 1
if (-not (Test-UseIncrementalChangePath 'midday' $refundOrder $unchangedGift $false $false)) {
    throw 'refund changes should use the targeted path'
}
$interventionOrder = $changedOrder.PSObject.Copy(); $interventionOrder.change_details = [pscustomobject]@{summary=$true;items=$true;interventions=$true}
if (-not (Test-UseIncrementalChangePath 'midday' $interventionOrder $unchangedGift $false $false)) {
    throw 'intervention changes should use the targeted path'
}
if (-not (Test-UseIncrementalChangePath 'midday' $changedOrder $changedGift $false $false)) {
    throw 'gift changes should use the targeted path'
}
$existingOrderChange = $changedOrder.PSObject.Copy(); $existingOrderChange.new_item_order_nos = @()
if (-not (Test-UseIncrementalChangePath 'midday' $existingOrderChange $unchangedGift $false $false)) {
    throw 'existing-order changes should use the targeted path'
}

$orderTokens = $null
$orderErrors = $null
$orderAst = [Management.Automation.Language.Parser]::ParseFile($orderSyncPath, [ref]$orderTokens, [ref]$orderErrors)
if (@($orderErrors).Count -gt 0) { throw "sync_meituan_orders.ps1 parse failed: $(@($orderErrors).Message -join '; ')" }
$requiredFunctions = @('Convert-ToExcelSerial','Convert-DateDisplayToSerial','Convert-ToComparableText','Convert-ToComparableNumber','Convert-ToComparableDate','Get-RecordsFingerprint','New-SummaryBusinessRecord')
foreach ($functionName in $requiredFunctions) {
    $matches = @($orderAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true))
    if ($matches.Count -ne 1) { throw "Expected exactly one $functionName function" }
    Invoke-Expression $matches[0].Extent.Text
}

$baseSummary = [pscustomobject]@{
    OrderDate='2026-10-11';Platform='Meituan';OrderNo='1001';Status='completed';ProductSales=20;CustomerPay=18
    ExpectedIncome=16.5;PlatformFees=1.5;DeliveryFee=3;ItemCount=1;Source='order page';Note='';DailyOrderNo='Meituan 1'
}
$sameSummary = $baseSummary.PSObject.Copy()
$changedSummary = $baseSummary.PSObject.Copy()
$changedSummary.ExpectedIncome = 15.5
$baseFingerprint = Get-RecordsFingerprint @((New-SummaryBusinessRecord $baseSummary))
$sameFingerprint = Get-RecordsFingerprint @((New-SummaryBusinessRecord $sameSummary))
$changedFingerprint = Get-RecordsFingerprint @((New-SummaryBusinessRecord $changedSummary))
if ($baseFingerprint -ne $sameFingerprint) { throw 'identical business data must have the same fingerprint' }
if ($baseFingerprint -eq $changedFingerprint) { throw 'changed business data must have a different fingerprint' }

$orderSyncSource = [IO.File]::ReadAllText($orderSyncPath, [Text.Encoding]::UTF8)
if ($orderSyncSource -notmatch '\$dataChanged\s*=.+\$summaryChanged.+\$itemsChanged.+\$interventionsChanged') {
    throw 'order data_changed must include summary, item and intervention changes'
}
if ($orderSyncSource -notmatch 'if \(\$dataChanged\)\s*\{[\s\S]+Set-CellMatrix[\s\S]+Set-CellMatrix[\s\S]+Set-CellMatrix') {
    throw 'order writes must be guarded by data_changed'
}

$costRebuildSource = [IO.File]::ReadAllText($costRebuildPath, [Text.Encoding]::UTF8)
$hasIncrementalSwitch = [bool]($costRebuildSource -match '\[switch\]\$Incremental')
$hasSparseItemWrite = [bool]($costRebuildSource -match 'incremental_item_rows' -and $costRebuildSource -match '\$itemCostUpdate')
if (-not ($hasIncrementalSwitch -and $hasSparseItemWrite)) {
    throw "incremental cost rebuild must write only targeted order-item rows (switch=$hasIncrementalSwitch sparse=$hasSparseItemWrite)"
}
$hasFullItemRepair = [bool]($costRebuildSource -match 'Clear-Range' -and $costRebuildSource -match 'L2:N4001')
$hasFullSummaryRepair = [bool]($costRebuildSource -match 'K2:K2001' -and $costRebuildSource -match 'L2:L2001')
if (-not ($hasFullItemRepair -and $hasFullSummaryRepair)) {
    throw "full cost rebuild must retain whole-table repair behavior (item=$hasFullItemRepair summary=$hasFullSummaryRepair length=$($costRebuildSource.Length) path=$costRebuildPath)"
}
$cliSource = [IO.File]::ReadAllText($cliPath, [Text.Encoding]::UTF8)
if ($cliSource -notmatch '--full-rebuild' -or $cliSource -notmatch '\("FullRebuild", args\.full_rebuild, True\)') {
    throw 'CLI must expose an explicit full rebuild escape hatch'
}

[ordered]@{
    ok = $true
    cases = 27
    covered = @('midday_no_change_skips','reminder_no_change_skips','night_no_change_uses_light_audit','manual_no_change_skips','explicit_full_does_not_skip','order_change_is_targeted','gift_change_is_targeted','missing_evidence_does_not_skip','order_missing_ok_does_not_skip','order_ok_false_does_not_skip','gift_missing_ok_does_not_skip','gift_ok_false_does_not_skip','new_order_uses_incremental','manual_recovery_uses_incremental','night_change_is_targeted','explicit_full_retains_full','refund_is_targeted','intervention_is_targeted','gift_is_targeted','existing_order_change_is_targeted','same_data_same_fingerprint','changed_data_different_fingerprint','all_order_domains_compared','writes_guarded_by_change','incremental_cost_writes_are_targeted','full_cost_repair_is_preserved','cli_full_rebuild_escape_hatch')
} | ConvertTo-Json -Depth 4
