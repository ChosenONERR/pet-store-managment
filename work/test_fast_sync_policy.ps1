$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$syncPath = Join-Path $projectRoot 'agent\run_sync_and_notify.ps1'
$orderSyncPath = Join-Path $projectRoot 'work\sync_meituan_orders.ps1'

$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($syncPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count -gt 0) { throw "run_sync_and_notify.ps1 parse failed: $(@($errors).Message -join '; ')" }
$functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-UseNoChangeFastPath' }, $true))
if ($functionAst.Count -ne 1) { throw 'Expected exactly one Test-UseNoChangeFastPath function' }
Invoke-Expression $functionAst[0].Extent.Text

$unchangedOrder = [pscustomobject]@{ok=$true;data_changed=$false}
$changedOrder = [pscustomobject]@{ok=$true;data_changed=$true}
$unchangedGift = [pscustomobject]@{ok=$true;updated_cells=0}
$changedGift = [pscustomobject]@{ok=$true;updated_cells=2}

if (-not (Test-UseNoChangeFastPath 'midday' $unchangedOrder $unchangedGift $false $false)) {
    throw 'midday unchanged data should use the fast path'
}
if (-not (Test-UseNoChangeFastPath 'reminder' $unchangedOrder $unchangedGift $false $false)) {
    throw 'reminder unchanged data should use the fast path'
}
if (Test-UseNoChangeFastPath 'night' $unchangedOrder $unchangedGift $false $false) {
    throw 'night sync must retain the full flow'
}
if (Test-UseNoChangeFastPath 'manual' $unchangedOrder $unchangedGift $false $false) {
    throw 'manual repair sync must retain the full flow'
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

[ordered]@{
    ok = $true
    cases = 15
    covered = @('midday_no_change_skips','reminder_no_change_skips','night_does_not_skip','manual_does_not_skip','order_change_does_not_skip','gift_change_does_not_skip','missing_evidence_does_not_skip','order_missing_ok_does_not_skip','order_ok_false_does_not_skip','gift_missing_ok_does_not_skip','gift_ok_false_does_not_skip','same_data_same_fingerprint','changed_data_different_fingerprint','all_order_domains_compared','writes_guarded_by_change')
} | ConvertTo-Json -Depth 4
