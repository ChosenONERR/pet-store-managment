$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$agentRoot = Join-Path $projectRoot 'agent'
$bridgePath = Join-Path $agentRoot 'lark_event_bridge.ps1'
$watchdogPath = Join-Path $agentRoot 'ensure_lark_event_bridge.ps1'
$wrapperPath = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$hiddenWatchdogPath = Join-Path $agentRoot 'run_hidden_watchdog.vbs'
$petstarCliPath = Join-Path $agentRoot 'petstar_cli.py'
$financeEventPath = Join-Path $agentRoot 'record_finance_event.ps1'
$stateRoot = Join-Path $agentRoot 'state'
$runtimePowerShellFiles = @()
$runtimePowerShellFiles += @(Get-ChildItem -LiteralPath $agentRoot -Filter '*.ps1' -File)
$runtimePowerShellFiles += Get-Item -LiteralPath (Join-Path $projectRoot 'work\sync_meituan_orders.ps1')
$runtimePowerShellFiles += Get-Item -LiteralPath (Join-Path $projectRoot 'work\sync_meituan_inventory.ps1')
$runtimePowerShellFiles += Get-Item -LiteralPath (Join-Path $projectRoot 'work\reconcile_manual_delivery_fees.ps1')
$runtimePowerShellFiles += Get-Item -LiteralPath (Join-Path $projectRoot 'work\repair_dashboard_offline_formulas.ps1')
$runtimePowerShellFiles += Get-Item -LiteralPath (Join-Path $projectRoot 'work\repair_dashboard_and_replenishment_payments_20261025.ps1')
$runtimePowerShellFiles += @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'work') -Filter 'test_*.ps1' -File)

$parseFailures = [Collections.Generic.List[string]]::new()
foreach ($file in $runtimePowerShellFiles) {
    $errors = $null
    [Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$null,[ref]$errors) | Out-Null
    foreach ($parseError in @($errors)) { $parseFailures.Add("$($file.Name): $($parseError.Message)") }
    $bytes = [IO.File]::ReadAllBytes($file.FullName)
    $text = [IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191
    if ($text -match '[^\x00-\x7F]' -and -not $hasBom) { $parseFailures.Add("$($file.Name): non-ASCII Windows PowerShell entry must have UTF-8 BOM") }
}
if ($parseFailures.Count -gt 0) { throw "Runtime PowerShell parse failures: $($parseFailures -join '; ')" }

$bridge = [IO.File]::ReadAllText($bridgePath,[Text.Encoding]::UTF8)
$watchdog = [IO.File]::ReadAllText($watchdogPath,[Text.Encoding]::UTF8)
$wrapper = [IO.File]::ReadAllText($wrapperPath,[Text.Encoding]::UTF8)
$hiddenWatchdog = [IO.File]::ReadAllText($hiddenWatchdogPath,[Text.Encoding]::UTF8)
$hiddenScheduledJob = [IO.File]::ReadAllText((Join-Path $agentRoot 'run_hidden_scheduled_job.vbs'),[Text.Encoding]::UTF8)
$scheduledJob = [IO.File]::ReadAllText((Join-Path $agentRoot 'run_scheduled_job.ps1'),[Text.Encoding]::UTF8)
$desktopLarkAuth = [IO.File]::ReadAllText((Join-Path $agentRoot 'authorize_lark_desktop.ps1'),[Text.Encoding]::UTF8)
$scheduledTaskInstaller = [IO.File]::ReadAllText((Join-Path $agentRoot 'configure_petstar_scheduled_tasks.ps1'),[Text.Encoding]::UTF8)
$weeklyMaintenance = [IO.File]::ReadAllText((Join-Path $agentRoot 'run_weekly_maintenance.ps1'),[Text.Encoding]::UTF8)
$petstarCli = [IO.File]::ReadAllText($petstarCliPath,[Text.Encoding]::UTF8)
$financeEvent = [IO.File]::ReadAllText($financeEventPath,[Text.Encoding]::UTF8)
$replenishmentBlock = [regex]::Match($financeEvent, "(?s)'Replenishment'\s*\{(.*?)\n\s*'SetSkuCost'\s*\{").Groups[1].Value
$setSkuCostBlock = [regex]::Match($financeEvent, "(?s)'SetSkuCost'\s*\{(.*?)\n\s*'SetSkuExpiry'\s*\{").Groups[1].Value
$skuSync = [IO.File]::ReadAllText((Join-Path $agentRoot 'sync_new_skus.ps1'),[Text.Encoding]::UTF8)
$skuIdentitySync = [IO.File]::ReadAllText((Join-Path $agentRoot 'sync_sku_master_by_identity.ps1'),[Text.Encoding]::UTF8)
$financeTransaction = [IO.File]::ReadAllText((Join-Path $agentRoot 'finance_transaction.ps1'),[Text.Encoding]::UTF8)
$scheduledSync = [IO.File]::ReadAllText((Join-Path $agentRoot 'run_sync_and_notify.ps1'),[Text.Encoding]::UTF8)
$costRebuild = [IO.File]::ReadAllText((Join-Path $agentRoot 'rebuild_cost_snapshots.ps1'),[Text.Encoding]::UTF8)
$ledgerSync = [IO.File]::ReadAllText((Join-Path $agentRoot 'sync_daily_ledger.ps1'),[Text.Encoding]::UTF8)
$codexBridge = [IO.File]::ReadAllText((Join-Path $agentRoot 'codex_chat_bridge.py'),[Text.Encoding]::UTF8)
$businessQuery = [IO.File]::ReadAllText((Join-Path $agentRoot 'query_business_data.ps1'),[Text.Encoding]::UTF8)
$runtimePaths = [IO.File]::ReadAllText((Join-Path $agentRoot 'runtime_paths.ps1'),[Text.Encoding]::UTF8)
$browserPreflight = [IO.File]::ReadAllText((Join-Path $agentRoot 'check_meituan_browser.mjs'),[Text.Encoding]::UTF8)
$deliveryReconciliation = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\reconcile_manual_delivery_fees.ps1'),[Text.Encoding]::UTF8)
$dashboardFormulaRepair = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\repair_dashboard_offline_formulas.ps1'),[Text.Encoding]::UTF8)
$dashboardPaymentRepair = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\repair_dashboard_and_replenishment_payments_20261025.ps1'),[Text.Encoding]::UTF8)
$simpleFinanceBuilder = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\create_simple_finance.ps1'),[Text.Encoding]::UTF8)
$orderExtractor = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\meituan_order_extract.mjs'),[Text.Encoding]::UTF8)
$orderSync = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\sync_meituan_orders.ps1'),[Text.Encoding]::UTF8)
$refundAdjustments = [IO.File]::ReadAllText((Join-Path $agentRoot 'refund_adjustments.ps1'),[Text.Encoding]::UTF8)
$orderActualOverrides = [IO.File]::ReadAllText((Join-Path $agentRoot 'order_actual_overrides.ps1'),[Text.Encoding]::UTF8)
$inventoryExporter = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\meituan_inventory_export.mjs'),[Text.Encoding]::UTF8)
$productExportFieldInspectorPath = Join-Path $projectRoot 'work\inspect_meituan_product_export_fields.mjs'
$productExportFieldInspector = [IO.File]::ReadAllText($productExportFieldInspectorPath,[Text.Encoding]::UTF8)
$meituanBrowserContext = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\meituan_browser_context.mjs'),[Text.Encoding]::UTF8)
$inventorySync = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\sync_meituan_inventory.ps1'),[Text.Encoding]::UTF8)
$interventionRecord = [IO.File]::ReadAllText((Join-Path $agentRoot 'intervention_record.ps1'),[Text.Encoding]::UTF8)
    $promotionSync = [IO.File]::ReadAllText((Join-Path $agentRoot 'sync_promotion_expense.ps1'),[Text.Encoding]::UTF8)
    $promotionExtractor = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\meituan_promotion_extract.mjs'),[Text.Encoding]::UTF8)
    $integrityAudit = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\audit_finance_integrity.ps1'),[Text.Encoding]::UTF8)

$contracts = [ordered]@{
    infinite_event_stream = ($bridge -match 'event consume.+--timeout 0.+--quiet' -and $bridge -notmatch '--timeout 10m')
    durable_ack_outbox = ($bridge -match 'pending_acknowledgements\.json' -and $bridge -match 'function Flush-PendingAcknowledgements')
    confirmation_atomic_transaction = ($bridge -match 'New-FinanceTransaction' -and $bridge -match '飞书确认' -and $bridge -match 'Invoke-FinanceTransactionRollback' -and $bridge -match '本条消息产生的改动已全部回滚')
    state_before_success_ack = ($bridge.IndexOf('Complete-PendingConfirmation ([string]$pending.change_id)') -lt $bridge.IndexOf("Send-Acknowledgement `$AuthorizedOpenId `$successReply"))
    shared_pending_mutex = ($bridge -match 'Local\\PetstarLarkPendingConfirmations')
    retryable_failures = ($bridge.Contains("Add-HandledMessageRecord `$messageId 'failed_retryable'"))
    watchdog_replay = ($watchdog -match 'function Start-ReplayWorker' -and $watchdog -match "-Mode','Replay")
    watchdog_does_not_count_pending_retry_as_live_bridge = ($watchdog -match 'FlushAcks\|RetryPending')
    watchdog_repairs_orphan_consumers = ($watchdog -match 'function Repair-OrphanEventConsumers' -and $watchdog -match 'consumer_pids')
    replenishment_remote_idempotency = ($financeEvent -match "eventBatchId" -and $financeEvent -match "T\$\{rowNumber\}:T\$\{rowNumber\}" -and $financeEvent -match 'existingEventIndex' -and $financeEvent -match 'existingAfterQuantity')
    replenishment_duplicate_sku_guard = ($financeEvent -match 'Get-SkuCodeCandidates' -and $financeEvent -match '已停止补货写入，避免新增重复档案')
    batch_sku_json_array_is_flattened_for_windows_powershell = ($financeEvent -match 'Windows PowerShell 5\.1 preserves a top-level JSON array' -and $financeEvent -match '\$item -is \[Array\]' -and $financeEvent -match 'Write-Output -NoEnumerate \$operation' -and $financeEvent -match '\$beforeInventoryValue = if' -and $financeEvent -match '\$cell -is \[Collections\.IDictionary\]')
    replenishment_duplicate_code_uses_only_unique_active_platform_identity = ($financeEvent -match '\$activeMatches\.Count -ne 1' -and $financeEvent -match "\$candidateStatus -eq '售卖中'" -and $financeEvent -match 'matched_unique_active_platform_sku' -and $bridge -match '未删除或修改其他同号行')
    sku_sync_prewrite_duplicate_guard = ($skuSync -match 'duplicateNewKw' -and $skuSync -match '已在写入前停止')
    recovery_cooldown = ($watchdog -match 'Test-NightRecoveryCooldown' -and $watchdog -match 'night_recovery_attempt\.json')
    noexit_cli_wrapper = ($wrapper -notmatch '(?m)^\s*exit\s+' -and $wrapper -match 'node')
    powershell_file_entry_preserves_script_root = ($petstarCli -match '"-File"' -and $petstarCli -notmatch 'ScriptBlock|EncodedCommand')
    hidden_watchdog_launcher = ($hiddenWatchdog -match 'shell\.Run\(command, 0, True\)' -and $hiddenWatchdog -match 'WScript\.Quit exitCode')
    hidden_scheduled_jobs_report_real_exit_code = ($hiddenScheduledJob -match 'shell\.Run\(command, 0, True\)' -and $hiddenScheduledJob -match 'WScript\.Quit exitCode' -and $scheduledTaskInstaller -match 'New-TimeSpan -Minutes 90')
    scheduled_jobs_use_local_lark_credentials = (
        $scheduledJob -match "'HERMES_HOME'.+'HERMES_GATEWAY_TOKEN'.+'HERMES_SESSION_KEY'" -and
        $scheduledJob -match "'OPENCLAW_CLI'.+'OPENCLAW_HOME'" -and
        $scheduledJob -match "'LARK_CHANNEL'" -and
        $scheduledJob -match 'Remove-Item -LiteralPath "Env:\$variableName"' -and
        $scheduledJob.Contains("`$env:LARKSUITE_CLI_CONFIG_DIR = Join-Path `$env:USERPROFILE '.lark-cli'")
    )
    desktop_lark_auth_uses_local_workspace_without_persisting_tokens = (
        $desktopLarkAuth -match 'auth login --recommend --no-wait --json' -and
        $desktopLarkAuth -match 'auth login --device-code \$deviceCode --json' -and
        $desktopLarkAuth -match 'LARKSUITE_CLI_CONFIG_DIR' -and
        $desktopLarkAuth -notmatch 'accessToken|refreshToken'
    )
    one_global_mutation_mutex = ($financeTransaction -match 'Local\\PetstarFinanceMutationV2')
    durable_before_images = ($financeTransaction -match 'Register-FinanceRangeBeforeImage' -and $financeTransaction -match 'status = ''rollback_pending''')
    processed_event_after_commit = ($financeEvent -match 'Register-FinanceProcessedEvent' -and $financeEvent -notmatch 'Register-FinanceProcessedEvent[^\r\n]+\r?\n\s*\[ordered\].+Add-Content')
    cost_and_ledger_join_transaction = ($costRebuild -match 'Join-FinanceTransaction' -and $ledgerSync -match 'Join-FinanceTransaction')
    scheduled_sync_runs_integrity_audit = ($scheduledSync -match "Invoke-WithRetry 'integrity_audit' \{ Invoke-FinanceIntegrityAudit")
    scheduled_sync_commits_before_notification = ($scheduledSync.LastIndexOf("`n    Complete-SyncFinanceTransaction") -gt 0 -and $scheduledSync.LastIndexOf("`n    Complete-SyncFinanceTransaction") -lt $scheduledSync.LastIndexOf('Send-FeishuMessage $message'))
    committed_data_not_reported_as_sync_failure = ($scheduledSync -match '账目状态：已校验并提交，无需重复录入')
    codex_has_read_only_business_query = ($codexBridge -match 'run_business_queries' -and $codexBridge -match '不能说“看不到表格”')
    business_query_has_no_write_command = ($businessQuery -match '\+cells-get' -and $businessQuery -notmatch '\+cells-set|\+cells-clear|record_finance_event|sync_meituan')
    pending_cost_actions_use_stable_sku_id = ($businessQuery.Contains('"sku_id:$productId"') -and $financeEvent.Contains("StartsWith('sku_id:')"))
    dry_run_skips_idempotency_journal_scan = ($financeEvent -match 'if \(-not \$DryRun -and \$EventId -and \(Test-ProcessedEvent \$EventId\)\)')
    unsupported_dry_run_stops_before_write = ($financeEvent -match 'nativeDryRunActions' -and $financeEvent -match '已在任何写入前停止')
    committed_event_recovery_is_direct_lookup = ($financeTransaction -match 'Get-FinanceTransactionPath \$StateRoot \("event-\$EventId"\)' -and $financeTransaction -notmatch "function Test-FinanceCommittedEvent[\s\S]{0,900}Get-ChildItem")
    finalized_transactions_drop_cell_payloads = ($financeTransaction -match 'ConvertTo-CompactedFinanceTransactionJournal' -and $financeTransaction -match "property\.Name -in @\('entries'" -and $financeTransaction -match "entry_manifest")
    failed_journal_writes_clean_temp_files = ($financeTransaction -match 'function Write-FinanceJournalAtomically' -and $financeTransaction -match 'finally[\s\S]{0,300}Remove-Item -LiteralPath \$temp')
    rollback_is_chunked_and_retried = ($financeTransaction -match 'Get-FinanceRollbackChunks' -and $financeTransaction -match 'attempt -le 3')
    rollback_operations_remain_json_array = ($financeTransaction -match 'ConvertTo-Json -InputObject \$operations')
    rollback_explicitly_clears_new_tail_rows = ($financeTransaction -match '\+cells-batch-clear' -and $financeTransaction -match '先显式清空内容')
    order_sync_clears_stale_tail_rows = ($orderSync -match 'existingItemsLastRow' -and $orderSync -match 'Clear-CellRange.+订单商品' -and $orderSync -match 'existingSummaryLastRow')
    integrity_audit_checks_order_item_quantity = ($integrityAudit -match 'order_item_quantity_mismatches' -and $integrityAudit -match 'itemQuantityByOrder')
    rollback_restores_numeric_storage_type = ($financeTransaction -match 'Writing those strings back during rollback' -and $financeTransaction -match '\[double\]::TryParse' -and $financeTransaction -match '\$numberFormat -notin @\(''General'',''@''\)')
    singleton_pending_cost_is_array_safe = ($ledgerSync -match '@\(\$missingCodes\)\.Count')
    duplicate_before_images_are_skipped = ($financeTransaction -match 'alreadyCaptured' -and $financeTransaction -match 'if \(\$alreadyCaptured\) \{ return \}')
    preflight_checks_state_drive = ($scheduledSync -match 'stateDriveRoot' -and $scheduledSync -match 'AvailableFreeSpace -lt 1GB')
    scheduled_sync_refreshes_auth_before_guarded_sku_schema_read = (
        $scheduledSync -match '\$skuSchema = \$null' -and
        $scheduledSync.Contains("Invoke-WithRetry 'sku_schema' { Get-SkuMasterSchema -SpreadsheetToken `$spreadsheetToken } 4") -and
        $scheduledSync.IndexOf("Invoke-WithRetry 'preflight'") -lt $scheduledSync.IndexOf("Invoke-WithRetry 'sku_schema'") -and
        $scheduledSync.IndexOf("Invoke-WithRetry 'sku_schema'") -gt $scheduledSync.IndexOf("try {", $scheduledSync.IndexOf('$operationalDegraded = $false'))
    )
    scheduled_sync_diagnoses_windows_lark_credential_store = ($scheduledSync -match 'HKEY_CURRENT_USER\\Software\\LarkCli\\keychain\\lark-cli' -and $scheduledSync -match 'lark_registry_value_count')
    pending_repair_skips_full_committed_journals = ($financeTransaction -match 'Get-FinanceTransactionStatusFast' -and $financeTransaction.Contains("if (`$status -notin @('active','rollback_pending')) { continue }"))
    order_cards_expand_before_extraction = ($orderExtractor -match 'await expandOrderCards\(targetId\)[\s\S]{0,120}extractPageExpression')
    order_retries_restart_from_first_page = ($orderExtractor -match 'function resetOrderPagination' -and $orderExtractor -match 'await resetOrderPagination\(targetId\)' -and $orderExtractor -match "querySelector\('\.roo-pagination'\)")
    lark_1_0_98_formula_verify_uses_continuous_ranges = ($orderSync -match "Range='E2:L2001'" -and $orderSync -match "Range='E2:P4001'" -and $inventorySync -match "--range 'H2:M2000'" -and $orderSync -notmatch "Range='[^']*,[^']*'" -and $scheduledSync -match "\$Range -split ','" -and $scheduledSync -match '--range \$continuousRange' -and $costRebuild -match "\$Range -split ','" -and $costRebuild -match '--range \$area' -and $ledgerSync -match "\$Range -split ','" -and $ledgerSync -match '--range \$area' -and $financeEvent -match "\$Range -split ','" -and $financeEvent -match '--range \$area')
    cost_rebuild_preserves_error_detail = ($scheduledSync -match '成本快照刷新失败：\$detail' -and $scheduledSync -match '\$costScript @costArguments 2>&1')
    order_page_reloads_once_before_failure = ($orderExtractor -match 'function waitForOrderPageReady' -and $orderExtractor -match '/navigate\?target=' -and $orderExtractor -match 'not ready after one safe reload')
    meituan_login_is_detected_without_slow_retry = ($orderExtractor -match 'function waitForOrderReadyOrAuth' -and $orderExtractor -match 'MEITUAN_AUTH_REQUIRED' -and $orderExtractor -match 'login_gw' -and $scheduledSync -match '\$nonRetryable[\s\S]{0,500}MEITUAN_AUTH_REQUIRED')
    meituan_login_alert_is_persistent_and_self_resolving = ($scheduledSync -match "return 'meituan-auth'" -and $scheduledSync -match 'active-meituan-auth\.json|active-\$safeCategory\.json' -and $scheduledSync -match "Resolve-FailureNotification 'meituan-auth'" -and $scheduledSync -match 'authenticated_order_read_succeeded')
    order_card_completeness_uses_quantity_sum = ($orderExtractor -match "reduce\(\(total, row\)" -and $orderExtractor -match 'expectedItemCount')
    empty_order_days_are_valid = ($orderExtractor -match 'if \(/共\\\\s\*0\\\\s\*单/\.test\(bodyText\)\) return true')
    empty_order_days_do_not_use_array_property_enumeration = ($orderSync -match '\$extractedOrderNos = @\(\$extract\.orders \| ForEach-Object' -and $orderSync -notmatch '\$extract\.orders\.orderNo')
    inventory_export_waits_for_loaded_fields = ($inventoryExporter -match 'required inventory export fields to finish loading' -and $inventoryExporter -match '!/正在加载/' -and $inventoryExporter -match 'required\.every')
    inventory_failure_returns_structured_detail = ($inventorySync -match 'Console\]::Out\.WriteLine' -and $inventorySync -match 'ok=\$false;error=\$reason' -and $scheduledSync -match 'parsedError\.error')
    business_warnings_do_not_mark_runtime_degraded = ($scheduledSync -match 'degraded = \[bool\]\$operationalDegraded' -and $scheduledSync -match 'sync_degraded' -and $scheduledSync -match '已完成·需核对')
    incomplete_orders_validate_quantity_sum = ($orderSync -match 'Measure-Object -Property quantity -Sum' -and $orderSync -match 'expectedItemCount')
    incomplete_orders_fail_before_sheet_read = ($orderSync.IndexOf('$incompleteOrders') -lt $orderSync.IndexOf("Get-SheetRows '订单汇总'"))
    system_powershell_resolver_avoids_pshome = ($runtimePaths -match 'System32\\WindowsPowerShell\\v1\.0\\powershell\.exe' -and $runtimePaths -notmatch '\$PSHOME')
    production_scripts_do_not_assume_pshome_powershell = (-not (($runtimePowerShellFiles | Where-Object { $_.Name -notlike 'test_*' } | ForEach-Object { [IO.File]::ReadAllText($_.FullName,[Text.Encoding]::UTF8) }) -match '\$PSHOME'))
    child_script_failure_preserves_detail = ($scheduledSync -match '2>&1' -and $scheduledSync -match '子脚本未返回错误详情' -and $orderSync -match 'throw \$originalError' -and $orderSync -notmatch '(?m)^\s*trap\s*\{')
    codex_query_uses_system_powershell = ($codexBridge -match 'WINDOWS_POWERSHELL' -and $codexBridge -notmatch 'command = \[\s*"powershell\.exe"')
    natural_language_enums_are_normalized_before_cli = ($bridge -match 'function Convert-PetstarReceiptStatus' -and $bridge -match "'当天到账'.+return '已到账'" -and $bridge -match 'function Convert-PetstarPlatform' -and $bridge -match "'京东外卖'.+return '京东到家'")
    withdrawn_interventions_are_auditable_and_excluded = ($financeEvent -match 'audit_tombstone' -and $financeEvent -match 'Get-InterventionTombstoneOrderNo' -and $interventionRecord -match 'Test-ActiveInterventionOrderNo' -and $businessQuery -match 'Test-ActiveInterventionOrderNo' -and $ledgerSync -match 'Test-ActiveInterventionOrderNo' -and $costRebuild -match 'Test-ActiveInterventionOrderNo' -and $orderSync -match 'Test-ActiveInterventionOrderNo')
    intervention_withdrawal_single_row_readback = ($financeEvent -match 'function Get-SingleSheetRow' -and $financeEvent -match "Get-SingleSheetRow '干预订单'")
    promotion_capture_supports_safe_next_day_recovery = ($promotionSync -match 'Today\.AddDays\(-1\)' -and $promotionSync -notmatch 'Now\.Hour -lt 3' -and $promotionExtractor -match 'yesterday.*settled spend' -and $promotionExtractor -match '/ad/v3/flow/r/history')
    promotion_capture_uses_network_response_bridge = ($promotionExtractor -match 'captureJsonResource' -and $promotionExtractor -match '/capture\?url=' -and $promotionExtractor -match 'shared CDP proxy' -and $promotionExtractor -match '/ad/v4/homepage/account/info')
    promotion_capture_uses_authenticated_merchant_route = ($promotionExtractor -match 'isAuthenticatedPromotionPortalUrl' -and $promotionExtractor -match 'shangoue\.meituan\.com')
    promotion_failure_notification_is_consolidated = ($promotionSync -match "'deferred_to_night_sync'" -and $promotionSync -match "'retry_after_settlement'" -and $promotionSync -notmatch 'Send-FailureNotification')
    promotion_failure_does_not_block_late_order_sync = ($scheduledSync -match '\$promotionCaptureError = \$null' -and $scheduledSync -match '推广费读取失败不能阻断晚到订单补抓' -and $scheduledSync -match 'order_sync_continued=\$true' -and $scheduledSync.IndexOf("\$promotionCapture = Get-PromotionCaptureState") -lt $scheduledSync.IndexOf("Invoke-WithRetry 'order_sync'"))
    promotion_failure_marks_night_sync_degraded = ($scheduledSync -match '\$operationalDegraded = \$true' -and $scheduledSync -match '本次赚/亏未扣除待补推广费')
    promotion_cdp_transport_failure_is_fast = ($promotionExtractor -notmatch 'new WebSocket' -and $promotionExtractor -match 'AbortSignal\.timeout' -and $promotionSync -match 'CDP proxy request failed.+break')
    repeated_failure_notification_is_idempotent = ($scheduledSync -match 'Get-FailureNotificationCategory' -and $scheduledSync -match 'failure_notifications' -and $scheduledSync -match 'notification_suppressed' -and $scheduledSync -notmatch 'petstar-failed-.+HHmm')
    acknowledgement_json_quotes_are_transport_safe = ($bridge -match 'transportText = \$Text\.Replace' -and $bridge -match '--text \$transportText')
    oversized_sheet_reads_are_split = ($scheduledSync -match 'data\.has_more' -and $scheduledSync -match '\$leftRange' -and $scheduledSync -match '\$rightRange' -and $scheduledSync -match '飞书单行读取仍不完整')
    summary_failure_cannot_rollback_valid_ledger = ($scheduledSync.LastIndexOf("`n    Complete-SyncFinanceTransaction") -gt $scheduledSync.LastIndexOf("Invoke-WithRetry 'integrity_audit'") -and $scheduledSync.LastIndexOf("`n    Complete-SyncFinanceTransaction") -lt $scheduledSync.LastIndexOf("Invoke-WithRetry 'summary_readback' { Build-Snapshot }"))
    completed_refunds_are_keyed_by_record_id = ($orderExtractor -match 'recordId' -and $refundAdjustments -match 'meituan:\$\(\$order\.orderNo\):\$\(\$refund\.recordId\)')
    platform_refunds_preserve_pending_and_unknown_states = ($orderExtractor -match 'classifyRefundGroup' -and $orderExtractor -match "state !== 'closed'" -and $orderSync -match 'refund_observation_details' -and $scheduledSync -match '退款状态观察' -and $scheduledSync -match 'refund_observation_alert')
    meituan_return_goods_refunds_use_new_structure = ($orderExtractor -match 'refundGoodsInfoList' -and $orderExtractor -match '全部退货\|整单退货' -and $orderExtractor -match 'chargeInfo\?\.userPayTotalAmount')
    refund_income_preserves_cents = ($orderSync.Contains('[math]::Max([double]0.0, [double][math]::Round'))
    zero_income_refunds_are_labeled_full_refund = ($orderSync -match 'effectiveExpectedIncome.+-le 0\.004999' -and $orderSync -match '已完成（全额退款）' -and $orderSync -match '已完成（部分退款）')
    returned_inventory_reverses_only_returned_cost = ($costRebuild -match 'returned_restocked' -and $costRebuild -match 'effective_quantity')
    optional_sheet_cells_and_order_fields_are_safe = ($orderSync -match "PSObject\.Properties\['value'\]" -and $ledgerSync -match "PSObject\.Properties\['is_offline'\]")
    actual_fulfillment_override_is_durable = ($orderActualOverrides -match 'order_actual_overrides\.json' -and $orderSync -match 'Get-OrderActualOverrides' -and $financeEvent -match "'SetOrderActuals'")
    codex_schema_exposes_order_actuals = (([IO.File]::ReadAllText((Join-Path $agentRoot 'codex_response_schema.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json).properties.actions.items.properties.action.enum -contains 'SetOrderActuals')
    order_refund_cannot_be_recorded_as_actual_receipt = ($financeEvent -match 'ActualReceipt 只用于平台资金实际到账' -and $codexBridge -match '绝不能使用 ActualReceipt')
    actual_fulfillment_override_is_atomic = ($financeEvent -match 'orderActualOverrideStateBackup' -and $financeEvent -match 'Invoke-OrderSyncForDate' -and $financeEvent -match '店主已手工恢复')
    actual_fulfillment_override_does_not_write_platform_inventory = ($codexBridge -match 'SetOrderActuals' -and $codexBridge -match '该动作不修改平台库存' -and $financeEvent -notmatch "SetOrderActuals[\s\S]{0,7000}(?:sync_meituan_inventory|meituan_inventory_export)")
    single_private_addon_uses_income_delta = ($financeEvent -match '\$newRequestedItems\.Count -eq 1 -and \$allOriginalItemsPreserved' -and $financeEvent -match '\$actualIncome - \$oldIncome' -and $financeEvent -match "additional_income_delta" -and $codexBridge -match '原商家预计收入\+额外付款')
    unspecified_actuals_delivery_fee_uses_current_order_value = ($financeEvent -match '\$actualDeliveryFee = if \(\$null -eq \$DeliveryFee\) \{ \$oldDeliveryFee \}' -and $bridge -match "\$actionName -eq 'SetOrderActuals'.+\$entry.Key -eq 'delivery_fee'" -and $codexBridge -match 'delivery_fee 必须填 null')
    table_put_styles_use_stdin = ($ledgerSync -match '--sheets\s+\$sheetsFileArgument\s+--styles\s+-' -and $ledgerSync -match 'daily_ledger_sheets_' -and $ledgerSync -match 'finally[\s\S]{0,180}Remove-Item -LiteralPath \$requestPath')
    table_put_resolves_physical_state_path = ($ledgerSync -match '\$stateDirectory\.LinkType' -and $ledgerSync -match '\$stateDirectory\.Target' -and $ledgerSync -match 'Push-Location -LiteralPath \$requestRoot')
    table_put_preserves_native_exit_code = ($ledgerSync -match '\[pscustomobject\]@\{raw=\$raw;exit_code=\$exitCode\}' -and $ledgerSync -match '\[int\]\$response\.exit_code')
    dropdown_json_uses_stdin_and_relative_file = ($ledgerSync -match 'function Invoke-DropdownSetJson' -and $ledgerSync -match '--options\s+-\s+--colors\s+\$colorsFileArgument' -and $ledgerSync -match 'daily_ledger_dropdown_colors_')
    runtime_has_no_inline_composite_dropdown_json = ($ledgerSync -notmatch '--options\s+\$options(?!\w)' -and $ledgerSync -notmatch '--colors\s+\$colors(?!\w)')
    expense_dropdown_repairs_partial_month_setup = ($ledgerSync -match 'Reapply this idempotently' -and $ledgerSync -match 'Invoke-DropdownSetJson \$expenseName')
    ledger_date_formats_are_reapplied = ($ledgerSync -match 'function Set-DateDisplayFormat' -and $ledgerSync.Contains("Set-DateDisplayFormat `$incomeName 'C2:C1001'") -and $ledgerSync.Contains("Set-DateDisplayFormat `$expenseName 'A2:A1001'"))
    ledger_sheet_reads_preserve_single_row_shape = ($ledgerSync -match '\[pscustomobject\]@\{rows=\[object\[\]\]@\(' -and $ledgerSync -match '\$checkIncome = \(Get-SheetRows[^\r\n]+\)\.rows' -and $ledgerSync -match '\$checkExpense = \(Get-SheetRows[^\r\n]+\)\.rows')
    confirmed_event_sheet_reads_are_retried = ($financeEvent -match 'function Invoke-LarkReadJson' -and $financeEvent -match 'MaxAttempts = 4' -and $financeEvent -match '读取飞书失败：\$SheetName')
    ledger_source_reads_are_retried = ($ledgerSync -match 'function Invoke-LarkReadJson' -and $ledgerSync -match 'MaxAttempts = 4' -and $ledgerSync -match '读取飞书失败：\$SheetName')
    integrity_audit_discovers_month_sheets = ($integrityAudit -match 'Get-LedgerMonth' -and $integrityAudit -match '\+workbook-info' -and $integrityAudit -match '\$incomeSheets' -and $integrityAudit -match '\$expenseSheets' -and $integrityAudit -notmatch "Get-Rows \$ledgerToken '8月收入'")
    integrity_audit_preserves_failure_detail = ($scheduledSync -match '自然单缺失' -and $scheduledSync -match '经营账目一致性审计未通过：\$detail')
    invalid_cli_arguments_do_not_retry = ($scheduledSync -match '\$nonRetryable[\s\S]{0,500}invalid JSON\|invalid_argument')
    codex_uses_persistent_app_server_with_fallback = ($codexBridge -match 'class CodexAppRuntime' -and $codexBridge -match 'app-server.+--stdio' -and $codexBridge -match 'return output\(process_message')
    codex_persistent_runtime_follows_account_switches = ($codexBridge -match 'get_auth_fingerprint' -and $codexBridge -match 'ensure_current_auth' -and $codexBridge -match 'auth_reloaded')
    meituan_automation_binds_verified_store_context = ($meituanBrowserContext -match '宠速星宠物用品店' -and $meituanBrowserContext -match 'authenticatedStore' -and $orderExtractor -match 'openPetstarMerchantTab' -and $inventoryExporter -match 'openPetstarMerchantTab' -and $promotionExtractor -match 'openPetstarMerchantTab')
    meituan_automation_uses_only_bound_chrome_profile = ($meituanBrowserContext -match 'loadBinding' -and $meituanBrowserContext -match 'profile_directory' -and $meituanBrowserContext -match 'resolveBoundContext' -and $meituanBrowserContext -match 'boundPages' -and $meituanBrowserContext -match 'MEITUAN_AUTH_REQUIRED')
    bound_online_chrome_opens_merchant_tab_without_profile_probe = ($meituanBrowserContext -match '/new\?url=' -and $meituanBrowserContext -match 'target\.browserContextId !== resolved\.browserContextId' -and $meituanBrowserContext -match 'if \(!target\) \{[\s\S]{0,500}/new\?url=')
    meituan_inventory_export_supports_explicit_extra_fields = ($inventoryExporter -match 'MEITUAN_EXPORT_EXTRA_FIELDS' -and $inventoryExporter -match '\.\.\.extraFields')
    meituan_product_field_inspector_is_read_only = ((Test-Path -LiteralPath $productExportFieldInspectorPath) -and $productExportFieldInspector -match 'openPetstarMerchantTab' -and $productExportFieldInspector -match 'releasePetstarMerchantTab' -and $productExportFieldInspector -notmatch 'createExport|clickCreateExport|修改库存')
    business_queries_are_parallel_and_bounded = ($codexBridge -match 'ThreadPoolExecutor' -and $codexBridge -match 'max_workers=min\(3')
    query_cache_is_short_and_bypassed_for_writes = ($codexBridge -match 'QUERY_CACHE_SECONDS = 30' -and $codexBridge -match 'use_cache=\(mode == "normal"\)' -and $bridge -match 'Clear-PetstarCodexQueryCache')
    message_latency_has_stage_telemetry = ($codexBridge -match 'codex_performance\.ndjson' -and $codexBridge -match 'model_seconds' -and $codexBridge -match 'query_seconds' -and $bridge -match 'message timing')
    order_local_ledger_rebuild_is_month_scoped = ($ledgerSync -match '\[string\]\$MonthFilter' -and $ledgerSync -match 'selected_months' -and $financeEvent -match 'monthLocalActions')
    replenishment_updates_only_local_cost_state = ($replenishmentBlock -match 'afterAverage' -and $replenishmentBlock -match "affected_tables=@\('补货记录','SKU成本主表'\)" -and $replenishmentBlock -notmatch 'Invoke-CostSnapshotRebuild|Invoke-DailyLedgerSync|库存批次')
    replenishment_defaults_to_same_day_full_payment = ($financeEvent -match 'explicitlyUnpaid' -and $financeEvent -match '未付款\|未支付\|赊账\|挂账\|欠款' -and $financeEvent -match '\$paymentAmountValue = if \(\$null -ne \$PaymentAmount\).+\$purchaseAmountValue' -and $codexBridge -match '默认当天已全额付款')
    plain_sku_cost_does_not_rebuild_history = ($setSkuCostBlock -notmatch 'Invoke-CostSnapshotRebuild' -and $setSkuCostBlock -match 'affectedTables = if \(\$OrderNo\)' -and $setSkuCostBlock -match "else \{ @\('SKU成本主表'\) \}")
    targeted_sku_cost_requires_order_and_month_scope = ($setSkuCostBlock -match '订单 \$OrderNo 中找不到' -and $financeEvent -match '\$Action -eq ''SetSkuCost'' -and \$null -ne \$explicitHistoricalOrderDate' -and $financeEvent -match 'Invoke-DailyLedgerSync \$explicitHistoricalOrderDate')
    order_sync_moves_cost_snapshot_with_item_identity = ($orderSync -match 'Get-OrderItemSnapshotKey' -and $orderSync -match 'Take-CostSnapshotFromQueue' -and $orderSync -match 'Set-CellMatrix ''订单商品'' "L2:N\$itemsLastRow"')
    scheduled_sync_initializes_only_targeted_order_costs = ($orderSync -match 'changed_order_nos' -and $scheduledSync -match 'Get-TargetedOrderNumbers' -and $scheduledSync -match 'Invoke-CostSnapshotRebuild \$targetedOrderNumbers \$true')
    order_sync_detects_business_changes_before_writing = ($orderSync -match 'Get-RecordsFingerprint' -and $orderSync -match '\$dataChanged = \[bool\]\$Initialize -or \$summaryChanged -or \$itemsChanged -or \$interventionsChanged' -and $orderSync -match 'if \(\$dataChanged\)[\s\S]{0,500}Set-CellMatrix ''订单汇总''')
    ordinary_no_change_sync_skips_writes = ($scheduledSync -match "SyncMode -notin @\('midday','reminder','night','manual'\)" -and $scheduledSync -match 'OrderSyncResult\.data_changed -eq \$false' -and $scheduledSync -match '00:00 仅执行本月只读审计')
    changed_orders_refunds_and_interventions_are_targeted = ($scheduledSync -match 'Get-TargetedOrderNumbers' -and $scheduledSync -match 'Get-TargetedMonthKeys' -and $scheduledSync -match 'Invoke-CostSnapshotRebuild \$targetedOrderNumbers \$true' -and $scheduledSync -match 'Invoke-DailyLedgerSync \$targetedMonthKeys' -and $costRebuild -match '\[switch\]\$Incremental')
    weekly_read_only_audit_gates_full_rebuild = ($weeklyMaintenance -match '第一阶段严格只读' -and $weeklyMaintenance -match 'if \(\$audit\.ok\)' -and $weeklyMaintenance -match "'-FullRebuild'" -and $scheduledTaskInstaller -match "Name='Petstar-Weekly-Audit'" -and $scheduledTaskInstaller -match "Days=@\('Monday'\)")
    historical_order_cost_snapshots_are_frozen = ($costRebuild -match 'previous_unit_cost.+eligibleOrderSet' -and $costRebuild -match '只有调用方明确标记的本次新增订单' -and $costRebuild -match 'item\.canonical -and \$eligibleOrderSet\.Contains')
    pending_historical_cost_stays_pending = ($costRebuild -match '\[string\[\]\]\$EligibleOrderNumbers = @\(\)' -and $costRebuild -match 'elseif \(\$eligibleOrderSet\.Contains')
    missing_sku_cost_never_uses_reference_fallback = ($costRebuild -notmatch "Get-SheetRows '商品成本'" -and $costRebuild -notmatch 'referenceCostByCode' -and $costRebuild -match '\$null -ne \$item\.previous_unit_cost -and \$null -ne \$item\.previous_item_cost')
    cost_confirmation_preserves_kw_conflict_warning = ($financeEvent -match "oldStatus -match '当前KW冲突'" -and $financeEvent -match '当前KW冲突待核对（成本已确认）')
    unique_active_sku_owns_shared_kw = ($skuIdentitySync -match 'active\.Count -ne 1' -and $skuIdentitySync -match 'retired\.Count -ne \(\$group\.Count - 1\)' -and $skuIdentitySync -match '\$retiredUpdate\.kw = \$null' -and $skuIdentitySync -match '已下架/未分配KW')
    unit_cost_precision_is_preserved = ($financeEvent -match 'function New-UnitCostCell' -and $financeEvent -match "number_format='#,##0\.####'" -and $costRebuild -match 'New-UnitCostCell \$item\.unit_cost' -and $costRebuild -match 'New-UnitCostCell \$state\.avg')
    scheduled_sync_yields_to_confirmed_writes = ($bridge -match 'interactive_priority\.json' -and $scheduledSync -match '人工确认写账优先，定时同步已让路')
    scheduled_sync_does_not_lock_before_chrome_check = ($scheduledSync.IndexOf("Invoke-WithRetry 'chrome_cdp'") -lt $scheduledSync.IndexOf('Start-SyncFinanceTransaction', $scheduledSync.IndexOf("Invoke-WithRetry 'chrome_cdp'")))
    chrome_preflight_is_bounded_and_nonretryable = ($browserPreflight -match 'waitForProxyTargets' -and $browserPreflight -match '45000' -and $browserPreflight -match 'MEITUAN_CDP_UNAVAILABLE' -and $scheduledSync -match "Invoke-WithRetry 'chrome_cdp'[\s\S]{0,700}\} 1" -and $scheduledSync -match '\$nonRetryable[\s\S]{0,300}MEITUAN_CDP_UNAVAILABLE')
    chrome_preflight_self_heals_stale_proxy_and_prompts_once = ($browserPreflight -match 'readChromeInstanceFingerprint' -and $browserPreflight -match 'readPromptedFingerprint\(\) === instanceFingerprint' -and $browserPreflight -match 'savePromptedFingerprint\(instanceFingerprint\)' -and $browserPreflight -match 'Repeated /targets calls' -and $browserPreflight -match '后续定时任务只保留飞书提醒' -and $scheduledSync -match "return 'meituan-cdp-auth'" -and $scheduledSync -match "Resolve-FailureNotification 'meituan-cdp-auth'")
    runtime_prefers_project_browser_preflight = ($runtimePaths -match 'Join-Path \$PSScriptRoot ''check_meituan_browser\.mjs''')
    manual_delivery_fee_fork_self_heals_to_authoritative_source = ($ledgerSync -match 'function Set-SourceCellMatrix' -and $ledgerSync -match 'auto_reconciled_delivery_fee_count' -and $ledgerSync -match 'Set-SourceCellMatrix ''订单汇总'' "J\$\{sourceRowNumber\}:J\$\{sourceRowNumber\}"' -and $ledgerSync -match '配送费自动归一回读失败' -and $ledgerSync -match '人工确认配送费' -and $ledgerSync -notmatch '请先运行 work\\reconcile_manual_delivery_fees\.ps1' -and $integrityAudit -match 'unmigrated_manual_delivery_fees' -and $integrityAudit -notmatch 'preserved_manual_delivery_fees')
    nightly_intervention_reconciliation_repairs_and_alerts = ($orderSync -match 'intervention_repair_count' -and $orderSync -match '订单汇总中找不到原订单' -and $orderSync -match 'New-NullableNumberCell \$record\.PlatformFee' -and $scheduledSync -match '补单数据自动补救' -and $scheduledSync -match '补单数据待处理' -and $ledgerSync -match '平台扣款尚未识别；未知不等于0元' -and $integrityAudit -match 'intervention_data_errors' -and $integrityAudit -match '-not \$interventionSet\.ContainsKey\(\$orderNo\)')
    manual_delivery_repair_is_atomic_and_read_back = ($deliveryReconciliation -match 'New-FinanceTransaction' -and $deliveryReconciliation -match 'Register-FinanceRangeBeforeImage' -and $deliveryReconciliation -match 'Invoke-FinanceTransactionRollback' -and $deliveryReconciliation -match '配送费回读订单数量不一致')
    dashboard_includes_offline_orders_and_matches_ledger = ($dashboardFormulaRepair -match "''线下订单''" -and $dashboardFormulaRepair -match 'dashboard_gross_profit' -and $dashboardFormulaRepair -match 'ledger_profit' -and $dashboardFormulaRepair -match 'Invoke-FinanceTransactionRollback')
    dashboard_uses_current_cost_inventory_and_compact_metrics = ($simpleFinanceBuilder -match '\(D\{r\}\+IFERROR\(M\{r\}\*1,0\)\)\*H\{r\}' -and $simpleFinanceBuilder -match 'SKU成本主表.+G.2.+G.3000' -and $simpleFinanceBuilder -notmatch '本月经济利润.+\$null' -and $dashboardPaymentRepair -match 'october_purchase_amount' -and $dashboardPaymentRepair -match 'Invoke-FinanceTransactionRollback')
    confirmed_writes_wait_for_running_sync = ($financeTransaction -match 'LockTimeoutSeconds' -and $bridge -match 'Invoke-StoredConfirmationOperations \$pending \$executeEvent 5' -and $bridge -match "Mode', 'RetryPending'" -and $bridge -match 'Invoke-StoredConfirmationOperations \$pending \$executeEvent 2700')
    codex_mutation_preview_reuses_first_parse = ($bridge -match 'function New-ConfirmationDraftFromCodexResult' -and $bridge -match 'Reuse them[\s\S]{0,240}New-ConfirmationDraftFromCodexResult')
}
$failedContracts = @($contracts.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object { $_.Key })
if ($failedContracts.Count -gt 0) { throw "Missing reliability contracts: $($failedContracts -join ', ')" }

$refundCentResult = [math]::Max([double]0.0, [double][math]::Round(([double]46.22 - [double]23.98), [int]2))
if ([math]::Abs($refundCentResult - 22.24) -gt 0.000001) { throw "Refund cent-preservation check failed: $refundCentResult" }

foreach ($stateName in @('pending_confirmations.json','pending_acknowledgements.json')) {
    $path = Join-Path $stateRoot $stateName
    if (-not (Test-Path -LiteralPath $path)) { continue }
    try { $null = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { throw "Runtime state is not valid JSON: $stateName" }
}

. (Join-Path $agentRoot 'runtime_paths.ps1')
$python = Get-PetstarPython
if (-not (Test-Path -LiteralPath $python)) { throw 'Project Python runtime was not found' }
& $python -m py_compile (Join-Path $agentRoot 'codex_chat_bridge.py') (Join-Path $agentRoot 'petstar_cli.py')
if ($LASTEXITCODE -ne 0) { throw 'Python entry-point compile check failed' }
$nodeFiles = @(
    (Join-Path $projectRoot 'work\meituan_order_extract.mjs'),
    (Join-Path $projectRoot 'work\test_meituan_refund_extraction.mjs'),
    (Join-Path $projectRoot 'work\meituan_inventory_export.mjs'),
    (Join-Path $projectRoot 'work\inspect_meituan_product_export_fields.mjs'),
    (Join-Path $projectRoot 'work\meituan_promotion_extract.mjs')
    (Join-Path $projectRoot 'work\meituan_browser_context.mjs')
    (Join-Path $agentRoot 'check_meituan_browser.mjs')
)
foreach ($nodeFile in $nodeFiles) {
    & node --check $nodeFile
    if ($LASTEXITCODE -ne 0) { throw "Node entry-point syntax check failed: $nodeFile" }
}
& node (Join-Path $projectRoot 'work\test_meituan_refund_extraction.mjs') | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'Meituan refund-state classification test failed' }

$resolvedWindowsPowerShell = Get-PetstarWindowsPowerShell
if (-not (Test-Path -LiteralPath $resolvedWindowsPowerShell -PathType Leaf)) { throw 'Windows PowerShell resolver returned a missing file' }
if ((Split-Path $resolvedWindowsPowerShell -Leaf) -ne 'powershell.exe') { throw 'Windows PowerShell resolver returned the wrong executable' }

[ordered]@{
    ok = $true
    powershell_files_parsed = $runtimePowerShellFiles.Count
    python_entries_compiled = 2
    node_entries_checked = $nodeFiles.Count
    contracts = $contracts
    external_writes = 0
} | ConvertTo-Json -Depth 5
