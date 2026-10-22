param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('AddIntervention','RemoveIntervention','RestoreInventory','ConfirmSelfDelivery','SetExternalDeliveryFee','SetOrderDeliveryFee','SetOrderActuals','OfflineSale','AddOrderGift','Replenishment','BatchSkuMaintenance','CashExpense','OperatingExpense','OwnerWithdrawal','OwnerLabor','MarketRent','BeginningCash','ActualReceipt','AssetAmortization','SetSkuExpiry','SetSkuBatchDate','SetSkuCost','ChangeSkuCode','SetRefundDisposition')]
    [string]$Action,
    [datetime]$Date = [datetime]::Today,
    [string]$OrderNo,
    [string]$DailyOrderNo,
    [string]$Person,
    [string]$Items,
    [string]$GoodsCode,
    [string]$TransitionCode,
    [string]$NewGoodsCode,
    [ValidateSet('returned_restocked','refund_no_return','returned_damaged')][string]$RefundDisposition,
    [switch]$RetireExistingNewCode,
    [Nullable[datetime]]$ExpiryDate,
    [switch]$ClearExpiry,
    [Nullable[datetime]]$ProductionDate,
    [Nullable[double]]$ShelfLifeValue,
    [ValidateSet('天','月','年')][string]$ShelfLifeUnit = '月',
    [Nullable[double]]$Quantity,
    [Nullable[double]]$PurchaseAmount,
    [Nullable[double]]$PaymentAmount,
    [Nullable[datetime]]$PaymentDate,
    [Nullable[double]]$Amount,
    [Nullable[double]]$DeliveryFee,
    [switch]$ReplaceMatchingProject,
    [string]$ExpenseSource = '手工',
    [string]$ExpenseMemo = '通过飞书确认录入',
    [ValidateSet('已到账','未到账')][string]$ReceiptStatus = '已到账',
    [ValidateSet('美团','淘宝闪购','京东到家','线下')][string]$Platform = '美团',
    [string]$Note,
    [string]$EventId,
    [string]$OperationsJson,
    [string]$OperationsPath,
    [switch]$DryRun,
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$projectRoot = $ProjectRoot
$metaPath = Join-Path $projectRoot 'work\simple_finance_meta.json'
$ledgerMetaPath = Join-Path $projectRoot 'work\daily_ledger_meta.json'
$ledgerScript = Join-Path $projectRoot 'agent\sync_daily_ledger.ps1'
$costScript = Join-Path $projectRoot 'agent\rebuild_cost_snapshots.ps1'
$orderSyncScript = Join-Path $projectRoot 'work\sync_meituan_orders.ps1'
$powershellExe = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$larkCli = Join-Path $PSScriptRoot 'lark_cli_noexit.ps1'
$stateRoot = Join-Path $projectRoot 'agent\state'
$eventLogPath = Join-Path $stateRoot 'processed_events.ndjson'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少飞书经营驾驶舱元数据' }
$meta = [IO.File]::ReadAllText($metaPath,[Text.Encoding]::UTF8) | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$ledgerSpreadsheetToken = $null
if (Test-Path -LiteralPath $ledgerMetaPath) {
    $ledgerMeta = [IO.File]::ReadAllText($ledgerMetaPath,[Text.Encoding]::UTF8) | ConvertFrom-Json
    $ledgerSpreadsheetToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
}
. (Join-Path $PSScriptRoot 'finance_transaction.ps1')
. (Join-Path $PSScriptRoot 'sku_master_schema.ps1')
. (Join-Path $PSScriptRoot 'intervention_record.ps1')
. (Join-Path $PSScriptRoot 'refund_adjustments.ps1')
. (Join-Path $PSScriptRoot 'order_actual_overrides.ps1')
$skuSchema = Get-SkuMasterSchema -SpreadsheetToken $spreadsheetToken
$script:financeTransaction = $null
$script:financeTransactionOwner = $false
$script:financeTransactionClosed = $false
$script:refundStateBackup = $null
$script:refundStateBackupPath = $null
$script:orderActualOverrideStateBackup = $null
$script:orderActualOverrideStateBackupPath = $null

function Get-FinanceWorkbookTokens {
    $tokens = @{source=$spreadsheetToken}
    if ($ledgerSpreadsheetToken) { $tokens['ledger'] = $ledgerSpreadsheetToken }
    return $tokens
}
function Complete-CurrentFinanceTransaction {
    if ($null -eq $script:financeTransaction -or $script:financeTransactionClosed) { return }
    if ($script:financeTransactionOwner) { Complete-FinanceTransaction $script:financeTransaction }
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
}
function Fail-CurrentFinanceTransaction([string]$Reason) {
    if ($null -eq $script:financeTransaction -or $script:financeTransactionClosed) { return $null }
    $rollback = $null
    if ($script:financeTransactionOwner) {
        $rollback = Invoke-FinanceTransactionRollback $script:financeTransaction (Get-FinanceWorkbookTokens) $larkCli $Reason
    }
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
    return $rollback
}

trap {
    $reason = [string]$_.Exception.Message
    $failureLocation = [string]$_.ScriptStackTrace
    if ($failureLocation) { $reason += "；位置：$failureLocation" }
    $refundBackupVariable = Get-Variable -Name refundStateBackup -Scope Script -ErrorAction SilentlyContinue
    $refundBackupPathVariable = Get-Variable -Name refundStateBackupPath -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $refundBackupVariable -and $null -ne $refundBackupVariable.Value -and
        $null -ne $refundBackupPathVariable -and [string]$refundBackupPathVariable.Value) {
        try { [IO.File]::WriteAllText([string]$refundBackupPathVariable.Value, [string]$refundBackupVariable.Value, [Text.UTF8Encoding]::new($false)) }
        catch { $reason += "；退款状态恢复失败：$($_.Exception.Message)" }
    }
    $overrideBackupVariable = Get-Variable -Name orderActualOverrideStateBackup -Scope Script -ErrorAction SilentlyContinue
    $overrideBackupPathVariable = Get-Variable -Name orderActualOverrideStateBackupPath -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $overrideBackupVariable -and $null -ne $overrideBackupVariable.Value -and
        $null -ne $overrideBackupPathVariable -and [string]$overrideBackupPathVariable.Value) {
        try { [IO.File]::WriteAllText([string]$overrideBackupPathVariable.Value, [string]$overrideBackupVariable.Value, [Text.UTF8Encoding]::new($false)) }
        catch { $reason += "；订单实际履约覆盖恢复失败：$($_.Exception.Message)" }
    }
    $rollback = $null
    if (Get-Command Fail-CurrentFinanceTransaction -CommandType Function -ErrorAction SilentlyContinue) {
        $rollback = Fail-CurrentFinanceTransaction $reason
    }
    if ($null -ne $rollback -and -not $rollback.ok) {
        [Console]::Error.WriteLine("经营事件失败，且自动回滚待重试：$reason；事务 $($rollback.transaction_id)")
    } elseif ($null -ne $rollback -and $rollback.status -eq 'rolled_back') {
        [Console]::Error.WriteLine("经营事件失败，本次改动已全部回滚：$reason；事务 $($rollback.transaction_id)")
    } else {
        [Console]::Error.WriteLine($reason)
    }
    exit 1
}

function Convert-ToExcelSerial([datetime]$Value) { return ($Value.Date - [datetime]'1899-12-30').TotalDays }
function Get-StableEventKey([string]$Value, [int]$Length = 20) {
    if (-not $Value) { return $null }
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString(
            $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value))
        ).Replace('-', '').Substring(0,$Length).ToUpperInvariant()
    } finally {
        $sha.Dispose()
    }
}
function New-ValueCell($Value) { if ($null -eq $Value -or [string]$Value -eq '') { return @{} }; return @{value=$Value} }
function New-UnitCostCell($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return @{} }
    return @{value=$Value;cell_styles=@{number_format='#,##0.####'}}
}
function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    $valueProperty = $cell.PSObject.Properties['value']
    if ($null -eq $valueProperty -or $null -eq $valueProperty.Value) { return $null }
    return [string]$valueProperty.Value
}
function Convert-ToNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return 0.0 }
    $number = 0.0
    if (-not [double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { throw "金额不是有效数字：$Value" }
    return $number
}
function Convert-ToNullableNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    return Convert-ToNumber $Value
}
function Convert-ToDateValue($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.Date }
    $serial = 0.0
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$serial)) { return ([datetime]'1899-12-30').AddDays($serial).Date }
    return $null
}
function Invoke-LarkReadJson([scriptblock]$Command, [string]$Label, [int]$MaxAttempts = 4) {
    $errors = [Collections.Generic.List[string]]::new()
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $global:LASTEXITCODE = 0
        try {
            $raw = & $Command | Out-String
            $exitCode = $LASTEXITCODE
            $result = $raw | ConvertFrom-Json
            if ($exitCode -eq 0 -and $result.ok -ne $false) { return $result }
            $detail = if ($result.error -and $result.error.message) { [string]$result.error.message } else { "退出码=$exitCode" }
            $errors.Add("第${attempt}次：$detail")
        } catch {
            $errors.Add("第${attempt}次：$($_.Exception.Message)")
        }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds ([math]::Min(8, $attempt * 2)) }
    }
    throw "$Label；重试${MaxAttempts}次仍失败：$($errors -join '；')"
}
function Get-SheetRows([string]$SheetName, [string]$Range) {
    $result = Invoke-LarkReadJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --include value --max-chars 5000000 --format json } "读取飞书失败：$SheetName"
    if ($result.data.has_more) { throw "飞书读取结果不完整：$SheetName" }
    return @($result.data.ranges[0].cells)
}
function Get-LedgerSheetRows([string]$SheetName, [string]$Range) {
    if (-not $ledgerSpreadsheetToken) { throw '缺少每日经营简账元数据，无法登记支出明细' }
    $result = Invoke-LarkReadJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $ledgerSpreadsheetToken --sheet-name $SheetName --range $Range --include value --max-chars 5000000 --format json } "读取飞书简账失败：$SheetName"
    if ($result.data.has_more) { throw "飞书简账读取结果不完整：$SheetName" }
    return @($result.data.ranges[0].cells)
}
function Get-SingleSheetRow([string]$SheetName, [string]$Range) {
    $result = Invoke-LarkReadJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --include value --max-chars 10000 --format json } "读取飞书失败：$SheetName $Range"
    if ($result.data.has_more) { throw "飞书读取结果不完整：$SheetName $Range" }
    $matrix = $result.data.ranges[0].cells
    if ($null -eq $matrix -or @($matrix).Count -eq 0) { return $null }
    return @($matrix)[0]
}
function Get-SkuCostReadback([int]$RowNumber, [double]$ExpectedCost, [string]$ExpectedStatus, [Nullable[double]]$ExpectedMovingCost = $null, [int]$MaxAttempts = 12) {
    $lastResult = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        # 不经过 Get-SheetRows：PowerShell 会把“仅一行”的二维数组展开成单元格列表，
        # 再取 [0] 就只剩 A 列，导致 J 列状态被误读为空。
        $raw = & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range "A${RowNumber}:Z${RowNumber}" --include value --max-chars 10000 --format json | Out-String
        if ($LASTEXITCODE -ne 0) { throw "读取飞书失败：SKU成本主表 A${RowNumber}:Z${RowNumber}" }
        $result = $raw | ConvertFrom-Json
        if (-not $result.ok -or $result.data.has_more) { throw "飞书读取结果不完整：SKU成本主表 A${RowNumber}:Z${RowNumber}" }
        $row = @($result.data.ranges[0].cells)[0]
        $actualBaseline = Convert-ToNumber (Get-CellValue $row $skuSchema.BaselineUnitCost.Index)
        $actualMoving = Convert-ToNumber (Get-CellValue $row $skuSchema.MovingAverageCost.Index)
        $actualStatus = Get-CellValue $row $skuSchema.CostStatus.Index
        $lastResult = [pscustomobject]@{
            row = $row
            baseline_cost = $actualBaseline
            moving_average_cost = $actualMoving
            status = $actualStatus
            attempt = $attempt
        }
        $movingMatches = $null -eq $ExpectedMovingCost -or [math]::Abs($actualMoving - [double]$ExpectedMovingCost) -le 0.0001
        if ([math]::Abs($actualBaseline - $ExpectedCost) -le 0.0001 -and $movingMatches -and $actualStatus -eq $ExpectedStatus) {
            return $lastResult
        }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Milliseconds 1000 }
    }
    return $lastResult
}
function Set-CellMatrix([string]$SheetName, [string]$Range, $Rows) {
    if ($null -ne $script:financeTransaction) {
        Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $spreadsheetToken $SheetName $Range $larkCli
    }
    $payload = ConvertTo-Json -InputObject $Rows -Depth 10 -Compress
    $raw = $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入飞书失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "写入飞书失败：$SheetName $Range" }
}
function Set-OneRow([string]$SheetName, [string]$Range, [object[]]$Values) {
    $row = @()
    foreach ($value in $Values) { $row += ,(New-ValueCell $value) }
    $matrix = @(); $matrix += ,$row
    Set-CellMatrix $SheetName $Range $matrix
}
function Set-LedgerOneRow([string]$SheetName, [string]$Range, [object[]]$Values) {
    if (-not $ledgerSpreadsheetToken) { throw '缺少每日经营简账元数据，无法登记支出明细' }
    if ($null -ne $script:financeTransaction) {
        Register-FinanceRangeBeforeImage $script:financeTransaction 'ledger' $ledgerSpreadsheetToken $SheetName $Range $larkCli
    }
    $row = @()
    foreach ($value in $Values) { $row += ,(New-ValueCell $value) }
    $matrix = @(); $matrix += ,$row
    $payload = ConvertTo-Json -InputObject $matrix -Depth 10 -Compress
    $raw = $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $ledgerSpreadsheetToken --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入飞书简账失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "写入飞书简账失败：$SheetName $Range" }
}
function Clear-CellRange([string]$SheetName, [string]$Range) {
    if ($null -ne $script:financeTransaction) {
        Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $spreadsheetToken $SheetName $Range $larkCli
    }
    $ranges = ConvertTo-Json -InputObject @("$SheetName!$Range") -Compress
    $raw = $ranges | & $larkCli sheets +cells-batch-clear --as user --spreadsheet-token $spreadsheetToken --ranges - --scope content --yes --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "清除飞书范围失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "清除飞书范围失败：$SheetName $Range" }
}
function Set-AtomicCellOperations([object[]]$Operations) {
    if ($null -ne $script:financeTransaction) {
        foreach ($operation in @($Operations)) {
            if ([string]$operation.shortcut -ne '+cells-set') { throw '事务批量写入只允许 +cells-set' }
            Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $spreadsheetToken ([string]$operation.input.sheet_name) ([string]$operation.input.range) $larkCli
        }
    }
    $payload = ConvertTo-Json -InputObject $Operations -Depth 12 -Compress
    $raw = $payload | & $larkCli sheets +batch-update --as user --spreadsheet-token $spreadsheetToken --operations - --yes --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '飞书原子批量写入失败' }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw '飞书原子批量写入失败' }
}
function Assert-FormulaRange([string]$SheetName, [string]$Range) {
    $raw = & $larkCli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --format json --exit-on-error | Out-String
    if ($LASTEXITCODE -ne 0) { throw "公式校验失败：$SheetName" }
    $result = $raw | ConvertFrom-Json
    if ($result.data.status -ne 'success') { throw "公式存在错误：$SheetName" }
}
function Assert-Positive([Nullable[double]]$Value, [string]$Name) {
    if ($null -eq $Value -or [double]$Value -le 0) { throw "$Name 必须大于0" }
}
function Assert-NonNegative([Nullable[double]]$Value, [string]$Name) {
    if ($null -eq $Value -or [double]$Value -lt 0) { throw "$Name 不能小于0" }
}
function Add-ShelfLife([datetime]$StartDate, [double]$Value, [string]$Unit) {
    if ($Value -le 0) { throw '保质期必须大于0' }
    if ([math]::Abs($Value - [math]::Round($Value)) -gt 0.000001 -and $Unit -in @('月','年')) { throw '按月或按年填写保质期时必须是整数' }
    switch ($Unit) {
        '天' { return $StartDate.Date.AddDays($Value) }
        '月' { return $StartDate.Date.AddMonths([int][math]::Round($Value)) }
        '年' { return $StartDate.Date.AddYears([int][math]::Round($Value)) }
        default { throw '保质期单位只支持天、月、年' }
    }
}

function Get-CanonicalSkuMap($Rows) {
    $rowList = @($Rows)
    $candidates = @{}
    for ($i = 0; $i -lt $rowList.Count; $i++) {
        $row = $rowList[$i]
        $kw = (Get-CellValue $row 0)
        $aliases = (Get-CellValue $row 1)
        $productId = (Get-CellValue $row $skuSchema.ProductId.Index)
        $codes = @()
        if ($kw) { $codes += $kw.Trim().ToLowerInvariant() }
        if ($aliases) {
            $codes += @([regex]::Split($aliases, '[/,，、;；\s]+') | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLowerInvariant() })
        }
        if ($productId) { $codes += ('sku_id:' + $productId.Trim().ToLowerInvariant()) }
        foreach ($code in @($codes | Select-Object -Unique)) {
            if (-not $candidates.ContainsKey($code)) {
                $candidates[$code] = [Collections.Generic.List[object]]::new()
            }
            $candidates[$code].Add([pscustomobject]@{ row_number = 2 + $i; cells = $row })
        }
    }
    $result = @{}
    foreach ($code in $candidates.Keys) {
        # Each master row contributes a code at most once (the source code list is
        # already de-duplicated), so a single candidate is unambiguous. Avoid
        # Sort-Object here: thousands of codes made preview construction CPU-bound.
        if ($candidates[$code].Count -eq 1) { $result[$code] = $candidates[$code][0] }
    }
    return $result
}

function Get-SkuCodeCandidates($Rows, [string]$RequestedCode) {
    $rowList = @($Rows)
    $normalizedCode = $RequestedCode.Trim().ToLowerInvariant()
    $matches = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $rowList.Count; $i++) {
        $row = $rowList[$i]
        $codes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $kw = Get-CellValue $row 0
        $aliases = Get-CellValue $row 1
        if ($kw) { $null = $codes.Add($kw.Trim()) }
        if ($aliases) {
            foreach ($alias in @([regex]::Split($aliases, '[/,，、;；\s]+') | Where-Object { $_ })) {
                $null = $codes.Add($alias.Trim())
            }
        }
        if ($codes.Contains($normalizedCode)) {
            $matches.Add([pscustomobject]@{row_number=2+$i;cells=$row})
        }
    }
    return @($matches)
}

function Resolve-UniqueSkuReference($Rows, [string]$RequestedCode) {
    $rowList = @($Rows)
    $normalizedCode = $RequestedCode.Trim().ToLowerInvariant()
    $matches = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $rowList.Count; $i++) {
        $row = $rowList[$i]
        $matched = $false
        if ($normalizedCode.StartsWith('sku_id:')) {
            $requestedProductId = $normalizedCode.Substring(7)
            $productId = Get-CellValue $row $skuSchema.ProductId.Index
            $matched = $productId -and $productId.Trim().ToLowerInvariant() -eq $requestedProductId
        } else {
            $kw = Get-CellValue $row 0
            if ($kw -and $kw.Trim().ToLowerInvariant() -eq $normalizedCode) {
                $matched = $true
            } else {
                $aliases = Get-CellValue $row 1
                if ($aliases) {
                    foreach ($alias in [regex]::Split($aliases, '[/,，、;；\s]+')) {
                        if ($alias -and $alias.Trim().ToLowerInvariant() -eq $normalizedCode) { $matched = $true; break }
                    }
                }
            }
        }
        if ($matched) { $matches.Add([pscustomobject]@{row_number=2+$i;cells=$row}) }
    }
    if ($matches.Count -eq 0) { throw "SKU成本主表中找不到货号：$RequestedCode" }
    if ($matches.Count -gt 1) { throw "SKU成本主表中标识不唯一：$RequestedCode（匹配 $($matches.Count) 行），请使用系统返回的 SKU ID 定位" }
    return $matches[0]
}

function Resolve-OrderForGift([string]$RequestedOrderNo, [string]$RequestedDailyOrderNo, [datetime]$RequestedDate) {
    $matches = [Collections.Generic.List[object]]::new()
    foreach ($row in (Get-SheetRows '订单汇总' 'A2:Q2001')) {
        $platformOrderNo = Get-CellValue $row 2
        $dailyNo = Get-CellValue $row 16
        $orderDate = Convert-ToDateValue (Get-CellValue $row 0)
        if (-not $platformOrderNo -or $null -eq $orderDate) { continue }
        $isMatch = if ($RequestedOrderNo) {
            $platformOrderNo -eq $RequestedOrderNo
        } else {
            $dailyNo -eq $RequestedDailyOrderNo -and $orderDate.Date -eq $RequestedDate.Date
        }
        if ($isMatch) {
            $matches.Add([pscustomobject]@{
                order_no = $platformOrderNo
                daily_order_no = $dailyNo
                date = $orderDate.Date
                platform = $(if (Get-CellValue $row 1) { Get-CellValue $row 1 } else { '美团' })
                order_type = Get-CellValue $row 4
                order_status = Get-CellValue $row 3
            })
        }
    }
    foreach ($row in (Get-SheetRows '线下订单' 'A2:J1001')) {
        $offlineOrderNo = Get-CellValue $row 2
        $dailyNo = Get-CellValue $row 1
        $orderDate = Convert-ToDateValue (Get-CellValue $row 0)
        if (-not $offlineOrderNo -or $null -eq $orderDate) { continue }
        $isMatch = if ($RequestedOrderNo) {
            $offlineOrderNo -eq $RequestedOrderNo
        } else {
            $dailyNo -eq $RequestedDailyOrderNo -and $orderDate.Date -eq $RequestedDate.Date
        }
        if ($isMatch) {
            $matches.Add([pscustomobject]@{order_no=$offlineOrderNo;daily_order_no=$dailyNo;date=$orderDate.Date;platform='线下';order_type='自然'})
        }
    }
    if ($matches.Count -eq 0) {
        $label = if ($RequestedOrderNo) { $RequestedOrderNo } else { "$($RequestedDate.ToString('yyyy-MM-dd')) $RequestedDailyOrderNo" }
        throw "找不到要登记赠品的订单：$label。请先同步订单后再登记"
    }
    if ($matches.Count -gt 1) { throw '赠品对应订单不唯一，请提供完整订单编号' }
    if ($matches[0].order_status -eq '已取消') { throw '已取消订单不能登记赠品' }
    return $matches[0]
}

function Get-OrderNumberDistance([string]$Left, [string]$Right) {
    if (-not $Left -or -not $Right -or $Left.Length -ne $Right.Length) { return [int]::MaxValue }
    $distance = 0
    for ($index = 0; $index -lt $Left.Length; $index++) {
        if ($Left[$index] -ne $Right[$index]) { $distance++ }
    }
    return $distance
}
function Convert-ChildOutputJson($Output, [string]$Label) {
    $text = [string](@($Output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine)
    $start = $text.IndexOf('{')
    $end = $text.LastIndexOf('}')
    if ($start -lt 0 -or $end -lt $start) { throw "$Label未返回有效结果" }
    try {
        return ($text.Substring($start, $end - $start + 1) | ConvertFrom-Json)
    } catch {
        throw "$Label返回结果无法解析：$($_.Exception.Message)"
    }
}
function Invoke-DailyLedgerSync([Nullable[datetime]]$TargetMonth) {
    if (-not (Test-Path -LiteralPath $ledgerScript)) { throw '缺少每日经营简账刷新脚本' }
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $arguments = @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$ledgerScript,'-ThroughDate',([datetime]::Today).ToString('yyyy-MM-dd'),'-ProjectRoot',$projectRoot)
        if ($null -ne $TargetMonth) { $arguments += @('-MonthFilter',([datetime]$TargetMonth).ToString('yyyy-MM')) }
        $output = @(& $powershellExe @arguments 2>&1)
        $childExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
    if ($childExitCode -ne 0) {
        $detail = [string](@($output | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() } | Select-Object -Last 1) -join '')
        throw "每日经营简账刷新失败：$detail"
    }
    return (Convert-ChildOutputJson $output '每日经营简账')
}
function Invoke-CostSnapshotRebuild {
    if (-not (Test-Path -LiteralPath $costScript)) { throw '缺少成本快照刷新脚本' }
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $costScript -ProjectRoot $projectRoot 2>&1)
        $childExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
    if ($childExitCode -ne 0) {
        $detail = [string](@($output | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() } | Select-Object -Last 1) -join '')
        throw "成本快照刷新失败：$detail"
    }
    return (Convert-ChildOutputJson $output '成本快照')
}
function Invoke-OrderSyncForDate([datetime]$TargetDate) {
    if (-not (Test-Path -LiteralPath $orderSyncScript)) { throw '缺少美团订单同步脚本' }
    $dateText = $TargetDate.ToString('yyyy-MM-dd')
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $orderSyncScript -StartDate $dateText -EndDate $dateText -WorkRoot (Join-Path $projectRoot 'work') 2>&1)
        $childExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
    if ($childExitCode -ne 0) {
        $detail = [string](@($output | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() } | Select-Object -Last 1) -join '')
        throw "订单实际履约覆盖后重新同步失败：$detail"
    }
    return (Convert-ChildOutputJson $output '美团订单同步')
}
function Test-ProcessedEvent([string]$Id) {
    if (-not $Id) { return $false }
    if (Test-Path -LiteralPath $eventLogPath) {
        foreach ($line in Get-Content -LiteralPath $eventLogPath) {
            if (-not $line) { continue }
            try { if (($line | ConvertFrom-Json).event_id -eq $Id) { return $true } } catch { continue }
        }
    }
    return (Test-FinanceCommittedEvent $stateRoot $Id)
}
function Save-ProcessedEvent([string]$Id, [string]$EventAction, [string]$Reference) {
    if (-not $Id) { return }
    if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
    if ($null -ne $script:financeTransaction) {
        Register-FinanceProcessedEvent $script:financeTransaction $Id $EventAction $Reference
        return
    }
    [System.IO.File]::AppendAllText(
        $eventLogPath,
        (([ordered]@{event_id=$Id;action=$EventAction;reference=$Reference;processed_at=[datetime]::Now.ToString('o')} | ConvertTo-Json -Compress) + [Environment]::NewLine),
        [Text.UTF8Encoding]::new($false)
    )
}
function Get-FirstEmptyRow($Rows, [int]$KeyColumn, [int]$StartRow, [int]$LastRow) {
    for ($i = 0; $i -lt @($Rows).Count; $i++) {
        if (-not (Get-CellValue $Rows[$i] $KeyColumn)) { return $StartRow + $i }
    }
    $candidate = $StartRow + @($Rows).Count
    if ($candidate -le $LastRow) { return $candidate }
    throw '表格容量已满'
}
function Get-FirstCompletelyEmptyRow($Rows, [int]$ColumnCount, [int]$StartRow, [int]$LastRow) {
    for ($i = 0; $i -lt @($Rows).Count; $i++) {
        $hasValue = $false
        for ($column = 0; $column -lt $ColumnCount; $column++) {
            if (Get-CellValue $Rows[$i] $column) { $hasValue = $true; break }
        }
        if (-not $hasValue) { return $StartRow + $i }
    }
    $candidate = $StartRow + @($Rows).Count
    if ($candidate -le $LastRow) { return $candidate }
    throw '表格容量已满'
}
function Get-MonthRow([datetime]$TargetDate) {
    $rows = Get-SheetRows '月度结算' 'A2:M25'
    $prefix = $TargetDate.ToString('yyyy-MM')
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $value = Get-CellValue $rows[$i] 0
        if ($value -and $value.StartsWith($prefix)) { return [pscustomobject]@{ row_number=2+$i; cells=$rows[$i] } }
    }
    throw "月度结算中不存在月份：$prefix"
}
function Update-MonthlyCell([int]$ColumnIndex, [string]$ColumnName, [double]$Value, [bool]$Add, [string]$Description) {
    $month = Get-MonthRow $Date
    $current = Convert-ToNumber (Get-CellValue $month.cells $ColumnIndex)
    $newValue = if ($Add) { $current + $Value } else { $Value }
    Set-OneRow '月度结算' "$ColumnName$($month.row_number):$ColumnName$($month.row_number)" @($newValue)
    if ($Description) {
        $oldNote = Get-CellValue $month.cells 11
        $entry = "$($Date.ToString('MM-dd')) $Description"
        $newNote = if ($oldNote) { "$oldNote；$entry" } else { $entry }
        Set-OneRow '月度结算' "L$($month.row_number):L$($month.row_number)" @($newNote)
    }
    return [pscustomobject]@{row=$month.row_number;value=$newValue}
}

function Normalize-BusinessText([string]$Value) {
    if (-not $Value) { return '' }
    return (($Value.Trim().ToLowerInvariant()) -replace '[\s，,。；;：:]+','')
}

function Get-OperatingExpenseState([datetime]$TargetDate, [double]$TargetAmount, [string]$TargetNote) {
    if (-not $ledgerSpreadsheetToken) { throw '缺少每日经营简账元数据，无法登记其他经营费用' }
    $project = if ($TargetNote -and $TargetNote.Trim()) { $TargetNote.Trim() } else { '其他经营费用' }
    $sheetPrefix = if ($TargetDate.Year -eq 2026) { "$($TargetDate.Month)月" } else { "$($TargetDate.Year)年$($TargetDate.Month)月" }
    $sheetName = "${sheetPrefix}支出"
    $ledgerRows = Get-LedgerSheetRows $sheetName 'A2:F1001'
    $matches = [Collections.Generic.List[object]]::new()
    $projectMatches = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $ledgerRows.Count; $index++) {
        $row = $ledgerRows[$index]
        $rowDate = Convert-ToDateValue (Get-CellValue $row 0)
        $rowProject = Get-CellValue $row 1
        $rowAmount = Convert-ToNumber (Get-CellValue $row 2)
        if ($null -ne $rowDate -and $rowDate.Date -eq $TargetDate.Date -and
            (Normalize-BusinessText $rowProject) -eq (Normalize-BusinessText $project)) {
            $item = [pscustomobject]@{row_number=2+$index;cells=$row;amount=$rowAmount}
            $projectMatches.Add($item)
            if ($null -ne $rowAmount -and [math]::Abs($rowAmount - $TargetAmount) -le 0.0001) {
                $matches.Add($item)
            }
        }
    }

    $month = Get-MonthRow $TargetDate
    $entry = "$($TargetDate.ToString('MM-dd')) 其他经营费用 $TargetAmount：$project"
    $entryKey = Normalize-BusinessText $entry
    $noteParts = @([regex]::Split([string](Get-CellValue $month.cells 11), '[；;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $monthNoteCount = @($noteParts | Where-Object { (Normalize-BusinessText $_) -eq $entryKey }).Count
    $projectNotePrefix = Normalize-BusinessText "$($TargetDate.ToString('MM-dd')) 其他经营费用"
    $projectNoteSuffix = Normalize-BusinessText $project
    $projectNoteIndexes = [Collections.Generic.List[int]]::new()
    for ($noteIndex = 0; $noteIndex -lt $noteParts.Count; $noteIndex++) {
        $normalized = Normalize-BusinessText $noteParts[$noteIndex]
        if ($normalized.StartsWith($projectNotePrefix) -and $normalized.EndsWith($projectNoteSuffix)) {
            $projectNoteIndexes.Add($noteIndex)
        }
    }

    return [pscustomobject]@{
        project = $project
        sheet_name = $sheetName
        ledger_rows = $ledgerRows
        ledger_matches = $matches.ToArray()
        ledger_match_count = $matches.Count
        project_matches = $projectMatches.ToArray()
        project_match_count = $projectMatches.Count
        month = $month
        month_note_parts = $noteParts
        month_note_count = $monthNoteCount
        project_note_indexes = $projectNoteIndexes.ToArray()
        project_note_count = $projectNoteIndexes.Count
        month_note_entry = $entry
    }
}

function Invoke-OperatingExpenseWrite([datetime]$TargetDate, [double]$TargetAmount, [string]$TargetNote, [switch]$PreviewOnly) {
    $state = Get-OperatingExpenseState $TargetDate $TargetAmount $TargetNote
    if ($ReplaceMatchingProject) {
        if ($state.project_match_count -gt 1 -or $state.project_note_count -gt 1) {
            throw "检测到同日同项目重复记录：$($TargetDate.ToString('yyyy-MM-dd')) $($state.project)；已停止自动更新，请先核对重复行"
        }
        if (($state.project_match_count -eq 1) -xor ($state.project_note_count -eq 1)) {
            throw "同日同项目跨表状态不一致：$($TargetDate.ToString('yyyy-MM-dd')) $($state.project)；为避免错账，本次未更新"
        }
        if ($state.project_match_count -eq 1 -and $state.project_note_count -eq 1) {
            $existing = $state.project_matches[0]
            $oldAmount = Convert-ToNumber $existing.amount
            if ($null -eq $oldAmount) { throw "现有推广费金额不是有效数字：$($state.sheet_name) 第 $($existing.row_number) 行" }
            $delta = [math]::Round($TargetAmount - $oldAmount, 2)
            if ([math]::Abs($delta) -le 0.0001) {
                return [pscustomobject]@{
                    status = 'duplicate'; changed = $false; date = $TargetDate.ToString('yyyy-MM-dd')
                    project = $state.project; amount = $TargetAmount; previous_amount = $oldAmount
                    expense_sheet = $state.sheet_name; expense_row = [int]$existing.row_number; month_row = [int]$state.month.row_number
                }
            }
            if ($PreviewOnly) {
                return [pscustomobject]@{
                    status = 'ready_to_update'; changed = $false; date = $TargetDate.ToString('yyyy-MM-dd')
                    project = $state.project; amount = $TargetAmount; previous_amount = $oldAmount; delta = $delta
                    expense_sheet = $state.sheet_name; expense_row = [int]$existing.row_number; month_row = [int]$state.month.row_number
                }
            }

            Set-LedgerOneRow $state.sheet_name "A$($existing.row_number):F$($existing.row_number)" @(
                (Convert-ToExcelSerial $TargetDate), $state.project, $TargetAmount, $null, $ExpenseSource, $ExpenseMemo
            )
            $currentMonthlyValue = Convert-ToNumber (Get-CellValue $state.month.cells 6)
            $newMonthlyValue = [math]::Round($currentMonthlyValue + $delta, 2)
            Set-OneRow '月度结算' "G$($state.month.row_number):G$($state.month.row_number)" @($newMonthlyValue)
            $newParts = @($state.month_note_parts)
            $newParts[[int]$state.project_note_indexes[0]] = $state.month_note_entry
            Set-OneRow '月度结算' "L$($state.month.row_number):L$($state.month.row_number)" @(($newParts -join '；'))

            $check = Get-OperatingExpenseState $TargetDate $TargetAmount $TargetNote
            $actualMonthlyValue = Convert-ToNumber (Get-CellValue $check.month.cells 6)
            if ($check.project_match_count -ne 1 -or $check.ledger_match_count -ne 1 -or
                $check.project_note_count -ne 1 -or $check.month_note_count -ne 1 -or
                [math]::Abs($actualMonthlyValue - $newMonthlyValue) -gt 0.0001) {
                throw "推广费更新后回读校验失败：$($TargetDate.ToString('yyyy-MM-dd')) $($state.project) ¥$TargetAmount"
            }
            return [pscustomobject]@{
                status = 'updated'; changed = $true; date = $TargetDate.ToString('yyyy-MM-dd')
                project = $state.project; amount = $TargetAmount; previous_amount = $oldAmount; delta = $delta
                expense_sheet = $state.sheet_name; expense_row = [int]$existing.row_number
                month_row = [int]$state.month.row_number; month_value = $newMonthlyValue
            }
        }
    }
    if ($state.ledger_match_count -gt 1 -or $state.month_note_count -gt 1) {
        throw "检测到重复的其他经营费用：$($TargetDate.ToString('yyyy-MM-dd')) $($state.project) ¥$TargetAmount；已停止继续累计，请先清理重复记录"
    }
    if (($state.ledger_match_count -eq 1) -xor ($state.month_note_count -eq 1)) {
        throw "其他经营费用跨表状态不一致：$($TargetDate.ToString('yyyy-MM-dd')) $($state.project) ¥$TargetAmount；为避免重复记账，本次未写入"
    }
    if ($state.ledger_match_count -eq 1 -and $state.month_note_count -eq 1) {
        return [pscustomobject]@{
            status = 'duplicate'
            changed = $false
            date = $TargetDate.ToString('yyyy-MM-dd')
            project = $state.project
            amount = $TargetAmount
            expense_sheet = $state.sheet_name
            expense_row = [int]$state.ledger_matches[0].row_number
            month_row = [int]$state.month.row_number
        }
    }

    $targetRow = Get-FirstCompletelyEmptyRow $state.ledger_rows 6 2 1001
    if ($PreviewOnly) {
        return [pscustomobject]@{
            status = 'ready'
            changed = $false
            date = $TargetDate.ToString('yyyy-MM-dd')
            project = $state.project
            amount = $TargetAmount
            expense_sheet = $state.sheet_name
            expense_row = $targetRow
            month_row = [int]$state.month.row_number
        }
    }

    Set-LedgerOneRow $state.sheet_name "A${targetRow}:F${targetRow}" @(
        (Convert-ToExcelSerial $TargetDate),
        $state.project,
        $TargetAmount,
        $null,
        $ExpenseSource,
        $ExpenseMemo
    )

    $currentMonthlyValue = Convert-ToNumber (Get-CellValue $state.month.cells 6)
    $newMonthlyValue = [math]::Round($currentMonthlyValue + $TargetAmount, 2)
    Set-OneRow '月度结算' "G$($state.month.row_number):G$($state.month.row_number)" @($newMonthlyValue)
    $oldNote = Get-CellValue $state.month.cells 11
    $newNote = if ($oldNote) { "$oldNote；$($state.month_note_entry)" } else { $state.month_note_entry }
    Set-OneRow '月度结算' "L$($state.month.row_number):L$($state.month.row_number)" @($newNote)

    $check = Get-OperatingExpenseState $TargetDate $TargetAmount $TargetNote
    $actualMonthlyValue = Convert-ToNumber (Get-CellValue $check.month.cells 6)
    if ($check.ledger_match_count -ne 1 -or $check.month_note_count -ne 1 -or [math]::Abs($actualMonthlyValue - $newMonthlyValue) -gt 0.0001) {
        throw "其他经营费用写入后回读校验失败：$($TargetDate.ToString('yyyy-MM-dd')) $($state.project) ¥$TargetAmount"
    }

    return [pscustomobject]@{
        status = 'recorded'
        changed = $true
        date = $TargetDate.ToString('yyyy-MM-dd')
        project = $state.project
        amount = $TargetAmount
        expense_sheet = $state.sheet_name
        expense_row = $targetRow
        month_row = [int]$state.month.row_number
        month_value = $newMonthlyValue
    }
}

function Get-BatchArgument($Arguments, [string]$Name) {
    if ($null -eq $Arguments) { return $null }
    $property = $Arguments.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Copy-CellRow($Row, [int]$Width) {
    $copy = [object[]]::new($Width)
    for ($index = 0; $index -lt $Width; $index++) {
        $copy[$index] = if ($null -ne $Row -and $index -lt @($Row).Count -and $null -ne @($Row)[$index]) { @($Row)[$index] } else { @{} }
    }
    return ,$copy
}

function Set-WorkingRow([Collections.ArrayList]$Rows, [int]$RowNumber, [object[]]$Cells) {
    $index = $RowNumber - 2
    if ($index -lt 0) { throw "无效的工作表行号：$RowNumber" }
    while ($Rows.Count -le $index) { $null = $Rows.Add([object[]]@()) }
    $Rows[$index] = $Cells
}

function Convert-BatchDate($Value, [datetime]$Fallback, [string]$Name) {
    if ($null -eq $Value -or -not ([string]$Value).Trim()) { return $Fallback.Date }
    $parsed = [datetime]::MinValue
    if (-not [datetime]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)) {
        throw "$Name 不是有效日期：$Value"
    }
    return $parsed.Date
}

function Get-BatchOperationsPayload {
    $json = $OperationsJson
    if ($OperationsPath) {
        if (-not (Test-Path -LiteralPath $OperationsPath -PathType Leaf)) { throw "找不到批量经营操作文件：$OperationsPath" }
        $json = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $OperationsPath).Path,[Text.Encoding]::UTF8)
    }
    if (-not $json) { throw '批量 SKU 维护必须提供 operations JSON' }
    try { $operations = @($json | ConvertFrom-Json) } catch { throw "批量 SKU 维护参数不是有效 JSON：$($_.Exception.Message)" }
    if ($operations.Count -eq 0) { throw '批量 SKU 维护没有可执行项目' }
    return $operations
}

function Test-NullableNumberEqual($Actual, $Expected) {
    $actualNumber = Convert-ToNullableNumber $Actual
    if ($null -eq $Expected) { return $null -eq $actualNumber }
    return $null -ne $actualNumber -and [math]::Abs([double]$actualNumber - [double]$Expected) -le 0.0001
}

function Invoke-BatchSkuMaintenance([object[]]$Operations) {
    # Compatible ordinary SKU actions share these two reads and one atomic write.
    # Historical-order corrections are deliberately rejected here and keep their existing path.
    $skuRows = [Collections.ArrayList]::new()
    foreach ($row in @(Get-SheetRows 'SKU成本主表' 'A2:Z3000')) { $null = $skuRows.Add((Copy-CellRow $row 26)) }
    $replenishmentRows = [Collections.ArrayList]::new()
    foreach ($row in @(Get-SheetRows '补货记录' 'A2:T1001')) { $null = $replenishmentRows.Add((Copy-CellRow $row 20)) }
    $writes = [Collections.Generic.List[object]]::new()
    $results = [Collections.Generic.List[object]]::new()
    $expectedSku = @{}
    $expectedReplenishment = @{}
    $formulaRows = [Collections.Generic.HashSet[int]]::new()

    for ($operationIndex = 0; $operationIndex -lt $Operations.Count; $operationIndex++) {
        $operation = $Operations[$operationIndex]
        $operationAction = [string](Get-BatchArgument $operation 'action')
        $arguments = Get-BatchArgument $operation 'arguments'
        if ($operationAction -notin @('Replenishment','SetSkuCost')) { throw "批量 SKU 维护不支持动作：$operationAction" }
        if ($operationAction -eq 'SetSkuCost' -and (Get-BatchArgument $arguments 'order_no')) { throw '指定历史订单的成本修正不能与普通 SKU 维护合并执行' }

        if ($operationAction -eq 'SetSkuCost') {
            $goodsCode = [string](Get-BatchArgument $arguments 'goods_code')
            $amountValue = Convert-ToNullableNumber (Get-BatchArgument $arguments 'amount')
            if (-not $goodsCode -or $goodsCode.Trim().Length -gt 100) { throw "第$($operationIndex+1)项成本修改缺少可唯一定位的货号" }
            Assert-NonNegative $amountValue "第$($operationIndex+1)项销售单位成本"
            $operationDate = Convert-BatchDate (Get-BatchArgument $arguments 'date') $Date '成本确认日期'
            $target = Resolve-UniqueSkuReference $skuRows $goodsCode.Trim().ToLowerInvariant()
            $rowNumber = [int]$target.row_number
            $cells = Copy-CellRow $skuRows[$rowNumber-2] 26
            $canonicalKw = Get-CellValue $cells $skuSchema.CurrentKw.Index
            $transitionCode = Get-CellValue $cells $skuSchema.TransitionCode.Index
            $productId = Get-CellValue $cells $skuSchema.ProductId.Index
            $displayCode = if ($canonicalKw) { $canonicalKw.Trim().ToLowerInvariant() } elseif ($transitionCode) { @([regex]::Split($transitionCode,'[/,，、;；\s]+') | Where-Object { $_ } | Select-Object -First 1)[0].Trim().ToLowerInvariant() } elseif ($productId) { "SKU ID $productId" } else { $goodsCode.Trim().ToLowerInvariant() }
            $oldBaseline = Convert-ToNullableNumber (Get-CellValue $cells $skuSchema.BaselineUnitCost.Index)
            $oldMoving = Convert-ToNullableNumber (Get-CellValue $cells $skuSchema.MovingAverageCost.Index)
            $oldStatus = Get-CellValue $cells $skuSchema.CostStatus.Index
            $newStatus = if ([string]$oldStatus -match '当前KW冲突') { '当前KW冲突待核对（成本已确认）' } else { '用户确认' }
            $newCost = [math]::Round([double]$amountValue,4)
            $noteValue = [string](Get-BatchArgument $arguments 'note')
            $noteEntry = "店主通过飞书确认销售单位成本为 $($newCost.ToString('0.####',[Globalization.CultureInfo]::InvariantCulture)) 元（$($operationDate.ToString('yyyy-MM-dd'))）"
            $newNote = ((@([regex]::Split([string](Get-CellValue $cells $skuSchema.Note.Index),'[；;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) + @($noteEntry,$noteValue) | Where-Object { $_ } | Select-Object -Unique) -join '；')

            $cells[$skuSchema.BaselineUnitCost.Index] = New-UnitCostCell $newCost
            $cells[$skuSchema.MovingAverageCost.Index] = New-UnitCostCell $newCost
            $cells[$skuSchema.CostStatus.Index] = New-ValueCell $newStatus
            $cells[$skuSchema.LastCostUpdate.Index] = New-ValueCell (Convert-ToExcelSerial ([datetime]::Now))
            $cells[$skuSchema.ConversionBasis.Index] = New-ValueCell '店主确认的销售单位成本'
            $cells[$skuSchema.Note.Index] = New-ValueCell $newNote
            Set-WorkingRow $skuRows $rowNumber $cells

            $costCells = @(); $costCells += ,@((New-UnitCostCell $newCost),(New-UnitCostCell $newCost))
            $statusCells = @(); $statusCells += ,@((New-ValueCell $newStatus))
            $updatedCells = @(); $updatedCells += ,@((New-ValueCell (Convert-ToExcelSerial ([datetime]::Now))))
            $noteCells = @(); $noteCells += ,@((New-ValueCell '店主确认的销售单位成本'),(New-ValueCell $newNote))
            $writes.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.BaselineUnitCost.Column)${rowNumber}:$($skuSchema.MovingAverageCost.Column)${rowNumber}";cells=$costCells}})
            $writes.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.CostStatus.Column)${rowNumber}:$($skuSchema.CostStatus.Column)${rowNumber}";cells=$statusCells}})
            $writes.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.LastCostUpdate.Column)${rowNumber}:$($skuSchema.LastCostUpdate.Column)${rowNumber}";cells=$updatedCells}})
            $writes.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.ConversionBasis.Column)${rowNumber}:$($skuSchema.Note.Column)${rowNumber}";cells=$noteCells}})
            $expectedSku[[string]$rowNumber] = [pscustomobject]@{baseline=$newCost;moving=$newCost;quantity=Convert-ToNullableNumber (Get-CellValue $cells $skuSchema.CostBookQuantity.Index);status=$newStatus}
            $results.Add([pscustomobject]@{sheet='SKU成本主表';row=$rowNumber;range="$($skuSchema.BaselineUnitCost.Column)${rowNumber}:$($skuSchema.MovingAverageCost.Column)${rowNumber},$($skuSchema.CostStatus.Column)${rowNumber},$($skuSchema.LastCostUpdate.Column)${rowNumber}:$($skuSchema.Note.Column)${rowNumber}";goods_code=$displayCode;sku_reference=$goodsCode.Trim().ToLowerInvariant();product_id=$productId;title=Get-CellValue $cells 2;old_baseline_cost=$oldBaseline;old_moving_average_cost=$oldMoving;new_unit_cost=$newCost;new_moving_average_cost=$newCost;old_status=$oldStatus;new_status=$newStatus;historical_order_no=$null;affected_tables=@('SKU成本主表')})
            continue
        }

        $goodsCode = [string](Get-BatchArgument $arguments 'goods_code')
        $transitionCode = [string](Get-BatchArgument $arguments 'transition_code')
        $quantityValue = Convert-ToNullableNumber (Get-BatchArgument $arguments 'quantity')
        $purchaseAmountValue = Convert-ToNullableNumber (Get-BatchArgument $arguments 'purchase_amount')
        $paymentAmountValue = Convert-ToNullableNumber (Get-BatchArgument $arguments 'payment_amount')
        if (-not $goodsCode -or $goodsCode -notmatch '(?i)^(?:kw|csx)\d+$') { throw "第$($operationIndex+1)项补货必须提供有效的 KW 或 CSX 货号" }
        if ($transitionCode -and $transitionCode -notmatch '(?i)^csx\d+$') { throw "第$($operationIndex+1)项补货关联的过渡货号无效" }
        Assert-Positive $quantityValue "第$($operationIndex+1)项补货数量"
        Assert-NonNegative $purchaseAmountValue "第$($operationIndex+1)项实际采购金额"
        if ($null -ne $paymentAmountValue) { Assert-NonNegative $paymentAmountValue "第$($operationIndex+1)项付款金额" }
        $operationDate = Convert-BatchDate (Get-BatchArgument $arguments 'date') $Date '补货日期'
        $normalizedCode = $goodsCode.Trim().ToLowerInvariant()
        $normalizedTransition = if ($transitionCode) { $transitionCode.Trim().ToLowerInvariant() } else { $null }
        $requestedKw = if ($normalizedCode -match '^kw') { $normalizedCode } else { $null }
        $requestedCsx = if ($normalizedTransition) { $normalizedTransition } elseif ($normalizedCode -match '^csx') { $normalizedCode } else { $null }
        $candidateRows = @{}
        foreach ($lookupCode in @($normalizedCode,$normalizedTransition | Where-Object { $_ } | Select-Object -Unique)) {
            $matches = @(Get-SkuCodeCandidates $skuRows $lookupCode)
            if ($matches.Count -gt 1) { throw "SKU成本主表中的货号 $lookupCode 重复出现在第 $(@($matches|ForEach-Object{$_.row_number}) -join '、') 行；批量补货未写入" }
            if ($matches.Count -eq 1) { $candidateRows[[string]$matches[0].row_number] = $matches[0] }
        }
        if ($candidateRows.Count -gt 1) { throw "第$($operationIndex+1)项提供的 KW/CSX 指向不同 SKU 行；批量操作未写入" }
        $newSku = $candidateRows.Count -eq 0
        $mappingUpdate = $false
        if ($newSku) {
            $skuRowNumber = Get-FirstCompletelyEmptyRow $skuRows 26 2 3000
            $cells = Copy-CellRow $null 26
            $currentKw = $requestedKw
            $productId = $null
            $unitsPerPack = 1.0
        } else {
            $target = @($candidateRows.Values)[0]
            $skuRowNumber = [int]$target.row_number
            $cells = Copy-CellRow $skuRows[$skuRowNumber-2] 26
            $existingKw = Get-CellValue $cells $skuSchema.CurrentKw.Index
            $existingCsx = Get-CellValue $cells $skuSchema.TransitionCode.Index
            if ($requestedKw) {
                $kwMatches = @(Get-SkuCodeCandidates $skuRows $requestedKw)
                if ($kwMatches.Count -eq 1 -and [int]$kwMatches[0].row_number -ne $skuRowNumber) { throw "当前 KW $requestedKw 已属于 SKU成本主表第 $($kwMatches[0].row_number) 行" }
                $mappingUpdate = -not $existingKw -or $existingKw.Trim().ToLowerInvariant() -ne $requestedKw
            }
            if ($requestedCsx -and $existingCsx -and $existingCsx.Trim().ToLowerInvariant() -ne $requestedCsx) { throw "SKU成本主表第 $skuRowNumber 行已有过渡货号 $existingCsx，与 $requestedCsx 不一致" }
            $currentKw = if ($requestedKw) { $requestedKw } else { $existingKw }
            $productId = Get-CellValue $cells $skuSchema.ProductId.Index
            $unitsPerPack = Convert-ToNumber (Get-CellValue $cells $skuSchema.UnitsPerPack.Index)
            if ($unitsPerPack -le 0) { $unitsPerPack = 1.0 }
        }
        $productionRaw = Get-BatchArgument $arguments 'production_date'
        $shelfLife = Convert-ToNullableNumber (Get-BatchArgument $arguments 'shelf_life_value')
        $shelfUnit = [string](Get-BatchArgument $arguments 'shelf_life_unit'); if (-not $shelfUnit) { $shelfUnit = '月' }
        if (($null -ne $productionRaw) -xor ($null -ne $shelfLife)) { throw "第$($operationIndex+1)项补货的生产日期和保质期必须一起提供" }
        $productionDateValue = $null; $expiry = $null
        if ($null -ne $productionRaw) { $productionDateValue = Convert-BatchDate $productionRaw $operationDate '生产日期'; Assert-Positive $shelfLife '保质期'; $expiry = Add-ShelfLife $productionDateValue ([double]$shelfLife) $shelfUnit }
        $salesQuantity = [double]$quantityValue * $unitsPerPack
        $incomingUnitCost = [double]$purchaseAmountValue / $salesQuantity
        $beforeQuantity = if ($newSku) { 0.0 } else { Convert-ToNullableNumber (Get-CellValue $cells $skuSchema.CostBookQuantity.Index) }; if ($null -eq $beforeQuantity) { $beforeQuantity = 0.0 }
        $beforeAverage = if ($newSku) { $null } else { Convert-ToNullableNumber (Get-CellValue $cells $skuSchema.MovingAverageCost.Index) }
        $afterQuantity = [double]$beforeQuantity + $salesQuantity
        $afterAverage = if ($null -eq $beforeAverage -and [double]$beforeQuantity -gt 0) { $null } else { ((if($null-eq$beforeAverage){0.0}else{[double]$beforeAverage*[double]$beforeQuantity}) + [double]$purchaseAmountValue) / $afterQuantity }
        $noteValue = [string](Get-BatchArgument $arguments 'note')
        if ($newSku) {
            $zeroCost = [double]$purchaseAmountValue -eq 0
            $cells[$skuSchema.CurrentKw.Index] = New-ValueCell $requestedKw; $cells[$skuSchema.TransitionCode.Index] = New-ValueCell $requestedCsx
            $cells[$skuSchema.InitialInventory.Index] = New-ValueCell 0; $cells[$skuSchema.BaselineUnitCost.Index] = New-ValueCell $(if($zeroCost){0.0}else{$null})
            $cells[$skuSchema.UnitsPerPack.Index] = New-ValueCell 1; $cells[$skuSchema.PurchaseUnit.Index] = New-ValueCell '个'
            $cells[$skuSchema.CostStatus.Index] = New-ValueCell $(if($zeroCost){'零成本'}else{'新SKU待核对'})
            $cells[$skuSchema.InitializationDate.Index] = New-ValueCell (Convert-ToExcelSerial $operationDate)
            $cells[$skuSchema.ConversionBasis.Index] = New-ValueCell $(if($zeroCost){'用户明确试吃装/赠品成本为0'}else{'新SKU按本次补货建立成本，采购包装换算默认1'})
            $cells[$skuSchema.Note.Index] = New-ValueCell $(if($noteValue){"新货号待平台库存同步补全商品资料；$noteValue"}else{'新货号待平台库存同步补全商品资料'})
            $cells[$skuSchema.EarliestExpiry.Index] = New-ValueCell $(if($expiry){Convert-ToExcelSerial $expiry}else{$null}); $cells[$skuSchema.PlatformStatus.Index] = New-ValueCell '待平台同步'
            $matrix = @(); $matrix += ,$cells; $writes.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="A${skuRowNumber}:Z${skuRowNumber}";cells=$matrix}})
        } elseif ($mappingUpdate) {
            $cells[$skuSchema.CurrentKw.Index] = New-ValueCell $requestedKw
            $matrix = @(); $matrix += ,@((New-ValueCell $requestedKw)); $writes.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="A${skuRowNumber}:A${skuRowNumber}";cells=$matrix}})
        }
        $cells[$skuSchema.MovingAverageCost.Index] = New-UnitCostCell $afterAverage; $cells[$skuSchema.CostBookQuantity.Index] = New-UnitCostCell $afterQuantity; $cells[$skuSchema.LastCostUpdate.Index] = New-ValueCell (Convert-ToExcelSerial ([datetime]::Now))
        Set-WorkingRow $skuRows $skuRowNumber $cells
        $matrix = @(); $matrix += ,@((New-UnitCostCell $afterAverage),(New-UnitCostCell $afterQuantity)); $writes.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.MovingAverageCost.Column)${skuRowNumber}:$($skuSchema.CostBookQuantity.Column)${skuRowNumber}";cells=$matrix}})
        $matrix = @(); $matrix += ,@((New-ValueCell (Convert-ToExcelSerial ([datetime]::Now)))); $writes.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.LastCostUpdate.Column)${skuRowNumber}:$($skuSchema.LastCostUpdate.Column)${skuRowNumber}";cells=$matrix}})
        $rowNumber = Get-FirstEmptyRow $replenishmentRows 1 2 1001
        $operationEventId = [string](Get-BatchArgument $arguments 'event_id'); if (-not $operationEventId) { $operationEventId = "$EventId-$operationIndex-Replenishment" }
        $batchId = 'EVT-' + (Get-StableEventKey $operationEventId)
        $paymentDateValue = if ($null-ne$paymentAmountValue) { Convert-ToExcelSerial (Convert-BatchDate (Get-BatchArgument $arguments 'payment_date') $operationDate '付款日期') } else { $null }
        $displayCode = if ($currentKw) { $currentKw } else { $normalizedCode }
        $rowCells = Copy-CellRow $null 20
        $rowCells[0]=New-ValueCell (Convert-ToExcelSerial $operationDate);$rowCells[1]=New-ValueCell $displayCode;$rowCells[2]=New-ValueCell ([double]$quantityValue);$rowCells[3]=New-ValueCell $unitsPerPack
        $rowCells[4]=@{formula="=IF(OR(C${rowNumber}=`"`",D${rowNumber}=`"`"),`"`",C${rowNumber}*D${rowNumber})"};$rowCells[5]=New-ValueCell ([double]$purchaseAmountValue);$rowCells[6]=@{formula="=IF(OR(E${rowNumber}=`"`",F${rowNumber}=`"`"),`"`",F${rowNumber}/E${rowNumber})"}
        $rowCells[7]=New-UnitCostCell $beforeQuantity;$rowCells[8]=New-UnitCostCell $beforeAverage;$rowCells[9]=New-UnitCostCell $afterQuantity;$rowCells[10]=New-UnitCostCell $afterAverage
        $rowCells[11]=New-ValueCell $paymentDateValue;$rowCells[12]=New-ValueCell $paymentAmountValue;$rowCells[13]=New-ValueCell $noteValue;$rowCells[14]=New-ValueCell $productId
        $rowCells[15]=New-ValueCell $(if($productionDateValue){Convert-ToExcelSerial $productionDateValue}else{$null});$rowCells[16]=New-ValueCell $shelfLife;$rowCells[17]=New-ValueCell $(if($productionDateValue){$shelfUnit}else{$null});$rowCells[18]=New-ValueCell $(if($expiry){Convert-ToExcelSerial $expiry}else{$null});$rowCells[19]=New-ValueCell $batchId
        Set-WorkingRow $replenishmentRows $rowNumber $rowCells
        $matrix=@();$matrix+=,$rowCells;$writes.Add(@{shortcut='+cells-set';input=@{sheet_name='补货记录';range="A${rowNumber}:T${rowNumber}";cells=$matrix}})
        $expectedSku[[string]$skuRowNumber]=[pscustomobject]@{baseline=Convert-ToNullableNumber (Get-CellValue $cells $skuSchema.BaselineUnitCost.Index);moving=$afterAverage;quantity=$afterQuantity;status=Get-CellValue $cells $skuSchema.CostStatus.Index}
        $expectedReplenishment[[string]$rowNumber]=[pscustomobject]@{goods_code=$displayCode;quantity=[double]$quantityValue;purchase_amount=[double]$purchaseAmountValue;batch_id=$batchId;after_quantity=$afterQuantity;after_average=$afterAverage};$null=$formulaRows.Add($rowNumber)
        $results.Add([pscustomobject]@{row=$rowNumber;sku_master_row=$skuRowNumber;goods_code=$displayCode;transition_code=$requestedCsx;product_id=$productId;package_quantity=[double]$quantityValue;sales_unit_quantity=$salesQuantity;purchase_amount=[double]$purchaseAmountValue;payment_amount=$paymentAmountValue;unit_cost=$incomingUnitCost;before_quantity=$beforeQuantity;before_average_cost=$beforeAverage;after_quantity=$afterQuantity;after_average_cost=$afterAverage;production_date=$(if($productionDateValue){$productionDateValue.ToString('yyyy-MM-dd')}else{$null});expiry_date=$(if($expiry){$expiry.ToString('yyyy-MM-dd')}else{$null});sku_status=$(if($newSku){$(if($DryRun){'will_create_pending_profile'}else{'created_pending_profile'})}elseif($mappingUpdate){$(if($DryRun){'will_attach_kw_by_csx'}else{'attached_kw_by_csx'})}else{'matched'});affected_tables=@('补货记录','SKU成本主表');remote_replay=$false})
    }

    if ($DryRun) {
        return [pscustomobject]@{operations=$results.ToArray();operation_count=$results.Count;write_operation_count=0;master_read_count=1;replenishment_read_count=1;preview=$true}
    }
    Set-AtomicCellOperations $writes.ToArray()
    foreach($entry in $expectedSku.GetEnumerator()) { $row=[int]$entry.Key;$e=$entry.Value;$a=Get-SingleSheetRow 'SKU成本主表' "A${row}:Z${row}";if($null-eq$a -or -not(Test-NullableNumberEqual (Get-CellValue $a $skuSchema.BaselineUnitCost.Index) $e.baseline) -or -not(Test-NullableNumberEqual (Get-CellValue $a $skuSchema.MovingAverageCost.Index) $e.moving) -or -not(Test-NullableNumberEqual (Get-CellValue $a $skuSchema.CostBookQuantity.Index) $e.quantity) -or [string](Get-CellValue $a $skuSchema.CostStatus.Index)-ne[string]$e.status){throw "批量 SKU 维护写后校验失败：SKU成本主表第 $row 行"} }
    foreach($entry in $expectedReplenishment.GetEnumerator()) { $row=[int]$entry.Key;$e=$entry.Value;$a=Get-SingleSheetRow '补货记录' "A${row}:T${row}";if($null-eq$a -or [string](Get-CellValue $a 1)-ne[string]$e.goods_code -or -not(Test-NullableNumberEqual (Get-CellValue $a 2) $e.quantity) -or -not(Test-NullableNumberEqual (Get-CellValue $a 5) $e.purchase_amount) -or [string](Get-CellValue $a 19)-ne[string]$e.batch_id -or -not(Test-NullableNumberEqual (Get-CellValue $a 9) $e.after_quantity) -or -not(Test-NullableNumberEqual (Get-CellValue $a 10) $e.after_average)){throw "批量补货写后校验失败：补货记录第 $row 行"} }
    if($formulaRows.Count-gt 0){$ranges=@($formulaRows|Sort-Object|ForEach-Object{"D${_}:E${_},G${_}:G${_}"})-join',';Assert-FormulaRange '补货记录' $ranges}
    return [pscustomobject]@{operations=$results.ToArray();operation_count=$results.Count;write_operation_count=$writes.Count;master_read_count=1;replenishment_read_count=1}
}

if (-not $DryRun -and $EventId -and (Test-ProcessedEvent $EventId)) {
    [ordered]@{ok=$true;duplicate=$true;event_id=$EventId;action=$Action;message='该事件已处理，本次未重复写入'} | ConvertTo-Json -Depth 6
    Complete-CurrentFinanceTransaction
    exit 0
}

$nativeDryRunActions = @('AddIntervention','RemoveIntervention','SetOrderDeliveryFee','SetOrderActuals','OfflineSale','AddOrderGift','Replenishment','BatchSkuMaintenance','SetSkuCost','SetSkuExpiry','SetSkuBatchDate','ChangeSkuCode','OperatingExpense','SetRefundDisposition')
if ($DryRun -and $Action -notin $nativeDryRunActions) {
    throw "动作 $Action 暂不支持直接预演；已在任何写入前停止"
}

if (-not $DryRun) {
    $script:financeTransaction = Join-FinanceTransaction $stateRoot
    if ($null -eq $script:financeTransaction) {
        $transactionId = if ($EventId) { "event-$EventId" } else { $null }
        $script:financeTransaction = New-FinanceTransaction $stateRoot "$Action $EventId" $transactionId -AcquireLock -WorkbookTokens (Get-FinanceWorkbookTokens) -LarkCli $larkCli
        $script:financeTransactionOwner = $true
    }
}

$reference = ''
$writeResult = $null
$costResult = $null
$explicitHistoricalOrderDate = $null
switch ($Action) {
    'BatchSkuMaintenance' {
        $batchOperations = @(Get-BatchOperationsPayload)
        $batchResult = Invoke-BatchSkuMaintenance $batchOperations
        if ([int]$batchResult.operation_count -ne $batchOperations.Count) {
            throw "批量 SKU 维护结果不完整：计划 $($batchOperations.Count) 项，实际 $($batchResult.operation_count) 项"
        }
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=$batchResult;message='批量预演成功，未写入飞书'} | ConvertTo-Json -Depth 12
            Complete-CurrentFinanceTransaction
            exit 0
        }
        $reference = "batch-sku-maintenance x$($batchOperations.Count)"
        $writeResult = $batchResult
    }
    'SetOrderActuals' {
        if (-not $OrderNo -or $OrderNo -notmatch '^\d{10,30}$') { throw '修正订单实际履约必须提供有效平台订单号' }
        Assert-NonNegative $Amount '订单实际收入'
        if ($null -ne $DeliveryFee) { Assert-NonNegative $DeliveryFee '订单实际配送费' }
        if (-not $Items) { throw '修正订单实际履约必须提供实际售出商品，例如 kw3001×1' }
        $actualItemMatches = @([regex]::Matches($Items, '(?i)\b((?:kw|csx)\d+)\s*[×xX*]\s*(\d+(?:\.\d+)?)'))
        if ($actualItemMatches.Count -eq 0) { throw '订单实际履约修正必须明确提供最终实际交付商品，例如 kw1096×2、kw1095×1；漏送或退回商品不要写入' }
        $requestedQuantities = [ordered]@{}
        foreach ($match in $actualItemMatches) {
            $requestedCode = $match.Groups[1].Value.ToLowerInvariant()
            $actualQuantity = [double]::Parse($match.Groups[2].Value, [Globalization.CultureInfo]::InvariantCulture)
            if ($actualQuantity -le 0) { throw "实际售出数量必须大于0：$requestedCode" }
            if (-not $requestedQuantities.Contains($requestedCode)) { $requestedQuantities[$requestedCode] = 0.0 }
            $requestedQuantities[$requestedCode] = [double]$requestedQuantities[$requestedCode] + $actualQuantity
        }

        $orderRows = Get-SheetRows '订单汇总' 'A2:Q2001'
        $orderMatches = [Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $orderRows.Count; $i++) {
            if ((Get-CellValue $orderRows[$i] 2) -eq $OrderNo) { $orderMatches.Add([pscustomobject]@{row_number=2+$i;cells=$orderRows[$i]}) }
        }
        if ($orderMatches.Count -ne 1) { throw "订单汇总中无法唯一定位订单：$OrderNo（匹配 $($orderMatches.Count) 行）" }
        $orderDate = Convert-ToDateValue (Get-CellValue $orderMatches[0].cells 0)
        if ($null -eq $orderDate) { throw "订单 $OrderNo 的日期无法识别" }
        if ((Get-CellValue $orderMatches[0].cells 1) -ne '美团') { throw '订单实际履约覆盖目前仅支持美团线上订单' }

        $originalItems = [Collections.Generic.List[string]]::new()
        $originalItemRecords = [Collections.Generic.List[object]]::new()
        foreach ($row in (Get-SheetRows '订单商品' 'A2:Q4001')) {
            if ((Get-CellValue $row 2) -ne $OrderNo) { continue }
            $code = Get-CellValue $row 5
            $qty = Convert-ToNumber (Get-CellValue $row 8)
            $qtyText = if ([math]::Abs($qty - [math]::Round($qty)) -lt 0.000001) { [string][int][math]::Round($qty) } else { $qty.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture) }
            $originalItems.Add("$code×$qtyText")
            $originalItemRecords.Add([pscustomobject]@{
                code = ([string]$code).Trim().ToLowerInvariant()
                title = Get-CellValue $row 6
                spec = Get-CellValue $row 7
                quantity = $qty
                unit_price = Convert-ToNumber (Get-CellValue $row 9)
                product_id = ([string](Get-CellValue $row 16)).Trim()
            })
        }
        if ($originalItems.Count -eq 0) { throw "订单商品中找不到订单 $OrderNo 的原始商品；请先同步订单" }

        $skuRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $actualIncome = [double][math]::Round([double]$Amount, 2)
        $oldIncome = Convert-ToNumber (Get-CellValue $orderMatches[0].cells 7)
        $oldDeliveryFee = Convert-ToNumber (Get-CellValue $orderMatches[0].cells 9)
        $oldCost = Convert-ToNumber (Get-CellValue $orderMatches[0].cells 10)
        $actualDeliveryFee = if ($null -eq $DeliveryFee) { $oldDeliveryFee } else { [double][math]::Round([double]$DeliveryFee, 2) }
        $resolvedRequestedItems = [Collections.Generic.List[object]]::new()
        foreach ($requested in $requestedQuantities.GetEnumerator()) {
            $requestedCode = [string]$requested.Key
            $actualQuantity = [double]$requested.Value
            $skuTarget = Resolve-UniqueSkuReference $skuRows $requestedCode
            $canonicalKw = (Get-CellValue $skuTarget.cells $skuSchema.CurrentKw.Index)
            if (-not $canonicalKw -or $canonicalKw -notmatch '(?i)^kw\d+$') { throw "SKU成本主表第$($skuTarget.row_number)行没有有效当前KW，无法登记实际售出商品" }
            $canonicalKw = $canonicalKw.ToLowerInvariant()
            $title = Get-CellValue $skuTarget.cells $skuSchema.CustomTitle.Index
            if (-not $title) { $title = Get-CellValue $skuTarget.cells $skuSchema.PlatformTitle.Index }
            $spec = Get-CellValue $skuTarget.cells $skuSchema.CurrentSpec.Index
            $productId = ([string](Get-CellValue $skuTarget.cells $skuSchema.ProductId.Index)).Trim()
            $unitCostText = Get-CellValue $skuTarget.cells $skuSchema.MovingAverageCost.Index
            if ($null -eq $unitCostText -or $unitCostText -eq '') { $unitCostText = Get-CellValue $skuTarget.cells $skuSchema.BaselineUnitCost.Index }
            if ($null -eq $unitCostText -or $unitCostText -eq '') { throw "$canonicalKw 当前成本为空，不能安全重算该订单" }
            $originalMatch = @($originalItemRecords | Where-Object {
                $_.code -ieq $requestedCode -or $_.code -ieq $canonicalKw -or ($productId -and $_.product_id -eq $productId)
            } | Select-Object -First 1)
            $resolvedRequestedItems.Add([pscustomobject]@{
                requested_code = $requestedCode
                canonical_kw = $canonicalKw
                title = $title
                spec = $spec
                product_id = $productId
                quantity = $actualQuantity
                unit_cost = Convert-ToNumber $unitCostText
                sku_master_row = $skuTarget.row_number
                original_item = $(if ($originalMatch.Count -gt 0) { $originalMatch[0] } else { $null })
            })
        }

        $newRequestedItems = @($resolvedRequestedItems | Where-Object { $null -eq $_.original_item })
        $allOriginalItemsPreserved = $true
        foreach ($originalRecord in $originalItemRecords) {
            $preservedQuantity = [double](($resolvedRequestedItems | Where-Object {
                $_.requested_code -ieq $originalRecord.code -or $_.canonical_kw -ieq $originalRecord.code -or
                ($originalRecord.product_id -and $_.product_id -eq $originalRecord.product_id)
            } | Measure-Object -Property quantity -Sum).Sum)
            if ([math]::Abs($preservedQuantity - [double]$originalRecord.quantity) -gt 0.000001) {
                $allOriginalItemsPreserved = $false
                break
            }
        }

        $actualItems = [Collections.Generic.List[object]]::new()
        $costDetails = [Collections.Generic.List[object]]::new()
        $lineIndex = 0
        foreach ($resolvedItem in $resolvedRequestedItems) {
            $lineIndex++
            $actualQuantity = [double]$resolvedItem.quantity
            if ($null -ne $resolvedItem.original_item) {
                $actualUnitPrice = [double]$resolvedItem.original_item.unit_price
                $unitPriceBasis = 'original_order_item'
            } elseif ($requestedQuantities.Count -eq 1) {
                $actualUnitPrice = [double][math]::Round($actualIncome / $actualQuantity, 4)
                $unitPriceBasis = 'single_item_actual_income'
            } elseif ($newRequestedItems.Count -eq 1 -and $allOriginalItemsPreserved) {
                $additionalIncome = [double][math]::Round($actualIncome - $oldIncome, 2)
                if ($additionalIncome -lt 0) {
                    throw "新增商品 $($resolvedItem.canonical_kw) 无法从收入差额安全定价：最终收入低于原商家收入；请明确该商品成交金额"
                }
                $actualUnitPrice = [double][math]::Round($additionalIncome / $actualQuantity, 4)
                $unitPriceBasis = 'additional_income_delta'
            } else {
                throw "实际交付包含多个替换商品，且 $($resolvedItem.canonical_kw) 不在原订单中；请分别说明各商品成交金额后再处理"
            }
            $actualItems.Add([pscustomobject]@{
                goods_code = $resolvedItem.canonical_kw
                title = $resolvedItem.title
                spec = $resolvedItem.spec
                product_id = $resolvedItem.product_id
                quantity = $actualQuantity
                unit_price = $actualUnitPrice
                unit_price_basis = $unitPriceBasis
                line_index = $lineIndex
            })
            $costDetails.Add([pscustomobject]@{
                goods_code = $resolvedItem.canonical_kw
                quantity = $actualQuantity
                unit_cost = $resolvedItem.unit_cost
                item_cost = [math]::Round([double]$resolvedItem.unit_cost * $actualQuantity, 2)
                sku_master_row = $resolvedItem.sku_master_row
                title = $resolvedItem.title
            })
        }
        $newCost = [math]::Round((($costDetails | Measure-Object -Property item_cost -Sum).Sum), 2)
        $newProfit = [math]::Round($actualIncome - $actualDeliveryFee - $newCost, 2)
        $newItemsText = @($actualItems | ForEach-Object {
            $qty = [double]$_.quantity
            $qtyText = if ([math]::Abs($qty - [math]::Round($qty)) -lt 0.000001) { [string][int][math]::Round($qty) } else { $qty.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture) }
            "$($_.goods_code)×$qtyText"
        }) -join '、'
        $details = [ordered]@{
            order_no = $OrderNo
            order_date = $orderDate.ToString('yyyy-MM-dd')
            daily_order_no = Get-CellValue $orderMatches[0].cells 16
            order_summary_row = $orderMatches[0].row_number
            old_income = $oldIncome
            new_income = $actualIncome
            old_delivery_fee = $oldDeliveryFee
            new_delivery_fee = $actualDeliveryFee
            old_items = @($originalItems)
            new_items = $newItemsText
            actual_items = @($actualItems)
            sku_master_rows = @($costDetails | ForEach-Object { $_.sku_master_row })
            cost_details = @($costDetails)
            old_cost = $oldCost
            expected_new_cost = $newCost
            expected_new_profit = $newProfit
            platform_inventory = '不修改；店主已手工恢复'
            affected_tables = @('订单实际履约覆盖','订单汇总','订单商品','对应月份收入','SKU成本主表成本账数量','每日经营简账')
        }
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=$details} | ConvertTo-Json -Depth 10
            Complete-CurrentFinanceTransaction
            exit 0
        }

        $script:orderActualOverrideStateBackupPath = Get-OrderActualOverrideStatePath $projectRoot
        if (Test-Path -LiteralPath $script:orderActualOverrideStateBackupPath) {
            $script:orderActualOverrideStateBackup = [IO.File]::ReadAllText($script:orderActualOverrideStateBackupPath, [Text.Encoding]::UTF8)
        } else {
            $script:orderActualOverrideStateBackup = (New-OrderActualOverrideState | ConvertTo-Json -Depth 10)
        }
        $fulfillmentNote = if ($Note) { $Note } else { '订单最终实收或实际履约由店主确认；平台库存不由本动作修改' }
        $null = Set-OrderActualOverride $projectRoot $OrderNo $orderDate.ToString('yyyy-MM-dd') $actualIncome $actualDeliveryFee @($actualItems) 'manual_order_adjustment' $fulfillmentNote
        $orderSyncResult = Invoke-OrderSyncForDate $orderDate
        $details['order_sync'] = $orderSyncResult
        $reference = "$OrderNo $newItemsText 实收$actualIncome"
        $writeResult = [pscustomobject]$details
    }
    'SetRefundDisposition' {
        if (-not $OrderNo -or $OrderNo -notmatch '^\d{10,30}$') { throw '确认退款商品状态必须提供有效平台订单号' }
        if (-not $GoodsCode -or $GoodsCode -notmatch '(?i)^(?:kw|csx)\d+$') { throw '确认退款商品状态必须提供有效 KW 或 CSX 货号' }
        if (-not $RefundDisposition) { throw '请明确退款商品是已退回入库、未退货，还是退回报损' }
        $events = @(Get-RefundEventsForOrder $projectRoot $OrderNo)
        $matched = @($events | ForEach-Object {
            $event = $_
            @($event.items | Where-Object { ([string]$_.goods_code).Trim().ToLowerInvariant() -eq $GoodsCode.Trim().ToLowerInvariant() } | ForEach-Object {
                [pscustomobject]@{event_id=$event.event_id;record_id=$event.record_id;order_no=$OrderNo;goods_code=$_.goods_code;title=$_.title;quantity=$_.quantity;refund_amount=$event.customer_refund_amount;old_disposition=$_.disposition;new_disposition=$RefundDisposition}
            })
        })
        if ($matched.Count -eq 0) { throw "退款记录中找不到订单 $OrderNo 的商品 $GoodsCode；请先同步美团订单" }
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{order_no=$OrderNo;goods_code=$GoodsCode.ToLowerInvariant();title=$matched[0].title;quantity=$matched[0].quantity;refund_amount=$matched[0].refund_amount;old_disposition=$matched[0].old_disposition;new_disposition=$RefundDisposition;affected_tables=@('退款补偿状态','订单商品成本','订单汇总','对应月份收入','SKU成本主表账面库存')}} | ConvertTo-Json -Depth 10
            Complete-CurrentFinanceTransaction
            exit 0
        }
        $script:refundStateBackupPath = Get-RefundAdjustmentStatePath $projectRoot
        $script:refundStateBackup = [IO.File]::ReadAllText($script:refundStateBackupPath, [Text.Encoding]::UTF8)
        $updated = @(Set-RefundItemDisposition $projectRoot $OrderNo $GoodsCode $RefundDisposition $Note)
        $reference = "$OrderNo $($GoodsCode.ToLowerInvariant()) $RefundDisposition"
        $writeResult = [pscustomobject]@{order_no=$OrderNo;goods_code=$GoodsCode.ToLowerInvariant();title=$matched[0].title;quantity=$matched[0].quantity;refund_amount=$matched[0].refund_amount;old_disposition=$matched[0].old_disposition;new_disposition=$RefundDisposition;updated_events=@($updated.event_id);affected_tables=@('退款补偿状态','订单商品成本','订单汇总','对应月份收入','SKU成本主表账面库存')}
    }
    'RemoveIntervention' {
        if (-not $OrderNo -or $OrderNo -notmatch '^\d{10,30}$') { throw '撤销干预必须提供有效平台订单号' }
        $orderRows = Get-SheetRows '订单汇总' 'A2:Q2001'
        $orderMatches = @($orderRows | Where-Object { (Get-CellValue $_ 2) -eq $OrderNo })
        if ($orderMatches.Count -ne 1) { throw "订单汇总中无法唯一定位订单：$OrderNo（匹配 $($orderMatches.Count) 行）" }

        $rows = Get-SheetRows '干预订单' 'A2:M1001'
        $matches = [Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $rows.Count; $i++) {
            if ((Get-CellValue $rows[$i] 3) -eq $OrderNo) {
                $matches.Add([pscustomobject]@{ index=$i; row_number=(2 + $i); cells=$rows[$i] })
            }
        }
        if ($matches.Count -eq 0) { throw "干预订单中找不到要撤销的订单：$OrderNo" }
        if ($matches.Count -gt 1) { throw "干预订单中发现 $($matches.Count) 条重复记录：$OrderNo；为避免误删，本次未修改" }

        $target = $matches[0]
        $details = [ordered]@{
            row = $target.row_number
            order_no = $OrderNo
            person = Get-CellValue $target.cells 2
            delivery_type = Get-CellValue $target.cells 5
            delivery_fee = Convert-ToNumber (Get-CellValue $target.cells 7)
            inventory_state = Get-CellValue $target.cells 9
            new_order_type = '自然'
        }
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=$details} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }

        $tombstoneOrderNo = Get-InterventionTombstoneOrderNo $OrderNo
        $oldNote = Get-CellValue $target.cells 12
        $reason = if ($Note) { $Note } else { '店主确认该订单不属于干预单' }
        $auditNote = "已撤销干预；原订单号：$OrderNo；原因：$reason"
        if ($oldNote) { $auditNote += "；原备注：$oldNote" }
        $deliveryType = Get-CellValue $target.cells 5
        if ($deliveryType -notin @('自配送','骑手配送')) { $deliveryType = '自配送' }
        $restoreState = Get-CellValue $target.cells 9
        if ($restoreState -notin @('待恢复','已恢复')) { $restoreState = '已恢复' }
        $reviewState = Get-CellValue $target.cells 10
        $tombstoneRow = @(); $tombstoneRow += ,@(
            (New-ValueCell (Get-CellValue $target.cells 0)),
            (New-ValueCell (Get-CellValue $target.cells 1)),
            (New-ValueCell (Get-CellValue $target.cells 2)),
            (New-ValueCell $tombstoneOrderNo),
            (New-ValueCell 0),
            (New-ValueCell $deliveryType),
            (New-ValueCell 0),
            (New-ValueCell 0),
            @{formula="=IF(D$($target.row_number)=`"`",`"`",IFERROR(G$($target.row_number)*1,0)+IFERROR(H$($target.row_number)*1,0))"},
            (New-ValueCell $restoreState),
            (New-ValueCell $reviewState),
            (New-ValueCell (Get-CellValue $target.cells 11)),
            (New-ValueCell $auditNote)
        )
        Set-CellMatrix '干预订单' "A$($target.row_number):M$($target.row_number)" $tombstoneRow
        $withdrawn = $false
        for ($attempt = 1; $attempt -le 12; $attempt++) {
            $readback = Get-SingleSheetRow '干预订单' "A$($target.row_number):M$($target.row_number)"
            if ($null -ne $readback -and
                (Get-CellValue $readback 3) -eq $tombstoneOrderNo -and
                [math]::Abs((Convert-ToNumber (Get-CellValue $readback 6))) -le 0.0001 -and
                [math]::Abs((Convert-ToNumber (Get-CellValue $readback 7))) -le 0.0001) {
                $withdrawn = $true
                break
            }
            if ($attempt -lt 12) { Start-Sleep -Seconds 1 }
        }
        if (-not $withdrawn) {
            throw "撤销干预写后校验失败：$OrderNo 未在第 $($target.row_number) 行正确标记为已撤销"
        }
        $reference = $OrderNo
        $details['removal_mode'] = 'audit_tombstone'
        $details['tombstone_order_no'] = $tombstoneOrderNo
        $writeResult = [pscustomobject]$details
    }
    'AddIntervention' {
        if (-not $OrderNo -or $OrderNo -notmatch '^\d{10,30}$') { throw '干预单必须提供有效平台订单号' }
        $rows = Get-SheetRows '干预订单' 'A2:M1001'
        for ($i = 0; $i -lt $rows.Count; $i++) {
            if ((Get-CellValue $rows[$i] 3) -eq $OrderNo) {
                $existingRow = 2 + $i
                if ($DryRun) {
                    [ordered]@{ok=$true;dry_run=$true;duplicate=$true;action=$Action;result=[ordered]@{row=$existingRow;order_no=$OrderNo;person=(Get-CellValue $rows[$i] 2);state=(Get-CellValue $rows[$i] 9)};message='预演成功，现有干预记录不会被修改'} | ConvertTo-Json -Depth 8
                    Complete-CurrentFinanceTransaction
                    exit 0
                }
                if ($Person -and (Get-CellValue $rows[$i] 2) -ne $Person) {
                    Set-OneRow '干预订单' "C${existingRow}:C${existingRow}" @($Person)
                }
                [ordered]@{ok=$true;duplicate=$true;action=$Action;order_no=$OrderNo;row=$existingRow;person=$Person;message='该干预单已存在；已保留原记录并补充明确字段'} | ConvertTo-Json -Depth 6
                Complete-CurrentFinanceTransaction
                exit 0
            }
        }
        $orderRows = Get-SheetRows '订单汇总' 'A2:E2001'
        $exactOrder = @($orderRows | Where-Object { (Get-CellValue $_ 2) -eq $OrderNo })
        if ($exactOrder.Count -eq 0) {
            $nearbyOrders = @($orderRows | Where-Object {
                $candidate = Get-CellValue $_ 2
                $candidateDate = Convert-ToDateValue (Get-CellValue $_ 0)
                $candidateDate -and $candidateDate.Date -eq $Date.Date -and (Get-OrderNumberDistance $OrderNo $candidate) -eq 1
            } | ForEach-Object { Get-CellValue $_ 2 } | Select-Object -Unique)
            if ($nearbyOrders.Count -eq 1) {
                throw "干预订单号未匹配到真实订单：$OrderNo；发现仅差一位的订单 $($nearbyOrders[0])，请核对。系统未写入错误编号"
            }
        }
        $rowNumber = Get-FirstEmptyRow $rows 3 2 1001
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;duplicate=$false;action=$Action;result=[ordered]@{row=$rowNumber;order_no=$OrderNo;person=$Person;state='待恢复'};message='预演成功，未写入飞书'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
        $newRow = @(); $newRow += ,@(
            (New-ValueCell (Convert-ToExcelSerial $Date)),(New-ValueCell $null),(New-ValueCell $Person),(New-ValueCell $OrderNo),
            (New-ValueCell $null),(New-ValueCell $null),(New-ValueCell $null),(New-ValueCell $null),
            @{formula="=IF(D$rowNumber=`"`",`"`",IFERROR(G$rowNumber*1,0)+IFERROR(H$rowNumber*1,0))"},
            (New-ValueCell '待恢复'),(New-ValueCell $null),(New-ValueCell $null),(New-ValueCell $Note)
        )
        Set-CellMatrix '干预订单' "A${rowNumber}:M${rowNumber}" $newRow
        $reference = $OrderNo
        $writeResult = [pscustomobject]@{row=$rowNumber;state='待恢复';person=$Person}
    }
    'RestoreInventory' {
        if (-not $OrderNo) { throw '库存恢复必须提供平台订单号' }
        $rows = Get-SheetRows '干预订单' 'A2:M1001'
        $found = $false
        for ($i = 0; $i -lt $rows.Count; $i++) {
            if ((Get-CellValue $rows[$i] 3) -eq $OrderNo) {
                $rowNumber = 2 + $i
                Set-OneRow '干预订单' "J${rowNumber}:J${rowNumber}" @('已恢复')
                if ($Note) { Set-OneRow '干预订单' "M${rowNumber}:M${rowNumber}" @($Note) }
                $writeResult = [pscustomobject]@{row=$rowNumber;state='已恢复'}
                $found = $true
                break
            }
        }
        if (-not $found) { throw "未找到干预订单：$OrderNo" }
        $reference = $OrderNo
    }
    'ConfirmSelfDelivery' {
        if (-not $OrderNo -or $OrderNo -notmatch '^\d{10,30}$') { throw '自配送确认必须提供有效平台订单号' }
        $rows = Get-SheetRows '干预订单' 'A2:M1001'
        $found = $false
        for ($i = 0; $i -lt $rows.Count; $i++) {
            if ((Get-CellValue $rows[$i] 3) -ne $OrderNo) { continue }
            $rowNumber = 2 + $i
            $platformFee = Convert-ToNumber (Get-CellValue $rows[$i] 6)
            $totalCost = $platformFee
            Set-OneRow '干预订单' "F${rowNumber}:F${rowNumber}" @('自配送')
            Set-OneRow '干预订单' "H${rowNumber}:H${rowNumber}" @(0)
            if ($Note) {
                $oldNote = Get-CellValue $rows[$i] 12
                $newNote = if ($oldNote) { "$oldNote；$Note" } else { $Note }
                Set-OneRow '干预订单' "M${rowNumber}:M${rowNumber}" @($newNote)
            }
            $writeResult = [pscustomobject]@{
                row=$rowNumber
                delivery_type='自配送'
                delivery_fee=0
                external_delivery_fee=0
                platform_fee=$platformFee
                total_cost=$totalCost
            }
            $found = $true
            break
        }
        if (-not $found) { throw "未找到干预订单：$OrderNo；请先把它登记为补单" }
        $reference = $OrderNo
    }
    'SetExternalDeliveryFee' {
        if (-not $OrderNo -or $OrderNo -notmatch '^\d{10,30}$') { throw '配送费必须提供有效平台订单号' }
        Assert-Positive $Amount '配送费'
        $rows = Get-SheetRows '干预订单' 'A2:M1001'
        $found = $false
        for ($i = 0; $i -lt $rows.Count; $i++) {
            if ((Get-CellValue $rows[$i] 3) -ne $OrderNo) { continue }
            $rowNumber = 2 + $i
            $platformFee = Convert-ToNumber (Get-CellValue $rows[$i] 6)
            $amountValue = [double]$Amount
            $totalCost = $platformFee + $amountValue
            Set-OneRow '干预订单' "F${rowNumber}:F${rowNumber}" @('骑手配送')
            Set-OneRow '干预订单' "H${rowNumber}:H${rowNumber}" @($amountValue)
            if ($Note) {
                $oldNote = Get-CellValue $rows[$i] 12
                $newNote = if ($oldNote) { "$oldNote；$Note" } else { $Note }
                Set-OneRow '干预订单' "M${rowNumber}:M${rowNumber}" @($newNote)
            }
            $writeResult = [pscustomobject]@{
                row=$rowNumber
                delivery_type='骑手配送'
                delivery_fee=$amountValue
                external_delivery_fee=$amountValue
                platform_fee=$platformFee
                total_cost=$totalCost
            }
            $found = $true
            break
        }
        if (-not $found) { throw "未找到干预订单：$OrderNo；请先把它登记为补单" }
        $reference = $OrderNo
    }
    'SetOrderDeliveryFee' {
        if (-not $OrderNo -or $OrderNo -notmatch '^\d{10,30}$') { throw '配送费必须提供有效平台订单号' }
        Assert-NonNegative $Amount '配送费'
        $amountValue = [math]::Round([double]$Amount, 2)

        $interventionRows = Get-SheetRows '干预订单' 'A2:M1001'
        $interventionIndex = -1
        for ($i = 0; $i -lt $interventionRows.Count; $i++) {
            if ((Get-CellValue $interventionRows[$i] 3) -eq $OrderNo) { $interventionIndex = $i; break }
        }

        if ($interventionIndex -ge 0) {
            $rowNumber = 2 + $interventionIndex
            $row = $interventionRows[$interventionIndex]
            $oldAmount = Convert-ToNumber (Get-CellValue $row 7)
            $deliveryType = if ($amountValue -eq 0) { '自配送' } else { '骑手配送' }
            $platformFee = Convert-ToNumber (Get-CellValue $row 6)
            if ($DryRun) {
                [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{order_no=$OrderNo;order_type='干预单';date=(Get-CellValue $row 0);daily_order_no=(Get-CellValue $row 1);old_delivery_fee=$oldAmount;delivery_fee=$amountValue;delivery_type=$deliveryType;total_cost=[math]::Round($platformFee+$amountValue,2)}} | ConvertTo-Json -Depth 8
                Complete-CurrentFinanceTransaction
                exit 0
            }
            Set-OneRow '干预订单' "F${rowNumber}:F${rowNumber}" @($deliveryType)
            Set-OneRow '干预订单' "H${rowNumber}:H${rowNumber}" @($amountValue)
            if ($Note) {
                $oldNote = Get-CellValue $row 12
                $newNote = if ($oldNote) { "$oldNote；$Note" } else { $Note }
                Set-OneRow '干预订单' "M${rowNumber}:M${rowNumber}" @($newNote)
            }
            $writeResult = [pscustomobject]@{row=$rowNumber;order_no=$OrderNo;order_type='干预单';old_delivery_fee=$oldAmount;delivery_fee=$amountValue;delivery_type=$deliveryType;total_cost=[math]::Round($platformFee+$amountValue,2)}
        } else {
            $orderRows = Get-SheetRows '订单汇总' 'A2:Q2001'
            $orderIndex = -1
            for ($i = 0; $i -lt $orderRows.Count; $i++) {
                if ((Get-CellValue $orderRows[$i] 2) -eq $OrderNo) { $orderIndex = $i; break }
            }
            if ($orderIndex -lt 0) { throw "未找到订单：$OrderNo；请先同步对应日期的订单" }
            $rowNumber = 2 + $orderIndex
            $row = $orderRows[$orderIndex]
            if ((Get-CellValue $row 3) -eq '已取消') { throw '已取消订单不需要补录配送费' }
            $oldAmount = Convert-ToNumber (Get-CellValue $row 9)
            $orderDate = Convert-ToDateValue (Get-CellValue $row 0)
            if ($DryRun) {
                [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{order_no=$OrderNo;order_type='自然单';date=$(if($orderDate){$orderDate.ToString('yyyy-MM-dd')}else{$null});daily_order_no=(Get-CellValue $row 16);old_delivery_fee=$oldAmount;delivery_fee=$amountValue}} | ConvertTo-Json -Depth 8
                Complete-CurrentFinanceTransaction
                exit 0
            }
            Set-OneRow '订单汇总' "J${rowNumber}:J${rowNumber}" @($amountValue)
            $oldNote = [string](Get-CellValue $row 15)
            $cleanNote = [regex]::Replace($oldNote, '(?:^|[；;])\s*人工确认配送费：?\s*\d+(?:\.\d+)?\s*元?', '').Trim('；',';',' ')
            $manualMarker = "人工确认配送费：$($amountValue.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture))元"
            $newNote = ((@($cleanNote,$manualMarker,$Note) | Where-Object { $_ }) -join '；')
            Set-OneRow '订单汇总' "P${rowNumber}:P${rowNumber}" @($newNote)
            $writeResult = [pscustomobject]@{row=$rowNumber;order_no=$OrderNo;order_type='自然单';date=$(if($orderDate){$orderDate.ToString('yyyy-MM-dd')}else{$null});daily_order_no=(Get-CellValue $row 16);old_delivery_fee=$oldAmount;delivery_fee=$amountValue}
        }
        $reference = $OrderNo
    }
    'OfflineSale' {
        Assert-Positive $Amount '线下销售实收金额'
        if (-not $Items) { throw '线下销售必须提供售出货品，例如 kw1001×2、kw2001×1' }
        $matches = @([regex]::Matches($Items, '(?i)\b((?:kw|csx)\d+)\s*[×xX*]\s*(\d+(?:\.\d+)?)'))
        if ($matches.Count -eq 0) { throw '没有识别到“货号×数量”，例如 kw1001×2' }

        $skuRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $skuByCode = @{}
        foreach ($skuRow in $skuRows) {
            $kw = Get-CellValue $skuRow 0
            $aliases = Get-CellValue $skuRow 1
            if ($kw) { $skuByCode[$kw.ToLowerInvariant()] = $skuRow }
            if ($aliases) {
                foreach ($alias in ([regex]::Split($aliases, '[/,，;；\s]+') | Where-Object { $_ })) {
                    $skuByCode[$alias.ToLowerInvariant()] = $skuRow
                }
            }
        }
        $quantityByCode = [ordered]@{}
        foreach ($match in $matches) {
            $code = $match.Groups[1].Value.ToLowerInvariant()
            $quantityValue = [double]::Parse($match.Groups[2].Value, [Globalization.CultureInfo]::InvariantCulture)
            if (-not $quantityByCode.Contains($code)) { $quantityByCode[$code] = 0.0 }
            $quantityByCode[$code] += $quantityValue
        }
        $soldParts = @()
        $pendingCodes = @()
        $totalProductCost = 0.0
        foreach ($entry in $quantityByCode.GetEnumerator()) {
            $rawCode = [string]$entry.Key
            $displayCode = $rawCode
            $quantityValue = [double]$entry.Value
            if ($skuByCode.ContainsKey($rawCode)) {
                $skuRow = $skuByCode[$rawCode]
                $kw = Get-CellValue $skuRow 0
                if ($kw) { $displayCode = $kw.ToLowerInvariant() }
                $status = Get-CellValue $skuRow $skuSchema.CostStatus.Index
                $unitCost = Get-CellValue $skuRow $skuSchema.MovingAverageCost.Index
                if ($null -eq $unitCost -or [string]$unitCost -eq '') { $unitCost = Get-CellValue $skuRow $skuSchema.BaselineUnitCost.Index }
                if ($null -eq $unitCost -or [string]$unitCost -eq '' -or $status -in @('待人工核对','包装换算待核对','缺少成本','缺少参考成本')) {
                    $pendingCodes += $rawCode
                } else {
                    $totalProductCost += (Convert-ToNumber $unitCost) * $quantityValue
                }
            } else {
                $pendingCodes += $rawCode
            }
            $qtyText = if ([math]::Abs($quantityValue - [math]::Round($quantityValue)) -lt 0.000001) { [string][int][math]::Round($quantityValue) } else { $quantityValue.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture) }
            $soldParts += "$displayCode×$qtyText"
        }

        $rows = Get-SheetRows '线下订单' 'A2:J1001'
        $rowNumber = Get-FirstEmptyRow $rows 2 2 1001
        $sameDayCount = @($rows | Where-Object {
            (Convert-ToDateValue (Get-CellValue $_ 0)) -eq $Date.Date
        }).Count
        $offlineOrderNo = if ($EventId) {
            $sha = [Security.Cryptography.SHA256]::Create()
            try { 'OFF-' + $Date.ToString('yyyyMMdd') + '-' + ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($EventId))).Replace('-','').Substring(0,8)) } finally { $sha.Dispose() }
        } else { 'OFF-' + $Date.ToString('yyyyMMdd-HHmmss') }
        $deliveryValue = if ($null -eq $DeliveryFee) { 0.0 } else { [double]$DeliveryFee }
        if ($deliveryValue -lt 0) { throw '配送费不能小于0' }
        $costCell = if ($pendingCodes.Count -gt 0) { New-ValueCell $null } else { New-ValueCell ([math]::Round($totalProductCost,2)) }
        $notes = @()
        if ($Note) { $notes += $Note }
        if ($pendingCodes.Count -gt 0) { $notes += "货物成本待核对：$($pendingCodes -join '、')" }
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{order_no=$offlineOrderNo;daily_order_no="线下$($sameDayCount+1)号单";income=[double]$Amount;delivery_fee=$deliveryValue;product_cost=$(if($pendingCodes.Count){$null}else{[math]::Round($totalProductCost,2)});pending_cost_codes=$pendingCodes;sold_goods=($soldParts -join '、');receipt_status=$ReceiptStatus};message='预演成功，未写入飞书'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
        $newRow = @(); $newRow += ,@(
            (New-ValueCell (Convert-ToExcelSerial $Date)),(New-ValueCell "线下$($sameDayCount+1)号单"),(New-ValueCell $offlineOrderNo),
            (New-ValueCell ([double]$Amount)),(New-ValueCell $deliveryValue),$costCell,
            @{formula="=IF(C$rowNumber=`"`",`"`",IF(F$rowNumber=`"`",`"`",IFERROR(D$rowNumber,0)-F$rowNumber-IFERROR(E$rowNumber,0)))"},
            (New-ValueCell ($soldParts -join '、')),(New-ValueCell $ReceiptStatus),(New-ValueCell ($notes -join '；'))
        )
        Set-CellMatrix '线下订单' "A${rowNumber}:J${rowNumber}" $newRow
        if ($ReceiptStatus -eq '已到账') {
            $null = Update-MonthlyCell 4 'E' ([double]$Amount) $true "线下订单 $offlineOrderNo 实际到账 $([double]$Amount)"
        }
        $reference = $offlineOrderNo
        $writeResult = [pscustomobject]@{row=$rowNumber;order_no=$offlineOrderNo;daily_order_no="线下$($sameDayCount+1)号单";income=[double]$Amount;delivery_fee=$deliveryValue;product_cost=$(if($pendingCodes.Count){$null}else{[math]::Round($totalProductCost,2)});pending_cost_codes=$pendingCodes;sold_goods=($soldParts -join '、');receipt_status=$ReceiptStatus}
    }
    'AddOrderGift' {
        if (-not $OrderNo -and -not $DailyOrderNo) { throw '登记赠品必须提供完整订单编号或当天单号' }
        if ($OrderNo -and $OrderNo -notmatch '^(?:\d{10,30}|OFF-\d{8}-[A-Za-z0-9-]+)$') { throw '赠品订单编号格式无效' }
        if (-not $Items) { throw '登记赠品必须提供赠品货号和数量，例如 kw4189×1' }
        $giftMatches = @([regex]::Matches($Items, '(?i)\b((?:kw|csx)\d+)\s*[×xX*]\s*(\d+(?:\.\d+)?)'))
        if ($giftMatches.Count -eq 0) { throw '没有识别到赠品“货号×数量”，例如 kw4189×1' }

        $pendingOrder = $false
        try {
            $resolvedOrder = Resolve-OrderForGift $OrderNo $DailyOrderNo $Date
        } catch {
            if ($OrderNo -and $_.Exception.Message -match '^找不到要登记赠品的订单') {
                $pendingOrder = $true
                $resolvedOrder = [pscustomobject]@{
                    order_no = $OrderNo
                    daily_order_no = '待同步'
                    date = $Date.Date
                    platform = $Platform
                    order_type = '待匹配'
                    order_status = ''
                }
            } else {
                throw
            }
        }
        $giftRows = Get-SheetRows '赠品记录' 'A2:L2001'
        if ($EventId) {
            $existingGiftRows = @($giftRows | Where-Object { (Get-CellValue $_ 11) -eq $EventId })
            if ($existingGiftRows.Count -gt 0) {
                $reference = $resolvedOrder.order_no
                $costResult = Invoke-CostSnapshotRebuild
                $ledgerResult = Invoke-DailyLedgerSync ([datetime]$resolvedOrder.date)
                Save-ProcessedEvent $EventId $Action $reference
                [ordered]@{ok=$true;duplicate=$true;event_id=$EventId;action=$Action;reference=$reference;result=[ordered]@{status=$(if($pendingOrder){'pending_order'}else{'matched'});order_no=$resolvedOrder.order_no;daily_order_no=$resolvedOrder.daily_order_no;message=$(if($pendingOrder){'赠品已暂存，等待订单同步'}else{'赠品已登记；本次完成成本和简账刷新'})};cost_snapshot_rebuild=$costResult;daily_ledger_sync=$ledgerResult;formula_verification='success';workbook_url=$meta.workbook_url} | ConvertTo-Json -Depth 8
                Complete-CurrentFinanceTransaction
                exit 0
            }
        }

        $skuRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $skuByCode = Get-CanonicalSkuMap $skuRows
        $quantityByCode = [ordered]@{}
        foreach ($match in $giftMatches) {
            $rawCode = $match.Groups[1].Value.ToLowerInvariant()
            $quantityValue = [double]::Parse($match.Groups[2].Value, [Globalization.CultureInfo]::InvariantCulture)
            if ($quantityValue -le 0) { throw "赠品数量必须大于0：$rawCode" }
            if (-not $quantityByCode.Contains($rawCode)) { $quantityByCode[$rawCode] = 0.0 }
            $quantityByCode[$rawCode] += $quantityValue
        }

        $rowsToWrite = @()
        $giftParts = @()
        $pendingCodes = @()
        $totalGiftCost = 0.0
        foreach ($entry in $quantityByCode.GetEnumerator()) {
            $rawCode = [string]$entry.Key
            $quantityValue = [double]$entry.Value
            $displayCode = $rawCode
            $unitCost = $null
            $costStatus = '成本待核对'
            if ($skuByCode.ContainsKey($rawCode)) {
                $skuRow = $skuByCode[$rawCode].cells
                $canonicalKw = Get-CellValue $skuRow 0
                if ($canonicalKw) { $displayCode = $canonicalKw.Trim().ToLowerInvariant() }
                $masterStatus = Get-CellValue $skuRow $skuSchema.CostStatus.Index
                $unitCostText = Get-CellValue $skuRow $skuSchema.MovingAverageCost.Index
                if ($null -eq $unitCostText -or $unitCostText -eq '') { $unitCostText = Get-CellValue $skuRow $skuSchema.BaselineUnitCost.Index }
                if ($null -ne $unitCostText -and $unitCostText -ne '' -and $masterStatus -notin @('待人工核对','包装换算待核对','缺少成本','缺少参考成本')) {
                    $unitCost = Convert-ToNumber $unitCostText
                    $costStatus = '已计成本'
                }
            }
            $giftCost = if ($null -eq $unitCost) { $null } else { [math]::Round(([double]$unitCost * $quantityValue),2) }
            if ($null -eq $giftCost) { $pendingCodes += $displayCode } else { $totalGiftCost += [double]$giftCost }
            $qtyText = if ([math]::Abs($quantityValue - [math]::Round($quantityValue)) -lt 0.000001) { [string][int][math]::Round($quantityValue) } else { $quantityValue.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture) }
            $giftParts += "$displayCode×$qtyText"
            $baseNote = if ($null -eq $giftCost) { "赠品成本待核对：${displayCode}未知成本价" } else { $Note }
            $rowNote = if ($pendingOrder) {
                if ($baseNote) { "等待订单同步；$baseNote" } else { '等待订单同步' }
            } else { $baseNote }
            $rowsToWrite += ,@(
                (New-ValueCell (Convert-ToExcelSerial $resolvedOrder.date)),(New-ValueCell $resolvedOrder.platform),(New-ValueCell $resolvedOrder.daily_order_no),(New-ValueCell $resolvedOrder.order_no),
                (New-ValueCell $displayCode),(New-ValueCell $quantityValue),(New-ValueCell $unitCost),(New-ValueCell $giftCost),
                (New-ValueCell $costStatus),(New-ValueCell $rowNote),(New-ValueCell (Convert-ToExcelSerial ([datetime]::Now))),(New-ValueCell $EventId)
            )
        }

        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{status=$(if($pendingOrder){'pending_order'}else{'matched'});order_no=$resolvedOrder.order_no;daily_order_no=$resolvedOrder.daily_order_no;order_type=$resolvedOrder.order_type;gift_items=($giftParts -join '、');gift_cost=$(if($pendingCodes.Count -gt 0){$null}else{[math]::Round($totalGiftCost,2)});pending_cost_codes=$pendingCodes};message=$(if($pendingOrder){'预演成功：确认后先暂存赠品，等待订单同步自动匹配'}else{'预演成功，未写入飞书'})} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
        $firstRow = Get-FirstEmptyRow $giftRows 3 2 2001
        $lastRow = $firstRow + $rowsToWrite.Count - 1
        if ($lastRow -gt 2001) { throw '赠品记录表容量已满' }
        Set-CellMatrix '赠品记录' "A${firstRow}:L${lastRow}" $rowsToWrite
        $reference = $resolvedOrder.order_no
        $writeResult = [pscustomobject]@{
            status=$(if($pendingOrder){'pending_order'}else{'matched'})
            first_row=$firstRow
            last_row=$lastRow
            order_no=$resolvedOrder.order_no
            daily_order_no=$resolvedOrder.daily_order_no
            platform=$resolvedOrder.platform
            order_type=$resolvedOrder.order_type
            gift_items=($giftParts -join '、')
            gift_cost=$(if($pendingCodes.Count -gt 0){$null}else{[math]::Round($totalGiftCost,2)})
            pending_cost_codes=$pendingCodes
            accounting=$(if($resolvedOrder.order_type -eq '干预'){'营销赠品支出'}else{'并入货物成本'})
        }
    }
    'Replenishment' {
        if (-not $GoodsCode) { throw '补货必须提供货号' }
        Assert-Positive $Quantity '补货数量'
        Assert-NonNegative $PurchaseAmount '实际采购金额'
        if ($null -ne $PaymentAmount) { Assert-NonNegative $PaymentAmount '付款金额' }
        if ($GoodsCode -notmatch '(?i)^(?:kw|csx)\d+$') { throw '补货货号必须是有效的 KW 或 CSX 编号' }
        if ($TransitionCode -and $TransitionCode -notmatch '(?i)^csx\d+$') { throw '补货关联的过渡货号必须是有效的 CSX 编号' }
        $skuRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $normalizedCode = $GoodsCode.Trim().ToLowerInvariant()
        $normalizedTransitionCode = if ($TransitionCode) { $TransitionCode.Trim().ToLowerInvariant() } else { $null }
        $requestedKw = if ($normalizedCode -match '^kw') { $normalizedCode } else { $null }
        $requestedCsx = if ($normalizedTransitionCode) { $normalizedTransitionCode } elseif ($normalizedCode -match '^csx') { $normalizedCode } else { $null }
        $lookupCodes = @($normalizedCode,$normalizedTransitionCode | Where-Object { $_ } | Select-Object -Unique)
        $candidateRows = @{}
        foreach ($lookupCode in $lookupCodes) {
            $matches = @(Get-SkuCodeCandidates $skuRows $lookupCode)
            if ($matches.Count -gt 1) {
                $duplicateRows = @($matches | ForEach-Object { $_.row_number }) -join '、'
                throw "SKU成本主表中的货号 $lookupCode 重复出现在第 $duplicateRows 行；已停止补货写入，避免新增重复档案。请先合并或确认该货号当前对应的商品"
            }
            if ($matches.Count -eq 1) { $candidateRows[[string]$matches[0].row_number] = $matches[0] }
        }
        if ($candidateRows.Count -gt 1) {
            $conflictRows = @($candidateRows.Keys | Sort-Object {[int]$_}) -join '、'
            throw "本次提供的 KW/CSX 分别指向 SKU成本主表第 $conflictRows 行，无法判定为同一商品；没有写入，请先确认对应关系"
        }
        $newSkuPending = $candidateRows.Count -eq 0
        $mappingUpdate = $null
        if ($newSkuPending) {
            $skuTarget = [pscustomobject]@{row_number=Get-FirstCompletelyEmptyRow $skuRows 26 2 3000;cells=@()}
            $currentKw = $requestedKw
            $productId = $null
            $unitsPerPack = 1.0
        } else {
            $skuTarget = @($candidateRows.Values)[0]
            $existingKw = Get-CellValue $skuTarget.cells 0
            $existingCsx = Get-CellValue $skuTarget.cells 1
            if ($requestedKw) {
                $kwMatches = @(Get-SkuCodeCandidates $skuRows $requestedKw)
                if ($kwMatches.Count -eq 1 -and [int]$kwMatches[0].row_number -ne [int]$skuTarget.row_number) {
                    throw "当前 KW $requestedKw 已属于 SKU成本主表第 $($kwMatches[0].row_number) 行，不能自动绑定到 CSX $requestedCsx"
                }
                if (-not $existingKw -or $existingKw.Trim().ToLowerInvariant() -ne $requestedKw) {
                    $mappingUpdate = [pscustomobject]@{column='A';old=$existingKw;new=$requestedKw}
                }
            }
            if ($requestedCsx -and $existingCsx -and $existingCsx.Trim().ToLowerInvariant() -ne $requestedCsx) {
                throw "SKU成本主表第 $($skuTarget.row_number) 行已有 CSX/过渡货号 $existingCsx，与本次提供的 $requestedCsx 不一致；没有覆盖原值"
            }
            $currentKw = if ($requestedKw) { $requestedKw } else { $existingKw }
            $productId = Get-CellValue $skuTarget.cells $skuSchema.ProductId.Index
            $unitsPerPack = Convert-ToNumber (Get-CellValue $skuTarget.cells $skuSchema.UnitsPerPack.Index)
            if ($unitsPerPack -le 0) { $unitsPerPack = 1.0 }
        }
        $displayCode = if ($currentKw) { $currentKw } else { $normalizedCode }
        $salesUnitQuantity = [double]$Quantity * $unitsPerPack
        $purchaseAmountValue = [double]$PurchaseAmount
        $incomingUnitCost = if ($salesUnitQuantity -gt 0) { $purchaseAmountValue / $salesUnitQuantity } else { $null }
        $beforeQuantity = if ($newSkuPending) { 0.0 } else { Convert-ToNullableNumber (Get-CellValue $skuTarget.cells $skuSchema.CostBookQuantity.Index) }
        if ($null -eq $beforeQuantity) { $beforeQuantity = 0.0 }
        $beforeAverage = if ($newSkuPending) { $null } else { Convert-ToNullableNumber (Get-CellValue $skuTarget.cells $skuSchema.MovingAverageCost.Index) }
        $computedExpiry = $null
        if (($null -ne $ProductionDate) -xor ($null -ne $ShelfLifeValue)) { throw '登记批次时，生产日期和保质期必须一起提供' }
        if ($null -ne $ProductionDate) {
            Assert-Positive $ShelfLifeValue '保质期'
            $computedExpiry = Add-ShelfLife ([datetime]$ProductionDate) ([double]$ShelfLifeValue) $ShelfLifeUnit
        }
        $rows = Get-SheetRows '补货记录' 'A2:T1001'
        # 批次ID同时承担远端幂等键：本地进程即使在飞书写入后、事件日志落盘前崩溃，
        # 重试也会复用原行，而不是再次增加库存。
        $eventBatchId = if ($EventId) { 'EVT-' + (Get-StableEventKey $EventId) } else { $null }
        $existingEventIndex = -1
        if ($eventBatchId) {
            for ($i = 0; $i -lt $rows.Count; $i++) {
                if ((Get-CellValue $rows[$i] 19) -eq $eventBatchId) { $existingEventIndex = $i; break }
            }
        }
        $rowNumber = if ($existingEventIndex -ge 0) { 2 + $existingEventIndex } else { Get-FirstEmptyRow $rows 1 2 1001 }
        $quantityValue = [double]$Quantity
        $paymentDateValue = if ($null -ne $PaymentAmount) { Convert-ToExcelSerial $(if ($null -ne $PaymentDate) { [datetime]$PaymentDate } else { $Date }) } else { $null }
        $paymentAmountValue = if ($null -ne $PaymentAmount) { [double]$PaymentAmount } else { $null }
        $batchId = if ($eventBatchId) { $eventBatchId } elseif ($computedExpiry) { "BATCH-$($Date.ToString('yyyyMMdd'))-$('{0:D4}' -f $rowNumber)" } else { $null }
        $existingBeforeQuantity = if ($existingEventIndex -ge 0) { Convert-ToNullableNumber (Get-CellValue $rows[$existingEventIndex] 7) } else { $null }
        $existingBeforeAverage = if ($existingEventIndex -ge 0) { Convert-ToNullableNumber (Get-CellValue $rows[$existingEventIndex] 8) } else { $null }
        $existingAfterQuantity = if ($existingEventIndex -ge 0) { Convert-ToNullableNumber (Get-CellValue $rows[$existingEventIndex] 9) } else { $null }
        $existingAfterAverage = if ($existingEventIndex -ge 0) { Convert-ToNullableNumber (Get-CellValue $rows[$existingEventIndex] 10) } else { $null }
        if ($null -ne $existingAfterQuantity) {
            $beforeQuantity = $existingBeforeQuantity
            $beforeAverage = $existingBeforeAverage
            $afterQuantity = $existingAfterQuantity
            $afterAverage = $existingAfterAverage
        } else {
            $afterQuantity = [double]$beforeQuantity + [double]$salesUnitQuantity
            if ($null -eq $beforeAverage -and [double]$beforeQuantity -gt 0) {
                # 旧库存成本未知时，不能只用本次进价伪造整批库存成本。
                $afterAverage = $null
            } elseif ($afterQuantity -le 0) {
                $afterAverage = $incomingUnitCost
            } else {
                $beforeValue = if ($null -eq $beforeAverage) { 0.0 } else { [double]$beforeAverage * [double]$beforeQuantity }
                $afterAverage = ($beforeValue + $purchaseAmountValue) / $afterQuantity
            }
        }
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{row=$rowNumber;sku_master_row=[int]$skuTarget.row_number;goods_code=$displayCode;transition_code=$requestedCsx;product_id=$productId;package_quantity=$quantityValue;sales_unit_quantity=$salesUnitQuantity;purchase_amount=$purchaseAmountValue;unit_cost=$incomingUnitCost;before_quantity=$beforeQuantity;before_average_cost=$beforeAverage;after_quantity=$afterQuantity;after_average_cost=$afterAverage;production_date=$(if($ProductionDate){([datetime]$ProductionDate).ToString('yyyy-MM-dd')}else{$null});expiry_date=$(if($computedExpiry){$computedExpiry.ToString('yyyy-MM-dd')}else{$null});sku_status=$(if($newSkuPending){'will_create_pending_profile'}elseif($mappingUpdate){'will_attach_kw_by_csx'}else{'matched'});mapping_update=$mappingUpdate;affected_tables=@('补货记录','SKU成本主表');remote_replay=($existingEventIndex -ge 0)};message='预演成功，未写入飞书'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
        if ($eventBatchId -and $existingEventIndex -lt 0) {
            # 先占用远端幂等行；后续任何阶段中断，重试都会回到这一行补完。
            Set-OneRow '补货记录' "T${rowNumber}:T${rowNumber}" @($eventBatchId)
        }
        if ($newSkuPending) {
            $kwValue = $requestedKw
            $csxValue = $requestedCsx
            $zeroCost = [double]$PurchaseAmount -eq 0
            $initialCost = if ($zeroCost) { 0.0 } else { $null }
            $costStatus = if ($zeroCost) { '零成本' } else { '新SKU待核对' }
            $basis = if ($zeroCost) { '用户明确试吃装/赠品成本为0' } else { '新SKU按本次补货建立成本，采购包装换算默认1' }
            $profileNote = if ($Note) { "新货号待平台库存同步补全商品资料；$Note" } else { '新货号待平台库存同步补全商品资料' }
            $newProfile = @(); for ($columnIndex = 0; $columnIndex -lt 26; $columnIndex++) { $newProfile += ,@{} }
            $newProfile[$skuSchema.CurrentKw.Index] = New-ValueCell $kwValue
            $newProfile[$skuSchema.TransitionCode.Index] = New-ValueCell $csxValue
            $newProfile[$skuSchema.InitialInventory.Index] = New-ValueCell 0
            $newProfile[$skuSchema.BaselineUnitCost.Index] = New-ValueCell $initialCost
            $newProfile[$skuSchema.MovingAverageCost.Index] = New-ValueCell $initialCost
            $newProfile[$skuSchema.CostBookQuantity.Index] = New-ValueCell 0
            $newProfile[$skuSchema.UnitsPerPack.Index] = New-ValueCell 1
            $newProfile[$skuSchema.PurchaseUnit.Index] = New-ValueCell '个'
            $newProfile[$skuSchema.CostStatus.Index] = New-ValueCell $costStatus
            $newProfile[$skuSchema.InitializationDate.Index] = New-ValueCell (Convert-ToExcelSerial $Date)
            $newProfile[$skuSchema.LastCostUpdate.Index] = New-ValueCell (Convert-ToExcelSerial ([datetime]::Now))
            $newProfile[$skuSchema.ConversionBasis.Index] = New-ValueCell $basis
            $newProfile[$skuSchema.Note.Index] = New-ValueCell $profileNote
            $newProfile[$skuSchema.EarliestExpiry.Index] = New-ValueCell $(if($computedExpiry){Convert-ToExcelSerial $computedExpiry}else{$null})
            $newProfile[$skuSchema.PlatformStatus.Index] = New-ValueCell '待平台同步'
            $profileMatrix = @(); $profileMatrix += ,$newProfile
            Set-CellMatrix 'SKU成本主表' "A$($skuTarget.row_number):Z$($skuTarget.row_number)" $profileMatrix
        } elseif ($mappingUpdate) {
            Set-OneRow 'SKU成本主表' "A$($skuTarget.row_number):A$($skuTarget.row_number)" @($requestedKw)
        }
        Set-OneRow '补货记录' "A${rowNumber}:C${rowNumber}" @((Convert-ToExcelSerial $Date),$displayCode,$quantityValue)
        # D 列必须使用 SKU 主表维护的“每采购包装含销售单位数”。
        # 旧模板曾误把商品单位成本写进此列，导致 6 袋被显示成 600 个等错误。
        Set-OneRow '补货记录' "D${rowNumber}:D${rowNumber}" @($unitsPerPack)
        $eFormulaMatrix = @(); $eFormulaMatrix += ,@(@{formula="=IF(OR(C${rowNumber}=`"`",D${rowNumber}=`"`"),`"`",C${rowNumber}*D${rowNumber})"})
        Set-CellMatrix '补货记录' "E${rowNumber}:E${rowNumber}" $eFormulaMatrix
        Set-OneRow '补货记录' "F${rowNumber}:F${rowNumber}" @($purchaseAmountValue)
        $gFormulaMatrix = @(); $gFormulaMatrix += ,@(@{formula="=IF(OR(E${rowNumber}=`"`",F${rowNumber}=`"`"),`"`",F${rowNumber}/E${rowNumber})"})
        Set-CellMatrix '补货记录' "G${rowNumber}:G${rowNumber}" $gFormulaMatrix
        Set-OneRow '补货记录' "L${rowNumber}:N${rowNumber}" @($paymentDateValue,$paymentAmountValue,$Note)
        Set-OneRow '补货记录' "O${rowNumber}:T${rowNumber}" @(
            $productId,
            $(if($ProductionDate){Convert-ToExcelSerial ([datetime]$ProductionDate)}else{$null}),
            $ShelfLifeValue,
            $(if($ProductionDate){$ShelfLifeUnit}else{$null}),
            $(if($computedExpiry){Convert-ToExcelSerial $computedExpiry}else{$null}),
            $batchId
        )
        $snapshotCells = @(); $snapshotCells += ,@(
            (New-UnitCostCell $beforeQuantity),(New-UnitCostCell $beforeAverage),
            (New-UnitCostCell $afterQuantity),(New-UnitCostCell $afterAverage)
        )
        $masterCostCells = @(); $masterCostCells += ,@(
            (New-UnitCostCell $afterAverage),(New-UnitCostCell $afterQuantity)
        )
        $masterUpdatedAtCells = @(); $masterUpdatedAtCells += ,@((New-ValueCell (Convert-ToExcelSerial ([datetime]::Now))))
        Set-AtomicCellOperations @(
            @{shortcut='+cells-set';input=@{sheet_name='补货记录';range="H${rowNumber}:K${rowNumber}";cells=$snapshotCells}},
            @{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.MovingAverageCost.Column)$($skuTarget.row_number):$($skuSchema.CostBookQuantity.Column)$($skuTarget.row_number)";cells=$masterCostCells}},
            @{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.LastCostUpdate.Column)$($skuTarget.row_number):$($skuSchema.LastCostUpdate.Column)$($skuTarget.row_number)";cells=$masterUpdatedAtCells}}
        )
        $reference = "$GoodsCode x$quantityValue"
        $writeResult = [pscustomobject]@{row=$rowNumber;sku_master_row=[int]$skuTarget.row_number;goods_code=$displayCode;transition_code=$requestedCsx;product_id=$productId;package_quantity=$quantityValue;sales_unit_quantity=$salesUnitQuantity;purchase_amount=$purchaseAmountValue;payment_amount=$paymentAmountValue;unit_cost=$incomingUnitCost;before_quantity=$beforeQuantity;before_average_cost=$beforeAverage;after_quantity=$afterQuantity;after_average_cost=$afterAverage;batch_id=$batchId;expiry_date=$(if($computedExpiry){$computedExpiry.ToString('yyyy-MM-dd')}else{$null});sku_status=$(if($newSkuPending){'created_pending_profile'}elseif($mappingUpdate){'attached_kw_by_csx'}else{'matched'});mapping_update=$mappingUpdate;affected_tables=@('补货记录','SKU成本主表');remote_replay=($existingEventIndex -ge 0)}
    }
    'SetSkuCost' {
        if (-not $GoodsCode -or $GoodsCode.Trim().Length -gt 100) { throw '修改成本必须提供 SKU成本主表中可唯一定位的当前货号或过渡标识' }
        Assert-NonNegative $Amount '销售单位成本'
        $normalizedCode = $GoodsCode.Trim().ToLowerInvariant()
        $skuRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $target = Resolve-UniqueSkuReference $skuRows $normalizedCode
        $rowNumber = [int]$target.row_number
        $canonicalKw = Get-CellValue $target.cells 0
        $transitionCode = Get-CellValue $target.cells 1
        $productId = Get-CellValue $target.cells $skuSchema.ProductId.Index
        $displayCode = if ($canonicalKw) {
            $canonicalKw.Trim().ToLowerInvariant()
        } elseif ($transitionCode) {
            @([regex]::Split($transitionCode, '[/,，、;；\s]+') | Where-Object { $_ } | Select-Object -First 1)[0].Trim().ToLowerInvariant()
        } elseif ($productId) {
            "SKU ID $productId"
        } else {
            $normalizedCode
        }
        $title = Get-CellValue $target.cells 2
        $oldBaselineCostText = Get-CellValue $target.cells $skuSchema.BaselineUnitCost.Index
        $oldBaselineCost = if ($oldBaselineCostText -eq $null -or $oldBaselineCostText -eq '') { $null } else { Convert-ToNumber $oldBaselineCostText }
        $oldMovingCostText = Get-CellValue $target.cells $skuSchema.MovingAverageCost.Index
        $oldMovingCost = if ($oldMovingCostText -eq $null -or $oldMovingCostText -eq '') { $null } else { Convert-ToNumber $oldMovingCostText }
        $newCost = [math]::Round([double]$Amount,4)
        $oldStatus = Get-CellValue $target.cells $skuSchema.CostStatus.Index
        $newStatus = if ([string]$oldStatus -match '当前KW冲突') { '当前KW冲突待核对（成本已确认）' } else { '用户确认' }
        $oldNote = Get-CellValue $target.cells $skuSchema.Note.Index
        $newNoteEntry = "店主通过飞书确认销售单位成本为 $($newCost.ToString('0.####',[Globalization.CultureInfo]::InvariantCulture)) 元（$($Date.ToString('yyyy-MM-dd'))）"
        $oldNoteParts = @([regex]::Split([string]$oldNote, '[；;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $newNote = ((@($oldNoteParts + @($newNoteEntry,$Note)) | Where-Object { $_ } | Select-Object -Unique) -join '；')
        $targetOrderItems = [Collections.Generic.List[object]]::new()
        if ($OrderNo) {
            if ($OrderNo -notmatch '^\d{10,30}$') { throw '更新历史订单成本必须提供有效的平台订单号' }
            $orderItemRows = Get-SheetRows '订单商品' 'A2:Q4001'
            $matchingCodes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($code in @($normalizedCode,$canonicalKw) + @([regex]::Split([string]$transitionCode, '[/,，、;；\s]+'))) {
                if ($code) { $null = $matchingCodes.Add(([string]$code).Trim().ToLowerInvariant()) }
            }
            for ($itemIndex = 0; $itemIndex -lt $orderItemRows.Count; $itemIndex++) {
                $itemRow = $orderItemRows[$itemIndex]
                if ((Get-CellValue $itemRow 2) -ne $OrderNo) { continue }
                $itemProductId = ([string](Get-CellValue $itemRow 16)).Trim()
                $itemCode = ([string](Get-CellValue $itemRow 5)).Trim().ToLowerInvariant()
                $sameProduct = ($productId -and $itemProductId -eq $productId) -or ($itemCode -and $matchingCodes.Contains($itemCode))
                if (-not $sameProduct) { continue }
                $itemDate = Convert-ToDateValue (Get-CellValue $itemRow 0)
                if ($null -eq $itemDate) { throw "订单 $OrderNo 的商品日期缺失，不能安全更新成本" }
                $itemQuantity = Convert-ToNumber (Get-CellValue $itemRow 8)
                $itemSalesAmount = Convert-ToNumber (Get-CellValue $itemRow 10)
                $targetOrderItems.Add([pscustomobject]@{
                    row=2+$itemIndex;date=$itemDate;quantity=$itemQuantity;sales_amount=$itemSalesAmount
                    old_unit_cost=Convert-ToNullableNumber (Get-CellValue $itemRow 11)
                    old_item_cost=Convert-ToNullableNumber (Get-CellValue $itemRow 12)
                    new_unit_cost=$newCost
                    new_item_cost=[math]::Round(($newCost * $itemQuantity),4)
                    new_margin=[math]::Round(($itemSalesAmount - ($newCost * $itemQuantity)),4)
                })
            }
            if ($targetOrderItems.Count -eq 0) { throw "订单 $OrderNo 中找不到与 $displayCode 对应的商品；没有修改任何成本" }
            $distinctOrderDates = @($targetOrderItems | ForEach-Object { $_.date.ToString('yyyy-MM-dd') } | Select-Object -Unique)
            if ($distinctOrderDates.Count -ne 1) { throw "订单 $OrderNo 的商品日期不一致，不能安全刷新对应月份" }
            $explicitHistoricalOrderDate = [datetime]$targetOrderItems[0].date
        }
        $affectedTables = if ($OrderNo) { @('SKU成本主表','指定订单的订单商品','订单汇总','对应月份收入','驾驶舱') } else { @('SKU成本主表') }

        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{
                sheet='SKU成本主表';row=$rowNumber;range="$($skuSchema.BaselineUnitCost.Column)${rowNumber}:$($skuSchema.MovingAverageCost.Column)${rowNumber},$($skuSchema.CostStatus.Column)${rowNumber},$($skuSchema.LastCostUpdate.Column)${rowNumber}:$($skuSchema.Note.Column)${rowNumber}"
                goods_code=$displayCode;sku_reference=$normalizedCode;product_id=$productId;title=$title;old_baseline_cost=$oldBaselineCost;old_moving_average_cost=$oldMovingCost
                new_unit_cost=$newCost;new_moving_average_cost=$newCost;old_status=$oldStatus;new_status=$newStatus;historical_order_no=$OrderNo
                historical_order_items=@($targetOrderItems);affected_tables=$affectedTables
            };message='预演成功，未写入飞书'} | ConvertTo-Json -Depth 10
            Complete-CurrentFinanceTransaction
            exit 0
        }

        if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
        $backupPath = Join-Path $stateRoot ("sku_cost_change_backup_{0}_{1}.json" -f $displayCode,[datetime]::Now.ToString('yyyyMMdd_HHmmss'))
        [ordered]@{source="SKU成本主表!A${rowNumber}:Z${rowNumber}";captured_at=[datetime]::Now.ToString('o');cells=@($target.cells)} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $backupPath -Encoding UTF8

        $costMatrix = @(); $costMatrix += ,@((New-UnitCostCell $newCost),(New-UnitCostCell $newCost))
        $jMatrix = @(); $jMatrix += ,@((New-ValueCell $newStatus))
        $updatedAtMatrix = @(); $updatedAtMatrix += ,@((New-ValueCell (Convert-ToExcelSerial ([datetime]::Now))))
        $mnMatrix = @(); $mnMatrix += ,@((New-ValueCell '店主确认的销售单位成本'),(New-ValueCell $newNote))
        $costOperations = [Collections.Generic.List[object]]::new()
        $costOperations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.BaselineUnitCost.Column)${rowNumber}:$($skuSchema.MovingAverageCost.Column)${rowNumber}";cells=$costMatrix}})
        $costOperations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.CostStatus.Column)${rowNumber}:$($skuSchema.CostStatus.Column)${rowNumber}";cells=$jMatrix}})
        $costOperations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.LastCostUpdate.Column)${rowNumber}:$($skuSchema.LastCostUpdate.Column)${rowNumber}";cells=$updatedAtMatrix}})
        $costOperations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.ConversionBasis.Column)${rowNumber}:$($skuSchema.Note.Column)${rowNumber}";cells=$mnMatrix}})
        foreach ($targetItem in $targetOrderItems) {
            $itemCostCells = @(); $itemCostCells += ,@(
                (New-UnitCostCell $targetItem.new_unit_cost),(New-UnitCostCell $targetItem.new_item_cost),(New-UnitCostCell $targetItem.new_margin)
            )
            $costOperations.Add(@{shortcut='+cells-set';input=@{sheet_name='订单商品';range="L$($targetItem.row):N$($targetItem.row)";cells=$itemCostCells}})
        }
        Set-AtomicCellOperations $costOperations.ToArray()

        $finalReadback = Get-SkuCostReadback $rowNumber $newCost $newStatus $newCost
        $actualBaseline = [double]$finalReadback.baseline_cost
        $actualMoving = [double]$finalReadback.moving_average_cost
        $actualStatus = [string]$finalReadback.status
        if ([math]::Abs($actualBaseline - $newCost) -gt 0.0001 -or [math]::Abs($actualMoving - $newCost) -gt 0.0001 -or $actualStatus -ne $newStatus) {
            throw "SKU成本写入后回读不一致：$displayCode（基准成本=$actualBaseline，当前成本=$actualMoving，状态=$actualStatus，重试=$($finalReadback.attempt)次）"
        }
        $reference = "$displayCode $newCost"
        $writeResult = [pscustomobject]@{
            sheet='SKU成本主表';row=$rowNumber;range="$($skuSchema.BaselineUnitCost.Column)${rowNumber}:$($skuSchema.MovingAverageCost.Column)${rowNumber},$($skuSchema.CostStatus.Column)${rowNumber},$($skuSchema.LastCostUpdate.Column)${rowNumber}:$($skuSchema.Note.Column)${rowNumber}"
            goods_code=$displayCode;sku_reference=$normalizedCode;product_id=$productId;title=$title;old_baseline_cost=$oldBaselineCost;old_moving_average_cost=$oldMovingCost
            new_unit_cost=$newCost;new_moving_average_cost=$actualMoving;old_status=$oldStatus;new_status=$actualStatus
            historical_order_no=$OrderNo;historical_order_items=@($targetOrderItems);affected_tables=$affectedTables
        }
    }
    'SetSkuExpiry' {
        if (-not $GoodsCode -or $GoodsCode -notmatch '(?i)^(?:kw|csx)\d+$') { throw '更新最早过期日期必须提供有效的 KW 或 CSX 货号' }
        if ($ClearExpiry -and $null -ne $ExpiryDate) { throw '清空过期日期时不能同时提供新日期' }
        if (-not $ClearExpiry -and $null -eq $ExpiryDate) { throw '请提供最早过期日期，或明确要求清空' }

        $normalizedCode = $GoodsCode.Trim().ToLowerInvariant()
        $skuRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $kwMatches = [Collections.Generic.List[object]]::new()
        $aliasMatches = [Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $skuRows.Count; $i++) {
            $row = $skuRows[$i]
            $kw = (Get-CellValue $row 0)
            if ($kw -and $kw.Trim().ToLowerInvariant() -eq $normalizedCode) {
                $kwMatches.Add([pscustomobject]@{row_number=2+$i;cells=$row})
                continue
            }
            $transitionCodes = Get-CellValue $row 1
            if (-not $transitionCodes) { continue }
            $transitionSet = @([regex]::Split($transitionCodes, '[/,，、;；\s]+') | Where-Object { $_ -match '^(?i)csx\d+$' } | ForEach-Object { $_.Trim().ToLowerInvariant() })
            if ($normalizedCode -in $transitionSet) { $aliasMatches.Add([pscustomobject]@{row_number=2+$i;cells=$row}) }
        }
        $skuMatches = if ($kwMatches.Count -gt 0) { @($kwMatches) } else { @($aliasMatches) }
        if ($skuMatches.Count -eq 0) { throw "SKU成本主表中找不到货号：$GoodsCode" }
        if ($skuMatches.Count -gt 1) { throw "货号 $GoodsCode 在 SKU成本主表中匹配到多行，请先核对 KW/CSX 映射" }

        $target = $skuMatches[0]
        $rowNumber = [int]$target.row_number
        $canonicalKw = Get-CellValue $target.cells 0
        $displayCode = if ($canonicalKw) { $canonicalKw } else { $normalizedCode }
        $oldExpiry = Convert-ToDateValue (Get-CellValue $target.cells $skuSchema.EarliestExpiry.Index)
        $newExpiryText = if ($ClearExpiry) { '-' } else { ([datetime]$ExpiryDate).ToString('yyyy-MM-dd') }

        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{row=$rowNumber;goods_code=$displayCode;old_expiry=$(if($oldExpiry){$oldExpiry.ToString('yyyy-MM-dd')}else{$null});new_expiry=$newExpiryText};message='预演成功，未写入飞书'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }

        $expiryCell = if ($ClearExpiry) {
            @{value='-';cell_styles=@{number_format='yyyy-mm-dd';background_color='#FFF8E1';horizontal_alignment='center'}}
        } else {
            @{value=(Convert-ToExcelSerial ([datetime]$ExpiryDate));cell_styles=@{number_format='yyyy-mm-dd';background_color='#FFF8E1';horizontal_alignment='center'}}
        }
        $matrix = @(); $matrix += ,@($expiryCell)
        Set-CellMatrix 'SKU成本主表' "$($skuSchema.EarliestExpiry.Column)${rowNumber}:$($skuSchema.EarliestExpiry.Column)${rowNumber}" $matrix

        $checkRaw = & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range "A${rowNumber}:Z${rowNumber}" --include value --max-chars 10000 --format json | Out-String
        if ($LASTEXITCODE -ne 0) { throw "最早过期日期回读失败：$displayCode" }
        $checkResult = $checkRaw | ConvertFrom-Json
        if (-not $checkResult.ok -or $checkResult.data.has_more) { throw "最早过期日期回读不完整：$displayCode" }
        $checkRow = @($checkResult.data.ranges[0].cells[0])
        $actualValue = Get-CellValue $checkRow $skuSchema.EarliestExpiry.Index
        if ($ClearExpiry) {
            if ($actualValue -ne '-') { throw "最早过期日期清空回读不一致：$displayCode" }
        } else {
            $actualDate = Convert-ToDateValue $actualValue
            if ($null -eq $actualDate -or $actualDate.Date -ne ([datetime]$ExpiryDate).Date) { throw "最早过期日期回读不一致：$displayCode" }
        }
        $reference = "$displayCode $newExpiryText"
        $writeResult = [pscustomobject]@{
            row = $rowNumber
            goods_code = $displayCode
            old_expiry = if ($oldExpiry) { $oldExpiry.ToString('yyyy-MM-dd') } else { $null }
            earliest_expiry = $newExpiryText
        }
    }
    'SetSkuBatchDate' {
        if (-not $GoodsCode -or $GoodsCode -notmatch '(?i)^(?:kw|csx)\d+$') { throw '登记生产日期必须提供有效的 KW 或 CSX 货号' }
        if ($null -eq $ProductionDate) { throw '请提供生产日期' }
        Assert-Positive $ShelfLifeValue '保质期'
        $computedExpiry = Add-ShelfLife ([datetime]$ProductionDate) ([double]$ShelfLifeValue) $ShelfLifeUnit
        $normalizedCode = $GoodsCode.Trim().ToLowerInvariant()
        $skuRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $skuMatches = [Collections.Generic.List[object]]::new()
        for ($i=0; $i -lt $skuRows.Count; $i++) {
            $row = $skuRows[$i]
            $kw = Get-CellValue $row 0
            $transitionCodes = Get-CellValue $row 1
            $isKw = $kw -and $kw.Trim().ToLowerInvariant() -eq $normalizedCode
            $isCsx = $false
            if ($transitionCodes) {
                $isCsx = $normalizedCode -in @([regex]::Split($transitionCodes, '[/,，、;；\s]+') | Where-Object { $_ -match '^(?i)csx\d+$' } | ForEach-Object { $_.Trim().ToLowerInvariant() })
            }
            if ($isKw -or $isCsx) { $skuMatches.Add([pscustomobject]@{row_number=2+$i;cells=$row}) }
        }
        if ($skuMatches.Count -eq 0) { throw "SKU成本主表中找不到货号：$GoodsCode" }
        if ($skuMatches.Count -gt 1) { throw "货号 $GoodsCode 匹配到多件商品，请先核对当前KW/CSX" }
        $target = $skuMatches[0]
        $rowNumber = [int]$target.row_number
        $currentKw = Get-CellValue $target.cells 0
        $productId = Get-CellValue $target.cells $skuSchema.ProductId.Index
        $batchRows = Get-SheetRows '库存批次' 'A2:M2000'
        $batchRow = Get-FirstEmptyRow $batchRows 1 2 2000
        $batchId = "BATCH-$($Date.ToString('yyyyMMdd'))-$('{0:D4}' -f $batchRow)"
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{goods_code=$currentKw;production_date=([datetime]$ProductionDate).ToString('yyyy-MM-dd');shelf_life="$ShelfLifeValue$ShelfLifeUnit";expiry_date=$computedExpiry.ToString('yyyy-MM-dd')};message='预演成功，未写入飞书'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
        $currentStock = Convert-ToNumber (Get-CellValue $target.cells $skuSchema.CostBookQuantity.Index)
        Set-OneRow '库存批次' "A${batchRow}:M${batchRow}" @(
            $batchId,$productId,$currentKw,(Convert-ToExcelSerial $Date),$currentStock,(Convert-ToExcelSerial ([datetime]$ProductionDate)),[double]$ShelfLifeValue,$ShelfLifeUnit,
            (Convert-ToExcelSerial $computedExpiry),$currentStock,'用户通过飞书补充当前库存日期','在库',$Note
        )
        $oldExpiry = Convert-ToDateValue (Get-CellValue $target.cells $skuSchema.EarliestExpiry.Index)
        $summaryExpiry = if ($null -eq $oldExpiry -or $computedExpiry.Date -lt $oldExpiry.Date) { $computedExpiry } else { $oldExpiry }
        $expiryCell = @(@(@{value=(Convert-ToExcelSerial $summaryExpiry);cell_styles=@{number_format='yyyy-mm-dd';background_color='#FFF8E1';horizontal_alignment='center'}}))
        Set-CellMatrix 'SKU成本主表' "$($skuSchema.EarliestExpiry.Column)${rowNumber}:$($skuSchema.EarliestExpiry.Column)${rowNumber}" $expiryCell
        $reference = "$currentKw $($computedExpiry.ToString('yyyy-MM-dd'))"
        $writeResult = [pscustomobject]@{row=$rowNumber;goods_code=$currentKw;product_id=$productId;batch_id=$batchId;production_date=([datetime]$ProductionDate).ToString('yyyy-MM-dd');shelf_life_value=[double]$ShelfLifeValue;shelf_life_unit=$ShelfLifeUnit;expiry_date=$computedExpiry.ToString('yyyy-MM-dd')}
    }
    'ChangeSkuCode' {
        if (-not $GoodsCode -or $GoodsCode -notmatch '(?i)^(?:kw|csx)\d+$') { throw '原货号必须是有效的 KW 或 CSX 货号' }
        if (-not $NewGoodsCode -or $NewGoodsCode -notmatch '(?i)^kw\d+$') { throw '新货号必须是有效的 KW 货号' }
        $oldCode = $GoodsCode.Trim().ToLowerInvariant()
        $newCode = $NewGoodsCode.Trim().ToLowerInvariant()
        if ($oldCode -eq $newCode) {
            [ordered]@{ok=$true;duplicate=$true;action=$Action;result=[ordered]@{status='unchanged';old_code=$oldCode;new_code=$newCode;changed=$false};formula_verification='success'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }

        $skuRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $oldMatches = [Collections.Generic.List[object]]::new()
        $newMatches = [Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $skuRows.Count; $i++) {
            $row = $skuRows[$i]
            $kw = Get-CellValue $row 0
            $transitionCodes = @([regex]::Split([string](Get-CellValue $row 1), '[/,，、;；\s]+') | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLowerInvariant() })
            $isOldMatch = if ($oldCode -match '^kw\d+$') {
                $kw -and $kw.Trim().ToLowerInvariant() -eq $oldCode
            } else {
                $oldCode -in $transitionCodes
            }
            if ($isOldMatch) { $oldMatches.Add([pscustomobject]@{row_number=2+$i;cells=$row}) }
            if ($kw -and $kw.Trim().ToLowerInvariant() -eq $newCode) { $newMatches.Add([pscustomobject]@{row_number=2+$i;cells=$row}) }
        }
        $sourceMatches = @($oldMatches)
        if ($oldCode -match '^csx\d+$' -and $sourceMatches.Count -gt 1) {
            # 本地成本参考行可能和已上架商品行同时保留；优先选择唯一带稳定商品ID的上架记录。
            $stableMatches = @($sourceMatches | Where-Object { Get-CellValue $_.cells $skuSchema.ProductId.Index })
            if ($stableMatches.Count -eq 1) { $sourceMatches = $stableMatches }
        }
        if ($sourceMatches.Count -eq 0) {
            [ordered]@{ok=$true;duplicate=$false;action=$Action;result=[ordered]@{status='not_found';old_code=$oldCode;new_code=$newCode;changed=$false};formula_verification='success'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
        if ($sourceMatches.Count -gt 1) { throw "原货号 $oldCode 在 SKU成本主表中对应多件上架商品，请先人工核对" }

        $source = $sourceMatches[0]
        $sourceTitle = Get-CellValue $source.cells 2
        $sourceCurrentKw = ([string](Get-CellValue $source.cells 0)).Trim().ToLowerInvariant()
        $otherNewMatches = @($newMatches | Where-Object { [int]$_.row_number -ne [int]$source.row_number })
        if ($sourceCurrentKw -eq $newCode -and $otherNewMatches.Count -eq 0) {
            [ordered]@{ok=$true;duplicate=$true;action=$Action;result=[ordered]@{status='unchanged';old_code=$oldCode;new_code=$newCode;changed=$false;row=[int]$source.row_number};formula_verification='success'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
        if ($otherNewMatches.Count -gt 0 -and -not $RetireExistingNewCode) {
            $target = $otherNewMatches[0]
            $targetTitle = Get-CellValue $target.cells 2
            [ordered]@{
                ok = $true
                duplicate = $false
                action = $Action
                result = [ordered]@{
                    status = 'conflict'
                    changed = $false
                    old_code = $oldCode
                    new_code = $newCode
                    old_title = $sourceTitle
                    new_title = $targetTitle
                    old_row = [int]$source.row_number
                    new_row = [int]$target.row_number
                }
                formula_verification = 'success'
            } | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
        if ($otherNewMatches.Count -gt 1) { throw "新货号 $newCode 在 SKU成本主表中出现多次，请先人工核对" }

        $transitionParts = @([regex]::Split([string](Get-CellValue $source.cells 1), '[/,，、;；\s]+') | Where-Object { $_ -match '^(?i)csx\d+$' } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique)
        $newAliases = $transitionParts -join '、'

        $retiredTarget = if ($otherNewMatches.Count -eq 1) { $otherNewMatches[0] } else { $null }
        $retiredTitle = if ($null -ne $retiredTarget) { Get-CellValue $retiredTarget.cells 2 } else { $null }
        $retiredAliases = $null
        $retiredNote = $null
        if ($null -ne $retiredTarget) {
            $retiredAliases = (@([regex]::Split([string](Get-CellValue $retiredTarget.cells 1), '[/,，、;；\s]+') | Where-Object { $_ -match '^(?i)csx\d+$' } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique)) -join '、'
            $oldRetiredNote = Get-CellValue $retiredTarget.cells $skuSchema.Note.Index
            $retiredNote = ((@($oldRetiredNote, "原商品已下架/移走；当前KW $newCode 自 $($Date.ToString('yyyy-MM-dd')) 起分配给 $sourceTitle") | Where-Object { $_ }) -join '；')
        }

        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=[ordered]@{status=$(if($retiredTarget){'ready_takeover'}else{'ready'});changed=$false;old_code=$oldCode;new_code=$newCode;title=$sourceTitle;transition_codes=$newAliases;row=[int]$source.row_number;retired_title=$retiredTitle;retired_row=$(if($retiredTarget){[int]$retiredTarget.row_number}else{$null})};formula_verification='success'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }

        if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
        $backupPath = Join-Path $stateRoot ("sku_cost_master_auto_backup_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
        [ordered]@{source='SKU成本主表!A2:Z3000';captured_at=[datetime]::Now.ToString('o');cells=$skuRows} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $backupPath -Encoding UTF8

        $rowNumber = [int]$source.row_number
        $sourceRow = @((New-ValueCell $newCode),(New-ValueCell $newAliases)); $sourceMatrix = @(); $sourceMatrix += ,$sourceRow
        $oldSourceNote = Get-CellValue $source.cells $skuSchema.Note.Index
        $ownerConfirmedNote = ((@($oldSourceNote, "当前KW由店主确认：$newCode（$($Date.ToString('yyyy-MM-dd'))）") | Where-Object { $_ }) -join '；')
        $sourceNoteRow = @((New-ValueCell $ownerConfirmedNote)); $sourceNoteMatrix = @(); $sourceNoteMatrix += ,$sourceNoteRow
        $headerRow = @((New-ValueCell 'CSX/过渡货号')); $headerMatrix = @(); $headerMatrix += ,$headerRow
        $operations = [Collections.Generic.List[object]]::new()
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="A${rowNumber}:B${rowNumber}";cells=$sourceMatrix}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.Note.Column)${rowNumber}:$($skuSchema.Note.Column)${rowNumber}";cells=$sourceNoteMatrix}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range='B1:B1';cells=$headerMatrix}})
        if ($null -ne $retiredTarget) {
            $retiredRowNumber = [int]$retiredTarget.row_number
            $retiredAB = @(@{value=''},(New-ValueCell $retiredAliases)); $retiredABMatrix = @(); $retiredABMatrix += ,$retiredAB
            $retiredJ = @((New-ValueCell '已下架/未分配KW')); $retiredJMatrix = @(); $retiredJMatrix += ,$retiredJ
            $retiredN = @((New-ValueCell $retiredNote)); $retiredNMatrix = @(); $retiredNMatrix += ,$retiredN
            $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="A${retiredRowNumber}:B${retiredRowNumber}";cells=$retiredABMatrix}})
            $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.CostStatus.Column)${retiredRowNumber}:$($skuSchema.CostStatus.Column)${retiredRowNumber}";cells=$retiredJMatrix}})
            $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($skuSchema.Note.Column)${retiredRowNumber}:$($skuSchema.Note.Column)${retiredRowNumber}";cells=$retiredNMatrix}})
        }
        Set-AtomicCellOperations $operations.ToArray()

        $checkRows = Get-SheetRows 'SKU成本主表' 'A2:Z3000'
        $checkRow = @($checkRows)[$rowNumber - 2]
        $actualKw = (Get-CellValue $checkRow 0)
        $actualAliases = (Get-CellValue $checkRow 1)
        if (-not $actualKw -or $actualKw.Trim().ToLowerInvariant() -ne $newCode) { throw "SKU改号回读失败：$oldCode → $newCode" }
        $actualAliasSet = @([regex]::Split($actualAliases, '[/,，、;；\s]+') | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLowerInvariant() })
        if ($actualAliasSet | Where-Object { $_ -match '^(?i)kw\d+$' }) { throw "SKU改号后仍残留KW历史别名：$oldCode" }
        if ($null -ne $retiredTarget) {
            $retiredCheck = @($checkRows)[([int]$retiredTarget.row_number) - 2]
            if (Get-CellValue $retiredCheck 0) { throw "原 $newCode 商品的 KW 解绑失败" }
            if ((Get-CellValue $retiredCheck $skuSchema.CostStatus.Index) -ne '已下架/未分配KW') { throw "原 $newCode 商品的解绑状态回读失败" }
        }

        $reference = "$oldCode -> $newCode"
        $writeResult = [pscustomobject]@{
            status = 'changed'
            changed = $true
            old_code = $oldCode
            new_code = $newCode
            title = $sourceTitle
            transition_codes = $actualAliases
            row = $rowNumber
            retired_title = $retiredTitle
            retired_row = $(if($retiredTarget){[int]$retiredTarget.row_number}else{$null})
        }
    }
    'CashExpense' {
        Assert-Positive $Amount '其他现金支出'
        $writeResult = Update-MonthlyCell 5 'F' ([double]$Amount) $true "其他现金支出 $([double]$Amount)：$Note"
        $reference = [string]([double]$Amount)
    }
    'OperatingExpense' {
        Assert-Positive $Amount '其他经营费用'
        $writeResult = Invoke-OperatingExpenseWrite $Date ([double]$Amount) $Note -PreviewOnly:$DryRun
        $reference = [string]([double]$Amount)
        if ($DryRun) {
            [ordered]@{ok=$true;dry_run=$true;action=$Action;result=$writeResult;formula_verification='not_required'} | ConvertTo-Json -Depth 8
            Complete-CurrentFinanceTransaction
            exit 0
        }
    }
    'OwnerWithdrawal' {
        Assert-Positive $Amount '老板取款'
        $writeResult = Update-MonthlyCell 7 'H' ([double]$Amount) $true "老板取款 $([double]$Amount)：$Note"
        $reference = [string]([double]$Amount)
    }
    'OwnerLabor' {
        Assert-Positive $Amount '老板人工'
        $writeResult = Update-MonthlyCell 8 'I' ([double]$Amount) $false "老板人工设为 $([double]$Amount)：$Note"
        $reference = [string]([double]$Amount)
    }
    'MarketRent' {
        Assert-Positive $Amount '市场租金机会成本'
        $writeResult = Update-MonthlyCell 9 'J' ([double]$Amount) $false "市场租金机会成本设为 $([double]$Amount)：$Note"
        $reference = [string]([double]$Amount)
    }
    'BeginningCash' {
        if ($null -eq $Amount) { throw '期初现金必须提供金额' }
        $writeResult = Update-MonthlyCell 10 'K' ([double]$Amount) $false "期初现金设为 $([double]$Amount)：$Note"
        $reference = [string]([double]$Amount)
    }
    'ActualReceipt' {
        Assert-Positive $Amount '实际到账'
        if ($OrderNo) { throw 'ActualReceipt 只用于平台资金实际到账，不能携带订单号或用于订单退款；请改用 SetOrderActuals' }
        $mapping = @{'美团'=@{index=1;column='B'};'淘宝闪购'=@{index=2;column='C'};'京东到家'=@{index=3;column='D'};'线下'=@{index=4;column='E'}}
        $target = $mapping[$Platform]
        $writeResult = Update-MonthlyCell $target.index $target.column ([double]$Amount) $true "$Platform 实际到账 $([double]$Amount)：$Note"
        $reference = "$Platform $([double]$Amount)"
    }
    'AssetAmortization' {
        Assert-Positive $Amount '摊销折旧'
        $writeResult = Update-MonthlyCell 12 'M' ([double]$Amount) $false "品牌/装修/设备摊销折旧设为 $([double]$Amount)：$Note"
        $reference = [string]([double]$Amount)
    }
}

Assert-FormulaRange '驾驶舱' 'C2:C19'
if ($Action -in @('AddIntervention','RemoveIntervention','RestoreInventory','ConfirmSelfDelivery','SetExternalDeliveryFee','SetOrderDeliveryFee','SetOrderActuals','AddOrderGift','SetRefundDisposition')) {
    Assert-FormulaRange '干预订单' 'I2:I1001'
    Assert-FormulaRange '订单汇总' 'E2:E2001,K2:L2001'
    Assert-FormulaRange '订单商品' 'E2:E4001,K2:K4001,P2:P4001'
    Assert-FormulaRange '库存快照' 'M2:M2000'
}
if ($Action -eq 'Replenishment') { Assert-FormulaRange '补货记录' 'D2:E1001,G2:G1001' }
if ($Action -eq 'OfflineSale') { Assert-FormulaRange '线下订单' 'G2:G1001' }
if ($Action -eq 'SetSkuCost' -and $OrderNo) {
    Assert-FormulaRange '订单汇总' 'E2:E2001,K2:L2001'
    Assert-FormulaRange '订单商品' 'E2:E4001,K2:K4001,P2:P4001'
}

$ledgerResult = $null
if ($Action -in @('AddIntervention','RemoveIntervention','SetOrderActuals','AddOrderGift','SetRefundDisposition') -and $null -eq $costResult) {
    $costResult = Invoke-CostSnapshotRebuild
}
if ($Action -in @('AddIntervention','RemoveIntervention','ConfirmSelfDelivery','SetExternalDeliveryFee','SetOrderDeliveryFee','SetOrderActuals','OfflineSale','AddOrderGift','SetRefundDisposition')) {
    # Order-local changes only require rebuilding their own month. Cost-basis changes
    # can affect FIFO snapshots across multiple months and deliberately keep the full rebuild.
    $ledgerTargetMonth = $null
    $monthLocalActions = @('AddIntervention','RemoveIntervention','ConfirmSelfDelivery','SetExternalDeliveryFee','SetOrderDeliveryFee','SetOrderActuals','OfflineSale','AddOrderGift','SetRefundDisposition')
    if ($Action -in $monthLocalActions) {
        $candidateDate = $null
        if ($null -ne $writeResult -and $null -ne $writeResult.PSObject.Properties['date']) {
            $parsedDate = [datetime]::MinValue
            if ([datetime]::TryParse([string]$writeResult.date,[ref]$parsedDate)) { $candidateDate = $parsedDate }
        }
        if ($null -eq $candidateDate -and $Action -eq 'OfflineSale') { $candidateDate = $Date }
        if ($null -ne $candidateDate) { $ledgerTargetMonth = [datetime]$candidateDate }
    }
    $ledgerResult = Invoke-DailyLedgerSync $ledgerTargetMonth
}
if ($Action -eq 'SetSkuCost' -and $null -ne $explicitHistoricalOrderDate) {
    # 只有店主明确指定历史订单时才刷新其所属月份；普通 SKU 成本修改不得改写历史订单或简账。
    $ledgerResult = Invoke-DailyLedgerSync $explicitHistoricalOrderDate
}

Save-ProcessedEvent $EventId $Action $reference
$finalResult = [ordered]@{
    ok = $true
    duplicate = $false
    event_id = $EventId
    action = $Action
    reference = $reference
    result = $writeResult
    cost_snapshot_rebuild = $costResult
    daily_ledger_sync = $ledgerResult
    formula_verification = 'success'
    workbook_url = $meta.workbook_url
    processed_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
}
Complete-CurrentFinanceTransaction
$script:refundStateBackup = $null
$script:orderActualOverrideStateBackup = $null
$finalResult | ConvertTo-Json -Depth 8
