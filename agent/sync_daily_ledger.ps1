param(
    [datetime]$ThroughDate = [datetime]::Today,
    [string]$MonthFilter,
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$sourceMetaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
$ledgerMetaPath = Join-Path $ProjectRoot 'work\daily_ledger_meta.json'
if (-not (Test-Path -LiteralPath $sourceMetaPath)) { throw '缺少经营驾驶舱元数据' }
if (-not (Test-Path -LiteralPath $ledgerMetaPath)) { throw '缺少每日经营简账元数据；请先创建简账' }
$sourceMeta = [System.IO.File]::ReadAllText($sourceMetaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [System.IO.File]::ReadAllText($ledgerMetaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$ledgerStartDate = [datetime]'2026-08-24'
$selectedMonthKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
if ($MonthFilter) {
    foreach ($monthKey in @($MonthFilter -split '[,，;；\s]+' | Where-Object { $_ })) {
        if ($monthKey -notmatch '^20\d{2}-(?:0[1-9]|1[0-2])$') {
            [Console]::Error.WriteLine("月份筛选格式无效：$monthKey（应为 yyyy-MM）")
            exit 1
        }
        $null = $selectedMonthKeys.Add($monthKey)
    }
}
$larkCli = Join-Path $PSScriptRoot 'lark_cli_noexit.ps1'
. (Join-Path $PSScriptRoot 'finance_transaction.ps1')
. (Join-Path $PSScriptRoot 'intervention_record.ps1')
. (Join-Path $PSScriptRoot 'refund_adjustments.ps1')
. (Join-Path $PSScriptRoot 'expense_accounting.ps1')
$stateRoot = Join-Path $ProjectRoot 'agent\state'
$script:financeTransaction = Join-FinanceTransaction $stateRoot
$script:financeTransactionOwner = $false
$script:financeTransactionClosed = $false
if ($null -eq $script:financeTransaction) {
    $script:financeTransaction = New-FinanceTransaction $stateRoot '刷新每日经营简账' $null -AcquireLock -WorkbookTokens @{source=$sourceToken;ledger=$ledgerToken} -LarkCli $larkCli
    $script:financeTransactionOwner = $true
}

function Close-LedgerTransaction([bool]$Succeeded, [string]$Reason = '') {
    if ($script:financeTransactionClosed -or $null -eq $script:financeTransaction) { return }
    if ($script:financeTransactionOwner) {
        if ($Succeeded) { Complete-FinanceTransaction $script:financeTransaction }
        else {
            $rollback = Invoke-FinanceTransactionRollback $script:financeTransaction @{source=$sourceToken;ledger=$ledgerToken} $larkCli $Reason
            if (-not $rollback.ok) { [Console]::Error.WriteLine("每日简账回滚待恢复：事务 $($rollback.transaction_id)；$($rollback.error)") }
        }
    }
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
}

trap {
    $reason = [string]$_.Exception.Message
    Close-LedgerTransaction $false $reason
    [Console]::Error.WriteLine($reason)
    exit 1
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    $valueProperty = $cell.PSObject.Properties['value']
    if ($null -eq $valueProperty -or $null -eq $valueProperty.Value) { return $null }
    return [string]$valueProperty.Value
}

function Convert-ToNullableNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}

function Convert-ToDate($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $text = [string]$Value
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.Date }
    $number = Convert-ToNullableNumber $Value
    if ($null -ne $number) { return ([datetime]'1899-12-30').AddDays($number).Date }
    return $null
}

function Convert-ToExcelSerial([datetime]$Value) { return ($Value.Date - [datetime]'1899-12-30').TotalDays }
function New-ValueCell($Value) { if ($null -eq $Value -or [string]$Value -eq '') { return @{} }; return @{value=$Value} }
function New-FormulaCell([string]$Formula) { if (-not $Formula) { return @{} }; return @{formula=$Formula} }

function Get-PlatformSortRank([string]$Platform) {
    switch ($Platform) {
        '美团' { return 10 }
        '饿了么' { return 20 }
        '京东到家' { return 30 }
        '淘宝闪购' { return 40 }
        '线下' { return 50 }
        default { return 90 }
    }
}

function Get-DailyOrderSequence([string]$DailyOrderNo) {
    if ($DailyOrderNo -and $DailyOrderNo -match '(\d+)\s*号单') { return [int]$Matches[1] }
    return [int]::MaxValue
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

function Get-SheetRows([string]$Token, [string]$SheetName, [string]$Range) {
    $result = Invoke-LarkReadJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $SheetName --range $Range --include 'value,formula' --max-chars 5000000 --format json } "读取飞书失败：$SheetName"
    if ($result.data.has_more) { throw "飞书读取结果不完整：$SheetName" }
    # Wrap rows in an object so PowerShell cannot flatten a one-row result into
    # its individual cells when the function writes to the success stream.
    return [pscustomobject]@{rows=[object[]]@($result.data.ranges[0].cells);row_indices=[int[]]@($result.data.ranges[0].row_indices)}
}

function Set-CellMatrix([string]$SheetName, [string]$Range, $Rows) {
    if (@($Rows).Count -eq 0) { return }
    Register-FinanceRangeBeforeImage $script:financeTransaction 'ledger' $ledgerToken $SheetName $Range $larkCli
    $payload = ConvertTo-Json -InputObject $Rows -Depth 12 -Compress
    $raw = $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $ledgerToken --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入飞书失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "写入飞书失败：$SheetName $Range" }
}

function Set-SourceCellMatrix([string]$SheetName, [string]$Range, $Rows) {
    if (@($Rows).Count -eq 0) { return }
    Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $sourceToken $SheetName $Range $larkCli
    $payload = ConvertTo-Json -InputObject $Rows -Depth 12 -Compress
    $raw = $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $sourceToken --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入经营主表失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "写入经营主表失败：$SheetName $Range" }
}

function Clear-SheetData([string]$SheetName, [string]$LastColumn) {
    Register-FinanceRangeBeforeImage $script:financeTransaction 'ledger' $ledgerToken $SheetName "A2:${LastColumn}1001" $larkCli
    $ranges = ConvertTo-Json -InputObject @("$SheetName!A2:${LastColumn}1001") -Compress
    $raw = $ranges | & $larkCli sheets +cells-batch-clear --as user --spreadsheet-token $ledgerToken --ranges - --scope content --yes --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "清理旧简账数据失败：$SheetName" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "清理旧简账数据失败：$SheetName" }
}

function Assert-FormulaRange([string]$SheetName, [string]$Range) {
    foreach ($area in @($Range -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
        $raw = & $larkCli sheets +formula-verify --as user --spreadsheet-token $ledgerToken --sheet-name $SheetName --range $area --format json --exit-on-error 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) { throw "公式校验失败：$SheetName $area；$($raw.Trim())" }
        $result = $raw | ConvertFrom-Json
        if ($result.data.status -ne 'success') { throw "公式存在错误：$SheetName $area" }
    }
}

function Get-MonthSheetPrefix([datetime]$Month) {
    if ($Month.Year -eq 2026) { return "$($Month.Month)月" }
    return "$($Month.Year)年$($Month.Month)月"
}

function Invoke-TablePutWithStyles([string]$SheetsJson, [string]$StylesJson) {
    # lark-cli only allows one composite JSON flag to read from stdin. On Windows,
    # passing --styles as an inline argument can lose its JSON quotes when the
    # PowerShell wrapper starts Node. Keep styles on stdin and put sheets in a
    # short-lived @file so neither payload is interpreted by the command shell.
    $stateDirectory = Get-Item -LiteralPath $stateRoot -Force
    $requestRoot = $stateDirectory.FullName
    if ($stateDirectory.LinkType -and $stateDirectory.Target) {
        $linkTarget = [string]@($stateDirectory.Target)[0]
        if (-not [IO.Path]::IsPathRooted($linkTarget)) { $linkTarget = Join-Path $stateDirectory.Parent.FullName $linkTarget }
        $requestRoot = [IO.Path]::GetFullPath($linkTarget)
    }
    $requestFileName = "daily_ledger_sheets_{0}.json" -f ([guid]::NewGuid().ToString('N'))
    $requestPath = Join-Path $requestRoot $requestFileName
    try {
        [IO.File]::WriteAllText($requestPath, $SheetsJson, [Text.UTF8Encoding]::new($false))
        # @file is intentionally restricted by lark-cli to the current directory.
        # Execute from the D: state directory and pass a relative path.
        Push-Location -LiteralPath $requestRoot
        try {
            $sheetsFileArgument = "@./$requestFileName"
            $raw = $StylesJson | & $larkCli sheets +table-put --as user --spreadsheet-token $ledgerToken --sheets $sheetsFileArgument --styles - --format json | Out-String
            $exitCode = $LASTEXITCODE
            return [pscustomobject]@{raw=$raw;exit_code=$exitCode}
        } finally {
            Pop-Location
        }
    } finally {
        if (Test-Path -LiteralPath $requestPath) { Remove-Item -LiteralPath $requestPath -Force }
    }
}

function Invoke-DropdownSetJson([string]$SheetName, [string]$Range, [string]$OptionsJson, [string]$ColorsJson) {
    $stateDirectory = Get-Item -LiteralPath $stateRoot -Force
    $requestRoot = $stateDirectory.FullName
    if ($stateDirectory.LinkType -and $stateDirectory.Target) {
        $linkTarget = [string]@($stateDirectory.Target)[0]
        if (-not [IO.Path]::IsPathRooted($linkTarget)) { $linkTarget = Join-Path $stateDirectory.Parent.FullName $linkTarget }
        $requestRoot = [IO.Path]::GetFullPath($linkTarget)
    }
    $requestFileName = "daily_ledger_dropdown_colors_{0}.json" -f ([guid]::NewGuid().ToString('N'))
    $requestPath = Join-Path $requestRoot $requestFileName
    try {
        [IO.File]::WriteAllText($requestPath, $ColorsJson, [Text.UTF8Encoding]::new($false))
        Push-Location -LiteralPath $requestRoot
        try {
            $colorsFileArgument = "@./$requestFileName"
            $raw = $OptionsJson | & $larkCli sheets +dropdown-set --as user --spreadsheet-token $ledgerToken --sheet-name $SheetName --range $Range --options - --colors $colorsFileArgument --format json | Out-String
            $exitCode = $LASTEXITCODE
            return [pscustomobject]@{raw=$raw;exit_code=$exitCode}
        } finally {
            Pop-Location
        }
    } finally {
        if (Test-Path -LiteralPath $requestPath) { Remove-Item -LiteralPath $requestPath -Force }
    }
}

function Set-DateDisplayFormat([string]$SheetName, [string]$Range) {
    $raw = & $larkCli sheets +cells-set-style --as user --spreadsheet-token $ledgerToken --sheet-name $SheetName --range $Range --number-format 'yyyy-mm-dd' --horizontal-alignment center --format json | Out-String
    $exitCode = $LASTEXITCODE
    try { $result = $raw | ConvertFrom-Json } catch { throw "设置日期格式失败：$SheetName $Range；返回内容无法解析" }
    if ($exitCode -ne 0 -or -not $result.ok) {
        $detail = if ($result.error -and $result.error.message) { [string]$result.error.message } else { $raw.Trim() }
        throw "设置日期格式失败：$SheetName $Range；$detail"
    }
}

function Ensure-MonthSheets([datetime]$Month) {
    $prefix = Get-MonthSheetPrefix $Month
    $incomeName = "${prefix}收入"
    $expenseName = "${prefix}支出"
    $infoRaw = & $larkCli sheets +workbook-info --as user --spreadsheet-token $ledgerToken --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '读取每日经营简账结构失败' }
    $info = $infoRaw | ConvertFrom-Json
    $names = @($info.data.sheets | ForEach-Object { $_.sheet_name })
    $missing = @()
    if ($incomeName -notin $names) {
        $missing += @{name=$incomeName;columns=@('当日单号','订单编号','日期','扣佣后收入','货物成本','配送费','这一单赚/亏','售出货品（KW）','赠品（KW）','当天赚/亏','备注');data=@();dtypes=@{'日期'='datetime64[ns]';'扣佣后收入'='float64';'货物成本'='float64';'配送费'='float64';'这一单赚/亏'='float64';'当天赚/亏'='float64'};formats=@{'日期'='yyyy-mm-dd';'扣佣后收入'='#,##0.00';'货物成本'='#,##0.00';'配送费'='#,##0.00';'这一单赚/亏'='#,##0.00;[Red]-#,##0.00';'当天赚/亏'='#,##0.00;[Red]-#,##0.00'}}
    }
    if ($expenseName -notin $names) {
        $missing += @{name=$expenseName;columns=@('日期','项目','费用','订单编号','来源','备注');data=@();dtypes=@{'日期'='datetime64[ns]';'费用'='float64'};formats=@{'日期'='yyyy-mm-dd';'费用'='#,##0.00'}}
    }
    if ($missing.Count -gt 0) {
        $styleItems = @()
        foreach ($sheet in $missing) {
            if ($sheet.name.EndsWith('收入')) {
                $styleItems += @{name=$sheet.name;cell_styles=@(@{range='A1:K1';background_color='#173A2D';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center'},@{range='A2:K1001';background_color='#F7FAF8'},@{range='G2:G1001';background_color='#E8F1EC';font_weight='bold'},@{range='I2:I1001';background_color='#EAF5EE'},@{range='J2:J1001';background_color='#DCEBE2';font_weight='bold'},@{range='K2:K1001';background_color='#FFF8E1'});col_sizes=@(@{range='A';type='pixel';size=120},@{range='B';type='pixel';size=205},@{range='C';type='pixel';size=110},@{range='D:G';type='pixel';size=120},@{range='H';type='pixel';size=300},@{range='I';type='pixel';size=220},@{range='J';type='pixel';size=125},@{range='K';type='pixel';size=260})}
            } else {
                $styleItems += @{name=$sheet.name;cell_styles=@(@{range='A1:F1';background_color='#8A5A13';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center'},@{range='A2:F1001';background_color='#FFF8E1'},@{range='D2:E1001';background_color='#F4F7F5'});col_sizes=@(@{range='A';type='pixel';size=110},@{range='B';type='pixel';size=160},@{range='C';type='pixel';size=110},@{range='D';type='pixel';size=205},@{range='E';type='pixel';size=105},@{range='F';type='pixel';size=300})}
            }
        }
        $payload = @{sheets=$missing} | ConvertTo-Json -Depth 20 -Compress
        $styleJson = @{styles=$styleItems} | ConvertTo-Json -Depth 20 -Compress
        $response = Invoke-TablePutWithStyles $payload $styleJson
        $raw = [string]$response.raw
        $result = $raw | ConvertFrom-Json
        if ([int]$response.exit_code -ne 0 -or -not $result.ok) {
            $errorProperty = $result.PSObject.Properties['error']
            $detail = if ($null -ne $errorProperty -and $null -ne $errorProperty.Value -and $errorProperty.Value.message) { [string]$errorProperty.Value.message } else { $raw.Trim() }
            throw "创建月份页签失败：$prefix；$detail"
        }
        foreach ($sheet in $missing) {
            $null = & $larkCli sheets +dim-freeze --as user --spreadsheet-token $ledgerToken --sheet-name $sheet.name --dimension row --count 1 --format json | Out-String
            if ($LASTEXITCODE -ne 0) { throw "冻结月份页签表头失败：$($sheet.name)" }
            $filterRange = if ($sheet.name.EndsWith('收入')) { 'A1:K1001' } else { 'A1:F1001' }
            $filterProperties = ConvertTo-Json -InputObject @{range=$filterRange;rules=@()} -Compress
            $null = $filterProperties | & $larkCli sheets +filter-create --as user --spreadsheet-token $ledgerToken --sheet-name $sheet.name --range $filterRange --properties - --format json | Out-String
            if ($LASTEXITCODE -ne 0) { throw "创建月份页签筛选器失败：$($sheet.name)" }
        }
    }
    # Reapply this idempotently even when the sheet already exists. This repairs a
    # month sheet left behind by a failed setup after table creation but before
    # data validation was attached.
    $options = '["平台自动识别","美团自动识别","待补录","人工补录","自配送确认","手工"]'
    $colors = '["#2563EB","#16A34A","#F59E0B","#7C3AED","#6B7280","#8A5A13"]'
    $dropdownResponse = Invoke-DropdownSetJson $expenseName 'E2:E1001' $options $colors
    $dropdownRaw = [string]$dropdownResponse.raw
    $dropdownResult = $dropdownRaw | ConvertFrom-Json
    if ([int]$dropdownResponse.exit_code -ne 0 -or -not $dropdownResult.ok) {
        $dropdownError = $dropdownResult.PSObject.Properties['error']
        $dropdownDetail = if ($null -ne $dropdownError -and $null -ne $dropdownError.Value -and $dropdownError.Value.message) { [string]$dropdownError.Value.message } else { $dropdownRaw.Trim() }
        throw "创建月份支出来源下拉失败：$expenseName；$dropdownDetail"
    }
    # 日期以 Excel/飞书日期序列存储，必须为现有和新建月份页签反复补上
    # 显示格式，否则一次全量重写后会直接显示 46296 之类的内部序列号。
    Set-DateDisplayFormat $incomeName 'C2:C1001'
    Set-DateDisplayFormat $expenseName 'A2:A1001'
    return [pscustomobject]@{month=$Month;income=$incomeName;expense=$expenseName}
}

$orderRange = Get-SheetRows $sourceToken '订单汇总' 'A2:Q2001'
$orderRows = $orderRange.rows
$orderRowNumbers = @($orderRange.row_indices)
$orderSourceRowByOrderNo = @{}
for ($orderIndex = 0; $orderIndex -lt $orderRows.Count; $orderIndex++) {
    $sourceOrderNo = ([string](Get-CellValue $orderRows[$orderIndex] 2)).Trim()
    if (-not $sourceOrderNo) { continue }
    if ($orderSourceRowByOrderNo.ContainsKey($sourceOrderNo)) { throw "订单汇总存在重复订单：$sourceOrderNo" }
    $sourceRowNumber = if ($orderRowNumbers.Count -eq $orderRows.Count) { [int]$orderRowNumbers[$orderIndex] } else { 2 + $orderIndex }
    $orderSourceRowByOrderNo[$sourceOrderNo] = $sourceRowNumber
}
$itemRows = (Get-SheetRows $sourceToken '订单商品' 'A2:N4001').rows
$inventoryRows = (Get-SheetRows $sourceToken '库存快照' 'B2:G2000').rows
$interventionRows = @((Get-SheetRows $sourceToken '干预订单' 'A2:M1001').rows | Where-Object { Test-ActiveInterventionOrderNo (Get-CellValue $_ 3) })
$offlineRows = (Get-SheetRows $sourceToken '线下订单' 'A2:J1001').rows
$giftRows = (Get-SheetRows $sourceToken '赠品记录' 'A2:L2001').rows

# ThroughDate is the date whose summary is requested, not a destructive cutoff for
# the month ledger. Historical replays/tests must retain newer source rows.
$dataThroughDate = $ThroughDate.Date
foreach ($datedRows in @($orderRows, $interventionRows, $offlineRows)) {
    foreach ($datedRow in @($datedRows)) {
        $sourceDate = Convert-ToDate (Get-CellValue $datedRow 0)
        if ($null -ne $sourceDate -and $sourceDate.Date -gt $dataThroughDate) {
            $dataThroughDate = $sourceDate.Date
        }
    }
}

$codeToKw = @{}
foreach ($row in $inventoryRows) {
    $currentCode = Get-CellValue $row 0
    $shelfCode = Get-CellValue $row 5
    if ($currentCode -and $currentCode -match '^(?i)kw\d+$') {
        $codeToKw[$currentCode] = $currentCode.ToLowerInvariant()
        if ($shelfCode) { $codeToKw[$shelfCode] = $currentCode.ToLowerInvariant() }
    }
}

$soldByOrder = @{}
$missingCostCodesByOrder = @{}
$remainingReturnedByOrderCode = @{}
$refundNotesByOrder = @{}
foreach ($refundEvent in @(Get-RefundEvents $ProjectRoot)) {
    $refundParts = @()
    foreach ($refundItem in @($refundEvent.items)) {
        $code = ([string]$refundItem.goods_code).Trim().ToLowerInvariant()
        $dispositionText = switch ([string]$refundItem.disposition) {
            'returned_restocked' { '已退回入库，已撤销成本' }
            'refund_no_return' { '未退货，保留成本' }
            'returned_damaged' { '退回报损，保留成本' }
            default { '退货状态待确认，暂保留成本' }
        }
        $refundParts += "$code×$($refundItem.quantity) $dispositionText"
        if ($refundItem.disposition -eq 'returned_restocked') {
            $key = "$($refundEvent.order_no)|$code"
            if (-not $remainingReturnedByOrderCode.ContainsKey($key)) { $remainingReturnedByOrderCode[$key] = 0.0 }
            $remainingReturnedByOrderCode[$key] += [double]$refundItem.quantity
        }
    }
    $line = "退款¥$([double]$refundEvent.customer_refund_amount)：$($refundParts -join '、')"
    if (-not $refundNotesByOrder.ContainsKey([string]$refundEvent.order_no)) { $refundNotesByOrder[[string]$refundEvent.order_no] = @() }
    $refundNotesByOrder[[string]$refundEvent.order_no] += $line
}
foreach ($row in $itemRows) {
    $orderNo = Get-CellValue $row 2
    $rawCode = Get-CellValue $row 5
    $quantity = Convert-ToNullableNumber (Get-CellValue $row 8)
    if (-not $orderNo -or $null -eq $quantity) { continue }
    $normalizedRawCode = ([string]$rawCode).Trim().ToLowerInvariant()
    $returnKey = "$orderNo|$normalizedRawCode"
    $effectiveQuantity = [double]$quantity
    if ($remainingReturnedByOrderCode.ContainsKey($returnKey) -and [double]$remainingReturnedByOrderCode[$returnKey] -gt 0) {
        $returnedQuantity = [math]::Min($effectiveQuantity, [double]$remainingReturnedByOrderCode[$returnKey])
        $effectiveQuantity -= $returnedQuantity
        $remainingReturnedByOrderCode[$returnKey] = [double]$remainingReturnedByOrderCode[$returnKey] - $returnedQuantity
    }
    $displayCode = if ($rawCode -and $rawCode -match '^(?i)kw\d+$') { $rawCode.ToLowerInvariant() } elseif ($rawCode -and $codeToKw.ContainsKey($rawCode)) { $codeToKw[$rawCode] } elseif ($rawCode) { $rawCode.ToLowerInvariant() } else { $null }
    if ($displayCode -and $effectiveQuantity -gt 0) {
        if (-not $soldByOrder.ContainsKey($orderNo)) { $soldByOrder[$orderNo] = @{} }
        if (-not $soldByOrder[$orderNo].ContainsKey($displayCode)) { $soldByOrder[$orderNo][$displayCode] = 0.0 }
        $soldByOrder[$orderNo][$displayCode] += $effectiveQuantity
    }
    $unitCost = Convert-ToNullableNumber (Get-CellValue $row 11)
    $isNaturalItem = (Get-CellValue $row 4) -eq '自然'
    $isCancelledItem = (Get-CellValue $row 3) -eq '已取消'
    if ($null -eq $unitCost -and $isNaturalItem -and -not $isCancelledItem -and $effectiveQuantity -gt 0) {
        $title = (([string](Get-CellValue $row 6)).Trim() -replace '\s*\([^)]*\)\s*$','').Trim()
        $spec = ([string](Get-CellValue $row 7)).Trim()
        if ($title.Length -gt 34) { $title = $title.Substring(0,34) + '…' }
        $missingCostKey = if ($displayCode -and $displayCode -match '^(?i)(?:kw|csx)\d+$') {
            $displayCode
        } elseif ($displayCode -and $title) {
            "${displayCode}「${title}」"
        } elseif ($displayCode) {
            $displayCode
        } elseif ($title -and $spec) {
            "未识别货号「$($title)，$spec」"
        } elseif ($title) {
            "未识别货号「$title」"
        } else {
            '未识别货号的商品'
        }
        if (-not $missingCostCodesByOrder.ContainsKey($orderNo)) { $missingCostCodesByOrder[$orderNo] = @{} }
        $missingCostCodesByOrder[$orderNo][$missingCostKey] = $true
    }
}

$giftTextByOrder = @{}
$giftCostByOrder = @{}
$pendingGiftCodesByOrder = @{}
foreach ($row in $giftRows) {
    $dailyOrderNo = Get-CellValue $row 2
    $orderNo = Get-CellValue $row 3
    $rawCode = ([string](Get-CellValue $row 4)).Trim().ToLowerInvariant()
    $quantity = Convert-ToNullableNumber (Get-CellValue $row 5)
    $giftCost = Convert-ToNullableNumber (Get-CellValue $row 7)
    $costStatus = Get-CellValue $row 8
    if ($dailyOrderNo -in @('待同步','订单已取消待处理')) { continue }
    if (-not $orderNo -or -not $rawCode -or $null -eq $quantity -or $quantity -le 0) { continue }
    if (-not $giftTextByOrder.ContainsKey($orderNo)) { $giftTextByOrder[$orderNo] = @{} }
    if (-not $giftTextByOrder[$orderNo].ContainsKey($rawCode)) { $giftTextByOrder[$orderNo][$rawCode] = 0.0 }
    $giftTextByOrder[$orderNo][$rawCode] += [double]$quantity
    if ($null -eq $giftCost -or $costStatus -eq '成本待核对') {
        if (-not $pendingGiftCodesByOrder.ContainsKey($orderNo)) { $pendingGiftCodesByOrder[$orderNo] = @{} }
        $pendingGiftCodesByOrder[$orderNo][$rawCode] = $true
    } else {
        if (-not $giftCostByOrder.ContainsKey($orderNo)) { $giftCostByOrder[$orderNo] = 0.0 }
        $giftCostByOrder[$orderNo] += [double]$giftCost
    }
}

function Get-GiftText([string]$OrderNo) {
    if (-not $OrderNo -or -not $giftTextByOrder.ContainsKey($OrderNo)) { return '' }
    $parts = @()
    foreach ($code in @($giftTextByOrder[$OrderNo].Keys | Sort-Object)) {
        $quantity = [double]$giftTextByOrder[$OrderNo][$code]
        $quantityText = if ([math]::Abs($quantity - [math]::Round($quantity)) -lt 0.000001) { [string][int][math]::Round($quantity) } else { $quantity.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture) }
        $parts += "$code×$quantityText"
    }
    return $parts -join '、'
}

$month = [datetime]::new($ledgerStartDate.Year,$ledgerStartDate.Month,1)
$monthSheets = @()
$throughMonth = [datetime]::new($dataThroughDate.Year,$dataThroughDate.Month,1)
while ($month -le $throughMonth) {
    $monthKey = $month.ToString('yyyy-MM')
    if ($selectedMonthKeys.Count -eq 0 -or $selectedMonthKeys.Contains($monthKey)) {
        $monthSheets += Ensure-MonthSheets $month
    }
    $month = $month.AddMonths(1)
}
if ($selectedMonthKeys.Count -gt 0 -and $monthSheets.Count -eq 0) {
    throw "月份筛选不在可刷新范围内：$MonthFilter"
}

$interventionOrderSet = @{}
foreach ($row in $interventionRows) {
    $orderNo = Get-CellValue $row 3
    if ($orderNo) { $interventionOrderSet[$orderNo] = $true }
}

$monthSettlementRowsByKey = @{}
$monthSettlementRange = Get-SheetRows $sourceToken '月度结算' 'A2:N25'
for ($settlementIndex = 0; $settlementIndex -lt $monthSettlementRange.rows.Count; $settlementIndex++) {
    $settlementMonth = Convert-ToDate (Get-CellValue $monthSettlementRange.rows[$settlementIndex] 0)
    if ($null -eq $settlementMonth) { continue }
    $settlementRowNumber = if ($monthSettlementRange.row_indices.Count -eq $monthSettlementRange.rows.Count) { [int]$monthSettlementRange.row_indices[$settlementIndex] } else { 2 + $settlementIndex }
    $monthSettlementRowsByKey[$settlementMonth.ToString('yyyy-MM')] = $settlementRowNumber
}

$summaries = @()
$requestedDateSummary = $null
$autoReconciledDeliveryFees = [Collections.Generic.List[object]]::new()
foreach ($monthSheet in $monthSheets) {
    $monthStart = $monthSheet.month
    $monthEnd = $monthStart.AddMonths(1).AddDays(-1)
    if ($monthStart -eq [datetime]::new(2026,8,1)) { $monthStart = $ledgerStartDate }

    $oldIncomeRows = (Get-SheetRows $ledgerToken $monthSheet.income 'A2:K1001').rows
    $oldIncomeByOrder = @{}
    foreach ($row in $oldIncomeRows) {
        $orderNo = Get-CellValue $row 1
        if (-not $orderNo) { continue }
        $oldIncomeByOrder[$orderNo] = [pscustomobject]@{
            delivery = Convert-ToNullableNumber (Get-CellValue $row 5)
            note = Get-CellValue $row 10
        }
    }

    $naturalOrders = @()
    foreach ($row in $orderRows) {
        $date = Convert-ToDate (Get-CellValue $row 0)
        $orderNo = Get-CellValue $row 2
        if ($null -eq $date -or $date -lt $monthStart -or $date -gt $monthEnd -or $date -gt $dataThroughDate) { continue }
        if ((Get-CellValue $row 3) -eq '已取消' -or (Get-CellValue $row 4) -ne '自然') { continue }
        if ($orderNo -and $interventionOrderSet.ContainsKey($orderNo)) { throw "收入筛选保护触发：干预单 $orderNo 被标记为自然单" }
        $platform = Get-CellValue $row 1
        if (-not $platform) { $platform = '美团' }
        $dailyOrderNo = Get-CellValue $row 16
        $naturalOrders += [pscustomobject]@{
            source_row=$row; date=$date; order_no=$orderNo; platform=$platform; daily_order_no=$dailyOrderNo
            platform_rank=(Get-PlatformSortRank $platform); order_sequence=(Get-DailyOrderSequence $dailyOrderNo)
        }
    }
    foreach ($row in $offlineRows) {
        $date = Convert-ToDate (Get-CellValue $row 0)
        $orderNo = Get-CellValue $row 2
        if ($null -eq $date -or $date -lt $monthStart -or $date -gt $monthEnd -or $date -gt $dataThroughDate -or -not $orderNo) { continue }
        $dailyOrderNo = Get-CellValue $row 1
        $naturalOrders += [pscustomobject]@{
            source_row=$row; date=$date; order_no=$orderNo; is_offline=$true; platform='线下'; daily_order_no=$dailyOrderNo
            platform_rank=(Get-PlatformSortRank '线下'); order_sequence=(Get-DailyOrderSequence $dailyOrderNo)
        }
    }
    $naturalOrders = @($naturalOrders | Sort-Object date, platform_rank, platform, order_sequence, order_no)
    if ($naturalOrders.Count -gt 1000) { throw "$($monthSheet.income) 超过1000行容量" }

    $incomeMatrix = @()
    $sequenceByDay = @{}
    for ($i = 0; $i -lt $naturalOrders.Count; $i++) {
        $entry = $naturalOrders[$i]
        $row = $entry.source_row
        $dayKey = $entry.date.ToString('yyyy-MM-dd')
        $isOfflineProperty = $entry.PSObject.Properties['is_offline']
        $isOffline = $null -ne $isOfflineProperty -and [bool]$isOfflineProperty.Value
        $platform = $entry.platform
        if (-not $platform) { $platform = '美团' }
        $sequenceKey = "$dayKey|$platform"
        if (-not $sequenceByDay.ContainsKey($sequenceKey)) { $sequenceByDay[$sequenceKey] = 0 }
        $sequenceByDay[$sequenceKey]++
        $dailyOrderNo = $entry.daily_order_no
        $income = Convert-ToNullableNumber (Get-CellValue $row $(if ($isOffline) { 3 } else { 7 }))
        $cost = Convert-ToNullableNumber (Get-CellValue $row $(if ($isOffline) { 5 } else { 10 }))
        if ($isOffline -and $entry.order_no -and $pendingGiftCodesByOrder.ContainsKey($entry.order_no)) {
            $cost = $null
        } elseif ($isOffline -and $entry.order_no -and $giftCostByOrder.ContainsKey($entry.order_no) -and $null -ne $cost) {
            $cost = [math]::Round(([double]$cost + [double]$giftCostByOrder[$entry.order_no]),2)
        }
        $delivery = Convert-ToNullableNumber (Get-CellValue $row $(if ($isOffline) { 4 } else { 9 }))
        $old = if ($entry.order_no -and $oldIncomeByOrder.ContainsKey($entry.order_no)) { $oldIncomeByOrder[$entry.order_no] } else { $null }
        if (($null -eq $delivery -or [math]::Abs($delivery) -lt 0.000001) -and $null -ne $old -and $null -ne $old.delivery -and $old.delivery -gt 0) {
            # 旧版本曾允许人工配送费只存在于月度收入表。该差异与当前
            # 订单修正无关，不应阻断用户已经确认的写入。把旧月表的明确
            # 正数配送费原子迁移到订单汇总权威字段，再继续刷新；订单号
            # 不唯一或无法定位时仍停止，避免把金额写到错误订单。
            if (-not $entry.order_no -or -not $orderSourceRowByOrderNo.ContainsKey([string]$entry.order_no)) {
                throw "配送费自动归一失败：订单 $($entry.order_no) 无法在订单汇总中唯一定位"
            }
            $sourceRowNumber = [int]$orderSourceRowByOrderNo[[string]$entry.order_no]
            $oldSourceNote = [string](Get-CellValue $row 15)
            $deliveryValue = [math]::Round([double]$old.delivery,2)
            $marker = '人工确认配送费：{0}元' -f $deliveryValue.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture)
            $cleanSourceNote = [regex]::Replace($oldSourceNote,'(?:^|[；;])\s*人工确认配送费：?\s*\d+(?:\.\d+)?\s*元?','').Trim('；',';',' ')
            $newSourceNote = (@($cleanSourceNote,$marker) | Where-Object { $_ }) -join '；'
            $feeCells = @(); $feeCells += ,@((New-ValueCell $deliveryValue))
            $noteCells = @(); $noteCells += ,@((New-ValueCell $newSourceNote))
            Set-SourceCellMatrix '订单汇总' "J${sourceRowNumber}:J${sourceRowNumber}" $feeCells
            Set-SourceCellMatrix '订单汇总' "P${sourceRowNumber}:P${sourceRowNumber}" $noteCells
            $reconciledSourceRows = (Get-SheetRows $sourceToken '订单汇总' "A${sourceRowNumber}:Q${sourceRowNumber}").rows
            if ($reconciledSourceRows.Count -ne 1) {
                throw "配送费自动归一回读失败：订单 $($entry.order_no) 未返回唯一订单行"
            }
            $reconciledSourceRow = $reconciledSourceRows[0]
            $reconciledOrderNo = ([string](Get-CellValue $reconciledSourceRow 2)).Trim()
            $reconciledDelivery = Convert-ToNullableNumber (Get-CellValue $reconciledSourceRow 9)
            $reconciledNote = [string](Get-CellValue $reconciledSourceRow 15)
            if ($reconciledOrderNo -ne [string]$entry.order_no -or $null -eq $reconciledDelivery -or [math]::Abs($reconciledDelivery - $deliveryValue) -gt 0.000001 -or $reconciledNote -notlike "*$marker*") {
                throw "配送费自动归一回读失败：订单 $($entry.order_no) 的金额或来源标记未正确保存"
            }
            $row[9] = New-ValueCell $deliveryValue
            $row[15] = New-ValueCell $newSourceNote
            $delivery = $deliveryValue
            $autoReconciledDeliveryFees.Add([pscustomobject]@{
                order_no=[string]$entry.order_no
                source_row=$sourceRowNumber
                amount=$deliveryValue
                month=$monthSheet.month.ToString('yyyy-MM')
            })
        }
        if ($null -eq $delivery) { $delivery = 0.0 }
        $soldText = if ($isOffline) { Get-CellValue $row 7 } else { '' }
        if (-not $isOffline -and $entry.order_no -and $soldByOrder.ContainsKey($entry.order_no)) {
            $parts = @()
            foreach ($code in @($soldByOrder[$entry.order_no].Keys | Sort-Object)) {
                $quantity = $soldByOrder[$entry.order_no][$code]
                $quantityText = if ([math]::Abs($quantity - [math]::Round($quantity)) -lt 0.000001) { [string][int][math]::Round($quantity) } else { $quantity.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture) }
                $parts += "$code×$quantityText"
            }
            $soldText = $parts -join '、'
        }
        $notes = @()
        if ($isOffline -and (Get-CellValue $row 9)) {
            $notes += Get-CellValue $row 9
        } elseif ($null -ne $old -and $old.note) {
            # 成本警告由系统根据本次商品成本重新生成，不能当成人工备注永久保留。
            $notes += @([regex]::Split([string]$old.note, '[；;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notmatch '^(?:货物|赠品)?成本待核对(?:[:：]|$)' })
        }
        if ($entry.order_no -and $pendingGiftCodesByOrder.ContainsKey($entry.order_no)) {
            $pendingGiftCodes = @($pendingGiftCodesByOrder[$entry.order_no].Keys | Sort-Object)
            $notes += '赠品成本待核对：' + (($pendingGiftCodes | ForEach-Object { "${_}未知成本价" }) -join '、')
        }
        if ($null -eq $cost -and -not ($entry.order_no -and $pendingGiftCodesByOrder.ContainsKey($entry.order_no))) {
            $missingCodes = if ($entry.order_no -and $missingCostCodesByOrder.ContainsKey($entry.order_no)) {
                @($missingCostCodesByOrder[$entry.order_no].Keys | Sort-Object)
            } else {
                @()
            }
            # PowerShell unwraps a one-item pipeline result into a scalar. Always
            # normalize before Count so a single missing SKU remains a valid
            # "pending cost" case instead of aborting the whole ledger refresh.
            if (@($missingCodes).Count -gt 0) {
                $notes += '成本待核对：' + (($missingCodes | ForEach-Object { "${_}未知成本价" }) -join '、')
            } else {
                $notes += '成本待核对：订单汇总未生成货物成本'
            }
        }
        if ($entry.order_no -and $refundNotesByOrder.ContainsKey($entry.order_no)) {
            $notes += @($refundNotesByOrder[$entry.order_no])
        }
        $rowNumber = 2 + $i
        $isLastOfDay = ($i -eq $naturalOrders.Count - 1) -or ($naturalOrders[$i+1].date -ne $entry.date)
        $dailyFormula = if ($isLastOfDay) { "=IF(COUNTIFS(`$C`$2:`$C`$1001,C$rowNumber,`$E`$2:`$E`$1001,`"`")>0,`"`",SUMIF(`$C`$2:`$C`$1001,C$rowNumber,`$G`$2:`$G`$1001))" } else { $null }
        $incomeMatrix += ,@(
            (New-ValueCell $(if ($dailyOrderNo) { $dailyOrderNo } else { "$platform$($sequenceByDay[$sequenceKey])号单" })),
            (New-ValueCell $entry.order_no),
            (New-ValueCell (Convert-ToExcelSerial $entry.date)),
            (New-ValueCell $income),
            (New-ValueCell $cost),
            (New-ValueCell $delivery),
            (New-FormulaCell ('=IF(B{0}="","",IF(E{0}="","",IFERROR(D{0},0)-E{0}-IFERROR(F{0},0)))' -f $rowNumber)),
            (New-ValueCell $soldText),
            (New-ValueCell (Get-GiftText $entry.order_no)),
            (New-FormulaCell $dailyFormula),
            (New-ValueCell (($notes | Select-Object -Unique) -join '；'))
        )
    }

    $oldExpenseRows = (Get-SheetRows $ledgerToken $monthSheet.expense 'A2:F1001').rows
    $expenseEntries = @()
    $oldManagedExpenseByKey = @{}
    foreach ($row in $oldExpenseRows) {
        $date = Convert-ToDate (Get-CellValue $row 0)
        $project = Get-CellValue $row 1
        $fee = Convert-ToNullableNumber (Get-CellValue $row 2)
        $orderNo = Get-CellValue $row 3
        $source = Get-CellValue $row 4
        $note = Get-CellValue $row 5
        if (-not $date -and -not $project -and $null -eq $fee -and -not $orderNo -and -not $source -and -not $note) { continue }
        if ($orderNo -and $project -in @('补单平台扣款','补单配送费','营销赠品支出')) {
            $oldManagedExpenseByKey["$orderNo|$project"] = [pscustomobject]@{date=$date;project=$project;fee=$fee;order_no=$orderNo;source=$source;note=$note}
            continue
        }
        $expenseEntries += [pscustomobject]@{date=$date;project=$project;fee=$fee;order_no=$orderNo;source=$source;note=$note;sort_order=9}
    }
    $autoExpenseCount = 0
    foreach ($row in $interventionRows) {
        $date = Convert-ToDate (Get-CellValue $row 0)
        $orderNo = Get-CellValue $row 3
        if ($null -eq $date -or $date -lt $monthStart -or $date -gt $monthEnd -or $date -gt $dataThroughDate -or -not $orderNo) { continue }
        $platformFee = Convert-ToNullableNumber (Get-CellValue $row 6)
        $deliveryFee = Convert-ToNullableNumber (Get-CellValue $row 7)
        $deliveryType = Get-CellValue $row 5
        $platformAmount = if ($null -eq $platformFee) { $null } else { [double]$platformFee }
        $platformSource = if ($null -eq $platformFee) { '待识别' } else { '平台自动识别' }
        $platformNote = if ($null -eq $platformFee) { '平台扣款尚未识别；未知不等于0元' } else { '不含商品进货成本' }
        if ($null -eq $platformAmount) {
            $oldPlatformKey = "$orderNo|补单平台扣款"
            $oldPlatform = if ($oldManagedExpenseByKey.ContainsKey($oldPlatformKey)) { $oldManagedExpenseByKey[$oldPlatformKey] } else { $null }
            if ($null -ne $oldPlatform -and $null -ne $oldPlatform.fee -and [double]$oldPlatform.fee -gt 0) {
                $platformAmount = [double]$oldPlatform.fee
                $platformSource = '历史已确认'
                $platformNote = '原始补单记录暂缺平台扣款，保留支出表已有已确认金额'
            }
        }
        $expenseEntries += [pscustomobject]@{date=$date;project='补单平台扣款';fee=$platformAmount;order_no=$orderNo;source=$platformSource;note=$platformNote;sort_order=1}
        $autoExpenseCount++

        $deliveryTotal = if ($null -eq $deliveryFee) { 0.0 } else { [double]$deliveryFee }
        $deliverySource = '待补录'
        $breakdown = @()
        if ($deliveryTotal -gt 0) {
            $breakdown += ('配送费 ¥{0:N2}' -f $deliveryTotal)
            $deliverySource = '人工补录'
        } else {
            $oldKey = "$orderNo|补单配送费"
            $oldDelivery = if ($oldManagedExpenseByKey.ContainsKey($oldKey)) { $oldManagedExpenseByKey[$oldKey] } else { $null }
            if ($null -ne $oldDelivery -and $null -ne $oldDelivery.fee -and $oldDelivery.fee -gt 0) {
                $deliveryTotal = [double]$oldDelivery.fee
                $deliverySource = '人工补录'
                $breakdown += '保留支出表人工填写金额'
            } elseif ($deliveryType -eq '自配送' -or ($null -ne $oldDelivery -and $oldDelivery.source -eq '自配送确认')) {
                $deliveryTotal = 0.0
                $deliverySource = '自配送确认'
                $breakdown += '自配送，确认配送费为0'
            } else {
                $deliveryTotal = 0.0
                $deliverySource = '待补录'
                $breakdown += '暂未识别配送费；未知不等于确认0元'
            }
        }
        $expenseEntries += [pscustomobject]@{date=$date;project='补单配送费';fee=$deliveryTotal;order_no=$orderNo;source=$deliverySource;note=($breakdown -join '；');sort_order=2}
        $autoExpenseCount++

        if ($giftTextByOrder.ContainsKey($orderNo)) {
            $giftText = Get-GiftText $orderNo
            $giftPending = $pendingGiftCodesByOrder.ContainsKey($orderNo)
            $giftAmount = if ($giftPending) { $null } elseif ($giftCostByOrder.ContainsKey($orderNo)) { [math]::Round([double]$giftCostByOrder[$orderNo],2) } else { 0.0 }
            $giftSource = if ($giftPending) { '待补录' } else { '赠品记录' }
            $giftNote = if ($giftPending) {
                $pendingCodes = @($pendingGiftCodesByOrder[$orderNo].Keys | Sort-Object)
                '赠品成本待核对：' + (($pendingCodes | ForEach-Object { "${_}未知成本价" }) -join '、')
            } else {
                "赠品：$giftText；不计补单购买商品成本，仅计实际赠品成本"
            }
            $expenseEntries += [pscustomobject]@{date=$date;project='营销赠品支出';fee=$giftAmount;order_no=$orderNo;source=$giftSource;note=$giftNote;sort_order=3}
            $autoExpenseCount++
        }
    }
    $expenseEntries = @($expenseEntries | Sort-Object @{Expression={if ($null -eq $_.date) {[datetime]::MaxValue} else {$_.date}}}, order_no, sort_order, project)
    $promotionExpenseTotal = 0.0
    foreach ($expenseEntry in @($expenseEntries)) {
        if (([string]$expenseEntry.project).Trim() -ne '投流推广费' -or $null -eq $expenseEntry.fee) { continue }
        $promotionExpenseTotal += [double]$expenseEntry.fee
    }
    $promotionExpenseTotal = [math]::Round($promotionExpenseTotal,2)
    $monthKey = $monthSheet.month.ToString('yyyy-MM')
    if (-not $monthSettlementRowsByKey.ContainsKey($monthKey)) { throw "月度结算缺少月份：$monthKey" }
    $monthSettlementRow = [int]$monthSettlementRowsByKey[$monthKey]
    Set-SourceCellMatrix '月度结算' "N${monthSettlementRow}:N${monthSettlementRow}" @(@(New-ValueCell $promotionExpenseTotal))
    $promotionReadback = (Get-SheetRows $sourceToken '月度结算' "N${monthSettlementRow}:N${monthSettlementRow}").rows
    $actualPromotionExpense = if ($promotionReadback.Count -gt 0) { Convert-ToNullableNumber (Get-CellValue $promotionReadback[0] 0) } else { $null }
    if ($null -eq $actualPromotionExpense -or [math]::Abs($actualPromotionExpense - $promotionExpenseTotal) -gt 0.005) {
        throw "平台推广费用回读失败：$monthKey，预期=$promotionExpenseTotal，实际=$actualPromotionExpense"
    }
    if ($expenseEntries.Count -gt 1000) { throw "$($monthSheet.expense) 超过1000行容量" }
    $expenseMatrix = @()
    foreach ($entry in $expenseEntries) {
        $expenseMatrix += ,@(
            (New-ValueCell $(if ($null -ne $entry.date) { Convert-ToExcelSerial $entry.date } else { $null })),
            (New-ValueCell $entry.project),
            (New-ValueCell $entry.fee),
            (New-ValueCell $entry.order_no),
            (New-ValueCell $entry.source),
            (New-ValueCell $entry.note)
        )
    }

    Clear-SheetData $monthSheet.income 'K'
    Clear-SheetData $monthSheet.expense 'F'
    if ($incomeMatrix.Count -gt 0) { Set-CellMatrix $monthSheet.income "A2:K$($incomeMatrix.Count+1)" $incomeMatrix }
    if ($expenseMatrix.Count -gt 0) { Set-CellMatrix $monthSheet.expense "A2:F$($expenseMatrix.Count+1)" $expenseMatrix }
    if ($incomeMatrix.Count -gt 0) {
        Assert-FormulaRange $monthSheet.income "G2:G$($incomeMatrix.Count+1),J2:J$($incomeMatrix.Count+1)"
    }

    $checkIncome = @()
    $checkExpense = @()
    if ($incomeMatrix.Count -gt 0) { $checkIncome = (Get-SheetRows $ledgerToken $monthSheet.income "A2:K$($incomeMatrix.Count+1)").rows }
    if ($expenseMatrix.Count -gt 0) { $checkExpense = (Get-SheetRows $ledgerToken $monthSheet.expense "A2:F$($expenseMatrix.Count+1)").rows }
    if (@($checkIncome).Count -ne $incomeMatrix.Count) { throw "$($monthSheet.income) 回读行数不一致" }
    if (@($checkExpense).Count -ne $expenseMatrix.Count) { throw "$($monthSheet.expense) 回读行数不一致" }
    # One source of truth: expense rows, never a second intervention deduction.
    $expenseSummary = Update-PetstarMonthExpenses -Month $monthStart -SourceToken $sourceToken -LedgerToken $ledgerToken -Cli $larkCli -Transaction $script:financeTransaction -Entries $expenseEntries
    for ($checkIndex = 0; $checkIndex -lt $checkIncome.Count; $checkIndex++) {
        $row = $checkIncome[$checkIndex]
        $orderNo = Get-CellValue $row 1
        if ($orderNo -and $interventionOrderSet.ContainsKey($orderNo)) { throw "$($monthSheet.income) 仍包含干预单：$orderNo" }
        $incomeValue = Convert-ToNullableNumber (Get-CellValue $row 3)
        $costValue = Convert-ToNullableNumber (Get-CellValue $row 4)
        $deliveryValue = Convert-ToNullableNumber (Get-CellValue $row 5)
        $profitValue = Convert-ToNullableNumber (Get-CellValue $row 6)
        if ($null -eq $costValue) {
            if ($null -ne $profitValue) { throw "$($monthSheet.income) 待核对成本订单不应显示利润：$orderNo" }
        } else {
            if ($null -eq $profitValue) { throw "$($monthSheet.income) 单笔公式没有计算结果：$orderNo" }
            $expected = $(if ($null -eq $incomeValue) {0}else{$incomeValue}) - $costValue - $(if ($null -eq $deliveryValue) {0}else{$deliveryValue})
            if ([math]::Abs($expected - $profitValue) -gt 0.02) { throw "$($monthSheet.income) 单笔公式抽查失败：$orderNo" }
        }
        $currentDate = Get-CellValue $row 2
        $nextDate = if ($checkIndex + 1 -lt $checkIncome.Count) { Get-CellValue $checkIncome[$checkIndex+1] 2 } else { '' }
        if ($currentDate -and $currentDate -ne $nextDate) {
            $dayHasPendingCost = @($checkIncome | Where-Object {
                (Get-CellValue $_ 2) -eq $currentDate -and $null -eq (Convert-ToNullableNumber (Get-CellValue $_ 4))
            }).Count -gt 0
            $dailyProfit = Convert-ToNullableNumber (Get-CellValue $row 9)
            if ($dayHasPendingCost -and $null -ne $dailyProfit) { throw "$($monthSheet.income) 成本待核对时当天利润不应显示：$currentDate" }
            if (-not $dayHasPendingCost -and $null -eq $dailyProfit) { throw "$($monthSheet.income) 当天汇总公式没有计算结果：$currentDate" }
        }
    }

    if ($ThroughDate.Date -ge $monthStart.Date -and $ThroughDate.Date -le $monthEnd.Date) {
        $requestedRows = @($checkIncome | Where-Object {
            $rowDate = Convert-ToDate (Get-CellValue $_ 2)
            $null -ne $rowDate -and $rowDate.Date -eq $ThroughDate.Date
        })
        $pendingOrders = @()
        $incomeTotal = 0.0
        $costTotal = 0.0
        $deliveryTotal = 0.0
        $profitTotal = 0.0
        foreach ($requestedRow in $requestedRows) {
            $incomeValue = Convert-ToNullableNumber (Get-CellValue $requestedRow 3)
            $costValue = Convert-ToNullableNumber (Get-CellValue $requestedRow 4)
            $deliveryValue = Convert-ToNullableNumber (Get-CellValue $requestedRow 5)
            $profitValue = Convert-ToNullableNumber (Get-CellValue $requestedRow 6)
            if ($null -ne $incomeValue) { $incomeTotal += [double]$incomeValue }
            if ($null -ne $costValue) { $costTotal += [double]$costValue }
            if ($null -ne $deliveryValue) { $deliveryTotal += [double]$deliveryValue }
            if ($null -ne $profitValue) { $profitTotal += [double]$profitValue }
            if ($null -eq $costValue) {
                $pendingOrders += [pscustomobject]@{
                    daily_order_no=Get-CellValue $requestedRow 0
                    order_no=Get-CellValue $requestedRow 1
                    note=Get-CellValue $requestedRow 10
                }
            }
        }
        $requestedDateSummary = [pscustomobject]@{
            date=$ThroughDate.ToString('yyyy-MM-dd')
            order_count=$requestedRows.Count
            income=[math]::Round($incomeTotal,2)
            product_cost=$(if ($pendingOrders.Count -gt 0) { $null } else { [math]::Round($costTotal,2) })
            delivery_fee=[math]::Round($deliveryTotal,2)
            profit=$(if ($pendingOrders.Count -gt 0) { $null } else { [math]::Round($profitTotal,2) })
            pending_cost_count=$pendingOrders.Count
            pending_orders=$pendingOrders
        }
    }

    $summaries += [pscustomobject]@{
        month = $monthSheet.month.ToString('yyyy-MM')
        income_sheet = $monthSheet.income
        natural_orders = $incomeMatrix.Count
        expense_sheet = $monthSheet.expense
        expense_rows = $expenseMatrix.Count
        automatic_intervention_expense_rows = $autoExpenseCount
        preserved_manual_expense_rows = $expenseMatrix.Count - $autoExpenseCount
        promotion_expense = $promotionExpenseTotal
        operating_expenses = $expenseSummary
    }
}

$finalResult = [ordered]@{
    ok = $true
    workbook_url = $ledgerMeta.workbook_url
    start_date = $ledgerStartDate.ToString('yyyy-MM-dd')
    through_date = $ThroughDate.ToString('yyyy-MM-dd')
    data_through_date = $dataThroughDate.ToString('yyyy-MM-dd')
    scope = $(if ($selectedMonthKeys.Count -gt 0) { 'selected_months' } else { 'all_months' })
    month_filter = @($selectedMonthKeys)
    auto_reconciled_delivery_fee_count = $autoReconciledDeliveryFees.Count
    auto_reconciled_delivery_fee_amount = $(if ($autoReconciledDeliveryFees.Count -gt 0) { [math]::Round((($autoReconciledDeliveryFees | Measure-Object amount -Sum).Sum),2) } else { 0.0 })
    auto_reconciled_delivery_fees = $autoReconciledDeliveryFees.ToArray()
    months = $summaries
    requested_date_summary = $requestedDateSummary
    formula_verification = 'success'
    completed_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
}
Close-LedgerTransaction $true
$finalResult | ConvertTo-Json -Depth 8
