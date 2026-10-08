param(
    [int]$GraceHours = 24,
    [switch]$DryRun,
    [string]$WorkRoot,
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少飞书经营驾驶舱元数据' }
$meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $ProjectRoot 'agent\finance_transaction.ps1')
$stateRoot = Join-Path $ProjectRoot 'agent\state'
$script:financeTransaction = if ($DryRun) { $null } else { Join-FinanceTransaction $stateRoot }
$script:financeTransactionOwner = $false
$script:financeTransactionClosed = $false
if (-not $DryRun -and $null -eq $script:financeTransaction) {
    $script:financeTransaction = New-FinanceTransaction $stateRoot '匹配待同步赠品' $null -AcquireLock -WorkbookTokens @{source=$spreadsheetToken} -LarkCli 'lark-cli'
    $script:financeTransactionOwner = $true
}
function Close-GiftTransaction([bool]$Succeeded, [string]$Reason = '') {
    if ($script:financeTransactionClosed -or $null -eq $script:financeTransaction) { return }
    if ($script:financeTransactionOwner) {
        if ($Succeeded) { Complete-FinanceTransaction $script:financeTransaction }
        else { $null = Invoke-FinanceTransactionRollback $script:financeTransaction @{source=$spreadsheetToken} 'lark-cli' $Reason }
    }
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
}
trap { $reason=[string]$_.Exception.Message; Close-GiftTransaction $false $reason; [Console]::Error.WriteLine($reason); exit 1 }

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    $valueProperty = $cell.PSObject.Properties['value']
    if ($null -eq $valueProperty -or $null -eq $valueProperty.Value) { return $null }
    return [string]$valueProperty.Value
}
function Convert-ToDateValue($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $serial = 0.0
    if ([double]::TryParse([string]$Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$serial)) {
        return ([datetime]'1899-12-30').AddDays($serial)
    }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value,[ref]$parsed)) { return $parsed }
    return $null
}
function Convert-ToExcelSerial([datetime]$Value) { return ($Value - [datetime]'1899-12-30').TotalDays }
function New-ValueCell($Value) { if ($null -eq $Value -or [string]$Value -eq '') { return @{} }; return @{value=$Value} }
function Get-SheetRows([string]$SheetName, [string]$Range) {
    $raw = lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --include value --max-chars 5000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取飞书失败：$SheetName" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "飞书读取结果不完整：$SheetName" }
    return @($result.data.ranges[0].cells)
}
function Set-AtomicOperations([object[]]$Operations) {
    if ($Operations.Count -eq 0) { return }
    foreach ($operation in $Operations) {
        if ([string]$operation.shortcut -ne '+cells-set') { throw "事务暂不允许未登记快照的批量动作：$($operation.shortcut)" }
        Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $spreadsheetToken ([string]$operation.input.sheet_name) ([string]$operation.input.range) 'lark-cli'
    }
    $payload = ConvertTo-Json -InputObject $Operations -Depth 12 -Compress
    $raw = $payload | lark-cli sheets +batch-update --as user --spreadsheet-token $spreadsheetToken --operations - --yes --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '待匹配赠品批量更新失败' }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw '待匹配赠品批量更新失败' }
}
function Get-CleanNote([string]$Note) {
    if (-not $Note) { return '' }
    return (($Note -replace '^等待订单同步[；;]?\s*','').Trim())
}

$ordersByNo = @{}
foreach ($row in (Get-SheetRows '订单汇总' 'A2:Q2001')) {
    $orderNo = Get-CellValue $row 2
    $date = Convert-ToDateValue (Get-CellValue $row 0)
    if (-not $orderNo -or $null -eq $date) { continue }
    if (-not $ordersByNo.ContainsKey($orderNo)) { $ordersByNo[$orderNo] = @() }
    $ordersByNo[$orderNo] += [pscustomobject]@{
        order_no=$orderNo;date=$date.Date;platform=$(if(Get-CellValue $row 1){Get-CellValue $row 1}else{'美团'});
        daily_order_no=Get-CellValue $row 16;status=Get-CellValue $row 3
    }
}
foreach ($row in (Get-SheetRows '线下订单' 'A2:J1001')) {
    $orderNo = Get-CellValue $row 2
    $date = Convert-ToDateValue (Get-CellValue $row 0)
    if (-not $orderNo -or $null -eq $date) { continue }
    if (-not $ordersByNo.ContainsKey($orderNo)) { $ordersByNo[$orderNo] = @() }
    $ordersByNo[$orderNo] += [pscustomobject]@{order_no=$orderNo;date=$date.Date;platform='线下';daily_order_no=Get-CellValue $row 1;status='已完成'}
}

$giftRows = Get-SheetRows '赠品记录' 'A2:L2001'
$pendingByOrder = @{}
for ($index = 0; $index -lt $giftRows.Count; $index++) {
    $row = $giftRows[$index]
    $dailyNo = Get-CellValue $row 2
    $orderNo = Get-CellValue $row 3
    if (-not $orderNo -or $dailyNo -ne '待同步') { continue }
    if (-not $pendingByOrder.ContainsKey($orderNo)) { $pendingByOrder[$orderNo] = @() }
    $pendingByOrder[$orderNo] += [pscustomobject]@{
        row_number=2+$index;cells=$row;created_at=Convert-ToDateValue (Get-CellValue $row 10)
    }
}

$operations = [Collections.Generic.List[object]]::new()
$matchedOrders = [Collections.Generic.List[string]]::new()
$cancelledOrders = [Collections.Generic.List[string]]::new()
$overdueOrders = [Collections.Generic.List[string]]::new()
$ambiguousOrders = [Collections.Generic.List[string]]::new()
$now = [datetime]::Now
foreach ($orderNo in @($pendingByOrder.Keys | Sort-Object)) {
    $pendingRows = @($pendingByOrder[$orderNo])
    $matches = if ($ordersByNo.ContainsKey($orderNo)) { @($ordersByNo[$orderNo]) } else { @() }
    if ($matches.Count -gt 1) {
        $ambiguousOrders.Add($orderNo)
        continue
    }
    if ($matches.Count -eq 0) {
        $created = $pendingRows | ForEach-Object { $_.created_at } | Where-Object { $null -ne $_ } | Sort-Object | Select-Object -First 1
        if ($null -ne $created -and ($now - [datetime]$created).TotalHours -ge $GraceHours) { $overdueOrders.Add($orderNo) }
        continue
    }

    $order = $matches[0]
    if ($order.status -eq '已取消') {
        foreach ($pending in $pendingRows) {
            $note = Get-CleanNote (Get-CellValue $pending.cells 9)
            $newNote = if ($note) { "订单已取消，赠品未计成本；$note" } else { '订单已取消，赠品未计成本' }
            $cMatrix = @(); $cMatrix += ,@((New-ValueCell '订单已取消待处理'))
            $jMatrix = @(); $jMatrix += ,@((New-ValueCell $newNote))
            $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='赠品记录';range="C$($pending.row_number):C$($pending.row_number)";cells=$cMatrix}})
            $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='赠品记录';range="J$($pending.row_number):J$($pending.row_number)";cells=$jMatrix}})
        }
        $cancelledOrders.Add($orderNo)
        continue
    }

    foreach ($pending in $pendingRows) {
        $note = Get-CleanNote (Get-CellValue $pending.cells 9)
        $abcMatrix = @(); $abcMatrix += ,@((New-ValueCell (Convert-ToExcelSerial $order.date)),(New-ValueCell $order.platform),(New-ValueCell $order.daily_order_no))
        $jMatrix = @(); $jMatrix += ,@((New-ValueCell $note))
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='赠品记录';range="A$($pending.row_number):C$($pending.row_number)";cells=$abcMatrix}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='赠品记录';range="J$($pending.row_number):J$($pending.row_number)";cells=$jMatrix}})
    }
    $matchedOrders.Add($orderNo)
}

if (-not $DryRun) { Set-AtomicOperations $operations.ToArray() }
if (-not $DryRun -and $operations.Count -gt 0) {
    $checkRows = Get-SheetRows '赠品记录' 'A2:L2001'
    foreach ($orderNo in $matchedOrders) {
        $remaining = @($checkRows | Where-Object { (Get-CellValue $_ 3) -eq $orderNo -and (Get-CellValue $_ 2) -eq '待同步' })
        if ($remaining.Count -gt 0) { throw "赠品订单自动匹配回读失败：$orderNo" }
    }
}

$finalResult = [ordered]@{
    ok=$true;dry_run=[bool]$DryRun;pending_order_count=$pendingByOrder.Count;matched_order_count=$matchedOrders.Count;
    matched_orders=@($matchedOrders);cancelled_orders=@($cancelledOrders);overdue_orders=@($overdueOrders);
    ambiguous_orders=@($ambiguousOrders);updated_cells=$operations.Count;grace_hours=$GraceHours;
    formula_verification='not_applicable';processed_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
}
Close-GiftTransaction $true
$finalResult | ConvertTo-Json -Depth 8
