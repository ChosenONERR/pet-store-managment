param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent),
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$agentRoot = Join-Path $ProjectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$sourceMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\daily_ledger_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$tokens = @{source=$sourceToken;ledger=$ledgerToken}
. (Join-Path $agentRoot 'finance_transaction.ps1')

function Invoke-LarkJson([scriptblock]$Action,[string]$Message) {
    $global:LASTEXITCODE = 0
    $raw = @(& $Action 2>&1 | ForEach-Object { [string]$_ }) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "$Message：$raw" }
    try { $result = $raw | ConvertFrom-Json } catch { throw "$Message：返回内容无法解析" }
    if (-not $result.ok) { throw "$Message：$raw" }
    return $result
}
function Get-CellValue($Row,[int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return $cell.value
}
function Convert-ToNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) { return $number }
    return $null
}
function Get-Range([string]$Token,[string]$Sheet,[string]$Range) {
    $result = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    return $result.data.ranges[0]
}
function Get-LedgerMonth([string]$SheetName) {
    if ($SheetName -match '^(\d{4})年(\d{1,2})月收入$') { return [datetime]::new([int]$Matches[1],[int]$Matches[2],1) }
    if ($SheetName -match '^(\d{1,2})月收入$') { return [datetime]::new(2026,[int]$Matches[1],1) }
    return $null
}
function New-ColumnMatrix([object[]]$Values) {
    $matrix = [Collections.Generic.List[object]]::new()
    foreach ($value in $Values) { $matrix.Add(@(@{value=$value})) }
    return ,$matrix.ToArray()
}
function Set-Range([string]$Sheet,[string]$Range,$Cells) {
    $payload = ConvertTo-Json -InputObject $Cells -Depth 8 -Compress
    $null = Invoke-LarkJson { $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $sourceToken --sheet-name $Sheet --range $Range --cells - --format json } "写入 $Sheet $Range 失败"
}

$sourceRange = Get-Range $sourceToken '订单汇总' 'A2:Q2001'
$sourceRows = @($sourceRange.cells)
$sourceRowNumbers = @($sourceRange.row_indices)
$sourceByOrder = @{}
$sourceByRow = @{}
for ($index=0; $index -lt $sourceRows.Count; $index++) {
    $row = @($sourceRows[$index])
    $rowNumber = if ($sourceRowNumbers.Count -eq $sourceRows.Count) { [int]$sourceRowNumbers[$index] } else { 2 + $index }
    $sourceByRow[$rowNumber] = $row
    $orderNo = ([string](Get-CellValue $row 2)).Trim()
    if (-not $orderNo) { continue }
    if ($sourceByOrder.ContainsKey($orderNo)) { throw "订单汇总存在重复订单：$orderNo" }
    $sourceByOrder[$orderNo] = [pscustomobject]@{row_number=$rowNumber;cells=$row}
}

$info = Invoke-LarkJson { & $larkCli sheets +workbook-info --as user --spreadsheet-token $ledgerToken --format json } '读取每日经营简账结构失败'
$incomeSheets = @($info.data.sheets | ForEach-Object {
    $month = Get-LedgerMonth ([string]$_.sheet_name)
    if ($null -ne $month -and $month -ge [datetime]'2026-08-01' -and $month -le [datetime]::Today.Date) {
        [pscustomobject]@{name=[string]$_.sheet_name;month=$month}
    }
} | Sort-Object month,name)
if ($incomeSheets.Count -eq 0) { throw '每日经营简账中没有可核对的收入表' }

$ledgerFees = @{}
foreach ($sheet in $incomeSheets) {
    $range = Get-Range $ledgerToken $sheet.name 'A2:K1001'
    foreach ($rawRow in @($range.cells)) {
        $row = @($rawRow)
        $orderNo = ([string](Get-CellValue $row 1)).Trim()
        $fee = Convert-ToNumber (Get-CellValue $row 5)
        if (-not $orderNo -or $null -eq $fee -or $fee -le 0) { continue }
        if ($ledgerFees.ContainsKey($orderNo) -and [math]::Abs([double]$ledgerFees[$orderNo].fee-$fee) -gt 0.001) {
            throw "同一订单在月度收入表出现冲突配送费：$orderNo"
        }
        $ledgerFees[$orderNo] = [pscustomobject]@{fee=[double]$fee;sheet=$sheet.name;month=$sheet.month.ToString('yyyy-MM')}
    }
}

$targets = [Collections.Generic.List[object]]::new()
foreach ($orderNo in @($ledgerFees.Keys | Sort-Object)) {
    if (-not $sourceByOrder.ContainsKey($orderNo)) { throw "月度收入表存在配送费，但订单汇总找不到订单：$orderNo" }
    $source = $sourceByOrder[$orderNo]
    $row = $source.cells
    if (([string](Get-CellValue $row 3)).Trim() -eq '已取消' -or ([string](Get-CellValue $row 4)).Trim() -ne '自然') {
        throw "配送费回写目标不是有效自然单：$orderNo"
    }
    $ledgerFee = [double]$ledgerFees[$orderNo].fee
    $sourceFee = Convert-ToNumber (Get-CellValue $row 9)
    if ($null -eq $sourceFee) { $sourceFee = 0.0 }
    if ($sourceFee -gt 0 -and [math]::Abs($sourceFee-$ledgerFee) -gt 0.02) {
        throw "订单汇总与月度收入表配送费冲突：$orderNo｜汇总=$sourceFee｜月表=$ledgerFee"
    }
    $oldNote = [string](Get-CellValue $row 15)
    $marker = '人工确认配送费：{0}元' -f $ledgerFee.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture)
    $cleanNote = [regex]::Replace($oldNote,'(?:^|[；;])\s*人工确认配送费：?\s*\d+(?:\.\d+)?\s*元?','').Trim('；',';',' ')
    $newNote = (@($cleanNote,$marker) | Where-Object { $_ }) -join '；'
    if ([math]::Abs($sourceFee-$ledgerFee) -gt 0.001 -or $oldNote -ne $newNote) {
        $targets.Add([pscustomobject]@{
            order_no=$orderNo
            row_number=[int]$source.row_number
            fee=$ledgerFee
            note=$newNote
            month=$ledgerFees[$orderNo].month
            fee_changed=([math]::Abs($sourceFee-$ledgerFee) -gt 0.001)
        })
    }
}

$totalsByMonth = @($targets | Group-Object month | ForEach-Object {
    $feeChanges = @($_.Group | Where-Object fee_changed)
    [ordered]@{
        month=$_.Name
        orders=$_.Count
        amount=[math]::Round((($_.Group.fee | Measure-Object -Sum).Sum),2)
        source_fee_updates=$feeChanges.Count
        source_fee_update_amount=[math]::Round((($feeChanges.fee | Measure-Object -Sum).Sum),2)
    }
})
$feeChangeTargets = @($targets | Where-Object fee_changed)
$resultSummary = [ordered]@{
    updates=$targets.Count
    amount=[math]::Round((($targets.fee | Measure-Object -Sum).Sum),2)
    source_fee_updates=$feeChangeTargets.Count
    source_fee_update_amount=[math]::Round((($feeChangeTargets.fee | Measure-Object -Sum).Sum),2)
    marker_only_updates=($targets.Count-$feeChangeTargets.Count)
    by_month=$totalsByMonth
}
if ($DryRun -or $targets.Count -eq 0) {
    $result = [ordered]@{ok=$true;dry_run=[bool]$DryRun}
    foreach ($key in $resultSummary.Keys) { $result[$key] = $resultSummary[$key] }
    $result | ConvertTo-Json -Depth 6
    exit 0
}

$transaction = New-FinanceTransaction $stateRoot '把月度收入表人工配送费回写订单汇总' ('TX-DELIVERY-RECON-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')) -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli
$closed = $false
try {
    $minRow = [int](($targets.row_number | Measure-Object -Minimum).Minimum)
    $maxRow = [int](($targets.row_number | Measure-Object -Maximum).Maximum)
    $targetByRow = @{}; foreach ($target in $targets) { $targetByRow[[int]$target.row_number] = $target }
    $deliveryValues = [Collections.Generic.List[object]]::new()
    $noteValues = [Collections.Generic.List[object]]::new()
    for ($rowNumber=$minRow; $rowNumber -le $maxRow; $rowNumber++) {
        $row = if ($sourceByRow.ContainsKey($rowNumber)) { $sourceByRow[$rowNumber] } else { $null }
        if ($targetByRow.ContainsKey($rowNumber)) {
            $deliveryValues.Add([double]$targetByRow[$rowNumber].fee)
            $noteValues.Add([string]$targetByRow[$rowNumber].note)
        } else {
            $deliveryValues.Add($(if ($null -ne $row) { Get-CellValue $row 9 } else { '' }))
            $noteValues.Add($(if ($null -ne $row) { Get-CellValue $row 15 } else { '' }))
        }
    }
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '订单汇总' "J${minRow}:J${maxRow}" $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '订单汇总' "P${minRow}:P${maxRow}" $larkCli
    Set-Range '订单汇总' "J${minRow}:J${maxRow}" (New-ColumnMatrix $deliveryValues.ToArray())
    Set-Range '订单汇总' "P${minRow}:P${maxRow}" (New-ColumnMatrix $noteValues.ToArray())

    $verifyRange = Get-Range $sourceToken '订单汇总' "A${minRow}:Q${maxRow}"
    $verified = @{}
    foreach ($rawRow in @($verifyRange.cells)) {
        $row = @($rawRow)
        $orderNo = ([string](Get-CellValue $row 2)).Trim()
        if (-not $orderNo -or -not ($targets.order_no -contains $orderNo)) { continue }
        $actual = Convert-ToNumber (Get-CellValue $row 9)
        $note = [string](Get-CellValue $row 15)
        $expected = @($targets | Where-Object order_no -eq $orderNo)[0]
        if ($null -eq $actual -or [math]::Abs($actual-[double]$expected.fee) -gt 0.001) { throw "配送费回读不一致：$orderNo" }
        if ($note -notmatch '人工确认配送费') { throw "配送费来源标记回读缺失：$orderNo" }
        $verified[$orderNo] = $true
    }
    if ($verified.Count -ne $targets.Count) { throw "配送费回读订单数量不一致：预期 $($targets.Count)，实际 $($verified.Count)" }
    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    $result = [ordered]@{ok=$true;dry_run=$false}
    foreach ($key in $resultSummary.Keys) { $result[$key] = $resultSummary[$key] }
    $result['source_protected'] = $true
    $result | ConvertTo-Json -Depth 6
} catch {
    $failure = $_.Exception.Message
    if (-not $closed) {
        $rollback = Invoke-FinanceTransactionRollback $transaction $tokens $larkCli $failure
        Close-FinanceTransaction $transaction
        $closed = $true
        if (-not $rollback.ok) { throw "$failure；自动回滚失败：$($rollback.error)" }
    }
    throw $failure
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
