param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent),
    [double]$ConfirmedUnitCost = 0.17
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$agentRoot = Join-Path $ProjectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$powershellExe = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$sourceMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\daily_ledger_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$workbooks = @{ source = $sourceToken; ledger = $ledgerToken }

. (Join-Path $agentRoot 'finance_transaction.ps1')

function Invoke-LarkJson([scriptblock]$Action, [string]$Message) {
    $lastError = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $global:LASTEXITCODE = 0
        $lines = @(& $Action 2>&1 | ForEach-Object { [string]$_ })
        $raw = $lines -join "`n"
        try {
            if ($LASTEXITCODE -ne 0) { throw "CLI退出码 $LASTEXITCODE：$raw" }
            $result = $raw | ConvertFrom-Json
            if (-not $result.ok) { throw $raw }
            return $result
        } catch {
            $lastError = [string]$_.Exception.Message
            if ($attempt -lt 3) { Start-Sleep -Seconds $attempt }
        }
    }
    throw "$Message：$lastError"
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}

function Get-Rows([string]$Token, [string]$Sheet, [string]$Range, [string]$Include = 'value,formula') {
    $result = Invoke-LarkJson {
        & $larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include $Include --max-chars 5000000 --format json
    } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    $rangeResult = @($result.data.ranges)[0]
    if ($rangeResult.truncated) { throw "读取 $Sheet 被截断" }
    return @($rangeResult.cells)
}

function Convert-ToNumber($Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $number = 0.0
    if (-not [double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        throw "无法识别数值：$Value"
    }
    return $number
}

function Test-RowEmpty($Row) {
    foreach ($cell in @($Row)) {
        if ($null -ne $cell -and $null -ne $cell.value -and -not [string]::IsNullOrWhiteSpace([string]$cell.value)) { return $false }
        if ($null -ne $cell -and $cell.formula) { return $false }
    }
    return $true
}

function Invoke-Child([string]$Script, [string[]]$Arguments) {
    $lastError = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $global:LASTEXITCODE = 0
        $lines = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Script @Arguments 2>&1 | ForEach-Object { [string]$_ })
        if ($LASTEXITCODE -eq 0) { return $lines -join "`n" }
        $lastError = @($lines | Where-Object { $_.Trim() } | Select-Object -Last 12) -join ' | '
        if ($attempt -lt 3) { Start-Sleep -Seconds (2 * $attempt) }
    }
    throw "子脚本失败：$(Split-Path $Script -Leaf)：$lastError"
}

function Assert-Formula([string]$Token, [string]$Sheet, [string]$Range) {
    $result = Invoke-LarkJson {
        & $larkCli sheets +formula-verify --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --format json --exit-on-error
    } "公式校验失败：$Sheet $Range"
    if ($result.data.status -ne 'success') { throw "公式存在错误：$Sheet $Range" }
}

$transaction = New-FinanceTransaction $stateRoot '清理KW2099重复档案并回填2026-10-04赠品成本' $null -AcquireLock -WorkbookTokens $workbooks -LarkCli $larkCli
$closed = $false
try {
    $keptRow = @(Get-Rows $sourceToken 'SKU成本主表' 'A199:Z199')[0]
    $duplicateRows = @(Get-Rows $sourceToken 'SKU成本主表' 'A1409:Z1410')
    if ((Get-CellValue $keptRow 0).Trim().ToLowerInvariant() -ne 'kw2099' -or (Get-CellValue $keptRow 16) -ne '51185331395') {
        throw 'SKU成本主表第199行身份已变化；未写入'
    }
    $baselineCost = Convert-ToNumber (Get-CellValue $keptRow 5)
    $movingCost = Convert-ToNumber (Get-CellValue $keptRow 6)
    if ([math]::Abs($baselineCost - $ConfirmedUnitCost) -gt 0.0001 -or [math]::Abs($movingCost - $ConfirmedUnitCost) -gt 0.0001) {
        throw "第199行成本不是确认值$ConfirmedUnitCost；未写入"
    }
    if ($duplicateRows.Count -ne 2 -or
        (Get-CellValue $duplicateRows[0] 0).Trim().ToLowerInvariant() -ne 'kw2099' -or (Get-CellValue $duplicateRows[0] 16) -ne '52834971444' -or
        (Get-CellValue $duplicateRows[1] 0).Trim().ToLowerInvariant() -ne 'kw2099' -or (Get-CellValue $duplicateRows[1] 16) -ne '52834971445') {
        throw 'SKU成本主表第1409、1410行身份已变化；未写入'
    }

    $giftRows = @(Get-Rows $sourceToken '赠品记录' 'A11:L12')
    $gift = $giftRows[0]
    if ((Get-CellValue $gift 3) -ne '2502329220931197851' -or (Get-CellValue $gift 4).Trim().ToLowerInvariant() -ne 'kw2099' -or [math]::Abs((Convert-ToNumber (Get-CellValue $gift 5)) - 1) -gt 0.0001) {
        throw '赠品记录第11行身份已变化；未写入'
    }

    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' 'A1409:Z1410' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' 'K199:K199' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' 'O199:O199' $larkCli

    $clearPayload = ConvertTo-Json -InputObject @('SKU成本主表!A1409:Z1410', 'SKU成本主表!O199:O199') -Compress
    $null = Invoke-LarkJson {
        $clearPayload | & $larkCli sheets +cells-batch-clear --as user --spreadsheet-token $sourceToken --ranges - --scope content --yes --format json
    } '清理KW2099重复档案失败'

    # 单单元格二维数组会被 PowerShell 管道自动摊平，直接保留明确的 JSON 形状。
    $statusPayload = '[[{"value":"默认整件"}]]'
    $null = Invoke-LarkJson {
        $statusPayload | & $larkCli sheets +cells-set --as user --spreadsheet-token $sourceToken --sheet-name 'SKU成本主表' --range 'K199:K199' --cells - --format json
    } '恢复KW2099成本状态失败'

    $costOutput = Invoke-Child (Join-Path $agentRoot 'rebuild_cost_snapshots.ps1') @('-ProjectRoot', $ProjectRoot)
    $ledgerOutput = Invoke-Child (Join-Path $agentRoot 'sync_daily_ledger.ps1') @('-ThroughDate', '2026-10-05', '-ProjectRoot', $ProjectRoot)

    $masterAll = @(Get-Rows $sourceToken 'SKU成本主表' 'A2:Z3000')
    $kw2099Rows = [Collections.Generic.List[int]]::new()
    for ($index = 0; $index -lt $masterAll.Count; $index++) {
        if (([string](Get-CellValue $masterAll[$index] 0)).Trim().ToLowerInvariant() -eq 'kw2099') { $kw2099Rows.Add(2 + $index) }
    }
    if ($kw2099Rows.Count -ne 1 -or $kw2099Rows[0] -ne 199) { throw "KW2099唯一性校验失败：行 $($kw2099Rows -join ',')" }
    $keptVerified = $masterAll[197]
    if ((Get-CellValue $keptVerified 10) -ne '默认整件' -or (Get-CellValue $keptVerified 14)) { throw 'KW2099保留档案的冲突状态未清理干净' }
    if (-not (Test-RowEmpty (@(Get-Rows $sourceToken 'SKU成本主表' 'A1409:Z1409')[0]))) { throw 'SKU成本主表第1409行未完全清空' }
    if (-not (Test-RowEmpty (@(Get-Rows $sourceToken 'SKU成本主表' 'A1410:Z1410')[0]))) { throw 'SKU成本主表第1410行未完全清空' }

    $giftVerified = @(Get-Rows $sourceToken '赠品记录' 'A11:L11')[0]
    $giftUnitCost = Convert-ToNumber (Get-CellValue $giftVerified 6)
    $giftTotalCost = Convert-ToNumber (Get-CellValue $giftVerified 7)
    if ([math]::Abs($giftUnitCost - $ConfirmedUnitCost) -gt 0.0001 -or [math]::Abs($giftTotalCost - $ConfirmedUnitCost) -gt 0.0001 -or (Get-CellValue $giftVerified 8) -ne '已计成本') {
        throw "赠品成本回填失败：单价=$giftUnitCost，总成本=$giftTotalCost，状态=$(Get-CellValue $giftVerified 8)"
    }

    $orderRows = @(Get-Rows $sourceToken '订单汇总' 'A2:L2001')
    $orderMatch = $null
    foreach ($row in $orderRows) { if ((Get-CellValue $row 2) -eq '2502329220931197851') { $orderMatch = $row; break } }
    if ($null -eq $orderMatch) { throw '订单汇总中找不到昨日美团2号单' }
    $orderCost = Convert-ToNumber (Get-CellValue $orderMatch 10)
    $orderProfit = Convert-ToNumber (Get-CellValue $orderMatch 11)
    if ($null -eq $orderCost -or $null -eq $orderProfit) { throw '昨日美团2号单成本或毛利仍为空' }

    $ledgerRows = @(Get-Rows $ledgerToken '10月收入' 'A2:K1001')
    $ledgerMatch = $null
    foreach ($row in $ledgerRows) { if ((Get-CellValue $row 1) -eq '2502329220931197851') { $ledgerMatch = $row; break } }
    if ($null -eq $ledgerMatch) { throw '10月收入表中找不到昨日美团2号单' }
    $ledgerIncome = Convert-ToNumber (Get-CellValue $ledgerMatch 3)
    $ledgerCost = Convert-ToNumber (Get-CellValue $ledgerMatch 4)
    $ledgerDelivery = Convert-ToNumber (Get-CellValue $ledgerMatch 5)
    $ledgerProfit = Convert-ToNumber (Get-CellValue $ledgerMatch 6)
    $expectedLedgerProfit = [math]::Round($ledgerIncome - $ledgerCost - $ledgerDelivery, 10)
    if ([math]::Abs($ledgerCost - $orderCost) -gt 0.0001 -or [math]::Abs($ledgerProfit - $expectedLedgerProfit) -gt 0.0001) {
        throw "10月收入表与订单汇总不一致：成本$ledgerCost/$orderCost，单笔赚亏$ledgerProfit/$expectedLedgerProfit"
    }

    Assert-Formula $sourceToken '订单汇总' 'K2:L2001'
    Assert-Formula $ledgerToken '10月收入' 'G2:G1001'

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok = $true
        transaction_id = $transaction.id
        cleared_rows = @(1409, 1410)
        kept_row = 199
        kw = 'kw2099'
        unit_cost = $giftUnitCost
        gift_total_cost = $giftTotalCost
        order_no = '2502329220931197851'
        order_cost = $orderCost
        order_profit = $orderProfit
        ledger_cost = $ledgerCost
        ledger_delivery = $ledgerDelivery
        ledger_profit = $ledgerProfit
        unique_kw_rows = $kw2099Rows.ToArray()
        formula_verification = 'success'
        atomic = $true
        cost_rebuild = $costOutput
        ledger_sync = $ledgerOutput
    } | ConvertTo-Json -Depth 8
} catch {
    $reason = [string]$_.Exception.Message
    $rollback = Invoke-FinanceTransactionRollback $transaction $workbooks $larkCli $reason
    if ($rollback.ok) { throw "$reason；本次清理与成本重算已全部回滚" }
    throw "$reason；自动回滚尚未完成：$($rollback.error)"
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
