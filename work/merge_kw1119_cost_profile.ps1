param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent),
    [double]$ConfirmedUnitCost = 56.15
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$agentRoot = Join-Path $ProjectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$powershellExe = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$meta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\daily_ledger_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$workbooks = @{ source = $sourceToken; ledger = $ledgerToken }

. (Join-Path $agentRoot 'finance_transaction.ps1')
. (Join-Path $agentRoot 'sku_master_schema.ps1')
$schema = Get-SkuMasterSchema -SpreadsheetToken $sourceToken

function Invoke-LarkJson([scriptblock]$Action, [string]$Message) {
    $lines = @(& $Action 2>&1 | ForEach-Object { [string]$_ })
    $raw = $lines -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "$Message：$raw" }
    try { $result = $raw | ConvertFrom-Json } catch { throw "$Message：返回内容无法解析" }
    if (-not $result.ok) { throw "$Message：$raw" }
    return $result
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}

function Get-Rows([string]$Sheet, [string]$Range) {
    $result = Invoke-LarkJson {
        & $larkCli sheets +cells-get --as user --spreadsheet-token $sourceToken --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json
    } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    return @($result.data.ranges[0].cells)
}

function Convert-ToNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if (-not [double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
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
    $lines = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Script @Arguments 2>&1 | ForEach-Object { [string]$_ })
    if ($LASTEXITCODE -ne 0) {
        $detail = @($lines | Where-Object { $_.Trim() } | Select-Object -Last 8) -join ' | '
        throw "子脚本失败：$(Split-Path $Script -Leaf)：$detail"
    }
    return $lines -join "`n"
}

$transaction = New-FinanceTransaction $stateRoot '合并KW1119重复SKU档案并登记56.15元成本' $null -AcquireLock -WorkbookTokens $workbooks -LarkCli $larkCli
$closed = $false
try {
    $rows = Get-Rows 'SKU成本主表' 'A2:Z3000'
    $matches = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $rows.Count; $index++) {
        $kw = ([string](Get-CellValue $rows[$index] 0)).Trim().ToLowerInvariant()
        if ($kw -eq 'kw1119') { $matches.Add([pscustomobject]@{ row = 2 + $index; cells = $rows[$index] }) }
    }
    if ($matches.Count -ne 2) { throw "预期KW1119有2条档案，实际为$($matches.Count)条；未写入" }

    $canonical = @($matches | Where-Object { (Get-CellValue $_.cells $schema.ProductId.Index) -eq '51180542081' })
    $pending = @($matches | Where-Object {
        -not (Get-CellValue $_.cells $schema.ProductId.Index) -and
        (Get-CellValue $_.cells $schema.PlatformStatus.Index) -eq '待平台同步'
    })
    if ($canonical.Count -ne 1 -or $pending.Count -ne 1) { throw 'KW1119两条档案身份不符合预期；未写入' }
    if ([int]$canonical[0].row -ne 1100 -or [int]$pending[0].row -ne 1397) { throw 'KW1119档案行号已变化；未写入' }
    if ((Get-CellValue $canonical[0].cells 1) -ne 'csx811') { throw 'KW1119完整档案的CSX已变化；未写入' }
    if ((Get-CellValue $pending[0].cells 2) -or (Get-CellValue $pending[0].cells $schema.ProductId.Index)) { throw 'KW1119待同步档案已被补全，拒绝自动清理' }

    $replenishments = Get-Rows '补货记录' 'A2:T1001'
    $purchaseMatches = @($replenishments | Where-Object { ([string](Get-CellValue $_ 1)).Trim().ToLowerInvariant() -eq 'kw1119' })
    if ($purchaseMatches.Count -ne 1) { throw "KW1119补货记录应有1条，实际为$($purchaseMatches.Count)条；未写入" }
    $purchaseQuantity = Convert-ToNumber (Get-CellValue $purchaseMatches[0] 2)
    $purchaseAmount = Convert-ToNumber (Get-CellValue $purchaseMatches[0] 5)
    if ([math]::Abs($purchaseQuantity - 2) -gt 0.0001 -or [math]::Abs($purchaseAmount - 112.3) -gt 0.0001) {
        throw "KW1119补货记录不是2个/112.3元；未写入"
    }

    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' 'A1100:Z1100' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' 'A1397:Z1397' $larkCli
    $clearPayload = ConvertTo-Json -InputObject @('SKU成本主表!A1397:Z1397') -Compress
    $null = Invoke-LarkJson {
        $clearPayload | & $larkCli sheets +cells-batch-clear --as user --spreadsheet-token $sourceToken --ranges - --scope content --yes --format json
    } '清理KW1119空白重复档案失败'

    $eventId = 'maintenance-kw1119-cost-56.15-20260925'
    $eventOutput = Invoke-Child (Join-Path $agentRoot 'record_finance_event.ps1') @(
        '-Action', 'SetSkuCost',
        '-GoodsCode', 'sku_id:51180542081',
        '-Amount', $ConfirmedUnitCost.ToString('0.####', [Globalization.CultureInfo]::InvariantCulture),
        '-Date', '2026-09-25',
        '-Note', '合并误建的待平台同步重复档案；本次补货2个共112.3元',
        '-EventId', $eventId,
        '-ProjectRoot', $ProjectRoot
    )

    $verifyRows = Get-Rows 'SKU成本主表' 'A2:Z3000'
    $verifyMatches = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $verifyRows.Count; $index++) {
        $kw = ([string](Get-CellValue $verifyRows[$index] 0)).Trim().ToLowerInvariant()
        if ($kw -eq 'kw1119') { $verifyMatches.Add([pscustomobject]@{ row = 2 + $index; cells = $verifyRows[$index] }) }
    }
    if ($verifyMatches.Count -ne 1 -or [int]$verifyMatches[0].row -ne 1100) { throw 'KW1119重复档案合并回读失败' }
    $verified = $verifyMatches[0].cells
    $baselineCost = Convert-ToNumber (Get-CellValue $verified $schema.BaselineUnitCost.Index)
    $movingCost = Convert-ToNumber (Get-CellValue $verified $schema.MovingAverageCost.Index)
    $costQuantity = Convert-ToNumber (Get-CellValue $verified $schema.CostBookQuantity.Index)
    if ([math]::Abs($baselineCost - $ConfirmedUnitCost) -gt 0.0001 -or [math]::Abs($movingCost - $ConfirmedUnitCost) -gt 0.0001 -or [math]::Abs($costQuantity - 2) -gt 0.0001) {
        throw "KW1119成本回读不一致：基准=$baselineCost，移动平均=$movingCost，成本账数量=$costQuantity"
    }
    $clearedRow = @(Get-Rows 'SKU成本主表' 'A1397:Z1397')[0]
    if (-not (Test-RowEmpty $clearedRow)) { throw 'SKU成本主表第1397行未完全清空' }

    foreach ($check in @(
        @{ sheet = 'SKU成本主表'; range = 'E2:G3000' },
        @{ sheet = '订单商品'; range = 'L2:N4001' },
        @{ sheet = '订单汇总'; range = 'K2:L2001' },
        @{ sheet = '驾驶舱'; range = 'C2:C19' }
    )) {
        $formula = Invoke-LarkJson {
            & $larkCli sheets +formula-verify --as user --spreadsheet-token $sourceToken --sheet-name $check.sheet --range $check.range --format json --exit-on-error
        } "公式校验失败：$($check.sheet)"
        if ($formula.data.status -ne 'success') { throw "公式校验未通过：$($check.sheet)" }
    }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok = $true
        transaction_id = $transaction.id
        kept_row = 1100
        cleared_row = 1397
        product_id = '51180542081'
        goods_code = 'kw1119'
        baseline_unit_cost = $baselineCost
        moving_average_cost = $movingCost
        cost_book_quantity = $costQuantity
        replenishment_quantity = $purchaseQuantity
        replenishment_amount = $purchaseAmount
        atomic = $true
        formula_verification = 'success'
        event_result = $eventOutput
    } | ConvertTo-Json -Depth 8
} catch {
    $reason = [string]$_.Exception.Message
    $rollback = Invoke-FinanceTransactionRollback $transaction $workbooks $larkCli $reason
    if ($rollback.ok) { throw "$reason；本次合并与成本修改已全部回滚" }
    throw "$reason；自动回滚尚未完成：$($rollback.error)"
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
