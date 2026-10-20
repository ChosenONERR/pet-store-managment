param([string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent))

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
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) { return $number }
    return $null
}

$formulas = @(
    '=COUNTIFS(''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$D$2:$D$2001,"<>已取消",''订单汇总''!$E$2:$E$2001,"自然")+COUNTIFS(''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''线下订单''!$C$2:$C$1001,"<>")',
    '=SUMIFS(''订单汇总''!$F$2:$F$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$D$2:$D$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))',
    '=SUMIFS(''订单汇总''!$H$2:$H$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$D$2:$D$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))',
    '=SUMIFS(''订单汇总''!$K$2:$K$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$F$2:$F$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))',
    '=SUMIFS(''订单汇总''!$L$2:$L$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$D$2:$D$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))-SUMIFS(''线下订单''!$F$2:$F$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))-SUMIFS(''线下订单''!$E$2:$E$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))'
)
$cells = [Collections.Generic.List[object]]::new()
foreach ($formula in $formulas) { $cells.Add(@(@{formula=$formula})) }
$descriptions = @(
    '线上自然单与线下订单合计；不含干预单和取消单',
    '线上商品小计加线下成交收入',
    '线上商家预计收入加线下成交收入，非实际到账',
    '线上与线下订单货物成本合计',
    '线上与线下收入-商品成本-配送费'
)
$descriptionCells = [Collections.Generic.List[object]]::new()
foreach ($description in $descriptions) { $descriptionCells.Add(@(@{value=$description})) }
$transaction = New-FinanceTransaction $stateRoot '驾驶舱纳入线下订单' ('TX-DASHBOARD-OFFLINE-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')) -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli
$closed = $false
try {
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '驾驶舱' 'C2:C6' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '驾驶舱' 'D2:D6' $larkCli
    $payload = ConvertTo-Json -InputObject $cells.ToArray() -Depth 8 -Compress
    $null = Invoke-LarkJson { $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'C2:C6' --cells - --format json } '更新驾驶舱公式失败'
    $descriptionPayload = ConvertTo-Json -InputObject $descriptionCells.ToArray() -Depth 8 -Compress
    $null = Invoke-LarkJson { $descriptionPayload | & $larkCli sheets +cells-set --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'D2:D6' --cells - --format json } '更新驾驶舱说明失败'
    $verify = Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'C2:C6' --format json --exit-on-error } '驾驶舱公式校验失败'
    if ($verify.data.status -ne 'success') { throw '驾驶舱公式校验未通过' }
    $readback = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'B2:C6' --include value --max-chars 10000 --format json } '驾驶舱回读失败'
    $rows = @($readback.data.ranges[0].cells)
    if ($rows.Count -ne 5) { throw '驾驶舱回读行数不正确' }
    $grossProfit = Convert-ToNumber (Get-CellValue $rows[4] 1)

    $incomeSheet = '{0}月收入' -f [datetime]::Today.Month
    $ledgerRead = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $ledgerToken --sheet-name $incomeSheet --range 'A2:K1001' --include value --max-chars 5000000 --format json } '月度收入表回读失败'
    if ($ledgerRead.data.has_more) { throw '月度收入表回读不完整' }
    $ledgerProfit = 0.0
    foreach ($row in @($ledgerRead.data.ranges[0].cells)) {
        if (-not (Get-CellValue $row 1)) { continue }
        $profit = Convert-ToNumber (Get-CellValue $row 6)
        if ($null -ne $profit) { $ledgerProfit += $profit }
    }
    $ledgerProfit = [math]::Round($ledgerProfit,2)
    if ($null -eq $grossProfit -or [math]::Abs($grossProfit-$ledgerProfit) -gt 0.02) {
        throw "驾驶舱毛利与月度收入利润不一致：驾驶舱=$grossProfit，月表=$ledgerProfit"
    }
    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{ok=$true;month=[datetime]::Today.ToString('yyyy-MM');dashboard_gross_profit=$grossProfit;ledger_profit=$ledgerProfit;matched=$true} | ConvertTo-Json
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
