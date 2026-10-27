param([string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent))

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$agentRoot = Join-Path $ProjectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$meta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
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
function Convert-ToDate($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $serial = Convert-ToNumber $Value
    if ($null -ne $serial -and $serial -gt 30000 -and $serial -lt 100000) { return [datetime]::FromOADate($serial) }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value,[ref]$date)) { return $date }
    return $null
}
function Convert-ToExcelSerial([datetime]$Value) { return [math]::Floor($Value.Date.ToOADate()) }
function Set-Cells([string]$Sheet,[string]$Range,$Cells,[string]$Message) {
    $payload = ConvertTo-Json -InputObject $Cells -Depth 12 -Compress
    $null = Invoke-LarkJson { $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $Sheet --range $Range --cells - --format json } $Message
}
function New-FormulaCell([string]$Formula,[string]$Format) {
    $cell = @{formula=$Formula}
    if ($Format) { $cell.cell_styles = @{number_format=$Format} }
    return $cell
}

$transaction = New-FinanceTransaction $stateRoot '修正驾驶舱口径并补齐2026年10月补货付款' ('TX-DASHBOARD-PAYMENT-20261025-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')) -AcquireLock -WorkbookTokens @{source=$token} -LarkCli $larkCli
$closed = $false
try {
    $replenishmentRead = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'A2:T1001' --include value --max-chars 2000000 --format json } '读取补货记录失败'
    if ($replenishmentRead.data.has_more) { throw '补货记录读取不完整' }
    $replenishmentRows = @($replenishmentRead.data.ranges[0].cells)
    $paymentRows = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $replenishmentRows.Count; $index++) {
        $row = $replenishmentRows[$index]
        $date = Convert-ToDate (Get-CellValue $row 0)
        $purchaseAmount = Convert-ToNumber (Get-CellValue $row 5)
        $paymentDate = Get-CellValue $row 11
        $paymentAmount = Get-CellValue $row 12
        $note = [string](Get-CellValue $row 13)
        if ($null -eq $date -or $date.Year -ne 2026 -or $date.Month -ne 10 -or $null -eq $purchaseAmount) { continue }
        if ($paymentDate -or $null -ne (Convert-ToNumber $paymentAmount)) { continue }
        $paymentSignalText = [regex]::Replace($note,'(?i)(未说明|没有说明|未提及|没有提及)(?:未付款|未支付|赊账|挂账|欠款|稍后付款|以后付款)','')
        if ($paymentSignalText -match '(?i)(未付款|未支付|赊账|挂账|欠款|稍后付款|以后付款)') { continue }
        $paymentRows.Add([pscustomobject]@{row=2+$index;date=$date;amount=[double]$purchaseAmount;goods_code=[string](Get-CellValue $row 1)})
    }

    foreach ($paymentRow in $paymentRows) {
        Register-FinanceRangeBeforeImage $transaction 'source' $token '补货记录' "L$($paymentRow.row):M$($paymentRow.row)" $larkCli
        $cells = @(@(
            @{value=(Convert-ToExcelSerial $paymentRow.date);cell_styles=@{number_format='yyyy-mm-dd'}},
            @{value=$paymentRow.amount;cell_styles=@{number_format='#,##0.00'}}
        ))
        Set-Cells '补货记录' "L$($paymentRow.row):M$($paymentRow.row)" $cells "补齐补货付款失败：第$($paymentRow.row)行"
    }

    Register-FinanceRangeBeforeImage $transaction 'source' $token '库存快照' 'H2:I2000' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $token '库存快照' 'K2:K2000' $larkCli
    $unitCostCells = [Collections.Generic.List[object]]::new()
    $inventoryValueCells = [Collections.Generic.List[object]]::new()
    $reminderCells = [Collections.Generic.List[object]]::new()
    for ($rowNumber = 2; $rowNumber -le 2000; $rowNumber++) {
        $unitCostFormula = '=IF(B{0}="","",IFERROR(IF(ISBLANK(INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(N{0},''SKU成本主表''!$Q$2:$Q$3000,0))),"",INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(N{0},''SKU成本主表''!$Q$2:$Q$3000,0))),IFERROR(IF(ISBLANK(INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(B{0},''SKU成本主表''!$A$2:$A$3000,0))),"",INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(B{0},''SKU成本主表''!$A$2:$A$3000,0))),IFERROR(IF(ISBLANK(INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(G{0},''SKU成本主表''!$B$2:$B$3000,0))),"",INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(G{0},''SKU成本主表''!$B$2:$B$3000,0))),""))))' -f $rowNumber
        $inventoryValueFormula = '=IF(OR(B{0}="",D{0}="",H{0}=""),"",(D{0}+IFERROR(M{0}*1,0))*H{0})' -f $rowNumber
        $reminderFormula = '=IF(C{0}="","",IF(B{0}="","缺少货号",IF(H{0}="","成本待核对",IF(ISNUMBER(SEARCH("重复货号",L{0})),"重复货号待核对",IF(F{0}<>"售卖中","不监控",IF(D{0}<=0,"缺货",IF(D{0}<=J{0},"低库存","正常")))))))' -f $rowNumber
        $unitCostCells.Add(@((New-FormulaCell $unitCostFormula '#,##0.####')))
        $inventoryValueCells.Add(@((New-FormulaCell $inventoryValueFormula '#,##0.00')))
        $reminderCells.Add(@((New-FormulaCell $reminderFormula $null)))
    }
    Set-Cells '库存快照' 'H2:H2000' $unitCostCells.ToArray() '更新库存单位成本公式失败'
    Set-Cells '库存快照' 'I2:I2000' $inventoryValueCells.ToArray() '更新库存货值公式失败'
    Set-Cells '库存快照' 'K2:K2000' $reminderCells.ToArray() '更新库存提醒公式失败'

    $dashboardDefinitions = @(
        @('经营结果','本月自然订单数','=COUNTIFS(''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$D$2:$D$2001,"<>已取消",''订单汇总''!$E$2:$E$2001,"自然")+COUNTIFS(''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''线下订单''!$C$2:$C$1001,"<>")','线上自然单与线下订单合计；不含干预单和取消单','0'),
        @('经营结果','本月自然商品销售额','=SUMIFS(''订单汇总''!$F$2:$F$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$D$2:$D$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))','线上订单实际商品小计加线下成交收入；不是商品目录标价','#,##0.00'),
        @('经营结果','本月预计结算收入','=SUMIFS(''订单汇总''!$H$2:$H$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$D$2:$D$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))','线上商家预计收入加线下成交收入；取消单排除，已确认退款按退款后金额','#,##0.00'),
        @('经营结果','本月商品成本','=SUMIFS(''订单汇总''!$K$2:$K$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$F$2:$F$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))','线上与线下订单货物成本合计','#,##0.00'),
        @('经营结果','本月经营毛利','=SUMIFS(''订单汇总''!$L$2:$L$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$D$2:$D$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))-SUMIFS(''线下订单''!$F$2:$F$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))-SUMIFS(''线下订单''!$E$2:$E$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))','线上与线下收入-商品成本-配送费','#,##0.00'),
        @('经营结果','本月经营利润','=C6-SUMIFS(''月度结算''!$G$2:$G$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))','经营毛利-其他经营费用；不再扣老板人工、摊销折旧和市场租金机会成本','#,##0.00'),
        @('现金与库存','本月补货付款','=SUMIFS(''补货记录''!$M$2:$M$1001,''补货记录''!$L$2:$L$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''补货记录''!$L$2:$L$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))','按付款日期汇总补货记录付款金额','#,##0.00'),
        @('现金与库存','本月现金净流入','=SUMIFS(''月度结算''!$B$2:$B$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$C$2:$C$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$D$2:$D$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$E$2:$E$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))-C8-SUMIFS(''月度结算''!$F$2:$F$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))-SUMIFS(''月度结算''!$H$2:$H$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))','平台与线下到账-补货付款-其他现金支出-老板取款','#,##0.00'),
        @('现金与库存','本月期末现金','=C9+SUMIFS(''月度结算''!$K$2:$K$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))','期初现金+本月现金净流入','#,##0.00'),
        @('现金与库存','当前库存货值','=SUMIFS(''库存快照''!$I$2:$I$2000,''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000))','（美团库存+待恢复库存数量）×SKU成本主表当前移动平均成本','#,##0.00'),
        @('库存提醒','当前缺货SKU数','=COUNTIFS(''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000),''库存快照''!$K$2:$K$2000,"缺货")','仅监控售卖中商品','0'),
        @('库存提醒','当前低库存SKU数','=COUNTIFS(''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000),''库存快照''!$K$2:$K$2000,"低库存")','售卖中且库存不高于补货阈值','0'),
        @('数据质量','待核对数据数','=COUNTIFS(''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000),''库存快照''!$K$2:$K$2000,"*核对")+COUNTIFS(''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000),''库存快照''!$K$2:$K$2000,"缺少*")+COUNTIFS(''订单商品''!$A$2:$A$4001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单商品''!$A$2:$A$4001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单商品''!$C$2:$C$4001,"<>",''订单商品''!$L$2:$L$4001,"")','缺货号、缺成本、重复货号及订单成本异常','0'),
        @('系统状态','最近库存快照日期','=IF(COUNTA(''库存快照''!$A$2:$A$2000)=0,"",MAX(''库存快照''!$A$2:$A$2000))','RPA最近一次读取日期','yyyy-mm-dd hh:mm'),
        @('系统状态','最近订单同步日期','=IF(COUNTA(''订单汇总''!$O$2:$O$2001)=0,"",MAX(''订单汇总''!$O$2:$O$2001))','RPA最近一次读取日期','yyyy-mm-dd hh:mm')
    )
    $dashboardCells = [Collections.Generic.List[object]]::new()
    foreach ($definition in $dashboardDefinitions) {
        $dashboardCells.Add(@(
            @{value=$definition[0]}, @{value=$definition[1]}, (New-FormulaCell $definition[2] $definition[4]),
            @{value=$definition[3]}, @{value=''}, @{value=''}
        ))
    }
    while ($dashboardCells.Count -lt 18) { $dashboardCells.Add(@(@{value=''},@{value=''},(New-FormulaCell '=""' $null),@{value=''},@{value=''},@{value=''})) }
    Register-FinanceRangeBeforeImage $transaction 'source' $token '驾驶舱' 'A2:F19' $larkCli
    Set-Cells '驾驶舱' 'A2:F19' $dashboardCells.ToArray() '更新驾驶舱失败'

    foreach ($range in @('H2:H2000','I2:I2000','K2:K2000')) {
        $verify = Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $token --sheet-name '库存快照' --range $range --format json --exit-on-error } "库存快照公式校验失败：$range"
        if ($verify.data.status -ne 'success') { throw "库存快照公式未通过：$range" }
    }
    $dashboardVerify = Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $token --sheet-name '驾驶舱' --range 'C2:C19' --format json --exit-on-error } '驾驶舱公式校验失败'
    if ($dashboardVerify.data.status -ne 'success') { throw '驾驶舱公式未通过' }

    $readback = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '补货记录' --range 'A2:M1001' --include value --max-chars 1000000 --format json } '补货付款回读失败'
    $octoberPurchase = 0.0
    $octoberPayment = 0.0
    $missingPaymentRows = [Collections.Generic.List[int]]::new()
    $rows = @($readback.data.ranges[0].cells)
    for ($index = 0; $index -lt $rows.Count; $index++) {
        $row = $rows[$index]
        $date = Convert-ToDate (Get-CellValue $row 0)
        if ($null -eq $date -or $date.Year -ne 2026 -or $date.Month -ne 10) { continue }
        $purchase = Convert-ToNumber (Get-CellValue $row 5)
        $payment = Convert-ToNumber (Get-CellValue $row 12)
        if ($null -ne $purchase) { $octoberPurchase += $purchase }
        if ($null -ne $payment) { $octoberPayment += $payment }
        if ($null -ne $purchase -and $null -eq $payment) { $missingPaymentRows.Add(2+$index) }
    }
    $dashboardReadback = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '驾驶舱' --range 'A2:D16' --include value --max-chars 50000 --format json } '驾驶舱回读失败'
    $dashboardRows = @($dashboardReadback.data.ranges[0].cells)
    $paymentMetric = Convert-ToNumber (Get-CellValue $dashboardRows[6] 2)
    if ([math]::Abs($octoberPurchase-$octoberPayment) -gt 0.005 -or $missingPaymentRows.Count -gt 0) { throw "10月补货付款仍不完整：采购=$octoberPurchase，付款=$octoberPayment，缺付款行=$($missingPaymentRows -join ',')" }
    if ($null -eq $paymentMetric -or [math]::Abs($paymentMetric-$octoberPayment) -gt 0.005) { throw "驾驶舱补货付款与补货记录不一致：驾驶舱=$paymentMetric，补货记录=$octoberPayment" }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok = $true
        backfilled_payment_rows = @($paymentRows | ForEach-Object { $_.row })
        october_purchase_amount = [math]::Round($octoberPurchase,2)
        october_payment_amount = [math]::Round($octoberPayment,2)
        dashboard_payment_amount = [math]::Round($paymentMetric,2)
        inventory_valuation = '（美团库存+待恢复库存数量）×SKU成本主表当前移动平均成本'
        removed_dashboard_metrics = @('本月经济利润','本月实际到账','待恢复库存SKU数')
    } | ConvertTo-Json -Depth 6
} catch {
    $failure = $_.Exception.Message
    if (-not $closed) {
        $rollback = Invoke-FinanceTransactionRollback $transaction @{source=$token} $larkCli $failure
        Close-FinanceTransaction $transaction
        $closed = $true
        if (-not $rollback.ok) { throw "$failure；自动回滚失败：$($rollback.error)" }
    }
    throw $failure
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
