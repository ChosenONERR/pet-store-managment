param(
    [datetime]$StartDate = [datetime]'2026-08-24',
    [datetime]$EndDate = [datetime]::Today,
    [string]$WorkRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $WorkRoot) { $WorkRoot = $PSScriptRoot }
if ($EndDate.Date -lt $StartDate.Date) { throw '结束日期不能早于开始日期' }

$metaPath = Join-Path $WorkRoot 'simple_finance_meta.json'
$interventionPath = Join-Path $WorkRoot 'intervention_orders.json'
$orderSyncPath = Join-Path $WorkRoot 'sync_meituan_orders.ps1'
$inventorySyncPath = Join-Path $WorkRoot 'sync_meituan_inventory.ps1'
foreach ($path in @($metaPath,$interventionPath,$orderSyncPath,$inventorySyncPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "缺少重建依赖：$path" }
}
$meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$interventionIds = @((Get-Content -LiteralPath $interventionPath -Raw | ConvertFrom-Json).intervention_order_ids | ForEach-Object { [string]$_ } | Select-Object -Unique)
if ($interventionIds.Count -eq 0) { throw '干预订单清单为空' }

function Invoke-Lark([string[]]$Arguments, [string]$FailureMessage) {
    $raw = & lark-cli @Arguments | Out-String
    if ($LASTEXITCODE -ne 0) { throw "$FailureMessage：$raw" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "$FailureMessage：$raw" }
    return $result
}
function New-ValueCell($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return @{value=''} }
    return @{value=$Value}
}
function Get-CellValue($Row,[int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}
function Set-CellMatrix([string]$SheetName,[string]$Range,$Rows) {
    $payload = ConvertTo-Json -InputObject $Rows -Depth 12 -Compress
    $null = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入飞书失败：$SheetName $Range" }
}
function Set-FormulaColumn([string]$Sheet,[string]$Column,[int]$LastRow,[string]$Template) {
    $cells = @()
    for ($row=2; $row -le $LastRow; $row++) {
        $cells += ,@(@{formula=$Template.Replace('{r}',[string]$row)})
    }
    Set-CellMatrix $Sheet "$Column`2:$Column$LastRow" $cells
}
function Assert-FormulaRange([string]$Sheet,[string]$Range) {
    $raw = lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name $Sheet --range $Range --format json --exit-on-error | Out-String
    if ($LASTEXITCODE -ne 0) { throw "公式校验失败：$Sheet $Range" }
    $result = $raw | ConvertFrom-Json
    if ($result.data.status -ne 'success') { throw "公式存在错误：$Sheet $Range" }
}
function Invoke-JsonScript([string]$Path,[hashtable]$Parameters) {
    $output = @(& $Path @Parameters 2>&1 | ForEach-Object { [string]$_ })
    if ($LASTEXITCODE -ne 0) { throw "执行失败：$(Split-Path $Path -Leaf)：$($output -join ' ')" }
    $line = @($output | Where-Object { $_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}') } | Select-Object -Last 1)
    if (-not $line) { throw "脚本未返回 JSON：$(Split-Path $Path -Leaf)" }
    return ($line | ConvertFrom-Json)
}

# 重跑历史时保留人工维护字段；兼容升级前的五列旧结构。
$preservedInterventions = @{}
$existingRaw = lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'A1:O1001' --include value --max-chars 500000 --format json | Out-String
if ($LASTEXITCODE -eq 0) {
    $existingResult = $existingRaw | ConvertFrom-Json
    $existingRows = @($existingResult.data.ranges[0].cells)
    $newSchema = $existingRows.Count -gt 0 -and (Get-CellValue $existingRows[0] 3) -eq '订单编号'
    foreach ($row in @($existingRows | Select-Object -Skip 1)) {
        $id = if ($newSchema) { Get-CellValue $row 3 } else { Get-CellValue $row 2 }
        if (-not $id) { continue }
        $preservedInterventions[$id] = [ordered]@{
            Person = if ($newSchema) { Get-CellValue $row 2 } else { $null }
            RestoreState = if ($newSchema) { Get-CellValue $row 9 } else { Get-CellValue $row 3 }
            Reviewed = if ($newSchema) { Get-CellValue $row 10 } else { $null }
            Note = if ($newSchema) { Get-CellValue $row 12 } else { Get-CellValue $row 4 }
            DeliveryType = if ($newSchema) { Get-CellValue $row 5 } else { $null }
            DeliveryFee = if ($newSchema) { Get-CellValue $row 7 } else { $null }
            DailyOrderNo = if ($newSchema) { Get-CellValue $row 1 } else { $null }
        }
    }
}

# 用户已明确授权“先备份再清空”；调用前已创建并回读验证整表备份。
$clearRanges = '["订单汇总!A2:D2001","订单汇总!F2:J2001","订单汇总!M2:Q2001","订单商品!A2:D4001","订单商品!F2:J4001","订单商品!O2:O4001","干预订单!A2:O1001"]'
$null = lark-cli sheets +cells-batch-clear --as user --spreadsheet-token $spreadsheetToken --ranges $clearRanges --scope content --dry-run --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '清空预检失败' }
$null = lark-cli sheets +cells-batch-clear --as user --spreadsheet-token $spreadsheetToken --ranges $clearRanges --scope content --yes --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '清空自动同步区域失败' }

$headers = @('日期','单号','补单人','订单编号','补单金额','配送方式','平台费用（自动）','配送费','总成本','是否已恢复库存','是否已好评','购买内容','备注')
$headerCells = @(); foreach ($header in $headers) { $headerCells += ,(New-ValueCell $header) }
Set-CellMatrix '干预订单' 'A1:M1' @(,$headerCells)
Set-CellMatrix '订单汇总' 'Q1:Q1' @(,@((New-ValueCell '当日单号')))
$interventionCells = @()
foreach ($id in $interventionIds) {
    $manual = if ($preservedInterventions.ContainsKey($id)) { $preservedInterventions[$id] } else { $null }
    $person = if ($manual) { $manual.Person } else { $null }
    $state = if ($manual -and $manual.RestoreState) { $manual.RestoreState } else { '待恢复' }
    $reviewed = if ($manual) { $manual.Reviewed } else { $null }
    $note = if ($manual -and $manual.Note) { $manual.Note } else { '用户确认的历史干预订单' }
    $deliveryType = if ($manual -and $manual.DeliveryType -in @('自配送','骑手配送')) { $manual.DeliveryType } else { $null }
    $deliveryFee = if ($manual) { $manual.DeliveryFee } else { $null }
    $dailyOrderNo = if ($manual) { $manual.DailyOrderNo } else { $null }
    $rowNumber = $interventionCells.Count + 2
    $interventionCells += ,@((New-ValueCell $null),(New-ValueCell $dailyOrderNo),(New-ValueCell $person),(New-ValueCell $id),(New-ValueCell $null),(New-ValueCell $deliveryType),(New-ValueCell $null),(New-ValueCell $deliveryFee),@{formula=('=IF(D{0}="","",IFERROR(G{0}*1,0)+IFERROR(H{0}*1,0))' -f $rowNumber)},(New-ValueCell $state),(New-ValueCell $reviewed),(New-ValueCell $null),(New-ValueCell $note))
}
Set-CellMatrix '干预订单' "A2:M$($interventionIds.Count+1)" $interventionCells

$null = lark-cli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'A1:M1' --background-color '#7A2E2E' --font-color '#FFFFFF' --font-weight bold --horizontal-alignment center --word-wrap auto-wrap --format json | Out-String
$null = lark-cli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'A2:M1001' --background-color '#FFF8E1' --word-wrap auto-wrap --format json | Out-String
$null = lark-cli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'A2:A1001' --number-format 'yyyy-mm-dd' --format json | Out-String
$null = lark-cli sheets +cols-resize --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --widths '{"A":110,"B:C":120,"D":210,"E:I":120,"J:K":130,"L":220,"M":260}' --format json | Out-String

# 补单人和购买内容保持自由文本；配送、库存恢复和好评状态使用下拉。
$wrongDropdownRanges = '["干预订单!C2:C1001","干预订单!L2:L1001"]'
$null = lark-cli sheets +dropdown-delete --as user --spreadsheet-token $spreadsheetToken --ranges $wrongDropdownRanges --yes --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '移除购买内容列的错误下拉失败' }
$null = lark-cli sheets +dropdown-set --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'F2:F1001' --options '["自配送","骑手配送"]' --colors '["#2563EB","#F59E0B"]' --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '设置配送方式下拉失败' }
$null = lark-cli sheets +dropdown-set --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'J2:J1001' --options '["待恢复","已恢复"]' --colors '["#F59E0B","#16A34A"]' --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '设置库存恢复状态下拉失败' }
$null = lark-cli sheets +dropdown-set --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'K2:K1001' --options '["未好评","已好评"]' --colors '["#9CA3AF","#16A34A"]' --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '设置好评状态下拉失败' }

Set-FormulaColumn '补货记录' 'D' 1001 '=IF(B{r}="","",IFERROR(VLOOKUP(B{r},''商品成本''!$A$2:$D$2000,4,FALSE),IFERROR(VLOOKUP(VLOOKUP(B{r},''库存快照''!$B$2:$G$2000,6,FALSE),''商品成本''!$A$2:$D$2000,4,FALSE),"")))'
Set-FormulaColumn '订单汇总' 'E' 2001 '=IF(C{r}="","",IF(COUNTIF(''干预订单''!$D$2:$D$1001,C{r})>0,"干预","自然"))'
Set-FormulaColumn '订单汇总' 'K' 2001 '=IF(C{r}="","",SUMIFS(''订单商品''!$M$2:$M$4001,''订单商品''!$C$2:$C$4001,C{r}))'
Set-FormulaColumn '订单汇总' 'L' 2001 '=IF(C{r}="","",IFERROR(H{r}*1,0)-IFERROR(K{r}*1,0)-IFERROR(J{r}*1,0))'
Set-FormulaColumn '订单商品' 'E' 4001 '=IF(C{r}="","",IF(COUNTIF(''干预订单''!$D$2:$D$1001,C{r})>0,"干预","自然"))'
Set-FormulaColumn '订单商品' 'K' 4001 '=IF(F{r}="","",IFERROR(I{r}*1,0)*IFERROR(J{r}*1,0))'
Set-FormulaColumn '订单商品' 'L' 4001 '=IF(F{r}="","",IFERROR(IF(VLOOKUP(F{r},''商品成本''!$A$2:$E$2000,5,FALSE)="待人工核对","",VLOOKUP(F{r},''商品成本''!$A$2:$D$2000,4,FALSE)),IFERROR(IF(VLOOKUP(VLOOKUP(F{r},''库存快照''!$B$2:$G$2000,6,FALSE),''商品成本''!$A$2:$E$2000,5,FALSE)="待人工核对","",VLOOKUP(VLOOKUP(F{r},''库存快照''!$B$2:$G$2000,6,FALSE),''商品成本''!$A$2:$D$2000,4,FALSE)),"")))'
Set-FormulaColumn '订单商品' 'M' 4001 '=IF(F{r}="","",IF(E{r}="干预",0,IF(L{r}="","",IFERROR(I{r}*1,0)*IFERROR(L{r}*1,0))))'
Set-FormulaColumn '订单商品' 'N' 4001 '=IF(F{r}="","",IF(M{r}="","",K{r}-M{r}))'
Set-FormulaColumn '订单商品' 'P' 4001 '=IF(C{r}="","",IF(E{r}<>"干预","",IF(COUNTIFS(''干预订单''!$D$2:$D$1001,C{r},''干预订单''!$J$2:$J$1001,"已恢复")>0,"已恢复","待恢复")))'
Set-FormulaColumn '库存快照' 'H' 2000 '=IF(B{r}="","",IFERROR(IF(VLOOKUP(B{r},''商品成本''!$A$2:$E$2000,5,FALSE)="待人工核对","",VLOOKUP(B{r},''商品成本''!$A$2:$D$2000,4,FALSE)),IFERROR(IF(VLOOKUP(G{r},''商品成本''!$A$2:$E$2000,5,FALSE)="待人工核对","",VLOOKUP(G{r},''商品成本''!$A$2:$D$2000,4,FALSE)),"")))'
Set-FormulaColumn '库存快照' 'I' 2000 '=IF(OR(B{r}="",D{r}="",H{r}=""),"",D{r}*H{r})'
Set-FormulaColumn '库存快照' 'K' 2000 '=IF(C{r}="","",IF(B{r}="","缺少货号",IF(H{r}="",IF(OR(IFERROR(VLOOKUP(B{r},''商品成本''!$A$2:$E$2000,5,FALSE),"")="待人工核对",IFERROR(VLOOKUP(G{r},''商品成本''!$A$2:$E$2000,5,FALSE),"")="待人工核对"),"成本待核对","缺少成本"),IF(ISNUMBER(SEARCH("重复货号",L{r})),"重复货号待核对",IF(F{r}<>"售卖中","不监控",IF(D{r}<=0,"缺货",IF(D{r}<=J{r},"低库存","正常")))))))'
Set-FormulaColumn '库存快照' 'M' 2000 '=IF(B{r}="","",SUMIFS(''订单商品''!$I$2:$I$4001,''订单商品''!$E$2:$E$4001,"干预",''订单商品''!$D$2:$D$4001,"<>已取消",''订单商品''!$F$2:$F$4001,B{r},''订单商品''!$P$2:$P$4001,"待恢复")+IF(OR(G{r}="",G{r}=B{r}),0,SUMIFS(''订单商品''!$I$2:$I$4001,''订单商品''!$E$2:$E$4001,"干预",''订单商品''!$D$2:$D$4001,"<>已取消",''订单商品''!$F$2:$F$4001,G{r},''订单商品''!$P$2:$P$4001,"待恢复")))'

foreach ($target in @(
    @('干预订单','I2:I1001'),@('补货记录','D2:D1001'),@('订单汇总','E2:E2001,K2:L2001'),
    @('订单商品','E2:E4001,K2:N4001,P2:P4001'),@('库存快照','H2:I2000,K2:K2000,M2:M2000'),@('驾驶舱','C2:C19')
)) { Assert-FormulaRange $target[0] $target[1] }

$dailyResults = @()
for ($date=$StartDate.Date; $date -le $EndDate.Date; $date=$date.AddDays(1)) {
    $dateText = $date.ToString('yyyy-MM-dd')
    Write-Host "SYNC_DAY $dateText"
    $dailyResults += Invoke-JsonScript $orderSyncPath @{StartDate=$dateText;EndDate=$dateText;WorkRoot=$WorkRoot;InterventionOrdersPath=$interventionPath}
}
$inventoryResult = Invoke-JsonScript $inventorySyncPath @{WorkRoot=$WorkRoot}

foreach ($target in @(
    @('干预订单','I2:I1001'),@('补货记录','D2:D1001'),@('订单汇总','E2:E2001,K2:L2001'),
    @('订单商品','E2:E4001,K2:N4001,P2:P4001'),@('库存快照','H2:I2000,K2:K2000,M2:M2000'),@('驾驶舱','C2:C19')
)) { Assert-FormulaRange $target[0] $target[1] }

[ordered]@{
    start_date=$StartDate.ToString('yyyy-MM-dd')
    end_date=$EndDate.ToString('yyyy-MM-dd')
    days=$dailyResults.Count
    extracted_orders=@($dailyResults | Measure-Object extracted_orders -Sum).Sum
    extracted_items=@($dailyResults | Measure-Object extracted_items -Sum).Sum
    natural_orders=@($dailyResults | Measure-Object natural_orders -Sum).Sum
    intervention_orders=@($dailyResults | Measure-Object intervention_orders -Sum).Sum
    intervention_ids=$interventionIds.Count
    inventory=$inventoryResult
    formula_verification='success'
    completed_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
} | ConvertTo-Json -Depth 10
