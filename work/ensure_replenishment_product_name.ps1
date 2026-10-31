param([string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent))

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$agentRoot = Join-Path $ProjectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$meta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
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
    return [string]$cell.value
}

function Get-Rows([string]$Sheet,[string]$Range,[string]$Include = 'value') {
    $result = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $Sheet --range $Range --include $Include --max-chars 3000000 --format json } "读取飞书失败：$Sheet $Range"
    if ($result.data.has_more) { throw "读取飞书结果不完整：$Sheet $Range" }
    return [pscustomobject]@{
        rows = [object[]]@($result.data.ranges[0].cells)
        row_indices = [int[]]@($result.data.ranges[0].row_indices)
    }
}

function New-ProductNameFormula([int]$RowNumber) {
    $kwLookup = "IFERROR(IF(VLOOKUP(B${RowNumber},'SKU成本主表'!`$A`$2:`$D`$3000,4,FALSE)<>`"`",VLOOKUP(B${RowNumber},'SKU成本主表'!`$A`$2:`$D`$3000,4,FALSE),VLOOKUP(B${RowNumber},'SKU成本主表'!`$A`$2:`$D`$3000,3,FALSE)),`"`")"
    $idLookup = "IFERROR(IF(INDEX('SKU成本主表'!`$D`$2:`$D`$3000,MATCH(O${RowNumber},'SKU成本主表'!`$Q`$2:`$Q`$3000,0))<>`"`",INDEX('SKU成本主表'!`$D`$2:`$D`$3000,MATCH(O${RowNumber},'SKU成本主表'!`$Q`$2:`$Q`$3000,0)),INDEX('SKU成本主表'!`$C`$2:`$C`$3000,MATCH(O${RowNumber},'SKU成本主表'!`$Q`$2:`$Q`$3000,0))),${kwLookup})"
    return "=IF(B${RowNumber}=`"`",`"`",IF(O${RowNumber}<>`"`",${idLookup},${kwLookup}))"
}

$workbook = Invoke-LarkJson { & $larkCli sheets +workbook-info --as user --spreadsheet-token $spreadsheetToken --format json } '读取工作簿结构失败'
$replenishmentSheet = @($workbook.data.sheets | Where-Object { [string]$_.sheet_name -eq '补货记录' } | Select-Object -First 1)
if ($replenishmentSheet.Count -ne 1) { throw '工作簿中找不到唯一的“补货记录”工作表' }
$columnCount = [int]$replenishmentSheet[0].column_count
if ($columnCount -lt 20) { throw "补货记录列数不足：当前 $columnCount 列，预期至少 20 列" }
$currentRange = if ($columnCount -ge 21) { 'A1:U1001' } else { 'A1:T1001' }
$current = Get-Rows '补货记录' $currentRange 'value,formula,style'
$headerRow = if ($current.rows.Count -gt 0) { @($current.rows[0]) } else { @() }
$headers = @{}
for ($index = 0; $index -lt $headerRow.Count; $index++) {
    $header = Get-CellValue $headerRow $index
    if ($header) { $headers[$header.Trim()] = $index }
}
if ($headers.ContainsKey('商品名') -and [int]$headers['商品名'] -ne 20) { throw '补货记录已在其他列存在“商品名”，为避免重复列已停止' }
$currentUHeader = Get-CellValue $headerRow 20
if ($currentUHeader -and $currentUHeader -ne '商品名') { throw "补货记录 U1 已被占用：$currentUHeader" }
if (-not $currentUHeader) {
    for ($rowIndex = 1; $rowIndex -lt $current.rows.Count; $rowIndex++) {
        $row = @($current.rows[$rowIndex])
        if ($row.Count -le 20) { continue }
        $cell = $row[20]
        if ($null -ne $cell -and ($cell.PSObject.Properties['value'] -or $cell.PSObject.Properties['formula'])) {
            $value = if ($cell.PSObject.Properties['value']) { [string]$cell.value } else { '' }
            $formula = if ($cell.PSObject.Properties['formula']) { [string]$cell.formula } else { '' }
            if ($value -or $formula) { throw '补货记录 U 列已有未识别内容，未覆盖' }
        }
    }
}

$formulaCells = [Collections.Generic.List[object]]::new()
for ($rowNumber = 2; $rowNumber -le 1001; $rowNumber++) { $formulaCells.Add(@(@{formula=(New-ProductNameFormula $rowNumber)})) }
$payload = ConvertTo-Json -InputObject $formulaCells.ToArray() -Depth 10 -Compress

$transaction = $null
$closed = $false
$insertedColumn = $false
try {
    if ($columnCount -eq 20) {
        $null = Invoke-LarkJson { & $larkCli sheets +dim-insert --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --position U --count 1 --inherit-style before --format json } '追加商品名列失败'
        $insertedColumn = $true
    }
    $transaction = New-FinanceTransaction $stateRoot '补货记录增加商品名列' ('TX-REPLENISHMENT-PRODUCT-NAME-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')) -AcquireLock -WorkbookTokens @{source=$spreadsheetToken} -LarkCli $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $spreadsheetToken '补货记录' 'U1:U1001' $larkCli
    $null = Invoke-LarkJson { & $larkCli sheets +range-copy --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --source-range 'B1:B1001' --target-range 'U1' --paste-type formats --format json } '复制商品名列格式失败'
    $headerPayload = ConvertTo-Json -InputObject @(@(@{value='商品名'})) -Depth 6 -Compress
    $null = Invoke-LarkJson { $headerPayload | & $larkCli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range 'U1:U1' --cells - --format json } '写入商品名表头失败'
    $null = Invoke-LarkJson { $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range 'U2:U1001' --cells - --format json } '写入商品名公式失败'
    $null = Invoke-LarkJson { & $larkCli sheets +cols-resize --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range 'U' --width 360 --format json } '设置商品名列宽失败'
    $verify = Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range 'U2:U1001' --format json --exit-on-error } '商品名公式校验失败'
    if ($verify.data.status -ne 'success') { throw '商品名公式未通过校验' }

    $after = Get-Rows '补货记录' 'A1:U1001' 'value,formula,style'
    if ($after.rows.Count -eq 0 -or (Get-CellValue $after.rows[0] 20) -ne '商品名') { throw '商品名表头回读不一致' }
    $master = Get-Rows 'SKU成本主表' 'A2:Q3000' 'value'
    $byKw = [ordered]@{}
    $byProductId = [ordered]@{}
    foreach ($row in @($master.rows)) {
        $kw = (Get-CellValue $row 0)
        $customTitle = Get-CellValue $row 3
        $platformTitle = Get-CellValue $row 2
        $title = if ($customTitle) { $customTitle } else { $platformTitle }
        if (-not $title) { continue }
        if ($kw -and -not $byKw.Contains($kw.Trim().ToLowerInvariant())) { $byKw[$kw.Trim().ToLowerInvariant()] = $title }
        $productId = Get-CellValue $row 16
        if ($productId -and -not $byProductId.Contains($productId.Trim())) { $byProductId[$productId.Trim()] = $title }
    }
    $mapped = 0
    $mismatches = [Collections.Generic.List[string]]::new()
    for ($rowIndex = 1; $rowIndex -lt $after.rows.Count; $rowIndex++) {
        $row = @($after.rows[$rowIndex])
        $goodsCode = Get-CellValue $row 1
        if (-not $goodsCode) { continue }
        $productId = Get-CellValue $row 14
        $expected = $null
        if ($productId -and $byProductId.Contains($productId.Trim())) { $expected = [string]$byProductId[$productId.Trim()] }
        elseif ($byKw.Contains($goodsCode.Trim().ToLowerInvariant())) { $expected = [string]$byKw[$goodsCode.Trim().ToLowerInvariant()] }
        if (-not $expected) { continue }
        $mapped++
        $actual = Get-CellValue $row 20
        if ($actual -ne $expected) {
            $absoluteRow = if ($rowIndex -lt $after.row_indices.Count) { [int]$after.row_indices[$rowIndex] } else { $rowIndex + 1 }
            $mismatches.Add("第${absoluteRow}行 $goodsCode")
        }
    }
    if ($mismatches.Count -gt 0) { throw "商品名映射回读不一致：$($mismatches -join '、')" }
    if ($mapped -eq 0) { throw '没有找到可用于抽查的补货商品名映射' }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{ok=$true;column='U';header='商品名';formula_rows=1000;mapped_rows_verified=$mapped} | ConvertTo-Json -Compress
} catch {
    $failure = $_.Exception.Message
    if ($null -ne $transaction -and -not $closed) {
        $rollback = Invoke-FinanceTransactionRollback $transaction @{source=$spreadsheetToken} $larkCli $failure
        Close-FinanceTransaction $transaction
        $closed = $true
        if (-not $rollback.ok) { throw "$failure；自动回滚失败：$($rollback.error)" }
    }
    if ($insertedColumn) {
        try {
            $null = Invoke-LarkJson { & $larkCli sheets +dim-delete --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range U --yes --format json } '删除失败升级新增列失败'
            $insertedColumn = $false
        } catch {
            throw "$failure；单元格已回滚，但新增空列 U 删除失败：$($_.Exception.Message)"
        }
    }
    throw $failure
} finally {
    if ($null -ne $transaction -and -not $closed) { Close-FinanceTransaction $transaction }
}
