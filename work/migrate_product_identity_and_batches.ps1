param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')

function Invoke-LarkJson([scriptblock]$Command, [string]$Label) {
    $raw = & $Command | Out-String
    if ($LASTEXITCODE -ne 0) { throw "$Label：$($raw.Trim())" }
    $result = $raw | ConvertFrom-Json
    if ($result.ok -eq $false) { throw "$Label：$($result.error.message)" }
    return $result
}
function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count -or $null -eq @($Row)[$Index].value) { return $null }
    return [string]@($Row)[$Index].value
}
function New-ValueCell($Value) { if ($null -eq $Value -or [string]$Value -eq '') { return @{value=''} }; return @{value=$Value} }
function New-FormulaCell([string]$Formula) { if (-not $Formula) { return @{value=''} }; return @{formula=$Formula} }
function New-ReplenishmentProductNameFormula([int]$RowNumber) {
    $kwLookup = "IFERROR(IF(VLOOKUP(B${RowNumber},'SKU成本主表'!`$A`$2:`$D`$3000,4,FALSE)<>`"`",VLOOKUP(B${RowNumber},'SKU成本主表'!`$A`$2:`$D`$3000,4,FALSE),VLOOKUP(B${RowNumber},'SKU成本主表'!`$A`$2:`$D`$3000,3,FALSE)),`"`")"
    $idLookup = "IFERROR(IF(INDEX('SKU成本主表'!`$D`$2:`$D`$3000,MATCH(O${RowNumber},'SKU成本主表'!`$Q`$2:`$Q`$3000,0))<>`"`",INDEX('SKU成本主表'!`$D`$2:`$D`$3000,MATCH(O${RowNumber},'SKU成本主表'!`$Q`$2:`$Q`$3000,0)),INDEX('SKU成本主表'!`$C`$2:`$C`$3000,MATCH(O${RowNumber},'SKU成本主表'!`$Q`$2:`$Q`$3000,0))),${kwLookup})"
    return "=IF(B${RowNumber}=`"`",`"`",IF(O${RowNumber}<>`"`",${idLookup},${kwLookup}))"
}
function Set-Cells([string]$Sheet,[string]$Range,$Rows) {
    $rowValues = @($Rows)
    $matrix = [Collections.ArrayList]::new()
    if ($rowValues.Count -gt 0 -and $rowValues[0] -is [Collections.IDictionary]) {
        $null = $matrix.Add([object[]]$rowValues)
    } else {
        foreach ($row in $rowValues) { $null = $matrix.Add([object[]]@($row)) }
    }
    $payload = ConvertTo-Json -InputObject $matrix -Depth 12 -Compress
    $null = Invoke-LarkJson { $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name $Sheet --range $Range --cells - --format json } "写入失败 $Sheet $Range"
}

$info = Invoke-LarkJson { lark-cli sheets +workbook-info --as user --spreadsheet-token $spreadsheetToken --format json } '读取工作簿结构失败'
$sheetTitles = @($info.data.sheets | ForEach-Object { [string]$_.sheet_name })
if ('库存批次' -notin $sheetTitles) {
    $null = Invoke-LarkJson { lark-cli sheets +sheet-create --as user --spreadsheet-token $spreadsheetToken --title '库存批次' --row-count 2000 --col-count 13 --format json } '创建库存批次表失败'
}

$headerStyle = @{background_color='#153F32';font_weight='bold';font_color='#FFFFFF';horizontal_alignment='center';vertical_alignment='center';word_wrap='auto-wrap'}
Set-Cells '库存快照' 'N1:O1' @(@(@{value='美团SKU ID';cell_styles=$headerStyle},@{value='规格名称';cell_styles=$headerStyle}))
Set-Cells 'SKU成本主表' 'A1:B1' @(@(@{value='当前KW';cell_styles=$headerStyle},@{value='CSX/过渡货号';cell_styles=$headerStyle}))
Set-Cells 'SKU成本主表' 'O1:R1' @(@(
    @{value='当前库存最早到期日期';cell_styles=$headerStyle},@{value='商品ID（美团SKU ID）';cell_styles=$headerStyle},
    @{value='平台销售状态';cell_styles=$headerStyle},@{value='当前规格';cell_styles=$headerStyle}
))
Set-Cells '订单商品' 'Q1:Q1' @(@(@{value='商品ID（美团SKU ID）';cell_styles=$headerStyle}))
Set-Cells '补货记录' 'O1:T1' @(@(
    @{value='商品ID';cell_styles=$headerStyle},@{value='生产日期';cell_styles=$headerStyle},@{value='保质期数值';cell_styles=$headerStyle},
    @{value='保质期单位';cell_styles=$headerStyle},@{value='到期日期';cell_styles=$headerStyle},@{value='批次ID';cell_styles=$headerStyle}
))
Set-Cells '补货记录' 'U1:U1' @(@(@{value='商品名';cell_styles=$headerStyle}))
$productNameCells = @()
for ($rowNumber = 2; $rowNumber -le 1001; $rowNumber++) { $productNameCells += ,@((New-FormulaCell (New-ReplenishmentProductNameFormula $rowNumber))) }
Set-Cells '补货记录' 'U2:U1001' $productNameCells
$batchHeaders = @('批次ID','商品ID','当前KW','入库日期','入库销售单位数量','生产日期','保质期数值','保质期单位','到期日期','剩余数量','来源记录','状态','备注')
$batchCells = @(); foreach ($header in $batchHeaders) { $batchCells += ,@{value=$header;cell_styles=$headerStyle} }
Set-Cells '库存批次' 'A1:M1' @($batchCells)

# B列只保留当前 CSX 过渡货号，删除过去被错误保存进去的 KW 历史别名。
$master = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'A2:B3000' --include value --max-chars 500000 --format json } '读取SKU过渡货号失败'
$rows = @($master.data.ranges[0].cells)
$rowNumbers = @($master.data.ranges[0].row_indices)
$cleaned = 0
for ($i=0; $i -lt $rows.Count; $i++) {
    $value = Get-CellValue $rows[$i] 1
    if (-not $value) { continue }
    $parts = @([regex]::Split($value, '[/,，、;；\s]+') | Where-Object { $_ -and $_ -notmatch '^(?i)kw\d+$' } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique)
    $next = $parts -join '、'
    if ($next -ne $value) {
        $rowNumber = if ($i -lt $rowNumbers.Count) { [int]$rowNumbers[$i] } else { 2 + $i }
        Set-Cells 'SKU成本主表' "B${rowNumber}:B${rowNumber}" @(@(New-ValueCell $next))
        $cleaned++
    }
}

$null = Invoke-LarkJson { lark-cli sheets +dim-freeze --as user --spreadsheet-token $spreadsheetToken --sheet-name '库存批次' --dimension row --count 1 --format json } '冻结库存批次表头失败'
$null = Invoke-LarkJson { lark-cli sheets +range-copy --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --source-range 'B2:B1001' --target-range 'U2' --paste-type formats --format json } '复制补货商品名列格式失败'
$null = Invoke-LarkJson { lark-cli sheets +cols-resize --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range 'U' --width 360 --format json } '设置补货商品名列宽失败'
$null = Invoke-LarkJson { lark-cli sheets +cols-resize --as user --spreadsheet-token $spreadsheetToken --sheet-name '库存批次' --widths '{"A:B":150,"C:D":105,"E:J":115,"K":170,"L":90,"M":220}' --format json } '设置库存批次列宽失败'
$null = Invoke-LarkJson { lark-cli sheets +sheet-hide --as user --spreadsheet-token $spreadsheetToken --sheet-name '库存批次' --format json } '隐藏库存批次表失败'

$verify = Invoke-LarkJson { lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'A1:R3000' --format json --exit-on-error } 'SKU成本主表公式校验失败'
if ($verify.data.status -ne 'success') { throw 'SKU成本主表公式存在错误' }
$replenishmentVerify = Invoke-LarkJson { lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range 'U2:U1001' --format json --exit-on-error } '补货商品名公式校验失败'
if ($replenishmentVerify.data.status -ne 'success') { throw '补货商品名公式存在错误' }

[ordered]@{ok=$true;cleaned_historical_kw_aliases=$cleaned;batch_sheet='库存批次';formula_verification='success'} | ConvertTo-Json -Compress
