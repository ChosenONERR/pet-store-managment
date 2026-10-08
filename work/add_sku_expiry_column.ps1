param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
if (-not (Test-Path -LiteralPath $metaPath -PathType Leaf)) { throw '缺少飞书经营驾驶舱元数据' }
$meta = [System.IO.File]::ReadAllText($metaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$larkCli = Join-Path $env:APPDATA 'npm\lark-cli.cmd'
if (-not (Test-Path -LiteralPath $larkCli -PathType Leaf)) { throw "找不到飞书 CLI：$larkCli" }

function Invoke-LarkJson([scriptblock]$Command, [string]$Label) {
    $raw = & $Command | Out-String
    if ($LASTEXITCODE -ne 0) { throw $Label }
    $result = $raw | ConvertFrom-Json
    if ($result.ok -eq $false) { throw "$Label：$($result.error.message)" }
    return $result
}

$existing = Invoke-LarkJson {
    & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'O1:O2' --include value --max-chars 2000 --format json
} '无法读取 SKU成本主表 O 列'
$cells = @($existing.data.ranges[0].cells)
$header = if ($cells.Count -gt 0 -and @($cells[0]).Count -gt 0) { [string]$cells[0][0].value } else { '' }
if ($header -and $header -ne '最早过期日期') { throw "SKU成本主表 O 列已有其他字段：$header，未覆盖" }

$headerMatrix = @(); $headerMatrix += ,@(@{
    value = '最早过期日期'
    cell_styles = @{
        background_color = '#315B7D'
        font_color = '#FFFFFF'
        font_weight = 'bold'
        horizontal_alignment = 'center'
        word_wrap = 'auto-wrap'
    }
})
$headerPayload = ConvertTo-Json -InputObject $headerMatrix -Depth 8 -Compress
$null = Invoke-LarkJson {
    $headerPayload | & $larkCli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'O1:O1' --cells - --format json
} '写入最早过期日期表头失败'
$null = Invoke-LarkJson {
    & $larkCli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'O2:O1500' --background-color '#FFF8E1' --number-format 'yyyy-mm-dd' --horizontal-alignment center --format json
} '设置最早过期日期列样式失败'
$null = Invoke-LarkJson {
    & $larkCli sheets +cols-resize --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'O' --width 130 --format json
} '设置最早过期日期列宽失败'
$filterState = Invoke-LarkJson {
    & $larkCli sheets +filter-list --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --format json
} '读取 SKU成本主表筛选器失败'
$filterSheet = @($filterState.data.sheets | Where-Object { $_.sheet_name -eq 'SKU成本主表' } | Select-Object -First 1)
if ($filterSheet.Count -ne 1 -or -not $filterSheet[0].sheet_id) { throw '无法取得 SKU成本主表 sheet_id' }
$sheetId = [string]$filterSheet[0].sheet_id
$filterProperties = @{rules=@()} | ConvertTo-Json -Compress
$null = Invoke-LarkJson {
    $filterProperties | & $larkCli sheets +filter-update --as user --spreadsheet-token $spreadsheetToken --sheet-id $sheetId --range 'A1:O1500' --properties - --format json
} '扩展 SKU成本主表筛选范围失败'
$formula = Invoke-LarkJson {
    & $larkCli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name '驾驶舱' --range 'C2:C19' --format json --exit-on-error
} '驾驶舱公式校验失败'
if ($formula.data.status -ne 'success') { throw '驾驶舱公式存在错误' }

$readback = Invoke-LarkJson {
    & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'O1:O2' --include value,styles --max-chars 5000 --format json
} '回读最早过期日期列失败'
$readbackCells = @($readback.data.ranges[0].cells)
$readbackHeader = [string]$readbackCells[0][0].value
if ($readbackHeader -ne '最早过期日期') { throw '最早过期日期表头回读不一致' }

[ordered]@{
    ok = $true
    sheet = 'SKU成本主表'
    column = 'O'
    header = $readbackHeader
    filter_range = 'A1:O1500'
    formula_verification = 'success'
    completed_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
} | ConvertTo-Json -Compress
