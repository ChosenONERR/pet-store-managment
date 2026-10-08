param(
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少经营驾驶舱元数据' }
$meta = [System.IO.File]::ReadAllText($metaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')

foreach ($dateRange in @('A2:A1001','L2:L1001','P2:P1001','S2:S1001')) {
    $raw = lark-cli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range $dateRange --number-format 'yyyy-mm-dd' --horizontal-alignment center --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "更新补货记录日期格式失败：$dateRange" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "更新补货记录日期格式未成功：$dateRange" }
}

# 早期个别日期曾以文本“46277”写入。数字格式无法作用于文本，需原位改回数值日期。
$normalizedTextSerials = @()
foreach ($dateColumn in @('A','L','P','S')) {
    $columnRaw = lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range "${dateColumn}2:${dateColumn}1001" --include value --max-chars 200000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取补货记录日期列失败：$dateColumn" }
    $columnResult = $columnRaw | ConvertFrom-Json
    if (-not $columnResult.ok -or $columnResult.data.has_more) { throw "补货记录日期列读取不完整：$dateColumn" }
    $rangeResult = $columnResult.data.ranges[0]
    for ($index = 0; $index -lt @($rangeResult.cells).Count; $index++) {
        $cell = @($rangeResult.cells[$index])[0]
        if ($null -eq $cell -or $null -eq $cell.value) { continue }
        $rawValue = [string]$cell.value
        if ($rawValue -notmatch '^\d{5}(?:\.\d+)?$') { continue }
        $serial = [double]::Parse($rawValue, [Globalization.CultureInfo]::InvariantCulture)
        if ($serial -lt 30000 -or $serial -gt 80000) { continue }
        $rowNumber = [int]$rangeResult.row_indices[$index]
        $cellRow = @()
        $cellRow += ,@{value=$serial;cell_styles=@{number_format='yyyy-mm-dd';horizontal_alignment='center'}}
        $cellMatrix = @()
        $cellMatrix += ,$cellRow
        $payload = ConvertTo-Json -InputObject $cellMatrix -Depth 8 -Compress
        $writeRaw = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range "${dateColumn}${rowNumber}:${dateColumn}${rowNumber}" --cells - --format json | Out-String
        if ($LASTEXITCODE -ne 0 -or -not (($writeRaw | ConvertFrom-Json).ok)) { throw "日期文本转数值失败：${dateColumn}${rowNumber}" }
        $normalizedTextSerials += "${dateColumn}${rowNumber}"
    }
}

$formulaRaw = lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range 'D2:E1001,G2:G1001' --format json --exit-on-error | Out-String
if ($LASTEXITCODE -ne 0) { throw '补货记录公式校验失败' }
$formulaResult = $formulaRaw | ConvertFrom-Json
if ($formulaResult.data.status -ne 'success') { throw '补货记录公式存在错误' }

$readRaw = lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name '补货记录' --range 'A1:T4' --include value,formula,style --max-chars 100000 --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '补货记录回读失败' }
$readResult = $readRaw | ConvertFrom-Json
if (-not $readResult.ok -or $readResult.data.has_more) { throw '补货记录回读结果不完整' }

$rows = @($readResult.data.ranges[0].cells)
$replenishmentDates = @()
for ($index = 1; $index -lt $rows.Count; $index++) {
    if ($rows[$index].Count -gt 0 -and $null -ne $rows[$index][0].value) {
        $rawDate = [string]$rows[$index][0].value
        $parsedDate = [datetime]::MinValue
        $serial = 0.0
        if ([datetime]::TryParse($rawDate, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
            $replenishmentDates += $parsedDate.ToString('yyyy-MM-dd')
        } elseif ([double]::TryParse($rawDate, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$serial)) {
            $replenishmentDates += ([datetime]'1899-12-30').AddDays($serial).ToString('yyyy-MM-dd')
        } else {
            throw "补货记录日期无法解析：$rawDate"
        }
    }
}

[ordered]@{
    ok = $true
    formatted_ranges = @('A2:A1001','L2:L1001','P2:P1001','S2:S1001')
    normalized_text_serial_cells = $normalizedTextSerials
    replenishment_dates = $replenishmentDates
    formula_verification = $formulaResult.data.status
    workbook_url = $meta.workbook_url
} | ConvertTo-Json -Depth 6 -Compress
