param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent),
    [string]$SheetName = '9月支出',
    [double]$TestAmount = 12.34
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\daily_ledger_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$syncScript = Join-Path $ProjectRoot 'agent\sync_daily_ledger.ps1'

function Get-Rows {
    $raw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name $SheetName --range 'A2:F1001' --include value --max-chars 500000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '读取支出表失败' }
    $result = $raw | ConvertFrom-Json
    return @($result.data.ranges[0].cells)
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    return $cell.value
}

function Find-DeliveryRow([string]$OrderNo = '') {
    $rows = Get-Rows
    for ($index = 0; $index -lt $rows.Count; $index++) {
        $row = @($rows[$index])
        $project = [string](Get-CellValue $row 1)
        $fee = Get-CellValue $row 2
        $currentOrderNo = [string](Get-CellValue $row 3)
        if ($project -ne '补单配送费') { continue }
        if ($OrderNo -and $currentOrderNo -ne $OrderNo) { continue }
        if (-not $OrderNo -and $null -ne $fee -and [double]$fee -ne 0) { continue }
        return [pscustomobject]@{
            row = $index + 2
            order_no = $currentOrderNo
            fee = if ($null -eq $fee -or "$fee" -eq '') { 0.0 } else { [double]$fee }
            source = [string](Get-CellValue $row 4)
            note = [string](Get-CellValue $row 5)
        }
    }
    return $null
}

function Set-Matrix([string]$Range, $RowCells) {
    $matrix = @()
    $matrix += ,$rowCells
    $cells = ConvertTo-Json -InputObject $matrix -Depth 6 -Compress
    $raw = $cells | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入支出表失败：$raw" }
}

function Set-RowFields([int]$Row, [double]$Fee, [string]$Source, [string]$Note) {
    Set-Matrix -Range "C${Row}:C${Row}" -RowCells @(@{ value = $Fee })
    Set-Matrix -Range "E${Row}:F${Row}" -RowCells @(@{ value = $Source }, @{ value = $Note })
}

function Run-LedgerSync {
    # Never run a production-table regression test with a historical cutoff.
    # ThroughDate requests a summary date; using today prevents a test from
    # temporarily truncating newer month rows if an older sync implementation is present.
    $output = @(& $syncScript -ThroughDate ([datetime]::Today) -ProjectRoot $ProjectRoot)
    if ($LASTEXITCODE -ne 0) { throw '简账同步失败' }
    $jsonLine = @($output | ForEach-Object { [string]$_ } | Where-Object { $_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}') } | Select-Object -Last 1)
    if (-not $jsonLine) { throw '简账同步未返回结果' }
}

$target = Find-DeliveryRow
if ($null -eq $target) { throw '没有找到可测试的待补录配送费行' }
$passed = $false
try {
    Set-RowFields -Row $target.row -Fee $TestAmount -Source $target.source -Note $target.note
    Run-LedgerSync
    $after = Find-DeliveryRow -OrderNo $target.order_no
    if ($null -eq $after) { throw '同步后未找到测试订单' }
    $passed = ([math]::Abs($after.fee - $TestAmount) -lt 0.000001 -and $after.source -eq '人工补录')
    if (-not $passed) { throw "人工配送费未被保留：fee=$($after.fee), source=$($after.source)" }
} finally {
    $current = Find-DeliveryRow -OrderNo $target.order_no
    if ($null -ne $current) {
        Set-RowFields -Row $current.row -Fee $target.fee -Source $target.source -Note $target.note
        Run-LedgerSync
    }
}

$restored = Find-DeliveryRow -OrderNo $target.order_no
$clean = ($null -ne $restored -and [math]::Abs($restored.fee - $target.fee) -lt 0.000001 -and $restored.source -eq $target.source)
[pscustomobject]@{
    ok = ($passed -and $clean)
    order_no = $target.order_no
    manual_value_preserved = $passed
    test_data_removed = $clean
} | ConvertTo-Json -Compress
