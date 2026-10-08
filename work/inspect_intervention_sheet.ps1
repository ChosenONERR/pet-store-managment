param(
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
. (Join-Path $ProjectRoot 'agent\runtime_paths.ps1')
$larkCli = Get-PetstarLarkCli
if (-not (Test-Path -LiteralPath $larkCli)) { throw '未找到已授权的 lark-cli' }

$meta = [System.IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$statePath = Join-Path $ProjectRoot 'agent\state\intervention_sheet_inspect.json'

$raw = & $larkCli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '干预订单' --range 'A1:Z1001' --include value,formula --max-chars 5000000 --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '读取干预订单失败' }
$result = $raw | ConvertFrom-Json
if (-not $result.ok -or $result.data.has_more) { throw '干预订单读取结果不完整' }
$range = $result.data.ranges[0]
$headers = @()
if (@($range.cells).Count -gt 0) {
    for ($i=0; $i -lt @($range.cells[0]).Count; $i++) {
        $cell = @($range.cells[0])[$i]
        $headers += [pscustomobject]@{
            column = [string]@($range.col_indices)[$i]
            value = if ($null -ne $cell -and $null -ne $cell.value) { [string]$cell.value } else { '' }
            formula = if ($null -ne $cell -and $null -ne $cell.formula) { [string]$cell.formula } else { '' }
        }
    }
}
$usedRows = @()
for ($i=1; $i -lt @($range.cells).Count; $i++) {
    $row = @($range.cells)[$i]
    $values = @()
    $formulas = @()
    $hasContent = $false
    foreach ($cell in @($row)) {
        $value = if ($null -ne $cell -and $null -ne $cell.value) { [string]$cell.value } else { '' }
        $formula = if ($null -ne $cell -and $null -ne $cell.formula) { [string]$cell.formula } else { '' }
        $values += $value
        $formulas += $formula
        if ($value -or $formula) { $hasContent = $true }
    }
    if ($hasContent) {
        $usedRows += [pscustomobject]@{
            row = [int]@($range.row_indices)[$i]
            values = $values
            formulas = $formulas
        }
    }
}

$summary = [ordered]@{
    ok = $true
    inspected_at = [datetime]::Now.ToString('s')
    actual_range = $range.actual_range
    headers = $headers
    used_row_count = $usedRows.Count
    rows = $usedRows
}
[System.IO.File]::WriteAllText($statePath, ($summary | ConvertTo-Json -Depth 12), [System.Text.UTF8Encoding]::new($false))
[ordered]@{ok=$true; state_file=$statePath; used_row_count=$usedRows.Count; headers=$headers} | ConvertTo-Json -Depth 6 -Compress
