param(
    [string]$ProjectRoot,
    [string]$SourceBackup = 'agent\state\sku_cost_master_auto_backup_20260915_002358.json'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $ProjectRoot 'agent\sku_master_schema.ps1')
$schema = Get-SkuMasterSchema -SpreadsheetToken $token
$sourcePath = Join-Path $ProjectRoot $SourceBackup
if (-not (Test-Path -LiteralPath $sourcePath)) { throw "找不到恢复源：$sourcePath" }

function Get-Value($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count -or $null -eq @($Row)[$Index] -or $null -eq @($Row)[$Index].value) { return $null }
    $value = @($Row)[$Index].value
    if ([string]$value -eq '') { return $null }
    return $value
}
function New-Cell($Value) { if ($null -eq $Value -or [string]$Value -eq '') { return @{value=''} }; return @{value=$Value} }
function Set-Matrix([string]$Range, $Rows) {
    $payload = ConvertTo-Json -InputObject $Rows -Depth 10 -Compress
    $raw = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入失败：$Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "写入未成功：$Range" }
}

$raw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A2:Z3000' --include value --max-chars 1600000 --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '读取当前SKU成本主表失败' }
$currentResult = $raw | ConvertFrom-Json
if (-not $currentResult.ok -or $currentResult.data.has_more) { throw '当前SKU成本主表读取不完整' }
$currentRows = @($currentResult.data.ranges[0].cells)
$currentIndices = @($currentResult.data.ranges[0].row_indices)

$timestamp = [datetime]::Now.ToString('yyyyMMdd_HHmmss')
$safetyBackup = Join-Path $ProjectRoot "agent\state\sku_cost_master_schema_shift_backup_$timestamp.json"
[ordered]@{source='SKU成本主表!A2:Z3000';captured_at=[datetime]::Now.ToString('o');row_indices=$currentIndices;cells=$currentRows} |
    ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $safetyBackup -Encoding UTF8

$old = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8 | ConvertFrom-Json
$oldRows = @($old.cells)
$oldIndices = @($old.row_indices)
$oldByRow = @{}
for ($i = 0; $i -lt $oldRows.Count; $i++) {
    $rowNumber = if ($i -lt $oldIndices.Count) { [int]$oldIndices[$i] } else { 2 + $i }
    $oldByRow[$rowNumber] = $oldRows[$i]
}

$baselineMatrix = [Collections.Generic.List[object]]::new()
$initializationMatrix = [Collections.Generic.List[object]]::new()
$matched = 0
$unmatched = [Collections.Generic.List[int]]::new()
$lastRow = if ($currentIndices.Count -gt 0) { [int](($currentIndices | Measure-Object -Maximum).Maximum) } else { 1 }
$currentByRow = @{}
for ($i = 0; $i -lt $currentRows.Count; $i++) {
    $rowNumber = if ($i -lt $currentIndices.Count) { [int]$currentIndices[$i] } else { 2 + $i }
    $currentByRow[$rowNumber] = $currentRows[$i]
}

for ($rowNumber = 2; $rowNumber -le $lastRow; $rowNumber++) {
    $current = if ($currentByRow.ContainsKey($rowNumber)) { $currentByRow[$rowNumber] } else { @() }
    $baseline = Get-Value $current $schema.BaselineUnitCost.Index
    $initializationDate = Get-Value $current $schema.InitializationDate.Index
    $isPopulated = (Get-Value $current $schema.CurrentKw.Index) -or (Get-Value $current $schema.TransitionCode.Index) -or (Get-Value $current $schema.PlatformTitle.Index) -or (Get-Value $current $schema.ProductId.Index)
    if ($isPopulated -and $oldByRow.ContainsKey($rowNumber)) {
        $previous = $oldByRow[$rowNumber]
        # 恢复源是插入“自定义标题”列之前的18列表：基准成本=旧E，初始化日期=旧K。
        $sameIdentity = $false
        $currentProductId = [string](Get-Value $current $schema.ProductId.Index)
        $oldProductId = [string](Get-Value $previous 15)
        if ($currentProductId -and $oldProductId) {
            $sameIdentity = $currentProductId -eq $oldProductId
        } else {
            $sameIdentity = ([string](Get-Value $current $schema.CurrentKw.Index) -eq [string](Get-Value $previous 0)) -and
                ([string](Get-Value $current $schema.TransitionCode.Index) -eq [string](Get-Value $previous 1)) -and
                ([string](Get-Value $current $schema.PlatformTitle.Index) -eq [string](Get-Value $previous 2))
        }
        if ($sameIdentity) {
            $baseline = Get-Value $previous 4
            $initializationDate = Get-Value $previous 10
            $matched++
        } else {
            $unmatched.Add($rowNumber)
        }
    } elseif ($isPopulated) {
        $unmatched.Add($rowNumber)
    }
    $baselineMatrix.Add(@((New-Cell $baseline)))
    $initializationMatrix.Add(@((New-Cell $initializationDate)))
}

if ($matched -eq 0) { throw '没有匹配到可恢复的SKU，已停止写入' }
Set-Matrix "$($schema.BaselineUnitCost.Column)2:$($schema.BaselineUnitCost.Column)$lastRow" $baselineMatrix.ToArray()
Set-Matrix "$($schema.InitializationDate.Column)2:$($schema.InitializationDate.Column)$lastRow" $initializationMatrix.ToArray()

$verifyRaw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A2:Z3000' --include value --max-chars 1600000 --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '恢复后回读失败' }
$verify = $verifyRaw | ConvertFrom-Json
$verifyRows = @($verify.data.ranges[0].cells)
$verifyIndices = @($verify.data.ranges[0].row_indices)
$verifyByRow = @{}
for ($i = 0; $i -lt $verifyRows.Count; $i++) { $verifyByRow[[int]$verifyIndices[$i]] = $verifyRows[$i] }
$errors = [Collections.Generic.List[int]]::new()
foreach ($rowNumber in $oldByRow.Keys) {
    if (-not $verifyByRow.ContainsKey([int]$rowNumber)) { continue }
    $current = $verifyByRow[[int]$rowNumber]
    $previous = $oldByRow[[int]$rowNumber]
    $currentProductId = [string](Get-Value $current $schema.ProductId.Index)
    $oldProductId = [string](Get-Value $previous 15)
    if ($currentProductId -and $oldProductId -and $currentProductId -eq $oldProductId) {
        if ([string](Get-Value $current $schema.BaselineUnitCost.Index) -ne [string](Get-Value $previous 4) -or
            [string](Get-Value $current $schema.InitializationDate.Index) -ne [string](Get-Value $previous 10)) { $errors.Add([int]$rowNumber) }
    }
}
if ($errors.Count -gt 0) { throw "恢复回读校验失败，行：$($errors -join ',')" }

[ordered]@{
    ok = $true
    restored_rows = $matched
    unmatched_rows = @($unmatched)
    safety_backup = $safetyBackup
    source_backup = $sourcePath
    restored_columns = @($schema.BaselineUnitCost.Header, $schema.InitializationDate.Header)
    readback_verification = 'success'
} | ConvertTo-Json -Depth 5
