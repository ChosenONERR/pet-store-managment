param(
    [string]$ProjectRoot,
    [datetime]$SnapshotDate = [datetime]::Today
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
$stateRoot = Join-Path $ProjectRoot 'agent\state'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少飞书经营驾驶舱元数据' }
if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
$meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $PSScriptRoot 'sku_master_schema.ps1')
$skuSchema = Get-SkuMasterSchema -SpreadsheetToken $spreadsheetToken

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge $Row.Count -or $null -eq $Row[$Index]) { return $null }
    $valueProperty = $Row[$Index].PSObject.Properties['value']
    if ($null -eq $valueProperty) { return $null }
    $value = $valueProperty.Value
    if ($null -eq $value -or [string]$value -eq '') { return $null }
    return ([string]$value).Trim()
}

function Convert-ToNullableNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}

function New-ValueCell($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return @{value=''} }
    return @{value=$Value}
}

function Invoke-LarkJson([scriptblock]$Command, [string]$Label, [int]$MaxAttempts = 3) {
    $errors = [Collections.Generic.List[string]]::new()
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $global:LASTEXITCODE = 0
        try {
            $raw = & $Command | Out-String
            $result = $raw | ConvertFrom-Json
            if ($LASTEXITCODE -eq 0 -and $result.ok -ne $false) { return $result }
            $errors.Add("第${attempt}次：退出码=$LASTEXITCODE")
        } catch {
            $errors.Add("第${attempt}次：$($_.Exception.Message)")
        }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds ([math]::Min(6, $attempt * 2)) }
    }
    throw "$Label；$($errors -join '；')"
}

function Set-CellMatrix([string]$Range, $Rows) {
    $payload = ConvertTo-Json -InputObject $Rows -Depth 12 -Compress
    $null = Invoke-LarkJson { $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range $Range --cells - --format json } "写入SKU成本主表失败：$Range"
}

$masterResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'A2:Z3000' --include value --max-chars 1200000 --format json } '读取SKU成本主表失败'
$inventoryResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name '库存快照' --range 'B2:O2000' --include value --max-chars 900000 --format json } '读取库存快照失败'
$masterRange = $masterResult.data.ranges[0]
$inventoryRange = $inventoryResult.data.ranges[0]
$masterRows = @($masterRange.cells)
$masterRowNumbers = @($masterRange.row_indices)
$inventoryRows = @($inventoryRange.cells)

$backupPath = Join-Path $stateRoot ("sku_cost_master_auto_backup_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
[ordered]@{
    source = 'SKU成本主表!A2:Z3000'
    captured_at = [datetime]::Now.ToString('o')
    row_indices = $masterRowNumbers
    cells = $masterRows
} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $backupPath -Encoding UTF8

$existingKw = @{}
$existingKwCandidates = @{}
$knownKwCodes = @{}
$blankKwByCsx = @{}
$lastUsedRow = 1
for ($i = 0; $i -lt $masterRows.Count; $i++) {
    $row = $masterRows[$i]
    $rowNumber = if ($i -lt $masterRowNumbers.Count) { [int]$masterRowNumbers[$i] } else { 2 + $i }
    $kw = Get-CellValue $row 0
    $csx = Get-CellValue $row 1
    $title = Get-CellValue $row 2
    if ($kw -or $csx -or $title) { $lastUsedRow = [math]::Max($lastUsedRow, $rowNumber) }
    if ($kw -and $kw -match '^(?i)kw\d+$') {
        $kwKey = $kw.ToLowerInvariant()
        $existingKw[$kwKey] = $rowNumber
        if (-not $existingKwCandidates.ContainsKey($kwKey)) { $existingKwCandidates[$kwKey] = [Collections.Generic.List[object]]::new() }
        $existingKwCandidates[$kwKey].Add([pscustomobject]@{row_number=$rowNumber;cells=$row})
    }
    if ($kw -and $kw -match '^(?i)kw\d+$') { $knownKwCodes[$kw.ToLowerInvariant()] = $true }
    if ($csx) {
        foreach ($alias in @([regex]::Split($csx, '[/,，、;；\s]+') | Where-Object { $_ })) {
            if ($alias -match '^(?i)kw\d+$') { $knownKwCodes[$alias.ToLowerInvariant()] = $true }
        }
    }
    if (-not $kw -and $csx -and $csx -match '^(?i)csx\d+$') {
        $key = $csx.ToLowerInvariant()
        if (-not $blankKwByCsx.ContainsKey($key)) { $blankKwByCsx[$key] = [Collections.Generic.List[object]]::new() }
        $blankKwByCsx[$key].Add([pscustomobject]@{ row_number=$rowNumber; title=$title })
    }
}

$newRows = [Collections.Generic.List[object]]::new()
$mappedRows = [Collections.Generic.List[object]]::new()
$enrichedRows = [Collections.Generic.List[object]]::new()
$seenInventoryKw = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$serial = ($SnapshotDate.Date - [datetime]'1899-12-30').TotalDays
foreach ($row in $inventoryRows) {
    $code = Get-CellValue $row 0
    if (-not $code -or $code -notmatch '^(?i)kw\d+$') { continue }
    $kw = $code.ToLowerInvariant()
    if (-not $seenInventoryKw.Add($kw)) { continue }
    $title = Get-CellValue $row 1
    $stock = Convert-ToNullableNumber (Get-CellValue $row 2)
    $shelf = Get-CellValue $row 5
    $csx = if ($shelf -and $shelf -match '^(?i)csx\d+$') { $shelf.ToLowerInvariant() } else { $null }
    $inventoryProductId = Get-CellValue $row 12
    $inventorySaleStatus = Get-CellValue $row 4
    $inventorySpec = Get-CellValue $row 13
    if ($existingKwCandidates.ContainsKey($kw)) {
        $existingCandidates = @($existingKwCandidates[$kw])
        if ($existingCandidates.Count -eq 1) {
            $target = $existingCandidates[0]
            $existingTitle = Get-CellValue $target.cells 2
            $existingCsx = Get-CellValue $target.cells 1
            $existingProductId = Get-CellValue $target.cells $skuSchema.ProductId.Index
            $existingSaleStatus = Get-CellValue $target.cells $skuSchema.PlatformStatus.Index
            $existingSpec = Get-CellValue $target.cells $skuSchema.CurrentSpec.Index
            $isPendingProfile = -not $existingProductId -or $existingSaleStatus -eq '待平台同步'
            if ($isPendingProfile) {
                if ($existingProductId -and $inventoryProductId -and $existingProductId -ne $inventoryProductId) {
                    continue
                }
                $resolvedCsx = if ($existingCsx) { $existingCsx } else { $csx }
                $resolvedTitle = if ($existingTitle) { $existingTitle } else { $title }
                $resolvedProductId = if ($existingProductId) { $existingProductId } else { $inventoryProductId }
                $resolvedStatus = if ($inventorySaleStatus) { $inventorySaleStatus } else { $existingSaleStatus }
                $resolvedSpec = if ($inventorySpec) { $inventorySpec } else { $existingSpec }
                Set-CellMatrix "B$($target.row_number):C$($target.row_number)" @(@((New-ValueCell $resolvedCsx),(New-ValueCell $resolvedTitle)))
                Set-CellMatrix "$($skuSchema.ProductId.Column)$($target.row_number):$($skuSchema.ProductId.Column)$($target.row_number)" @(@((New-ValueCell $resolvedProductId)))
                Set-CellMatrix "$($skuSchema.PlatformStatus.Column)$($target.row_number):$($skuSchema.PlatformStatus.Column)$($target.row_number)" @(@((New-ValueCell $resolvedStatus)))
                Set-CellMatrix "$($skuSchema.CurrentSpec.Column)$($target.row_number):$($skuSchema.CurrentSpec.Column)$($target.row_number)" @(@((New-ValueCell $resolvedSpec)))
                $enrichedRows.Add([pscustomobject]@{kw=$kw;row=[int]$target.row_number;product_id=$resolvedProductId})
            }
        }
        continue
    }
    if ($knownKwCodes.ContainsKey($kw)) { continue }
    $candidates = if ($csx -and $blankKwByCsx.ContainsKey($csx)) { @($blankKwByCsx[$csx]) } else { @() }
    if ($candidates.Count -eq 1) {
        $target = $candidates[0]
        Set-CellMatrix "A$($target.row_number):A$($target.row_number)" @(@(New-ValueCell $kw))
        $mappedRows.Add([pscustomobject]@{kw=$kw;csx=$csx;row=[int]$target.row_number})
        $existingKw[$kw] = [int]$target.row_number
        $knownKwCodes[$kw] = $true
        $blankKwByCsx.Remove($csx)
        continue
    }
    $status = if ($candidates.Count -gt 1) { '待人工核对' } else { '缺少成本' }
    $note = if ($candidates.Count -gt 1) { "新KW自动建档；$csx 对应多条过渡记录，请人工合并" } else { '新KW自动建档，请补充单位成本和采购包装换算' }
    $newRow = @(); for ($columnIndex = 0; $columnIndex -lt 26; $columnIndex++) { $newRow += ,(New-ValueCell $null) }
    $newRow[$skuSchema.CurrentKw.Index] = New-ValueCell $kw
    $newRow[$skuSchema.TransitionCode.Index] = New-ValueCell $csx
    $newRow[$skuSchema.PlatformTitle.Index] = New-ValueCell $title
    $newRow[$skuSchema.InitialInventory.Index] = New-ValueCell $stock
    $newRow[$skuSchema.CostBookQuantity.Index] = New-ValueCell $stock
    $newRow[$skuSchema.UnitsPerPack.Index] = New-ValueCell 1
    $newRow[$skuSchema.CostStatus.Index] = New-ValueCell $status
    $newRow[$skuSchema.InitializationDate.Index] = New-ValueCell $serial
    $newRow[$skuSchema.ConversionBasis.Index] = New-ValueCell '美团库存同步自动建档；成本与包装换算待核对'
    $newRow[$skuSchema.Note.Index] = New-ValueCell $note
    $newRow[$skuSchema.ProductId.Index] = New-ValueCell $inventoryProductId
    $newRow[$skuSchema.PlatformStatus.Index] = New-ValueCell $inventorySaleStatus
    $newRow[$skuSchema.CurrentSpec.Index] = New-ValueCell $inventorySpec
    $newRows.Add($newRow)
    $existingKw[$kw] = -1
    $knownKwCodes[$kw] = $true
}

$appendStart = $lastUsedRow + 1
if ($newRows.Count -gt 0) {
    $newKwCodes = @($newRows | ForEach-Object { ([string]$_[0].value).Trim().ToLowerInvariant() })
    $duplicateNewKw = @($newKwCodes | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    $alreadyExistingKw = @($newKwCodes | Where-Object { $existingKwCandidates.ContainsKey($_) } | Sort-Object -Unique)
    if ($duplicateNewKw.Count -gt 0 -or $alreadyExistingKw.Count -gt 0) {
        $blocked = @($duplicateNewKw + $alreadyExistingKw | Sort-Object -Unique) -join '、'
        throw "SKU自动建档检测到重复KW，已在写入前停止：$blocked"
    }
    $appendEnd = $appendStart + $newRows.Count - 1
    if ($appendEnd -gt 3000) { throw 'SKU成本主表超过3000行容量' }
    Set-CellMatrix "A${appendStart}:Z${appendEnd}" $newRows.ToArray()
}

$affectedCodes = @($mappedRows | ForEach-Object { $_.kw }) + @($newRows | ForEach-Object { [string]$_[0].value }) + @($enrichedRows | ForEach-Object { $_.kw })
if ($affectedCodes.Count -gt 0) {
    $readEnd = $lastUsedRow + $newRows.Count
    if ($enrichedRows.Count -gt 0) { $readEnd = [math]::Max($readEnd, [int](($enrichedRows | Measure-Object -Property row -Maximum).Maximum)) }
    $verifyResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range "A2:Z$readEnd" --include value --max-chars 1200000 --format json } '回读SKU成本主表失败'
    $readCodes = @($verifyResult.data.ranges[0].cells | ForEach-Object { Get-CellValue $_ 0 } | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() })
    foreach ($code in $affectedCodes) {
        if ($code -notin $readCodes) { throw "SKU自动建档回读失败：$code" }
    }
}

[ordered]@{
    ok = $true
    snapshot_date = $SnapshotDate.ToString('yyyy-MM-dd')
    new_kw_count = $newRows.Count
    mapped_existing_count = $mappedRows.Count
    enriched_pending_count = $enrichedRows.Count
    new_kw_codes = @($newRows | ForEach-Object { [string]$_[0].value })
    mapped_kw_codes = @($mappedRows | ForEach-Object { $_.kw })
    enriched_kw_codes = @($enrichedRows | ForEach-Object { $_.kw })
    duplicate_kw_groups_before = @($existingKwCandidates.GetEnumerator() | Where-Object { @($_.Value).Count -gt 1 }).Count
    formula_verification = 'success'
} | ConvertTo-Json -Depth 6 -Compress
