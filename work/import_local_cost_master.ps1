param(
    [string]$ProjectRoot,
    [string]$SourcePath = 'D:\market-research\宠速星\对货表.xlsx'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
if (-not (Test-Path -LiteralPath $SourcePath)) { throw "找不到本地成本源：$SourcePath" }

$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少经营驾驶舱元数据' }
$meta = [System.IO.File]::ReadAllText($metaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$stateRoot = Join-Path $ProjectRoot 'agent\state'
if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }

function Get-ZipEntryText([System.IO.Compression.ZipArchive]$Zip, [string]$Name) {
    $entry = $Zip.GetEntry($Name)
    if (-not $entry) { throw "XLSX 缺少内部文件：$Name" }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Convert-ColumnNameToIndex([string]$Name) {
    $result = 0
    foreach ($character in $Name.ToCharArray()) { $result = $result * 26 + ([int][char]$character - [int][char]'A' + 1) }
    return $result - 1
}
function Get-XlsxCellValue($Cell, [string[]]$SharedStrings) {
    $type = [string]$Cell.t
    $raw = [string]$Cell.v
    if ($type -eq 's' -and $raw -match '^\d+$') { return $SharedStrings[[int]$raw] }
    if ($type -eq 'inlineStr') {
        return (($Cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.'#text' }) -join '')
    }
    if ($null -eq $raw -or $raw -eq '') { return $null }
    if ($raw -match '^-?\d+(\.\d+)?$') { return [double]$raw }
    return $raw
}
function Read-XlsxRows([string]$Path, [string]$PreferredSheet) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
    try {
        $workbook = [xml](Get-ZipEntryText $zip 'xl/workbook.xml')
        $relationships = [xml](Get-ZipEntryText $zip 'xl/_rels/workbook.xml.rels')
        $sharedStrings = @()
        if ($zip.GetEntry('xl/sharedStrings.xml')) {
            $sharedXml = [xml](Get-ZipEntryText $zip 'xl/sharedStrings.xml')
            $sharedStrings = @($sharedXml.SelectNodes("//*[local-name()='si']") | ForEach-Object {
                (($_.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.'#text' }) -join '')
            })
        }
        $sheetNode = @($workbook.SelectNodes("//*[local-name()='sheet']") | Where-Object { ([string]$_.name) -eq $PreferredSheet } | Select-Object -First 1)
        if (-not $sheetNode) { $sheetNode = @($workbook.SelectNodes("//*[local-name()='sheet']") | Select-Object -First 1) }
        $relationshipId = [string]$sheetNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        if (-not $relationshipId) { $relationshipId = [string]$sheetNode.'r:id' }
        $relationship = @($relationships.SelectNodes("//*[local-name()='Relationship']") | Where-Object { ([string]$_.Id) -eq $relationshipId } | Select-Object -First 1)
        $target = ([string]$relationship.Target).TrimStart('/')
        if ($target -notlike 'xl/*') { $target = "xl/$target" }
        $sheetXml = [xml](Get-ZipEntryText $zip $target)
        $rows = @()
        foreach ($rowNode in $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $cells = @{}
            foreach ($cell in $rowNode.SelectNodes("./*[local-name()='c']")) {
                $columnName = [regex]::Match([string]$cell.r, '^[A-Z]+').Value
                if ($columnName) { $cells[(Convert-ColumnNameToIndex $columnName)] = Get-XlsxCellValue $cell $sharedStrings }
            }
            $maximumIndex = if ($cells.Count) { ($cells.Keys | Measure-Object -Maximum).Maximum } else { -1 }
            if ($maximumIndex -ge 0) {
                $row = [object[]]::new($maximumIndex + 1)
                foreach ($index in $cells.Keys) { $row[$index] = $cells[$index] }
                $rows += ,$row
            }
        }
        return ,$rows
    } finally {
        $zip.Dispose()
        $stream.Dispose()
    }
}
function Get-At([object[]]$Row, [int]$Index) {
    if ($null -ne $Row -and $Index -ge 0 -and $Index -lt $Row.Count) { return $Row[$Index] }
    return $null
}
function Normalize-Code($Value) {
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}
function Split-Codes($Value) {
    $text = Normalize-Code $Value
    if (-not $text) { return @() }
    return @([regex]::Split($text, '[/,，、;；\s]+') | Where-Object { $_ } | Select-Object -Unique)
}
function Convert-ToNullableNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}
function Get-LarkCellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}
function New-OverwriteCell($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return @{value=''} }
    return @{value=$Value}
}
function Get-SheetRows([string]$SheetName, [string]$Range) {
    $raw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --include value,formula --max-chars 5000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取飞书失败：$SheetName" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "飞书读取结果不完整：$SheetName" }
    return @($result.data.ranges[0].cells)
}
function Set-CellMatrix([string]$SheetName, [int]$StartRow, $Rows) {
    $batchSize = 80
    for ($offset = 0; $offset -lt @($Rows).Count; $offset += $batchSize) {
        $count = [math]::Min($batchSize, @($Rows).Count - $offset)
        $batch = @($Rows[$offset..($offset + $count - 1)])
        $firstRow = $StartRow + $offset
        $lastRow = $firstRow + $count - 1
        $range = "A${firstRow}:N${lastRow}"
        $payload = ConvertTo-Json -InputObject $batch -Depth 12 -Compress
        $raw = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $SheetName --range $range --cells - --format json | Out-String
        if ($LASTEXITCODE -ne 0 -or -not (($raw | ConvertFrom-Json).ok)) { throw "写入飞书失败：$SheetName $range" }
    }
}

$masterRows = Get-SheetRows 'SKU成本主表' 'A1:N1500'
$referenceRows = Get-SheetRows '商品成本' 'A2:E2000'
$backupPath = Join-Path $stateRoot ("sku_cost_master_backup_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
$backupObject = [ordered]@{ created_at=[datetime]::Now.ToString('s'); source='SKU成本主表!A1:N1500'; cells=$masterRows }
[System.IO.File]::WriteAllText($backupPath, ($backupObject | ConvertTo-Json -Depth 20), [System.Text.UTF8Encoding]::new($false))

$referenceByCode = @{}
foreach ($row in $referenceRows) {
    $aliases = Split-Codes (Get-LarkCellValue $row 0)
    if ($aliases.Count -eq 0) { continue }
    $packCost = Convert-ToNullableNumber (Get-LarkCellValue $row 1)
    $splitQuantity = Convert-ToNullableNumber (Get-LarkCellValue $row 2)
    $unitCost = Convert-ToNullableNumber (Get-LarkCellValue $row 3)
    $status = [string](Get-LarkCellValue $row 4)
    if ($null -eq $splitQuantity -or $splitQuantity -le 0) { $splitQuantity = 1.0 }
    if ($null -eq $unitCost -and $null -ne $packCost) { $unitCost = [double]$packCost / [double]$splitQuantity }
    foreach ($alias in $aliases) {
        if (-not $referenceByCode.ContainsKey($alias)) {
            $referenceByCode[$alias] = [pscustomobject]@{ pack_cost=$packCost; split_quantity=$splitQuantity; unit_cost=$unitCost; status=$status }
        }
    }
}

$existingAliases = @{}
$lastUsedRow = 1
for ($index = 1; $index -lt $masterRows.Count; $index++) {
    $row = $masterRows[$index]
    $hasData = $false
    foreach ($column in 0..13) { if (Get-LarkCellValue $row $column) { $hasData = $true; break } }
    if ($hasData) { $lastUsedRow = $index + 1 }
    foreach ($alias in (Split-Codes (Get-LarkCellValue $row 1))) { $existingAliases[$alias] = $true }
}

$sourceRows = Read-XlsxRows $SourcePath '货品表'
$headerRowIndex = -1
$headers = @{}
for ($rowIndex=0; $rowIndex -lt [math]::Min(20,$sourceRows.Count); $rowIndex++) {
    $candidate = $sourceRows[$rowIndex]
    for ($columnIndex=0; $columnIndex -lt $candidate.Count; $columnIndex++) {
        $label = ([string](Get-At $candidate $columnIndex)).Trim()
        if ($label) { $headers[$label] = $columnIndex }
    }
    if ($headers.ContainsKey('货号') -and $headers.ContainsKey('参考进货价')) { $headerRowIndex = $rowIndex; break }
    $headers = @{}
}
if ($headerRowIndex -lt 0) { throw '本地对货表未找到“货号”和“参考进货价”表头' }

$newRows = @()
$sourceRecordCount = 0
$sourceAliasSet = @{}
$missingCostCount = 0
$splitCostCount = 0
for ($rowIndex=$headerRowIndex+1; $rowIndex -lt $sourceRows.Count; $rowIndex++) {
    $sourceRow = $sourceRows[$rowIndex]
    $sourceCode = Normalize-Code (Get-At $sourceRow $headers['货号'])
    $sourceAliases = Split-Codes $sourceCode
    if ($sourceAliases.Count -eq 0) { continue }
    $sourceRecordCount++
    foreach ($alias in $sourceAliases) { $sourceAliasSet[$alias] = $true }
    $missingAliases = @($sourceAliases | Where-Object { -not $existingAliases.ContainsKey($_) })
    if ($missingAliases.Count -eq 0) { continue }

    $title = if ($headers.ContainsKey('商品名称')) { [string](Get-At $sourceRow $headers['商品名称']) } else { '' }
    $specification = if ($headers.ContainsKey('商品规格')) { [string](Get-At $sourceRow $headers['商品规格']) } else { '' }
    $localPackCost = Convert-ToNullableNumber (Get-At $sourceRow $headers['参考进货价'])
    $referenceCandidates = @($missingAliases | Where-Object { $referenceByCode.ContainsKey($_) } | ForEach-Object { $referenceByCode[$_] })
    $unitCosts = @($referenceCandidates | Where-Object { $null -ne $_.unit_cost } | ForEach-Object { [double]$_.unit_cost } | Sort-Object -Unique)
    $splitQuantities = @($referenceCandidates | Where-Object { $null -ne $_.split_quantity } | ForEach-Object { [double]$_.split_quantity } | Sort-Object -Unique)
    $unitCost = if ($unitCosts.Count -eq 1) { $unitCosts[0] } elseif ($unitCosts.Count -eq 0) { $localPackCost } else { $null }
    $splitQuantity = if ($splitQuantities.Count -eq 1) { $splitQuantities[0] } else { 1.0 }
    $hasPendingReference = @($referenceCandidates | Where-Object { $_.status -in @('待人工核对','缺少成本') }).Count -gt 0
    if ($null -eq $unitCost) { $missingCostCount++ }
    if ($splitQuantity -gt 1) { $splitCostCount++ }

    $status = if ($null -eq $unitCost) {
        '缺少参考成本'
    } elseif ($hasPendingReference) {
        '包装换算待核对'
    } elseif ($splitQuantity -gt 1) {
        '拆分成本'
    } else {
        '待补KW'
    }
    $basis = if ($splitQuantity -gt 1) { "参考进货价÷$splitQuantity 个销售单位" } else { '参考进货价按单个销售单位' }
    $note = if ($hasPendingReference) { '待补KW；包装换算待人工确认；来源：本地对货表' } else { '待补KW；来源：本地对货表' }
    $nowSerial = ([datetime]::Now - [datetime]'1899-12-30').TotalDays
    $newRows += ,@(
        (New-OverwriteCell $null),
        (New-OverwriteCell ($missingAliases -join '、')),
        (New-OverwriteCell $title),
        (New-OverwriteCell 0),
        (New-OverwriteCell $unitCost),
        (New-OverwriteCell $unitCost),
        (New-OverwriteCell 0),
        (New-OverwriteCell $splitQuantity),
        (New-OverwriteCell $specification),
        (New-OverwriteCell $status),
        (New-OverwriteCell $nowSerial),
        (New-OverwriteCell $nowSerial),
        (New-OverwriteCell $basis),
        (New-OverwriteCell $note)
    )
    foreach ($alias in $missingAliases) { $existingAliases[$alias] = $true }
}

$startRow = $lastUsedRow + 1
$lastWriteRow = $startRow + $newRows.Count - 1
if ($lastWriteRow -gt 1500) { throw "SKU成本主表容量不足：需要写到第 $lastWriteRow 行" }
if ($newRows.Count -gt 0) { Set-CellMatrix 'SKU成本主表' $startRow $newRows }

$checkRows = Get-SheetRows 'SKU成本主表' 'A2:N1500'
$checkAliases = @{}
$blankKwRows = 0
$usedRows = 0
foreach ($row in $checkRows) {
    $csx = Get-LarkCellValue $row 1
    if (-not $csx) { continue }
    $usedRows++
    if (-not (Get-LarkCellValue $row 0)) { $blankKwRows++ }
    foreach ($alias in (Split-Codes $csx)) { $checkAliases[$alias] = $true }
}
$unrepresented = @($sourceAliasSet.Keys | Where-Object { -not $checkAliases.ContainsKey($_) } | Sort-Object)
if ($unrepresented.Count -gt 0) { throw "导入后仍有 $($unrepresented.Count) 个对货表货号未进入 SKU成本主表" }

[ordered]@{
    ok=$true
    source_records=$sourceRecordCount
    source_aliases=$sourceAliasSet.Count
    existing_last_row=$lastUsedRow
    appended_rows=$newRows.Count
    final_cost_rows=$usedRows
    blank_kw_rows=$blankKwRows
    split_cost_rows_added=$splitCostCount
    missing_cost_rows_added=$missingCostCount
    unrepresented_source_aliases=$unrepresented.Count
    backup_file=$backupPath
} | ConvertTo-Json -Compress
