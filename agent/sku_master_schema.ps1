$script:SkuMasterRequiredHeaders = [ordered]@{
    CurrentKw          = '当前KW'
    TransitionCode     = 'CSX/过渡货号'
    PlatformTitle      = '商品标题'
    CustomTitle        = '自定义标题'
    InitialInventory   = '初始化库存'
    BaselineUnitCost   = '基准单位成本'
    MovingAverageCost  = '当前移动平均成本'
    CostBookQuantity   = '成本账数量'
    UnitsPerPack       = '每采购包装含销售单位数'
    PurchaseUnit       = '采购单位'
    CostStatus         = '成本核对状态'
    InitializationDate = '初始化日期'
    LastCostUpdate     = '最近成本更新时间'
    ConversionBasis    = '换算依据'
    Note               = '备注'
    EarliestExpiry     = '当前库存最早到期日期'
    ProductId          = '商品ID（美团SKU ID）'
    PlatformStatus     = '平台销售状态'
    CurrentSpec        = '当前规格'
}

function Convert-ColumnNumberToLetter([int]$Number) {
    if ($Number -lt 1) { throw '列号必须大于0' }
    $letters = ''
    while ($Number -gt 0) {
        $Number--
        $letters = [char](65 + ($Number % 26)) + $letters
        $Number = [math]::Floor($Number / 26)
    }
    return $letters
}

function Get-SkuMasterSchema([string]$SpreadsheetToken) {
    if (-not $SpreadsheetToken) { throw '缺少飞书电子表格标识' }
    $raw = & $script:SkuSchemaLarkCli sheets +cells-get --as user --spreadsheet-token $SpreadsheetToken --sheet-name 'SKU成本主表' --range 'A1:Z1' --include value --max-chars 20000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '读取SKU成本主表表头失败' }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw 'SKU成本主表表头读取不完整' }
    $returnedCells = @($result.data.ranges[0].cells)
    if ($returnedCells.Count -eq 0) { throw 'SKU成本主表表头为空' }
    # 单行范围在不同 CLI/JSON 包装中可能返回“行数组”或直接返回“单元格数组”。
    $firstValueProperty = $returnedCells[0].PSObject.Properties['value']
    $headerRow = if ($null -ne $firstValueProperty) { $returnedCells } else { $returnedCells[0] }
    $byHeader = @{}
    for ($index = 0; $index -lt @($headerRow).Count; $index++) {
        $cell = @($headerRow)[$index]
        $valueProperty = if ($null -ne $cell) { $cell.PSObject.Properties['value'] } else { $null }
        $header = if ($null -ne $valueProperty -and $null -ne $valueProperty.Value) { ([string]$valueProperty.Value).Trim() } else { '' }
        if (-not $header) { continue }
        if ($byHeader.ContainsKey($header)) { throw "SKU成本主表存在重复表头：$header" }
        $byHeader[$header] = $index
    }

    $schema = [ordered]@{}
    foreach ($entry in $script:SkuMasterRequiredHeaders.GetEnumerator()) {
        if (-not $byHeader.ContainsKey($entry.Value)) { throw "SKU成本主表缺少必要表头：$($entry.Value)" }
        $index = [int]$byHeader[$entry.Value]
        $schema[$entry.Key] = [pscustomobject]@{
            Header = $entry.Value
            Index = $index
            Column = Convert-ColumnNumberToLetter ($index + 1)
        }
    }
    return [pscustomobject]$schema
}
$script:SkuSchemaLarkCli = Join-Path $PSScriptRoot 'lark_cli_noexit.ps1'
