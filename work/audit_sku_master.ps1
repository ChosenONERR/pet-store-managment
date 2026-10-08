param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
$larkCli = Join-Path $ProjectRoot 'agent\lark_cli_noexit.ps1'
$meta = Get-Content -LiteralPath $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri][string]$meta.workbook_url).Segments[-1].Trim('/')

function Get-CellValue($Cell) {
    if ($null -eq $Cell -or $null -eq $Cell.value) { return $null }
    return $Cell.value
}

function Convert-DateValue($Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $number = 0.0
    if ([double]::TryParse([string]$Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number) -and $number -gt 30000 -and $number -lt 90000) {
        return [datetime]::FromOADate($number).ToString('yyyy-MM-dd')
    }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value,[ref]$date)) { return $date.ToString('yyyy-MM-dd') }
    return [string]$Value
}

function Convert-NullableNumber($Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) { return $number }
    return $null
}

function Get-SheetTable([string]$SheetName,[string]$Range) {
    $global:LASTEXITCODE = 0
    $raw = & $larkCli sheets +cells-get --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --include 'value,formula' --max-chars 10000000 --format json 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "read failed: $SheetName" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "incomplete read: $SheetName" }
    $rangeResult = @($result.data.ranges)[0]
    if ($rangeResult.truncated) { throw "truncated read: $SheetName" }
    $matrix = @($rangeResult.cells)
    if ($matrix.Count -eq 0) { return @() }
    $headers = @($matrix[0] | ForEach-Object { [string](Get-CellValue $_) })
    $rowIndices = @($rangeResult.row_indices)
    $records = [Collections.Generic.List[object]]::new()
    for ($rowOffset=1; $rowOffset -lt $matrix.Count; $rowOffset++) {
        $cells = @($matrix[$rowOffset])
        $record = [ordered]@{_row=if($rowIndices.Count -gt $rowOffset){[int]$rowIndices[$rowOffset]}else{$rowOffset+1}}
        $hasValue = $false
        for ($column=0; $column -lt $headers.Count; $column++) {
            $header = $headers[$column].Trim()
            if (-not $header) { continue }
            $cell = if ($column -lt $cells.Count) { $cells[$column] } else { $null }
            $value = Get-CellValue $cell
            if ($header -match '日期|时间' -and $null -ne $value) { $value = Convert-DateValue $value }
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { $hasValue = $true }
            $record[$header] = $value
            if ($cell -and $cell.formula) { $record["__formula_$header"] = [string]$cell.formula }
        }
        if ($hasValue) { $records.Add([pscustomobject]$record) }
    }
    return @($records)
}

function Text($Value) { return ([string]$Value).Trim() }
function Lower($Value) { return (Text $Value).ToLowerInvariant() }
function Normalize-Title($Value) { return ([regex]::Replace((Lower $Value),'[\s\p{P}\p{S}]+','')) }
function Project-Row($Row) {
    return [ordered]@{
        row=$Row._row
        kw=Text $Row.'当前KW'
        transition_code=Text $Row.'CSX/过渡货号'
        product_id=Text $Row.'商品ID（美团SKU ID）'
        title=Text $Row.'商品标题'
        custom_title=Text $Row.'自定义标题'
        spec=Text $Row.'当前规格'
        baseline_cost=Convert-NullableNumber $Row.'基准单位成本'
        moving_cost=Convert-NullableNumber $Row.'当前移动平均成本'
        inventory=Convert-NullableNumber $Row.'初始化库存'
        cost_quantity=Convert-NullableNumber $Row.'成本账数量'
        units_per_pack=Convert-NullableNumber $Row.'每采购包装含销售单位数'
        cost_status=Text $Row.'成本核对状态'
        platform_status=Text $Row.'平台销售状态'
        expiry=Text $Row.'当前库存最早到期日期'
        note=Text $Row.'备注'
    }
}
function Group-Issues($Rows,[scriptblock]$KeySelector,[int]$Minimum=2) {
    return @($Rows | Group-Object -Property $KeySelector | Where-Object { $_.Name -and $_.Count -ge $Minimum } | ForEach-Object {
        [ordered]@{key=$_.Name;count=$_.Count;rows=@($_.Group | ForEach-Object { Project-Row $_ })}
    })
}

$skuRows = @(Get-SheetTable 'SKU成本主表' 'A1:Z3000')
$inventoryRows = @(Get-SheetTable '库存快照' 'A1:O2000')
$orderRows = @(Get-SheetTable '订单商品' 'A1:Q4001')

$kwGroups = @{}
foreach ($row in $skuRows) {
    $kw = Lower $row.'当前KW'
    if (-not $kw) { continue }
    if (-not $kwGroups.ContainsKey($kw)) { $kwGroups[$kw] = [Collections.Generic.List[object]]::new() }
    $kwGroups[$kw].Add($row)
}
$productGroups = @{}
foreach ($row in $skuRows) {
    $id = Text $row.'商品ID（美团SKU ID）'
    if (-not $id) { continue }
    if (-not $productGroups.ContainsKey($id)) { $productGroups[$id] = [Collections.Generic.List[object]]::new() }
    $productGroups[$id].Add($row)
}
$transitionGroups = @{}
foreach ($row in $skuRows) {
    $tokens = @([regex]::Split((Lower $row.'CSX/过渡货号'),'[/,，、;；\s]+') | Where-Object { $_ })
    foreach ($code in $tokens) {
        if (-not $transitionGroups.ContainsKey($code)) { $transitionGroups[$code] = [Collections.Generic.List[object]]::new() }
        $transitionGroups[$code].Add($row)
    }
}

$duplicateKw = @($kwGroups.GetEnumerator() | Where-Object {$_.Value.Count -gt 1} | Sort-Object Name | ForEach-Object {[ordered]@{key=$_.Key;count=$_.Value.Count;rows=@($_.Value|ForEach-Object{Project-Row $_})}})
$duplicateProductId = @($productGroups.GetEnumerator() | Where-Object {$_.Value.Count -gt 1} | Sort-Object Name | ForEach-Object {[ordered]@{key=$_.Key;count=$_.Value.Count;rows=@($_.Value|ForEach-Object{Project-Row $_})}})
$duplicateTransition = @($transitionGroups.GetEnumerator() | Where-Object {$_.Value.Count -gt 1} | Sort-Object Name | ForEach-Object {[ordered]@{key=$_.Key;count=$_.Value.Count;rows=@($_.Value|ForEach-Object{Project-Row $_})}})
$duplicateTitleSpec = Group-Issues $skuRows { "$(Normalize-Title $_.'商品标题')|$(Normalize-Title $_.'当前规格')" }
$exactDuplicates = Group-Issues $skuRows { "$(Lower $_.'当前KW')|$(Lower $_.'CSX/过渡货号')|$(Text $_.'商品ID（美团SKU ID）')|$(Normalize-Title $_.'商品标题')|$(Normalize-Title $_.'当前规格')" }

$blankKw = @($skuRows | Where-Object { -not (Text $_.'当前KW') } | ForEach-Object { Project-Row $_ })
$invalidKw = @($skuRows | Where-Object { (Text $_.'当前KW') -and (Lower $_.'当前KW') -notmatch '^kw\d+$' } | ForEach-Object { Project-Row $_ })
$blankProductId = @($skuRows | Where-Object { -not (Text $_.'商品ID（美团SKU ID）') } | ForEach-Object { Project-Row $_ })
$invalidProductId = @($skuRows | Where-Object { (Text $_.'商品ID（美团SKU ID）') -and (Text $_.'商品ID（美团SKU ID）') -notmatch '^\d{8,20}$' } | ForEach-Object { Project-Row $_ })
$blankTitle = @($skuRows | Where-Object { -not (Text $_.'商品标题') -and -not (Text $_.'自定义标题') } | ForEach-Object { Project-Row $_ })
$missingBaselineCost = @($skuRows | Where-Object { $null -eq (Convert-NullableNumber $_.'基准单位成本') } | ForEach-Object { Project-Row $_ })
$missingMovingCost = @($skuRows | Where-Object { $null -eq (Convert-NullableNumber $_.'当前移动平均成本') } | ForEach-Object { Project-Row $_ })
$negativeCost = @($skuRows | Where-Object {
    $baseline=Convert-NullableNumber $_.'基准单位成本'; $moving=Convert-NullableNumber $_.'当前移动平均成本'
    ($null -ne $baseline -and $baseline -lt 0) -or ($null -ne $moving -and $moving -lt 0)
} | ForEach-Object { Project-Row $_ })
$invalidUnitsPerPack = @($skuRows | Where-Object { $null -eq (Convert-NullableNumber $_.'每采购包装含销售单位数') -or (Convert-NullableNumber $_.'每采购包装含销售单位数') -le 0 } | ForEach-Object { Project-Row $_ })
$negativeCostQuantity = @($skuRows | Where-Object { (Convert-NullableNumber $_.'成本账数量') -lt 0 } | ForEach-Object { Project-Row $_ })
$staleMissingStatus = @($skuRows | Where-Object {
    (Text $_.'成本核对状态') -eq '缺少成本' -and $null -ne (Convert-NullableNumber $_.'基准单位成本') -and $null -ne (Convert-NullableNumber $_.'当前移动平均成本')
} | ForEach-Object { Project-Row $_ })
$staleConflictStatus = @($skuRows | Where-Object {
    (Text $_.'成本核对状态') -match '冲突' -and $kwGroups[(Lower $_.'当前KW')].Count -le 1
} | ForEach-Object { Project-Row $_ })
$unmarkedKwConflicts = @($skuRows | Where-Object {
    $kw=(Lower $_.'当前KW'); $kw -and $kwGroups[$kw].Count -gt 1 -and (Text $_.'成本核对状态') -notmatch '冲突'
} | ForEach-Object { Project-Row $_ })
$zeroCostWithStock = @($skuRows | Where-Object {
    $cost=Convert-NullableNumber $_.'当前移动平均成本'; $qty=Convert-NullableNumber $_.'成本账数量';
    $cost -eq 0 -and $qty -gt 0 -and "$(Text $_.'商品标题') $(Text $_.'自定义标题') $(Text $_.'备注')" -notmatch '试吃|尝鲜|赠送|赠品|成本为0|供货商送'
} | ForEach-Object { Project-Row $_ })

$activeRows = @($skuRows | Where-Object { (Text $_.'平台销售状态') -match '售卖' })
$activeBlankKw = @($activeRows | Where-Object { -not (Text $_.'当前KW') } | ForEach-Object { Project-Row $_ })
$activeMissingCost = @($activeRows | Where-Object { $null -eq (Convert-NullableNumber $_.'基准单位成本') -or $null -eq (Convert-NullableNumber $_.'当前移动平均成本') } | ForEach-Object { Project-Row $_ })
$activeMissingCostWithStock = @($activeRows | Where-Object {
    ($null -eq (Convert-NullableNumber $_.'基准单位成本') -or $null -eq (Convert-NullableNumber $_.'当前移动平均成本')) -and (Convert-NullableNumber $_.'成本账数量') -gt 0
} | ForEach-Object { Project-Row $_ })
$activeDuplicateKwGroups = @($duplicateKw | Where-Object { @($_.rows | Where-Object { $_.platform_status -match '售卖' }).Count -gt 1 })
$singleActiveDuplicateKwGroups = @($duplicateKw | Where-Object { @($_.rows | Where-Object { $_.platform_status -match '售卖' }).Count -le 1 })
$expiryPresent = @($skuRows | Where-Object { Text $_.'当前库存最早到期日期' })
$expiryCutoff = [datetime]::Today.AddMonths(8)
$expiredRows = @($skuRows | Where-Object { $text=Text $_.'当前库存最早到期日期'; if(-not $text){return $false};$d=[datetime]::MinValue;[datetime]::TryParse($text,[ref]$d) -and $d.Date -lt [datetime]::Today } | ForEach-Object { Project-Row $_ })
$expiringRows = @($skuRows | Where-Object { $text=Text $_.'当前库存最早到期日期'; if(-not $text){return $false};$d=[datetime]::MinValue;[datetime]::TryParse($text,[ref]$d) -and $d.Date -ge [datetime]::Today -and $d.Date -le $expiryCutoff.Date } | ForEach-Object { Project-Row $_ })
$foodPattern = '粮|罐|零食|冻干|猫条|狗条|奶粉|肉泥|湿粮|鲜肉|鸡肉|鸭肉|牛肉|鱼油|营养膏|益生菌|咬胶|洁齿骨|火腿|饼干'
$activeFoodStockMissingExpiry = @($activeRows | Where-Object {
    (Convert-NullableNumber $_.'成本账数量') -gt 0 -and -not (Text $_.'当前库存最早到期日期') -and "$(Text $_.'商品标题') $(Text $_.'自定义标题')" -match $foodPattern
} | ForEach-Object { Project-Row $_ })

$invalidDates = [Collections.Generic.List[object]]::new()
foreach ($row in $skuRows) {
    foreach ($column in @('初始化日期','最近成本更新时间','当前库存最早到期日期')) {
        $value = Text $row.$column
        if (-not $value) { continue }
        $parsed = [datetime]::MinValue
        if (-not [datetime]::TryParse($value,[ref]$parsed)) { $invalidDates.Add([ordered]@{row=$row._row;column=$column;value=$value;kw=Text $row.'当前KW';product_id=Text $row.'商品ID（美团SKU ID）'}) }
    }
}

$inventoryById = @{}
foreach ($row in $inventoryRows) {
    $id=Text $row.'美团SKU ID'; if(-not $id){continue}
    if(-not $inventoryById.ContainsKey($id)){$inventoryById[$id]=[Collections.Generic.List[object]]::new()}
    $inventoryById[$id].Add($row)
}
$inventoryDuplicateIds = @($inventoryById.GetEnumerator() | Where-Object {$_.Value.Count -gt 1} | ForEach-Object {[ordered]@{product_id=$_.Key;count=$_.Value.Count;rows=@($_.Value|ForEach-Object{[ordered]@{row=$_._row;kw=Text $_.'货号';title=Text $_.'商品名称';spec=Text $_.'规格名称'}})}})
$activeMasterMissingInventory = @($skuRows | Where-Object {
    (Text $_.'平台销售状态') -match '售卖' -and (Text $_.'商品ID（美团SKU ID）') -and -not $inventoryById.ContainsKey((Text $_.'商品ID（美团SKU ID）'))
} | ForEach-Object { Project-Row $_ })
$inventoryMissingMaster = @($inventoryRows | Where-Object { $id=Text $_.'美团SKU ID'; $id -and -not $productGroups.ContainsKey($id) } | ForEach-Object {[ordered]@{row=$_._row;product_id=Text $_.'美团SKU ID';kw=Text $_.'货号';shelf_code=Text $_.'货架码';title=Text $_.'商品名称';spec=Text $_.'规格名称';status=Text $_.'销售状态';stock=Convert-NullableNumber $_.'库存'}})
$kwInventoryMismatch = @($skuRows | Where-Object {
    $id=Text $_.'商品ID（美团SKU ID）'; if(-not $id -or -not $inventoryById.ContainsKey($id)){return $false}
    $inventoryKw=Lower $inventoryById[$id][0].'货号'; $masterKw=Lower $_.'当前KW';
    $inventoryKw -and $masterKw -and $inventoryKw -ne $masterKw
} | ForEach-Object {[ordered]@{row=$_._row;product_id=Text $_.'商品ID（美团SKU ID）';master_kw=Text $_.'当前KW';inventory_kw=Text $inventoryById[(Text $_.'商品ID（美团SKU ID）')][0].'货号';title=Text $_.'商品标题';spec=Text $_.'当前规格'}})

$masterProductIds = [Collections.Generic.HashSet[string]]::new([string[]]@($productGroups.Keys))
$orderProductIds = @($orderRows | Where-Object { (Text $_.'订单状态') -notmatch '取消' } | ForEach-Object { Text $_.'商品ID（美团SKU ID）' } | Where-Object { $_ } | Select-Object -Unique)
$orderProductMissingMaster = @($orderProductIds | Where-Object { -not $masterProductIds.Contains($_) } | ForEach-Object { $id=$_; $sample=@($orderRows|Where-Object{(Text $_.'商品ID（美团SKU ID）') -eq $id}|Select-Object -First 1)[0]; [ordered]@{product_id=$id;code=Text $sample.'货号';title=Text $sample.'商品标题';spec=Text $sample.'规格';order_no=Text $sample.'平台订单号'} })

$formulaErrors = @($skuRows | ForEach-Object {
    $row=$_
    foreach($property in $row.PSObject.Properties){
        if($property.Name -like '__formula_*'){
            $header=$property.Name.Substring(10); $value=Text $row.$header
            if($value -match '^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|SPILL!|CALC!)$'){[ordered]@{row=$row._row;column=$header;value=$value;formula=$property.Value}}
        }
    }
})

$statusSummary = @($skuRows | Group-Object -Property '成本核对状态' | Sort-Object Count -Descending | ForEach-Object {[ordered]@{status=if($_.Name){$_.Name}else{'(blank)'};count=$_.Count}})
$platformStatusSummary = @($skuRows | Group-Object -Property '平台销售状态' | Sort-Object Count -Descending | ForEach-Object {[ordered]@{status=if($_.Name){$_.Name}else{'(blank)'};count=$_.Count}})

[ordered]@{
    ok=$true
    audited_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    counts=[ordered]@{sku_rows=$skuRows.Count;inventory_rows=$inventoryRows.Count;order_item_rows=$orderRows.Count;unique_kw=$kwGroups.Count;unique_product_id=$productGroups.Count;unique_transition_code=$transitionGroups.Count}
    summaries=[ordered]@{
        cost_status=$statusSummary
        platform_status=$platformStatusSummary
        quality=[ordered]@{
            active_rows=$activeRows.Count
            active_blank_kw=$activeBlankKw.Count
            active_missing_cost=$activeMissingCost.Count
            active_missing_cost_with_stock=$activeMissingCostWithStock.Count
            duplicate_kw_groups=$duplicateKw.Count
            duplicate_kw_groups_with_multiple_active_skus=$activeDuplicateKwGroups.Count
            duplicate_kw_groups_with_at_most_one_active_sku=$singleActiveDuplicateKwGroups.Count
            expiry_present=$expiryPresent.Count
            expired=$expiredRows.Count
            expiring_within_8_months=$expiringRows.Count
            active_food_like_with_stock_missing_expiry=$activeFoodStockMissingExpiry.Count
        }
    }
    issues=[ordered]@{
        duplicate_kw=$duplicateKw
        duplicate_product_id=$duplicateProductId
        duplicate_transition_code=$duplicateTransition
        exact_duplicate_rows=$exactDuplicates
        duplicate_title_and_spec=$duplicateTitleSpec
        blank_kw=$blankKw
        invalid_kw=$invalidKw
        blank_product_id=$blankProductId
        invalid_product_id=$invalidProductId
        blank_title=$blankTitle
        missing_baseline_cost=$missingBaselineCost
        missing_moving_cost=$missingMovingCost
        negative_cost=$negativeCost
        invalid_units_per_pack=$invalidUnitsPerPack
        negative_cost_quantity=$negativeCostQuantity
        stale_missing_cost_status=$staleMissingStatus
        stale_conflict_status=$staleConflictStatus
        duplicate_kw_not_marked_conflict=$unmarkedKwConflicts
        suspicious_zero_cost_with_stock=$zeroCostWithStock
        active_blank_kw=$activeBlankKw
        active_missing_cost=$activeMissingCost
        active_missing_cost_with_stock=$activeMissingCostWithStock
        active_food_like_with_stock_missing_expiry=$activeFoodStockMissingExpiry
        expired=$expiredRows
        expiring_within_8_months=$expiringRows
        invalid_dates=@($invalidDates)
        inventory_duplicate_product_id=$inventoryDuplicateIds
        active_master_missing_inventory=$activeMasterMissingInventory
        inventory_missing_master=$inventoryMissingMaster
        kw_mismatch_vs_inventory=$kwInventoryMismatch
        used_order_product_missing_master=$orderProductMissingMaster
        formula_errors=$formulaErrors
    }
} | ConvertTo-Json -Depth 12 -Compress
