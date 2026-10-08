param(
    [string]$ProjectRoot,
    [switch]$SaveReport
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少经营驾驶舱元数据' }
$meta = Get-Content -LiteralPath $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $ProjectRoot 'agent\sku_master_schema.ps1')
$skuSchema = Get-SkuMasterSchema -SpreadsheetToken $token

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return '' }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return '' }
    return ([string]$cell.value).Trim()
}

$raw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A2:Z3000' --include value --max-chars 1200000 --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '读取SKU成本主表失败' }
$result = $raw | ConvertFrom-Json
if (-not $result.ok -or $result.data.has_more) { throw 'SKU成本主表读取不完整' }
$range = $result.data.ranges[0]
$cells = @($range.cells)
$rowNumbers = @($range.row_indices)

$items = [Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $cells.Count; $i++) {
    $row = @($cells[$i])
    $kw = (Get-CellValue $row $skuSchema.CurrentKw.Index).ToLowerInvariant()
    $csx = (Get-CellValue $row $skuSchema.TransitionCode.Index).ToLowerInvariant()
    $title = Get-CellValue $row $skuSchema.PlatformTitle.Index
    $skuId = Get-CellValue $row $skuSchema.ProductId.Index
    if (-not ($kw -or $csx -or $title -or $skuId)) { continue }
    $items.Add([pscustomobject]@{
        row = [int]$rowNumbers[$i]
        kw = $kw
        csx = $csx
        title = $title
        base_cost = Get-CellValue $row $skuSchema.BaselineUnitCost.Index
        average_cost = Get-CellValue $row $skuSchema.MovingAverageCost.Index
        cost_quantity = Get-CellValue $row $skuSchema.CostBookQuantity.Index
        cost_status = Get-CellValue $row $skuSchema.CostStatus.Index
        sku_id = $skuId
        sale_status = Get-CellValue $row $skuSchema.PlatformStatus.Index
        spec = Get-CellValue $row $skuSchema.CurrentSpec.Index
    })
}

$kwGroups = @($items | Where-Object { $_.kw } | Group-Object kw | Where-Object { $_.Count -gt 1 } | Sort-Object Name)
$csxGroups = @($items | Where-Object { $_.csx } | Group-Object csx | Where-Object { $_.Count -gt 1 } | Sort-Object Name)
$skuIdGroups = @($items | Where-Object { $_.sku_id } | Group-Object sku_id | Where-Object { $_.Count -gt 1 } | Sort-Object Name)
$exactGroups = @($items | Group-Object { '{0}|{1}|{2}|{3}' -f $_.kw,$_.csx,$_.title,$_.sku_id } | Where-Object { $_.Count -gt 1 })

$classified = [Collections.Generic.List[object]]::new()
foreach ($group in $kwGroups) {
    $legacy = @($group.Group | Where-Object { -not $_.sku_id })
    $platform = @($group.Group | Where-Object { $_.sku_id })
    $titles = @($group.Group.title | Where-Object { $_ } | Sort-Object -Unique)
    $specs = @($group.Group.spec | Where-Object { $_ } | Sort-Object -Unique)
    $category = if ($legacy.Count -gt 0 -and $platform.Count -gt 0) {
        if ($titles.Count -le 1) { '旧成本档案+平台规格（同商品）' } else { '旧成本档案+平台规格（KW换货或冲突）' }
    } elseif ($legacy.Count -eq 0) {
        if ($titles.Count -le 1) { '多个平台规格共用KW（同商品）' } else { '多个平台商品共用KW（冲突）' }
    } else {
        '多个旧成本档案共用KW'
    }
    $classified.Add([pscustomobject]@{
        kw = $group.Name
        category = $category
        row_count = $group.Count
        legacy_rows = @($legacy.row)
        platform_rows = @($platform.row)
        product_ids = @($platform.sku_id)
        distinct_title_count = $titles.Count
        distinct_spec_count = $specs.Count
        rows = @($group.Group | ForEach-Object {
            [pscustomobject]@{
                row = $_.row
                csx = $_.csx
                sku_id = $_.sku_id
                title = $_.title
                spec = $_.spec
                base_cost = $_.base_cost
                average_cost = $_.average_cost
                cost_quantity = $_.cost_quantity
                cost_status = $_.cost_status
                sale_status = $_.sale_status
            }
        })
    })
}

$audit = [ordered]@{
    ok = $true
    audited_at = [datetime]::Now.ToString('o')
    populated_rows = $items.Count
    last_populated_row = ($items.row | Measure-Object -Maximum).Maximum
    duplicate_kw_groups = $kwGroups.Count
    duplicate_csx_groups = $csxGroups.Count
    duplicate_sku_id_groups = $skuIdGroups.Count
    exact_identity_duplicate_groups = $exactGroups.Count
    classification_summary = @($classified | Group-Object category | ForEach-Object {
        [pscustomobject]@{
            category = $_.Name
            groups = $_.Count
            rows = (@($_.Group | ForEach-Object { $_.row_count }) | Measure-Object -Sum).Sum
        }
    })
    duplicate_kw_details = $classified
    duplicate_sku_id_details = @($skuIdGroups | ForEach-Object {
        [pscustomobject]@{ sku_id=$_.Name; rows=@($_.Group.row); kws=@($_.Group.kw | Sort-Object -Unique) }
    })
}

$json = $audit | ConvertTo-Json -Depth 12
if ($SaveReport) {
    $outputs = Join-Path $ProjectRoot 'outputs'
    if (-not (Test-Path -LiteralPath $outputs)) { $null = New-Item -ItemType Directory -Path $outputs }
    $reportPath = Join-Path $outputs ("sku_cost_master_duplicate_audit_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
    $json | Set-Content -LiteralPath $reportPath -Encoding UTF8
}
$json
