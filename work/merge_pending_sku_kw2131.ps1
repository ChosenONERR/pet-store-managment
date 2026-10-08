param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$meta = Get-Content -LiteralPath (Join-Path $ProjectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $ProjectRoot 'agent\sku_master_schema.ps1')
$schema = Get-SkuMasterSchema -SpreadsheetToken $token

function Value($Row,[int]$Index) { if ($Index -lt @($Row).Count -and $null -ne @($Row)[$Index] -and $null -ne @($Row)[$Index].value) { return [string]@($Row)[$Index].value }; return '' }
function Set-Matrix([string]$Range,$Cells) {
    $rows = [Collections.ArrayList]::new()
    $values = @($Cells)
    if ($values.Count -gt 0 -and $values[0] -is [Collections.IDictionary]) {
        $null = $rows.Add([object[]]$values)
    } else {
        foreach ($row in $values) { $null = $rows.Add([object[]]@($row)) }
    }
    $payload = ConvertTo-Json -InputObject $rows -Depth 10 -Compress
    $raw = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not ($raw|ConvertFrom-Json).ok) { throw "SKU合并写入失败：$Range" }
}

$raw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1399:S1400' --include value --max-chars 50000 --format json | Out-String
$result = $raw | ConvertFrom-Json
if (-not $result.ok) { throw '读取kw2131档案失败' }
$rows = @($result.data.ranges[0].cells)
if ($rows.Count -ne 2) { throw 'kw2131目标行结构变化，停止合并' }
$pending = $rows[0]; $platform = $rows[1]
if ((Value $pending $schema.CurrentKw.Index).ToLowerInvariant() -ne 'kw2131' -or (Value $platform $schema.CurrentKw.Index).ToLowerInvariant() -ne 'kw2131') { throw '目标行已不是两条kw2131，停止合并' }
if ((Value $pending $schema.PlatformStatus.Index) -ne '待平台同步' -or (Value $pending $schema.ProductId.Index)) { throw '第1399行已不是待平台同步占位档案，停止合并' }
if (-not (Value $platform $schema.ProductId.Index)) { throw '第1400行缺少稳定商品ID，停止合并' }

$backup = Join-Path $ProjectRoot ("agent\state\sku_kw2131_merge_backup_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
[ordered]@{source='SKU成本主表!A1399:S1400';captured_at=[datetime]::Now.ToString('o');cells=$rows} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $backup -Encoding UTF8

Set-Matrix "$($schema.PlatformTitle.Column)1399:$($schema.PlatformTitle.Column)1399" @(@(@{value=(Value $platform $schema.PlatformTitle.Index)}))
Set-Matrix "$($schema.ProductId.Column)1399:$($schema.CurrentSpec.Column)1399" @(@(
    @{value=(Value $platform $schema.ProductId.Index)},
    @{value=(Value $platform $schema.PlatformStatus.Index)},
    @{value=(Value $platform $schema.CurrentSpec.Index)}
))
$clearPayload = ConvertTo-Json -InputObject @('SKU成本主表!A1400:S1400') -Compress
$clearRaw = $clearPayload | lark-cli sheets +cells-batch-clear --as user --spreadsheet-token $token --ranges - --scope content --yes --format json | Out-String
if ($LASTEXITCODE -ne 0 -or -not ($clearRaw|ConvertFrom-Json).ok) { throw '清除已合并的kw2131重复平台档案失败' }

$checkRaw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1399:S1400' --include value --max-chars 50000 --format json | Out-String
$checkRows = @(($checkRaw|ConvertFrom-Json).data.ranges[0].cells)
if ((Value $checkRows[0] $schema.ProductId.Index) -ne (Value $platform $schema.ProductId.Index)) { throw 'kw2131合并后商品ID回读不一致' }
if (@($checkRows[1] | Where-Object {$null -ne $_.value -and [string]$_.value -ne ''}).Count -ne 0) { throw 'kw2131重复行未清空' }

[ordered]@{ok=$true;kept_row=1399;cleared_row=1400;backup=$backup;product_id=(Value $platform $schema.ProductId.Index);readback='success'} | ConvertTo-Json
