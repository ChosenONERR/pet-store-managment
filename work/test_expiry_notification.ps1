$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'agent\run_sync_and_notify.ps1'
$source = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) { throw "主脚本语法错误：$($errors.Message -join '；')" }

$requiredFunctions = @(
    'Get-CellValue',
    'Convert-ToNumber',
    'Convert-ToDateValue',
    'Get-ExpiryRemainingText',
    'Get-ExpiringProductRecords',
    'Format-Money',
    'Limit-Text',
    'Get-FailureGuidance',
    'Build-Card',
    'Build-Message'
)
foreach ($name in $requiredFunctions) {
    $functionAst = $ast.FindAll({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true) | Select-Object -First 1
    if ($null -eq $functionAst) { throw "找不到函数：$name" }
    Invoke-Expression $functionAst.Extent.Text
}

function New-TestCellRow([object[]]$Values) {
    return ,@($Values | ForEach-Object {
        if ($null -eq $_) { [pscustomobject]@{} } else { [pscustomobject]@{ value = [string]$_ } }
    })
}

$asOfDate = [datetime]'2026-09-13'
$insideCutoff = $asOfDate.AddMonths(8)
$outsideCutoff = $insideCutoff.AddDays(1)
$insideSerial = ($insideCutoff - [datetime]'1899-12-30').TotalDays
$masterRows = @(
    (New-TestCellRow @('kw1001','csx1','正好8个月',1,10,10,2,1,$null,'已核对','2026-01-01','2026-09-01',$null,$null,$insideSerial,'sku-1')),
    (New-TestCellRow @('kw1002','csx2','超过8个月',1,10,10,2,1,$null,'已核对','2026-01-01','2026-09-01',$null,$null,$outsideCutoff.ToString('yyyy-MM-dd'),'sku-2')),
    (New-TestCellRow @('kw1003','csx3','已经过期',1,10,10,3,1,$null,'已核对','2026-01-01','2026-09-01',$null,$null,'2026-09-01','sku-3')),
    (New-TestCellRow @('kw1004','csx4','零库存',1,10,10,5,1,$null,'已核对','2026-01-01','2026-09-01',$null,$null,'2026-10-01','sku-4')),
    (New-TestCellRow @($null,'csx5','过渡货号',1,10,10,1,1,$null,'已核对','2026-01-01','2026-09-01',$null,$null,'2027-01-01','sku-5')),
    (New-TestCellRow @('kw1006','csx6','未登记日期',1,10,10,1,1,$null,'已核对','2026-01-01','2026-09-01',$null,$null,$null,'sku-6'))
)
$inventoryRecords = @(
    [pscustomobject]@{code='kw1001';shelf_code='csx1';name='正好8个月';stock=2},
    [pscustomobject]@{code='kw1002';shelf_code='csx2';name='超过8个月';stock=2},
    [pscustomobject]@{code='kw1003';shelf_code='csx3';name='已经过期';stock=3},
    [pscustomobject]@{code='kw1004';shelf_code='csx4';name='零库存';stock=0},
    [pscustomobject]@{code='csx5';shelf_code='';name='过渡货号';stock=1}
)
$script:skuSchema = [pscustomobject]@{
    EarliestExpiry = [pscustomobject]@{Index=14}
    CostBookQuantity = [pscustomobject]@{Index=6}
}

$items = @(Get-ExpiringProductRecords $masterRows $inventoryRecords $asOfDate)
if ($items.Count -ne 3) { throw "临期筛选数量错误：预期3，实际$($items.Count)" }
if ($items[0].code -ne 'kw1003' -or $items[0].remaining_text -notmatch '^已过期') { throw '已过期商品没有排在最前或状态错误' }
if (-not ($items.code -contains 'kw1001')) { throw '正好8个自然月的商品未被包含' }
if ($items.code -contains 'kw1002') { throw '超过8个自然月的商品被错误包含' }
if ($items.code -contains 'kw1004') { throw '零库存商品被错误包含' }
if (-not ($items.code -contains 'csx5')) { throw '无KW的过渡货号未被包含' }

$script:Mode = 'night'
$script:Date = $asOfDate
$script:meta = [pscustomobject]@{ workbook_url = 'https://example.invalid/sheets/test' }
$snapshot = [pscustomobject]@{
    sync_warnings = @()
    today_net_profit = 12.34
    today_total = 3
    today_natural = 2
    today_intervention = 1
    today_offline = 0
    today_cancelled = 0
    pending_cost = @()
    pending_delivery = @()
    urgent_restock = @()
    restock_excluded_samples = @()
    restore_inventory = @()
    expiring_products = $items
    expiry_cutoff = $insideCutoff.ToString('yyyy-MM-dd')
    metrics = @{'当前缺货SKU数'=1;'当前低库存SKU数'=2;'待核对数据数'=3;'本月经营毛利'=12.34}
}
$card = Build-Card $snapshot $true
$cardJson = $card | ConvertTo-Json -Depth 30 -Compress
if ($cardJson -notmatch '临期商品' -or $cardJson -notmatch '正好8个月' -or $cardJson -notmatch '2027-05-13') { throw 'Card 2.0 未包含预期临期信息' }
$expiryTables = @($card.body.elements | Where-Object { $_.tag -eq 'table' -and @($_.columns).display_name -contains '到期日' })
if ($expiryTables.Count -ne 1 -or @($expiryTables[0].rows).Count -ne 3) { throw '临期表格结构或行数错误' }

$message = Build-Message $snapshot $true
if ($message -notmatch '临期商品：3 款' -or $message -notmatch 'csx5') { throw '文本兜底消息未包含预期临期信息' }

[ordered]@{
    ok = $true
    assertions = 11
    cutoff = $insideCutoff.ToString('yyyy-MM-dd')
    expiring_count = $items.Count
    card_expiry_rows = @($expiryTables[0].rows).Count
} | ConvertTo-Json
