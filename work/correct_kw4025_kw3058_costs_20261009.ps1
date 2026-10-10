param()

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$projectRoot = Split-Path $PSScriptRoot -Parent
$agentRoot = Join-Path $projectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$powershellExe = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$meta = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\daily_ledger_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$tokens = @{source=$sourceToken;ledger=$ledgerToken}
. (Join-Path $agentRoot 'finance_transaction.ps1')
. (Join-Path $agentRoot 'sku_master_schema.ps1')
$schema = Get-SkuMasterSchema -SpreadsheetToken $sourceToken

function Invoke-LarkJson([scriptblock]$Action,[string]$Message) {
    $raw = @(& $Action 2>&1 | ForEach-Object { [string]$_ }) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "$Message：$raw" }
    try { $result = $raw | ConvertFrom-Json } catch { throw "$Message：返回内容无法解析" }
    if (-not $result.ok) { throw "$Message：$raw" }
    return $result
}
function Get-CellValue($Row,[int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return $cell.value
}
function New-Cell($Value) {
    if ($null -eq $Value) { return @{} }
    return @{value=$Value}
}
function New-UnitCostCell($Value) {
    if ($null -eq $Value) { return @{} }
    return @{value=$Value;cell_styles=@{number_format='#,##0.####'}}
}
function New-Matrix([object[]]$Values) {
    $row = @(); foreach($value in $Values){$row += ,(New-Cell $value)}
    $matrix = @(); $matrix += ,$row; return ,$matrix
}
function Get-Rows([string]$Token,[string]$Sheet,[string]$Range) {
    $result = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    return @($result.data.ranges[0].cells)
}
function Invoke-Child([string]$Script,[string[]]$Arguments) {
    $output = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Script @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) {
        $detail = (@($output | ForEach-Object {[string]$_} | Where-Object {$_.Trim()} | Select-Object -Last 5) -join ' | ')
        throw "子脚本失败：$(Split-Path $Script -Leaf)：$detail"
    }
    return ($output -join "`n")
}
function Assert-Near($Actual,[double]$Expected,[string]$Label,[double]$Tolerance=0.001) {
    if ($null -eq $Actual -or [math]::Abs([double]$Actual-$Expected) -gt $Tolerance) {
        throw "$Label 校验失败：实际=$Actual，预期=$Expected"
    }
}

$targets = [ordered]@{
    '51180062747' = [pscustomobject]@{cost=1.188;expected_old=5.94;status='当前KW冲突待核对（成本已确认）';note='店主确认kw4025销售单位成本；5.94元为5条合计，不是单条成本'}
    '51180062748' = [pscustomobject]@{cost=1.188;expected_old=$null;status='当前KW冲突待核对（成本已确认）';note='店主确认kw4025销售单位成本；5.94元为5条合计，不是单条成本'}
    '51192927142' = [pscustomobject]@{cost=12.5;expected_old=$null;status='用户确认';note='店主确认kw3058销售单位成本'}
}

$transaction = New-FinanceTransaction $stateRoot '修正kw4025与kw3058成本并重算订单' ('TX-COST-FIX-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')) -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli
$closed = $false
try {
    $masterRows = Get-Rows $sourceToken 'SKU成本主表' 'A2:Z3000'
    $byProductId = @{}
    for($i=0;$i -lt $masterRows.Count;$i++) {
        $productId = ([string](Get-CellValue $masterRows[$i] $schema.ProductId.Index)).Trim()
        if (-not $productId) { continue }
        if ($byProductId.ContainsKey($productId)) { throw "SKU成本主表商品ID重复：$productId" }
        $byProductId[$productId] = [pscustomobject]@{row=2+$i;cells=$masterRows[$i]}
    }

    $operations = [Collections.Generic.List[object]]::new()
    foreach($entry in $targets.GetEnumerator()) {
        $productId = [string]$entry.Key
        $target = $entry.Value
        if (-not $byProductId.ContainsKey($productId)) { throw "SKU成本主表找不到商品ID：$productId" }
        $record = $byProductId[$productId]
        $rowNumber = [int]$record.row
        $oldCostText = Get-CellValue $record.cells $schema.BaselineUnitCost.Index
        if ($null -eq $target.expected_old) {
            if ($null -ne $oldCostText -and [string]$oldCostText -ne '') { throw "商品ID $productId 原成本已变化：$oldCostText" }
        } else {
            Assert-Near $oldCostText ([double]$target.expected_old) "商品ID $productId 原成本"
        }
        $oldNote = [string](Get-CellValue $record.cells $schema.Note.Index)
        $noteEntry = "店主确认销售单位成本为 $(([double]$target.cost).ToString('0.####',[Globalization.CultureInfo]::InvariantCulture)) 元（2026-10-09）"
        $newNote = ((@($oldNote,$noteEntry,[string]$target.note) | Where-Object { $_ } | Select-Object -Unique) -join '；')
        Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' "$($schema.BaselineUnitCost.Column)${rowNumber}:$($schema.BaselineUnitCost.Column)$rowNumber" $larkCli
        Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' "$($schema.CostStatus.Column)${rowNumber}:$($schema.CostStatus.Column)$rowNumber" $larkCli
        Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' "$($schema.ConversionBasis.Column)${rowNumber}:$($schema.Note.Column)$rowNumber" $larkCli
        $unitCostMatrix = @(); $unitCostMatrix += ,@((New-UnitCostCell ([double]$target.cost)))
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.BaselineUnitCost.Column)${rowNumber}:$($schema.BaselineUnitCost.Column)$rowNumber";cells=$unitCostMatrix}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.CostStatus.Column)${rowNumber}:$($schema.CostStatus.Column)$rowNumber";cells=(New-Matrix @([string]$target.status))}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.ConversionBasis.Column)${rowNumber}:$($schema.Note.Column)$rowNumber";cells=(New-Matrix @('店主确认的销售单位成本',$newNote))}})
    }

    $payload = ConvertTo-Json -InputObject $operations.ToArray() -Depth 12 -Compress
    $null = Invoke-LarkJson { $payload | & $larkCli sheets +batch-update --as user --spreadsheet-token $sourceToken --operations - --yes --format json } '批量修改SKU成本失败'
    $null = Invoke-Child (Join-Path $agentRoot 'rebuild_cost_snapshots.ps1') @('-ProjectRoot',$projectRoot)
    $null = Invoke-Child (Join-Path $agentRoot 'sync_daily_ledger.ps1') @('-ThroughDate','2026-10-09','-ProjectRoot',$projectRoot)

    $masterAfter = Get-Rows $sourceToken 'SKU成本主表' 'A2:Z3000'
    foreach($entry in $targets.GetEnumerator()) {
        $productId = [string]$entry.Key
        $target = $entry.Value
        $match = @($masterAfter | Where-Object { ([string](Get-CellValue $_ $schema.ProductId.Index)).Trim() -eq $productId })
        if ($match.Count -ne 1) { throw "商品ID $productId 回读不唯一" }
        Assert-Near (Get-CellValue $match[0] $schema.BaselineUnitCost.Index) ([double]$target.cost) "商品ID $productId 基准成本"
        Assert-Near (Get-CellValue $match[0] $schema.MovingAverageCost.Index) ([double]$target.cost) "商品ID $productId 移动平均成本"
    }

    $orderNo = '3202336871182598943'
    $itemRows = @(Get-Rows $sourceToken '订单商品' 'A2:Q4001' | Where-Object { [string](Get-CellValue $_ 2) -eq $orderNo })
    if ($itemRows.Count -ne 3) { throw "今日美团3号单商品行异常：$($itemRows.Count)" }
    $itemByProductId = @{}; foreach($row in $itemRows){$itemByProductId[[string](Get-CellValue $row 16)]=$row}
    Assert-Near (Get-CellValue $itemByProductId['51192927142'] 11) 12.5 'kw3058单位成本'
    Assert-Near (Get-CellValue $itemByProductId['51192927142'] 12) 12.5 'kw3058商品成本'
    Assert-Near (Get-CellValue $itemByProductId['51180062748'] 11) 1.188 'kw4025单位成本'
    Assert-Near (Get-CellValue $itemByProductId['51180062748'] 12) 5.94 'kw4025商品成本'

    $summaryMatch = @(Get-Rows $sourceToken '订单汇总' 'A2:P2001' | Where-Object { [string](Get-CellValue $_ 2) -eq $orderNo })
    if ($summaryMatch.Count -ne 1) { throw '今日美团3号单汇总回读不唯一' }
    Assert-Near (Get-CellValue $summaryMatch[0] 10) 73.73 '订单汇总货物成本' 0.011
    Assert-Near (Get-CellValue $summaryMatch[0] 11) 32.42 '订单汇总利润' 0.011

    $incomeMatch = @(Get-Rows $ledgerToken '10月收入' 'A2:J1001' | Where-Object { [string](Get-CellValue $_ 1) -eq $orderNo })
    if ($incomeMatch.Count -ne 1) { throw '10月收入今日美团3号单回读不唯一' }
    Assert-Near (Get-CellValue $incomeMatch[0] 4) 73.73 '10月收入货物成本' 0.011
    Assert-Near (Get-CellValue $incomeMatch[0] 6) 32.42 '10月收入利润' 0.011

    foreach($check in @(
        @{token=$sourceToken;sheet='订单商品';range='L2:N4001'},
        @{token=$sourceToken;sheet='订单汇总';range='K2:L2001'},
        @{token=$sourceToken;sheet='SKU成本主表';range='F2:G3000'},
        @{token=$ledgerToken;sheet='10月收入';range='D2:G1001'}
    )) {
        $verify = Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $check.token --sheet-name $check.sheet --range $check.range --format json --exit-on-error } "公式校验失败：$($check.sheet)"
        if ($verify.data.status -ne 'success') { throw "公式校验未通过：$($check.sheet)" }
    }
    $null = Invoke-Child (Join-Path $projectRoot 'work\audit_finance_integrity.ps1') @('-ProjectRoot',$projectRoot)

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok=$true
        transaction_id=$transaction.id
        corrected_sku_rows=3
        order_no=$orderNo
        product_cost=73.73
        profit=32.42
        formula_verification='success'
        integrity_audit='success'
        atomic=$true
    } | ConvertTo-Json -Depth 6
} catch {
    $reason = [string]$_.Exception.Message
    $rollback = Invoke-FinanceTransactionRollback $transaction $tokens $larkCli $reason
    if ($rollback.ok) { throw "$reason；本次改动已全部回滚" }
    throw "$reason；自动回滚未完整完成：$($rollback.error)"
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
