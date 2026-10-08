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
    return [string]$cell.value
}
function Get-Rows([string]$Token,[string]$Sheet,[string]$Range) {
    $result = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    return @($result.data.ranges[0].cells)
}
function New-Matrix([object[]]$Values) {
    $row = @(); foreach($value in $Values){$row += ,@{value=$value}}
    $matrix = @(); $matrix += ,$row; return ,$matrix
}
function Invoke-Child([string]$Script,[string[]]$Arguments) {
    $output = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Script @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) {
        $detail = (@($output | ForEach-Object {[string]$_} | Where-Object {$_.Trim()} | Select-Object -Last 4) -join ' | ')
        throw "子脚本失败：$(Split-Path $Script -Leaf)：$detail"
    }
    return ($output -join "`n")
}

$transactionId = 'TX-KW1106-DENTAL-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')
$transaction = New-FinanceTransaction $stateRoot '修正kw1106补货关联并拆分牙膏牙刷成本' $transactionId -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli
$closed = $false
try {
    $replenishmentRow = @(Get-Rows $sourceToken '补货记录' 'A2:T2')[0]
    if((Get-CellValue $replenishmentRow 1).ToLowerInvariant() -ne 'kw1003' -or
       [math]::Abs([double](Get-CellValue $replenishmentRow 2)-6) -gt 0.0001 -or
       [math]::Abs([double](Get-CellValue $replenishmentRow 5)-600) -gt 0.0001 -or
       (Get-CellValue $replenishmentRow 14) -ne '51185971656') {
        throw '补货记录第2行已变化，停止自动修正'
    }

    $masterRows = Get-Rows $sourceToken 'SKU成本主表' 'A2:Z3000'
    $masterByProductId = @{}
    for($i=0;$i -lt $masterRows.Count;$i++){
        $productId = ([string](Get-CellValue $masterRows[$i] $schema.ProductId.Index)).Trim()
        if(-not $productId){continue}
        if($masterByProductId.ContainsKey($productId)){throw "SKU主表商品ID重复：$productId"}
        $masterByProductId[$productId]=[pscustomobject]@{row=2+$i;cells=$masterRows[$i]}
    }
    $dentalCosts = [ordered]@{
        '51180542085' = [pscustomobject]@{cost=10.29;label='牙膏';ratio='按当时售价34:4拆分首批每套11.5元'}
        '51180542086' = [pscustomobject]@{cost=1.21;label='牙刷';ratio='按当时售价34:4拆分首批每套11.5元'}
    }
    foreach($productId in $dentalCosts.Keys){if(-not $masterByProductId.ContainsKey($productId)){throw "找不到牙膏牙刷商品ID：$productId"}}

    # Rebuild and ledger refresh touch these derived ranges. Capture all of them so
    # a downstream failure restores the complete accounting state.
    foreach($snapshot in @(
        @('source',$sourceToken,'补货记录','H2:O1001'),
        @('source',$sourceToken,'SKU成本主表','F2:O3000'),
        @('source',$sourceToken,'订单商品','L2:N4001'),
        @('source',$sourceToken,'赠品记录','G2:J2001'),
        @('source',$sourceToken,'订单汇总','K2:L2001'),
        @('ledger',$ledgerToken,'8月收入','A2:K1001'),
        @('ledger',$ledgerToken,'8月支出','A2:F1001'),
        @('ledger',$ledgerToken,'9月收入','A2:K1001'),
        @('ledger',$ledgerToken,'9月支出','A2:F1001')
    )){ Register-FinanceRangeBeforeImage $transaction $snapshot[0] $snapshot[1] $snapshot[2] $snapshot[3] $larkCli }

    $operations = [Collections.Generic.List[object]]::new()
    # The six 2kg bags belong to kw1003 (cost 100), not kw1106 400g.
    $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='补货记录';range='O2:O2';cells=(New-Matrix @('51185331394'))}})
    foreach($entry in $dentalCosts.GetEnumerator()){
        $target=$masterByProductId[[string]$entry.Key]
        $row=[int]$target.row
        $oldNote=Get-CellValue $target.cells $schema.Note.Index
        $note=((@($oldNote,"店主确认：$($entry.Value.label)首批单位成本$($entry.Value.cost)元；后补3套每套7.5元的日期待补录")|Where-Object{$_}|Select-Object -Unique)-join '；')
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.BaselineUnitCost.Column)${row}:$($schema.BaselineUnitCost.Column)$row";cells=(New-Matrix @([double]$entry.Value.cost))}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.CostStatus.Column)${row}:$($schema.CostStatus.Column)$row";cells=(New-Matrix @('用户确认'))}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.ConversionBasis.Column)${row}:$($schema.Note.Column)$row";cells=(New-Matrix @([string]$entry.Value.ratio,$note))}})
    }
    $payload=ConvertTo-Json -InputObject $operations.ToArray() -Depth 12 -Compress
    $null=Invoke-LarkJson { $payload | & $larkCli sheets +batch-update --as user --spreadsheet-token $sourceToken --operations - --yes --format json } '写入修正失败'

    $null=Invoke-Child (Join-Path $agentRoot 'rebuild_cost_snapshots.ps1') @('-ProjectRoot',$projectRoot)
    $null=Invoke-Child (Join-Path $agentRoot 'sync_daily_ledger.ps1') @('-ThroughDate','2026-09-22','-ProjectRoot',$projectRoot)

    $masterAfter=Get-Rows $sourceToken 'SKU成本主表' 'A2:Z3000'
    function Find-Master([string]$ProductId){
        for($i=0;$i -lt $masterAfter.Count;$i++){if((Get-CellValue $masterAfter[$i] $schema.ProductId.Index)-eq$ProductId){return $masterAfter[$i]}}
        return $null
    }
    $kw1003=Find-Master '51185331394'
    $kw1106=Find-Master '51185971656'
    if($null-eq$kw1003 -or [math]::Abs([double](Get-CellValue $kw1003 $schema.MovingAverageCost.Index)-100)-gt0.001 -or [math]::Abs([double](Get-CellValue $kw1003 $schema.CostBookQuantity.Index)-6)-gt0.001){throw 'kw1003重算结果异常'}
    if($null-eq$kw1106 -or [math]::Abs([double](Get-CellValue $kw1106 $schema.MovingAverageCost.Index)-20.976)-gt0.001){throw 'kw1106当前移动平均成本异常'}

    $items=Get-Rows $sourceToken '订单商品' 'A2:Q4001'
    $kw1106Sale=@($items|Where-Object{(Get-CellValue $_ 2)-eq'4202307393525100811' -and (Get-CellValue $_ 16)-eq'51185971656'})
    if($kw1106Sale.Count-ne1 -or [math]::Abs([double](Get-CellValue $kw1106Sale[0] 11)-21)-gt0.001){throw '9月19日kw1106订单成本未恢复为21元'}
    $dentalProductIds=@('51180542085','51180542086')
    $dentalSale=@($items|Where-Object{(Get-CellValue $_ 2)-eq'3202277921913852108' -and (Get-CellValue $_ 16)-in$dentalProductIds})
    $dentalTotal=0.0
    foreach($dentalItem in $dentalSale){$dentalTotal += [double](Get-CellValue $dentalItem 12)}
    $toothpaste=@($dentalSale|Where-Object{(Get-CellValue $_ 16)-eq'51180542085'})
    $toothbrush=@($dentalSale|Where-Object{(Get-CellValue $_ 16)-eq'51180542086'})
    if($dentalSale.Count-ne2 -or $toothpaste.Count-ne1 -or $toothbrush.Count-ne1 -or
       [math]::Abs([double](Get-CellValue $toothpaste[0] 11)-10.29)-gt0.001 -or
       [math]::Abs([double](Get-CellValue $toothbrush[0] 11)-1.21)-gt0.001 -or
       [math]::Abs([double]$dentalTotal-11.5)-gt0.001){
        throw "牙膏牙刷历史订单成本拆分异常：行数=$($dentalSale.Count)，牙膏=$(@($toothpaste|ForEach-Object{Get-CellValue $_ 11})-join ',')，牙刷=$(@($toothbrush|ForEach-Object{Get-CellValue $_ 11})-join ',')，合计=$dentalTotal"
    }

    foreach($check in @(
        @{sheet='订单商品';range='L2:N4001'},
        @{sheet='订单汇总';range='K2:L2001'},
        @{sheet='SKU成本主表';range='F2:H3000'},
        @{sheet='补货记录';range='D2:K1001'}
    )){
        $verify=Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $sourceToken --sheet-name $check.sheet --range $check.range --format json --exit-on-error } "公式校验失败：$($check.sheet)"
        if($verify.data.status-ne'success'){throw "公式校验未通过：$($check.sheet)"}
    }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed=$true
    [ordered]@{
        ok=$true;transaction_id=$transaction.id;atomic=$true
        kw1003_current_cost=100;kw1003_cost_book_quantity=6
        kw1106_sep19_cost=21;kw1106_current_moving_average=20.976
        dental_order='3202277921913852108';dental_order_total_cost=11.5
        dental_first_batch=@{toothpaste=10.29;toothbrush=1.21}
        dental_later_batch_pending_date=$true
    }|ConvertTo-Json -Depth 6
} catch {
    $reason=[string]$_.Exception.Message
    $rollback=Invoke-FinanceTransactionRollback $transaction $tokens $larkCli $reason
    if($rollback.ok){throw "$reason；本次改动已全部回滚"}
    throw "$reason；自动回滚未完整完成：$($rollback.error)"
} finally {
    if(-not$closed){Close-FinanceTransaction $transaction}
}
