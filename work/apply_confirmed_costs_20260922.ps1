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
function New-Cell($Value) {
    if ($null -eq $Value) { return @{} }
    return @{value=$Value}
}
function New-Matrix([object[]]$Values) {
    $row = @(); foreach($value in $Values){$row += ,(New-Cell $value)}
    $matrix = @(); $matrix += ,$row; return ,$matrix
}
function Get-RangeRows([string]$Sheet,[string]$Range) {
    $result = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $sourceToken --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    return @($result.data.ranges[0].cells)
}
function Get-SingleRow([string]$Sheet,[string]$Range) {
    $result = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $sourceToken --sheet-name $Sheet --range $Range --include value --max-chars 10000 --format json } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    $matrix = $result.data.ranges[0].cells
    if($null -eq $matrix -or @($matrix).Count -eq 0){return $null}
    return ,(@($matrix)[0])
}
function Get-SingleCellValue([string]$Sheet,[string]$Range,[int]$Index) {
    $result = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $sourceToken --sheet-name $Sheet --range $Range --include value --max-chars 10000 --format json } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    $matrix = $result.data.ranges[0].cells
    if($null -eq $matrix -or @($matrix).Count -eq 0){return $null}
    $row = @($matrix)[0]
    if($Index -ge @($row).Count -or $null -eq @($row)[$Index].value){return $null}
    return [string]@($row)[$Index].value
}
function Invoke-Child([string]$Script,[string[]]$Arguments) {
    $output = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Script @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) {
        $detail = (@($output | ForEach-Object {[string]$_} | Where-Object {$_.Trim()} | Select-Object -Last 3) -join ' | ')
        throw "子脚本失败：$(Split-Path $Script -Leaf)：$detail"
    }
    return ($output -join "`n")
}

$confirmedCosts = [ordered]@{
    '51186993204' = 0.0
    '51181657349' = 7.2
    '51189679006' = 0.0
    '51185971656' = 21.0
    '51763658234' = 3.4
    '51180062722' = 45.0
    '51185331453' = 55.0
    '51185331394' = 100.0
    '51180542081' = 67.0
    '51181657253' = 6.5
    '51192127313' = 4.5
    '51192127345' = 4.1
    '51187300745' = 21.0
    '51194362170' = 0.71
    '51194362171' = 0.71
    '51194362172' = 0.71
    '51180062725' = 1.83
    '51180062744' = 2.1
    '51180062745' = 2.1
    '51180062746' = 2.1
    '52888357492' = 10.4
    '51189678983' = 1.68
    '51189678984' = 1.68
    '51189678985' = 1.68
    '51189678986' = 1.68
}

$transactionId = 'TX-COST-CONFIRM-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')
$transaction = New-FinanceTransaction $stateRoot '批量登记店主确认的SKU成本' $transactionId -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli
$closed = $false
try {
    $rows = Get-RangeRows 'SKU成本主表' 'A2:Z3000'
    $byProductId = @{}
    for($i=0;$i -lt $rows.Count;$i++){
        $productId = ([string](Get-CellValue $rows[$i] $schema.ProductId.Index)).Trim()
        if(-not $productId){continue}
        if($byProductId.ContainsKey($productId)){throw "SKU主表商品ID重复：$productId"}
        $byProductId[$productId] = [pscustomobject]@{row=2+$i;cells=$rows[$i]}
    }
    $operations = [Collections.Generic.List[object]]::new()
    $verifiedTargets = [Collections.Generic.List[object]]::new()
    foreach($entry in $confirmedCosts.GetEnumerator()){
        $productId = [string]$entry.Key
        $cost = [math]::Round([double]$entry.Value,4)
        if(-not $byProductId.ContainsKey($productId)){throw "SKU主表找不到商品ID：$productId"}
        $target = $byProductId[$productId]
        $rowNumber = [int]$target.row
        $oldNote = Get-CellValue $target.cells $schema.Note.Index
        $noteEntry = "店主确认销售单位成本为 $($cost.ToString('0.####',[Globalization.CultureInfo]::InvariantCulture)) 元（2026-09-22批量复核）"
        $note = ((@($oldNote,$noteEntry) | Where-Object {$_} | Select-Object -Unique) -join '；')
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.BaselineUnitCost.Column)${rowNumber}:$($schema.BaselineUnitCost.Column)$rowNumber";cells=(New-Matrix @($cost))}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.CostStatus.Column)${rowNumber}:$($schema.CostStatus.Column)$rowNumber";cells=(New-Matrix @('用户确认'))}})
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.ConversionBasis.Column)${rowNumber}:$($schema.Note.Column)$rowNumber";cells=(New-Matrix @('店主确认的销售单位成本',$note))}})
        $verifiedTargets.Add([pscustomobject]@{product_id=$productId;row=$rowNumber;cost=$cost})
    }

    # 已下架的高爷家试吃装不再占用 kw1117；保留稳定商品ID和历史记录。
    $retiredProductId = '51180062762'
    if(-not $byProductId.ContainsKey($retiredProductId)){throw "找不到待解绑旧商品：$retiredProductId"}
    $retired = $byProductId[$retiredProductId]
    if((Get-CellValue $retired.cells 0) -ne 'kw1117'){throw '旧高爷家试吃装当前KW不是kw1117，已停止写入'}
    if((Get-CellValue $retired.cells $schema.PlatformStatus.Index) -ne '已下架'){throw '旧高爷家试吃装尚未下架，已停止写入'}
    $retiredRow = [int]$retired.row
    $retiredNote = ((@((Get-CellValue $retired.cells $schema.Note.Index),'商品已下架，kw1117自2026-09-22起分配给皇家奶糕罐') | Where-Object {$_}) -join '；')
    $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="A${retiredRow}:A$retiredRow";cells=(New-Matrix @(''))}})
    $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.CostStatus.Column)${retiredRow}:$($schema.CostStatus.Column)$retiredRow";cells=(New-Matrix @('已下架/未分配KW'))}})
    $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='SKU成本主表';range="$($schema.Note.Column)${retiredRow}:$($schema.Note.Column)$retiredRow";cells=(New-Matrix @($retiredNote))}})

    # 历史试吃商品没有稳定商品ID。成本引擎按店主确认的统一规则将试吃/尝鲜/试用/体验装成本记为0，
    # 不再给历史订单临时补商品ID，避免同步重排后身份字段再次丢失。
    $trialOrderNo = '3402274371651171019'

    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' 'A2:A3000' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken 'SKU成本主表' 'E2:N3000' $larkCli
    $payload = ConvertTo-Json -InputObject $operations.ToArray() -Depth 12 -Compress
    $batchResult = Invoke-LarkJson { $payload | & $larkCli sheets +batch-update --as user --spreadsheet-token $sourceToken --operations - --yes --format json } '批量登记成本失败'

    $null = Invoke-Child (Join-Path $agentRoot 'rebuild_cost_snapshots.ps1') @('-ProjectRoot',$projectRoot)
    $null = Invoke-Child (Join-Path $agentRoot 'sync_daily_ledger.ps1') @('-ThroughDate','2026-09-22','-ProjectRoot',$projectRoot)

    $checkRows = Get-RangeRows 'SKU成本主表' 'A2:Z3000'
    foreach($target in $verifiedTargets){
        $row = @($checkRows)[$target.row-2]
        $actualCost = [double](Get-CellValue $row $schema.BaselineUnitCost.Index)
        $actualStatus = Get-CellValue $row $schema.CostStatus.Index
        if([math]::Abs($actualCost-$target.cost) -gt 0.0001 -or $actualStatus -ne '用户确认'){
            throw "成本回读不一致：$($target.product_id)"
        }
    }
    $retiredCheck = @($checkRows)[$retiredRow-2]
    if(Get-CellValue $retiredCheck 0){throw '旧kw1117解绑回读失败'}
    $trialRows = Get-RangeRows '订单商品' 'A2:Q4001'
    $trialMatches = @($trialRows | Where-Object { (Get-CellValue $_ 2) -eq $trialOrderNo -and (Get-CellValue $_ 6) -match '高爷家' })
    if($trialMatches.Count -ne 1){throw "历史试吃商品回读定位异常：找到 $($trialMatches.Count) 行"}
    if([math]::Abs([double](Get-CellValue $trialMatches[0] 11)) -gt 0.0001 -or [math]::Abs([double](Get-CellValue $trialMatches[0] 12)) -gt 0.0001){throw '高爷家试吃装成本没有重算为0'}

    foreach($check in @(
        @{sheet='订单商品';range='L2:N4001'},
        @{sheet='订单汇总';range='K2:L2001'},
        @{sheet='SKU成本主表';range='E2:G3000'}
    )){
        $verify = Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $sourceToken --sheet-name $check.sheet --range $check.range --format json --exit-on-error } "公式校验失败：$($check.sheet)"
        if($verify.data.status -ne 'success'){throw "公式校验未通过：$($check.sheet)"}
    }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok=$true
        transaction_id=$transaction.id
        confirmed_sku_count=$verifiedTargets.Count
        retired_kw='kw1117'
        retired_product_id=$retiredProductId
        corrected_order='3402274371651171019'
        corrected_order_trial_cost=0
        formula_verification='success'
        atomic=$true
    } | ConvertTo-Json -Depth 6
} catch {
    $reason = [string]$_.Exception.Message
    $rollback = Invoke-FinanceTransactionRollback $transaction $tokens $larkCli $reason
    if($rollback.ok){throw "$reason；本次批量改动已全部回滚"}
    throw "$reason；自动回滚未完整完成：$($rollback.error)"
} finally {
    if(-not $closed){Close-FinanceTransaction $transaction}
}
