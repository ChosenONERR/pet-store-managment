param()
$ErrorActionPreference='Stop'
$OutputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$projectRoot=Split-Path $PSScriptRoot -Parent
$agentRoot=Join-Path $projectRoot 'agent'
$stateRoot=Join-Path $agentRoot 'state'
$larkCli=Join-Path $agentRoot 'lark_cli_noexit.ps1'
$powershellExe='C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$meta=Get-Content (Join-Path $projectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$ledgerMeta=Get-Content (Join-Path $projectRoot 'work\daily_ledger_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$sourceToken=([uri]$meta.workbook_url).Segments[-1].Trim('/')
$ledgerToken=([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$tokens=@{source=$sourceToken;ledger=$ledgerToken}
. (Join-Path $agentRoot 'finance_transaction.ps1')

function Invoke-LarkJson([scriptblock]$Action,[string]$Message){
    $raw=@(&$Action 2>&1|ForEach-Object{[string]$_})-join"`n"
    if($LASTEXITCODE-ne0){throw "$Message：$raw"}
    $result=$raw|ConvertFrom-Json
    if(-not$result.ok){throw "$Message：$raw"}
    return $result
}
function Get-CellValue($Row,[int]$Index){
    if($null-eq$Row-or$Index-ge@($Row).Count){return $null}
    $cell=@($Row)[$Index]
    if($null-eq$cell-or$null-eq$cell.value){return $null}
    return [string]$cell.value
}
function Convert-ToNumber($Value){
    if($null-eq$Value-or[string]$Value-eq''){return $null}
    $n=0.0
    if([double]::TryParse(([string]$Value).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){return $n}
    return $null
}
function Get-Rows([string]$Token,[string]$Sheet,[string]$Range,[string]$Include='value'){
    $result=Invoke-LarkJson {&$larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include $Include --max-chars 5000000 --format json} "读取$Sheet失败"
    if($result.data.has_more){throw "读取$Sheet不完整"}
    return @($result.data.ranges[0].cells)
}
function New-Matrix([object[]]$Values){
    $row=@();foreach($value in $Values){$row+=,@{value=$value}}
    $matrix=@();$matrix+=,$row;return,$matrix
}
function New-FormulaMatrix([string]$First,[string]$Second){
    $row=@(@{formula=$First},@{formula=$Second})
    $matrix=@();$matrix+=,$row;return,$matrix
}
function Invoke-Child([string]$Script,[string[]]$Arguments){
    $output=@(&$powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Script @Arguments 2>&1)
    if($LASTEXITCODE-ne0){throw "子脚本失败：$((@($output|ForEach-Object{[string]$_}|Select-Object -Last 4)-join' | '))"}
    return($output-join"`n")
}

$transactionId='TX-GIFT-NUMERIC-'+[datetime]::Now.ToString('yyyyMMdd-HHmmss')
$transaction=New-FinanceTransaction $stateRoot '规范赠品成本数值类型并刷新利润' $transactionId -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli
$closed=$false
try{
    foreach($snapshot in @(
        @('source',$sourceToken,'赠品记录','G2:J2001'),
        @('source',$sourceToken,'订单汇总','K2:L2001'),
        @('ledger',$ledgerToken,'8月收入','A2:K1001'),
        @('ledger',$ledgerToken,'8月支出','A2:F1001'),
        @('ledger',$ledgerToken,'9月收入','A2:K1001'),
        @('ledger',$ledgerToken,'9月支出','A2:F1001')
    )){Register-FinanceRangeBeforeImage $transaction $snapshot[0] $snapshot[1] $snapshot[2] $snapshot[3] $larkCli}

    $giftRows=Get-Rows $sourceToken '赠品记录' 'A2:L2001'
    $operations=[Collections.Generic.List[object]]::new()
    $normalized=0
    for($i=0;$i-lt$giftRows.Count;$i++){
        $row=$giftRows[$i]
        $orderNo=Get-CellValue $row 3
        $status=Get-CellValue $row 8
        $unitCost=Convert-ToNumber (Get-CellValue $row 6)
        $giftCost=Convert-ToNumber (Get-CellValue $row 7)
        if(-not$orderNo-or$status-eq'成本待核对'-or$null-eq$giftCost){continue}
        $sheetRow=2+$i
        $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='赠品记录';range="G${sheetRow}:J$sheetRow";cells=(New-Matrix @($unitCost,[double]$giftCost,'已计成本',(Get-CellValue $row 9)))}})
        $normalized++
    }
    if($operations.Count-eq0){throw '没有找到可规范化的赠品成本'}
    $payload=ConvertTo-Json -InputObject $operations.ToArray() -Depth 12 -Compress
    $null=Invoke-LarkJson {$payload|&$larkCli sheets +batch-update --as user --spreadsheet-token $sourceToken --operations - --yes --format json} '规范赠品成本失败'

    # Re-apply the existing formulas on affected rows to force Feishu to recalculate
    # after text amounts become numeric values.
    $summaryRows=Get-Rows $sourceToken '订单汇总' 'A2:Q2001' 'value,formula'
    $expected=[ordered]@{
        '4102299151749346084'=24.175
        '4102298960223685130'=43.17
        '3302299491023380387'=23.49
        '3402299431696905994'=4.51
    }
    $formulaOps=[Collections.Generic.List[object]]::new()
    $summaryRowByOrder=@{}
    for($i=0;$i-lt$summaryRows.Count;$i++){
        $row=$summaryRows[$i];$orderNo=Get-CellValue $row 2
        if(-not$orderNo-or-not$expected.Contains($orderNo)){continue}
        $sheetRow=2+$i;$summaryRowByOrder[$orderNo]=$sheetRow
        $k=@($row)[10].formula;$l=@($row)[11].formula
        if(-not$k-or-not$l){throw "订单汇总缺少公式：$orderNo"}
        $formulaOps.Add(@{shortcut='+cells-set';input=@{sheet_name='订单汇总';range="K${sheetRow}:L$sheetRow";cells=(New-FormulaMatrix $k $l)}})
    }
    if($formulaOps.Count-ne$expected.Count){throw '未完整定位4笔赠品订单'}
    $formulaPayload=ConvertTo-Json -InputObject $formulaOps.ToArray() -Depth 12 -Compress
    $null=Invoke-LarkJson {$formulaPayload|&$larkCli sheets +batch-update --as user --spreadsheet-token $sourceToken --operations - --yes --format json} '刷新订单汇总公式失败'
    Start-Sleep -Seconds 3

    $summaryAfter=Get-Rows $sourceToken '订单汇总' 'A2:Q2001'
    foreach($entry in $expected.GetEnumerator()){
        $row=@($summaryAfter)[$summaryRowByOrder[[string]$entry.Key]-2]
        $actual=Convert-ToNumber (Get-CellValue $row 10)
        if($null-eq$actual-or[math]::Abs([double]$actual-[double]$entry.Value)-gt0.02){throw "赠品订单成本未刷新：$($entry.Key)，实际=$actual，预期=$($entry.Value)"}
    }

    $null=Invoke-Child (Join-Path $agentRoot 'sync_daily_ledger.ps1') @('-ThroughDate','2026-09-22','-ProjectRoot',$projectRoot)
    foreach($check in @(
        @{sheet='赠品记录';range='G2:I2001'},
        @{sheet='订单汇总';range='K2:L2001'}
    )){
        $verify=Invoke-LarkJson {&$larkCli sheets +formula-verify --as user --spreadsheet-token $sourceToken --sheet-name $check.sheet --range $check.range --format json --exit-on-error} "公式校验失败：$($check.sheet)"
        if($verify.data.status-ne'success'){throw "公式校验未通过：$($check.sheet)"}
    }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed=$true
    [ordered]@{ok=$true;transaction_id=$transaction.id;atomic=$true;normalized_gift_rows=$normalized;corrected_orders=$expected.Keys}|ConvertTo-Json -Depth 5
}catch{
    $reason=[string]$_.Exception.Message
    $rollback=Invoke-FinanceTransactionRollback $transaction $tokens $larkCli $reason
    if($rollback.ok){throw "$reason；本次改动已全部回滚"}
    throw "$reason；自动回滚未完整完成：$($rollback.error)"
}finally{
    if(-not$closed){Close-FinanceTransaction $transaction}
}
