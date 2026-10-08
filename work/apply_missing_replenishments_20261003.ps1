param()

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$projectRoot = Split-Path $PSScriptRoot -Parent
$agentRoot = Join-Path $projectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$powershellExe = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$meta = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\simple_finance_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\daily_ledger_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$tokens = @{source=$sourceToken;ledger=$ledgerToken}
. (Join-Path $agentRoot 'finance_transaction.ps1')
. (Join-Path $agentRoot 'sku_master_schema.ps1')
$schema = Get-SkuMasterSchema -SpreadsheetToken $sourceToken

$requests = @(
    [pscustomobject]@{code='kw1035';date=[datetime]'2026-08-30';quantity=3.0;amount=74.80},
    [pscustomobject]@{code='kw4212';date=[datetime]'2026-08-31';quantity=3.0;amount=88.11},
    [pscustomobject]@{code='kw4213';date=[datetime]'2026-08-31';quantity=4.0;amount=130.45},
    [pscustomobject]@{code='kw2115';date=[datetime]'2026-09-18';quantity=3.0;amount=192.76}
)

function Invoke-LarkJson([scriptblock]$Action, [string]$Message) {
    $raw = @(& $Action 2>&1 | ForEach-Object { [string]$_ }) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "$Message：$raw" }
    try { $result = $raw | ConvertFrom-Json } catch { throw "$Message：返回内容无法解析" }
    if (-not $result.ok) { throw "$Message：$raw" }
    return $result
}

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}

function Convert-ToNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if (-not [double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $null
    }
    return $number
}

function Convert-ToDate($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) { return $date.Date }
    $number = Convert-ToNumber $Value
    if ($null -ne $number) { return ([datetime]'1899-12-30').AddDays($number).Date }
    return $null
}

function Convert-ToExcelSerial([datetime]$Value) { return ($Value.Date - [datetime]'1899-12-30').TotalDays }
function New-ValueCell($Value) { if ($null -eq $Value -or [string]$Value -eq '') { return @{} }; return @{value=$Value} }
function New-FormulaCell([string]$Formula) { return @{formula=$Formula} }

function Get-Rows([string]$Sheet, [string]$Range) {
    $result = Invoke-LarkJson {
        & $larkCli sheets +cells-get --as user --spreadsheet-token $sourceToken --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json
    } "读取 $Sheet 失败"
    if ($result.data.has_more) { throw "读取 $Sheet 不完整" }
    return @($result.data.ranges[0].cells)
}

function Invoke-Child([string]$Script, [string[]]$Arguments) {
    $output = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Script @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) {
        $detail = (@($output | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() } | Select-Object -Last 5) -join ' | ')
        throw "子脚本失败：$(Split-Path $Script -Leaf)：$detail"
    }
    return ($output -join "`n")
}

function Get-ItemSnapshot($Rows) {
    $result = @{}
    for ($i = 1; $i -lt $Rows.Count; $i++) {
        $orderNo = Get-CellValue $Rows[$i] 2
        if (-not $orderNo) { continue }
        $result[2 + $i - 1] = [pscustomobject]@{
            row = 1 + $i
            date = Convert-ToDate (Get-CellValue $Rows[$i] 0)
            order_no = $orderNo
            code = ([string](Get-CellValue $Rows[$i] 5)).Trim().ToLowerInvariant()
            title = Get-CellValue $Rows[$i] 6
            quantity = Convert-ToNumber (Get-CellValue $Rows[$i] 8)
            unit_cost = Convert-ToNumber (Get-CellValue $Rows[$i] 11)
            item_cost = Convert-ToNumber (Get-CellValue $Rows[$i] 12)
        }
    }
    return $result
}

function Test-NumberChanged($Before, $After, [double]$Tolerance = 0.0001) {
    if ($null -eq $Before -and $null -eq $After) { return $false }
    if ($null -eq $Before -or $null -eq $After) { return $true }
    return [math]::Abs([double]$Before - [double]$After) -gt $Tolerance
}

$transactionId = 'TX-MISSING-REPLENISHMENTS-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')
$transaction = New-FinanceTransaction $stateRoot '补录4笔历史补货并重算相关订单成本' $transactionId -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli
$closed = $false
try {
    $masterRows = Get-Rows 'SKU成本主表' 'A1:Z3000'
    $replenishmentRows = Get-Rows '补货记录' 'A1:T1001'
    $itemRowsBefore = Get-Rows '订单商品' 'A1:Q4001'
    $itemsBefore = Get-ItemSnapshot $itemRowsBefore

    $targets = [Collections.Generic.List[object]]::new()
    foreach ($request in $requests) {
        $matches = [Collections.Generic.List[object]]::new()
        for ($i = 1; $i -lt $masterRows.Count; $i++) {
            $currentKw = ([string](Get-CellValue $masterRows[$i] $schema.CurrentKw.Index)).Trim().ToLowerInvariant()
            if ($currentKw -eq $request.code) { $matches.Add([pscustomobject]@{row=1+$i;cells=$masterRows[$i]}) }
        }
        if ($matches.Count -ne 1) { throw "SKU成本主表无法唯一定位 $($request.code)：匹配 $($matches.Count) 行" }
        $target = $matches[0]
        $unitsPerPack = Convert-ToNumber (Get-CellValue $target.cells $schema.UnitsPerPack.Index)
        if ($null -eq $unitsPerPack -or $unitsPerPack -le 0) { throw "$($request.code) 的每采购包装含销售单位数无效" }
        $targets.Add([pscustomobject]@{
            request=$request
            master_row=[int]$target.row
            title=Get-CellValue $target.cells $schema.PlatformTitle.Index
            product_id=Get-CellValue $target.cells $schema.ProductId.Index
            units_per_pack=[double]$unitsPerPack
            old_baseline=Convert-ToNumber (Get-CellValue $target.cells $schema.BaselineUnitCost.Index)
            old_moving=Convert-ToNumber (Get-CellValue $target.cells $schema.MovingAverageCost.Index)
            old_book_quantity=Convert-ToNumber (Get-CellValue $target.cells $schema.CostBookQuantity.Index)
        })
    }

    $exactMatches = [Collections.Generic.List[string]]::new()
    $alreadyApplied = $false
    foreach ($target in $targets) {
        $request = $target.request
        $sameDateCode = [Collections.Generic.List[object]]::new()
        for ($i = 1; $i -lt $replenishmentRows.Count; $i++) {
            $code = ([string](Get-CellValue $replenishmentRows[$i] 1)).Trim().ToLowerInvariant()
            $date = Convert-ToDate (Get-CellValue $replenishmentRows[$i] 0)
            if ($code -eq $request.code -and $null -ne $date -and $date -eq $request.date.Date) {
                $sameDateCode.Add([pscustomobject]@{row=1+$i;quantity=Convert-ToNumber (Get-CellValue $replenishmentRows[$i] 2);amount=Convert-ToNumber (Get-CellValue $replenishmentRows[$i] 5)})
            }
        }
        foreach ($existing in $sameDateCode) {
            if ([math]::Abs([double]$existing.quantity - [double]$request.quantity) -le 0.0001 -and [math]::Abs([double]$existing.amount - [double]$request.amount) -le 0.0001) {
                $exactMatches.Add($request.code)
                $target | Add-Member -NotePropertyName replenishment_row -NotePropertyValue ([int]$existing.row)
            } else {
                throw "$($request.code) 在 $($request.date.ToString('yyyy-MM-dd')) 已有不同数量或金额的补货记录（第$($existing.row)行），已停止写入"
            }
        }
    }
    if ($exactMatches.Count -gt 0) {
        if ($exactMatches.Count -eq $requests.Count) {
            $alreadyApplied = $true
        } else {
            throw "只发现部分补货已经存在（$($exactMatches -join '、')），为避免部分重复，本次未写入"
        }
    }

    if (-not $alreadyApplied) {
        $emptyRows = [Collections.Generic.List[int]]::new()
        for ($i = 1; $i -lt 1001; $i++) {
            $row = if ($i -lt $replenishmentRows.Count) { $replenishmentRows[$i] } else { $null }
            $hasValue = $false
            foreach ($cell in @($row)) { if ($null -ne $cell -and $null -ne $cell.value -and [string]$cell.value -ne '') { $hasValue = $true; break } }
            if (-not $hasValue) { $emptyRows.Add(1 + $i) }
            if ($emptyRows.Count -eq $requests.Count) { break }
        }
        if ($emptyRows.Count -ne $requests.Count) { throw '补货记录没有足够的完整空行' }

        $operations = [Collections.Generic.List[object]]::new()
        for ($index = 0; $index -lt $targets.Count; $index++) {
            $target = $targets[$index]
            $request = $target.request
            $rowNumber = $emptyRows[$index]
            Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '补货记录' "A${rowNumber}:T${rowNumber}" $larkCli
            $row = @(
            (New-ValueCell (Convert-ToExcelSerial $request.date)),
            (New-ValueCell $request.code),
            (New-ValueCell ([double]$request.quantity)),
            (New-ValueCell ([double]$target.units_per_pack)),
            (New-FormulaCell "=IF(OR(C${rowNumber}=`"`",D${rowNumber}=`"`"),`"`",C${rowNumber}*D${rowNumber})"),
            (New-ValueCell ([double]$request.amount)),
            (New-FormulaCell "=IF(OR(E${rowNumber}=`"`",F${rowNumber}=`"`"),`"`",F${rowNumber}/E${rowNumber})"),
            @{},@{},@{},@{},
            @{},@{},
            (New-ValueCell '历史漏登补货；仅补采购成本记录，未补记付款现金流'),
            (New-ValueCell $target.product_id),
            @{},@{},@{},@{},
            (New-ValueCell ("MANUAL-{0}-{1}" -f $request.date.ToString('yyyyMMdd'),$request.code.ToUpperInvariant()))
            )
            $matrix = @(); $matrix += ,$row
            $operations.Add(@{shortcut='+cells-set';input=@{sheet_name='补货记录';range="A${rowNumber}:T${rowNumber}";cells=$matrix}})
            $target | Add-Member -NotePropertyName replenishment_row -NotePropertyValue $rowNumber
        }
        $payload = ConvertTo-Json -InputObject $operations.ToArray() -Depth 12 -Compress
        $null = Invoke-LarkJson {
            $payload | & $larkCli sheets +batch-update --as user --spreadsheet-token $sourceToken --operations - --yes --format json
        } '写入历史补货失败'
    }

    $costRaw = Invoke-Child (Join-Path $agentRoot 'rebuild_cost_snapshots.ps1') @('-ProjectRoot',$projectRoot)
    $ledgerRaw = Invoke-Child (Join-Path $agentRoot 'sync_daily_ledger.ps1') @('-ThroughDate',[datetime]::Today.ToString('yyyy-MM-dd'),'-ProjectRoot',$projectRoot)

    $replenishmentAfter = Get-Rows '补货记录' 'A1:T1001'
    $masterAfter = Get-Rows 'SKU成本主表' 'A1:Z3000'
    $itemRowsAfter = Get-Rows '订单商品' 'A1:Q4001'
    $itemsAfter = Get-ItemSnapshot $itemRowsAfter
    $targetByCode = @{}
    foreach ($target in $targets) { $targetByCode[$target.request.code] = $target }

    foreach ($target in $targets) {
        $request = $target.request
        $row = $replenishmentAfter[$target.replenishment_row - 1]
        $actualDate = Convert-ToDate (Get-CellValue $row 0)
        $actualCode = ([string](Get-CellValue $row 1)).Trim().ToLowerInvariant()
        $actualQuantity = Convert-ToNumber (Get-CellValue $row 2)
        $actualAmount = Convert-ToNumber (Get-CellValue $row 5)
        $actualUnitCost = Convert-ToNumber (Get-CellValue $row 6)
        $expectedUnitCost = [double]$request.amount / ([double]$request.quantity * [double]$target.units_per_pack)
        if ($actualDate -ne $request.date.Date -or $actualCode -ne $request.code -or
            [math]::Abs([double]$actualQuantity - [double]$request.quantity) -gt 0.0001 -or
            [math]::Abs([double]$actualAmount - [double]$request.amount) -gt 0.0001 -or
            [math]::Abs([double]$actualUnitCost - $expectedUnitCost) -gt 0.0001) {
            throw "$($request.code) 补货记录回读不一致"
        }
        $masterRow = $masterAfter[$target.master_row - 1]
        $target | Add-Member -NotePropertyName new_baseline -NotePropertyValue (Convert-ToNumber (Get-CellValue $masterRow $schema.BaselineUnitCost.Index))
        $target | Add-Member -NotePropertyName new_moving -NotePropertyValue (Convert-ToNumber (Get-CellValue $masterRow $schema.MovingAverageCost.Index))
        $target | Add-Member -NotePropertyName new_book_quantity -NotePropertyValue (Convert-ToNumber (Get-CellValue $masterRow $schema.CostBookQuantity.Index))
        if ($null -eq $target.new_moving -or $null -eq $target.new_book_quantity) { throw "$($request.code) 的成本主表重算结果为空" }
    }

    $changes = [Collections.Generic.List[object]]::new()
    $unrelatedChanges = [Collections.Generic.List[string]]::new()
    foreach ($rowNumber in $itemsAfter.Keys) {
        if (-not $itemsBefore.ContainsKey($rowNumber)) { continue }
        $before = $itemsBefore[$rowNumber]
        $after = $itemsAfter[$rowNumber]
        if (-not (Test-NumberChanged $before.unit_cost $after.unit_cost) -and -not (Test-NumberChanged $before.item_cost $after.item_cost)) { continue }
        if (-not $targetByCode.ContainsKey($after.code)) {
            $unrelatedChanges.Add("第${rowNumber}行/$($after.order_no)/$($after.code)")
            continue
        }
        $target = $targetByCode[$after.code]
        if ($null -eq $after.date -or $after.date -lt $target.request.date.Date) {
            throw "$($after.code) 在补货日期之前的订单成本发生变化：$($after.order_no)"
        }
        $changes.Add([pscustomobject]@{
            row=$rowNumber;date=$after.date.ToString('yyyy-MM-dd');order_no=$after.order_no;code=$after.code;title=$after.title;quantity=$after.quantity
            old_unit_cost=$before.unit_cost;new_unit_cost=$after.unit_cost;old_item_cost=$before.item_cost;new_item_cost=$after.item_cost
        })
    }
    if ($unrelatedChanges.Count -gt 0) { throw "成本重算改动了本次范围外的订单：$(@($unrelatedChanges | Select-Object -First 5) -join '、')" }

    foreach ($check in @(
        @{sheet='补货记录';range='D2:K1001'},
        @{sheet='SKU成本主表';range='F2:H3000'},
        @{sheet='订单商品';range='L2:N4001'},
        @{sheet='订单汇总';range='K2:L2001'}
    )) {
        $verify = Invoke-LarkJson {
            & $larkCli sheets +formula-verify --as user --spreadsheet-token $sourceToken --sheet-name $check.sheet --range $check.range --format json --exit-on-error
        } "公式校验失败：$($check.sheet)"
        if ($verify.data.status -ne 'success') { throw "公式校验未通过：$($check.sheet)" }
    }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok=$true
        transaction_id=$transaction.id
        atomic=$true
        replenishments_already_present=$alreadyApplied
        payment_cashflow_recorded=$false
        replenishments=@($targets | ForEach-Object {
            [ordered]@{
                code=$_.request.code;date=$_.request.date.ToString('yyyy-MM-dd');quantity=$_.request.quantity;amount=$_.request.amount
                purchase_unit_cost=[math]::Round($_.request.amount/($_.request.quantity*$_.units_per_pack),6)
                replenishment_row=$_.replenishment_row;sku_master_row=$_.master_row;title=$_.title
                old_baseline=$_.old_baseline;new_baseline=$_.new_baseline;old_moving=$_.old_moving;new_moving=$_.new_moving
                old_book_quantity=$_.old_book_quantity;new_book_quantity=$_.new_book_quantity
            }
        })
        changed_order_items=@($changes | Sort-Object date,order_no,code)
        changed_order_count=@($changes | Select-Object -ExpandProperty order_no -Unique).Count
        unrelated_order_cost_changes=0
        formula_verification='success'
        cost_rebuild=($costRaw | ConvertFrom-Json)
        ledger_sync=($ledgerRaw | ConvertFrom-Json)
    } | ConvertTo-Json -Depth 10
} catch {
    $reason = [string]$_.Exception.Message
    $rollback = Invoke-FinanceTransactionRollback $transaction $tokens $larkCli $reason
    if ($rollback.ok) { throw "$reason；本次4笔补货及其派生改动已全部回滚" }
    throw "$reason；自动回滚未完整完成：$($rollback.error)"
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
