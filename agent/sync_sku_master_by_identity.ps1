param(
    [string]$ProjectRoot,
    [datetime]$SnapshotDate = [datetime]::Today
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:LARK_CLI_NO_PROXY = '1'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
$stateRoot = Join-Path $ProjectRoot 'agent\state'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少飞书经营驾驶舱元数据' }
if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
$meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $PSScriptRoot 'finance_transaction.ps1')
. (Join-Path $PSScriptRoot 'sku_master_schema.ps1')
$skuSchema = Get-SkuMasterSchema -SpreadsheetToken $spreadsheetToken
$script:financeTransaction = Join-FinanceTransaction $stateRoot
$script:financeTransactionOwner = $false
$script:financeTransactionClosed = $false
if ($null -eq $script:financeTransaction) {
    $script:financeTransaction = New-FinanceTransaction $stateRoot "按商品身份同步SKU主表 $($SnapshotDate.ToString('yyyy-MM-dd'))" $null -AcquireLock -WorkbookTokens @{source=$spreadsheetToken} -LarkCli 'lark-cli'
    $script:financeTransactionOwner = $true
}
function Close-SkuSyncTransaction([bool]$Succeeded, [string]$Reason = '') {
    if ($script:financeTransactionClosed) { return }
    if ($script:financeTransactionOwner) {
        if ($Succeeded) { Complete-FinanceTransaction $script:financeTransaction }
        else { $null = Invoke-FinanceTransactionRollback $script:financeTransaction @{source=$spreadsheetToken} 'lark-cli' $Reason }
    }
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
}
trap { $reason=[string]$_.Exception.Message; Close-SkuSyncTransaction $false $reason; [Console]::Error.WriteLine($reason); exit 1 }

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count -or $null -eq @($Row)[$Index]) { return $null }
    $valueProperty = @($Row)[$Index].PSObject.Properties['value']
    if ($null -eq $valueProperty) { return $null }
    $value = $valueProperty.Value
    if ($null -eq $value -or [string]$value -eq '') { return $null }
    return ([string]$value).Trim()
}
function Convert-ToNullableNumber($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}
function New-ValueCell($Value) { if ($null -eq $Value -or [string]$Value -eq '') { return @{value=''} }; return @{value=$Value} }
function Normalize-Code([string]$Value) { if (-not $Value) { return $null }; return $Value.Trim().ToLowerInvariant() }
function Get-TransitionCodes([string]$Value) {
    if (-not $Value) { return @() }
    return @([regex]::Split($Value, '[/,，、;；\s]+') | Where-Object { $_ -and $_ -notmatch '^(?i)kw\d+$' } | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique)
}
function Invoke-LarkJson([scriptblock]$Command, [string]$Label, [int]$MaxAttempts = 3) {
    $errors = [Collections.Generic.List[string]]::new()
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $global:LASTEXITCODE = 0
        try {
            $raw = & $Command | Out-String
            $result = $raw | ConvertFrom-Json
            if ($LASTEXITCODE -eq 0 -and $result.ok -ne $false) { return $result }
            $errors.Add("第${attempt}次：退出码=$LASTEXITCODE")
        } catch { $errors.Add("第${attempt}次：$($_.Exception.Message)") }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds ([math]::Min(6, $attempt * 2)) }
    }
    throw "$Label；$($errors -join '；')"
}
function Set-CellMatrix([string]$Range, $Rows) {
    $rowValues = @($Rows)
    $matrix = [Collections.ArrayList]::new()
    if ($rowValues.Count -gt 0 -and $rowValues[0] -is [Collections.IDictionary]) {
        $null = $matrix.Add([object[]]$rowValues)
    } else {
        foreach ($row in $rowValues) { $null = $matrix.Add([object[]]@($row)) }
    }
    Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $spreadsheetToken 'SKU成本主表' $Range 'lark-cli'
    $payload = ConvertTo-Json -InputObject $matrix -Depth 12 -Compress
    $null = Invoke-LarkJson { $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range $Range --cells - --format json } "写入SKU成本主表失败：$Range"
}

$masterResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'A2:Z3000' --include value --max-chars 1500000 --format json } '读取SKU成本主表失败'
$inventoryResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name '库存快照' --range 'A2:O2000' --include value --max-chars 1000000 --format json } '读取库存快照失败'
$masterRows = @($masterResult.data.ranges[0].cells)
$masterRowNumbers = @($masterResult.data.ranges[0].row_indices)
$inventoryRows = @($inventoryResult.data.ranges[0].cells)

$backupPath = Join-Path $stateRoot ("sku_cost_master_auto_backup_{0}.json" -f [datetime]::Now.ToString('yyyyMMdd_HHmmss'))
[ordered]@{source='SKU成本主表!A2:Z3000';captured_at=[datetime]::Now.ToString('o');row_indices=$masterRowNumbers;cells=$masterRows} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $backupPath -Encoding UTF8

$master = [Collections.Generic.List[object]]::new()
$bySkuId = @{}
$masterByRow = @{}
$lastUsedRow = 1
for ($i=0; $i -lt $masterRows.Count; $i++) {
    $row = $masterRows[$i]
    $rowNumber = if ($i -lt $masterRowNumbers.Count) { [int]$masterRowNumbers[$i] } else { 2 + $i }
    $entry = [pscustomobject]@{row_number=$rowNumber;cells=$row;kw=(Normalize-Code (Get-CellValue $row $skuSchema.CurrentKw.Index));title=(Get-CellValue $row $skuSchema.PlatformTitle.Index);sku_id=(Get-CellValue $row $skuSchema.ProductId.Index)}
    if ($entry.kw -or (Get-CellValue $row 1) -or $entry.title -or $entry.sku_id) { $lastUsedRow=[math]::Max($lastUsedRow,$rowNumber) }
    $master.Add($entry)
    $masterByRow[$rowNumber] = $entry
    if ($entry.sku_id) {
        if (-not $bySkuId.ContainsKey($entry.sku_id)) { $bySkuId[$entry.sku_id]=[Collections.Generic.List[object]]::new() }
        $bySkuId[$entry.sku_id].Add($entry)
    }
}

$updates = [Collections.Generic.List[object]]::new()
$newRows = [Collections.Generic.List[object]]::new()
$seenSkuIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$assignedRows = [Collections.Generic.HashSet[int]]::new()
$serial = ($SnapshotDate.Date - [datetime]'1899-12-30').TotalDays

foreach ($row in $inventoryRows) {
    $skuId = Get-CellValue $row 13
    if (-not $skuId -or $skuId -like 'missing:*') { continue }
    $null = $seenSkuIds.Add($skuId)
    $rawGoodsCode = Normalize-Code (Get-CellValue $row 1)
    $kw = if ($rawGoodsCode -match '^(?i)kw\d+$') { $rawGoodsCode } else { $null }
    $title = Get-CellValue $row 2
    $stock = Convert-ToNullableNumber (Get-CellValue $row 3)
    $status = Get-CellValue $row 5
    $shelfCode = Normalize-Code (Get-CellValue $row 6)
    $transition = @($(if($rawGoodsCode -and $rawGoodsCode -notmatch '^(?i)kw\d+$'){$rawGoodsCode}),$(if($shelfCode -and $shelfCode -notmatch '^(?i)kw\d+$'){$shelfCode})) | Where-Object { $_ } | Select-Object -Unique
    $csx = $transition -join '、'
    $spec = Get-CellValue $row 14
    $target = $null
    if ($bySkuId.ContainsKey($skuId) -and @($bySkuId[$skuId]).Count -eq 1) { $target = @($bySkuId[$skuId])[0] }
    if ($null -eq $target) {
        $candidates = @($master | Where-Object {
            -not $_.sku_id -and -not $assignedRows.Contains([int]$_.row_number) -and (
                ($kw -and $_.kw -eq $kw -and (Get-CellValue $_.cells $skuSchema.PlatformStatus.Index) -eq '待平台同步') -or
                ($kw -and $_.kw -eq $kw -and (-not $title -or $_.title -eq $title)) -or
                ($title -and $_.title -eq $title -and $transition.Count -gt 0 -and @($transition | Where-Object { $_ -in (Get-TransitionCodes (Get-CellValue $_.cells 1)) }).Count -gt 0)
            )
        })
        if ($candidates.Count -eq 1) { $target = $candidates[0] }
    }
    if ($null -ne $target) {
        $null = $assignedRows.Add([int]$target.row_number)
        $targetNote = Get-CellValue $target.cells $skuSchema.Note.Index
        $hasOwnerConfirmedKw = $target.kw -and $targetNote -match '店主确认当前KW|当前KW由店主确认'
        # 店主通过对话确认的当前 KW 优先于平台旧值；否则平台明确返回 KW 时才更新。
        # 平台仍使用 CSX/过渡货号时，也保留主表已有的当前 KW。
        $effectiveKw = if ($hasOwnerConfirmedKw) { $target.kw } elseif ($kw) { $kw } else { $target.kw }
        $updates.Add([pscustomobject]@{row=[int]$target.row_number;kw=$effectiveKw;csx=$csx;title=$title;sku_id=$skuId;status=$status;spec=$spec})
        continue
    }
    $newRow = @(); for ($columnIndex = 0; $columnIndex -lt 26; $columnIndex++) { $newRow += ,(New-ValueCell $null) }
    $newRow[$skuSchema.CurrentKw.Index] = New-ValueCell $kw
    $newRow[$skuSchema.TransitionCode.Index] = New-ValueCell $csx
    $newRow[$skuSchema.PlatformTitle.Index] = New-ValueCell $title
    $newRow[$skuSchema.InitialInventory.Index] = New-ValueCell $stock
    $newRow[$skuSchema.CostBookQuantity.Index] = New-ValueCell $stock
    $newRow[$skuSchema.UnitsPerPack.Index] = New-ValueCell 1
    $newRow[$skuSchema.PurchaseUnit.Index] = New-ValueCell $spec
    $newRow[$skuSchema.CostStatus.Index] = New-ValueCell '缺少成本'
    $newRow[$skuSchema.InitializationDate.Index] = New-ValueCell $serial
    $newRow[$skuSchema.ConversionBasis.Index] = New-ValueCell '美团库存同步自动建档；成本与包装换算待核对'
    $newRow[$skuSchema.Note.Index] = New-ValueCell '新商品自动建档，请补充单位成本'
    $newRow[$skuSchema.ProductId.Index] = New-ValueCell $skuId
    $newRow[$skuSchema.PlatformStatus.Index] = New-ValueCell $status
    $newRow[$skuSchema.CurrentSpec.Index] = New-ValueCell $spec
    $newRows.Add($newRow)
}

$kwOwner = @{}
foreach ($update in $updates) { if ($update.kw) { $kwOwner[$update.kw] = [string]$update.sku_id } }
$clearKwRows = [Collections.Generic.HashSet[int]]::new()
foreach ($entry in $master) {
    if (-not $entry.kw -or -not $entry.sku_id -or -not $kwOwner.ContainsKey($entry.kw)) { continue }
    if ([string]$kwOwner[$entry.kw] -ne [string]$entry.sku_id) { $null = $clearKwRows.Add([int]$entry.row_number) }
}

$updateByRow = @{}
foreach ($update in $updates) { $updateByRow[[int]$update.row] = $update }
$acRows = [Collections.Generic.List[object]]::new()
$productIdRows = [Collections.Generic.List[object]]::new()
$platformStatusRows = [Collections.Generic.List[object]]::new()
$currentSpecRows = [Collections.Generic.List[object]]::new()
for ($rowNumber=2; $rowNumber -le $lastUsedRow; $rowNumber++) {
    $entry = if ($masterByRow.ContainsKey($rowNumber)) { $masterByRow[$rowNumber] } else { $null }
    $cells = if ($null -ne $entry) { $entry.cells } else { @() }
    $currentKw = Get-CellValue $cells 0
    $currentCsx = Get-CellValue $cells 1
    $currentTitle = Get-CellValue $cells 2
    $currentSkuId = Get-CellValue $cells $skuSchema.ProductId.Index
    $currentStatus = Get-CellValue $cells $skuSchema.PlatformStatus.Index
    $currentSpec = Get-CellValue $cells $skuSchema.CurrentSpec.Index
    if ($clearKwRows.Contains($rowNumber)) { $currentKw = $null }
    if ($updateByRow.ContainsKey($rowNumber)) {
        $update = $updateByRow[$rowNumber]
        $currentKw=$update.kw; $currentCsx=$update.csx; $currentTitle=$update.title
        $currentSkuId=$update.sku_id; $currentStatus=$update.status; $currentSpec=$update.spec
    }
    $acRows.Add(@((New-ValueCell $currentKw),(New-ValueCell $currentCsx),(New-ValueCell $currentTitle)))
    $productIdRows.Add(@((New-ValueCell $currentSkuId)))
    $platformStatusRows.Add(@((New-ValueCell $currentStatus)))
    $currentSpecRows.Add(@((New-ValueCell $currentSpec)))
}
if ($lastUsedRow -ge 2) {
    Set-CellMatrix "A2:C${lastUsedRow}" $acRows.ToArray()
    Set-CellMatrix "$($skuSchema.ProductId.Column)2:$($skuSchema.ProductId.Column)${lastUsedRow}" $productIdRows.ToArray()
    Set-CellMatrix "$($skuSchema.PlatformStatus.Column)2:$($skuSchema.PlatformStatus.Column)${lastUsedRow}" $platformStatusRows.ToArray()
    Set-CellMatrix "$($skuSchema.CurrentSpec.Column)2:$($skuSchema.CurrentSpec.Column)${lastUsedRow}" $currentSpecRows.ToArray()
}

$appendStart = $lastUsedRow + 1
if ($newRows.Count -gt 0) {
    $appendEnd = $appendStart + $newRows.Count - 1
    if ($appendEnd -gt 3000) { throw 'SKU成本主表超过3000行容量' }
    Set-CellMatrix "A${appendStart}:Z${appendEnd}" $newRows.ToArray()
}

$verifyResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name 'SKU成本主表' --range 'A2:Z3000' --include value --max-chars 1500000 --format json } '回读SKU成本主表失败'
$finalRows = @($verifyResult.data.ranges[0].cells)
$finalRowNumbers = @($verifyResult.data.ranges[0].row_indices)
$verifiedIds = @($finalRows | ForEach-Object { Get-CellValue $_ $skuSchema.ProductId.Index } | Where-Object { $_ })
foreach ($skuId in $seenSkuIds) { if ($skuId -notin $verifiedIds) { throw "商品身份写入回读失败：$skuId" } }

# 同一个当前 KW 若被多个稳定商品 ID 占用，只标记冲突，不擅自合并或挑选其中一行。
$kwAssignments = @{}
for ($i=0; $i -lt $finalRows.Count; $i++) {
    $kw = Normalize-Code (Get-CellValue $finalRows[$i] 0)
    $skuId = Get-CellValue $finalRows[$i] $skuSchema.ProductId.Index
    if (-not $kw -or -not $skuId) { continue }
    $rowNumber = if ($i -lt $finalRowNumbers.Count) { [int]$finalRowNumbers[$i] } else { 2 + $i }
    if (-not $kwAssignments.ContainsKey($kw)) { $kwAssignments[$kw] = [Collections.Generic.List[object]]::new() }
    $kwAssignments[$kw].Add([pscustomobject]@{row=$rowNumber;sku_id=$skuId;cells=$finalRows[$i]})
}
$conflictCodes = [Collections.Generic.List[string]]::new()
foreach ($kw in $kwAssignments.Keys) {
    $owners = @($kwAssignments[$kw] | ForEach-Object { $_.sku_id } | Select-Object -Unique)
    if ($owners.Count -le 1) { continue }
    $conflictCodes.Add($kw)
    foreach ($entry in $kwAssignments[$kw]) {
        $oldNote = Get-CellValue $entry.cells $skuSchema.Note.Index
        $conflictNote = "当前KW $kw 同时对应 $($owners.Count) 个美团SKU ID，请在美团后台为每个规格分配唯一货号"
        $newNote = if ($oldNote -and $oldNote -notmatch '当前KW.+同时对应') { "$oldNote；$conflictNote" } elseif ($oldNote -match [regex]::Escape($conflictNote)) { $oldNote } else { $conflictNote }
        if ((Get-CellValue $entry.cells $skuSchema.CostStatus.Index) -ne '当前KW冲突待核对') {
            Set-CellMatrix "$($skuSchema.CostStatus.Column)$($entry.row):$($skuSchema.CostStatus.Column)$($entry.row)" @((New-ValueCell '当前KW冲突待核对'))
        }
        if ($newNote -ne $oldNote) {
            Set-CellMatrix "$($skuSchema.Note.Column)$($entry.row):$($skuSchema.Note.Column)$($entry.row)" @((New-ValueCell $newNote))
        }
    }
}

$finalResult = [ordered]@{
    ok=$true;snapshot_date=$SnapshotDate.ToString('yyyy-MM-dd');new_kw_count=$newRows.Count;mapped_existing_count=$updates.Count
    new_kw_codes=@($newRows | ForEach-Object {[string]$_[0].value});mapped_kw_codes=@($updates | ForEach-Object {[string]$_.kw})
    stable_product_ids=$seenSkuIds.Count;kw_conflict_count=$conflictCodes.Count;kw_conflict_codes=@($conflictCodes);formula_verification='success'
}
Close-SkuSyncTransaction $true
$finalResult | ConvertTo-Json -Depth 6 -Compress
