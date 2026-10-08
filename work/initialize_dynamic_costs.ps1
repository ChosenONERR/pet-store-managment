param(
    [string]$ProjectRoot,
    [string]$AuditPath
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
if (-not $AuditPath) { $AuditPath = Join-Path $ProjectRoot 'work\cost_audit.json' }
$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少经营驾驶舱元数据' }
if (-not (Test-Path -LiteralPath $AuditPath)) { throw '缺少成本审计结果' }
$meta = [System.IO.File]::ReadAllText($metaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$audit = [System.IO.File]::ReadAllText($AuditPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$token = ([uri]$meta.workbook_url).Segments[-1].Trim('/')

function New-ValueCell($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return @{} }
    return @{value=$Value}
}
function New-FormulaCell([string]$Formula) {
    if (-not $Formula) { return @{} }
    return @{formula=$Formula}
}
function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}
function Convert-ToExcelSerial([datetime]$Value) { return ($Value.Date - [datetime]'1899-12-30').TotalDays }
function Get-SheetRows([string]$SheetName, [string]$Range) {
    $raw = lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --include value,formula --max-chars 2000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取飞书失败：$SheetName" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "飞书读取结果不完整：$SheetName" }
    return @($result.data.ranges[0].cells)
}
function Set-CellMatrix([string]$SheetName, [string]$Range, $Rows) {
    if (@($Rows).Count -eq 0) { return }
    $payload = ConvertTo-Json -InputObject $Rows -Depth 12 -Compress
    $raw = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入飞书失败：$SheetName $Range" }
    if (-not (($raw | ConvertFrom-Json).ok)) { throw "写入飞书未成功：$SheetName $Range" }
}
function Clear-Range([string]$Range) {
    $ranges = ConvertTo-Json -InputObject @($Range) -Compress
    $raw = $ranges | lark-cli sheets +cells-batch-clear --as user --spreadsheet-token $token --ranges - --scope content --yes --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not (($raw | ConvertFrom-Json).ok)) { throw "清理范围失败：$Range" }
}
function Assert-FormulaRange([string]$SheetName, [string]$Range) {
    $raw = lark-cli sheets +formula-verify --as user --spreadsheet-token $token --sheet-name $SheetName --range $Range --format json --exit-on-error | Out-String
    if ($LASTEXITCODE -ne 0) { throw "公式校验失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if ($result.data.status -ne 'success') { throw "公式存在错误：$SheetName $Range" }
}

$inventoryDateMatch = [regex]::Match([string]$audit.inventory_file, '(\d{4}-\d{2}-\d{2})')
if (-not $inventoryDateMatch.Success) { throw '无法从库存文件名识别快照日期' }
$snapshotDate = [datetime]::ParseExact($inventoryDateMatch.Groups[1].Value, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
$baselineDate = $snapshotDate.AddDays(1)

$workbookRaw = lark-cli sheets +workbook-info --as user --spreadsheet-token $token --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '读取工作簿结构失败' }
$workbook = $workbookRaw | ConvertFrom-Json
$sheetNames = @($workbook.data.sheets | ForEach-Object { $_.sheet_name })
if ('SKU成本主表' -notin $sheetNames) {
    $sheetPayload = @{sheets=@(@{
        name='SKU成本主表'
        columns=@('KW货号','CSX货架码','商品标题','初始化库存','基准单位成本','当前移动平均成本','成本账数量','每采购包装含销售单位数','采购单位','成本核对状态','初始化日期','最近成本更新时间','换算依据','备注')
        data=@()
        dtypes=@{'初始化库存'='float64';'基准单位成本'='float64';'当前移动平均成本'='float64';'成本账数量'='float64';'每采购包装含销售单位数'='float64';'初始化日期'='datetime64[ns]';'最近成本更新时间'='datetime64[ns]'}
        formats=@{'初始化库存'='#,##0.##';'基准单位成本'='#,##0.0000';'当前移动平均成本'='#,##0.0000';'成本账数量'='#,##0.##';'每采购包装含销售单位数'='#,##0.##';'初始化日期'='yyyy-mm-dd';'最近成本更新时间'='yyyy-mm-dd hh:mm'}
    })} | ConvertTo-Json -Depth 20 -Compress
    $styles = @{styles=@(@{
        name='SKU成本主表'
        cell_styles=@(
            @{range='A1:N1';background_color='#244C3A';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center'},
            @{range='A2:N1500';background_color='#F7FAF8'},
            @{range='F2:H1500';background_color='#E8F1EC'},
            @{range='J2:J1500';background_color='#FFF3CD'},
            @{range='N2:N1500';background_color='#FFF8E1'}
        )
        col_sizes=@(
            @{range='A:B';type='pixel';size=110},@{range='C';type='pixel';size=360},@{range='D:H';type='pixel';size=135},
            @{range='I:J';type='pixel';size=130},@{range='K:L';type='pixel';size=145},@{range='M:N';type='pixel';size=300}
        )
    })} | ConvertTo-Json -Depth 20 -Compress
    $raw = $sheetPayload | lark-cli sheets +table-put --as user --spreadsheet-token $token --sheets - --styles $styles --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not (($raw | ConvertFrom-Json).ok)) { throw '创建SKU成本主表失败' }
    $null = lark-cli sheets +dim-freeze --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --dimension row --count 1 --format json | Out-String
    $filter = @{range='A1:N1500';rules=@()} | ConvertTo-Json -Compress
    $null = $filter | lark-cli sheets +filter-create --as user --spreadsheet-token $token --sheet-name 'SKU成本主表' --range 'A1:N1500' --properties - --format json | Out-String
}

$masterRows = @()
$auditByCsx = @{}
$validAuditRows = @($audit.rows | Where-Object { [string]$_.kw_code -match '^(?i)kw\d+$' })
foreach ($entry in $validAuditRows) {
    $csx = ([string]$entry.csx_code).ToLowerInvariant()
    if ($csx) {
        if (-not $auditByCsx.ContainsKey($csx)) { $auditByCsx[$csx] = @() }
        $auditByCsx[$csx] += $entry
    }
}

$kwGroups = @($validAuditRows | Group-Object { ([string]$_.kw_code).ToLowerInvariant() } | Sort-Object Name)
foreach ($group in $kwGroups) {
    $entries = @($group.Group)
    $kw = [string]$group.Name
    $csxValues = @($entries | ForEach-Object { ([string]$_.csx_code).ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
    $titleValues = @($entries | ForEach-Object { [string]$_.title } | Where-Object { $_ } | Select-Object -Unique)
    $purchaseUnitValues = @($entries | ForEach-Object { [string]$_.purchase_unit } | Where-Object { $_ } | Select-Object -Unique)
    $statuses = @($entries | ForEach-Object { [string]$_.audit_status } | Select-Object -Unique)
    $usableEntries = @($entries | Where-Object { $_.audit_status -in @('已自动换算','默认整件') -and $null -ne $_.baseline_unit_cost })
    $costValues = @($usableEntries | ForEach-Object { [double]$_.baseline_unit_cost } | Sort-Object -Unique)
    $unitValues = @($usableEntries | ForEach-Object { [double]$_.units_per_pack } | Sort-Object -Unique)
    $mappingConflict = ($csxValues.Count -gt 1 -or $titleValues.Count -gt 1)
    $costConflict = ($costValues.Count -gt 1 -or $unitValues.Count -gt 1)

    $status = '默认整件'
    if ($mappingConflict -or $costConflict) { $status = '待人工核对' }
    elseif ('缺少CSX映射' -in $statuses) { $status = '缺少CSX映射' }
    elseif ('缺少成本' -in $statuses) { $status = '缺少成本' }
    elseif ('待人工核对' -in $statuses) { $status = '待人工核对' }
    elseif ('已自动换算' -in $statuses) { $status = '已自动换算' }

    $usableCost = if ($status -in @('已自动换算','默认整件') -and $costValues.Count -eq 1) { [double]$costValues[0] } else { $null }
    $unitsPerPack = if ($status -eq '已自动换算' -and $unitValues.Count -eq 1) { [double]$unitValues[0] } else { 1.0 }
    $stock = [double](($entries | ForEach-Object { if ($null -eq $_.inventory) { 0.0 } else { [double]$_.inventory } } | Measure-Object -Sum).Sum)
    $reasonParts = @()
    if ($entries.Count -gt 1) { $reasonParts += "同一KW合并$($entries.Count)个美团规格行，初始化库存已合计" }
    if ($mappingConflict) { $reasonParts += '同一KW对应多个CSX或多个商品标题，需人工确认' }
    elseif ($costConflict) { $reasonParts += '同一KW出现多个成本或包装换算，需人工确认' }
    else { $reasonParts += [string]$entries[0].conversion_reason }
    $noteParts = @()
    if ($entries.Count -gt 1) { $noteParts += "已合并$($entries.Count)行库存" }
    if ($status -notin @('已自动换算','默认整件')) { $noteParts += '成本未确认前不计入自然单利润' }
    $masterRows += ,@(
        (New-ValueCell $kw),(New-ValueCell ($csxValues -join '/')),
        (New-ValueCell $(if ($titleValues.Count -gt 0) { $titleValues[0] } else { $null })),(New-ValueCell $stock),(New-ValueCell $usableCost),(New-ValueCell $usableCost),(New-ValueCell $stock),
        (New-ValueCell $unitsPerPack),(New-ValueCell $(if ($purchaseUnitValues.Count -gt 0) { $purchaseUnitValues[0] } else { $null })),(New-ValueCell $status),
        (New-ValueCell (Convert-ToExcelSerial $baselineDate)),(New-ValueCell (Convert-ToExcelSerial ([datetime]::Today))),
        (New-ValueCell ($reasonParts -join '；')),(New-ValueCell ($noteParts -join '；'))
    )
}
if ($masterRows.Count -gt 1499) { throw 'SKU成本主表容量不足' }
Clear-Range 'SKU成本主表!A2:N1500'
Set-CellMatrix 'SKU成本主表' "A2:N$($masterRows.Count + 1)" $masterRows

$costRows = Get-SheetRows '商品成本' 'A2:F2000'
$costSplitCells = @()
$costStatusCells = @()
for ($i = 0; $i -lt $costRows.Count; $i++) {
    $code = ([string](Get-CellValue $costRows[$i] 0)).ToLowerInvariant()
    $split = $null
    $status = '默认整件'
    if ($code -in @('csx188','csx189','csx576')) { $status = '待人工核对' }
    if ($auditByCsx.ContainsKey($code)) {
        $entries = @($auditByCsx[$code])
        $auto = @($entries | Where-Object { $_.audit_status -eq '已自动换算' } | Select-Object -First 1)
        $pending = @($entries | Where-Object { $_.audit_status -in @('待人工核对','缺少成本') } | Select-Object -First 1)
        if ($auto.Count -gt 0) {
            $split = [double]$auto[0].units_per_pack
            $status = '自动拆分成本'
        } elseif ($pending.Count -gt 0) {
            $status = '待人工核对'
        }
    }
    $costSplitCells += ,@((New-ValueCell $split))
    $costStatusCells += ,@((New-ValueCell $status))
}
if ($costRows.Count -gt 0) {
    Set-CellMatrix '商品成本' "C2:C$($costRows.Count + 1)" $costSplitCells
    Set-CellMatrix '商品成本' "E2:E$($costRows.Count + 1)" $costStatusCells
}

$existingReplenishment = Get-SheetRows '补货记录' 'A2:N1001'
$nonblankReplenishment = @($existingReplenishment | Where-Object { Get-CellValue $_ 0 -or Get-CellValue $_ 1 -or Get-CellValue $_ 2 -or Get-CellValue $_ 4 })
if ($nonblankReplenishment.Count -gt 0) { throw '补货记录已有数据，不能自动重构；请先人工迁移' }
Clear-Range '补货记录!A1:N1001'
$headers = @('补货日期','货号','采购包装数量','每包装销售单位数','入库销售单位数量','实际采购金额','本次单位成本','补货前数量','补货前平均成本','补货后数量','补货后平均成本','付款日期','付款金额','备注')
$headerCells = @(); foreach ($header in $headers) { $headerCells += ,(New-ValueCell $header) }
$headerMatrix = @(); $headerMatrix += ,$headerCells
Set-CellMatrix '补货记录' 'A1:N1' $headerMatrix

$dCells=@();$eCells=@();$gCells=@()
for ($row=2; $row -le 1001; $row++) {
    $dCells += ,@((New-FormulaCell "=IF(B$row=`"`",`"`",IFERROR(VLOOKUP(B$row,'SKU成本主表'!`$A`$2:`$H`$1500,8,FALSE),IFERROR(VLOOKUP(B$row,'SKU成本主表'!`$B`$2:`$H`$1500,7,FALSE),`"`")))"))
    $eCells += ,@((New-FormulaCell "=IF(OR(C$row=`"`",D$row=`"`"),`"`",C$row*D$row)"))
    $gCells += ,@((New-FormulaCell "=IF(OR(E$row=`"`",F$row=`"`"),`"`",F$row/E$row)"))
}
Set-CellMatrix '补货记录' 'D2:D1001' $dCells
Set-CellMatrix '补货记录' 'E2:E1001' $eCells
Set-CellMatrix '补货记录' 'G2:G1001' $gCells

$styleCommands = @(
    @('A1:N1','#8A5A13','#FFFFFF','bold','center'),
    @('A2:N1001','#FFF8E1',$null,$null,$null),
    @('D2:E1001','#E8F1EC',$null,$null,$null),
    @('G2:K1001','#E8F1EC',$null,$null,$null)
)
foreach ($style in $styleCommands) {
    $arguments = @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$token,'--sheet-name','补货记录','--range',$style[0],'--background-color',$style[1],'--format','json')
    if ($style[2]) { $arguments += @('--font-color',$style[2]) }
    if ($style[3]) { $arguments += @('--font-weight',$style[3]) }
    if ($style[4]) { $arguments += @('--horizontal-alignment',$style[4]) }
    $null = & lark-cli @arguments | Out-String
    if ($LASTEXITCODE -ne 0) { throw "更新补货记录样式失败：$($style[0])" }
}
# 日期按真实日期序列存储，必须单独设置显示格式；否则飞书会直接显示 46277 这类序列号。
foreach ($dateRange in @('A2:A1001','L2:L1001','P2:P1001','S2:S1001')) {
    $null = lark-cli sheets +cells-set-style --as user --spreadsheet-token $token --sheet-name '补货记录' --range $dateRange --number-format 'yyyy-mm-dd' --horizontal-alignment center --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "更新补货记录日期格式失败：$dateRange" }
}
$widths = '{"A:B":115,"C:K":135,"L:M":120,"N":260}'
$null = lark-cli sheets +cols-resize --as user --spreadsheet-token $token --sheet-name '补货记录' --widths $widths --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '更新补货记录列宽失败' }
$null = lark-cli sheets +dim-freeze --as user --spreadsheet-token $token --sheet-name '补货记录' --dimension row --count 1 --format json | Out-String

Assert-FormulaRange '商品成本' 'D2:D2000'
Assert-FormulaRange '补货记录' 'D2:E1001,G2:G1001'

$readback = Get-SheetRows 'SKU成本主表' "A2:N$($masterRows.Count + 1)"
if ($readback.Count -ne $masterRows.Count) { throw 'SKU成本主表回读行数不一致' }
$csx104 = @($readback | Where-Object { (Get-CellValue $_ 1) -eq 'csx104' } | Select-Object -First 1)
if ($csx104.Count -ne 1 -or [math]::Abs([double](Get-CellValue $csx104[0] 4) - 1.25) -gt 0.0001) { throw 'csx104/kw2054成本抽查失败' }

[ordered]@{
    ok=$true
    master_rows=$masterRows.Count
    merged_variant_rows=($validAuditRows.Count - $masterRows.Count)
    malformed_kw_rows=@($audit.rows | Where-Object { [string]$_.kw_code -match '^(?i)kw' -and [string]$_.kw_code -notmatch '^(?i)kw\d+$' }).Count
    baseline_date=$baselineDate.ToString('yyyy-MM-dd')
    automatic_conversions=@($audit.rows | Where-Object { $_.audit_status -eq '已自动换算' -and $_.kw_code -match '^(?i)kw\d+$' }).Count
    pending_review=@($audit.rows | Where-Object { $_.audit_status -eq '待人工核对' -and $_.kw_code -match '^(?i)kw\d+$' }).Count
    missing_cost=@($audit.rows | Where-Object { $_.audit_status -eq '缺少成本' -and $_.kw_code -match '^(?i)kw\d+$' }).Count
    missing_mapping=@($audit.rows | Where-Object { $_.audit_status -eq '缺少CSX映射' -and $_.kw_code -match '^(?i)kw\d+$' }).Count
    csx104_unit_cost=1.25
} | ConvertTo-Json -Compress
