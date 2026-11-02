$ErrorActionPreference = 'Stop'

function Get-ZipEntryText([System.IO.Compression.ZipArchive]$Zip, [string]$Name) {
    $entry = $Zip.GetEntry($Name)
    if (-not $entry) { throw "Missing XLSX entry: $Name" }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Convert-ColumnNameToIndex([string]$Name) {
    $n = 0
    foreach ($ch in $Name.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1) }
    return $n - 1
}

function Get-CellValue($Cell, [string[]]$SharedStrings) {
    $type = [string]$Cell.t
    $raw = [string]$Cell.v
    if ($type -eq 's' -and $raw -match '^\d+$') { return $SharedStrings[[int]$raw] }
    if ($type -eq 'inlineStr') { return ([string]$Cell.is.t) }
    if ($null -eq $raw -or $raw -eq '') { return $null }
    if ($raw -match '^-?\d+(\.\d+)?$') { return [double]$raw }
    return $raw
}

function Read-XlsxRows([string]$Path, [string]$PreferredSheet) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $wb = [xml](Get-ZipEntryText $zip 'xl/workbook.xml')
        $rels = [xml](Get-ZipEntryText $zip 'xl/_rels/workbook.xml.rels')
        $shared = @()
        $sharedEntry = $zip.GetEntry('xl/sharedStrings.xml')
        if ($sharedEntry) {
            $ss = [xml](Get-ZipEntryText $zip 'xl/sharedStrings.xml')
            $shared = @($ss.SelectNodes("//*[local-name()='si']") | ForEach-Object {
                (($_.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.'#text' }) -join '')
            })
        }
        $sheetNode = @($wb.SelectNodes("//*[local-name()='sheet']") | Where-Object { ([string]$_.name) -eq $PreferredSheet } | Select-Object -First 1)
        if (-not $sheetNode) { $sheetNode = @($wb.SelectNodes("//*[local-name()='sheet']") | Select-Object -First 1) }
        $rid = [string]$sheetNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        if (-not $rid) { $rid = [string]$sheetNode.'r:id' }
        $rel = @($rels.SelectNodes("//*[local-name()='Relationship']") | Where-Object { ([string]$_.Id) -eq $rid } | Select-Object -First 1)
        $target = ([string]$rel.Target).TrimStart('/')
        if ($target -notlike 'xl/*') { $target = "xl/$target" }
        $sheet = [xml](Get-ZipEntryText $zip $target)
        $rows = @()
        foreach ($row in $sheet.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $cells = @{}
            foreach ($cell in $row.SelectNodes("./*[local-name()='c']")) {
                $ref = [string]$cell.r
                $col = [regex]::Match($ref, '^[A-Z]+').Value
                if ($col) { $cells[(Convert-ColumnNameToIndex $col)] = Get-CellValue $cell $shared }
            }
            $max = if ($cells.Count) { ($cells.Keys | Measure-Object -Maximum).Maximum } else { -1 }
            if ($max -ge 0) {
                $arr = [object[]]::new($max + 1)
                foreach ($k in $cells.Keys) { $arr[$k] = $cells[$k] }
                $rows += ,$arr
            }
        }
        return ,$rows
    } finally { $zip.Dispose() }
}

function Normalize([object]$Value) {
    if ($null -eq $Value) { return $null }
    $s = [string]$Value
    if ($s -eq '-' -or $s -eq '—') { return $null }
    return $s.Trim()
}

function To-Number([object]$Value) {
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    $n = 0.0
    if ([double]::TryParse(([string]$Value).Replace(',',''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    return $null
}

function Get-At([object[]]$Row, [int]$Index) {
    if ($null -ne $Row -and $Index -lt $Row.Count) { return $Row[$Index] }
    return $null
}

$costPath = Join-Path $env:TEMP 'pet_cost_source_copy.xlsx'
$stockPath = Join-Path $env:TEMP 'pet_stock_source_copy.xlsx'
if (-not (Test-Path $costPath)) { $costPath = 'D:\market-research\宠速星\对货表.xlsx' }
if (-not (Test-Path $stockPath)) { $stockPath = 'D:\market-research\宠速星\商品库存情况\2026-08-27 美团导出商品库存情况.xlsx' }
$costRows = Read-XlsxRows $costPath 'Sheet1'
$stockRows = Read-XlsxRows $stockPath 'Sheet1'

$costData = @()
foreach ($r in ($costRows | Select-Object -Skip 2)) {
    $code = Normalize (Get-At $r 1)
    if (-not $code) { continue }
    $refCost = To-Number (Get-At $r 7)
    if ($null -eq $refCost) { continue }
    $costData += ,@($code, $refCost, $null, $null, '默认按货号成本；拆包销售请补销售拆分数量', 'D:\market-research\宠速星\对货表.xlsx')
}

$stockGroups = @{}
$stockUncoded = @()
foreach ($r in ($stockRows | Select-Object -Skip 2)) {
    $code = Normalize (Get-At $r 2)
    $name = Normalize (Get-At $r 1)
    $stock = To-Number (Get-At $r 3)
    $monthSale = To-Number (Get-At $r 5)
    $status = Normalize (Get-At $r 4)
    $shelf = Normalize (Get-At $r 7)
    if (-not $code) {
        $stockUncoded += ,@([datetime]::Parse('2026-08-27'), $null, $name, $stock, $monthSale, $status, $shelf, $null, $null, 2, $null, '美团商品库存导出；缺少货号', $null)
        continue
    }
    if (-not $stockGroups.ContainsKey($code)) {
        $stockGroups[$code] = [ordered]@{Code=$code;Name=$name;Stock=$stock;MonthSale=$monthSale;Status=$status;Shelf=$shelf;Count=1}
    } else {
        $g = $stockGroups[$code]
        $g.Count++
        if ($null -ne $stock -and ($null -eq $g.Stock -or $stock -gt $g.Stock)) { $g.Stock = $stock }
        if ($null -ne $monthSale) { $g.MonthSale = [double]($g.MonthSale ?? 0) + $monthSale }
        if ($status -eq '售卖中') { $g.Status = '售卖中' }
        if (-not $g.Shelf -and $shelf) { $g.Shelf = $shelf }
    }
}
$stockData = @()
foreach ($g in ($stockGroups.Values | Sort-Object Code)) {
    $source = if ($g.Count -gt 1) { "美团商品库存导出；重复货号$($g.Count)行，库存暂取最大值" } else { '美团商品库存导出' }
    $stockData += ,@([datetime]::Parse('2026-08-27'), $g.Code, $g.Name, $g.Stock, $g.MonthSale, $g.Status, $g.Shelf, $null, $null, 2, $null, $source, $null)
}
$stockData += $stockUncoded
while ($stockData.Count -lt 1999) { $stockData += ,@($null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null) }

$blankRows = @()
for ($i = 0; $i -lt 2000; $i++) { $blankRows += ,@($null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null) }
$blankInputs = @()
for ($i = 0; $i -lt 1000; $i++) { $blankInputs += ,@($null,$null,$null,$null,$null,$null,$null) }
$interventionData = @()
$replenishmentData = @()
foreach ($r in $blankInputs) {
    $interventionData += ,@($r[0],$r[1],$r[2],$r[3],$r[4])
    $replenishmentData += ,@($r[0],$r[1],$r[2],$r[3],$r[4],$r[5],$r[6],$null)
}
$orderData = @()
foreach ($r in $blankRows) { $orderData += ,@($r[0],$r[1],$r[2],$r[3],$r[4],$r[5],$r[6],$r[7],$r[8],$r[9],$r[10],$r[11],$r[12],$r[13],$r[14],$r[15]) }
$orderItemData = @()
for ($i = 0; $i -lt 4000; $i++) { $orderItemData += ,@($null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null) }
$dashboardData = @()
$dashboardData += ,@('经营结果','本月自然订单数',$null,'','线上自然单与线下订单合计；不含干预单和取消单',$null,'')
$dashboardData += ,@('经营结果','本月自然商品销售额',$null,'','线上订单实际商品小计加线下成交收入；不是商品目录标价',$null,'')
$dashboardData += ,@('经营结果','本月预计结算收入',$null,'','线上商家预计收入加线下成交收入；取消单排除，已确认退款按退款后金额',$null,'')
$dashboardData += ,@('经营结果','本月商品成本',$null,'','线上与线下订单货物成本合计',$null,'')
$dashboardData += ,@('经营结果','本月经营毛利',$null,'','线上与线下扣佣后收入-商品及赠品成本-首次配送费',$null,'')
$dashboardData += ,@('经营结果','本月经营利润',$null,'','经营毛利-当月支出表经营费用合计，只扣一次',$null,'')
$dashboardData += ,@('现金与库存','本月补货付款',$null,'','来自补货记录付款金额',$null,'')
$dashboardData += ,@('现金与库存','本月现金净流入',$null,'','平台与线下到账-补货付款-其他现金支出-老板取款',$null,'')
$dashboardData += ,@('现金与库存','本月期末现金',$null,'','期初现金+本月现金净流入',$null,'')
$dashboardData += ,@('现金与库存','当前库存货值',$null,'','（美团库存+待恢复库存数量）×SKU成本主表当前移动平均成本',$null,'')
$dashboardData += ,@('库存提醒','当前缺货SKU数',$null,'','仅监控售卖中商品',$null,'')
$dashboardData += ,@('库存提醒','当前低库存SKU数',$null,'','售卖中且库存不高于补货阈值',$null,'')
$dashboardData += ,@('数据质量','待核对数据数',$null,'','缺货号、缺成本、重复货号及订单成本异常',$null,'')
$dashboardData += ,@('系统状态','最近库存快照日期',$null,'','RPA最近一次读取日期',$null,'')
$dashboardData += ,@('系统状态','最近订单同步日期',$null,'','RPA最近一次读取日期',$null,'')
for ($dashboardPadding = $dashboardData.Count; $dashboardPadding -lt 18; $dashboardPadding++) { $dashboardData += ,@('','',$null,'',$null,'','') }
$settlementData = @()
$settlementStartMonth = [datetime]::Parse('2026-08-01')
for ($monthOffset = 0; $monthOffset -lt 24; $monthOffset++) {
    $settlementData += ,@($settlementStartMonth.AddMonths($monthOffset),$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,'',$null,$null)
}

$sheets = @{ sheets = @(
    @{ name='驾驶舱'; columns=@('模块','指标','数值','计算明细','解释','状态','行动'); data=$dashboardData },
    @{ name='干预订单'; columns=@('日期','单号','补单人','订单编号','补单金额','配送方式','平台费用（自动）','配送费','总成本','是否已恢复库存','是否已好评','购买内容','备注'); data=@() },
    @{ name='补货记录'; columns=@('补货日期','货号','补货数量','参考单位成本','实际采购金额','付款日期','付款金额','备注'); data=$replenishmentData },
    @{ name='月度结算'; columns=@('月份','美团实际到账','淘宝闪购实际到账','京东到家实际到账','线下实际到账','其他现金支出','其他经营费用','老板取款','老板人工','市场租金机会成本','期初现金','备注','品牌/装修/设备摊销折旧','平台推广费用'); data=$settlementData; dtypes=@{'月份'='datetime64[ns]';'美团实际到账'='float64';'淘宝闪购实际到账'='float64';'京东到家实际到账'='float64';'线下实际到账'='float64';'其他现金支出'='float64';'其他经营费用'='float64';'老板取款'='float64';'老板人工'='float64';'市场租金机会成本'='float64';'期初现金'='float64';'品牌/装修/设备摊销折旧'='float64';'平台推广费用'='float64'}; formats=@{'月份'='yyyy-mm';'美团实际到账'='#,##0.00';'淘宝闪购实际到账'='#,##0.00';'京东到家实际到账'='#,##0.00';'线下实际到账'='#,##0.00';'其他现金支出'='#,##0.00';'其他经营费用'='#,##0.00';'老板取款'='#,##0.00';'老板人工'='#,##0.00';'市场租金机会成本'='#,##0.00';'期初现金'='#,##0.00';'品牌/装修/设备摊销折旧'='#,##0.00';'平台推广费用'='#,##0.00'} },
    @{ name='订单汇总'; columns=@('下单日期','平台','平台订单号','订单状态','订单类型','商品销售额','顾客实际支付','商家预计收入','平台费用','配送费','商品成本','订单毛利','商品数量','识别依据','同步时间','备注'); data=$orderData; dtypes=@{'下单日期'='datetime64[ns]';'同步时间'='datetime64[ns]'}; formats=@{'下单日期'='yyyy-mm-dd';'商品销售额'='#,##0.00';'顾客实际支付'='#,##0.00';'商家预计收入'='#,##0.00';'平台费用'='#,##0.00';'配送费'='#,##0.00';'商品成本'='#,##0.00';'订单毛利'='#,##0.00';'同步时间'='yyyy-mm-dd hh:mm'} },
    @{ name='订单商品'; columns=@('下单日期','平台','平台订单号','订单状态','订单类型','货号','商品名称','规格','数量','商品售价','商品销售额','单位成本','商品成本','商品毛利','同步时间','库存恢复状态'); data=$orderItemData; dtypes=@{'下单日期'='datetime64[ns]';'数量'='float64';'商品售价'='float64';'同步时间'='datetime64[ns]'}; formats=@{'下单日期'='yyyy-mm-dd';'数量'='#,##0.##';'商品售价'='#,##0.00';'商品销售额'='#,##0.00';'单位成本'='#,##0.00';'商品成本'='#,##0.00';'商品毛利'='#,##0.00';'同步时间'='yyyy-mm-dd hh:mm'} },
    @{ name='库存快照'; columns=@('快照日期','货号','商品名称','库存','月售','销售状态','货架码','单位成本','库存货值','补货阈值','库存提醒','来源','干预扣减待恢复'); data=$stockData; dtypes=@{'快照日期'='datetime64[ns]';'库存'='float64';'月售'='float64';'单位成本'='float64';'库存货值'='float64';'补货阈值'='float64';'干预扣减待恢复'='float64'}; formats=@{'快照日期'='yyyy-mm-dd';'库存'='#,##0.##';'月售'='#,##0.##';'单位成本'='#,##0.00';'库存货值'='#,##0.00';'补货阈值'='#,##0';'干预扣减待恢复'='#,##0.##'} },
    @{ name='商品成本'; columns=@('货号','参考进货价','销售拆分数量','单位成本','成本状态','来源'); data=$costData; dtypes=@{'参考进货价'='float64';'销售拆分数量'='float64';'单位成本'='float64'}; formats=@{'参考进货价'='#,##0.00';'销售拆分数量'='#,##0.##';'单位成本'='#,##0.00'} }
) }

$styles = @{ styles = @(
    @{name='驾驶舱'; cell_styles=@(
        @{range='A1:G1';background_color='#173A2D';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center';vertical_alignment='middle'};
        @{range='A2:G19';background_color='#F4F7F5'};
        @{range='C2:C19';background_color='#E3F0E8';font_weight='bold'};
        @{range='C2:C2';number_format='0'};
        @{range='C3:C11';number_format='#,##0.00'};
        @{range='C12:C14';number_format='0'};
        @{range='C15:C16';number_format='yyyy-mm-dd hh:mm'};
        @{range='D2:D19';background_color='#EEF5F0';word_wrap='auto-wrap';vertical_alignment='middle'};
        @{range='F2:F19';background_color='#FFF8E1'}
    ); row_sizes=@(@{range='1';type='pixel';size=40}); col_sizes=@(@{range='A';type='pixel';size=120},@{range='B';type='pixel';size=180},@{range='C';type='pixel';size=120},@{range='D';type='pixel';size=520},@{range='E';type='pixel';size=260},@{range='F';type='pixel';size=120},@{range='G';type='pixel';size=240})};
    @{name='干预订单'; cell_styles=@(@{range='A1:M1';background_color='#7A2E2E';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center'};@{range='A2:M1001';background_color='#FFF8E1'};@{range='A2:A1001';number_format='yyyy-mm-dd'};@{range='E2:I1001';number_format='#,##0.00'}); col_sizes=@(@{range='A';type='pixel';size=110},@{range='B:C';type='pixel';size=120},@{range='D';type='pixel';size=210},@{range='E:I';type='pixel';size=120},@{range='J:K';type='pixel';size=130},@{range='L';type='pixel';size=220},@{range='M';type='pixel';size=260})};
    @{name='补货记录'; cell_styles=@(@{range='A1:H1';background_color='#8A5A13';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center'};@{range='A2:H1001';background_color='#FFF8E1'};@{range='D2:D1001';background_color='#E8F1EC'}); col_sizes=@(@{range='A';type='pixel';size=110},@{range='B';type='pixel';size=100},@{range='C:G';type='pixel';size=125},@{range='H';type='pixel';size=220})};
    @{name='月度结算'; cell_styles=@(@{range='A1:N1';background_color='#173A2D';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center'};@{range='A2:N25';background_color='#FFF8E1'};@{range='A2:A25';background_color='#F2F4F3'}); col_sizes=@(@{range='A';type='pixel';size=100},@{range='B:K';type='pixel';size=130},@{range='L';type='pixel';size=220},@{range='M';type='pixel';size=180},@{range='N';type='pixel';size=130})};
    @{name='订单汇总'; cell_styles=@(@{range='A1:P1';background_color='#315B7D';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center';word_wrap='auto-wrap'};@{range='A2:P2001';background_color='#F4F7F5'}); col_sizes=@(@{range='A';type='pixel';size=110},@{range='B';type='pixel';size=80},@{range='C';type='pixel';size=190},@{range='D:E';type='pixel';size=110},@{range='F:L';type='pixel';size=115},@{range='M';type='pixel';size=90},@{range='N:P';type='pixel';size=160})};
    @{name='订单商品'; cell_styles=@(@{range='A1:P1';background_color='#315B7D';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center';word_wrap='auto-wrap'};@{range='A2:P4001';background_color='#F4F7F5'}); col_sizes=@(@{range='A';type='pixel';size=110},@{range='B';type='pixel';size=80},@{range='C';type='pixel';size=190},@{range='D:E';type='pixel';size=110},@{range='F';type='pixel';size=100},@{range='G';type='pixel';size=280},@{range='H';type='pixel';size=180},@{range='I:N';type='pixel';size=110},@{range='O:P';type='pixel';size=160})};
    @{name='库存快照'; cell_styles=@(@{range='A1:M1';background_color='#315B7D';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center';word_wrap='auto-wrap'};@{range='A2:M2000';background_color='#F4F7F5'};@{range='K2:K2000';background_color='#FFF8E1'}); col_sizes=@(@{range='A';type='pixel';size=110},@{range='B';type='pixel';size=100},@{range='C';type='pixel';size=280},@{range='D:E';type='pixel';size=90},@{range='F:G';type='pixel';size=120},@{range='H:I';type='pixel';size=120},@{range='J:M';type='pixel';size=150})};
    @{name='商品成本'; cell_styles=@(@{range='A1:F1';background_color='#315B7D';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center'};@{range='A2:F2000';background_color='#F4F7F5'};@{range='C2:C2000';background_color='#FFF8E1'};@{range='D2:D2000';background_color='#E8F1EC'}); col_sizes=@(@{range='A';type='pixel';size=100},@{range='B:D';type='pixel';size=130},@{range='E';type='pixel';size=300},@{range='F';type='pixel';size=260})}
) }

$payload = $sheets | ConvertTo-Json -Depth 20 -Compress
$styleJson = $styles | ConvertTo-Json -Depth 20 -Compress
$token = $env:SIMPLE_FINANCE_RESUME_TOKEN
if ($token) {
    $resultRaw = $payload | lark-cli sheets +table-put --as user --spreadsheet-token $token --sheets - --styles $styleJson --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw ("填充简版飞书工作簿失败：" + $resultRaw) }
} else {
    $resultRaw = $payload | lark-cli sheets +workbook-create --as user --title '宠速星｜经营驾驶舱（简版）' --sheets - --styles $styleJson --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw ("创建简版飞书工作簿失败：" + $resultRaw) }
    $result = $resultRaw | ConvertFrom-Json
    $token = $result.data.spreadsheet_token
    if (-not $token) { $token = $result.data.token }
    if (-not $token) { throw '创建成功但未能读取工作簿引用' }
}

function Set-Cells([string]$Sheet, [string]$Range, $Cells) {
    $json = $Cells | ConvertTo-Json -Depth 12 -Compress
    $null = $json | lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $Sheet --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入公式失败: $Sheet $Range" }
}

function Set-FormulaColumn([string]$Sheet, [string]$Column, [int]$LastRow, [string]$Template) {
    $cells = @()
    for ($row = 2; $row -le $LastRow; $row++) {
        $cells += ,@(@{formula=$Template.Replace('{r}', [string]$row)})
    }
    Set-Cells $Sheet ("$Column`2:$Column$LastRow") $cells
}

# “补单人”和“购买内容”是自由文本；三个状态列使用下拉。
$null = lark-cli sheets +dropdown-set --as user --spreadsheet-token $token --sheet-name '干预订单' --range 'F2:F1001' --options '["自配送","骑手配送"]' --colors '["#2563EB","#F59E0B"]' --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '设置干预订单配送方式下拉失败' }
$null = lark-cli sheets +dropdown-set --as user --spreadsheet-token $token --sheet-name '干预订单' --range 'J2:J1001' --options '["待恢复","已恢复"]' --colors '["#F59E0B","#16A34A"]' --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '设置干预订单库存恢复状态下拉失败' }
$null = lark-cli sheets +dropdown-set --as user --spreadsheet-token $token --sheet-name '干预订单' --range 'K2:K1001' --options '["未好评","已好评"]' --colors '["#9CA3AF","#16A34A"]' --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '设置干预订单好评状态下拉失败' }

$dashboardMonthOptions = [Collections.Generic.List[string]]::new()
$dashboardMonthCursor = [datetime]'2026-08-01'
$dashboardMonthEnd = [datetime]::Today.AddDays(1-[datetime]::Today.Day).Date
while ($dashboardMonthCursor -le $dashboardMonthEnd) {
    $dashboardMonthOptions.Add($dashboardMonthCursor.ToString('yyyy-MM'))
    $dashboardMonthCursor = $dashboardMonthCursor.AddMonths(1)
}
$dashboardMonthOptionsJson = @($dashboardMonthOptions) | ConvertTo-Json -Compress
$null = $dashboardMonthOptionsJson | lark-cli sheets +dropdown-set --as user --spreadsheet-token $token --sheet-name '驾驶舱' --range 'I1:I1' --options - --highlight=false --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '设置驾驶舱月份下拉失败' }
Set-Cells '驾驶舱' 'H1:I1' @(@(@{value='查看月份'},@{value=[datetime]::Today.ToString('yyyy-MM')}))
$null = lark-cli sheets +cells-set-style --as user --spreadsheet-token $token --sheet-name '驾驶舱' --range 'H1:H1' --background-color '#173A2D' --font-color '#FFFFFF' --font-weight bold --horizontal-alignment center --vertical-alignment middle --format json | Out-String
$null = lark-cli sheets +cells-set-style --as user --spreadsheet-token $token --sheet-name '驾驶舱' --range 'I1:I1' --background-color '#E3F0E8' --font-weight bold --horizontal-alignment center --vertical-alignment middle --number-format '@' --format json | Out-String
$null = lark-cli sheets +cols-resize --as user --spreadsheet-token $token --sheet-name '驾驶舱' --range 'H:I' --width 120 --format json | Out-String

Set-FormulaColumn '干预订单' 'I' 1001 '=IF(D{r}="","",IFERROR(G{r}*1,0)+IFERROR(H{r}*1,0))'
Set-FormulaColumn '商品成本' 'D' 2000 '=IF(A{r}="","",IF(C{r}>0,B{r}/C{r},B{r}))'
Set-FormulaColumn '商品成本' 'E' 2000 '=IF(A{r}="","",IF(OR(A{r}="csx188",A{r}="csx576"),"待人工核对",IF(C{r}>0,"拆分成本","默认成本")))'
Set-FormulaColumn '补货记录' 'D' 1001 '=IF(B{r}="","",IFERROR(VLOOKUP(B{r},''商品成本''!$A$2:$D$2000,4,FALSE),IFERROR(VLOOKUP(VLOOKUP(B{r},''库存快照''!$B$2:$G$2000,6,FALSE),''商品成本''!$A$2:$D$2000,4,FALSE),"")))'
Set-FormulaColumn '订单汇总' 'E' 2001 '=IF(C{r}="","",IF(COUNTIF(''干预订单''!$D$2:$D$1001,C{r})>0,"干预","自然"))'
Set-FormulaColumn '订单汇总' 'K' 2001 '=IF(C{r}="","",SUMIFS(''订单商品''!$M$2:$M$4001,''订单商品''!$C$2:$C$4001,C{r}))'
Set-FormulaColumn '订单汇总' 'L' 2001 '=IF(C{r}="","",IFERROR(H{r}*1,0)-IFERROR(K{r}*1,0)-IFERROR(J{r}*1,0))'
Set-FormulaColumn '订单商品' 'E' 4001 '=IF(C{r}="","",IF(COUNTIF(''干预订单''!$D$2:$D$1001,C{r})>0,"干预","自然"))'
Set-FormulaColumn '订单商品' 'K' 4001 '=IF(F{r}="","",IFERROR(I{r}*1,0)*IFERROR(J{r}*1,0))'
Set-FormulaColumn '订单商品' 'L' 4001 '=IF(F{r}="","",IFERROR(IF(VLOOKUP(F{r},''商品成本''!$A$2:$E$2000,5,FALSE)="待人工核对","",VLOOKUP(F{r},''商品成本''!$A$2:$D$2000,4,FALSE)),IFERROR(IF(VLOOKUP(VLOOKUP(F{r},''库存快照''!$B$2:$G$2000,6,FALSE),''商品成本''!$A$2:$E$2000,5,FALSE)="待人工核对","",VLOOKUP(VLOOKUP(F{r},''库存快照''!$B$2:$G$2000,6,FALSE),''商品成本''!$A$2:$D$2000,4,FALSE)),"")))'
Set-FormulaColumn '订单商品' 'M' 4001 '=IF(F{r}="","",IF(E{r}="干预",0,IF(L{r}="","",IFERROR(I{r}*1,0)*IFERROR(L{r}*1,0))))'
Set-FormulaColumn '订单商品' 'N' 4001 '=IF(F{r}="","",IF(M{r}="","",K{r}-M{r}))'
Set-FormulaColumn '订单商品' 'P' 4001 '=IF(C{r}="","",IF(E{r}<>"干预","",IF(COUNTIFS(''干预订单''!$D$2:$D$1001,C{r},''干预订单''!$J$2:$J$1001,"已恢复")>0,"已恢复","待恢复")))'
Set-FormulaColumn '库存快照' 'H' 2000 '=IF(B{r}="","",IFERROR(IF(ISBLANK(INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(N{r},''SKU成本主表''!$Q$2:$Q$3000,0))),"",INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(N{r},''SKU成本主表''!$Q$2:$Q$3000,0))),IFERROR(IF(ISBLANK(INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(B{r},''SKU成本主表''!$A$2:$A$3000,0))),"",INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(B{r},''SKU成本主表''!$A$2:$A$3000,0))),IFERROR(IF(ISBLANK(INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(G{r},''SKU成本主表''!$B$2:$B$3000,0))),"",INDEX(''SKU成本主表''!$G$2:$G$3000,MATCH(G{r},''SKU成本主表''!$B$2:$B$3000,0))),""))))'
Set-FormulaColumn '库存快照' 'I' 2000 '=IF(OR(B{r}="",D{r}="",H{r}=""),"",(D{r}+IFERROR(M{r}*1,0))*H{r})'
Set-FormulaColumn '库存快照' 'K' 2000 '=IF(C{r}="","",IF(B{r}="","缺少货号",IF(H{r}="","成本待核对",IF(ISNUMBER(SEARCH("重复货号",L{r})),"重复货号待核对",IF(F{r}<>"售卖中","不监控",IF(D{r}<=0,"缺货",IF(D{r}<=J{r},"低库存","正常")))))))'
Set-FormulaColumn '库存快照' 'M' 2000 '=IF(B{r}="","",SUMIFS(''订单商品''!$I$2:$I$4001,''订单商品''!$E$2:$E$4001,"干预",''订单商品''!$D$2:$D$4001,"<>已取消",''订单商品''!$F$2:$F$4001,B{r},''订单商品''!$P$2:$P$4001,"待恢复")+IF(OR(G{r}="",G{r}=B{r}),0,SUMIFS(''订单商品''!$I$2:$I$4001,''订单商品''!$E$2:$E$4001,"干预",''订单商品''!$D$2:$D$4001,"<>已取消",''订单商品''!$F$2:$F$4001,G{r},''订单商品''!$P$2:$P$4001,"待恢复")))'

$currentMonth = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)'
$nextMonth = 'EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1)'
$deliveryExpenseExpression = 'SUMIFS(''订单汇总''!$J$2:$J$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$E$2:$E$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))'
$promotionExpenseExpression = 'SUMIFS(''月度结算''!$N$2:$N$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
. (Join-Path $PSScriptRoot '..\agent\expense_accounting.ps1')
$expenseFormulas = Get-PetstarDashboardExpenseFormulas $currentMonth $deliveryExpenseExpression
$cashReceiptExpression = 'SUMIFS(''月度结算''!$B$2:$B$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$C$2:$C$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$D$2:$D$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$E$2:$E$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$otherCashExpenseExpression = 'SUMIFS(''月度结算''!$F$2:$F$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$ownerWithdrawalExpression = 'SUMIFS(''月度结算''!$H$2:$H$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$beginningCashExpression = 'SUMIFS(''月度结算''!$K$2:$K$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'
$dashboardFormulas = @(
    '=COUNTIFS(''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$D$2:$D$2001,"<>已取消",''订单汇总''!$E$2:$E$2001,"自然")+COUNTIFS(''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''线下订单''!$C$2:$C$1001,"<>")',
    '=SUMIFS(''订单汇总''!$F$2:$F$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$D$2:$D$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))',
    '=SUMIFS(''订单汇总''!$H$2:$H$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$D$2:$D$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))',
    '=SUMIFS(''订单汇总''!$K$2:$K$2001,''订单汇总''!$A$2:$A$2001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单汇总''!$A$2:$A$2001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单汇总''!$E$2:$E$2001,"自然",''订单汇总''!$D$2:$D$2001,"<>已取消")+SUMIFS(''线下订单''!$F$2:$F$1001,''线下订单''!$A$2:$A$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''线下订单''!$A$2:$A$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))',
    $expenseFormulas.gross,
    $expenseFormulas.profit,
    '=SUMIFS(''补货记录''!$M$2:$M$1001,''补货记录''!$L$2:$L$1001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''补货记录''!$L$2:$L$1001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))',
    '=SUMIFS(''月度结算''!$B$2:$B$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$C$2:$C$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$D$2:$D$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))+SUMIFS(''月度结算''!$E$2:$E$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))-C8-SUMIFS(''月度结算''!$F$2:$F$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))-SUMIFS(''月度结算''!$H$2:$H$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
    '=C9+SUMIFS(''月度结算''!$K$2:$K$25,''月度结算''!$A$2:$A$25,DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
    '=SUMIFS(''库存快照''!$I$2:$I$2000,''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000))',
    '=COUNTIFS(''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000),''库存快照''!$K$2:$K$2000,"缺货")',
    '=COUNTIFS(''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000),''库存快照''!$K$2:$K$2000,"低库存")',
    '=COUNTIFS(''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000),''库存快照''!$K$2:$K$2000,"*核对")+COUNTIFS(''库存快照''!$A$2:$A$2000,MAX(''库存快照''!$A$2:$A$2000),''库存快照''!$K$2:$K$2000,"缺少*")+COUNTIFS(''订单商品''!$A$2:$A$4001,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),''订单商品''!$A$2:$A$4001,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),''订单商品''!$C$2:$C$4001,"<>",''订单商品''!$L$2:$L$4001,"")',
    '=IF(COUNTA(''库存快照''!$A$2:$A$2000)=0,"",MAX(''库存快照''!$A$2:$A$2000))',
    '=IF(COUNTA(''订单汇总''!$O$2:$O$2001)=0,"",MAX(''订单汇总''!$O$2:$O$2001))',
    '=""',
    '=""',
    '=""'
)
$systemCurrentGrossProfitFormula = $dashboardFormulas[4]
$dashboardSelectedMonthExpression = 'DATE(VALUE(LEFT($I$1,4)),VALUE(RIGHT($I$1,2)),1)'
$dashboardFormulas = @($dashboardFormulas | ForEach-Object { $_.Replace('DATE(YEAR(TODAY()),MONTH(TODAY()),1)',$dashboardSelectedMonthExpression) })
$dashboardCells = @()
for ($i = 0; $i -lt $dashboardFormulas.Count; $i++) { $dashboardCells += ,@(@{formula=$dashboardFormulas[$i]}) }
Set-Cells '驾驶舱' 'C2:C19' $dashboardCells

$dashboardDetailFormulas = @{
    6 = $expenseFormulas.gross_detail
    7 = $expenseFormulas.profit_detail
    9 = '="=平台与线下到账-补货付款-其他现金支出-老板取款="&TEXT((' + $cashReceiptExpression + '),"#,##0.00")&"-"&TEXT(C8,"#,##0.00")&"-"&TEXT((' + $otherCashExpenseExpression + '),"#,##0.00")&"-"&TEXT((' + $ownerWithdrawalExpression + '),"#,##0.00")'
    10 = '="=期初现金+本月现金净流入="&TEXT((' + $beginningCashExpression + '),"#,##0.00")&IF(C9>=0,"+","")&TEXT(C9,"#,##0.00")'
}
foreach ($detailRow in @($dashboardDetailFormulas.Keys)) {
    $dashboardDetailFormulas[$detailRow] = $dashboardDetailFormulas[$detailRow].Replace('DATE(YEAR(TODAY()),MONTH(TODAY()),1)',$dashboardSelectedMonthExpression)
}
$dashboardDetailCells = @()
for ($rowNumber = 2; $rowNumber -le 19; $rowNumber++) {
    if ($dashboardDetailFormulas.ContainsKey($rowNumber)) { $dashboardDetailCells += ,@(@{formula=$dashboardDetailFormulas[$rowNumber]}) }
    else { $dashboardDetailCells += ,@(@{value=''}) }
}
Set-Cells '驾驶舱' 'D2:D19' $dashboardDetailCells
Set-Cells '驾驶舱' 'F7:F7' @(@(@{formula=$expenseFormulas.status.Replace($currentMonth,$dashboardSelectedMonthExpression)}))
Set-Cells '驾驶舱' 'J1:J2' @(@(@{value='系统当月经营毛利'}),@(@{formula=$systemCurrentGrossProfitFormula}))
$null = lark-cli sheets +dim-hide --as user --spreadsheet-token $token --sheet-name '驾驶舱' --range 'J:J' --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw '隐藏驾驶舱系统辅助列失败' }

$verifyTargets = @(
    @{Sheet='干预订单';Range='I2:I1001'},
    @{Sheet='驾驶舱';Range='C2:C19'},
    @{Sheet='驾驶舱';Range='D6:D7'},
    @{Sheet='驾驶舱';Range='D9:D10'},
    @{Sheet='驾驶舱';Range='J2:J2'},
    @{Sheet='补货记录';Range='D2:D1001'},
    @{Sheet='订单汇总';Range='E2:L2001'},
    @{Sheet='订单商品';Range='E2:P4001'},
    @{Sheet='库存快照';Range='H2:M2000'},
    @{Sheet='商品成本';Range='D2:E2000'}
)
foreach ($target in $verifyTargets) {
    $verifyRaw = lark-cli sheets +formula-verify --as user --spreadsheet-token $token --sheet-name $target.Sheet --range $target.Range --format json --exit-on-error | Out-String
    $verify = $verifyRaw | ConvertFrom-Json
    if ($verify.data.status -ne 'success') { throw "公式校验未通过: $($target.Sheet) $($target.Range)" }
}

Write-Output ("SIMPLE_WORKBOOK_CREATED: https://my.feishu.cn/sheets/$token")
