param(
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$metaPath = Join-Path $ProjectRoot 'work\daily_ledger_meta.json'
if (Test-Path -LiteralPath $metaPath) {
    $existing = [System.IO.File]::ReadAllText($metaPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    [ordered]@{
        ok = $true
        created = $false
        workbook_url = $existing.workbook_url
        message = '每日经营简账已存在，本次未重复创建'
    } | ConvertTo-Json -Depth 5
    exit 0
}

$incomeColumns = @('当日单号','订单编号','日期','扣佣后收入','货物成本','配送费','这一单赚/亏','售出货品（KW）','赠品（KW）','当天赚/亏','备注')
$expenseColumns = @('日期','项目','费用','订单编号','来源','备注')
$sheetDefinitions = @(
    @{name='8月收入';columns=$incomeColumns;data=@();dtypes=@{'当日单号'='object';'订单编号'='object';'日期'='datetime64[ns]';'扣佣后收入'='float64';'货物成本'='float64';'配送费'='float64';'这一单赚/亏'='float64';'售出货品（KW）'='object';'赠品（KW）'='object';'当天赚/亏'='float64';'备注'='object'};formats=@{'日期'='yyyy-mm-dd';'扣佣后收入'='#,##0.00';'货物成本'='#,##0.00';'配送费'='#,##0.00';'这一单赚/亏'='#,##0.00;[Red]-#,##0.00';'当天赚/亏'='#,##0.00;[Red]-#,##0.00'}},
    @{name='8月支出';columns=$expenseColumns;data=@();dtypes=@{'日期'='datetime64[ns]';'项目'='object';'费用'='float64';'订单编号'='object';'来源'='object';'备注'='object'};formats=@{'日期'='yyyy-mm-dd';'费用'='#,##0.00'}},
    @{name='9月收入';columns=$incomeColumns;data=@();dtypes=@{'当日单号'='object';'订单编号'='object';'日期'='datetime64[ns]';'扣佣后收入'='float64';'货物成本'='float64';'配送费'='float64';'这一单赚/亏'='float64';'售出货品（KW）'='object';'赠品（KW）'='object';'当天赚/亏'='float64';'备注'='object'};formats=@{'日期'='yyyy-mm-dd';'扣佣后收入'='#,##0.00';'货物成本'='#,##0.00';'配送费'='#,##0.00';'这一单赚/亏'='#,##0.00;[Red]-#,##0.00';'当天赚/亏'='#,##0.00;[Red]-#,##0.00'}},
    @{name='9月支出';columns=$expenseColumns;data=@();dtypes=@{'日期'='datetime64[ns]';'项目'='object';'费用'='float64';'订单编号'='object';'来源'='object';'备注'='object'};formats=@{'日期'='yyyy-mm-dd';'费用'='#,##0.00'} }
)

$incomeStyle = @(
    @{range='A1:K1';background_color='#173A2D';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center';vertical_alignment='middle';word_wrap='auto-wrap'},
    @{range='A2:K1001';background_color='#F7FAF8';vertical_alignment='middle'},
    @{range='D2:F1001';number_format='#,##0.00'},
    @{range='G2:G1001';background_color='#E8F1EC';font_weight='bold';number_format='#,##0.00;[Red]-#,##0.00'},
    @{range='I2:I1001';background_color='#EAF5EE'},
    @{range='J2:J1001';background_color='#DCEBE2';font_weight='bold';number_format='#,##0.00;[Red]-#,##0.00'},
    @{range='K2:K1001';background_color='#FFF8E1'},
    @{range='C2:C1001';number_format='yyyy-mm-dd'}
)
$expenseStyle = @(
    @{range='A1:F1';background_color='#8A5A13';font_color='#FFFFFF';font_weight='bold';horizontal_alignment='center';vertical_alignment='middle'},
    @{range='A2:F1001';background_color='#FFF8E1';vertical_alignment='middle'},
    @{range='A2:A1001';number_format='yyyy-mm-dd'},
    @{range='C2:C1001';number_format='#,##0.00'},
    @{range='D2:E1001';background_color='#F4F7F5'},
    @{range='F2:F1001';background_color='#FFFDF5'}
)
$styles = @{styles=@(
    @{name='8月收入';cell_styles=$incomeStyle;row_sizes=@(@{range='1';type='pixel';size=42});col_sizes=@(@{range='A';type='pixel';size=120},@{range='B';type='pixel';size=205},@{range='C';type='pixel';size=110},@{range='D:G';type='pixel';size=120},@{range='H';type='pixel';size=300},@{range='I';type='pixel';size=220},@{range='J';type='pixel';size=125},@{range='K';type='pixel';size=260})},
    @{name='8月支出';cell_styles=$expenseStyle;row_sizes=@(@{range='1';type='pixel';size=42});col_sizes=@(@{range='A';type='pixel';size=110},@{range='B';type='pixel';size=160},@{range='C';type='pixel';size=110},@{range='D';type='pixel';size=205},@{range='E';type='pixel';size=105},@{range='F';type='pixel';size=300})},
    @{name='9月收入';cell_styles=$incomeStyle;row_sizes=@(@{range='1';type='pixel';size=42});col_sizes=@(@{range='A';type='pixel';size=120},@{range='B';type='pixel';size=205},@{range='C';type='pixel';size=110},@{range='D:G';type='pixel';size=120},@{range='H';type='pixel';size=300},@{range='I';type='pixel';size=220},@{range='J';type='pixel';size=125},@{range='K';type='pixel';size=260})},
    @{name='9月支出';cell_styles=$expenseStyle;row_sizes=@(@{range='1';type='pixel';size=42});col_sizes=@(@{range='A';type='pixel';size=110},@{range='B';type='pixel';size=160},@{range='C';type='pixel';size=110},@{range='D';type='pixel';size=205},@{range='E';type='pixel';size=105},@{range='F';type='pixel';size=300})}
)}

$payload = @{sheets=$sheetDefinitions} | ConvertTo-Json -Depth 20 -Compress
$styleJson = $styles | ConvertTo-Json -Depth 20 -Compress
$raw = $payload | lark-cli sheets +workbook-create --as user --title '宠速星｜每日经营简账' --sheets - --styles $styleJson --format json | Out-String
if ($LASTEXITCODE -ne 0) { throw "创建每日经营简账失败：$($raw.Trim())" }
$result = $raw | ConvertFrom-Json
if (-not $result.ok) { throw "创建每日经营简账失败：$($result.error.message)" }
$token = $result.data.spreadsheet_token
if (-not $token) { $token = $result.data.token }
if (-not $token -and $result.data.spreadsheet) { $token = $result.data.spreadsheet.spreadsheet_token }
if (-not $token -and $result.data.spreadsheet) { $token = $result.data.spreadsheet.token }
if (-not $token) { throw '创建成功但未读取到工作簿引用' }
$workbookUrl = "https://my.feishu.cn/sheets/$token"

foreach ($sheetName in @('8月收入','8月支出','9月收入','9月支出')) {
    $freezeRaw = lark-cli sheets +dim-freeze --as user --spreadsheet-token $token --sheet-name $sheetName --dimension row --count 1 --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not (($freezeRaw | ConvertFrom-Json).ok)) { throw "冻结表头失败：$sheetName" }
    $range = if ($sheetName.EndsWith('收入')) { 'A1:K1001' } else { 'A1:F1001' }
    $filterProperties = ConvertTo-Json -InputObject @{range=$range;rules=@()} -Compress
    $filterRaw = $filterProperties | lark-cli sheets +filter-create --as user --spreadsheet-token $token --sheet-name $sheetName --range $range --properties - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "创建筛选器失败：$sheetName" }
}

[ordered]@{
    title = '宠速星｜每日经营简账'
    workbook_url = $workbookUrl
    created_on = [datetime]::Now.ToString('yyyy-MM-dd')
    identity = 'user-default'
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $metaPath -Encoding UTF8

[ordered]@{
    ok = $true
    created = $true
    workbook_url = $workbookUrl
    sheets = @('8月收入','8月支出','9月收入','9月支出')
} | ConvertTo-Json -Depth 5
