$ErrorActionPreference = 'Stop'
$startMonth = [datetime]'2026-08-01'
$monthCount = 24
$endDate = $startMonth.AddMonths($monthCount).AddDays(-1)

function New-NullRow([int]$Count) {
    $row = [object[]]::new($Count)
    return ,$row
}

$dailyColumns = @('日期','美团净销售额','淘宝闪购净销售额','京东到家净销售额','线下销售额','总净销售额','美团实际到账','淘宝闪购实际到账','京东到家实际到账','线下实际到账','总实际到账','采购入库成本','采购付款','其他现金支出','老板取款','备注','月份')
$dailyRows = [System.Collections.ArrayList]::new()
for ($date = $startMonth; $date -le $endDate; $date = $date.AddDays(1)) {
    $row = New-NullRow $dailyColumns.Count
    $row[0] = $date.ToString('yyyy-MM-dd')
    [void]$dailyRows.Add($row)
}

$supplementColumns = @('月份','期初库存','期末库存','平台服务费/配送费/支付费（应计）','其他经营费用（应计）','老板人工','市场租金机会成本','首月期初现金','备注')
$supplementRows = [System.Collections.ArrayList]::new()
for ($i = 0; $i -lt $monthCount; $i++) {
    $row = New-NullRow $supplementColumns.Count
    $row[0] = $startMonth.AddMonths($i).ToString('yyyy-MM-dd')
    [void]$supplementRows.Add($row)
}

$summaryColumns = @('月份','净销售额','实际到账','采购入库成本','采购付款','期初库存','期末库存','库存变化','已售商品成本','毛利','其他现金支出','平台服务费/配送费/支付费','其他经营费用','品牌服务费摊销','装修/货架/监控/无人仓折旧','经营费用','老板人工','经营利润','市场租金机会成本','经济利润','老板取款','现金净流入','期初现金','期末现金')
$summaryRows = [System.Collections.ArrayList]::new()
for ($i = 0; $i -lt $monthCount; $i++) {
    $row = New-NullRow $summaryColumns.Count
    $row[0] = $startMonth.AddMonths($i).ToString('yyyy-MM-dd')
    [void]$summaryRows.Add($row)
}

$assetColumns = @('资产/费用项目','核算分类','金额','开始摊销/折旧月','期限（月）','结束月','月摊销/折旧','备注')
$assetRows = [System.Collections.ArrayList]::new()
for ($i = 0; $i -lt 100; $i++) {
    [void]$assetRows.Add((New-NullRow $assetColumns.Count))
}

$moneyFormats = @{}
foreach ($name in $dailyColumns[1..14]) { $moneyFormats[$name] = '#,##0.00' }
$moneyFormats['日期'] = 'yyyy-mm-dd'
$moneyFormats['月份'] = '@'
$dailyDtypes = @{'日期'='datetime64[ns]'; '备注'='object'; '月份'='object'}
foreach ($name in $dailyColumns[1..14]) { $dailyDtypes[$name] = 'float64' }

$supplementDtypes = @{'月份'='datetime64[ns]'; '备注'='object'}
$supplementFormats = @{'月份'='yyyy-mm'}
foreach ($name in $supplementColumns[1..7]) { $supplementDtypes[$name] = 'float64'; $supplementFormats[$name] = '#,##0.00' }

$summaryDtypes = @{'月份'='datetime64[ns]'}
$summaryFormats = @{'月份'='yyyy-mm'}
foreach ($name in $summaryColumns[1..23]) { $summaryDtypes[$name] = 'float64'; $summaryFormats[$name] = '#,##0.00' }

$assetDtypes = @{'资产/费用项目'='object'; '核算分类'='object'; '金额'='float64'; '开始摊销/折旧月'='datetime64[ns]'; '期限（月）'='float64'; '结束月'='datetime64[ns]'; '月摊销/折旧'='float64'; '备注'='object'}
$assetFormats = @{'金额'='#,##0.00'; '开始摊销/折旧月'='yyyy-mm'; '期限（月）'='0'; '结束月'='yyyy-mm'; '月摊销/折旧'='#,##0.00'}

$sheetsPayload = @{sheets=@(
    @{name='每日记录'; columns=$dailyColumns; data=$dailyRows; dtypes=$dailyDtypes; formats=$moneyFormats},
    @{name='月度补充'; columns=$supplementColumns; data=$supplementRows; dtypes=$supplementDtypes; formats=$supplementFormats},
    @{name='月度汇总'; columns=$summaryColumns; data=$summaryRows; dtypes=$summaryDtypes; formats=$summaryFormats},
    @{name='资产与口径'; columns=$assetColumns; data=$assetRows; dtypes=$assetDtypes; formats=$assetFormats}
)}

$header = @{range='A1:Q1'; background_color='#1F4D3B'; font_color='#FFFFFF'; font_weight='bold'; horizontal_alignment='center'; vertical_alignment='middle'; word_wrap='auto-wrap'; border_styles=@{bottom=@{style='solid';weight='medium';color='#173A2D'}}}
$stylesPayload = @{styles=@(
    @{name='每日记录'; cell_styles=@(
        $header,
        @{range='A2:A732';background_color='#F2F4F3';horizontal_alignment='center'},
        @{range='B2:E732';background_color='#FFF8E1'},
        @{range='F2:F732';background_color='#E8F1EC'},
        @{range='G2:J732';background_color='#FFF8E1'},
        @{range='K2:K732';background_color='#E8F1EC'},
        @{range='L2:P732';background_color='#FFF8E1'},
        @{range='Q2:Q732';background_color='#E8F1EC';horizontal_alignment='center'}
    ); row_sizes=@(@{range='1';type='pixel';size=44}); col_sizes=@(
        @{range='A';type='pixel';size=105},@{range='B:E';type='pixel';size=125},@{range='F';type='pixel';size=120},@{range='G:J';type='pixel';size=125},@{range='K:O';type='pixel';size=120},@{range='P';type='pixel';size=220},@{range='Q';type='pixel';size=90}
    )},
    @{name='月度补充'; cell_styles=@(
        (@{range='A1:I1'; background_color='#1F4D3B'; font_color='#FFFFFF'; font_weight='bold'; horizontal_alignment='center'; vertical_alignment='middle'; word_wrap='auto-wrap'}),
        @{range='A2:A25';background_color='#F2F4F3';horizontal_alignment='center'},
        @{range='B2:I25';background_color='#FFF8E1'},
        @{range='B3:B25';background_color='#E8F1EC'}
    ); row_sizes=@(@{range='1';type='pixel';size=52}); col_sizes=@(
        @{range='A';type='pixel';size=95},@{range='B:C';type='pixel';size=110},@{range='D:E';type='pixel';size=190},@{range='F:H';type='pixel';size=130},@{range='I';type='pixel';size=220}
    )},
    @{name='月度汇总'; cell_styles=@(
        (@{range='A1:X1'; background_color='#1F4D3B'; font_color='#FFFFFF'; font_weight='bold'; horizontal_alignment='center'; vertical_alignment='middle'; word_wrap='auto-wrap'}),
        @{range='A2:A25';background_color='#F2F4F3';horizontal_alignment='center'},
        @{range='B2:X25';background_color='#E8F1EC'},
        @{range='R2:R25';background_color='#DDEFE5';font_weight='bold'},
        @{range='T2:T25';background_color='#DDEFE5';font_weight='bold'},
        @{range='V2:V25';background_color='#E3EEF8';font_weight='bold'},
        @{range='X2:X25';background_color='#E3EEF8';font_weight='bold'}
    ); row_sizes=@(@{range='1';type='pixel';size=58}); col_sizes=@(
        @{range='A';type='pixel';size=90},@{range='B:K';type='pixel';size=115},@{range='L:O';type='pixel';size=165},@{range='P:X';type='pixel';size=120}
    )},
    @{name='资产与口径'; cell_styles=@(
        (@{range='A1:H1'; background_color='#1F4D3B'; font_color='#FFFFFF'; font_weight='bold'; horizontal_alignment='center'; vertical_alignment='middle'; word_wrap='auto-wrap'}),
        @{range='A2:E101';background_color='#FFF8E1'},
        @{range='F2:G101';background_color='#E8F1EC'},
        @{range='H2:H101';background_color='#FFF8E1'}
    ); row_sizes=@(@{range='1';type='pixel';size=44}); col_sizes=@(
        @{range='A';type='pixel';size=200},@{range='B';type='pixel';size=120},@{range='C';type='pixel';size=110},@{range='D';type='pixel';size=145},@{range='E';type='pixel';size=100},@{range='F:G';type='pixel';size=135},@{range='H';type='pixel';size=230}
    )}
)}

$sheetsJson = $sheetsPayload | ConvertTo-Json -Depth 20 -Compress
$stylesJson = $stylesPayload | ConvertTo-Json -Depth 20 -Compress
$sheetsJson | lark-cli sheets +workbook-create --as user --title '宠速星｜经营利润与现金流' --sheets - --styles $stylesJson --format json
