param([switch]$Apply)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$metaPath = Join-Path $PSScriptRoot 'simple_finance_meta.json'
$meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
if (-not $spreadsheetToken) { throw '无法解析飞书工作簿引用' }

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}

function Get-SheetRows([string]$SheetName, [string]$Range) {
    $raw = lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --include value --max-chars 500000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取飞书失败：$SheetName $Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "飞书读取不完整：$SheetName $Range" }
    return @($result.data.ranges[0].cells)
}

function New-ValueCell($Value) {
    if ($null -eq $Value) { return @{value=''} }
    return @{value=$Value}
}

function Set-CellMatrix([string]$SheetName, [string]$Range, $Rows) {
    $payload = ConvertTo-Json -InputObject $Rows -Depth 12 -Compress
    $raw = $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "写入飞书失败：$SheetName $Range $raw" }
}

$personPairs = @(
    @('姐','3402269981653848673'), @('一宁','3402270191343933411'), @('熊佳明','3402270270828412572'),
    @('李嘉奇','3402270470443155017'), @('张毅斌','3402270401483361644'), @('wikar','3202270563195646541'),
    @('李栩琪','3402270610009247219'), @('淑琪周','3402271532991943565'), @('妈妈','5302271762200734156'),
    @('遇见公寓老板','3402273224126547967'), @('郑海龙','3402272960303924940'), @('黄增添','3402273372661176866'),
    @('瑞丽菲林小姐','3402272993641635876'), @('姐姐','3402274460217435837'), @('李嘉奇','3402274781101310940'),
    @('wikar1乐先生','3202274664042626373'), @('杨梦情','2902275803792480143'), @('熊佳明','3402276011411606776'),
    @('吴钦先','3002276431320343097'), @('一宁','3402276222349855974'), @('赖富威','3402277263561610001'),
    @('陈根培','3402277542344370409'), @('妈妈','5302277981079431084'), @('郑海龙','3402278740071437615'),
    @('遇见公寓老板','3402278981316773812'), @('黄炫堃','3402279281868142897'), @('爸爸','3402278780658329499'),
    @('罗润发','3402280413723387549'), @('wikar1任-','3402280552056169204'), @('发哥老婆','3402280792848020666'),
    @('骑手 - 一路有李、','3402281682301110868'), @('骑手 - 一路有李、','3202281573109753390'), @('骑手 - 一路有李、','3402281842060426471'),
    @('李栩琪','3402282093659273951'), @('姐姐','3402283101653848673'), @('李嘉奇','3402283480179553935'),
    @('wikar','3202283593562353533'), @('杨梦情','2902283713609308571'), @('黄健勇','3402284482495345991'),
    @('妈妈','5302284540016891722'), @('熊佳明','3402284880828412572'), @('张毅斌','3402285160909284892'),
    @('林女士','3402285910676229064'), @('爸爸','3402285972819618069'), @('罗润发','3402286311190051916'),
    @('吴钦先','3002286461320343097'), @('赖富威','3402287391681847231'), @('黄炫堃','3402287610068617497'),
    @('一宁','3402287751621698940'), @('刀疤老婆','3402287921383791963'), @('姐姐','3402288972032515706'),
    @('黄健勇','3402290493192803160'), @('妈妈','5302290301362038118'), @('李嘉奇','3402290543464409264'),
    @('罗润丰','3402293480111496655'), @('熊佳明','3402293091985665554')
)
$personByOrder = @{}
foreach ($pair in $personPairs) { $personByOrder[[string]$pair[1]] = [string]$pair[0] }

$skuAliasToKw = @{}
foreach ($row in (Get-SheetRows 'SKU成本主表' 'A2:B3000')) {
    $kw = Get-CellValue $row 0
    if (-not $kw -or $kw -notmatch '(?i)^kw\d+$') { continue }
    $kw = $kw.ToLowerInvariant()
    $skuAliasToKw[$kw] = $kw
    $aliases = Get-CellValue $row 1
    if ($aliases) {
        foreach ($alias in ([regex]::Split($aliases, '[/,，;；\s]+') | Where-Object { $_ })) {
            $skuAliasToKw[$alias.ToLowerInvariant()] = $kw
        }
    }
}

$rows = Get-SheetRows '干预订单' 'A2:M1001'
$activeRows = @($rows | Where-Object { Get-CellValue $_ 3 })
$personCells = @()
$contentCells = @()
$matchedPeople = 0
$unmatchedPeople = [Collections.Generic.List[string]]::new()
$unmappedCodes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$compactContentCount = 0
$dailyNoCount = 0

foreach ($row in $activeRows) {
    $orderNo = Get-CellValue $row 3
    $person = Get-CellValue $row 2
    if ($personByOrder.ContainsKey($orderNo)) {
        $person = $personByOrder[$orderNo]
        $matchedPeople++
    }

    $contents = Get-CellValue $row 11
    if ($contents) {
        $quantityByKw = [ordered]@{}
        foreach ($match in [regex]::Matches($contents, '(?i)\b((?:kw|csx)\d+)\s*[×xX*]\s*(\d+(?:\.\d+)?)')) {
            $rawCode = $match.Groups[1].Value.ToLowerInvariant()
            $kw = if ($rawCode -match '^kw\d+$') { $rawCode } elseif ($skuAliasToKw.ContainsKey($rawCode)) { $skuAliasToKw[$rawCode] } else { [void]$unmappedCodes.Add($rawCode); $rawCode }
            $quantity = [double]::Parse($match.Groups[2].Value, [Globalization.CultureInfo]::InvariantCulture)
            if (-not $quantityByKw.Contains($kw)) { $quantityByKw[$kw] = 0.0 }
            $quantityByKw[$kw] += $quantity
        }
        if ($quantityByKw.Count -gt 0) {
            $contents = (@($quantityByKw.GetEnumerator() | ForEach-Object {
                $qty = if ([math]::Abs($_.Value - [math]::Round($_.Value)) -lt 0.000001) { [string][int][math]::Round($_.Value) } else { $_.Value.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) }
                "$($_.Key)×$qty"
            }) -join '、')
            $compactContentCount++
        }
    }

    # “美团N号单”只能来自美团订单卡片显示的 #N，严禁按订单编号或行序推算。
    if (Get-CellValue $row 1) { $dailyNoCount++ }

    $personCells += ,@((New-ValueCell $person))
    $contentCells += ,@((New-ValueCell $contents))
}

foreach ($orderNo in $personByOrder.Keys) {
    if (-not ($activeRows | Where-Object { (Get-CellValue $_ 3) -eq $orderNo })) { $unmatchedPeople.Add($orderNo) }
}

if ($Apply) {
    if ($activeRows.Count -gt 0) {
        $lastRow = $activeRows.Count + 1
        Set-CellMatrix '干预订单' "C2:C$lastRow" $personCells
        Set-CellMatrix '干预订单' "L2:L$lastRow" $contentCells
    }
    $null = lark-cli sheets +dropdown-delete --as user --spreadsheet-token $spreadsheetToken --ranges '["干预订单!C2:C1001","干预订单!L2:L1001"]' --yes --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '移除补单人/购买内容下拉失败' }
    $null = lark-cli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'A1:M1' --background-color '#7A2E2E' --font-color '#FFFFFF' --font-weight bold --horizontal-alignment center --word-wrap auto-wrap --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '设置表头样式失败' }
    $null = lark-cli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --range 'B2:B1001' --background-color '#FFF8E1' --word-wrap auto-wrap --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '设置单号列样式失败' }
    $null = lark-cli sheets +cols-resize --as user --spreadsheet-token $spreadsheetToken --sheet-name '干预订单' --widths '{"B:C":120,"L":220}' --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw '设置列宽失败' }
}

[pscustomobject]@{
    applied = [bool]$Apply
    intervention_rows = $activeRows.Count
    supplied_people = $personByOrder.Count
    matched_people = $matchedPeople
    missing_orders_for_people = @($unmatchedPeople)
    compact_content_rows = $compactContentCount
    unmapped_codes = @($unmappedCodes)
    daily_order_labels = $dailyNoCount
} | ConvertTo-Json -Depth 6 -Compress
