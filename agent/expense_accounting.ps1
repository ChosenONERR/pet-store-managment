function Get-PetstarExpenseSummary {
    param([object[]]$Entries)
    $totals = @{promotion=0.0;intervention=0.0;after_sales=0.0;other=0.0;excluded=0.0}
    $pending = [Collections.Generic.List[string]]::new()
    $managedKeys = @{}
    foreach ($entry in @($Entries)) {
        $project = ([string]$entry.project).Trim()
        if (-not $project) { continue }
        $orderNo = ([string]$entry.order_no).Trim()
        if ($orderNo -and $project -in @('补单平台扣款','补单配送费','营销赠品支出')) {
            $key = "$orderNo|$project"
            if ($managedKeys.ContainsKey($key)) { throw "重复经营费用：$key；禁止重复扣费" }
            $managedKeys[$key] = $true
        }
        $amount = 0.0
        $known = -not [string]::IsNullOrWhiteSpace([string]$entry.fee) -and
            [double]::TryParse(([string]$entry.fee).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$amount)
        if (-not $known -or [double]::IsNaN($amount) -or [double]::IsInfinity($amount)) {
            $pending.Add("$project $orderNo 金额待核对")
            continue
        }
        if ([string]$entry.source -match '待补录|待识别|待核对') { $pending.Add("$project $orderNo 来源待核对") }
        $category = if ($project -match '采购|补货付款|进货|老板取款|老板人工|市场租金|折旧|摊销|首次配送费|已冲减收入的退款') {
            $pending.Add("非经营或已扣费用须分离：$project")
            'excluded'
        } elseif ($project -match '补单|干预') { 'intervention' }
        elseif ($project -match '推广|广告|竞价') { 'promotion' }
        elseif ($project -match '售后.*配送|再次配送|二次配送|补送.*配送') { 'after_sales' }
        else { 'other' }
        $totals[$category] += $amount
    }
    $total = $totals.promotion + $totals.intervention + $totals.after_sales + $totals.other
    return [pscustomobject]@{
        total=[math]::Round($total,2);promotion=[math]::Round($totals.promotion,2)
        intervention=[math]::Round($totals.intervention,2);after_sales=[math]::Round($totals.after_sales,2)
        other=[math]::Round($totals.other,2);excluded=[math]::Round($totals.excluded,2)
        pending=@($pending);status=$(if($pending.Count){'利润暂估；'+($pending -join '；')}else{'已核对'})
    }
}

function Get-PetstarExpenseCellValue($Row,[int]$Index) {
    if($null -eq $Row -or $Index -ge @($Row).Count){return $null}
    $cell=@($Row)[$Index]
    if($null -eq $cell){return $null}
    $property=$cell.PSObject.Properties['value']
    if($null -ne $property){return $property.Value}
    return $null
}

function Invoke-PetstarExpenseCli {
    param([string]$Cli,[string[]]$Arguments,[string]$Payload)
    $global:LASTEXITCODE = 0
    $raw = if ($PSBoundParameters.ContainsKey('Payload')) { $Payload | & $Cli @Arguments | Out-String } else { & $Cli @Arguments | Out-String }
    if ($LASTEXITCODE -ne 0) { throw "经营费用读写失败：$raw" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok) { throw "经营费用读写失败：$raw" }
    return $result
}

function Get-PetstarExpenseRows {
    param([string]$Cli,[string]$Token,[string]$Sheet,[string]$Range)
    $result = Invoke-PetstarExpenseCli $Cli @('sheets','+cells-get','--as','user','--spreadsheet-token',$Token,'--sheet-name',$Sheet,'--range',$Range,'--include','value,formula','--max-chars','1000000','--format','json')
    if ($result.data.has_more -or $result.data.ranges[0].truncated) { throw "经营费用读取不完整：$Sheet" }
    return $result.data.ranges[0]
}

function Update-PetstarMonthExpenses {
    param([datetime]$Month,[string]$SourceToken,[string]$LedgerToken,[string]$Cli,$Transaction,[object[]]$Entries)
    if ($null -eq $Transaction) { throw '经营费用汇总必须加入财务事务' }
    if (-not $PSBoundParameters.ContainsKey('Entries')) {
        $prefix = if($Month.Year -eq 2026){"$($Month.Month)月"}else{"$($Month.Year)年$($Month.Month)月"}
        $range = Get-PetstarExpenseRows $Cli $LedgerToken "${prefix}支出" 'A2:F1001'
        $Entries = @(foreach($row in $range.cells) {
            if (-not (Get-PetstarExpenseCellValue $row 1)) { continue }
            $dateText = [string](Get-PetstarExpenseCellValue $row 0)
            $date = [datetime]::MinValue; $serial = 0.0
            if ([double]::TryParse($dateText,[ref]$serial) -and $serial -gt 30000 -and $serial -lt 100000) { $date=[datetime]::FromOADate($serial) }
            elseif (-not [datetime]::TryParse($dateText,[ref]$date)) { throw "支出日期无法识别：$dateText" }
            if ($date.ToString('yyyy-MM') -ne $Month.ToString('yyyy-MM')) { throw "支出表存在跨月记录：$dateText" }
            [pscustomobject]@{project=(Get-PetstarExpenseCellValue $row 1);fee=(Get-PetstarExpenseCellValue $row 2);order_no=(Get-PetstarExpenseCellValue $row 3);source=(Get-PetstarExpenseCellValue $row 4);note=(Get-PetstarExpenseCellValue $row 5)}
        })
    }
    foreach($entry in @($Entries)) {
        $dateProperty=$entry.PSObject.Properties['date']
        if($null -ne $dateProperty -and $entry.project) {
            if($null -eq $dateProperty.Value -or ([datetime]$dateProperty.Value).ToString('yyyy-MM') -ne $Month.ToString('yyyy-MM')){throw '月份支出存在缺失日期或跨月记录；未改费用汇总'}
        }
    }
    $summary = Get-PetstarExpenseSummary $Entries
    $months = Get-PetstarExpenseRows $Cli $SourceToken '月度结算' 'A1:R25'
    $headers = @('补单费用','售后追加配送费','其他经营费用明细','费用核对状态')
    $headersMissing=$false
    for($j=0;$j -lt 4;$j++) {
        $existing=[string](Get-PetstarExpenseCellValue $months.cells[0] (14+$j))
        if($existing -and $existing -ne $headers[$j]) { throw "月度结算费用辅助列已被占用：$existing" }
        if(-not $existing){$headersMissing=$true}
    }
    $matches = @(for($i=1;$i -lt $months.cells.Count;$i++) {
        $text=[string](Get-PetstarExpenseCellValue $months.cells[$i] 0)
        $d=[datetime]::MinValue; $serial=0.0
        if([double]::TryParse($text,[ref]$serial) -and $serial -gt 30000 -and $serial -lt 100000){$d=[datetime]::FromOADate($serial)}
        elseif(-not [datetime]::TryParse($text,[ref]$d)){continue}
        if($d.ToString('yyyy-MM') -eq $Month.ToString('yyyy-MM')){[int]$months.row_indices[$i]}
    })
    if($matches.Count -ne 1){throw '月度结算月份无法唯一定位'}
    $r=$matches[0]
    $writes=@(
        @{range='O1:R1';cells=,@($headers | ForEach-Object{@{value=$_;cell_styles=@{background_color='#173A2D';font_color='#FFFFFF';font_weight='bold';word_wrap='auto-wrap'}}})},
        @{range="G${r}:G${r}";cells=,@(@{value=$summary.total})},
        @{range="N${r}:R${r}";cells=,@(@{value=$summary.promotion},@{value=$summary.intervention},@{value=$summary.after_sales},@{value=$summary.other},@{value=$summary.status})}
    )
    foreach($cell in $writes[2].cells[0]) {
        $cell.cell_styles=@{background_color='#FFF8E1';vertical_alignment='middle';number_format='#,##0.00'}
    }
    $writes[2].cells[0][4].cell_styles=@{background_color='#FFF8E1';vertical_alignment='middle';word_wrap='auto-wrap'}
    foreach($write in $writes) {
        if($write.range -eq 'O1:R1' -and -not $headersMissing){continue}
        Register-FinanceRangeBeforeImage $Transaction 'source' $SourceToken '月度结算' $write.range $Cli
        $payload=ConvertTo-Json -InputObject $write.cells -Depth 10 -Compress
        $null=Invoke-PetstarExpenseCli $Cli @('sheets','+cells-set','--as','user','--spreadsheet-token',$SourceToken,'--sheet-name','月度结算','--range',$write.range,'--cells','-','--format','json') $payload
    }
    $check=Get-PetstarExpenseRows $Cli $SourceToken '月度结算' "G${r}:R${r}"
    foreach($pair in @(@(0,$summary.total),@(7,$summary.promotion),@(8,$summary.intervention),@(9,$summary.after_sales),@(10,$summary.other))) {
        $value=Get-PetstarExpenseCellValue $check.cells[0] $pair[0]
        if($null -eq $value -or [math]::Abs([double]$value-[double]$pair[1]) -gt 0.005){throw "经营费用汇总回读不一致：偏移$($pair[0]) 实际$value 预期$($pair[1])"}
    }
    if([string](Get-PetstarExpenseCellValue $check.cells[0] 11) -ne $summary.status){throw '费用核对状态回读不一致'}
    return $summary
}

function Get-PetstarDashboardExpenseFormulas {
    param([string]$MonthExpression,[string]$DeliveryExpression)
    $expressions=@{}
    foreach($pair in @(@('total','G'),@('promotion','N'),@('intervention','O'),@('after_sales','P'),@('other','Q'))) {
        $expressions[$pair[0]] = 'SUMIFS(''月度结算''!$'+$pair[1]+'$2:$'+$pair[1]+'$25,''月度结算''!$A$2:$A$25,'+$MonthExpression+')'
    }
    return @{
        gross='=C4-C5-('+$DeliveryExpression+')'
        profit='=C6-('+$expressions.total+')'
        gross_detail='="=线上与线下收入-商品成本-首次配送费="&TEXT(C4,"#,##0.00")&"-"&TEXT(C5,"#,##0.00")&"-"&TEXT(('+ $DeliveryExpression +'),"#,##0.00")&"="&TEXT(C6,"#,##0.00")'
        profit_detail='="=经营毛利-推广费-补单费用-售后追加配送费-其他经营费用="&TEXT(C6,"#,##0.00")&"-"&TEXT(('+ $expressions.promotion +'),"#,##0.00")&"-"&TEXT(('+ $expressions.intervention +'),"#,##0.00")&"-"&TEXT(('+ $expressions.after_sales +'),"#,##0.00")&"-"&TEXT(('+ $expressions.other +'),"#,##0.00")&"="&TEXT(C7,"#,##0.00")'
        status='=IFERROR(INDEX(''月度结算''!$R$2:$R$25,MATCH('+$MonthExpression+',''月度结算''!$A$2:$A$25,0)),"利润暂估；费用尚未核对")'
    }
}
