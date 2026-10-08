param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent),
    [datetime]$Date = [datetime]'2026-09-22',
    [double]$Amount = 29.76,
    [string]$Project = '投流推广费'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$sourceMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\daily_ledger_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$larkCli = Join-Path $ProjectRoot 'agent\lark_cli_noexit.ps1'
$stateRoot = Join-Path $ProjectRoot 'agent\state'
. (Join-Path $ProjectRoot 'agent\finance_transaction.ps1')

function Get-CellValue($Row, [int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return [string]$cell.value
}
function To-Number($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return 0.0 }
    $number = 0.0
    if (-not [double]::TryParse(([string]$Value).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) { throw "金额格式错误：$Value" }
    return $number
}
function To-Date($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)) { return $parsed.Date }
    $number = 0.0
    if ([double]::TryParse([string]$Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) { return ([datetime]'1899-12-30').AddDays($number).Date }
    return $null
}
function Normalize-Text([string]$Value) {
    if (-not $Value) { return '' }
    return (($Value.Trim().ToLowerInvariant()) -replace '[\s，,。；;：:]+','')
}
function Get-Rows([string]$Token,[string]$Sheet,[string]$Range) {
    $raw = & $larkCli sheets +cells-get --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --include value --max-chars 5000000 --format json | Out-String
    if ($LASTEXITCODE -ne 0) { throw "读取失败：$Sheet $Range" }
    $result = $raw | ConvertFrom-Json
    if (-not $result.ok -or $result.data.has_more) { throw "读取不完整：$Sheet $Range" }
    return @($result.data.ranges[0].cells)
}
function Set-Row([string]$WorkbookKey,[string]$Token,[string]$Sheet,[string]$Range,[object[]]$Values,$Transaction) {
    Register-FinanceRangeBeforeImage $Transaction $WorkbookKey $Token $Sheet $Range $larkCli
    $row = @()
    foreach ($value in $Values) {
        if ($null -eq $value -or [string]$value -eq '') { $row += ,@{} } else { $row += ,@{value=$value} }
    }
    $matrix = @(); $matrix += ,$row
    $payload = ConvertTo-Json -InputObject $matrix -Depth 10 -Compress
    $raw = $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $Token --sheet-name $Sheet --range $Range --cells - --format json | Out-String
    if ($LASTEXITCODE -ne 0 -or -not (($raw | ConvertFrom-Json).ok)) { throw "写入失败：$Sheet $Range" }
}
function Get-FirstEmptyRow($Rows) {
    for ($index = 0; $index -lt @($Rows).Count; $index++) {
        $hasValue = $false
        for ($column = 0; $column -lt 6; $column++) {
            if (Get-CellValue $Rows[$index] $column) { $hasValue = $true; break }
        }
        if (-not $hasValue) { return 2 + $index }
    }
    throw '9月支出没有可用空行'
}

$workbooks = @{source=$sourceToken;ledger=$ledgerToken}
$transaction = New-FinanceTransaction $stateRoot '修复2026-09-22投流推广费重复汇总与支出漏记' $null -AcquireLock -WorkbookTokens $workbooks -LarkCli $larkCli
$closed = $false
try {
    $monthRows = Get-Rows $sourceToken '月度结算' 'A2:M25'
    $monthIndex = -1
    for ($index = 0; $index -lt $monthRows.Count; $index++) {
        if ([string](Get-CellValue $monthRows[$index] 0) -like "$($Date.ToString('yyyy-MM'))*") { $monthIndex = $index; break }
    }
    if ($monthIndex -lt 0) { throw "月度结算找不到 $($Date.ToString('yyyy-MM'))" }
    $monthRowNumber = 2 + $monthIndex
    $monthRow = $monthRows[$monthIndex]
    $entry = "$($Date.ToString('MM-dd')) 其他经营费用 $Amount：$Project"
    $entryKey = Normalize-Text $entry
    $noteParts = @([regex]::Split([string](Get-CellValue $monthRow 11),'[；;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $matchingNotes = @($noteParts | Where-Object { (Normalize-Text $_) -eq $entryKey })
    $otherNotes = @($noteParts | Where-Object { (Normalize-Text $_) -ne $entryKey })
    $currentMonthlyValue = To-Number (Get-CellValue $monthRow 6)
    $newMonthlyValue = if ($matchingNotes.Count -eq 0) {
        [math]::Round($currentMonthlyValue + $Amount,2)
    } else {
        [math]::Round($currentMonthlyValue - ($Amount * ($matchingNotes.Count - 1)),2)
    }
    $newMonthlyNote = (@($otherNotes + $entry) -join '；')

    $expenseSheet = if ($Date.Year -eq 2026) { "$($Date.Month)月支出" } else { "$($Date.Year)年$($Date.Month)月支出" }
    $expenseRows = Get-Rows $ledgerToken $expenseSheet 'A2:F1001'
    $ledgerMatches = [Collections.Generic.List[int]]::new()
    for ($index = 0; $index -lt $expenseRows.Count; $index++) {
        $row = $expenseRows[$index]
        $rowDate = To-Date (Get-CellValue $row 0)
        $rowAmount = To-Number (Get-CellValue $row 2)
        if ($null -ne $rowDate -and $rowDate.Date -eq $Date.Date -and [math]::Abs($rowAmount-$Amount) -le 0.0001 -and (Normalize-Text (Get-CellValue $row 1)) -eq (Normalize-Text $Project)) {
            $ledgerMatches.Add(2+$index)
        }
    }
    if ($ledgerMatches.Count -gt 1) { throw "支出明细已存在 $($ledgerMatches.Count) 条相同推广费；拒绝自动判断保留哪一条" }
    $expenseRowNumber = if ($ledgerMatches.Count -eq 1) { $ledgerMatches[0] } else { Get-FirstEmptyRow $expenseRows }

    Set-Row 'source' $sourceToken '月度结算' "G${monthRowNumber}:G${monthRowNumber}" @($newMonthlyValue) $transaction
    Set-Row 'source' $sourceToken '月度结算' "L${monthRowNumber}:L${monthRowNumber}" @($newMonthlyNote) $transaction
    if ($ledgerMatches.Count -eq 0) {
        $serial = ($Date.Date - [datetime]'1899-12-30').TotalDays
        Set-Row 'ledger' $ledgerToken $expenseSheet "A${expenseRowNumber}:F${expenseRowNumber}" @($serial,$Project,$Amount,$null,'手工','通过飞书确认录入') $transaction
    }

    $checkMonthRows = Get-Rows $sourceToken '月度结算' 'A2:M25'
    $checkMonthRow = @($checkMonthRows)[$monthRowNumber-2]
    $checkValue = To-Number (Get-CellValue $checkMonthRow 6)
    $checkNote = Get-CellValue $checkMonthRow 11
    $checkNoteCount = @([regex]::Split([string]$checkNote,'[；;]') | Where-Object { (Normalize-Text $_) -eq $entryKey }).Count
    $checkExpense = Get-Rows $ledgerToken $expenseSheet 'A2:F1001'
    $checkLedgerCount = 0
    foreach ($row in $checkExpense) {
        $rowDate = To-Date (Get-CellValue $row 0)
        if ($null -ne $rowDate -and $rowDate.Date -eq $Date.Date -and [math]::Abs((To-Number (Get-CellValue $row 2))-$Amount) -le 0.0001 -and (Normalize-Text (Get-CellValue $row 1)) -eq (Normalize-Text $Project)) { $checkLedgerCount++ }
    }
    if ([math]::Abs($checkValue-$newMonthlyValue) -gt 0.0001 -or $checkNoteCount -ne 1 -or $checkLedgerCount -ne 1) {
        throw "修复回读不一致：月度值=$checkValue，月度备注数=$checkNoteCount，支出明细数=$checkLedgerCount"
    }
    $formulaRaw = & $larkCli sheets +formula-verify --as user --spreadsheet-token $sourceToken --sheet-name '驾驶舱' --range 'C2:C19' --format json --exit-on-error | Out-String
    if ($LASTEXITCODE -ne 0 -or (($formulaRaw | ConvertFrom-Json).data.status -ne 'success')) { throw '驾驶舱公式校验失败' }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok = $true
        date = $Date.ToString('yyyy-MM-dd')
        project = $Project
        amount = $Amount
        month_row = $monthRowNumber
        month_value_before = $currentMonthlyValue
        month_value_after = $newMonthlyValue
        duplicate_month_notes_removed = [math]::Max(0,$matchingNotes.Count-1)
        expense_sheet = $expenseSheet
        expense_row = $expenseRowNumber
        expense_inserted = ($ledgerMatches.Count -eq 0)
    } | ConvertTo-Json -Compress
} catch {
    $reason = [string]$_.Exception.Message
    $rollback = Invoke-FinanceTransactionRollback $transaction $workbooks $larkCli $reason
    if ($rollback.ok) { throw "$reason；本次修复已全部回滚" }
    throw "$reason；自动回滚待恢复：$($rollback.error)"
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
