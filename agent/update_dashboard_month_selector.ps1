param(
    [string]$ProjectRoot,
    [ValidatePattern('^\d{4}-\d{2}$')][string]$SelectedMonth,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$agentRoot = Join-Path $ProjectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$metaPath = Join-Path $ProjectRoot 'work\simple_finance_meta.json'
$statePath = Join-Path $stateRoot 'dashboard_month_selector.json'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$labels = @'
{
  "dashboard": "\u9a7e\u9a76\u8231",
  "selector_label": "\u67e5\u770b\u6708\u4efd",
  "system_gross_profit": "\u7cfb\u7edf\u5f53\u6708\u7ecf\u8425\u6bdb\u5229"
}
'@ | ConvertFrom-Json

if (-not (Test-Path -LiteralPath $metaPath -PathType Leaf)) { throw 'Missing simple_finance_meta.json' }
$meta = [IO.File]::ReadAllText($metaPath,[Text.Encoding]::UTF8) | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$currentMonth = [datetime]::Today.ToString('yyyy-MM')
$createdMonth = ([datetime]$meta.created_on).ToString('yyyy-MM')

$options = [Collections.Generic.List[string]]::new()
$cursor = [datetime]::ParseExact("$createdMonth-01",'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
$end = [datetime]::ParseExact("$currentMonth-01",'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
while ($cursor -le $end) {
    $options.Add($cursor.ToString('yyyy-MM'))
    $cursor = $cursor.AddMonths(1)
}
if ($SelectedMonth -and -not $options.Contains($SelectedMonth)) { $options.Add($SelectedMonth) }
$optionArray = @($options | Sort-Object -Unique)
$optionsSignature = $optionArray -join '|'

if (-not $Force -and -not $SelectedMonth -and (Test-Path -LiteralPath $statePath -PathType Leaf)) {
    try {
        $state = [IO.File]::ReadAllText($statePath,[Text.Encoding]::UTF8) | ConvertFrom-Json
        if ([string]$state.options_signature -eq $optionsSignature) {
            [ordered]@{ok=$true;skipped=$true;reason='month options unchanged';options=$optionArray} | ConvertTo-Json -Depth 5
            exit 0
        }
    } catch {}
}

function Invoke-LarkJson([scriptblock]$Action,[string]$Message) {
    $global:LASTEXITCODE = 0
    $raw = @(& $Action 2>&1 | ForEach-Object { [string]$_ }) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "$Message`: $raw" }
    try { $result = $raw | ConvertFrom-Json } catch { throw "$Message`: invalid JSON response" }
    if (-not $result.ok) { throw "$Message`: $raw" }
    return $result
}

function Get-CellValue($Row,[int]$Index,[string]$Property='value') {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell) { return $null }
    $propertyValue = $cell.PSObject.Properties[$Property]
    if ($null -eq $propertyValue) { return $null }
    return $propertyValue.Value
}

function Set-Cells([string]$Range,$Cells,[string]$Message) {
    $payload = ConvertTo-Json -InputObject $Cells -Depth 12 -Compress
    $null = Invoke-LarkJson { $payload | & $larkCli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range $Range --cells - --format json } $Message
}

function Convert-ToSelectorFormula([string]$Formula) {
    if (-not $Formula) { return $Formula }
    return $Formula.Replace('DATE(YEAR(TODAY()),MONTH(TODAY()),1)',$script:selectorMonthExpression)
}

function Convert-ToCurrentMonthFormula([string]$Formula) {
    if (-not $Formula) { return $Formula }
    return $Formula.Replace($script:selectorMonthExpression,'DATE(YEAR(TODAY()),MONTH(TODAY()),1)')
}

$script:selectorMonthExpression = 'DATE(VALUE(LEFT($I$1,4)),VALUE(RIGHT($I$1,2)),1)'
$dashboard = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range 'A1:J19' --include 'value,formula' --max-chars 200000 --format json } 'Read dashboard failed'
if ($dashboard.data.has_more) { throw 'Dashboard read was incomplete' }
$rows = @($dashboard.data.ranges[0].cells)
if ($rows.Count -lt 14) { throw 'Dashboard layout is incomplete' }

$existingLabel = [string](Get-CellValue $rows[0] 7)
$existingSelection = [string](Get-CellValue $rows[0] 8)
if ($existingLabel -and $existingLabel -ne [string]$labels.selector_label) { throw "Dashboard H1 is already occupied: $existingLabel" }
$effectiveSelection = if ($SelectedMonth) { $SelectedMonth } elseif ($existingSelection -and $existingSelection -match '^\d{4}-\d{2}$') { $existingSelection } else { $currentMonth }
if ($effectiveSelection -notin $optionArray) { throw "Selected month is outside available options: $effectiveSelection" }

$monthlyRows = @(2,3,4,5,6,7,8,9,10,14)
$formulaUpdates = [ordered]@{}
foreach ($rowNumber in $monthlyRows) {
    $formula = [string](Get-CellValue $rows[$rowNumber - 1] 2 'formula')
    if (-not $formula) { throw "Dashboard C$rowNumber has no formula" }
    $formulaUpdates[[string]$rowNumber] = Convert-ToSelectorFormula $formula
}
$detailRows = @(6,7,9,10)
$detailUpdates = [ordered]@{}
foreach ($rowNumber in $detailRows) {
    $formula = [string](Get-CellValue $rows[$rowNumber - 1] 3 'formula')
    if (-not $formula) { throw "Dashboard D$rowNumber has no formula" }
    $detailUpdates[[string]$rowNumber] = Convert-ToSelectorFormula $formula
}

$systemGrossFormula = [string](Get-CellValue $rows[1] 9 'formula')
if (-not $systemGrossFormula) { $systemGrossFormula = Convert-ToCurrentMonthFormula $formulaUpdates['6'] }

. (Join-Path $agentRoot 'finance_transaction.ps1')
$transaction = New-FinanceTransaction $stateRoot 'dashboard month selector' ('TX-DASHBOARD-MONTH-' + [datetime]::Now.ToString('yyyyMMdd-HHmmss')) -AcquireLock -WorkbookTokens @{source=$spreadsheetToken} -LarkCli $larkCli
$closed = $false
try {
    Register-FinanceRangeBeforeImage $transaction 'source' $spreadsheetToken $labels.dashboard 'C2:D19' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $spreadsheetToken $labels.dashboard 'H1:J2' $larkCli

    foreach ($rowNumber in $monthlyRows) {
        Set-Cells "C${rowNumber}:C${rowNumber}" @(@(@{formula=$formulaUpdates[[string]$rowNumber]})) "Write dashboard C$rowNumber failed"
    }
    foreach ($rowNumber in $detailRows) {
        Set-Cells "D${rowNumber}:D${rowNumber}" @(@(@{formula=$detailUpdates[[string]$rowNumber]})) "Write dashboard D$rowNumber failed"
    }
    Set-Cells 'H1:I1' @(@(@{value=[string]$labels.selector_label},@{value=$effectiveSelection})) 'Write month selector failed'
    Set-Cells 'J1:J2' @(@(@{value=[string]$labels.system_gross_profit}),@(@{formula=$systemGrossFormula})) 'Write system current-month metric failed'

    $optionsJson = $optionArray | ConvertTo-Json -Compress
    $null = Invoke-LarkJson { $optionsJson | & $larkCli sheets +dropdown-set --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range 'I1:I1' --options - --highlight=false --format json } 'Set month dropdown failed'
    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range 'H1:H1' --background-color '#173A2D' --font-color '#FFFFFF' --font-weight bold --horizontal-alignment center --vertical-alignment middle --format json } 'Style month selector label failed'
    $null = Invoke-LarkJson { & $larkCli sheets +cells-set-style --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range 'I1:I1' --background-color '#E3F0E8' --font-weight bold --horizontal-alignment center --vertical-alignment middle --number-format '@' --format json } 'Style month selector value failed'
    $null = Invoke-LarkJson { & $larkCli sheets +cols-resize --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range 'H:I' --width 120 --format json } 'Resize selector columns failed'
    $null = Invoke-LarkJson { & $larkCli sheets +dim-hide --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range 'J:J' --format json } 'Hide system metric column failed'

    foreach ($range in @('C2:C10','C14:C14','D6:D7','D9:D10','J2:J2')) {
        $verify = Invoke-LarkJson { & $larkCli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range $range --format json --exit-on-error } "Verify formulas failed: $range"
        if ($verify.data.status -ne 'success') { throw "Formula verification failed: $range" }
    }
    $readback = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $labels.dashboard --range 'H1:J2' --include 'value,formula' --max-chars 10000 --format json } 'Read back month selector failed'
    $readRows = @($readback.data.ranges[0].cells)
    if ([string](Get-CellValue $readRows[0] 0) -ne [string]$labels.selector_label -or [string](Get-CellValue $readRows[0] 1) -ne $effectiveSelection) { throw 'Month selector readback mismatch' }

    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true

    if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
    $statePayload = [ordered]@{options_signature=$optionsSignature;options=$optionArray;updated_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')}
    [IO.File]::WriteAllText($statePath,($statePayload | ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
    [ordered]@{ok=$true;selected_month=$effectiveSelection;options=$optionArray;system_current_month=$currentMonth;dashboard_url=$meta.workbook_url} | ConvertTo-Json -Depth 5
} catch {
    if (-not $closed) {
        try { $null = Invoke-FinanceTransactionRollback $transaction @{source=$spreadsheetToken} $larkCli $_.Exception.Message } catch {}
        try { Close-FinanceTransaction $transaction } catch {}
    }
    throw
}
