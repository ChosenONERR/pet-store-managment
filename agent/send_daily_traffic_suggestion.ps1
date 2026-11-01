param(
    [datetime]$Date = [datetime]::Today,
    [string]$ProjectRoot,
    [switch]$DryRun,
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
$configPath = Join-Path $agentRoot 'traffic_suggestion_rules.json'
$historyPath = Join-Path $stateRoot 'traffic_suggestion_history.json'
$identityPath = Join-Path $stateRoot 'lark_user_identity.json'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
if (-not (Test-Path -LiteralPath $metaPath -PathType Leaf)) { throw 'Missing simple_finance_meta.json' }
if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { throw 'Missing traffic_suggestion_rules.json' }
if (-not (Test-Path -LiteralPath $identityPath -PathType Leaf)) { throw 'Missing cached Lark recipient identity' }

$meta = [IO.File]::ReadAllText($metaPath,[Text.Encoding]::UTF8) | ConvertFrom-Json
$config = [IO.File]::ReadAllText($configPath,[Text.Encoding]::UTF8) | ConvertFrom-Json
$identity = [IO.File]::ReadAllText($identityPath,[Text.Encoding]::UTF8) | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
$openId = [string]$identity.open_id
if (-not $openId) { throw 'Cached Lark recipient identity has no open_id' }
$dateKey = $Date.ToString('yyyy-MM-dd')

function Invoke-LarkJson([scriptblock]$Action,[string]$Message,[int]$MaxAttempts=3) {
    $last = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $global:LASTEXITCODE = 0
        $raw = @(& $Action 2>&1 | ForEach-Object { [string]$_ }) -join "`n"
        try { $result = $raw | ConvertFrom-Json } catch { $result = $null }
        if ($LASTEXITCODE -eq 0 -and $null -ne $result -and $result.ok) { return $result }
        $last = $raw
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds ([math]::Min(8,2*$attempt)) }
    }
    throw "$Message`: $last"
}

function Get-CellValue($Row,[int]$Index) {
    if ($null -eq $Row -or $Index -ge @($Row).Count) { return $null }
    $cell = @($Row)[$Index]
    if ($null -eq $cell -or $null -eq $cell.value) { return $null }
    return $cell.value
}

function Convert-ToNumber($Value) {
    $number = 0.0
    if ($null -ne $Value -and [double]::TryParse(([string]$Value).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) { return $number }
    return 0.0
}

function Convert-ToDate($Value) {
    if ($null -eq $Value -or -not ([string]$Value).Trim()) { return $null }
    $number = Convert-ToNumber $Value
    if ($number -gt 30000 -and $number -lt 100000) { return [datetime]::FromOADate($number) }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value,[ref]$parsed)) { return $parsed }
    return $null
}

$history = [pscustomobject]@{items=@()}
if (Test-Path -LiteralPath $historyPath -PathType Leaf) {
    try { $history = [IO.File]::ReadAllText($historyPath,[Text.Encoding]::UTF8) | ConvertFrom-Json } catch {}
}
$existingToday = @($history.items | Where-Object { [string]$_.date -eq $dateKey })
if ($existingToday.Count -gt 0 -and -not $Force) {
    [ordered]@{ok=$true;skipped=$true;reason='suggestion already sent';suggestion=$existingToday[-1]} | ConvertTo-Json -Depth 8
    exit 0
}

$inventoryResult = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $config.inventory_sheet --range 'A2:O2000' --include value --max-chars 1000000 --format json } 'Read inventory failed'
$orderResult = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name $config.order_items_sheet --range 'A2:I4001' --include value --max-chars 1000000 --format json } 'Read order items failed'
if ($inventoryResult.data.has_more -or $orderResult.data.has_more) { throw 'Lark read was incomplete' }
$inventoryRows = @($inventoryResult.data.ranges[0].cells)
$orderRows = @($orderResult.data.ranges[0].cells)

$latestInventoryDate = [datetime]::MinValue
foreach ($row in $inventoryRows) {
    $rowDate = Convert-ToDate (Get-CellValue $row 0)
    if ($null -ne $rowDate -and $rowDate -gt $latestInventoryDate) { $latestInventoryDate = $rowDate }
}
if ($latestInventoryDate -eq [datetime]::MinValue) { throw 'Inventory snapshot has no valid date' }

$naturalSales = @{}
$windowStart = $Date.Date.AddDays(-29)
foreach ($row in $orderRows) {
    $rowDate = Convert-ToDate (Get-CellValue $row 0)
    if ($null -eq $rowDate -or $rowDate.Date -lt $windowStart -or $rowDate.Date -gt $Date.Date) { continue }
    if ([string](Get-CellValue $row 3) -eq [string]$config.cancelled_order_status) { continue }
    if ([string](Get-CellValue $row 4) -ne [string]$config.natural_order_type) { continue }
    $code = ([string](Get-CellValue $row 5)).Trim().ToLowerInvariant()
    if (-not $code) { continue }
    if (-not $naturalSales.ContainsKey($code)) { $naturalSales[$code] = 0.0 }
    $naturalSales[$code] += Convert-ToNumber (Get-CellValue $row 8)
}

$candidates = [Collections.Generic.List[object]]::new()
$fallbackCandidates = [Collections.Generic.List[object]]::new()
foreach ($row in $inventoryRows) {
    $snapshotDate = Convert-ToDate (Get-CellValue $row 0)
    if ($null -eq $snapshotDate -or $snapshotDate.Date -ne $latestInventoryDate.Date) { continue }
    $code = ([string](Get-CellValue $row 1)).Trim().ToLowerInvariant()
    $name = ([string](Get-CellValue $row 2)).Trim()
    $stock = Convert-ToNumber (Get-CellValue $row 3)
    $monthSales = Convert-ToNumber (Get-CellValue $row 4)
    $status = [string](Get-CellValue $row 5)
    if (-not $code -or -not $name -or $stock -le 0 -or $status -ne [string]$config.selling_status) { continue }
    $globallyExcluded = $false
    foreach ($pattern in @($config.global_exclude)) { if ($name -like "*$pattern*") { $globallyExcluded = $true; break } }
    if ($globallyExcluded) { continue }
    $recentNatural = if ($naturalSales.ContainsKey($code)) { [double]$naturalSales[$code] } else { 0.0 }
    $lastHistory = @($history.items | Where-Object { [string]$_.code -eq $code } | Sort-Object date -Descending | Select-Object -First 1)
    $daysSince = if ($lastHistory.Count -gt 0) { ($Date.Date - ([datetime]$lastHistory[0].date).Date).TotalDays } else { 999.0 }
    $score = ([math]::Min($stock,100.0) * 0.2) + ([math]::Min($monthSales,30.0) * 0.1) + ([math]::Min($recentNatural,30.0) * 0.05) + ([math]::Min($daysSince,60.0) * 0.5)
    $fallbackCandidates.Add([pscustomobject]@{code=$code;name=$name;stock=$stock;month_sales=$monthSales;recent_natural=$recentNatural;rule_index=-1;keyword=[string]$config.fallback_keyword;action=[string]$config.fallback_action;score=$score})
    for ($ruleIndex = 0; $ruleIndex -lt @($config.rules).Count; $ruleIndex++) {
        $rule = @($config.rules)[$ruleIndex]
        $matched = $false
        foreach ($pattern in @($rule.include)) { if ($name -like "*$pattern*") { $matched = $true; break } }
        if (-not $matched) { continue }
        foreach ($pattern in @($rule.exclude)) { if ($name -like "*$pattern*") { $matched = $false; break } }
        if (-not $matched) { continue }
        $candidates.Add([pscustomobject]@{code=$code;name=$name;stock=$stock;month_sales=$monthSales;recent_natural=$recentNatural;rule_index=$ruleIndex;keyword=[string]$rule.keyword;action=[string]$rule.action;score=$score})
    }
}
if ($candidates.Count -eq 0) {
    foreach ($candidate in $fallbackCandidates) { $candidates.Add($candidate) }
}
if ($candidates.Count -eq 0) { throw 'No in-stock selling product is available for the daily traffic suggestion' }

$primaryRule = $Date.DayOfYear % @($config.rules).Count
$selected = $null
for ($offset = 0; $offset -lt @($config.rules).Count; $offset++) {
    $ruleIndex = ($primaryRule + $offset) % @($config.rules).Count
    $pool = @($candidates | Where-Object { $_.rule_index -eq $ruleIndex } | Sort-Object @{Expression='score';Descending=$true},code)
    if ($pool.Count -gt 0) { $selected = $pool[0]; break }
}
if ($null -eq $selected) { $selected = @($candidates | Sort-Object @{Expression='score';Descending=$true},code)[0] }

$stockText = if ([math]::Abs($selected.stock - [math]::Round($selected.stock)) -lt 0.000001) { [string][int][math]::Round($selected.stock) } else { '{0:0.##}' -f $selected.stock }
$salesText = if ([math]::Abs($selected.recent_natural - [math]::Round($selected.recent_natural)) -lt 0.000001) { [string][int][math]::Round($selected.recent_natural) } else { '{0:0.##}' -f $selected.recent_natural }
$message = @(
    [string]$config.message_title,
    "\u641c\u7d22\u8bcd\uff1a$($selected.keyword)",
    "\u5546\u54c1\uff1a$($selected.code)\uff5c$($selected.name)",
    "\u5e93\u5b58\uff1a$stockText",
    "\u53c2\u8003\uff1a\u8fd130\u5929\u81ea\u7136\u552e\u51fa$salesText\u4ef6\uff08$($config.recent_sales_note)\uff09",
    "\u4eca\u5929\uff1a$($selected.action)"
) -join "`n"
$message = [regex]::Unescape($message)

$result = [ordered]@{date=$dateKey;keyword=$selected.keyword;code=$selected.code;name=$selected.name;stock=$selected.stock;recent_30d_natural_sales=$selected.recent_natural;action=$selected.action;inventory_snapshot=$latestInventoryDate.ToString('yyyy-MM-dd');message=$message}
if ($DryRun) {
    [ordered]@{ok=$true;dry_run=$true;suggestion=$result} | ConvertTo-Json -Depth 8
    exit 0
}

$idempotencyKey = "petstar-traffic-$($Date.ToString('yyyyMMdd'))"
$send = Invoke-LarkJson { & $larkCli im +messages-send --as bot --user-id $openId --text $message --idempotency-key $idempotencyKey --format json } 'Send traffic suggestion failed' 2
$keptHistory = @($history.items | Where-Object {
    try { ([datetime]$_.date).Date -ge $Date.Date.AddDays(-90) } catch { $false }
})
$newHistory = @($keptHistory) + @([pscustomobject]$result)
if (-not (Test-Path -LiteralPath $stateRoot)) { $null = New-Item -ItemType Directory -Path $stateRoot }
[IO.File]::WriteAllText($historyPath,([ordered]@{items=$newHistory;updated_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')} | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
[ordered]@{ok=$true;sent=$true;suggestion=$result;message_id=$send.data.message_id} | ConvertTo-Json -Depth 8
