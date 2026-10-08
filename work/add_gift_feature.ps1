param(
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }

$sourceMeta = [System.IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [System.IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\daily_ledger_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$larkCli = Join-Path $env:APPDATA 'npm\lark-cli.cmd'
if (-not (Test-Path -LiteralPath $larkCli)) { throw "找不到飞书 CLI：$larkCli" }

function Invoke-LarkJson([string[]]$Arguments, [string]$InputJson) {
    $previousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $raw = if ($InputJson) { $InputJson | & $larkCli @Arguments 2>$null | Out-String } else { & $larkCli @Arguments 2>$null | Out-String }
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorAction
    }
    if ($exitCode -ne 0) { throw "飞书命令失败 [$($Arguments -join ' ')]：$($raw.Trim())" }
    try { $result = $raw | ConvertFrom-Json } catch { throw "飞书命令返回无法解析：$($raw.Trim())" }
    if (-not $result.ok) { throw "飞书命令未成功：$($result.error.message)" }
    return $result
}

$sourceInfo = Invoke-LarkJson @('sheets','+workbook-info','--as','user','--spreadsheet-token',$sourceToken,'--format','json') $null
$sourceNames = @($sourceInfo.data.sheets | ForEach-Object { $_.sheet_name })
if ('赠品记录' -notin $sourceNames) {
    $null = Invoke-LarkJson @('sheets','+sheet-create','--as','user','--spreadsheet-token',$sourceToken,'--title','赠品记录','--row-count','2001','--col-count','12','--format','json') $null
    $headers = @('日期','平台','当日单号','订单编号','赠品货号','数量','冻结单位成本','赠品成本','成本状态','备注','登记时间','事件ID')
    $headerPayload = '[' + (($headers | ForEach-Object { @{value=$_} } | ConvertTo-Json -Depth 5 -Compress)) + ']'
    $null = Invoke-LarkJson @('sheets','+cells-set','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','A1:L1','--cells','-','--format','json') $headerPayload
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','A1:L1','--background-color','#173A2D','--font-color','#FFFFFF','--font-weight','bold','--horizontal-alignment','center','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','A2:L2001','--background-color','#F7FAF8','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','E2:J2001','--background-color','#FFF8E1','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','A2:A2001','--number-format','yyyy-mm-dd','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','F2:F2001','--number-format','0.##','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','G2:H2001','--number-format','#,##0.00','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','K2:K2001','--number-format','yyyy-mm-dd hh:mm','--format','json') $null
    foreach ($widthSpec in @(@('A:C','120'),@('D','205'),@('E','120'),@('F:I','115'),@('J','300'),@('K','155'),@('L','240'))) {
        $null = Invoke-LarkJson @('sheets','+cols-resize','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range',$widthSpec[0],'--width',$widthSpec[1],'--format','json') $null
    }
}
$null = Invoke-LarkJson @('sheets','+dim-freeze','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--dimension','row','--count','1','--format','json') $null
$null = Invoke-LarkJson @('sheets','+sheet-hide','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--format','json') $null

foreach ($sheetName in @('8月收入','9月收入')) {
    $headerResult = Invoke-LarkJson @('sheets','+cells-get','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','A1:K2','--include','value,formula','--max-chars','20000','--format','json') $null
    $header = @($headerResult.data.ranges[0].cells[0] | ForEach-Object { [string]$_.value })
    if ('赠品（KW）' -notin $header) {
        $null = Invoke-LarkJson @('sheets','+dim-insert','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--position','I','--count','1','--inherit-style','before','--format','json') $null
    }
    $headerPayload = '[[{"value":"赠品（KW）"}]]'
    $null = Invoke-LarkJson @('sheets','+cells-set','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','I1:I1','--cells','-','--format','json') $headerPayload
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','A1:K1','--background-color','#173A2D','--font-color','#FFFFFF','--font-weight','bold','--horizontal-alignment','center','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','I2:I1001','--background-color','#EAF5EE','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','J2:J1001','--background-color','#DCEBE2','--font-weight','bold','--number-format','#,##0.00;[Red]-#,##0.00','--format','json') $null
    $null = Invoke-LarkJson @('sheets','+cells-set-style','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','K2:K1001','--background-color','#FFF8E1','--format','json') $null
    foreach ($widthSpec in @(@('A','120'),@('B','205'),@('C','110'),@('D:G','120'),@('H','300'),@('I','220'),@('J','125'),@('K','260'))) {
        $null = Invoke-LarkJson @('sheets','+cols-resize','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range',$widthSpec[0],'--width',$widthSpec[1],'--format','json') $null
    }
    $filterProperties = '{"rules":[]}'
    try {
        $null = Invoke-LarkJson @('sheets','+filter-update','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','A1:K1001','--properties','-','--format','json') $filterProperties
    } catch {
        $null = Invoke-LarkJson @('sheets','+filter-create','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','A1:K1001','--properties','-','--format','json') $filterProperties
    }
    $null = Invoke-LarkJson @('sheets','+dim-freeze','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--dimension','row','--count','1','--format','json') $null
}

$sourceCheck = Invoke-LarkJson @('sheets','+cells-get','--as','user','--spreadsheet-token',$sourceToken,'--sheet-name','赠品记录','--range','A1:L2','--include','value,formula','--max-chars','20000','--format','json') $null
$ledgerChecks = @()
foreach ($sheetName in @('8月收入','9月收入')) {
    $ledgerChecks += Invoke-LarkJson @('sheets','+cells-get','--as','user','--spreadsheet-token',$ledgerToken,'--sheet-name',$sheetName,'--range','A1:K2','--include','value,formula','--max-chars','20000','--format','json') $null
}

[ordered]@{
    ok = $true
    source_header = @($sourceCheck.data.ranges[0].cells[0] | ForEach-Object { $_.value })
    ledger_headers = @($ledgerChecks | ForEach-Object { @($_.data.ranges[0].cells[0] | ForEach-Object { $_.value }) })
    completed_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
} | ConvertTo-Json -Depth 8
