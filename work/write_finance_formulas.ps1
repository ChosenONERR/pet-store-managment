$ErrorActionPreference = 'Stop'
$token = 'JV1aswYrChprbbthehHcJ5O0nCf'

function Set-FormulaCell([string]$Sheet, [string]$Range, [string]$Formula, [string]$CopyTo = '') {
    $cellJson = @{formula=$Formula; cell_styles=@{number_format='#,##0.00'}} | ConvertTo-Json -Depth 8 -Compress
    $payload = "[[$cellJson]]"
    $args = @('sheets','+cells-set','--as','user','--spreadsheet-token',$token,'--sheet-name',$Sheet,'--range',$Range,'--cells',$payload,'--format','json')
    if ($CopyTo) { $args += @('--copy-to-range',$CopyTo) }
    & lark-cli @args
    if ($LASTEXITCODE -ne 0) { throw "Formula write failed: $Sheet $Range" }
}

# Daily derived columns.
Set-FormulaCell '每日记录' 'F2' '=SUM(B2:E2)'
Set-FormulaCell '每日记录' 'K2' '=SUM(G2:J2)'
$monthCellJson = @{formula='=TEXT(A2,"yyyy-mm")'; cell_styles=@{number_format='@';horizontal_alignment='center'}} | ConvertTo-Json -Depth 8 -Compress
$monthPayload = "[[$monthCellJson]]"
& lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name '每日记录' --range 'Q2' --cells $monthPayload --format json
if ($LASTEXITCODE -ne 0) { throw 'Daily month formula write failed' }

# Month opening stock rolls from prior closing stock after the first month.
Set-FormulaCell '月度补充' 'B3' '=C2'

# Asset amortization/depreciation.
$assetRow = @(
    @{formula='=IF(OR(D2="",E2=""),"",EDATE(D2,E2-1))';cell_styles=@{number_format='yyyy-mm'}},
    @{formula='=IFERROR(C2/E2,0)';cell_styles=@{number_format='#,##0.00'}}
)
$assetRowJson = ConvertTo-Json -InputObject $assetRow -Depth 8 -Compress
$assetPayload = "[$assetRowJson]"
& lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name '资产与口径' --range 'F2:G2' --cells $assetPayload --format json
if ($LASTEXITCODE -ne 0) { throw 'Asset formula write failed' }

# Monthly summary formulas B:X (A is the month template).
$formulas = @(
    '=SUMIF(''每日记录''!$Q:$Q,TEXT($A2,"yyyy-mm"),''每日记录''!$F:$F)',
    '=SUMIF(''每日记录''!$Q:$Q,TEXT($A2,"yyyy-mm"),''每日记录''!$K:$K)',
    '=SUMIF(''每日记录''!$Q:$Q,TEXT($A2,"yyyy-mm"),''每日记录''!$L:$L)',
    '=SUMIF(''每日记录''!$Q:$Q,TEXT($A2,"yyyy-mm"),''每日记录''!$M:$M)',
    '=IFERROR(VLOOKUP($A2,''月度补充''!$A:$H,2,FALSE),0)',
    '=IFERROR(VLOOKUP($A2,''月度补充''!$A:$H,3,FALSE),0)',
    '=G2-F2',
    '=F2+D2-G2',
    '=B2-I2',
    '=SUMIF(''每日记录''!$Q:$Q,TEXT($A2,"yyyy-mm"),''每日记录''!$N:$N)',
    '=IFERROR(VLOOKUP($A2,''月度补充''!$A:$H,4,FALSE),0)',
    '=IFERROR(VLOOKUP($A2,''月度补充''!$A:$H,5,FALSE),0)',
    '=SUMIFS(''资产与口径''!$G:$G,''资产与口径''!$B:$B,"品牌服务费摊销",''资产与口径''!$D:$D,"<="&$A2,''资产与口径''!$F:$F,">="&$A2)',
    '=SUMIFS(''资产与口径''!$G:$G,''资产与口径''!$B:$B,"装修/设备折旧",''资产与口径''!$D:$D,"<="&$A2,''资产与口径''!$F:$F,">="&$A2)',
    '=SUM(L2:O2)',
    '=IFERROR(VLOOKUP($A2,''月度补充''!$A:$H,6,FALSE),0)',
    '=J2-P2-Q2',
    '=IFERROR(VLOOKUP($A2,''月度补充''!$A:$H,7,FALSE),0)',
    '=R2-S2',
    '=SUMIF(''每日记录''!$Q:$Q,TEXT($A2,"yyyy-mm"),''每日记录''!$O:$O)',
    '=C2-E2-K2-U2',
    '=IF(ROW()=2,IFERROR(VLOOKUP($A2,''月度补充''!$A:$H,8,FALSE),0),X1)',
    '=W2+V2'
)
$cells = @()
foreach ($formula in $formulas) { $cells += @{formula=$formula;cell_styles=@{number_format='#,##0.00'}} }
$summaryRowJson = ConvertTo-Json -InputObject $cells -Depth 8 -Compress
$summaryPayload = "[$summaryRowJson]"
& lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name '月度汇总' --range 'B2:X2' --cells $summaryPayload --format json
if ($LASTEXITCODE -ne 0) { throw 'Summary formula write failed' }
