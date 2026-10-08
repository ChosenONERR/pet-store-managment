param()
$ErrorActionPreference='Stop'
$projectRoot=Split-Path $PSScriptRoot -Parent
$cli=Join-Path $projectRoot 'agent\lark_cli_noexit.ps1'
$m=Get-Content (Join-Path $projectRoot 'work\simple_finance_meta.json') -Raw -Encoding UTF8|ConvertFrom-Json
$token=([uri]$m.workbook_url).Segments[-1].Trim('/')
$formulas=@(
    '=''赠品记录''!$A$6=A190',
    '=''赠品记录''!$B$6=B190',
    '=''赠品记录''!$C$6=Q190'
)
$cells=@();foreach($formula in $formulas){$cells+=,@(@{formula=$formula})}
$payload=$cells|ConvertTo-Json -Depth 10 -Compress
$write=$payload|&$cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name '订单汇总' --range 'Z1:Z3' --cells - --format json|Out-String
if(-not($write|ConvertFrom-Json).ok){throw $write}
try{
    Start-Sleep -Seconds 3
    $read=&$cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '订单汇总' --range 'Z1:Z3' --include 'value,formula' --max-chars 10000 --format json|Out-String
    if(-not($read|ConvertFrom-Json).ok){throw $read}
    $read
}finally{
    $ranges=ConvertTo-Json -InputObject @('订单汇总!Z1:Z3') -Compress
    $clear=$ranges|&$cli sheets +cells-batch-clear --as user --spreadsheet-token $token --ranges - --scope content --yes --format json|Out-String
    if(-not($clear|ConvertFrom-Json).ok){throw $clear}
}
