param([string]$ProjectRoot,[switch]$Apply)
$ErrorActionPreference='Stop'
$OutputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
if(-not $ProjectRoot){$ProjectRoot=Split-Path $PSScriptRoot -Parent}
$agentRoot=Join-Path $ProjectRoot 'agent'
. (Join-Path $agentRoot 'keyword_library.ps1')
. (Join-Path $agentRoot 'finance_transaction.ps1')
$cli=Join-Path $agentRoot 'lark_cli_noexit.ps1'
$meta=[IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8)|ConvertFrom-Json
$config=[IO.File]::ReadAllText((Join-Path $agentRoot 'traffic_suggestion_rules.json'),[Text.Encoding]::UTF8)|ConvertFrom-Json
$token=([uri]$meta.workbook_url).Segments[-1].Trim('/')
$sheet=[string]$config.keyword_sheet
function Invoke-KeywordLark([scriptblock]$Command){
    $global:LASTEXITCODE=0;$raw=& $Command|Out-String
    if($LASTEXITCODE -ne 0){throw '关键词库操作失败，请检查飞书授权或命令输出'}
    $r=$raw|ConvertFrom-Json
    if($r.ok -ne $true){throw '关键词库操作未成功'}
    return $r
}
function Read-KeywordRows([string]$Name,[string]$Range){
    $r=Invoke-KeywordLark { & $cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name $Name --range $Range --include value --max-chars 1000000 --format json }
    if($r.data.has_more){throw '关键词库初始化读取不完整'}
    return ,@($r.data.ranges[0].cells)
}
$transaction=$null;$closed=$false
try{
    if($Apply){$transaction=New-FinanceTransaction (Join-Path $agentRoot 'state') '初始化搜索关键词库' $null -AcquireLock -WorkbookTokens @{source=$token} -LarkCli $cli}
    $book=Invoke-KeywordLark { & $cli sheets +workbook-info --as user --spreadsheet-token $token --format json }
    $found=@($book.data.sheets|Where-Object {$_.sheet_name -eq $sheet})
    if($found.Count -gt 1){throw '关键词库工作表不唯一'}
    $existing=@()
    if($found.Count){$existing=Read-KeywordRows $sheet 'A1:D1001'}
    $hasContent=@($existing|ForEach-Object {for($i=0;$i -lt 4;$i++){Get-KeywordCell $_ $i}}|Where-Object {$_}).Count -gt 0
    if($hasContent){
        $entries=@(ConvertFrom-KeywordLibrary $existing)
        if($Apply){Complete-FinanceTransaction $transaction;Close-FinanceTransaction $transaction;$closed=$true}
        [ordered]@{ok=$true;skipped=$true;reason='Existing keyword library preserved';enabled=$entries.Count;sheet_name=$sheet;workbook_url=$meta.workbook_url}|ConvertTo-Json -Depth 4
        return
    }
    $inventory=Read-KeywordRows ([string]$config.inventory_sheet) 'A2:O2000'
    $latest=@($inventory|ForEach-Object {ConvertTo-KeywordDay (Get-KeywordCell $_ 0)}|Where-Object {$_}|Sort-Object -Descending|Select-Object -First 1)
    if(-not $latest.Count){throw '缺少库存快照，不能核实品牌词'}
    $matrix=@();$matrix+=,@('启用','搜索关键词','商品品类','关联货号（选填）')
    $index=0
    foreach($seed in @(Get-KeywordSeeds)){
        $enabled=$index -lt 5
        if(-not $enabled){
            $entry=[pscustomobject]@{keyword=$seed[0];category=$seed[1];codes=@()}
            foreach($r in $inventory){
                if((ConvertTo-KeywordDay (Get-KeywordCell $r 0)) -ne $latest[0] -or (Get-KeywordCell $r 5) -ne [string]$config.selling_status){continue}
                $name=Get-KeywordCell $r 2;$excluded=$false
                foreach($pattern in @($config.global_exclude)){if($name -like "*$pattern*"){$excluded=$true;break}}
                if($excluded){continue}
                $stock=0.0
                if(-not [double]::TryParse((Get-KeywordCell $r 3),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$stock) -or $stock -le 0){continue}
                if(Test-KeywordProductMatch $entry (Get-KeywordCell $r 1) (Get-KeywordCell $r 2) $config){$enabled=$true;break}
            }
        }
        $matrix+=,@($(if($enabled){'是'}else{'否'}),$seed[0],$seed[1],'');$index++
    }
    if(-not $Apply){[ordered]@{ok=$true;dry_run=$true;sheet_name=$sheet;rows=$matrix;workbook_url=$meta.workbook_url}|ConvertTo-Json -Depth 5;return}
    if(-not $found.Count){$null=Invoke-KeywordLark { & $cli sheets +sheet-create --as user --spreadsheet-token $token --title $sheet --row-count 1001 --col-count 4 --format json }}
    $range='A1:D'+$matrix.Count
    Register-FinanceRangeBeforeImage $transaction 'source' $token $sheet $range $cli
    $cells=@();foreach($r in $matrix){$cells+=,@($r|ForEach-Object {@{value=[string]$_}})}
    $payload=ConvertTo-Json -InputObject $cells -Depth 6 -Compress
    $null=Invoke-KeywordLark { $payload | & $cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $sheet --range $range --cells - --allow-overwrite=false --format json }
    $null=Invoke-KeywordLark { & $cli sheets +cells-set-style --as user --spreadsheet-token $token --sheet-name $sheet --range 'A1:D1001' --number-format '@' --font-family 'Microsoft YaHei' --font-size 11 --vertical-alignment middle --word-wrap auto-wrap --format json }
    $null=Invoke-KeywordLark { & $cli sheets +cells-set-style --as user --spreadsheet-token $token --sheet-name $sheet --range 'A1:D1' --background-color '#173A2D' --font-color '#FFFFFF' --font-weight bold --format json }
    $options=ConvertTo-Json -InputObject @('是','否') -Compress
    $null=Invoke-KeywordLark { $options | & $cli sheets +dropdown-set --as user --spreadsheet-token $token --sheet-name $sheet --range 'A2:A1001' --options - --highlight=false --format json }
    $widths=ConvertTo-Json -InputObject @{A=80;B=190;C=140;D=240} -Compress
    $null=Invoke-KeywordLark { $widths | & $cli sheets +cols-resize --as user --spreadsheet-token $token --sheet-name $sheet --widths - --format json }
    $null=Invoke-KeywordLark { & $cli sheets +dim-freeze --as user --spreadsheet-token $token --sheet-name $sheet --rows 1 --format json }
    $readback=Read-KeywordRows $sheet $range
    for($i=0;$i -lt $matrix.Count;$i++){for($j=0;$j -lt 4;$j++){if((Get-KeywordCell $readback[$i] $j) -ne [string]$matrix[$i][$j]){throw '关键词库回读与初始化内容不一致'}}}
    $verify=Invoke-KeywordLark { & $cli sheets +formula-verify --as user --spreadsheet-token $token --sheet-name $sheet --range $range --exit-on-error --format json }
    if($verify.data.status -ne 'success'){throw '关键词库单元格校验失败'}
    Complete-FinanceTransaction $transaction;Close-FinanceTransaction $transaction;$closed=$true
    [ordered]@{ok=$true;applied=$true;sheet_name=$sheet;keyword_count=($matrix.Count-1);enabled=@(ConvertFrom-KeywordLibrary $readback).Count;workbook_url=$meta.workbook_url}|ConvertTo-Json -Depth 4
} catch {
    if($null -ne $transaction -and -not $closed){$null=Invoke-FinanceTransactionRollback $transaction @{source=$token} $cli $_.Exception.Message}
    throw
} finally {if($null -ne $transaction -and -not $closed){Close-FinanceTransaction $transaction}}
