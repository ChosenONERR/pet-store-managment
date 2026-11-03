$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'agent\keyword_library.ps1')
$config=[IO.File]::ReadAllText((Join-Path $root 'agent\traffic_suggestion_rules.json'),[Text.Encoding]::UTF8)|ConvertFrom-Json
function Make-KeywordRow([string[]]$Values){return ,@($Values|ForEach-Object {[pscustomobject]@{value=$_}})}
function New-TestRows($Entries){
    if(@($Entries).Count -gt 0 -and @($Entries)[0] -is [string]){$Entries=,@($Entries)}
    $r=@();$r+=,(Make-KeywordRow @('启用','搜索关键词','商品品类','关联货号（选填）'))
    foreach($e in $Entries){$r+=,(Make-KeywordRow $e)}
    return ,$r
}
$cases=0
foreach($date in @('11/2/2026','2026-11-02','46328')){if((ConvertTo-KeywordDay $date) -ne '2026-11-02'){throw 'Mixed inventory date formats were not normalized'};$cases++}
$rows=New-TestRows @(@('是','皇家猫粮','猫粮',''),@('否','猫砂','猫砂',''),@('','狗粮','狗粮',''))
$words=@(ConvertFrom-KeywordLibrary $rows)
if($words.Count -ne 1 -or $words[0].keyword -ne '皇家猫粮'){throw 'Enable/disable parsing failed'};$cases++
if(-not (Test-KeywordProductMatch $words[0] 'kw1106' '皇家成猫粮400g' $config)){throw 'Brand did not match'};$cases++
if(Test-KeywordProductMatch $words[0] 'kw1106' '网易严选猫粮' $config){throw 'Wrong brand matched'};$cases++
if(Test-KeywordProductMatch $words[0] 'kw1106' '皇家猫条零食' $config){throw 'Wrong category matched'};$cases++
$new=@(ConvertFrom-KeywordLibrary (New-TestRows @(@('是','鸡胸肉','零食',''))))
if(-not (Test-KeywordProductMatch $new[0] 'kw2001' '猫用鸡胸肉100g' $config)){throw 'Manual new keyword was ignored'};$cases++
$linked=@(ConvertFrom-KeywordLibrary (New-TestRows @(@('是','猫粮','猫粮','KW1001，kw1002'))))
if(-not (Test-KeywordProductMatch $linked[0] 'kw1002' '测试商品' $config)){throw 'Explicit KW did not match'};$cases++
if(Test-KeywordProductMatch $linked[0] 'kw1003' '全价猫粮' $config){throw 'Explicit KW silently fell back to unrelated product'};$cases++
$dupes=@(ConvertFrom-KeywordLibrary (New-TestRows @(@('是','猫粮','猫粮',''),@('是',' 猫粮 ','猫粮',''))))
if($dupes.Count -ne 1){throw 'Identical duplicate keyword not deduplicated'};$cases++
foreach($bad in @(@(@('是','猫粮','猫粮',''),@('是','猫粮','狗粮','')),@(@('开','猫粮','猫粮','')),@(@('是','','猫粮','')),@(@('是','猫粮','猫粮','123')))){
    $thrown=$false;try{$null=ConvertFrom-KeywordLibrary (New-TestRows $bad)}catch{$thrown=$true}
    if(-not $thrown){throw 'Invalid keyword configuration was accepted'};$cases++
}
$entries=@(ConvertFrom-KeywordLibrary (New-TestRows @(@('是','猫粮','猫粮',''),@('是','猫砂','猫砂',''),@('是','狗粮','狗粮',''))))
$history=@([pscustomobject]@{date='2026-11-02';keyword='猫粮'},[pscustomobject]@{date='2026-11-01';keyword='猫砂'})
$order=@(Get-RotatedKeywords $entries $history ([datetime]'2026-11-03'))
if(($order.keyword -join ',') -ne '狗粮,猫砂,猫粮'){throw 'Keyword recency rotation failed'};$cases++
$empty=@(ConvertFrom-KeywordLibrary (New-TestRows @(@('否','猫粮','猫粮',''))))
if($empty.Count -ne 0){throw 'Disabled library must stay empty, no fixed fallback'};$cases++
$seeds=@(Get-KeywordSeeds)
if($seeds.Count -ne 13 -or $seeds[0].Count -ne 2){throw 'Seed library shape incorrect'};$cases++
$runtime=[IO.File]::ReadAllText((Join-Path $root 'agent\send_daily_traffic_suggestion.ps1'),[Text.Encoding]::UTF8)
foreach($pattern in @('config.keyword_sheet','ConvertFrom-KeywordLibrary','Get-RotatedKeywords','Test-KeywordProductMatch','-not \$DryRun','\$stock -le 0','keyword_source','未回退固定关键词')){
    if($runtime -notmatch $pattern){throw "Missing runtime keyword contract $pattern"};$cases++
}
$setup=[IO.File]::ReadAllText((Join-Path $root 'work\ensure_keyword_library.ps1'),[Text.Encoding]::UTF8)
if($setup -notmatch 'Existing keyword library preserved' -or $setup -notmatch '--allow-overwrite=false' -or $setup -match 'sheet-delete'){throw 'Setup would overwrite/delete user edits'};$cases++
[ordered]@{ok=$true;cases=$cases;external_writes=0}|ConvertTo-Json -Compress
