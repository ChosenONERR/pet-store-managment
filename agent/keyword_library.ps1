# Pure helpers. No network, file writes or message sends.
function Get-KeywordCell($Row,[int]$Index) {
    if($null -eq $Row -or $Index -ge @($Row).Count){return ''}
    $cell=@($Row)[$Index]
    if($null -eq $cell){return ''}
    $p=$cell.PSObject.Properties['value']
    if($null -eq $p){return ''}
    return ([string]$p.Value).Trim()
}
function ConvertTo-KeywordText([string]$Text) {
    return [regex]::Replace($Text.ToLowerInvariant(),'[\s\p{P}\p{S}]+','')
}
function ConvertTo-KeywordDay([string]$Value) {
    if(-not $Value){return ''}
    $number=0.0
    if([double]::TryParse($Value,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)){
        return ([datetime]::FromOADate($number)).ToString('yyyy-MM-dd')
    }
    return ([datetime]::Parse($Value,[Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
}
function ConvertFrom-KeywordLibrary($Rows) {
    $headers=@('启用','搜索关键词','商品品类','关联货号（选填）')
    if(@($Rows).Count -eq 0){throw '关键词库为空，请先初始化'}
    for($i=0;$i -lt 4;$i++){if((Get-KeywordCell $Rows[0] $i) -ne $headers[$i]){throw "关键词库第$($i+1)列表头应为：$($headers[$i])"}}
    $result=[Collections.Generic.List[object]]::new();$seen=@{}
    for($i=1;$i -lt @($Rows).Count;$i++){
        $r=$Rows[$i];$on=Get-KeywordCell $r 0;$word=Get-KeywordCell $r 1
        if(-not $on -or $on -in @('否','false','FALSE','0')){continue}
        if($on -notin @('是','true','TRUE','1')){throw ('关键词库第{0}行启用值请填写“是”或“否”' -f ($i+1))}
        if(-not $word){throw "关键词库第$($i+1)行缺少搜索关键词"}
        $codes=@([regex]::Split((Get-KeywordCell $r 3),'[,，、;；\s]+')|Where-Object {$_}|ForEach-Object {$_.ToLowerInvariant()}|Sort-Object -Unique)
        foreach($c in $codes){if($c -notmatch '^kw\d+$'){throw "关键词库第$($i+1)行关联货号格式不正确：$c"}}
        $key=ConvertTo-KeywordText $word
        if($seen.ContainsKey($key)){
            $old=$seen[$key]
            if($old.category -ne (Get-KeywordCell $r 2) -or ($old.codes -join ',') -ne ($codes -join ',')){throw "关键词库同词配置冲突：$word"}
            continue
        }
        $entry=[pscustomobject]@{keyword=$word;category=(Get-KeywordCell $r 2);codes=$codes;index=$result.Count}
        $result.Add($entry);$seen[$key]=$entry
    }
    return $result.ToArray()
}
function Get-KeywordCategoryRule($Entry,$Config) {
    $category=$Entry.category
    if($category -in @('猫主粮','猫食品')){$category='猫粮'}
    if($category -in @('狗主粮','犬粮')){$category='狗粮'}
    if($category -in @('保健用品','益生菌')){$category='宠物益生菌'}
    if($category -in @('清洁用品','毛巾')){$category='宠物清洁'}
    return @($Config.rules|Where-Object {$_.keyword -eq $category}|Select-Object -First 1)
}
function Test-KeywordProductMatch($Entry,[string]$Code,[string]$Name,$Config) {
    if(@($Entry.codes).Count -gt 0){return $Code.ToLowerInvariant() -in @($Entry.codes)}
    $rule=@(Get-KeywordCategoryRule $Entry $Config)
    $nameText=ConvertTo-KeywordText $Name
    $wordText=ConvertTo-KeywordText $Entry.keyword
    if($rule.Count -eq 0){return $nameText.Contains($wordText)}
    $r=$rule[0];$categoryMatched=$false
    foreach($part in @($r.include)){if($Name -like "*$part*"){$categoryMatched=$true;break}}
    foreach($part in @($r.exclude)){if($Name -like "*$part*"){return $false}}
    if(-not $categoryMatched){return $false}
    # Keep brand terms: Royal cat food must never match an unrelated cat food.
    $specific=$wordText.Replace((ConvertTo-KeywordText $r.keyword),'')
    if($specific.StartsWith('宠物')){$specific=$specific.Substring(2)}
    if(-not $specific){return $true}
    return $nameText.Contains($specific)
}
function Get-RotatedKeywords($Entries,$History,[datetime]$Date) {
    return @($Entries|Sort-Object @{Expression={
        $key=ConvertTo-KeywordText $_.keyword
        $last=@($History|Where-Object {([string]$_.date -lt $Date.ToString('yyyy-MM-dd')) -and (ConvertTo-KeywordText ([string]$_.keyword)) -eq $key}|Sort-Object date -Descending|Select-Object -First 1)
        if($last.Count){[string]$last[0].date}else{'0000-00-00'}
    }},index)
}
function Get-KeywordSeeds {
    return @(
        @('猫粮','猫粮'),@('猫砂','猫砂'),@('狗粮','狗粮'),@('宠物益生菌','宠物益生菌'),@('宠物毛巾','宠物清洁'),
        @('网易猫粮','猫粮'),@('纯福猫粮','猫粮'),@('皇家猫粮','猫粮'),@('诚实一口猫粮','猫粮'),@('顽皮猫粮','猫粮'),@('艾尔猫粮','猫粮'),
        @('小猫班长猫砂','猫砂'),@('城市理享猫砂','猫砂')
    )
}
