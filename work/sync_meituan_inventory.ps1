param(
    [string]$InventoryPath,
    [string]$WorkRoot,
    [datetime]$SnapshotDate = [datetime]::Today
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if (-not $WorkRoot) { $WorkRoot = $PSScriptRoot }
if (-not [IO.Path]::IsPathRooted($WorkRoot)) { $WorkRoot = (Resolve-Path -LiteralPath $WorkRoot).Path }
$metaPath = Join-Path $WorkRoot 'simple_finance_meta.json'
$exporterPath = Join-Path $WorkRoot 'meituan_inventory_export.mjs'
$projectRoot = Split-Path $WorkRoot -Parent
$skuDiscoveryPath = Join-Path $projectRoot 'agent\sync_sku_master_by_identity.ps1'
$inventoryDirectory = 'D:\market-research\宠速星\商品库存情况'
$snapshotDate = $SnapshotDate.Date
if (-not (Test-Path -LiteralPath $metaPath)) { throw '缺少简版工作簿元数据' }
$meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
$spreadsheetToken = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
. (Join-Path $projectRoot 'agent\finance_transaction.ps1')
$stateRoot = Join-Path $projectRoot 'agent\state'
$script:financeTransaction = Join-FinanceTransaction $stateRoot
$script:financeTransactionOwner = $false
$script:financeTransactionClosed = $false
if ($null -eq $script:financeTransaction) {
    $script:financeTransaction = New-FinanceTransaction $stateRoot "同步美团库存 $($snapshotDate.ToString('yyyy-MM-dd'))" $null -AcquireLock -WorkbookTokens @{source=$spreadsheetToken} -LarkCli 'lark-cli'
    $script:financeTransactionOwner = $true
}
function Close-InventorySyncTransaction([bool]$Succeeded, [string]$Reason = '') {
    if ($script:financeTransactionClosed) { return }
    if ($script:financeTransactionOwner) {
        if ($Succeeded) { Complete-FinanceTransaction $script:financeTransaction }
        else { $null = Invoke-FinanceTransactionRollback $script:financeTransaction @{source=$spreadsheetToken} 'lark-cli' $Reason }
    }
    Close-FinanceTransaction $script:financeTransaction
    $script:financeTransactionClosed = $true
}
trap {
    $reason = [string]$_.Exception.Message
    Close-InventorySyncTransaction $false $reason
    # Windows PowerShell 5.1 may not pipe Console.Error from a nested .ps1
    # back to its caller. Emit structured stdout so the parent can preserve
    # the real stage and reason in the Feishu failure card.
    [Console]::Out.WriteLine(([ordered]@{ok=$false;error=$reason} | ConvertTo-Json -Compress))
    exit 1
}
$downloadSource = if ($InventoryPath) { 'provided_file' } else { 'new_export' }
$degraded = $false
$exportError = $null
$exportInfo = $null

function Get-ZipEntryText([System.IO.Compression.ZipArchive]$Zip, [string]$Name) {
    $entry = $Zip.GetEntry($Name)
    if (-not $entry) { throw "XLSX缺少文件：$Name" }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Convert-ColumnNameToIndex([string]$Name) {
    $n = 0
    foreach ($ch in $Name.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1) }
    return $n - 1
}
function Get-CellValue($Cell, [string[]]$SharedStrings) {
    $type = [string]$Cell.t; $raw = [string]$Cell.v
    if ($type -eq 's' -and $raw -match '^\d+$') { return $SharedStrings[[int]$raw] }
    if ($type -eq 'inlineStr') { return ([string]$Cell.is.t) }
    if ($null -eq $raw -or $raw -eq '') { return $null }
    if ($raw -match '^-?\d+(\.\d+)?$') { return [double]$raw }
    return $raw
}
function Read-XlsxRows([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $wb = [xml](Get-ZipEntryText $zip 'xl/workbook.xml')
        $rels = [xml](Get-ZipEntryText $zip 'xl/_rels/workbook.xml.rels')
        $shared = @()
        if ($zip.GetEntry('xl/sharedStrings.xml')) {
            $ss = [xml](Get-ZipEntryText $zip 'xl/sharedStrings.xml')
            $shared = @($ss.SelectNodes("//*[local-name()='si']") | ForEach-Object { (($_.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.'#text' }) -join '') })
        }
        $sheetNode = @($wb.SelectNodes("//*[local-name()='sheet']") | Select-Object -First 1)
        $rid = [string]$sheetNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        if (-not $rid) { $rid = [string]$sheetNode.'r:id' }
        $rel = @($rels.SelectNodes("//*[local-name()='Relationship']") | Where-Object { ([string]$_.Id) -eq $rid } | Select-Object -First 1)
        $target = ([string]$rel.Target).TrimStart('/')
        if ($target -notlike 'xl/*') { $target = "xl/$target" }
        $sheet = [xml](Get-ZipEntryText $zip $target)
        $rows = @()
        foreach ($row in $sheet.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $cells = @{}
            foreach ($cell in $row.SelectNodes("./*[local-name()='c']")) {
                $col = [regex]::Match([string]$cell.r, '^[A-Z]+').Value
                if ($col) { $cells[(Convert-ColumnNameToIndex $col)] = Get-CellValue $cell $shared }
            }
            $max = if ($cells.Count) { ($cells.Keys | Measure-Object -Maximum).Maximum } else { -1 }
            if ($max -ge 0) {
                $arr = [object[]]::new($max + 1)
                foreach ($k in $cells.Keys) { $arr[$k] = $cells[$k] }
                $rows += ,$arr
            }
        }
        return ,$rows
    } finally { $zip.Dispose() }
}
function Get-At([object[]]$Row, [int]$Index) { if ($null -ne $Row -and $Index -lt $Row.Count) { return $Row[$Index] }; return $null }
function Normalize($Value) { if ($null -eq $Value -or "$Value" -in @('-','—')) { return $null }; return ([string]$Value).Trim() }
function To-Number($Value) { $n=0.0; if ($null -ne $Value -and [double]::TryParse(([string]$Value).Replace(',',''),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){return $n}; return $null }
function Convert-ToExcelSerial([datetime]$Date) { return ($Date - [datetime]'1899-12-30').TotalDays }
function New-ValueCell($Value) { if ($null -eq $Value -or "$Value" -eq '') { return @{value=''} }; return @{value=$Value} }
function Invoke-LarkJson([scriptblock]$Command, [string]$Label, [int]$MaxAttempts = 3) {
    $errors = [Collections.Generic.List[string]]::new()
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $global:LASTEXITCODE = 0
        try {
            $raw = & $Command | Out-String
            $exitCode = $LASTEXITCODE
            $result = $raw | ConvertFrom-Json
            if ($exitCode -eq 0 -and $result.ok -ne $false) { return $result }
            $errors.Add("第${attempt}次：退出码=$exitCode")
        } catch {
            $errors.Add("第${attempt}次：$($_.Exception.Message)")
        }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds ([math]::Min(8, $attempt * 2)) }
    }
    throw "$Label；重试${MaxAttempts}次仍失败：$($errors -join '；')"
}
function Set-CellMatrix([string]$SheetName,[string]$Range,$Rows) {
    Register-FinanceRangeBeforeImage $script:financeTransaction 'source' $spreadsheetToken $SheetName $Range 'lark-cli'
    $payload = ConvertTo-Json -InputObject $Rows -Depth 10 -Compress
    $null = Invoke-LarkJson { $payload | lark-cli sheets +cells-set --as user --spreadsheet-token $spreadsheetToken --sheet-name $SheetName --range $Range --cells - --format json } "写入库存快照失败：$SheetName $Range"
}

function Test-XlsxIntegrity([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $file = Get-Item -LiteralPath $Path
    if ($file.Length -lt 1024) { return $false }
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            return $null -ne $zip.GetEntry('[Content_Types].xml') -and $null -ne $zip.GetEntry('xl/workbook.xml')
        } finally { $zip.Dispose() }
    } catch { return $false }
}

function Get-ExporterFailureDetail($Lines) {
    $candidate = @($Lines | ForEach-Object { [string]$_ } | Where-Object { $_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}') } | Select-Object -Last 1)
    if ($candidate) {
        try {
            $parsed = ([string]$candidate[-1]).Trim() | ConvertFrom-Json
            if ($parsed.message) {
                $stages = @($parsed.failures | ForEach-Object { "第$($_.attempt)次/$($_.stage)：$($_.message)" })
                return if ($stages.Count) { $stages -join '；' } else { [string]$parsed.message }
            }
        } catch {}
    }
    $text = (@($Lines | ForEach-Object { [string]$_ }) -join ' ').Trim()
    if ($text.Length -gt 1200) { return $text.Substring(0, 1200) }
    return $text
}

if (-not $InventoryPath) {
    if (-not (Test-Path -LiteralPath $inventoryDirectory)) {
        $null = New-Item -ItemType Directory -Path $inventoryDirectory
    }
    $InventoryPath = Join-Path $inventoryDirectory ($snapshotDate.ToString('yyyy-MM-dd') + ' 美团导出商品库存情况.xlsx')
    $partialPath = "$InventoryPath.partial"
    $exportMutex = [Threading.Mutex]::new($false, 'Local\PetstarMeituanInventoryExport')
    $lockAcquired = $false
    try {
        try { $lockAcquired = $exportMutex.WaitOne([TimeSpan]::FromSeconds(5)) } catch [Threading.AbandonedMutexException] { $lockAcquired = $true }
        if (-not $lockAcquired) {
            $waitStarted = [datetime]::Now
            while (([datetime]::Now - $waitStarted).TotalSeconds -lt 180) {
                if ((Test-XlsxIntegrity $InventoryPath) -and (Get-Item -LiteralPath $InventoryPath).LastWriteTime -ge $waitStarted.AddSeconds(-10)) {
                    $downloadSource = 'concurrent_export_reused'
                    break
                }
                Start-Sleep -Seconds 3
            }
            if ($downloadSource -ne 'concurrent_export_reused') { throw '另一个库存导出任务占用超过3分钟，且没有产生可复用文件' }
        } else {
            $global:LASTEXITCODE = 0
            # PowerShell 5.1 promotes native stderr to ErrorRecord when the
            # script-wide preference is Stop. Capture Node's JSON error first,
            # then decide whether a recent validated export can be reused.
            $previousErrorActionPreference = $ErrorActionPreference
            try {
                $ErrorActionPreference = 'Continue'
                $exportLines = @(& node $exporterPath 2>&1 | ForEach-Object { [string]$_ })
                $exportExitCode = $LASTEXITCODE
            } finally {
                $ErrorActionPreference = $previousErrorActionPreference
            }
            if ($exportExitCode -ne 0) {
                $exportError = Get-ExporterFailureDetail $exportLines
                $fileAgeHours = if (Test-Path -LiteralPath $InventoryPath) { ([datetime]::Now - (Get-Item -LiteralPath $InventoryPath).LastWriteTime).TotalHours } else { [double]::PositiveInfinity }
                $canReuse = [math]::Abs(([datetime]::Today - $snapshotDate).TotalDays) -le 1 -and $fileAgeHours -le 24 -and (Test-XlsxIntegrity $InventoryPath)
                if ($canReuse) {
                    $downloadSource = 'existing_file_fallback'
                    $degraded = $true
                } else {
                    throw "美团库存导出任务失败：$exportError"
                }
            } else {
                $exportJson = (($exportLines -join "`n").Trim())
                try { $exportInfo = $exportJson | ConvertFrom-Json } catch { throw "库存导出结果不是有效JSON：$($_.Exception.Message)" }
                if (-not $exportInfo.downloadUrl) { throw '未取得美团库存下载地址' }
                $global:LASTEXITCODE = 0
                $curlOutput = @(& curl.exe --fail --location --silent --show-error --connect-timeout 30 --max-time 180 --retry 3 --retry-delay 2 --retry-all-errors ([string]$exportInfo.downloadUrl) --output $partialPath 2>&1 | ForEach-Object { [string]$_ })
                if ($LASTEXITCODE -ne 0) { throw "库存文件下载失败：$($curlOutput -join ' ')" }
                if (-not (Test-XlsxIntegrity $partialPath)) { throw '库存下载结果不是完整的XLSX文件，未覆盖现有库存文件' }
                Move-Item -LiteralPath $partialPath -Destination $InventoryPath -Force
            }
        }
    } finally {
        if ($lockAcquired) { try { $exportMutex.ReleaseMutex() } catch {} }
        $exportMutex.Dispose()
    }
}
if (-not (Test-XlsxIntegrity $InventoryPath)) { throw "库存文件不存在或不是完整XLSX：$InventoryPath" }

$thresholdResult = Invoke-LarkJson { lark-cli sheets +cells-get --as user --spreadsheet-token $spreadsheetToken --sheet-name '库存快照' --range B2:J2000 --include value --max-chars 500000 --format json } '读取库存补货阈值失败'
$thresholds = @{}
foreach ($row in @($thresholdResult.data.ranges[0].cells)) {
    if ($row.Count -ge 9 -and $row[0].value) { $thresholds[[string]$row[0].value] = To-Number $row[8].value }
}

$rows = Read-XlsxRows $InventoryPath
$header = $rows[0]
function Find-HeaderIndex([object[]]$HeaderRow, [string]$Pattern, [string]$Label) {
    for ($i = 0; $i -lt $HeaderRow.Count; $i++) {
        if ([string](Get-At $HeaderRow $i) -match $Pattern) { return $i }
    }
    throw "库存导出缺少必要字段：$Label"
}
$idxName = Find-HeaderIndex $header '^商品名称$' '商品名称'
$idxSkuId = Find-HeaderIndex $header '^sku_id$' 'sku_id'
$idxCode = Find-HeaderIndex $header '货号' '店内码/货号'
$idxSpec = Find-HeaderIndex $header '^规格名称$' '规格名称'
$idxStock = Find-HeaderIndex $header '^库存$' '库存'
$idxStatus = Find-HeaderIndex $header '上/下架|售卖状态' '售卖状态'
$idxSales = Find-HeaderIndex $header '销量|月售' '销量/月售'
$idxShelf = Find-HeaderIndex $header '货架码' '货架码/位置码'
$groups = @{}; $uncoded = @()
foreach ($row in ($rows | Select-Object -Skip 2)) {
    $skuId=Normalize (Get-At $row $idxSkuId); $code=Normalize (Get-At $row $idxCode); $name=Normalize (Get-At $row $idxName); $spec=Normalize (Get-At $row $idxSpec); $stock=To-Number (Get-At $row $idxStock); $sales=To-Number (Get-At $row $idxSales); $status=Normalize (Get-At $row $idxStatus); $shelf=Normalize (Get-At $row $idxShelf)
    if (-not $name) { continue }
    if (-not $skuId) { $skuId = "missing:$code|$spec|$name" }
    if (-not $code) { $uncoded += [pscustomobject]@{SkuId=$skuId;Code=$null;Name=$name;Spec=$spec;Stock=$stock;Sales=$sales;Status=$status;Shelf=$shelf;Count=1}; continue }
    if (-not $groups.ContainsKey($skuId)) { $groups[$skuId]=[ordered]@{SkuId=$skuId;Code=$code;Name=$name;Spec=$spec;Stock=$stock;Sales=$sales;Status=$status;Shelf=$shelf;Count=1} }
    else {
        $g = $groups[$skuId]
        $g.Count++
        if ($null -ne $stock) {
            $existingStock = if ($null -eq $g.Stock) { 0.0 } else { [double]$g.Stock }
            $g.Stock = $existingStock + $stock
        }
        if ($null -ne $sales) {
            $existingSales = if ($null -eq $g.Sales) { 0.0 } else { [double]$g.Sales }
            $g.Sales = $existingSales + $sales
        }
        if ($status -eq '售卖中') { $g.Status = '售卖中' }
        if (-not $g.Shelf -and $shelf) { $g.Shelf = $shelf }
        if (-not $g.Code -and $code) { $g.Code = $code }
        if (-not $g.Spec -and $spec) { $g.Spec = $spec }
    }
}
$records = @($groups.Values | Sort-Object Code,SkuId) + @($uncoded)
if ($records.Count -gt 1999) { throw '库存商品超过1999行容量' }
$AG=@(); $J=@(); $L=@(); $NO=@(); $serial=Convert-ToExcelSerial $snapshotDate
foreach($g in $records){
    $source=if(-not$g.Code){'美团商品库存导出；缺少货号'}elseif($g.Count-gt1){"美团商品库存导出；同一美团SKU出现$($g.Count)行，库存和月售均已合计"}else{'美团商品库存导出'}
    $threshold=if($g.Code-and$thresholds.ContainsKey([string]$g.Code)-and$null-ne$thresholds[[string]$g.Code]){$thresholds[[string]$g.Code]}else{2}
    $AG += ,@((New-ValueCell $serial),(New-ValueCell $g.Code),(New-ValueCell $g.Name),(New-ValueCell $g.Stock),(New-ValueCell $g.Sales),(New-ValueCell $g.Status),(New-ValueCell $g.Shelf))
    $J += ,@((New-ValueCell $threshold)); $L += ,@((New-ValueCell $source))
    $NO += ,@((New-ValueCell $g.SkuId),(New-ValueCell $g.Spec))
}
while($AG.Count-lt1999){$AG+=,@((New-ValueCell $null),(New-ValueCell $null),(New-ValueCell $null),(New-ValueCell $null),(New-ValueCell $null),(New-ValueCell $null),(New-ValueCell $null));$J+=,@((New-ValueCell $null));$L+=,@((New-ValueCell $null));$NO+=,@((New-ValueCell $null),(New-ValueCell $null))}
Set-CellMatrix '库存快照' 'A2:G2000' $AG
Set-CellMatrix '库存快照' 'J2:J2000' $J
Set-CellMatrix '库存快照' 'L2:L2000' $L
Set-CellMatrix '库存快照' 'N2:O2000' $NO

$skuDiscovery = $null
if (Test-Path -LiteralPath $skuDiscoveryPath) {
    # 直接按文件调用，确保子脚本中的 $PSScriptRoot 指向 agent 目录；动态 ScriptBlock 会把它变成空字符串。
    $discoveryRaw = & $skuDiscoveryPath -ProjectRoot $projectRoot -SnapshotDate $snapshotDate | Out-String
    try { $skuDiscovery = $discoveryRaw | ConvertFrom-Json } catch { throw '新增KW自动建档返回了无法解析的结果' }
    if (-not $skuDiscovery.ok -or $skuDiscovery.formula_verification -ne 'success') { throw '新增KW自动建档或回读校验失败' }
}

$verify = Invoke-LarkJson { lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name '库存快照' --range 'H2:I2000,K2:K2000,M2:M2000' --format json --exit-on-error } '库存快照公式校验调用失败'
if ($verify.data.status -ne 'success') { throw '库存快照公式校验失败' }
$dashboardVerify = Invoke-LarkJson { lark-cli sheets +formula-verify --as user --spreadsheet-token $spreadsheetToken --sheet-name '驾驶舱' --range 'C2:C19' --format json --exit-on-error } '驾驶舱公式校验调用失败'
if ($dashboardVerify.data.status -ne 'success') { throw '驾驶舱公式校验失败' }

$finalResult = [pscustomobject]@{
    inventory_file = $InventoryPath
    exported_rows = ($rows.Count - 2)
    snapshot_skus = $records.Count
    duplicate_sku_ids = @($groups.Values | Where-Object { $_.Count -gt 1 }).Count
    missing_codes = $uncoded.Count
    snapshot_date = $snapshotDate.ToString('yyyy-MM-dd')
    source = $downloadSource
    degraded = $degraded
    export_attempts = if ($exportInfo) { $exportInfo.totalAttempts } else { $null }
    export_recovered = if ($exportInfo) { [bool]$exportInfo.recovered } else { $false }
    export_error = $exportError
    new_kw_count = if ($skuDiscovery) { [int]$skuDiscovery.new_kw_count } else { 0 }
    mapped_existing_count = if ($skuDiscovery) { [int]$skuDiscovery.mapped_existing_count } else { 0 }
    new_kw_codes = if ($skuDiscovery) { @($skuDiscovery.new_kw_codes) } else { @() }
    formula_verification = 'success'
}
Close-InventorySyncTransaction $true
$finalResult | ConvertTo-Json -Compress
