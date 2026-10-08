$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'agent\run_sync_and_notify.ps1'
$source = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) { throw "主脚本语法错误：$($errors.Message -join '；')" }

$functionAst = $ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-SheetRows'
}, $true) | Select-Object -First 1
if ($null -eq $functionAst) { throw '找不到 Get-SheetRows' }
Invoke-Expression $functionAst.Extent.Text

$script:spreadsheetToken = 'test-token'
$script:readCalls = [Collections.Generic.List[string]]::new()

function Invoke-LarkJson([scriptblock]$Operation, [string]$FailureMessage) {
    $global:LASTEXITCODE = 0
    $raw = & $Operation | Out-String
    if ($LASTEXITCODE -ne 0) { throw $FailureMessage }
    return ($raw | ConvertFrom-Json)
}

function global:lark-cli {
    $arguments = @($args)
    $rangeIndex = [array]::IndexOf($arguments, '--range')
    if ($rangeIndex -lt 0) { throw '测试调用缺少 --range' }
    $range = [string]$arguments[$rangeIndex + 1]
    $script:readCalls.Add($range)

    if ($range -eq 'A2:C9') {
        [ordered]@{ok=$true;data=[ordered]@{has_more=$true;ranges=@()}} | ConvertTo-Json -Depth 10 -Compress
        return
    }
    if ($range -notmatch '^A(\d+):C(\d+)$') { throw "测试收到意外范围：$range" }
    $startRow = [int]$Matches[1]
    $endRow = [int]$Matches[2]
    $cells = [Collections.Generic.List[object]]::new()
    for ($row = $startRow; $row -le $endRow; $row++) {
        $cells.Add(@(
            [ordered]@{value="kw$row"},
            [ordered]@{value="csx$row"},
            [ordered]@{value="商品$row"}
        ))
    }
    [ordered]@{ok=$true;data=[ordered]@{has_more=$false;ranges=@([ordered]@{range=$range;cells=$cells.ToArray()})}} | ConvertTo-Json -Depth 10 -Compress
}

$rows = @(Get-SheetRows 'SKU成本主表' 'A2:C9')
if ($rows.Count -ne 8) { throw "分段读取行数错误：预期8，实际$($rows.Count)" }
if ([string]$rows[0][0].value -ne 'kw2' -or [string]$rows[-1][0].value -ne 'kw9') {
    throw '分段读取没有保持原始行顺序'
}
if (($script:readCalls -join ',') -ne 'A2:C9,A2:C5,A6:C9') {
    throw "分段范围错误：$($script:readCalls -join ',')"
}

[ordered]@{
    ok = $true
    assertions = 3
    rows = $rows.Count
    calls = $script:readCalls.ToArray()
    external_writes = 0
} | ConvertTo-Json -Depth 5
