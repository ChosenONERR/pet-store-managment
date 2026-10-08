param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$agentRoot = Join-Path $ProjectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
. (Join-Path $agentRoot 'finance_transaction.ps1')

$sourceMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\simple_finance_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'work\daily_ledger_meta.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$sourceToken = ([uri]$sourceMeta.workbook_url).Segments[-1].Trim('/')
$ledgerToken = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
$tokens = @{source=$sourceToken;ledger=$ledgerToken}
$transaction = New-FinanceTransaction $stateRoot '撤销误记的平台实际到账 2026-10-07' 'repair-actual-receipt-20261008' -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli

function Invoke-LarkJson([scriptblock]$Action, [string]$Label) {
    $raw = & $Action | Out-String
    $exitCode = $LASTEXITCODE
    try { $result = $raw | ConvertFrom-Json } catch { throw "$Label：返回内容无法解析" }
    if ($exitCode -ne 0 -or -not $result.ok) {
        $detail = if ($result.error -and $result.error.message) { [string]$result.error.message } else { $raw.Trim() }
        throw "$Label：$detail"
    }
    return $result
}

try {
    $before = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $sourceToken --sheet-name '月度结算' --range 'A4:L4' --include 'value,formula' --max-chars 20000 --format json } '读取月度结算失败'
    $row = @($before.data.ranges[0].cells)[0]
    $month = [string]$row[0].value
    $receipt = [string]$row[1].value
    $note = [string]$row[11].value
    if ($month -ne '2026-10') { throw "月度结算第4行不是2026-10，已停止：$month" }
    if ($receipt -and [math]::Abs(([double]$receipt) - 34.65) -gt 0.001) { throw "2026-10美团实际到账不是误写的34.65，已停止：$receipt" }

    $noteParts = @([regex]::Split($note, '[；;]') | ForEach-Object { $_.Trim() } | Where-Object {
        $_ -and
        $_ -notmatch '^10-07\s+美团\s+实际到账\s+34\.65(?:：|$)' -and
        $_ -notmatch '^已与客户私下沟通并退款3\.80元，按商家预计收入38\.45元调整为最终实际收入34\.65元'
    })
    $newNote = $noteParts -join '；'

    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '月度结算' 'B4:B4' $larkCli
    Register-FinanceRangeBeforeImage $transaction 'source' $sourceToken '月度结算' 'L4:L4' $larkCli
    $ranges = '["月度结算!B4:B4"]'
    $null = Invoke-LarkJson { $ranges | & $larkCli sheets +cells-batch-clear --as user --spreadsheet-token $sourceToken --ranges - --scope content --yes --format json } '清除误记实际到账失败'
    $encodedNote = ConvertTo-Json -InputObject $newNote -Compress
    $notePayload = "[[{`"value`":$encodedNote}]]"
    $null = Invoke-LarkJson { $notePayload | & $larkCli sheets +cells-set --as user --spreadsheet-token $sourceToken --sheet-name '月度结算' --range 'L4:L4' --cells - --format json } '清理月度备注失败'

    $after = Invoke-LarkJson { & $larkCli sheets +cells-get --as user --spreadsheet-token $sourceToken --sheet-name '月度结算' --range 'A4:L4' --include 'value,formula' --max-chars 20000 --format json } '回读月度结算失败'
    $afterRow = @($after.data.ranges[0].cells)[0]
    $afterReceipt = [string]$afterRow[1].value
    $afterNote = [string]$afterRow[11].value
    if ($afterReceipt) { throw "误记实际到账清除后仍有值：$afterReceipt" }
    if ($afterNote -match '10-07\s+美团\s+实际到账\s+34\.65|已与客户私下沟通并退款3\.80元，按商家预计收入38\.45元调整为最终实际收入34\.65元') { throw '误记实际到账备注未清理' }

    Complete-FinanceTransaction $transaction
    [ordered]@{
        ok = $true
        month = '2026-10'
        cleared_meituan_actual_receipt = 34.65
        note_cleaned = $true
        completed_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    } | ConvertTo-Json -Depth 5
} catch {
    $rollback = Invoke-FinanceTransactionRollback $transaction $tokens $larkCli ([string]$_.Exception.Message)
    throw "撤销误记实际到账失败：$($_.Exception.Message)；回滚状态=$($rollback.status)"
} finally {
    Close-FinanceTransaction $transaction
}
