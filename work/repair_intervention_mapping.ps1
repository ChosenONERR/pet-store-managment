param(
    [Parameter(Mandatory=$true)][ValidatePattern('^\d{10,30}$')][string]$IncorrectOrderNo,
    [Parameter(Mandatory=$true)][ValidatePattern('^\d{10,30}$')][string]$CorrectOrderNo,
    [Parameter(Mandatory=$true)][string]$Person,
    [datetime]$Date = [datetime]::Today,
    [switch]$SelfDelivery,
    [switch]$InventoryRestored
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$agentRoot = Join-Path $projectRoot 'agent'
$stateRoot = Join-Path $agentRoot 'state'
$eventScript = Join-Path $agentRoot 'record_finance_event.ps1'
$transactionScript = Join-Path $agentRoot 'finance_transaction.ps1'
$larkCli = Join-Path $agentRoot 'lark_cli_noexit.ps1'
$powershellExe = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$meta = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\simple_finance_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$ledgerMeta = [IO.File]::ReadAllText((Join-Path $projectRoot 'work\daily_ledger_meta.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
$tokens = @{
    source = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
    ledger = ([uri]$ledgerMeta.workbook_url).Segments[-1].Trim('/')
}
. $transactionScript
$script:repairRunId = [datetime]::Now.ToString('yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)

function Invoke-RepairAction([string]$Action, [string]$OrderNo, [string]$EventSuffix, [hashtable]$Extra = @{}) {
    $arguments = @(
        '-NoProfile','-ExecutionPolicy','Bypass','-File',$eventScript,
        '-Action',$Action,'-OrderNo',$OrderNo,'-Date',$Date.ToString('yyyy-MM-dd'),
        '-EventId',"repair-intervention-$($script:repairRunId)-$EventSuffix",'-ProjectRoot',$projectRoot
    )
    foreach ($entry in $Extra.GetEnumerator()) {
        if ($entry.Value -is [System.Management.Automation.SwitchParameter]) {
            if ([bool]$entry.Value) { $arguments += "-$($entry.Key)" }
        } elseif ($null -ne $entry.Value -and [string]$entry.Value -ne '') {
            $arguments += @("-$($entry.Key)",[string]$entry.Value)
        }
    }
    $output = & $powershellExe @arguments 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "Action $Action failed: $($output.Trim())" }
    try { return ($output | ConvertFrom-Json) } catch { throw "Action $Action returned invalid JSON: $($output.Trim())" }
}

$transactionId = "repair-intervention-$($script:repairRunId)"
$transaction = New-FinanceTransaction $stateRoot "Repair intervention order $IncorrectOrderNo -> $CorrectOrderNo" $transactionId -AcquireLock -WorkbookTokens $tokens -LarkCli $larkCli
$closed = $false
try {
    $results = [Collections.Generic.List[object]]::new()
    $results.Add((Invoke-RepairAction 'RemoveIntervention' $IncorrectOrderNo 'remove-incorrect' @{Note='Remove incorrectly classified intervention order'}))
    $results.Add((Invoke-RepairAction 'RemoveIntervention' $CorrectOrderNo 'remove-incomplete' @{Note='Remove incomplete correction record before atomic rebuild'}))
    $results.Add((Invoke-RepairAction 'AddIntervention' $CorrectOrderNo 'add-correct' @{Person=$Person;Note='Re-register corrected intervention order'}))
    if ($SelfDelivery) { $results.Add((Invoke-RepairAction 'ConfirmSelfDelivery' $CorrectOrderNo 'self-delivery' @{Note='Owner confirmed self delivery'})) }
    if ($InventoryRestored) { $results.Add((Invoke-RepairAction 'RestoreInventory' $CorrectOrderNo 'restore-inventory' @{Note='Owner confirmed inventory restored'})) }
    Complete-FinanceTransaction $transaction
    Close-FinanceTransaction $transaction
    $closed = $true
    [ordered]@{
        ok = $true
        transaction_id = $transactionId
        incorrect_order_no = $IncorrectOrderNo
        correct_order_no = $CorrectOrderNo
        actions = @($results | ForEach-Object { $_.action })
        atomic = $true
    } | ConvertTo-Json -Depth 8
} catch {
    $reason = [string]$_.Exception.Message
    $rollback = Invoke-FinanceTransactionRollback $transaction $tokens $larkCli $reason
    if ($rollback.ok) { throw "$reason; all changes rolled back" }
    throw "$reason; automatic rollback incomplete: $($rollback.error)"
} finally {
    if (-not $closed) { Close-FinanceTransaction $transaction }
}
