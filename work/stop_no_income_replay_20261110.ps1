# Owner-authorized one-off; default read-only. Never use for another message.
param([switch]$Apply,[string]$ProjectRoot=(Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference='Stop'
$hash='39877c253d2e3a5f5874be53568e06bd39622cf59f29fb43188823d52de22686'
$path=Join-Path $ProjectRoot 'agent/state/handled_inbound_messages.ndjson'
$messageMutex=[Threading.Mutex]::new($false,('Local\PetstarLarkMessage_'+$hash.Substring(0,32)))
$ledgerMutex=[Threading.Mutex]::new($false,'Local\PetstarLarkHandledMessageLedger')
$messageTaken=$false;$ledgerTaken=$false
try{
    try{$messageTaken=$messageMutex.WaitOne([timespan]::FromSeconds(30))}catch [Threading.AbandonedMutexException]{$messageTaken=$true}
    if(-not$messageTaken){throw 'Target message is busy; no state changed'}
    try{$ledgerTaken=$ledgerMutex.WaitOne([timespan]::FromSeconds(15))}catch [Threading.AbandonedMutexException]{$ledgerTaken=$true}
    if(-not$ledgerTaken){throw 'Message ledger busy; no state changed'}
    $records=@(Get-Content -LiteralPath $path -Encoding UTF8 | ForEach-Object {if($_.Trim()){$_|ConvertFrom-Json}} | Where-Object {$_.message_hash-eq$hash})
    if(-not$records.Count){throw 'Known failing message not found; no state changed'}
    $terminal=@($records|Where-Object{$_.status-in@('completed','baseline','failed_replied','needs_clarification')})
    if($terminal.Count){[ordered]@{ok=$true;already_stopped=$true;writes=0}|ConvertTo-Json -Compress;return}
    if($records[-1].status-ne'failed_retryable' -or $records[-1].create_time-ne'1794320160000'){throw 'Target status changed; no state changed'}
    if($Apply){
        $record=[ordered]@{message_hash=$hash;status='needs_clarification';source='owner_requested_stop_20261110';create_time=$records[-1].create_time;recorded_at=[datetime]::Now.ToString('o')}
        [IO.File]::AppendAllText($path,(($record|ConvertTo-Json -Compress)+[Environment]::NewLine),[Text.UTF8Encoding]::new($false))
        $last=Get-Content -LiteralPath $path -Tail 1 -Encoding UTF8|ConvertFrom-Json
        if($last.message_hash-ne$hash -or $last.status-ne'needs_clarification'){throw 'Stop record readback failed'}
    }
    [ordered]@{ok=$true;applied=[bool]$Apply;attempts=$records.Count;business_writes=0}|ConvertTo-Json -Compress
}finally{
    if($ledgerTaken){$ledgerMutex.ReleaseMutex()};if($messageTaken){$messageMutex.ReleaseMutex()}
    $ledgerMutex.Dispose();$messageMutex.Dispose()
}
