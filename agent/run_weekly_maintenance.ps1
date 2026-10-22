param(
    [datetime]$Date = [datetime]::Today.AddDays(-1),
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
. (Join-Path $PSScriptRoot 'runtime_paths.ps1')
$powershellExe = Get-PetstarWindowsPowerShell
$auditScript = Join-Path $ProjectRoot 'work\audit_finance_integrity.ps1'
$syncScript = Join-Path $PSScriptRoot 'run_sync_and_notify.ps1'
$statePath = Join-Path $ProjectRoot 'agent\state\weekly_maintenance_last_run.json'

function Write-AtomicJson([string]$Path, $Value) {
    $directory = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $directory)) { $null = New-Item -ItemType Directory -Path $directory }
    $temporary = "$Path.$PID.tmp"
    [IO.File]::WriteAllText($temporary,($Value | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Invoke-JsonChild([string]$Path, [string[]]$Arguments, [string]$Label) {
    $global:LASTEXITCODE = 0
    $output = @(& $powershellExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $Path @Arguments 2>&1 | ForEach-Object { [string]$_ })
    $raw = ($output -join "`n").Trim()
    $parsed = $null
    if ($raw) {
        try { $parsed = $raw | ConvertFrom-Json } catch {}
    }
    if ($null -eq $parsed) {
        $jsonLine = @($output | Where-Object { $_ -and $_.Trim().StartsWith('{') -and $_.Trim().EndsWith('}') } | Select-Object -Last 1)
        if ($jsonLine.Count -gt 0) {
            try { $parsed = ([string]$jsonLine[-1]).Trim() | ConvertFrom-Json } catch {}
        }
    }
    if ($LASTEXITCODE -ne 0 -or $null -eq $parsed) {
        $detail = (($output | Where-Object { $_ -and $_.Trim() }) -join ' ').Trim()
        throw "$Label：$(if($detail){$detail}else{'子进程未返回有效结果'})"
    }
    return $parsed
}

$startedAt = [datetime]::Now
try {
    # 第一阶段严格只读。健康时不创建大事务、不改写任何历史数据。
    $audit = Invoke-JsonChild $auditScript @('-ProjectRoot',$ProjectRoot) '每周全历史审计失败'
    if ($audit.ok) {
        $result = [ordered]@{
            ok=$true;date=$Date.ToString('yyyy-MM-dd');audit_ok=$true;full_rebuild_triggered=$false
            reason='全历史只读审计通过，无需重建';completed_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
            duration_seconds=[math]::Round(([datetime]::Now-$startedAt).TotalSeconds,2)
        }
        Write-AtomicJson $statePath $result
        $result | ConvertTo-Json -Depth 8
        exit 0
    }

    # 只有完整性审计明确失败，才进入原子化全量重建。重建失败时主同步器会发飞书告警并回滚。
    $rebuild = Invoke-JsonChild $syncScript @(
        '-Date',$Date.ToString('yyyy-MM-dd'),'-Mode','manual','-SkipSync','-SkipInventory','-Silent','-FullRebuild','-ProjectRoot',$ProjectRoot
    ) '每周全量重建失败'
    $result = [ordered]@{
        ok=[bool]$rebuild.ok;date=$Date.ToString('yyyy-MM-dd');audit_ok=$false;full_rebuild_triggered=$true
        audit=$audit;rebuild=$rebuild;completed_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
        duration_seconds=[math]::Round(([datetime]::Now-$startedAt).TotalSeconds,2)
    }
    Write-AtomicJson $statePath $result
    $result | ConvertTo-Json -Depth 10
} catch {
    $failure = [ordered]@{
        ok=$false;date=$Date.ToString('yyyy-MM-dd');error=$_.Exception.Message
        completed_at=[datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
        duration_seconds=[math]::Round(([datetime]::Now-$startedAt).TotalSeconds,2)
    }
    try { Write-AtomicJson $statePath $failure } catch {}
    $failure | ConvertTo-Json -Depth 8
    exit 1
}
