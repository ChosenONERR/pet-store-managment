param(
    [string]$ProjectRoot,
    [string]$ProfileName = 'viral-video-analyzer'
)

$ErrorActionPreference = 'Stop'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
. (Join-Path $PSScriptRoot 'runtime_paths.ps1')

$profileRoot = Join-Path $env:LOCALAPPDATA "hermes\profiles\$ProfileName"
$gatewayPidPath = Join-Path $profileRoot 'gateway.pid'
if (Test-Path -LiteralPath $gatewayPidPath -PathType Leaf) {
    $rawPid = (Get-Content -LiteralPath $gatewayPidPath -Raw -ErrorAction SilentlyContinue).Trim()
    $gatewayPid = 0
    if ($rawPid -match '^\d+$') { $gatewayPid = [int]$rawPid }
    else { try { $gatewayPid = [int](($rawPid | ConvertFrom-Json).pid) } catch {} }
    if ($gatewayPid -gt 0 -and (Get-Process -Id $gatewayPid -ErrorAction SilentlyContinue)) {
        [ordered]@{ok=$true;already_running=$true;pid=$gatewayPid;console_window=$false} | ConvertTo-Json -Compress
        exit 0
    }
}

$pythonExe = Get-PetstarHermesPython
$pythonwExe = Join-Path (Split-Path $pythonExe -Parent) 'pythonw.exe'
if (-not (Test-Path -LiteralPath $pythonwExe -PathType Leaf)) {
    throw "Hermes virtual environment does not contain pythonw.exe: $pythonwExe"
}
$venvRoot = Split-Path (Split-Path $pythonExe -Parent) -Parent
$agentRoot = Split-Path $venvRoot -Parent

$env:HERMES_HOME = $profileRoot
$env:PYTHONIOENCODING = 'utf-8'
$env:HERMES_GATEWAY_DETACHED = '1'
$env:VIRTUAL_ENV = $venvRoot
$env:PYTHONPATH = if ($env:PYTHONPATH) { "$agentRoot;$env:PYTHONPATH" } else { $agentRoot }

$process = Start-Process -FilePath $pythonwExe -ArgumentList @('-m','hermes_cli.main','--profile',$ProfileName,'gateway','run') -WorkingDirectory $profileRoot -WindowStyle Hidden -PassThru
[ordered]@{ok=$true;already_running=$false;launcher='pythonw.exe';pid=$process.Id;console_window=$false} | ConvertTo-Json -Compress

