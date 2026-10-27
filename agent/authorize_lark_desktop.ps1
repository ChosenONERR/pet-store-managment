param(
    [Parameter(Mandatory = $true)]
    [string]$RequestPath,
    [Parameter(Mandatory = $true)]
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

# Force the normal Windows-user Lark workspace. Host-agent markers otherwise
# redirect lark-cli to an isolated Hermes/OpenClaw/Lark Channel credential set.
foreach ($variableName in @(
    'HERMES_HOME','HERMES_QUIET','HERMES_EXEC_ASK','HERMES_GATEWAY_TOKEN','HERMES_SESSION_KEY',
    'OPENCLAW_CLI','OPENCLAW_HOME','OPENCLAW_STATE_DIR','OPENCLAW_CONFIG_PATH',
    'OPENCLAW_SERVICE_MARKER','OPENCLAW_SERVICE_VERSION','OPENCLAW_GATEWAY_PORT','OPENCLAW_SHELL',
    'LARK_CHANNEL'
)) {
    Remove-Item -LiteralPath "Env:$variableName" -ErrorAction SilentlyContinue
}
$env:LARKSUITE_CLI_CONFIG_DIR = Join-Path $env:USERPROFILE '.lark-cli'
$larkCli = Join-Path $env:APPDATA 'npm\lark-cli.cmd'
if (-not (Test-Path -LiteralPath $larkCli -PathType Leaf)) { throw "lark-cli not found: $larkCli" }

$requestRaw = & $larkCli auth login --recommend --no-wait --json | Out-String
if ($LASTEXITCODE -ne 0) { throw "Unable to start Lark authorization: exit=$LASTEXITCODE" }
$request = $requestRaw | ConvertFrom-Json
if ($request.ok -eq $false) { throw "Unable to start Lark authorization: $($request.error.message)" }

$requestDirectory = Split-Path $RequestPath -Parent
if ($requestDirectory -and -not (Test-Path -LiteralPath $requestDirectory)) {
    $null = New-Item -ItemType Directory -Path $requestDirectory
}
[IO.File]::WriteAllText($RequestPath, $requestRaw, [Text.UTF8Encoding]::new($false))

$deviceCode = [string]$request.device_code
if (-not $deviceCode -and $request.data) { $deviceCode = [string]$request.data.device_code }
if (-not $deviceCode) { throw 'Lark authorization did not return a device code.' }

Write-Host 'Open the authorization URL or scan the QR code shown by Codex.'
$completionRaw = & $larkCli auth login --device-code $deviceCode --json | Out-String
if ($LASTEXITCODE -ne 0) { throw "Lark authorization did not complete: exit=$LASTEXITCODE" }
$completion = $completionRaw | ConvertFrom-Json
if ($completion.ok -eq $false) { throw "Lark authorization did not complete: $($completion.error.message)" }

$statusRaw = & $larkCli auth status --json --verify | Out-String
$status = $statusRaw | ConvertFrom-Json
$ready = ($LASTEXITCODE -eq 0 -and $status.identities.user.status -eq 'ready' -and $status.identities.user.verified -eq $true)
$result = [ordered]@{
    ok = $ready
    user_status = [string]$status.identities.user.status
    verified = [bool]$status.identities.user.verified
    token_status = [string]$status.identities.user.tokenStatus
    completed_at = [datetime]::Now.ToString('o')
}
[IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
if (-not $ready) { throw 'Lark authorization completed but verification is not ready.' }
Write-Host 'Lark authorization is ready. This window can close.'
Start-Sleep -Seconds 5

