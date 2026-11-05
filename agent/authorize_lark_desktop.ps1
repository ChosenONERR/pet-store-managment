param(
    [Parameter(Mandatory = $true)]
    [string]$RequestPath,
    [Parameter(Mandatory = $true)]
    [string]$ResultPath,
    [switch]$Resume
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

function Invoke-DesktopLarkJson([string[]]$Arguments) {
    # Device-flow progress is native stderr, not an authorization failure.
    # Windows PowerShell 5.1 + Stop otherwise aborts before the user can approve.
    $priorPreference=$ErrorActionPreference
    try {
        $ErrorActionPreference='Continue'
        $raw=& $larkCli @Arguments 2>$null | Out-String
        $code=$LASTEXITCODE
    } finally {$ErrorActionPreference=$priorPreference}
    if($code -ne 0){
        $kind='authorization_command_failed'
        try {
            $errorResult=$raw|ConvertFrom-Json
            $message=[string]$errorResult.error.message
            if($message -match '(?i)expired|过期'){$kind='authorization_expired'}
            elseif($message -match '(?i)pending|等待'){$kind='authorization_pending'}
            elseif($message -match '(?i)denied|拒绝'){$kind='authorization_denied'}
            elseif($message -match '(?i)decrypt|keychain|credential'){$kind='credential_store_error'}
        }catch{}
        throw "$kind (exit=$code)"
    }
    $parsed=$raw|ConvertFrom-Json
    if($parsed.ok -eq $false){throw 'Lark authorization returned an unsuccessful result'}
    return [pscustomobject]@{raw=$raw;value=$parsed}
}

try {
if($Resume) {
    if(-not (Test-Path -LiteralPath $RequestPath)){throw 'No pending desktop authorization to resume'}
    $request=Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8|ConvertFrom-Json
} else {
    $started=Invoke-DesktopLarkJson @('auth','login','--recommend','--no-wait','--json')
    $requestRaw=$started.raw;$request=$started.value

$requestDirectory = Split-Path $RequestPath -Parent
if ($requestDirectory -and -not (Test-Path -LiteralPath $requestDirectory)) {
    $null = New-Item -ItemType Directory -Path $requestDirectory
}
[IO.File]::WriteAllText($RequestPath, $requestRaw, [Text.UTF8Encoding]::new($false))
}

$deviceCode = [string]$request.device_code
if (-not $deviceCode -and $request.data) { $deviceCode = [string]$request.data.device_code }
if (-not $deviceCode) { throw 'Lark authorization did not return a device code.' }

Write-Host 'Open the authorization URL or scan the QR code shown by Codex.'
$null=Invoke-DesktopLarkJson @('auth','login','--device-code',$deviceCode,'--json')

$status=(Invoke-DesktopLarkJson @('auth','status','--json','--verify')).value
$ready = ($status.identities.user.status -eq 'ready' -and $status.identities.user.verified -eq $true)
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
} catch {
    $safeError='desktop_authorization_failed'
    if($_.Exception.Message -match '^(authorization_command_failed|authorization_expired|authorization_pending|authorization_denied|credential_store_error) \(exit=\d+\)$'){$safeError=$_.Exception.Message}
    $failure=[ordered]@{ok=$false;error=$safeError;completed_at=[datetime]::Now.ToString('o')}
    [IO.File]::WriteAllText($ResultPath,($failure|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
    throw
}

