param([string]$OutputPath = (Join-Path $PSScriptRoot 'state\lark_background_probe.json'))
$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'runtime_paths.ps1')
Initialize-PetstarLarkUserEnvironment
$env:LARK_CLI_NO_PROXY = '1'
$cli = Join-Path $PSScriptRoot 'lark_cli_noexit.ps1'
$result = [ordered]@{
    checked_at = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    windows_identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    session_id = [Diagnostics.Process]::GetCurrentProcess().SessionId
    config_directory = $env:LARKSUITE_CLI_CONFIG_DIR
    registry_value_count = 0
    read_ok = $false
    read_exit_code = -1
    error_type = $null
    error_code = $null
    failure_stage = $null
    exception_type = $null
    user_status = 'unknown'
    token_status = 'unknown'
}
$key = Get-Item -LiteralPath 'Registry::HKEY_CURRENT_USER\Software\LarkCli\keychain\lark-cli' -ErrorAction SilentlyContinue
if ($null -ne $key) { $result.registry_value_count = @($key.Property).Count }
try {
    $meta = Get-Content -LiteralPath (Join-Path (Split-Path $PSScriptRoot -Parent) 'work\simple_finance_meta.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $book = ([uri]$meta.workbook_url).Segments[-1].Trim('/')
    $result.failure_stage = 'read'
    $raw = & $cli sheets +cells-get --as user --spreadsheet-token $book --sheet-name '驾驶舱' --range A1:A1 --include value --max-chars 2000 --format json 2>$null | Out-String
    $result.read_exit_code = $LASTEXITCODE
    $result.failure_stage = 'parse'
    $response = $raw | ConvertFrom-Json
    $result.failure_stage = 'classify'
    $result.read_ok = ($result.read_exit_code -eq 0 -and $response.ok -ne $false)
    if (-not $result.read_ok -and $response.error) {
        $result.error_type = [string]$response.error.type
        $result.error_code = [string]$response.error.code
    }
    if ($result.read_ok) { $result.failure_stage = $null }
} catch {
    $result.error_type = 'probe_failed'
    $result.exception_type = $_.Exception.GetType().FullName
}
try {
    $status = (@(& $cli auth status --json 2>$null) -join "`n") | ConvertFrom-Json
    $result.user_status = [string]$status.identities.user.status
    $result.token_status = [string]$status.identities.user.tokenStatus
} catch {}
# The CLI response is never persisted: only this fixed allowlist is written.
[IO.File]::WriteAllText($OutputPath, ($result | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
if (-not $result.read_ok) { exit 1 }
