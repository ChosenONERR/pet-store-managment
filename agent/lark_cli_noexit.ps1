# The event consumer writes normal lifecycle messages to stderr.
# Windows PowerShell 5.1 wraps native stderr as ErrorRecord, so stopping on it
# would terminate the listener even while the CLI is healthy.
$ErrorActionPreference = 'Continue'
$npmRoot = Join-Path $env:APPDATA 'npm'
$bundledNode = Join-Path $npmRoot 'node.exe'
$nodeExe = if (Test-Path -LiteralPath $bundledNode) { $bundledNode } else { 'node.exe' }
$runner = Join-Path $npmRoot 'node_modules\@larksuite\cli\scripts\run.js'
if (-not (Test-Path -LiteralPath $runner)) { throw "Lark CLI runner not found: $runner" }

if ($MyInvocation.ExpectingInput) {
    $input | & $nodeExe $runner @args
} else {
    & $nodeExe $runner @args
}
$global:LASTEXITCODE = $LASTEXITCODE
