param(
    [string]$ProjectRoot,
    [string]$StartupName = 'Hermes_Gateway_viral-video-analyzer.vbs'
)

$ErrorActionPreference = 'Stop'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$launcherPath = Join-Path $ProjectRoot 'agent\run_hidden_gateway.vbs'
if (-not (Test-Path -LiteralPath $launcherPath -PathType Leaf)) { throw "Hidden gateway launcher not found: $launcherPath" }

$startupRoot = [Environment]::GetFolderPath('Startup')
$startupPath = Join-Path $startupRoot $StartupName
if (Test-Path -LiteralPath $startupPath -PathType Leaf) {
    $stamp = [datetime]::Now.ToString('yyyyMMdd-HHmmss')
    Copy-Item -LiteralPath $startupPath -Destination "$startupPath.pre_petstar_$stamp" -Force
}

$escapedLauncher = $launcherPath.Replace('"','""')
$content = @"
Option Explicit
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "wscript.exe //B //Nologo ""$escapedLauncher""", 0, False
"@
[IO.File]::WriteAllText($startupPath, $content, [Text.UTF8Encoding]::new($false))

$written = [IO.File]::ReadAllText($startupPath)
if ($written.IndexOf($launcherPath, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
    throw 'The Windows startup launcher could not be verified.'
}
[ordered]@{ok=$true;startup_file=$startupPath;launcher=$launcherPath;applies_after_next_sign_in=$true} | ConvertTo-Json -Compress

