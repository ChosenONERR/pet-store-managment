param(
    [string]$ProjectRoot,
    [string]$TaskName = 'Petstar-Lark-Event-Watchdog'
)

$ErrorActionPreference = 'Stop'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path

$wrapperPath = Join-Path $ProjectRoot 'agent\run_hidden_watchdog.vbs'
$wscriptPath = Join-Path $env:SystemRoot 'System32\wscript.exe'
if (-not (Test-Path -LiteralPath $wrapperPath -PathType Leaf)) { throw "Hidden watchdog launcher not found: $wrapperPath" }
if (-not (Test-Path -LiteralPath $wscriptPath -PathType Leaf)) { throw "Windows Script Host not found: $wscriptPath" }

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
$action = New-ScheduledTaskAction -Execute $wscriptPath -Argument ('//B //Nologo "{0}"' -f $wrapperPath)
$null = Set-ScheduledTask -TaskName $TaskName -TaskPath $task.TaskPath -Action $action

$updated = Get-ScheduledTask -TaskName $TaskName -TaskPath $task.TaskPath
$actualAction = @($updated.Actions)[0]
$isHiddenLauncher = (
    [IO.Path]::GetFullPath([string]$actualAction.Execute) -eq [IO.Path]::GetFullPath($wscriptPath) -and
    [string]$actualAction.Arguments -like "*$wrapperPath*"
)
if (-not $isHiddenLauncher) { throw 'The watchdog task did not switch to the hidden launcher.' }

[ordered]@{
    ok = $true
    task_name = $TaskName
    state = [string]$updated.State
    launcher = 'wscript.exe'
    interactive_console = $false
} | ConvertTo-Json -Compress
