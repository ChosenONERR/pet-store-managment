param([string]$ProjectRoot)

$ErrorActionPreference = 'Stop'
if (-not $ProjectRoot) { $ProjectRoot = Split-Path $PSScriptRoot -Parent }
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
. (Join-Path $PSScriptRoot 'runtime_paths.ps1')

function Test-ResolvedPath([string]$Name, [scriptblock]$Resolver, [bool]$Required = $true) {
    try {
        $value = & $Resolver
        return [ordered]@{ name=$Name; required=$Required; ok=$true; path=[string]$value; detail='' }
    } catch {
        return [ordered]@{ name=$Name; required=$Required; ok=$false; path=''; detail=$_.Exception.Message }
    }
}

$checks = [Collections.Generic.List[object]]::new()
$checks.Add((Test-ResolvedPath 'Windows PowerShell' { Get-PetstarWindowsPowerShell }))
$checks.Add((Test-ResolvedPath 'Python' { Get-PetstarPython }))
$checks.Add((Test-ResolvedPath 'Hermes Python (legacy scheduler only)' { Get-PetstarHermesPython } $false))
$checks.Add((Test-ResolvedPath 'Lark CLI' { Get-PetstarLarkCli }))
$checks.Add((Test-ResolvedPath 'web-access dependency check' { Get-PetstarWebAccessCheck }))
$checks.Add((Test-ResolvedPath 'Node.js' {
    $node = Get-Command node.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1
    $node.Source
}))
$checks.Add((Test-ResolvedPath 'Codex CLI' {
    $codex = Get-Command codex.exe,codex.cmd -CommandType Application -ErrorAction Stop | Select-Object -First 1
    $codex.Source
}))

foreach ($relativePath in @('work\simple_finance_meta.json','work\daily_ledger_meta.json')) {
    $path = Join-Path $ProjectRoot $relativePath
    $checks.Add([ordered]@{
        name = $relativePath
        required = $true
        ok = Test-Path -LiteralPath $path -PathType Leaf
        path = $path
        detail = if (Test-Path -LiteralPath $path -PathType Leaf) { '' } else { 'Copy the matching .example.json file and fill in the workbook URL.' }
    })
}

$statePath = Join-Path $ProjectRoot 'agent\state'
$stateItem = Get-Item -LiteralPath $statePath -Force -ErrorAction SilentlyContinue
$checks.Add([ordered]@{
    name = 'runtime state'
    required = $true
    ok = $null -ne $stateItem
    path = $statePath
    detail = if ($null -eq $stateItem) { 'Restore the state backup or create the state directory/junction.' } else { "link_type=$($stateItem.LinkType); target=$(@($stateItem.Target) -join ',')" }
})

$failed = @($checks | Where-Object { $_.required -and -not $_.ok })
$warnings = @($checks | Where-Object { -not $_.required -and -not $_.ok })
[ordered]@{
    ok = $failed.Count -eq 0
    project_root = $ProjectRoot
    checks = @($checks)
    failed_count = $failed.Count
    warning_count = $warnings.Count
} | ConvertTo-Json -Depth 5

if ($failed.Count -gt 0) { exit 1 }
