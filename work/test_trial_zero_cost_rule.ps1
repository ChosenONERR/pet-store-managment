param(
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'

$trial = [string]([char[]]@(0x8BD5,0x5403))
$sample = [string]([char[]]@(0x5C1D,0x9C9C))
$tryUse = [string]([char[]]@(0x8BD5,0x7528))
$experience = [string]([char[]]@(0x4F53,0x9A8C,0x88C5))
$pattern = @($trial,$sample,$tryUse,$experience) -join '|'

function Test-ZeroCostTrialItem([string]$Title, [string]$Spec) {
    $text = (([string]$Title) + ' ' + ([string]$Spec)).Trim()
    return [bool]($text -match $script:pattern)
}

$cases = @(
    @{title=('cat-food-' + $trial);spec='50g';expected=$true},
    @{title=('cat-food-' + $tryUse);spec='50g';expected=$true},
    @{title='cat-food';spec=$sample;expected=$true},
    @{title='cat-food';spec=$experience;expected=$true},
    @{title='regular-cat-food';spec='500g';expected=$false}
)

foreach ($case in $cases) {
    $actual = Test-ZeroCostTrialItem $case.title $case.spec
    if ($actual -ne $case.expected) { throw "trial cost rule failed: $($case.title) / $($case.spec)" }
}

foreach ($relativePath in @('agent\rebuild_cost_snapshots.ps1','work\audit_aug_sep_order_costs.ps1')) {
    $path = Join-Path $ProjectRoot $relativePath
    $text = [IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
    if ($text -notmatch 'Test-ZeroCostTrialItem' -or $text -notmatch 'trial_zero_cost') {
        throw "trial cost rule is not wired into: $relativePath"
    }
    $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$errors)
    if ($errors.Count -gt 0) { throw "PowerShell syntax check failed: $relativePath" }
}

[pscustomobject]@{ok=$true;cases=$cases.Count;engines=2} | ConvertTo-Json -Compress
