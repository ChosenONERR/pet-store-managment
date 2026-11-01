function Get-PetstarWindowsPowerShell {
    $candidates = [Collections.Generic.List[string]]::new()
    if ($env:SystemRoot) {
        $candidates.Add((Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'))
        $candidates.Add((Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'))
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    $command = Get-Command powershell.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command -and $command.Source -and (Test-Path -LiteralPath $command.Source -PathType Leaf)) {
        return $command.Source
    }
    throw 'Windows PowerShell was not found. Check the Windows system components.'
}

function Resolve-PetstarExecutable {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [string[]]$Candidates = @(),
        [string[]]$CommandNames = @()
    )

    foreach ($candidate in @($Candidates)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $expanded = [Environment]::ExpandEnvironmentVariables($candidate)
        if (Test-Path -LiteralPath $expanded -PathType Leaf) {
            return (Resolve-Path -LiteralPath $expanded).Path
        }
    }
    foreach ($commandName in @($CommandNames)) {
        $command = Get-Command $commandName -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command -and $command.Source -and (Test-Path -LiteralPath $command.Source -PathType Leaf)) {
            return $command.Source
        }
    }
    throw "$Label was not found. Configure the matching PETSTAR_* environment variable or run agent\check_environment.ps1."
}

function Get-PetstarUserProfile {
    $userProfile = [Environment]::GetFolderPath('UserProfile')
    if (-not $userProfile) { $userProfile = [string]$env:USERPROFILE }
    if (-not $userProfile) { throw 'Cannot resolve the current Windows user profile path' }
    return $userProfile
}

function Get-PetstarPython {
    $userProfile = Get-PetstarUserProfile
    return Resolve-PetstarExecutable -Label 'Python' -Candidates @(
        $env:PETSTAR_PYTHON,
        (Join-Path $userProfile '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'),
        'D:\CodexData\cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe',
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python313\python.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python311\python.exe')
    ) -CommandNames @('python.exe', 'python3.exe')
}

function Get-PetstarHermesPython {
    $profileRoot = Join-Path $env:LOCALAPPDATA 'hermes\profiles\viral-video-analyzer'
    $gatewayCommand = Join-Path $profileRoot 'gateway-service\Hermes_Gateway_viral-video-analyzer.cmd'
    $candidates = [Collections.Generic.List[string]]::new()
    if ($env:PETSTAR_HERMES_PYTHON) { $candidates.Add($env:PETSTAR_HERMES_PYTHON) }
    if (Test-Path -LiteralPath $gatewayCommand -PathType Leaf) {
        $text = [IO.File]::ReadAllText($gatewayCommand)
        $match = [regex]::Match($text, '(?im)^\s*(?<python>(?:"[^"]+python\.exe"|[^\r\n]+?python\.exe))\s+-m\s+hermes_cli\.main')
        if ($match.Success) { $candidates.Add($match.Groups['python'].Value.Trim().Trim('"')) }
    }
    $candidates.Add('D:\hermes\hermes-agent\venv\Scripts\python.exe')
    return Resolve-PetstarExecutable -Label 'Hermes Python' -Candidates $candidates.ToArray()
}

function Get-PetstarLarkCli {
    return Resolve-PetstarExecutable -Label 'Lark CLI' -Candidates @(
        $env:PETSTAR_LARK_CLI,
        (Join-Path $env:APPDATA 'npm\lark-cli.cmd')
    ) -CommandNames @('lark-cli.cmd', 'lark-cli.exe')
}

function Get-PetstarWebAccessCheck {
    $userProfile = Get-PetstarUserProfile
    $roots = @(
        $env:PETSTAR_WEB_ACCESS_CHECK,
        (Join-Path $PSScriptRoot 'check_meituan_browser.mjs'),
        (Join-Path $userProfile '.agents\skills\web-access\scripts\check-deps.mjs'),
        (Join-Path $userProfile '.codex\skills\web-access\scripts\check-deps.mjs')
    )
    return Resolve-PetstarExecutable -Label 'web-access dependency check' -Candidates $roots
}
