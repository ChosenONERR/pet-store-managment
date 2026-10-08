$ErrorActionPreference = 'Stop'
$script:cliArgs = @($args)
$statePath = [string]$env:PETSTAR_FAKE_LARK_STATE
if (-not $statePath) { throw '缺少 PETSTAR_FAKE_LARK_STATE' }

function Get-ArgumentValue([string]$Name) {
    $index = [array]::IndexOf([object[]]$script:cliArgs, $Name)
    if ($index -lt 0 -or $index + 1 -ge $script:cliArgs.Count) { return $null }
    return [string]$script:cliArgs[$index + 1]
}
function Read-State { return ([IO.File]::ReadAllText($statePath,[Text.Encoding]::UTF8) | ConvertFrom-Json) }
function Write-State($State) { [IO.File]::WriteAllText($statePath,(($State|ConvertTo-Json -Depth 50)+[Environment]::NewLine),[Text.UTF8Encoding]::new($false)) }
function Get-RangeValue($State,[string]$Key) {
    $property = $State.ranges.PSObject.Properties[$Key]
    if ($null -eq $property) { return @() }
    return @($property.Value)
}
function Set-RangeValue($State,[string]$Key,$Value) {
    $State.ranges | Add-Member -NotePropertyName $Key -NotePropertyValue @($Value) -Force
}
function ConvertFrom-ColumnLetters([string]$Letters) {
    $value = 0
    foreach ($character in $Letters.ToUpperInvariant().ToCharArray()) { $value = ($value * 26) + ([int][char]$character - [int][char]'A' + 1) }
    return $value
}
function ConvertTo-ColumnLetters([int]$Number) {
    $letters = ''
    while ($Number -gt 0) {
        $Number--
        $letters = [char](65 + ($Number % 26)) + $letters
        $Number = [math]::Floor($Number / 26)
    }
    return $letters
}

$shortcut = if ($script:cliArgs.Count -gt 1) { [string]$script:cliArgs[1] } else { '' }
$token = Get-ArgumentValue '--spreadsheet-token'
if ($shortcut -eq '+cells-get') {
    $sheet = Get-ArgumentValue '--sheet-name'
    $range = Get-ArgumentValue '--range'
    $state = Read-State
    $cells = Get-RangeValue $state "$token|$sheet|$range"
    $startRow = if ($range -match '^[A-Z]+(\d+):') { [int]$Matches[1] } else { 1 }
    $rowIndices = @()
    for ($i=0;$i-lt$cells.Count;$i++) { $rowIndices += $startRow + $i }
    $columnIndices = @()
    if ($range -match '^([A-Z]+)\d+:([A-Z]+)\d+$') {
        $startColumn = ConvertFrom-ColumnLetters $Matches[1]
        $endColumn = ConvertFrom-ColumnLetters $Matches[2]
        for ($column=$startColumn; $column -le $endColumn; $column++) { $columnIndices += ConvertTo-ColumnLetters $column }
    }
    [ordered]@{ok=$true;data=[ordered]@{has_more=$false;ranges=@([ordered]@{range=$range;cells=$cells;row_indices=$rowIndices;col_indices=$columnIndices})}} | ConvertTo-Json -Depth 50 -Compress
    exit 0
}
if ($shortcut -eq '+cells-batch-clear') {
    $payload = ($input | Out-String).Trim()
    $ranges = @($payload | ConvertFrom-Json)
    $state = Read-State
    $clearCalls = @()
    $clearProperty = $state.PSObject.Properties['clear_calls']
    if ($null -ne $clearProperty) { $clearCalls = @($clearProperty.Value) }
    foreach ($qualifiedRange in $ranges) {
        $text = [string]$qualifiedRange
        $separator = $text.IndexOf('!')
        if ($separator -lt 1) { throw "unsupported fake clear range: $text" }
        $sheet = $text.Substring(0,$separator)
        $range = $text.Substring($separator + 1)
        $clearCalls += "$token|$sheet|$range"
        Set-RangeValue $state "$token|$sheet|$range" @()
    }
    $state | Add-Member -NotePropertyName clear_calls -NotePropertyValue @($clearCalls) -Force
    Write-State $state
    [ordered]@{ok=$true;data=[ordered]@{cleared=$ranges.Count}} | ConvertTo-Json -Compress
    exit 0
}
if ($shortcut -eq '+batch-update') {
    if ($env:PETSTAR_FAKE_LARK_FAIL_BATCH -eq '1') { [Console]::Error.WriteLine('injected batch failure'); exit 9 }
    $payload = ($input | Out-String).Trim()
    $operations = @($payload | ConvertFrom-Json)
    $state = Read-State
    foreach ($operation in $operations) {
        if ([string]$operation.shortcut -ne '+cells-set') { throw "unsupported fake operation: $($operation.shortcut)" }
        $sheet = [string]$operation.input.sheet_name
        $range = [string]$operation.input.range
        Set-RangeValue $state "$token|$sheet|$range" @($operation.input.cells)
    }
    Write-State $state
    [ordered]@{ok=$true;data=[ordered]@{operation_count=$operations.Count}} | ConvertTo-Json -Compress
    exit 0
}
throw "unsupported fake shortcut: $shortcut"
