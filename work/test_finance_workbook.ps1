$ErrorActionPreference = 'Stop'
$token = 'JV1aswYrChprbbthehHcJ5O0nCf'

function Set-Cells([string]$Sheet,[string]$Range,[string]$Payload) {
    & lark-cli sheets +cells-set --as user --spreadsheet-token $token --sheet-name $Sheet --range $Range --cells $Payload --format json | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Test write failed: $Sheet $Range" }
}

$daily = @(@(
    @{value=100},@{value=50},@{value=30},@{value=20},@{},
    @{value=80},@{value=40},@{value=25},@{},@{},
    @{value=120},@{value=100},@{value=30},@{value=10},@{value='公式校验临时数据'}
)) | ConvertTo-Json -Depth 8 -Compress
$dailyRow = (ConvertTo-Json -InputObject @(
    @{value=100},@{value=50},@{value=30},@{value=20},@{},
    @{value=80},@{value=40},@{value=25},@{},@{},
    @{value=120},@{value=100},@{value=30},@{value=10},@{value='公式校验临时数据'}
) -Depth 8 -Compress)
Set-Cells '每日记录' 'B2:P2' "[$dailyRow]"

$monthRow = ConvertTo-Json -InputObject @(
    @{value=500},@{value=570},@{value=20},@{value=10},@{value=50},@{value=100},@{value=1000},@{value='公式校验临时数据'}
) -Depth 8 -Compress
Set-Cells '月度补充' 'B2:I2' "[$monthRow]"

$assets = ConvertTo-Json -InputObject @(
    @(@{value='品牌服务费测试'},@{value='品牌服务费摊销'},@{value=120},@{value=46235;cell_styles=@{number_format='yyyy-mm'}},@{value=12}),
    @(@{value='货架测试'},@{value='装修/设备折旧'},@{value=240},@{value=46235;cell_styles=@{number_format='yyyy-mm'}},@{value=24})
) -Depth 8 -Compress
Set-Cells '资产与口径' 'A2:E3' $assets
$assetNotes = ConvertTo-Json -InputObject @(@(@{value='公式校验临时数据'}),@(@{value='公式校验临时数据'})) -Depth 8 -Compress
if ($assetNotes -notmatch '^\[\[') { $assetNotes = '[[{"value":"公式校验临时数据"}],[{"value":"公式校验临时数据"}]]' }
Set-Cells '资产与口径' 'H2:H3' $assetNotes

Start-Sleep -Seconds 2
$raw = & lark-cli sheets +cells-get --as user --spreadsheet-token $token --sheet-name '月度汇总' --range 'A1:X2' --include value,formula --format json
if ($LASTEXITCODE -ne 0) { throw 'Summary readback failed' }
$response = ($raw | Out-String) | ConvertFrom-Json
$row = $response.data.ranges[0].cells[1]
$expected = [ordered]@{
    B=200; C=165; D=120; E=100; F=500; G=570; H=70; I=50; J=150;
    K=30; L=20; M=10; N=10; O=10; P=50; Q=50; R=50; S=100; T=-50;
    U=10; V=25; W=1000; X=1025
}
$letters = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X')
foreach ($entry in $expected.GetEnumerator()) {
    $index = [array]::IndexOf($letters,$entry.Key)
    $actual = [double]$row[$index].value
    if ([math]::Abs($actual - [double]$entry.Value) -gt 0.001) {
        throw "Assertion failed $($entry.Key)2: expected $($entry.Value), got $actual"
    }
}
Write-Output 'TEST_ASSERTIONS_PASSED: 23/23'
