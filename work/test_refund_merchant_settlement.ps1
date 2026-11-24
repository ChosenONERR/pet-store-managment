$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'agent/refund_adjustments.ps1')

# Exercise the real refund rules with an in-memory state store.
$script:fixtureState = New-RefundAdjustmentState
function Read-RefundAdjustmentState([string]$ProjectRoot) { return $script:fixtureState }
function Write-RefundAdjustmentState([string]$ProjectRoot, $State) { $script:fixtureState = $State }
function Assert([bool]$Pass, [string]$Message) { if (-not $Pass) { throw $Message } }
$order = [pscustomobject]@{
    orderNo='3000000000000000001'; orderDate='2026-11-22'; expectedIncome=72.38
    refunds=@([pscustomobject]@{
        recordId='99901'; completed=$true; refundType='refund'; completedAt='2026-11-24 12:00'; requestedAt='2026-11-24 10:00'; reason='test'
        amount=63; items=@([pscustomobject]@{goodsCode='kw9999';productId='123';title='test';spec='single';quantity=1;refundAmount=63})
    })
}
$result = Merge-RefundDiscoveries 'offline' @($order)
Assert ($result.newly_detected[0].income_adjustment -eq 63) 'Existing behavior changed without finance evidence'
$null = Set-RefundItemDisposition 'offline' $order.orderNo 'kw9999' returned_restocked 'confirmed return'
$refund = Set-RefundMerchantSettlement 'offline' $order.orderNo '99901' 72.38
Assert ((Test-VerifiedRefundMerchantSettlement $refund) -and $refund.customer_refund_amount -eq 63) 'Seller reversal overwritten into buyer refund'
$after = Merge-RefundDiscoveries 'offline' @($order)
$event = @($after.state.events)[0]
Assert ($event.income_adjustment -eq 72.38 -and $event.base_expected_income -eq 72.38) 'Later order sync overwrote verified merchant settlement'
Assert ($event.items[0].disposition -eq 'returned_restocked' -and $event.items[0].quantity -eq 1) 'Merchant settlement changed item disposition or quantity'
$summary = Get-RefundAccountingSummary 'offline' $order.orderNo
Assert ($summary.income_adjustment -eq 72.38 -and $summary.returned_quantity_by_code['kw9999'] -eq 1) 'Merchant adjustment or returns were counted twice'
$blocked = $false
try { $null = Set-RefundMerchantSettlement 'offline' $order.orderNo '99902' 72.38 } catch { $blocked = $true }
Assert $blocked 'Wrong refund identity was accepted'
$order.refunds[0].amount = 62
$next = Merge-RefundDiscoveries 'offline' @($order)
Assert ($next.state.events[0].income_adjustment -eq 72.38 -and $next.state.events[0].customer_refund_amount -eq 62) 'Amounts did not remain independent after resync'
'{"ok":true,"cases":7,"external_writes":0}'
