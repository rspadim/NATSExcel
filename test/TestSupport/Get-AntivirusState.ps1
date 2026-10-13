#Requires -Version 5.1
<#
    Get-AntivirusState.ps1 - reports the active antivirus products and their
    real-time state, so a test run is attributable to its environment.

    Excel COM automation is inherently flaky and security software widens the
    window (see AGENTS.md). A run that fails with RPC_E_CALL_REJECTED, a bogus
    InvalidCastException or a hung COM call is very likely environment noise,
    not an add-in defect - but only if the AV state is known. Every layer prints
    this banner first.

    productState is a bit field; the byte at bits 8-15 is the real-time state:
      0x10 = on (enabled)
      0x11 = on (enabled, and the definitions are up to date)  -> printed as "on"
      0x12 = paused / snoozed
    The full value is printed too, so a variant prefix is still readable.

    Usage:
      powershell -File Get-AntivirusState.ps1
      powershell -File Get-AntivirusState.ps1 -AsObject
#>
[CmdletBinding()]
param(
    [switch]$AsObject
)

$ErrorActionPreference = 'Continue'

function Get-RealTimeState([int]$productState) {
    # Bits 8-15 carry the real-time (scanner) state; the project documents
    # 0x41000 = on and 0x42000 = paused, i.e. the state byte is bits 8-15 with
    # 0x10 = on and 0x20 = paused. Anything else is reported with its raw byte so
    # a different engine is read correctly instead of guessed at.
    $byte = ($productState -shr 8) -band 0xFF
    switch ($byte) {
        0x01 { return 'on' }
        0x10 { return 'on' }
        0x11 { return 'on' }
        0x20 { return 'paused' }
        0x21 { return 'paused' }
        default { return ('unknown (0x{0:X2})' -f $byte) }
    }
}

$products = @()
try {
    $found = Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntivirusProduct -ErrorAction Stop
    foreach ($p in $found) {
        $products += [pscustomobject]@{
            Name      = $p.displayName
            State     = (Get-RealTimeState ([int]$p.productState))
            Raw       = ('0x{0:X}' -f $p.productState)
        }
    }
} catch {
    # SecurityCenter2 is not available on a server SKU or a hardened image.
    $products += [pscustomobject]@{ Name = '(security center unavailable)'; State = 'unknown'; Raw = '' }
}

# A third-party engine with self-defense cannot be paused by any script; only
# report it. The state that matters is the one printed here.
if ($AsObject) { return $products }

$summary = ($products | ForEach-Object { "$($_.Name)=$($_.State)" }) -join '; '
Write-Output ("Antivirus: " + $summary)
