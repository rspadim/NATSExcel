#Requires -Version 5.1
<#
    Set-Antivirus.ps1 - toggles Windows Defender real-time protection so an Excel
    E2E run can be compared with antivirus ON and OFF (PLAN V8.7, the AV matrix).

    Excel COM automation is inherently flaky and security software widens the
    window (see AGENTS.md "Excel automation lessons"): the same script can pass
    with AV off and fail with AV on. This helper exists so a matrix run can
    attribute a failure to the environment instead of the add-in. It reports the
    state BEFORE and AFTER so the change is auditable.

    It ONLY touches Windows Defender's real-time monitoring (Set-MpPreference).
    A third-party engine with self-defense cannot be toggled by any script; this
    helper refuses to pretend otherwise and reports what it found.

    Elevation is required. Nothing here runs as part of a normal test; it is a
    manual, deliberate step.

    Usage:
      powershell -File Set-Antivirus.ps1 -State off
      powershell -File Set-Antivirus.ps1 -State on
      powershell -File Set-Antivirus.ps1 -State query     # report only
#>
[CmdletBinding()]
param(
    [ValidateSet('on', 'off', 'query')]
    [string]$State = 'query'
)

$ErrorActionPreference = 'Stop'

function Show-State([string]$when) {
    try {
        $pref = Get-MpPreference -ErrorAction Stop
        $disabled = [bool]$pref.DisableRealtimeMonitoring
        Write-Output ("[$when] Windows Defender real-time monitoring: " + $(if ($disabled) { 'OFF' } else { 'ON' }))
        return -not $disabled
    } catch {
        Write-Output ("[$when] Windows Defender not available or not readable: " + $_.Exception.Message)
        return $null
    }
}

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Not elevated: Windows Defender preferences cannot be changed without administrator rights."
    if ($State -ne 'query') { exit 2 }
}

$before = Show-State 'before'

switch ($State) {
    'query' {
        # Nothing to change; the banner above is the answer.
    }
    'off' {
        try {
            Set-MpPreference -DisableRealtimeMonitoring $true
            Write-Output "Real-time monitoring disabled (remember to re-enable with -State on)."
        } catch {
            Write-Warning ("Could not disable real-time monitoring: " + $_.Exception.Message)
        }
    }
    'on' {
        try {
            Set-MpPreference -DisableRealtimeMonitoring $false
            Write-Output "Real-time monitoring enabled."
        } catch {
            Write-Warning ("Could not enable real-time monitoring: " + $_.Exception.Message)
        }
    }
}

$after = Show-State 'after'

# Exit non-zero only when a requested change did not take effect, so a caller can
# treat it as a failed matrix setup rather than reading a stale state.
if ($State -eq 'off' -and $after -eq $true) { exit 1 }
if ($State -eq 'on' -and $after -eq $false) { exit 1 }
exit 0
