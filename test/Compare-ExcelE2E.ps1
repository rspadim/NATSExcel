#Requires -Version 5.1
<#
    Compare-ExcelE2E.ps1 - compares two Excel E2E matrix runs
    (test\Run-ExcelE2E-Matrix.ps1 -Tag ...) side by side.

    Read-only reporter: it parses the per-mode logs and prints a table. It does
    NOT run Excel and touches nothing on disk beyond reading the logs, so it is
    safe to run any time.

    It reads %TEMP%\natsexcel-e2e-<tag>-<mode>.log for the four modes and shows,
    per mode: the run status (PASS / FAIL xN / incomplete), the COM-retry total
    (a load/health signal, not an antivirus fingerprint - Excel COM is
    inherently flaky) and the SKIP count. Use the same script for both tags so
    the only variable is the environment.

    Examples:
      powershell -ExecutionPolicy Bypass -File test\Compare-ExcelE2E.ps1 -TagA before -TagB after
      powershell -ExecutionPolicy Bypass -File test\Compare-ExcelE2E.ps1 -TagA wip -TagB base -Modes default,sync
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$TagA,
    [Parameter(Mandatory = $true)][string]$TagB,
    # A comma-separated subset of: default, sync, ffall, async (default all).
    [string[]]$Modes
)

$ErrorActionPreference = 'Stop'

$allModes = @('default', 'sync', 'ffall', 'async')
# -File passes arguments as raw strings, so "default,sync" arrives as one
# element; split any commas so both "-Modes default,sync" and "-Modes default sync" work.
$Modes = @($Modes | ForEach-Object { $_ -split ',' } | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() })
if (-not $Modes -or $Modes.Count -eq 0) { $Modes = $allModes }
foreach ($m in $Modes) {
    if ($allModes -notcontains $m) {
        throw "unknown mode '$m' (expected: $($allModes -join ', '))"
    }
}

# Reads one tag's per-mode logs into an ordered map. A missing log is reported
# as "(absent)" rather than skipped, so an incomplete tag set is obvious.
# (The local map is named $map, NOT $modes: PowerShell variables are
# case-insensitive, so a local $modes would clobber the $Modes script parameter.)
function Read-MatrixOne([string]$tag) {
    $map = [ordered]@{}
    foreach ($mode in $Modes) {
        $log = Join-Path ([System.IO.Path]::GetTempPath()) ("natsexcel-e2e-{0}-{1}.log" -f $tag, $mode)
        if (-not (Test-Path -LiteralPath $log)) {
            $map[$mode] = [pscustomobject]@{ Status = '(absent)'; Fails = -1; Skips = -1; Retries = -1 }
            continue
        }
        $text = Get-Content -LiteralPath $log -Raw
        $fails = -1; $skips = -1; $retries = -1
        if ($text -match 'E2E DONE: fails=(\d+), skips=(\d+), comRetries=(\d+)') {
            $fails = [int]$Matches[1]; $skips = [int]$Matches[2]; $retries = [int]$Matches[3]
        }
        $status = if ($fails -eq 0) { 'PASS' }
                  elseif ($fails -gt 0) { "FAIL x$fails" }
                  else { 'incomplete' }
        $map[$mode] = [pscustomobject]@{ Status = $status; Fails = $fails; Skips = $skips; Retries = $retries }
    }
    # -NoEnumerate: an OrderedDictionary is IEnumerable, so a bare `return $map`
    # would splat it into an array of DictionaryEntry and break keyed lookup.
    return , $map
}

$a = Read-MatrixOne $TagA
$b = Read-MatrixOne $TagB

$a = Read-MatrixOne $TagA
$b = Read-MatrixOne $TagB

Write-Host ("{0,-8} | {1,-16} {2,-8} {3,-6} | {4,-16} {5,-8} {6,-6}" -f `
    'mode', "$TagA status", 'retries', 'skip', "$TagB status", 'retries', 'skip')
Write-Host ('-' * 82)
foreach ($mode in $Modes) {
    $ra = $a[$mode]; $rb = $b[$mode]
    Write-Host ("{0,-8} | {1,-16} {2,-8} {3,-6} | {4,-16} {5,-8} {6,-6}" -f `
        $mode,
        $ra.Status, $ra.Retries, $ra.Skips,
        $rb.Status, $rb.Retries, $rb.Skips)
}
Write-Host ''
Write-Host ("Tag A logs: %TEMP%\natsexcel-e2e-{0}-<mode>.log" -f $TagA) -ForegroundColor DarkGray
Write-Host ("Tag B logs: %TEMP%\natsexcel-e2e-{0}-<mode>.log" -f $TagB) -ForegroundColor DarkGray
Write-Host 'COM retries is a load/health signal (Excel COM is inherently flaky; the count is similar with the AV active or paused), not an antivirus indicator. Status/SKIP differences between the tags are the environment''s observable impact.'
