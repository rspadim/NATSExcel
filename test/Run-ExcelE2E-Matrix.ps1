#Requires -Version 5.1
<#
    Run-ExcelE2E-Matrix.ps1 - runs the Excel E2E once per write mode and prints
    a one-line result per mode.

    The write mode is a CONFIGURATION choice (NATSExcel.json SyncWrite /
    AsyncWrites), not a code path a parameter can flip from outside Excel: each
    run is a fresh hidden Excel whose throwaway add-in folder carries a config
    with that mode. So this driver passes -SyncWrite / -AsyncWrites (and a
    -LogPath) to Run-ExcelE2E.ps1, which rewrites ONLY the copy of the config,
    never the shipped sample.

    Each mode's full log is %TEMP%\natsexcel-e2e-<tag>-<mode>.log;
    test\Compare-ExcelE2E.ps1 diffs two tags (e.g. before/after an AV pause).

    Modes: default (as shipped), sync (SyncWrite=sync), ffall
    (fireforget-all), async (AsyncWrites=true).

    It is not part of CI (it needs Excel). It never leaves an Excel process
    behind: the hidden instance each run uses is quit, and any empty-title
    automation instance is cleaned up between modes. Nothing needs a click.

    NOTE: the destructive node kill is SKIPPED by default (-SkipNodeKill is
    passed to each run) so a 3-node test cluster survives all four modes. Pass
    -IncludeNodeKill to run it (only against a disposable cluster: it kills the
    node the URL points at, so later modes on the same URL would fail).

    Examples:
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E-Matrix.ps1
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E-Matrix.ps1 -Tag wip -Fast
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E-Matrix.ps1 -Modes default,sync
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E-Matrix.ps1 -NatsUrl "nats://127.0.0.1:4222"
#>
[CmdletBinding()]
param(
    [string]$Tag = 'run',
    [string]$RepoRoot,
    [string]$NatsUrl = 'nats://127.0.0.1:4222',
    [string]$SubjectPrefix = 'test.natsexcel',
    # A comma-separated subset of: default, sync, ffall, async.
    [string[]]$Modes,
    # Shortens the per-assertion bound, like Run-ExcelE2E.ps1 -Fast.
    [switch]$Fast,
    # Run the destructive node kill in every mode (disposable cluster only).
    [switch]$IncludeNodeKill,
    # Run only the original checks in each mode.
    [switch]$SkipExtraChecks
)

$ErrorActionPreference = 'Stop'

if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath) }
$RepoRoot = (Resolve-Path $RepoRoot).Path

$e2e = Join-Path $RepoRoot 'test\Run-ExcelE2E.ps1'
if (-not (Test-Path -LiteralPath $e2e)) { throw "E2E script not found: $e2e" }

# name -> extra args handed to Run-ExcelE2E.ps1. "default" carries none, so the
# starter config is left exactly as the copy helper wrote it.
$allModes = [ordered]@{
    default = @()
    sync    = @('-SyncWrite', 'sync')
    ffall   = @('-SyncWrite', 'fireforget-all')
    async   = @('-AsyncWrites')
}

if (-not $Modes -or $Modes.Count -eq 0) { $Modes = @($allModes.Keys) }
# -File passes arguments as raw strings, so "default,sync" arrives as one
# element; split any commas so both "-Modes default,sync" and "-Modes default sync" work.
$Modes = @($Modes | ForEach-Object { $_ -split ',' } | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() })
if (-not $Modes -or $Modes.Count -eq 0) { $Modes = @($allModes.Keys) }
foreach ($m in $Modes) {
    if (-not $allModes.Contains($m)) {
        throw "unknown mode '$m' (expected: $(($allModes.Keys) -join ', '))"
    }
}

# Kills only a hidden automation Excel (no window title) - never the user's own.
function Stop-HiddenExcel {
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) } |
        ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
}

# One line per mode, parsed from that mode's log. The E2E ends with
# "=== E2E DONE: fails=N, skips=N, comRetries=N ==="; a FAIL line is "<name> FAIL".
function Get-RunSummary([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) {
        return [pscustomobject]@{ Status = 'no log'; Fails = -1; Skips = -1; Retries = -1 }
    }
    $text = Get-Content -LiteralPath $path -Raw
    $fails = -1; $skips = -1; $retries = -1
    if ($text -match 'E2E DONE: fails=(\d+), skips=(\d+), comRetries=(\d+)') {
        $fails = [int]$Matches[1]; $skips = [int]$Matches[2]; $retries = [int]$Matches[3]
    }
    $status = if ($fails -eq 0) { 'PASS' }
              elseif ($fails -gt 0) { "FAIL x$fails" }
              else { 'incomplete' }
    return [pscustomobject]@{ Status = $status; Fails = $fails; Skips = $skips; Retries = $retries }
}

$results = [ordered]@{}
Write-Host ("=== NATSExcel E2E matrix [{0}] ===" -f $Tag) -ForegroundColor Cyan
Write-Host ("repo: {0}" -f $RepoRoot)
Write-Host ("nats: {0}  (node kill {1})" -f $NatsUrl, $(if ($IncludeNodeKill) { 'ON' } else { 'skipped' }))
Write-Host (powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot 'test\TestSupport\Get-AntivirusState.ps1'))

foreach ($mode in $Modes) {
    $log = Join-Path ([System.IO.Path]::GetTempPath()) ("natsexcel-e2e-{0}-{1}.log" -f $Tag, $mode)
    Write-Host ("") -NoNewline
    Write-Host ("=== MODE {0} [{1}] ===" -f $mode, $Tag) -ForegroundColor Cyan

    $runArgs = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $e2e,
        '-RepoRoot', $RepoRoot,
        '-NatsUrl', $NatsUrl,
        '-SubjectPrefix', $SubjectPrefix,
        '-LogPath', $log
    )
    if ($Fast) { $runArgs += '-Fast' }
    if ($SkipExtraChecks) { $runArgs += '-SkipExtraChecks' }
    if (-not $IncludeNodeKill) { $runArgs += '-SkipNodeKill' }
    $runArgs += $allModes[$mode]

    # The E2E writes its own full log to -LogPath, so do NOT tee it here (that
    # would duplicate every line). Capture the output once and echo only the
    # assertion lines and the DONE summary, so a mode is a one-screen report.
    $out = & powershell @runArgs 2>&1
    $exit = $LASTEXITCODE
    $out | Select-String -Pattern '^.*\s(PASS|FAIL)$|^SKIP |^=== E2E DONE' |
        ForEach-Object { $_.Line.Trim() }

    $results[$mode] = Get-RunSummary $log
    $r = $results[$mode]
    $color = if ($r.Fails -eq 0) { 'Green' } else { 'Red' }
    Write-Host ("MODE {0,-8} : {1,-10} fails={2} skips={3} comRetries={4}  (exit={5})" -f `
        $mode, $r.Status, $r.Fails, $r.Skips, $r.Retries, $exit) -ForegroundColor $color
    Write-Host ("  log: {0}" -f $log) -ForegroundColor DarkGray

    # Quit any hidden instance this run left, then let Excel settle before the
    # next mode so the modes do not contend for the same COM server.
    Stop-HiddenExcel
    Start-Sleep -Seconds 3
}

Write-Host ("") -NoNewline
Write-Host ("=== DONE [{0}] ===" -f $Tag) -ForegroundColor Cyan
foreach ($mode in $Modes) {
    $r = $results[$mode]
    Write-Host ("  {0,-8} {1,-10} fails={2} skips={3} comRetries={4}" -f `
        $mode, $r.Status, $r.Fails, $r.Skips, $r.Retries)
}

$anyFail = @($results.Values | Where-Object { $_.Fails -ne 0 }).Count
if ($anyFail -gt 0) { exit 1 }
exit 0
