#Requires -Version 5.1
<#
    New-AddInFolder.ps1 - builds a throwaway add-in folder by COPYING the build
    output. No installer is involved, so no wizard, no COM registration, no
    autoload entry and nothing to click - a test can create and discard the
    folder as many times as it likes.

    Excel loads the XLL straight from the folder (RegisterXLL), so the same
    bits the release ships are the bits under test. The dll, NLog.config and the
    schema sit next to the XLL, exactly as the installer would lay them out.

    Usage:
      powershell -File New-AddInFolder.ps1 -Path C:\temp\myaddin
      # optionally write a config (JSON text or a path to one)
      powershell -File New-AddInFolder.ps1 -Path C:\temp\myaddin -Config "C:\temp\c.json"
      powershell -File New-AddInFolder.ps1 -Path C:\temp\myaddin -Url "nats://127.0.0.1:4222"

    It prints the XLL path on success and exits non-zero when the build output is
    missing (run `dotnet build NATSExcel.csproj -c Release` first).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$RepoRoot,
    # A NATS URL to put in the starter config (ignored when -Config is given).
    [string]$Url = 'nats://localhost:4222',
    # A JSON file to copy in as NATSExcel.json; or a raw JSON string.
    [string]$Config,
    # Keep an existing folder's NATSExcel.json (a test that wants a custom one).
    [switch]$KeepConfig,
    [ValidateSet('x64', 'x86')]
    [string]$Bitness = 'x64'
)

$ErrorActionPreference = 'Stop'

if (-not $RepoRoot) {
    # This script lives in <repo>\test\TestSupport\, so the repo root is two
    # levels up from its own folder. Do not use $MyInvocation: it points at the
    # CALLER when the script is invoked by path.
    $here = Split-Path -Parent $PSCommandPath
    $RepoRoot = (Resolve-Path (Join-Path $here '..\..')).Path
}
$RepoRoot = (Resolve-Path $RepoRoot).Path

$publish = Join-Path $RepoRoot 'bin\Release\net48\publish'
$xllName = if ($Bitness -eq 'x64') { 'NATSExcel64-packed.xll' } else { 'NATSExcel-packed.xll' }
$xllSource = Join-Path $publish $xllName
if (-not (Test-Path $xllSource)) {
    throw "build output missing: $xllSource`nrun: dotnet build NATSExcel.csproj -c Release"
}

New-Item -ItemType Directory -Force -Path $Path | Out-Null

# A crashed previous run can leave its HIDDEN Excel instance alive; that instance
# keeps the copied XLL loaded and every later run then dies on "being used by
# another process" (a message that names no culprit). The instance is provably
# this harness's - it is invisible AND it has OUR destination XLL loaded - so it
# is stopped here, and a VISIBLE Excel (the user's own session, which would not
# have our temp copy loaded anyway) is only reported, never killed.
function Clear-StaleAddInHost([string]$xllPath) {
    $stale = @()
    foreach ($proc in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        $loads = $false
        try {
            $loads = [bool]($proc.Modules | Where-Object {
                $_.FileName -and $_.FileName.Equals($xllPath, [System.StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1)
        }
        catch { $loads = $false }   # a process we cannot inspect is not ours to judge
        if ($loads) { $stale += $proc }
    }
    if (-not $stale.Count) { return }

    foreach ($proc in $stale) {
        if ($proc.MainWindowHandle -ne 0) {
            throw ("$xllPath is loaded by the VISIBLE Excel pid $($proc.Id); close that workbook " +
                   "and re-run (it is not this harness's instance, so it is never stopped here).")
        }
    }
    foreach ($proc in $stale) {
        # stderr, never stdout: this script's STDOUT IS ITS RETURN VALUE (the caller
        # runs `$Xll = powershell -File New-AddInFolder.ps1 ...`, so a host-write here
        # lands inside the returned path and every later Test-Path / RegisterXLL
        # fails with a string full of progress text). stderr is shown and not
        # captured.
        [Console]::Error.WriteLine("[addin] stopping the stale hidden Excel pid $($proc.Id) that still held $xllPath")
        try { Stop-Process -Id $proc.Id -Force -ErrorAction Stop } catch {
            throw ("could not stop the stale hidden Excel pid $($proc.Id): $($_.Exception.Message)")
        }
    }
    # Give Windows a moment to release the file lock before the copy below.
    for ($i = 0; $i -lt 20; $i++) {
        try {
            $fs = [System.IO.File]::Open($xllPath, 'Open', 'ReadWrite', 'None')
            $fs.Close()
            return
        }
        catch { Start-Sleep -Milliseconds 250 }
    }
    throw "$xllPath is still locked after stopping the stale instance(s)"
}

$destXll = Join-Path $Path $xllName
if (Test-Path $destXll) { Clear-StaleAddInHost $destXll }

# The add-in folder holds the XLL, the managed dll it loads, the log config and
# the schema. The other packed XLL is deliberately NOT copied, mirroring what
# the installer keeps for the detected bitness.
Copy-Item $xllSource $destXll -Force
foreach ($name in 'NATSExcel.dll', 'NLog.config', 'NATSExcel.schema.json') {
    $src = Join-Path $publish $name
    if (Test-Path $src) { Copy-Item $src (Join-Path $Path $name) -Force }
    else {
        $srcRepo = Join-Path $RepoRoot $name
        if (Test-Path $srcRepo) { Copy-Item $srcRepo (Join-Path $Path $name) -Force }
    }
}

# Starter config: a caller-provided one wins; otherwise write a minimal file that
# points at the requested URL, unless the folder already has one and -KeepConfig.
$configPath = Join-Path $Path 'NATSExcel.json'
$writeConfig = -not ($KeepConfig -and (Test-Path $configPath))
if ($writeConfig) {
    if ($Config) {
        if (Test-Path $Config) { Copy-Item $Config $configPath -Force }
        else {
            $utf8 = New-Object System.Text.UTF8Encoding($false)
            [System.IO.File]::WriteAllText($configPath, $Config, $utf8)
        }
    } else {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $json = @"
{
  // Written by test\TestSupport\New-AddInFolder.ps1 - a throwaway test config.
  "Default": "local",
  "Servers": { "dev": "$Url" },
  "Clusters": { "local": { "urls": ["dev"] } },
  "RTD": {
    "NatsUpdateRateMs": 1000,
    "ExcelUpdateRateMs": 100,
    "MessageCounterThreshold": 10000,
    "NatsUpdateStyle": "Automatic"
  },
  "JetStream": { "enabled": true, "defaultTimeoutMs": 5000 },
  "AllowAdminWrites": false,
  "SkipRepeatedMessages": true,
  "CoalesceRealtimeUpdates": true,
  "ConflationMs": 200,
  "PublishDedupCacheSize": 10000,
  "SyncWrite": "fireforget",
  "AsyncWrites": true,
  "UpdateCheck": false
}
"@
        [System.IO.File]::WriteAllText($configPath, $json, $utf8)
    }
}

Write-Output (Join-Path $Path $xllName)
