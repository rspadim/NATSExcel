#Requires -Version 5.1
<#
    Run-Tests.ps1 - the fast, single-command local test run for NATSExcel.

    It builds the solution ONCE and then runs the independent test layers in
    PARALLEL off that one build (--no-build --no-restore), so no layer pays for
    a second restore/compile and a slow layer no longer blocks the others.

    Layers and what keeps them from colliding:
      unit      - no server; runs off the same build output.
      smoke     - its OWN nats-server on -SmokePort (default 4322) in its OWN
                  temp work dir, started with -Force; --require-jetstream.
      liveness  - its OWN nats-server on -LivenessPort (default 4422) in its OWN
                  temp work dir; --quick --skip-restart by default (the CI set),
                  -LivenessFull drops both flags for the destructive run.
      e2e       - NOT started unless -WithE2E; needs Excel; its OWN server on
                  -E2EPort (default 4222) and the script starts a second endpoint
                  on -E2EPort + 10 for the multi-cluster journey.

    The layers keep their own semantics - this is a convenience wrapper, it never
    skips or weakens an assertion. Every layer is re-runnable: each cluster is
    started with -Force and each work dir is private to its layer.

    Concurrency note: the layers share one machine, so running them together adds
    CPU/ThreadPool pressure on top of each layer's own. That is how CI behaves,
    so a healthy run is a real result; the liveness delivery bounds are the most
    load-sensitive, and that layer is safe to re-run on its own (or run the
    layers sequentially with -Sequential) if the machine is otherwise busy.

    Usage:
      powershell -ExecutionPolicy Bypass -File test\TestSupport\Run-Tests.ps1
      ... -NoBuild                 # reuse the current build (fastest re-run)
      ... -WithE2E                 # add the Excel end-to-end layer
      ... -SkipSmoke -SkipLiveness # unit only
      ... -LivenessFull            # the full destructive liveness run
      ... -ServerExe C:\nats\nats-server.exe
      ... -KeepCluster             # leave the test servers up (debugging)
#>
[CmdletBinding()]
param(
    [string]$RepoRoot,
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    # Reuse the last build instead of building the solution again.
    [switch]$NoBuild,
    # Add the Excel end-to-end layer (requires Excel and a built XLL).
    [switch]$WithE2E,
    [switch]$SkipSmoke,
    [switch]$SkipLiveness,
    # Run liveness WITHOUT --quick --skip-restart (the full, destructive run).
    [switch]$LivenessFull,
    # Add the load layer (a throughput benchmark, not a correctness gate, so it is
    # opt-in here; the CI always runs it). -LoadSeconds keeps it short.
    [switch]$WithLoad,
    [int]$LoadSeconds = 5,
    # Explicit nats-server.exe. Otherwise NATSEXCEL_NATS_SERVER, PATH and the
    # shared %TEMP%\natsexcel-cluster server are tried (Start-Cluster's own
    # fallback downloads once if none is found).
    [string]$ServerExe,
    [int]$SmokePort = 4322,
    [int]$LivenessPort = 4422,
    [int]$E2EPort = 4222,
    [int]$LoadPort = 4522,
    [string]$SubjectPrefix = 'test.natsexcel',
    # Run the layers one after another instead of in parallel (slowest, but the
    # least machine pressure for a flaky machine or a debug run).
    [switch]$Sequential,
    # Leave the test servers running when the run ends.
    [switch]$KeepCluster
)

$ErrorActionPreference = 'Stop'

# This script lives in <repo>\test\TestSupport\; the repo root is two levels up
# (never $MyInvocation - it points at the caller when invoked by path).
if (-not $RepoRoot) {
    $here = Split-Path -Parent $PSCommandPath
    $RepoRoot = (Resolve-Path (Join-Path $here '..\..')).Path
}
$RepoRoot = (Resolve-Path $RepoRoot).Path
Set-Location $RepoRoot
$script:RepoRoot = $RepoRoot

# Full paths: Start-Process does not always resolve a bare name from PATH.
$script:DotNet = (Get-Command dotnet -ErrorAction Stop).Source
$script:PowerShell = (Get-Command powershell -ErrorAction Stop).Source
$script:StartCluster = Join-Path $RepoRoot 'test\TestSupport\Start-Cluster.ps1'
$script:StopCluster = Join-Path $RepoRoot 'test\TestSupport\Stop-Cluster.ps1'
$script:TestRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-tests'

function Write-Section([string]$m) {
    Write-Host ''
    Write-Host ("=== " + $m + " ===") -ForegroundColor Cyan
}
function Write-Step([string]$m) { Write-Host ("[run-tests] " + $m) }

# ---------------------------------------------------------------------------
# Locate a nats-server once, so the parallel layers do not each download it into
# their own work dir (Start-Cluster's fallback would).
# ---------------------------------------------------------------------------
function Resolve-NatsServer {
    param([string]$Explicit)
    if ($Explicit) {
        if (-not (Test-Path -LiteralPath $Explicit)) { throw "nats-server not found: $Explicit" }
        return (Resolve-Path -LiteralPath $Explicit).Path
    }
    if ($env:NATSEXCEL_NATS_SERVER -and (Test-Path -LiteralPath $env:NATSEXCEL_NATS_SERVER)) {
        return (Resolve-Path -LiteralPath $env:NATSEXCEL_NATS_SERVER).Path
    }
    $onPath = Get-Command nats-server -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }
    # The shared server Start-Cluster downloads by default; reuse it if present.
    $shared = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-cluster\server\nats-server.exe'
    if (Test-Path -LiteralPath $shared) { return $shared }
    return $null
}
$server = Resolve-NatsServer -Explicit $ServerExe
if ($server) { Write-Step "nats-server: $server" }
else { Write-Step 'nats-server: not found locally; Start-Cluster.ps1 will fetch one' }

# ---------------------------------------------------------------------------
# Job control: launch a layer, poll until every layer has exited, report each
# one's own wall time.
# ---------------------------------------------------------------------------
$script:running = [ordered]@{}

function Start-Layer {
    param(
        [string]$Name,
        [string]$Exe,
        [string[]]$Arguments,
        [string]$LogDir
    )
    New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
    $outLog = Join-Path $LogDir "$Name.out.log"
    $errLog = Join-Path $LogDir "$Name.err.log"
    $watch = [System.Diagnostics.Stopwatch]::StartNew()

    # System.Diagnostics.Process, not Start-Process -PassThru: with redirection
    # and no -Wait, Start-Process leaves Process.ExitCode $null, so a passing
    # layer would look like a failure. Reading the two streams async keeps the
    # child from blocking on a full pipe while the layers run in parallel.
    # Full path resolved at startup; a bare name is not always found by the
    # child process launch.
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.WorkingDirectory = $script:RepoRoot
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    # Quote each argument; ArgumentList does not exist on net48-era .NET but
    # this host is modern PowerShell, so join into a single command line.
    $psi.Arguments = ($Arguments | ForEach-Object { '"' + ($_ -replace '"', '\"') + '"' }) -join ' '

    $proc = [System.Diagnostics.Process]::Start($psi)
    $script:running[$Name] = [pscustomobject]@{
        Name    = $Name
        Process = $proc
        Watch   = $watch
        OutLog  = $outLog
        ErrLog  = $errLog
        OutTask = $proc.StandardOutput.ReadToEndAsync()
        ErrTask = $proc.StandardError.ReadToEndAsync()
        ExitCode = $null
        Seconds  = 0.0
    }
    Write-Step "started $Name (pid $($proc.Id))"
}

function Wait-AllLayers {
    while ($true) {
        $pending = $false
        foreach ($layer in @($script:running.Values)) {
            if ($null -ne $layer.ExitCode) { continue }
            if ($layer.Process.HasExited) {
                Complete-Layer -Layer $layer
            } else {
                $pending = $true
            }
        }
        if (-not $pending) { break }
        Start-Sleep -Milliseconds 200
    }
}

# -Sequential runs one layer at a time: wait for the just-started layer before
# launching the next one.
function Wait-Layer {
    param([string]$Name)
    $layer = $script:running[$Name]
    $layer.Process.WaitForExit()
    Complete-Layer -Layer $layer
}

function Complete-Layer {
    param($Layer)
    if ($null -ne $Layer.ExitCode) { return }
    $Layer.Watch.Stop()
    $Layer.Seconds = [math]::Round($Layer.Watch.Elapsed.TotalSeconds, 1)
    $Layer.ExitCode = $Layer.Process.ExitCode
    # Drain both pipes to disk for the tail/diagnosis.
    $out = try { $Layer.OutTask.GetAwaiter().GetResult() } catch { '' }
    $err = try { $Layer.ErrTask.GetAwaiter().GetResult() } catch { '' }
    [System.IO.File]::WriteAllText($Layer.OutLog, "$out", (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText($Layer.ErrLog, "$err", (New-Object System.Text.UTF8Encoding($false)))
    $verdict = if ($Layer.ExitCode -eq 0) { 'PASS' } else { 'FAIL' }
    $color = if ($Layer.ExitCode -eq 0) { 'Green' } else { 'Red' }
    Write-Host ("[run-tests] {0} {1} in {2}s (exit {3})" -f $Layer.Name, $verdict, $Layer.Seconds, $Layer.ExitCode) -ForegroundColor $color
}

# A layer needs its own port. If the preferred one is already taken (a foreign
# or leaked server from another session), shift up to the next free port rather
# than failing or silently sharing someone else's server - a shared server would
# break the "own port and own work dir per layer" guarantee. Pass -SmokePort /
# -LivenessPort / -E2EPort to pick a different starting point.
function Test-PortFree {
    param([int]$Port)
    return -not (Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1)
}
function Resolve-LayerPort {
    param([string]$Layer, [int]$Preferred, [switch]$AlsoNextFree)
    for ($port = $Preferred; $port -lt ($Preferred + 50); $port++) {
        $required = @($port)
        # The E2E script starts its own second endpoint on <port>+10.
        if ($AlsoNextFree) { $required += ($port + 10) }
        if (($required | Where-Object { -not (Test-PortFree $_) }).Count -eq 0) {
            if ($port -ne $Preferred) { Write-Step "$Layer port $Preferred is busy; using $port" }
            return $port
        }
    }
    throw "no free port for $Layer near $Preferred"
}

function Show-LayerTail {
    param($Layer, [int]$Lines = 15, [switch]$Force)
    if ($Layer.ExitCode -eq 0 -and -not $Force) { return }
    Write-Host ("--- {0} (last {1} lines) ---" -f $Layer.Name, $Lines) -ForegroundColor Yellow
    foreach ($f in @($Layer.OutLog, $Layer.ErrLog)) {
        if (Test-Path -LiteralPath $f) {
            Get-Content -LiteralPath $f -Tail $Lines | ForEach-Object { Write-Host ("  " + $_) }
        }
    }
}

# ---------------------------------------------------------------------------
# Cluster helpers (always -Nodes 1, private work dir, -Force so a stale node
# from a previous run is replaced rather than collided with).
# ---------------------------------------------------------------------------
function Start-LayerCluster {
    param([string]$Layer, [int]$Port)
    $workDir = Join-Path $script:TestRoot $Layer
    New-Item -ItemType Directory -Force -Path $workDir | Out-Null
    # Track the node PIDs by diffing the nats-server set around the launch.
    # Start-Cluster.ps1's own -Force cannot attribute a server on this machine
    # (Get-Process.Path is empty for nats-server), so the runner records exactly
    # the processes IT started and reaps those on the way out. Never kills a
    # server it did not start.
    $before = Get-NatsServerIds
    $startArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:StartCluster,
        '-Nodes', '1', '-BasePort', "$Port", '-WorkDir', $workDir, '-Force')
    if ($server) { $startArgs += @('-ServerExe', $server) }
    $out = & $script:PowerShell @startArgs
    $url = "$($out | Select-Object -Last 1)".Trim()
    if (-not $url -or $url -notmatch '^nats://') { $url = "nats://127.0.0.1:$Port" }
    $pids = @(Get-NatsServerIds | Where-Object { $before -notcontains $_ })
    Write-Step "$Layer cluster: $url (node pid(s) $($pids -join ', '))"
    return [pscustomobject]@{ Url = $url; Port = $Port; WorkDir = $workDir; Pids = $pids }
}

# The set of running nats-server PIDs (identity by PID; Path/CommandLine are not
# readable for this process on this host).
function Get-NatsServerIds {
    return @(Get-Process -Name 'nats-server' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

function Stop-LayerCluster {
    param($Cluster)
    if ($KeepCluster) { Write-Step "leaving the cluster on $($Cluster.Port) up (-KeepCluster)"; return }
    # Kill the exact nodes this run started (a no-op if they already exited).
    foreach ($nodeId in @($Cluster.Pids)) {
        Stop-Process -Id $nodeId -Force -ErrorAction SilentlyContinue
    }
    & $script:PowerShell -NoProfile -ExecutionPolicy Bypass -File $script:StopCluster `
        -Nodes 1 -BasePort "$($Cluster.Port)" -WorkDir $Cluster.WorkDir -Remove | Out-Null
}

# ---------------------------------------------------------------------------
# 1. Build once (the single restore/compile of the whole run).
# ---------------------------------------------------------------------------
$totalWatch = [System.Diagnostics.Stopwatch]::StartNew()
$buildSeconds = 0.0
if ($NoBuild) {
    Write-Section "Build ($Configuration) skipped (-NoBuild)"
} else {
    Write-Section "Build the solution ($Configuration)"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    & dotnet build NATSExcel.sln -c $Configuration "--nologo"
    $buildExit = $LASTEXITCODE
    $sw.Stop()
    $buildSeconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
    if ($buildExit -ne 0) { throw "the solution build failed (exit $buildExit) after ${buildSeconds}s" }
    Write-Step "solution built in ${buildSeconds}s"
}

# ---------------------------------------------------------------------------
# 2. Bring up each layer's private server BEFORE the parallel block, so the
#    per-layer wall time measures the layer, not its server's startup.
# ---------------------------------------------------------------------------
$clusters = @{}
if (-not $SkipSmoke) {
    $SmokePort = Resolve-LayerPort -Layer 'smoke' -Preferred $SmokePort
    $clusters.smoke = Start-LayerCluster -Layer 'smoke' -Port $SmokePort
}
if (-not $SkipLiveness) {
    $LivenessPort = Resolve-LayerPort -Layer 'liveness' -Preferred $LivenessPort
    $clusters.liveness = Start-LayerCluster -Layer 'liveness' -Port $LivenessPort
}
if ($WithE2E) {
    $E2EPort = Resolve-LayerPort -Layer 'e2e' -Preferred $E2EPort -AlsoNextFree
    $clusters.e2e = Start-LayerCluster -Layer 'e2e' -Port $E2EPort
}
if ($WithLoad) {
    $LoadPort = Resolve-LayerPort -Layer 'load' -Preferred $LoadPort
    $clusters.load = Start-LayerCluster -Layer 'load' -Port $LoadPort
}

# ---------------------------------------------------------------------------
# 3. Run the independent layers: IN PARALLEL by default (each on its own port
#    and work dir), or one at a time with -Sequential.
# ---------------------------------------------------------------------------
Write-Section $(if ($Sequential) { 'Run the layers sequentially' } else { 'Run the layers in parallel' })
$logsDir = Join-Path $script:TestRoot 'logs'

function Start-Unit {
    Start-Layer -Name 'unit' -Exe $script:DotNet -LogDir $logsDir -Arguments @(
        'test', 'test\NATSExcel.Tests\NATSExcel.Tests.csproj', '-c', $Configuration,
        '--no-build', '--no-restore')
}
function Start-Smoke {
    Start-Layer -Name 'smoke' -Exe $script:DotNet -LogDir $logsDir -Arguments @(
        'run', '--project', 'test\SmokeTests', '-c', $Configuration,
        '--no-build', '--no-restore', '--',
        $clusters.smoke.Url, '--require-jetstream')
}
function Start-Liveness {
    $livenessArgs = @(
        'run', '--project', 'test\LivenessTests', '-c', $Configuration,
        '--no-build', '--no-restore', '--',
        $clusters.liveness.Url)
    if (-not $LivenessFull) { $livenessArgs += @('--quick', '--skip-restart') }
    Start-Layer -Name 'liveness' -Exe $script:DotNet -LogDir $logsDir -Arguments $livenessArgs
}
function Start-E2E {
    Start-Layer -Name 'e2e' -Exe $script:PowerShell -LogDir $logsDir -Arguments @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
        (Join-Path $RepoRoot 'test\Run-ExcelE2E.ps1'),
        '-RepoRoot', $RepoRoot, '-NatsUrl', $clusters.e2e.Url,
        '-SubjectPrefix', $SubjectPrefix)
}

function Start-Load {
    # `manager` exercises the subscription broadcast path; `raw` is the plain
    # NATS.Client baseline. A short run keeps the runner quick: the numbers are
    # informational, and the layer fails only on a crash.
    Start-Layer -Name 'load' -Exe $script:DotNet -LogDir $logsDir -Arguments @(
        'run', '--project', 'test\LoadTests', '-c', $Configuration,
        '--no-build', '--no-restore', '--',
        'manager', $clusters.load.Url, "$LoadSeconds", '2', '1')
}

if ($Sequential) {
    # No -Wait on the first layer's launch (it would block before the others
    # start); wait on each one before starting the next.
    Start-Unit; Wait-Layer 'unit'
    if (-not $SkipSmoke) { Start-Smoke; Wait-Layer 'smoke' }
    if (-not $SkipLiveness) { Start-Liveness; Wait-Layer 'liveness' }
    if ($WithE2E) { Start-E2E; Wait-Layer 'e2e' }
    if ($WithLoad) { Start-Load; Wait-Layer 'load' }
} else {
    Start-Unit
    if (-not $SkipSmoke) { Start-Smoke }
    if (-not $SkipLiveness) { Start-Liveness }
    if ($WithE2E) { Start-E2E }
    if ($WithLoad) { Start-Load }
    Wait-AllLayers
}

# ---------------------------------------------------------------------------
# 4. Summary + cleanup.
# ---------------------------------------------------------------------------
Write-Section 'Summary'
$failed = @()
$header = '{0,-10} {1,-8} {2,8}' -f 'layer', 'result', 'wall(s)'
Write-Host $header
Write-Host ('-' * $header.Length)
foreach ($layer in @($script:running.Values)) {
    $verdict = if ($layer.ExitCode -eq 0) { 'PASS' } else { 'FAIL' }
    if ($layer.ExitCode -ne 0) { $failed += $layer.Name }
    Write-Host ('{0,-10} {1,-8} {2,8}' -f $layer.Name, $verdict, $layer.Seconds)
}
$totalWatch.Stop()
$totalSeconds = [math]::Round($totalWatch.Elapsed.TotalSeconds, 1)
Write-Host ('-' * $header.Length)
if ($NoBuild) { Write-Host ('{0,-10} {1,-8} {2,8}' -f 'build', 'SKIP', 0) }
else { Write-Host ('{0,-10} {1,-8} {2,8}' -f 'build', 'PASS', $buildSeconds) }
Write-Host ('{0,-10} {1,-8} {2,8}' -f 'TOTAL', '', $totalSeconds)
Write-Host "logs: $logsDir"

foreach ($layer in @($script:running.Values)) { Show-LayerTail -Layer $layer }

foreach ($c in @($clusters.Values)) { Stop-LayerCluster -Cluster $c }

if ($failed.Count -gt 0) {
    Write-Host ("FAILED layer(s): " + ($failed -join ', ')) -ForegroundColor Red
    exit 1
}
$ran = @($script:running.Values | ForEach-Object { $_.Name })
$notRun = @()
if (-not $WithE2E) { $notRun += 'e2e (needs Excel; -WithE2E)' }
if (-not $WithLoad) { $notRun += 'load (a benchmark, not a gate; -WithLoad)' }
Write-Host ("ALL STARTED LAYERS PASSED (ran: " + ($ran -join ', ') + ")") -ForegroundColor Green
Write-Host ("not run: " + ($notRun -join '; '))
exit 0
