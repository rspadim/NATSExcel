#Requires -Version 5.1
<#
    Run-ExcelE2E.ps1 - the Excel end-to-end test for NATSExcel.

    Loads the packed XLL into a hidden Excel and exercises what a user actually
    does: UDF write/read, RTD push, JetStream/KV paths, cluster failover (kill
    the node the add-in is on), the multi-workbook case, an Excel restart, two
    Excel instances on one subject, reopening a saved workbook, the
    ...NonVolatile single-run rule and a bad-endpoint bounded error.

    On top of that, a second "user journeys" block (skip with -SkipJourneys)
    walks the things a real person does with a workspace: a multi-cluster
    NATSExcel.json (a cluster name, a server alias and a raw URL each reaching
    their own endpoint), switching Default without restarting Excel, deleting
    the config and reloading, replaying JetStream messages from cells, keeping a
    JSON range in Key/Value and expanding it back into the sheet, a sheet with
    20 live SUB cells, and the #N/A pending marker of an async write. The
    journeys that seed a stream/KV bucket from outside the sheet need the
    `nats` CLI; without it they print SKIP, never FAIL.

    Requires: Excel, a built XLL, and a running NATS server (start one with
    test\TestSupport\Start-Cluster.ps1). Local hosts only for the destructive
    steps; the node kill is skipped automatically for a non-local URL.

    Every wait is CONDITION-BASED (a helper polls until the cell changes or the
    timeout expires), so a healthy run finishes as fast as Excel allows and a
    failing one still fails at a known bound. No fixed "sleep 5s and hope".

    Hosts default to loopback. Never commit a real endpoint: pass it as -NatsUrl.

    -SyncWrite / -AsyncWrites override those keys in the throwaway starter config
    the copy helper writes (never the shipped sample), so the same script can be
    run under each write mode by test\Run-ExcelE2E-Matrix.ps1.

    Examples:
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1 -NatsUrl "nats://127.0.0.1:4222" -Fast
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1 -SyncWrite sync -Tag wip
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1 -SkipExtraChecks   # original set only
#>
[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$NatsUrl = 'nats://127.0.0.1:4222',
    [string]$SubjectPrefix = 'test.natsexcel',
    [string]$Xll,
    [switch]$SkipNodeKill,
    [switch]$SkipMultiWorkbook,
    [switch]$KeepExcelOpen,
    # Run only the original check set (skip the extra scenarios below).
    [switch]$SkipExtraChecks,
    # Skip the user-journey block (journeys 1-8 below) and run the rest.
    [switch]$SkipJourneys,
    # Observe only: skip every write (publishes, stream/consumer/KV/object
    # creation) and the destructive node-kill step, keeping the read paths. It is
    # forced ON for a non-local endpoint (see below): D6 in PLAN.md says a real
    # server is read-only by default.
    [switch]$ReadOnly,
    # Opt IN to writing to a non-local endpoint. Without it the script refuses to
    # start, because pointing this suite at a real feed would publish to it.
    [switch]$AllowRemoteWrites,
    # A second, INDEPENDENT endpoint (a different server, not another node of
    # the same cluster) for the multi-cluster journey. On a local host with none
    # given, the script starts a disposable one-node server on <port>+10 itself
    # and stops it at the end. A non-local URL must supply -NatsUrl2 or the
    # multi-cluster isolate check is SKIPPED.
    [string]$NatsUrl2,
    # Write-mode selection, applied to the starter config the copy helper
    # writes (see New-AddInFolder.ps1). Empty leaves the config as written.
    [ValidateSet('', 'sync', 'fireforget', 'fireforget-all')]
    [string]$SyncWrite = '',
    [switch]$AsyncWrites,
    # Names a run: the log becomes %TEMP%\natsexcel-e2e-<Tag>.log. The matrix
    # (Run-ExcelE2E-Matrix.ps1) passes -LogPath instead for <tag>-<mode>.log.
    [string]$Tag,
    [string]$LogPath,
    # Shortens the per-assertion bound (a smoke run on a known-good machine).
    [switch]$Fast
)

$ErrorActionPreference = 'Stop'

if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath) }
$RepoRoot = (Resolve-Path $RepoRoot).Path

# Prefer a COPY of the build output over an install: no COM registration, no
# autoload entry, no wizard to click. Pass -Xll to point at a real install.
# The config journeys (1-3) rewrite NATSExcel.json in the add-in folder, so they
# only run when THIS script built the throwaway copy (never a real install).
$script:builtCopy = $false
if (-not $Xll) {
    $script:builtCopy = $true
    $copyDir = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-e2e-addin'
    Remove-Item -Recurse -Force $copyDir -ErrorAction SilentlyContinue
    $Xll = powershell -NoProfile -ExecutionPolicy Bypass -File `
        (Join-Path $RepoRoot 'test\TestSupport\New-AddInFolder.ps1') `
        -Path $copyDir -Url $NatsUrl

    # Write-mode selection: the helper wrote the starter config, so override just
    # the keys this run needs (SyncWrite / AsyncWrites) in THAT file. The shipped
    # sample is never touched - only the throwaway copy. Real JSON with the keys
    # left as written when neither switch is given (no config change at all).
    if ($SyncWrite -or $AsyncWrites) {
        $cfgPath = Join-Path $copyDir 'NATSExcel.json'
        $cfg = Get-Content -LiteralPath $cfgPath -Raw
        if ($SyncWrite) {
            $cfg = [regex]::Replace($cfg, '"SyncWrite"\s*:\s*"[^"]*"', '"SyncWrite": "' + $SyncWrite + '"')
        }
        if ($AsyncWrites) {
            $cfg = [regex]::Replace($cfg, '"AsyncWrites"\s*:\s*(true|false)', '"AsyncWrites": true')
        }
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($cfgPath, $cfg, $utf8)
    }
}
$Xll = "$Xll".Trim()
# The add-in searches its own folder first, so NATSExcel.json next to the XLL is
# the one the journeys edit (the copy helper writes it there).
$script:addInDir = Split-Path -Parent $Xll

# Bound per assertion: enough for a slow COM call, short enough not to waste
# minutes. A -Fast run trusts the machine and uses a tight bound.
$BoundMs = if ($Fast) { 4000 } else { 12000 }
$PollMs = 250

# Log destination: -LogPath wins (the matrix names it <tag>-<mode>.log);
# otherwise a -Tag run is <tag>.log and a plain run keeps the historical name.
if (-not $LogPath) {
    $LogPath = if ($Tag) {
        Join-Path ([System.IO.Path]::GetTempPath()) ("natsexcel-e2e-{0}.log" -f $Tag)
    } else {
        Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-e2e.log'
    }
}
$log = $LogPath
Remove-Item $log -ErrorAction SilentlyContinue
function Log([string]$m) { Write-Host $m; Add-Content -Path $log -Value $m }

$script:fails = 0
$script:skips = 0
$script:comRetries = 0
function Assert([string]$name, $cond) {
    Log ("{0,-52} {1}" -f $name, $(if ($cond) { 'PASS' } else { 'FAIL' }))
    if (-not $cond) { $script:fails++ }
}

# A scenario that could not run prints a uniform "SKIP <name>" line so the
# matrix/compare helpers can count skips without scraping prose.
function Skip([string]$name) {
    Log ("SKIP " + $name)
    $script:skips++
}

# COM is inherently flaky (a busied Excel rejects a call); retry with backoff.
function Invoke-Com([scriptblock]$block, [int]$attempts = 4) {
    for ($i = 1; $i -le $attempts; $i++) {
        try { return & $block }
        catch {
            if ($i -eq $attempts) { throw }
            $script:comRetries++
            Start-Sleep -Milliseconds (100 * $i)
        }
    }
}

# Poll a worksheet cell until the predicate holds (or the bound expires) instead
# of sleeping a fixed time. Returns the last text seen either way.
function Wait-Cell {
    param(
        [Parameter(Mandatory)][object]$Worksheet,
        [Parameter(Mandatory)][string]$Address,
        [Parameter(Mandatory)][scriptblock]$Until,
        [int]$TimeoutMs = $BoundMs
    )
    $deadline = (Get-Date).AddMilliseconds($TimeoutMs)
    $text = ''
    while ((Get-Date) -lt $deadline) {
        $text = Invoke-Com { $Worksheet.Range($Address).Text }
        if (& $Until $text) { return $text }
        Start-Sleep -Milliseconds $PollMs
    }
    return $text
}

function New-HiddenExcel {
    $x = New-Object -ComObject Excel.Application
    $x.Visible = $false
    $x.DisplayAlerts = $false
    Invoke-Com { $x.RegisterXLL($Xll) } | Out-Null
    return $x
}
function Close-Excel($x) {
    try { foreach ($wb in @($x.Workbooks)) { Invoke-Com { $wb.Close($false) } } } catch {}
    try { $x.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}

# Writes the copy folder's config as real JSON without a BOM. The template holds
# the multi-cluster shape the journeys exercise: a Default name, two Servers
# aliases and two Clusters, each alias resolving to its own endpoint.
$script:configTemplate = @'
{
  "Default": "__DEFAULT__",
  "Servers": { "srvA": "__URL__", "srvB": "__URL2__" },
  "Clusters": { "clusterA": { "urls": ["srvA"] }, "clusterB": { "urls": ["srvB"] } },
  "Auth": { "user": null, "password": null, "token": null, "jwt": null, "nkeySeed": null, "credsFile": null, "tls": false, "tlsCert": null, "tlsCertPassword": null },
  "Reconnect": { "maxAttempts": 60, "waitMs": 2000, "jitterMs": 100 },
  "RTD": { "NatsUpdateRateMs": 1000, "ExcelUpdateRateMs": 100, "MessageCounterThreshold": 10000, "NatsUpdateStyle": "Automatic" },
  "JetStream": { "enabled": true, "defaultTimeoutMs": 5000 },
  "AllowAdminWrites": __ADMIN__,
  "SkipRepeatedMessages": true,
  "CoalesceRealtimeUpdates": true,
  "ConflationMs": 200,
  "PublishDedupCacheSize": 10000,
  "SyncWrite": "__SYNC__",
  "AsyncWrites": __ASYNC__,
  "UpdateCheck": false
}
'@
function Write-Cfg {
    param([string]$Path, [string]$Default, [string]$Url, [string]$Url2, [bool]$Admin = $false)
    $text = $script:configTemplate.
        Replace('__DEFAULT__', $Default).
        Replace('__URL__', $Url).
        Replace('__URL2__', $Url2).
        Replace('__ADMIN__', $(if ($Admin) { 'true' } else { 'false' })).
        Replace('__SYNC__', $(if ($SyncWrite) { $SyncWrite } else { 'fireforget' })).
        Replace('__ASYNC__', $(if ($AsyncWrites) { 'true' } else { 'false' }))
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $text, $utf8)
}

# Is the nats-server answering $port one of the DISPOSABLE test nodes this
# repository starts (the cluster harness under %TEMP%\natsexcel-cluster, or this
# script's own second server), or somebody else's server?
#
# This is a SAFETY check, not a convenience. The failover journey KILLS the node
# it finds on the port, and killing a server this run did not start is destructive
# to whoever owns it (AGENTS.md lesson 3). .Path cannot answer the question: a
# Start-Process'd nats-server exposes an EMPTY Path. The reliable evidence is the
# command line naming a nats.conf under OUR work root - and a server started by
# another user or from an elevated shell does not expose its command line at all,
# which is itself the signal that it is not ours.
function Test-OwnedTestServer([int]$port) {
    $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $conn) { return $false }

    $proc = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
    if (-not $proc -or $proc.ProcessName -notlike 'nats-server*') { return $false }

    $cmd = $null
    try { $cmd = (Get-CimInstance Win32_Process -Filter "ProcessId=$($proc.Id)" -ErrorAction Stop).CommandLine }
    catch { return $false }
    if ([string]::IsNullOrWhiteSpace($cmd)) { return $false }

    $roots = @(
        (Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-cluster'),
        (Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-e2e-server2')
    )
    foreach ($dir in $roots) {
        if ($cmd -like "*$dir*") { return $true }
    }
    return $false
}


$script:secondServerPid = $null

# Starts a disposable one-node NATS server on $Port (JetStream on) using the
# cluster harness's own pinned server binary, in its own work dir. Only a server
# this script started is stopped again; the pid is recorded for that.
function Start-SecondServer {
    param([int]$Port)
    $root = [System.IO.Path]::GetTempPath()
    $work = Join-Path $root 'natsexcel-e2e-server2'
    $exe = Join-Path $root 'natsexcel-cluster\server\nats-server.exe'
    if (-not (Test-Path $exe)) { return $null }
    # Stop a stale second-server from a previous run - but only one this repository
    # started. Test-OwnedTestServer is the proof (the command line names a
    # nats.conf under OUR temp work root); .Path cannot answer it, because a
    # Start-Process'd nats-server exposes an EMPTY Path, and a conf file in our
    # work dir alone would let us kill a stranger that happens to hold that port.
    $existing = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existing -and (Test-OwnedTestServer $Port)) {
        Stop-Process -Id $existing.OwningProcess -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
    }
    elseif ($existing) {
        # Somebody else's server holds the port. Leave it alone and fail the start,
        # so the caller skips this journey instead of killing a server we do not own.
        Write-Host "  port $Port is already in use by a server this run did not start; leaving it alone"
        return $null
    }
    Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    $conf = Join-Path $work 'nats.conf'
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $lines = @('server_name: e2eb', "port: $Port", 'jetstream {',
        "  store_dir: $((Join-Path $work 'data') -replace '\\','/')", '}')
    [System.IO.File]::WriteAllText($conf, (($lines -join "`r`n") + "`r`n"), $utf8)
    $p = Start-Process -FilePath $exe -ArgumentList '-c', $conf -WorkingDirectory $work -WindowStyle Hidden -PassThru
    $script:secondServerPid = $p.Id
    # Condition-based readiness: a real INFO/PONG handshake, not a fixed sleep.
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        try {
            $c = New-Object System.Net.Sockets.TcpClient
            $c.Connect('127.0.0.1', $Port)
            $s = $c.GetStream(); $s.ReadTimeout = 3000
            $buf = New-Object byte[] 4096
            $null = $s.Read($buf, 0, $buf.Length)
            $ping = [Text.Encoding]::ASCII.GetBytes("PING`r`n")
            $s.Write($ping, 0, $ping.Length)
            Start-Sleep -Milliseconds 150
            $n = $s.Read($buf, 0, $buf.Length)
            $reply = [Text.Encoding]::ASCII.GetString($buf, 0, $n)
            $c.Close()
            if ($reply -match 'PONG') { return "nats://127.0.0.1:$Port" }
        } catch { }
        Start-Sleep -Milliseconds 300
    }
    return $null
}
function Stop-SecondServer {
    if ($script:secondServerPid) {
        Stop-Process -Id $script:secondServerPid -Force -ErrorAction SilentlyContinue
        $script:secondServerPid = $null
    }
}

# Runs the nats CLI without letting its stderr (a "stream not found" on a first
# cleanup, say) become a terminating error: with $ErrorActionPreference='Stop' a
# native command writing to stderr throws. Returns the exit code.
function Invoke-NatsCli {
    param([string[]]$Arguments)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'
    try {
        & nats @Arguments 2>&1 | Out-Null
        return $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previous
    }
}

$isLocal = $NatsUrl -match '127\.0\.0\.1|localhost|\[::1\]'
$pub = "$SubjectPrefix.core"
# The ProgId is frozen, so every RTD formula in this script builds the same prefix.
$rtdQuoted = '"NatsRtd",,'    # closing quote for the literal

# This suite PUBLISHES, creates streams/consumers and (in the node-kill step)
# stops a server. Pointing it at a real feed is therefore a production incident,
# not a test run: a non-local endpoint is refused unless writing to it is an
# explicit, deliberate choice. D6 in PLAN.md: a real server is read-only by
# default. -ReadOnly (or a non-local endpoint) also skips every write below and
# keeps only the read paths, so a feed can be observed without touching it.
$readOnly = $ReadOnly -or (-not $isLocal)
if (-not $isLocal -and -not $AllowRemoteWrites) {
    throw ("refusing to run against the non-local endpoint ${NatsUrl}: this suite writes " +
           "(publishes, stream/consumer/KV/object creation) and stops a node. If that server " +
           "really is a disposable sandbox, re-run with -AllowRemoteWrites; to only OBSERVE a " +
           "feed, pass -ReadOnly.")
}

Log "=== NATSExcel E2E ==="
Log "xll    : $Xll"
Log "nats   : $NatsUrl  (local=$isLocal, bound=${BoundMs}ms, readOnly=$readOnly)"

# Is the local endpoint a server THIS repository started (the cluster harness's
# node under %TEMP%\natsexcel-cluster)? If not, the run is using somebody else's
# server: the writes below land in their instance and the failover journey must
# never kill them. Say it out loud instead of finding out from a denied kill.
$ownsServer = $isLocal -and (Test-OwnedTestServer ([Uri]$NatsUrl).Port)
if ($isLocal -and -not $ownsServer) {
    Log ("WARNING: the server on port " + ([Uri]$NatsUrl).Port +
         " was not started by this repository's harness. The destructive failover journey is skipped; " +
         "start a disposable cluster with test\TestSupport\Start-Cluster.ps1 for a complete run.")
}

Log (powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot 'test\TestSupport\Get-AntivirusState.ps1'))
Assert "XLL exists" (Test-Path $Xll)

# ---------------------------------------------------------------------------
# 1. UDF round trip, RTD push, JetStream/KV paths, in one hidden Excel
# ---------------------------------------------------------------------------
$excel = New-HiddenExcel
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = 'E2E'

Invoke-Com { $ws.Range('A1').Formula = '=NatsUDFVersion()' } | Out-Null
if (-not $readOnly) {
    Invoke-Com { $ws.Range('A2').Formula = '=NatsUDFPublish("' + $pub + '","hello","' + $NatsUrl + '")' } | Out-Null
    Invoke-Com { $ws.Range('A3').Formula = '=NatsUDFSubjectLatest("' + $pub + '","","' + $NatsUrl + '")' } | Out-Null
}
Invoke-Com { $excel.CalculateFullRebuild() } | Out-Null

$version = Wait-Cell $ws 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^#' }
Log "  version=$version"
Assert "UDF: the XLL answers (packing OK)" (-not [string]::IsNullOrWhiteSpace($version) -and $version -notmatch '^#')

if ($readOnly) {
    # Observe-only: publishing would touch the endpoint, so the write/read pair is
    # skipped rather than asserted (the read paths below still run).
    Log "  (read-only: skipping the publish and the round trip)"
} else {
    # Wait for a TERMINAL marker: under AsyncWrites the cell shows the pending
    # "#N/A" first, and an earlier predicate that accepted "#" stopped there and
    # then asserted "OK" against it. Only "OK..." or a named "Error:" settles.
    $marker = Wait-Cell $ws 'A2' { param($t) $t -match '^OK' -or $t -match '^Error' }
    Log "  publish=$marker"
    # The publish must return the OK marker, never an error or a pending #N/A: the
    # earlier predicate accepted "Error" and "#", so a broken write could pass.
    Assert "UDF: publish returns the OK marker" ($marker -match '^OK')

    # The read is the actual write -> read proof (a bare "returns something" check
    # let the round trip fail silently). The SUB registers asynchronously and Core
    # NATS is at-most-once, so re-publish until the latest-read cell observes the
    # value; a one-shot read could legitimately race the subscription to (null).
    $latest = '(null)'
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        Invoke-Com { $ws.Range('A2').Formula = '=NatsUDFPublish("' + $pub + '","hello","' + $NatsUrl + '")' } | Out-Null
        $latest = Wait-Cell $ws 'A3' { param($t) $t -eq 'hello' } 2500
        if ($latest -eq 'hello') { break }
    }
    Log "  latest=$latest"
    Assert "UDF: the published value is read back (write -> read round trip)" ($latest -eq 'hello')
}

# RTD SUB: use a DEDICATED subject and a NON-VOLATILE publisher. A volatile
# publish cell (A2 above) is recalculated on every calculation and republishes,
# which would overwrite the RTD cell with a stale value - a test artifact, not a
# product bug. The NonVolatile twin fires only when its argument cell changes.
#
# A fresh SUB registers asynchronously, so the FIRST publish can land before the
# subscription is live (at-most-once Core NATS - a real, documented property,
# not a defect). Retry fresh values until both the first push and a later update
# are observed, so the check proves the push path and is not a one-shot race.
$rtdSubject = "$SubjectPrefix.rtd"
Invoke-Com { $ws.Range('A5').Formula = '=RTD("NatsRtd",,"SUB","' + $rtdSubject + '","","' + $NatsUrl + '")' } | Out-Null
$rtdReady = Wait-Cell $ws 'A5' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
Assert "RTD: SUB did not error" ($rtdReady -notmatch '#ERROR')

if ($readOnly) {
    Log "  RTD ready=$rtdReady (read-only: skipping the publishes and the push checks)"
} else {
    Invoke-Com { $ws.Range('C1').Value2 = 'rtd-0' } | Out-Null
    Invoke-Com { $ws.Range('A6').Formula = '=NatsUDFPublishNonVolatile("' + $rtdSubject + '",C1,"' + $NatsUrl + '")' } | Out-Null
    $rtdFirst = ''
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        $try = "rtd-$attempt"
        Invoke-Com { $ws.Range('C1').Value2 = $try } | Out-Null
        $rtdFirst = Wait-Cell $ws 'A5' { param($t) $t -eq $try }
        if ($rtdFirst -eq $try) { break }
    }
    # A later argument change must re-run the write and update the cell.
    $rtdTarget = 'rtd-after'
    Invoke-Com { $ws.Range('C1').Value2 = $rtdTarget } | Out-Null
    $rtdAfter = Wait-Cell $ws 'A5' { param($t) $t -eq $rtdTarget }
    Log "  RTD first=$rtdFirst after=$rtdAfter"
    Assert "RTD: a push reached the cell" ($rtdFirst -match '^rtd-\d+$')
    Assert "RTD: a later push updated the cell" ($rtdAfter -eq $rtdTarget)
}

# JetStream + KV: each must return a value or a NAMED error, never a hang/blank.
Invoke-Com { $ws.Range('A8').Formula = '=NatsUDFKVGet("NoSuchBucket","k","' + $NatsUrl + '")' } | Out-Null
Invoke-Com { $ws.Range('A9').Formula = '=NatsUDFJSStreams("' + $NatsUrl + '")' } | Out-Null
$kv = Wait-Cell $ws 'A8' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^#' }
$streams = Wait-Cell $ws 'A9' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^#' }
Log "  kv(missing bucket)=$($kv.Substring(0,[Math]::Min(90,$kv.Length)))  jsStreams=$streams"
Assert "KV: a missing bucket is a named error" ($kv -match 'Error|not found')
# The streams call must ANSWER, not fail: requiring only "non-empty" let an
# "Error: JetStream is not enabled ..." string satisfy the check.
Assert "JetStream: the streams call answers with a value, not an error" ($streams -notmatch '^(Error|#)')

# ---------------------------------------------------------------------------
# 2. Multi-workbook: a copy closing must not disturb the original
# ---------------------------------------------------------------------------
if ($isLocal -and -not $SkipMultiWorkbook) {
    $copy = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-e2e-copy.xlsx'
    # A crashed previous run can leave its own copy (and Excel's "~$" lock marker)
    # behind; opening a file with a stale marker is one of the ways Excel answers
    # with no workbook at all. Both files are this suite's own temp output.
    Remove-Item $copy -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path (Split-Path -Parent $copy) ('~$' + (Split-Path -Leaf $copy))) -Force -ErrorAction SilentlyContinue
    Invoke-Com { $wb.SaveCopyAs($copy) } | Out-Null

    $excel2 = New-HiddenExcel
    # Workbooks.Open can answer with $null instead of throwing (a transient COM/AV
    # rejection, a file still being written), and Invoke-Com only retries on an
    # exception - so retry here rather than dereferencing a null workbook.
    $wb2 = $null
    for ($attempt = 1; $attempt -le 5 -and -not $wb2; $attempt++) {
        $wb2 = Invoke-Com { $excel2.Workbooks.Open($copy) }
        if (-not $wb2) { Start-Sleep -Milliseconds 600 }
    }

    if (-not $wb2) {
        # Never leave the hidden instance behind (it holds the copy and the XLL, and
        # the NEXT run then fails on a locked file): quit it, then skip with a reason.
        Log "  Excel returned no workbook for the saved copy ($copy); skipping the multi-workbook check"
        try { $excel2.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) | Out-Null } catch {}
        Remove-Item $copy -Force -ErrorAction SilentlyContinue
        Skip "Multi-workbook: the original keeps receiving"
    }
    else {
        try {
            $null = Invoke-Com { $wb2.Worksheets.Item(1).Range('A5').Text }   # force a calc
            Invoke-Com { $wb2.Close($false) } | Out-Null
        }
        finally {
            try { $excel2.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) | Out-Null } catch {}
        }

        # The RTD subject above is dedicated, so a publish cannot be overwritten by a
        # recalculated volatile cell. Change the NonVolatile publisher's argument.
        Invoke-Com { $ws.Range('C1').Value2 = 'after-copy' } | Out-Null
        $afterCopy = Wait-Cell $ws 'A5' { param($t) $t -eq 'after-copy' }
        Log "  RTD after the copy closed = $afterCopy"
        Assert "Multi-workbook: the original keeps receiving" ($afterCopy -eq 'after-copy')
        Remove-Item $copy -Force -ErrorAction SilentlyContinue
    }
} else {
    Skip "multi-workbook (non-local host or -SkipMultiWorkbook)"
}

Close-Excel $excel
Log "  (steps 1-2 done, fails=$fails, comRetries=$comRetries)"

# ---------------------------------------------------------------------------
# 3. Extra scenarios a user actually hits (skip with -SkipExtraChecks). Every
#    wait stays condition-based, so a healthy run adds a few seconds at most.
#    These run BEFORE the destructive node kill below (which leaves the URL
#    pointing at a dead node, so any later cell on that URL could not reach
#    the cluster).
# ---------------------------------------------------------------------------
if (-not $SkipExtraChecks -and -not $readOnly) {

    # ----------------------------------------------------------------------
    # 3a. Two Excel processes at once on the SAME subject. The runtime is
    #     per-process, so one publish must reach both live cells; after one
    #     process exits, the survivor must keep receiving (isolation).
    #     A brand-new instance registers its RTD topic asynchronously, so the
    #     first publish can land before its SUB exists; publish fresh values
    #     until BOTH cells show the same one (condition-based, never a fixed
    #     sleep), then that single publish provably reached both.
    # ----------------------------------------------------------------------
    Log "  --- extra: two Excel instances, same subject ---"
    $twoSubject = "$SubjectPrefix.two"
    $xa = New-HiddenExcel
    $wa = $xa.Workbooks.Add(); $sa = $wa.Worksheets.Item(1)
    $xb = New-HiddenExcel
    $wbx = $xb.Workbooks.Add(); $sbx = $wbx.Worksheets.Item(1)
    Invoke-Com { $sa.Range('A1').Formula = '=RTD("NatsRtd",,"SUB","' + $twoSubject + '","","' + $NatsUrl + '")' } | Out-Null
    Invoke-Com { $sbx.Range('A1').Formula = '=RTD("NatsRtd",,"SUB","' + $twoSubject + '","","' + $NatsUrl + '")' } | Out-Null
    $readyA = Wait-Cell $sa 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
    $readyB = Wait-Cell $sbx 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
    # The NonVolatile publisher writes on an argument CHANGE; C1 drives it.
    Invoke-Com { $sa.Range('C1').Value2 = 'x' } | Out-Null
    Invoke-Com { $sa.Range('A2').Formula = '=NatsUDFPublishNonVolatile("' + $twoSubject + '",C1,"' + $NatsUrl + '")' } | Out-Null

    $twoValue = $null
    $twoA = ''
    $twoB = ''
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        $try = "two-$attempt"
        Invoke-Com { $sa.Range('C1').Value2 = $try } | Out-Null
        $twoA = Wait-Cell $sa 'A1' { param($t) $t -eq $try } -TimeoutMs 3000
        $twoB = Wait-Cell $sbx 'A1' { param($t) $t -eq $try } -TimeoutMs 3000
        if ($twoA -eq $try -and $twoB -eq $try) { $twoValue = $try; break }
    }
    Log "  instance A ready=$readyA got=$twoA ; instance B ready=$readyB got=$twoB (attempt value=$twoValue)"
    Assert "Two instances: one publish reached instance A" ($twoA -eq $twoValue -and $twoValue)
    Assert "Two instances: the SAME publish reached instance B" ($twoB -eq $twoValue -and $twoValue)

    # Close the second process; the first must survive and keep receiving.
    Close-Excel $xb
    $survivor = $null
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        $try = "surv-$attempt"
        Invoke-Com { $sa.Range('C1').Value2 = $try } | Out-Null
        $got = Wait-Cell $sa 'A1' { param($t) $t -eq $try } -TimeoutMs 3000
        if ($got -eq $try) { $survivor = $try; break }
    }
    Log "  survivor after the other process closed = $survivor"
    Assert "Two instances: the survivor keeps receiving after the other closed" ($null -ne $survivor)
    Close-Excel $xa

    # ----------------------------------------------------------------------
    # 3b. Reopen the SAME workbook file: a user closes the sheet and opens it
    #     again; the saved RTD cell must come alive with a fresh subscribe.
    # ----------------------------------------------------------------------
    Log "  --- extra: reopen the same workbook ---"
    $reopenSubject = "$SubjectPrefix.reopen"
    $reopenPath = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-e2e-reopen.xlsx'
    Remove-Item $reopenPath -ErrorAction SilentlyContinue
    $xr = New-HiddenExcel
    $wr = $xr.Workbooks.Add(); $sr = $wr.Worksheets.Item(1)
    Invoke-Com { $sr.Range('A1').Formula = '=RTD("NatsRtd",,"SUB","' + $reopenSubject + '","","' + $NatsUrl + '")' } | Out-Null
    $rReady = Wait-Cell $sr 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
    Invoke-Com { $sr.Range('C1').Value2 = 'reopen-1' } | Out-Null
    Invoke-Com { $sr.Range('A2').Formula = '=NatsUDFPublishNonVolatile("' + $reopenSubject + '",C1,"' + $NatsUrl + '")' } | Out-Null
    $rFirst = Wait-Cell $sr 'A1' { param($t) $t -eq 'reopen-1' }
    Invoke-Com { $wr.SaveAs($reopenPath) } | Out-Null
    Invoke-Com { $wr.Close($false) } | Out-Null

    $wrOpen = Invoke-Com { $xr.Workbooks.Open($reopenPath) }
    $srOpen = $wrOpen.Worksheets.Item(1)
    $null = Invoke-Com { $srOpen.Range('A1').Text }   # force the RTD re-register
    # A new value proves the fresh subscribe is live (a cached cell would not move).
    Invoke-Com { $srOpen.Range('C1').Value2 = 'reopen-2' } | Out-Null
    $rSecond = Wait-Cell $srOpen 'A1' { param($t) $t -eq 'reopen-2' } -TimeoutMs 15000
    Log "  reopen ready=$rReady first=$rFirst afterReopen=$rSecond"
    Assert "Reopen: the saved workbook's RTD cell comes alive again" ($rSecond -eq 'reopen-2')
    Invoke-Com { $wrOpen.Close($false) } | Out-Null
    Close-Excel $xr
    Remove-Item $reopenPath -ErrorAction SilentlyContinue

    # ----------------------------------------------------------------------
    # 3c. ...NonVolatile runs on entry only. A PLAIN recalculation (F9 /
    #    Application.Calculate) must not re-send it. A FULL calculation
    #    (Ctrl+Alt+F9 / CalculateFull[Rebuild]) deliberately recomputes EVERY
    #    formula, non-volatile ones included, so it does re-run the write -
    #    that is Excel's documented behavior (the sibling project notes the
    #    same), not an add-in defect. The strong "runs once" proof therefore
    #    uses plain recalculations; the requested full-recalc sequence is run
    #    too and asserted to stay healthy.
    #
    #    Detection: a tripwire volatile publish overwrites the dedicated
    #    subject once; with the tripwire removed, a plain recalculation that
    #    re-ran the NonVolatile write would publish its value back and the RTD
    #    cell would revert. If it stays on the tripwire value, the write did
    #    not re-run.
    # ----------------------------------------------------------------------
    Log "  --- extra: ...NonVolatile single-run rule ---"
    $nvSubject = "$SubjectPrefix.nv"
    $xnv = New-HiddenExcel
    $wnv = $xnv.Workbooks.Add(); $snv = $wnv.Worksheets.Item(1)
    Invoke-Com { $snv.Range('A1').Formula = '=RTD("NatsRtd",,"SUB","' + $nvSubject + '","","' + $NatsUrl + '")' } | Out-Null
    $null = Wait-Cell $snv 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
    Invoke-Com { $snv.Range('C1').Value2 = 'nv-once' } | Out-Null
    Invoke-Com { $snv.Range('A2').Formula = '=NatsUDFPublishNonVolatile("' + $nvSubject + '",C1,"' + $NatsUrl + '")' } | Out-Null
    # Retry the entry publish until the fresh SUB delivers it (at-most-once Core;
    # a first publish can precede the async subscription - a race, not a defect).
    $nvFirst = ''
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        $try = "nv-once-$attempt"
        Invoke-Com { $snv.Range('C1').Value2 = $try } | Out-Null
        $nvFirst = Wait-Cell $snv 'A1' { param($t) $t -eq $try } -TimeoutMs 3000
        if ($nvFirst -eq $try) { break }
    }
    Assert "NonVolatile: the write ran on entry" ($nvFirst -match '^nv-once-\d+$')

    # Tripwire: overwrite the subject once from a volatile cell, then remove it.
    # (Retry a few values in case this fresh SUB also registered late.)
    Invoke-Com { $snv.Range('C2').Value2 = 'nv-trip-0' } | Out-Null
    Invoke-Com { $snv.Range('A3').Formula = '=NatsUDFPublish("' + $nvSubject + '",C2,"' + $NatsUrl + '")' } | Out-Null
    $nvTrip = ''
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        $try = "nv-trip-$attempt"
        Invoke-Com { $snv.Range('C2').Value2 = $try } | Out-Null
        $nvTrip = Wait-Cell $snv 'A1' { param($t) $t -eq $try } -TimeoutMs 3000
        if ($nvTrip -eq $try) { break }
    }
    Invoke-Com { $snv.Range('A3').Clear() } | Out-Null   # drop the tripwire publisher

    # Two plain recalculations must NOT re-run the NonVolatile write, so the
    # cell must stay on the tripwire value (a re-run would publish its argument
    # back and revert the cell).
    Invoke-Com { $xnv.Calculate() } | Out-Null
    Start-Sleep -Milliseconds 500
    Invoke-Com { $xnv.Calculate() } | Out-Null
    Start-Sleep -Milliseconds 500
    $nvAfterPlain = Invoke-Com { $snv.Range('A1').Text }
    Log "  tripwire=$nvTrip after two plain recalcs=$nvAfterPlain"
    Assert "NonVolatile: two plain recalculations did not re-publish" `
        ($nvTrip -match '^nv-trip-\d+$' -and $nvAfterPlain -eq $nvTrip)

    # The requested full recalcs recompute every formula (Ctrl+Alt+F9). Record
    # what happened and assert the cell stays a healthy value either way.
    Invoke-Com { $xnv.CalculateFullRebuild() } | Out-Null
    $afterRebuild = Wait-Cell $snv 'A1' { param($t) $t -match '^nv-(once|trip)-' }
    Invoke-Com { $xnv.CalculateFull() } | Out-Null
    $afterFull = Wait-Cell $snv 'A1' { param($t) $t -match '^nv-(once|trip)-' }
    $nvHealthy = ($afterRebuild -match '^nv-(once|trip)-') -and ($afterFull -match '^nv-(once|trip)-')
    Log "  full recalc: afterRebuild=$afterRebuild afterFull=$afterFull (a full recalc recomputes every formula by design)"
    Assert "NonVolatile: a full recalculation stays healthy (never blank/errored)" $nvHealthy
    Close-Excel $xnv

    # ----------------------------------------------------------------------
    # 3d. A bad endpoint is a bounded named error, never a hang. Port 1 on
    #    loopback refuses immediately, so the contract "never block the Excel
    #    thread" is measurable at the cell level. Use a READ (the RTT probe),
    #    not a publish: under fireforget-all a write is sent fire-and-forget
    #    and legitimately returns a marker instead of the error, whereas a read
    #    always waits bounded and reports the failure the same in every mode.
    # ----------------------------------------------------------------------
    Log "  --- extra: bad endpoint is a bounded error ---"
    $badEp = 'nats://127.0.0.1:1'
    $xbe = New-HiddenExcel
    $wbe = $xbe.Workbooks.Add(); $sbe = $wbe.Worksheets.Item(1)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-Com { $sbe.Range('A1').Formula = '=NatsUDFRtt("' + $badEp + '")' } | Out-Null
    $beText = Wait-Cell $sbe 'A1' { param($t) $t -match '^Error' } -TimeoutMs 15000
    $beMs = $sw.ElapsedMilliseconds
    Log "  bad endpoint: '$($beText.Substring(0,[Math]::Min(80,$beText.Length)))' after ${beMs}ms"
    Assert "Bad endpoint: the UDF shows a named Error: (no #N/A, no blank)" ($beText -match '^Error')
    Assert "Bad endpoint: the error is bounded (no Excel-thread hang) [${beMs}ms]" ($beMs -lt 10000)
    Invoke-Com { $wbe.Close($false) } | Out-Null
    Close-Excel $xbe

    # ----------------------------------------------------------------------
    # The RTD surface BEYOND SUB: the frozen command vocabulary, the errors the
    # runtime must NAME, the cluster cell, the six status helpers and the 255
    # bound at runtime. Only SUB was ever created in any layer, so a broken polled
    # kind, a reworded error or a leaked raw payload would have shipped unseen.
    # ----------------------------------------------------------------------
    Log "  --- extra: RTD vocabulary, cluster cell, status helpers, 255 bound ---"
    $xrt = New-HiddenExcel
    $wrt = $xrt.Workbooks.Add(); $srt = $wrt.Worksheets.Item(1)
    # The ProgId is frozen, so every RTD formula builds the same prefix.
    # (Defined with the other script variables above.)

    # Three frozen texts: an unknown command, the RTDCLUSTER vocabulary, and an
    # unresolvable reference (which must NEVER silently fall back to a default).
    Invoke-Com { $srt.Range('A1').Formula = '=RTD(' + $rtdQuoted + '"NOPE","x","' + $NatsUrl + '")' } | Out-Null
    $unknown = Wait-Cell $srt 'A1' { param($t) $t -match '^#ERROR' }
    Assert "RTD: an unknown command is the frozen named error" `
        ($unknown -eq "#ERROR: ConnectData: unknown command 'NOPE', expected one of [SUB, KV, KVGET, KVKEYS, REQ, JS, MSG, OBJ, RTDCLUSTER]")

    Invoke-Com { $srt.Range('A2').Formula = '=RTD(' + $rtdQuoted + '"RTDCLUSTER","bogus","' + $NatsUrl + '")' } | Out-Null
    $vocab = Wait-Cell $srt 'A2' { param($t) $t -match '^#ERROR' }
    Assert "RTD: the RTDCLUSTER vocabulary error is the frozen text" `
        ($vocab -match 'RTDCLUSTER expects what to be one of \[cluster, nodes, leader, lag\]')

    Invoke-Com { $srt.Range('A3').Formula = '=RTD(' + $rtdQuoted + '"SUB","' + $pub + '.rtd","","no-such-reference")' } | Out-Null
    $dangling = Wait-Cell $srt 'A3' { param($t) $t -match '^#ERROR' }
    Assert "RTD: an unresolvable reference is a named #ERROR (never a hidden default)" `
        ($dangling -match 'does not resolve to an endpoint')

    # The two Tier-1 RTDCLUSTER values (no $SYS account needed).
    Invoke-Com { $srt.Range('A4').Formula = '=RTD(' + $rtdQuoted + '"RTDCLUSTER","cluster","' + $NatsUrl + '")' } | Out-Null
    $clusterCell = Wait-Cell $srt 'A4' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^\(Connect' }
    Assert "RTD: RTDCLUSTER cluster answers (never an #ERROR when INFO is available)" `
        ($clusterCell -and $clusterCell -notmatch '^#ERROR')

    Invoke-Com { $srt.Range('A5').Formula = '=RTD(' + $rtdQuoted + '"RTDCLUSTER","nodes","' + $NatsUrl + '")' } | Out-Null
    $nodesCell = Wait-Cell $srt 'A5' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^\(Connect' }
    Assert "RTD: RTDCLUSTER nodes answers" ($nodesCell -and $nodesCell -notmatch '^#ERROR')

    # The 255 bound at RUNTIME (the unit source-scan only pins that the call is
    # there). Excel-DNA truncates at 255 silently, so the marker must come from us.
    # The SUB cell subscribes asynchronously and Core is at-most-once, so the
    # publish is retried (with a payload that CHANGES, or SkipRepeatedMessages
    # would suppress the repeat) until the cell shows the marker.
    Invoke-Com { $srt.Range('B3').Formula = '=RTD(' + $rtdQuoted + '"SUB","' + $pub + '.rtd255","","' + $NatsUrl + '")' } | Out-Null
    $bounded = ''
    for ($attempt = 1; $attempt -le 6 -and $bounded -notmatch '\(truncated \d+ chars\)$'; $attempt++) {
        Invoke-Com { $srt.Range('B1').Value2 = ('x' * 400) + $attempt } | Out-Null
        Invoke-Com { $srt.Range('B2').Formula = '=NatsUDFPublishNonVolatile("' + $pub + '.rtd255",B1,"' + $NatsUrl + '")' } | Out-Null
        $bounded = Wait-Cell $srt 'B3' { param($t) $t -match '\(truncated \d+ chars\)$' } -TimeoutMs 4000
    }
    Assert "RTD: an oversized value is bounded at 255 with a visible marker" `
        ($bounded.Length -le 255 -and $bounded -match '\(truncated \d+ chars\)$')

    # The six status helpers: real numbers/strings for THIS run, never a throw.
    Invoke-Com { $srt.Range('C1').Formula = '=NatsRTDTopicCount()' } | Out-Null
    Invoke-Com { $srt.Range('C2').Formula = '=NatsRTDConnectionCount()' } | Out-Null
    Invoke-Com { $srt.Range('C3').Formula = '=NatsRTDSubscriptionCount()' } | Out-Null
    Invoke-Com { $srt.Range('C4').Formula = '=NatsRTDSubjectCount()' } | Out-Null
    Invoke-Com { $srt.Range('C5').Formula = '=NatsRTDDefaultHost()' } | Out-Null
    Invoke-Com { $srt.Range('C6').Formula = '=NatsRTDCluster()' } | Out-Null
    $topics = Wait-Cell $srt 'C1' { param($t) $t -match '^\d+$' }
    $conns = Wait-Cell $srt 'C2' { param($t) $t -match '^\d+$' }
    $defaultHost = Wait-Cell $srt 'C5' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
    Assert "RTD status: the topic count is live (this sheet registered topics)" ([int]$topics -ge 1)
    Assert "RTD status: the connection count is live" ([int]$conns -ge 1)
    Assert "RTD status: the default host reports the configured endpoint" ($defaultHost -eq $NatsUrl)
    Invoke-Com { $wrt.Close($false) } | Out-Null
    Close-Excel $xrt
} else {
    Skip "extra scenarios (-SkipExtraChecks)"
}

# ---------------------------------------------------------------------------
# 4. User journeys (skip with -SkipJourneys). Everything a person actually does
#    with a workspace: a real multi-cluster config, switching the default and
#    reloading without a restart, a missing config, replaying JetStream from
#    cells, a JSON range in Key/Value, a scaled sheet of live cells and the
#    async pending marker. Every wait is condition-based; a scenario that cannot
#    run prints SKIP, never FAIL.
# ---------------------------------------------------------------------------
if (-not $SkipJourneys -and -not $readOnly) {

    # The config file the add-in folder carries; rewriting it only happens when
    # this script built that throwaway copy.
    $cfgFile = Join-Path $script:addInDir 'NATSExcel.json'
    $haveNatsCli = [bool](Get-Command nats -ErrorAction SilentlyContinue)

    # A second, INdependent endpoint. -NatsUrl2 wins; a local host with none is
    # given a disposable one-node server; a non-local host gets none.
    $url2 = $NatsUrl2
    if (-not $url2 -and $isLocal) {
        $basePort = ([Uri]$NatsUrl).Port
        Log "  --- journeys: starting a second independent server on $($basePort + 10) ---"
        $url2 = Start-SecondServer ($basePort + 10)
    }
    $haveUrl2 = -not [string]::IsNullOrWhiteSpace($url2)
    if (-not $haveUrl2) { Log "  (no second endpoint: -NatsUrl2 unset or a server could not start)" }

    # ======================================================================
    # J1. A real multi-cluster config: cluster name, server alias and raw URL
    #     each reach their own endpoint. A value published on B must not appear
    #     when reading through A (a mix-up is then detectable, not merely absent).
    # ======================================================================
    if ($script:builtCopy -and $haveUrl2) {
        Log "  --- journey: multi-cluster config (name / alias / raw URL) ---"
        Write-Cfg -Path $cfgFile -Default 'clusterA' -Url $NatsUrl -Url2 $url2 -Admin $false
        $j1Subject = "$SubjectPrefix.j1"
        $xj = New-HiddenExcel
        $wj = $xj.Workbooks.Add(); $sj = $wj.Worksheets.Item(1)

        Invoke-Com { $sj.Range('A1').Formula = '=NatsUDFConfigPath()' } | Out-Null
        $cfgSeen = Wait-Cell $sj 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
        Assert "Journey1: the add-in loaded the multi-cluster config" ($cfgSeen -eq $cfgFile)

        # A SUB on clusterB and a SUB on clusterA, same subject.
        Invoke-Com { $sj.Range('B1').Formula = '=RTD("NatsRtd",,"SUB","' + $j1Subject + '","","clusterB")' } | Out-Null
        Invoke-Com { $sj.Range('B2').Formula = '=RTD("NatsRtd",,"SUB","' + $j1Subject + '","","clusterA")' } | Out-Null
        $null = Wait-Cell $sj 'B1' { param($t) $t -ne '(ConnectData)' -and -not [string]::IsNullOrWhiteSpace($t) }
        $null = Wait-Cell $sj 'B2' { param($t) $t -ne '(ConnectData)' -and -not [string]::IsNullOrWhiteSpace($t) }

        # Raw-URL publish to B: only the clusterB cell may move.
        Invoke-Com { $sj.Range('D1').Value2 = 'on-B' } | Out-Null
        Invoke-Com { $sj.Range('D2').Formula = '=NatsUDFPublishNonVolatile("' + $j1Subject + '",D1,"' + $url2 + '")' } | Out-Null
        $null = Wait-Cell $sj 'D2' { param($t) $t -notmatch '^\(Connect' -and -not [string]::IsNullOrWhiteSpace($t) }
        $onB = Wait-Cell $sj 'B1' { param($t) $t -eq 'on-B' }
        $notOnA = Invoke-Com { $sj.Range('B2').Text }
        Log "  raw URL -> B: cellB='$onB' cellA='$notOnA'"
        Assert "Journey1: a raw-URL publish reached cluster B" ($onB -eq 'on-B')
        Assert "Journey1: the cluster-A cell did NOT receive the B value (endpoints are distinct)" ($notOnA -ne 'on-B')

        # Cluster-name publish to A: only the clusterA cell may move.
        Invoke-Com { $sj.Range('D3').Value2 = 'on-A' } | Out-Null
        Invoke-Com { $sj.Range('D4').Formula = '=NatsUDFPublishNonVolatile("' + $j1Subject + '",D3,"clusterA")' } | Out-Null
        $null = Wait-Cell $sj 'D4' { param($t) $t -notmatch '^\(Connect' -and -not [string]::IsNullOrWhiteSpace($t) }
        $onA = Wait-Cell $sj 'B2' { param($t) $t -eq 'on-A' }
        $stillB = Invoke-Com { $sj.Range('B1').Text }
        Log "  cluster name -> A: cellA='$onA' cellB='$stillB'"
        Assert "Journey1: a cluster-name publish reached cluster A" ($onA -eq 'on-A')
        Assert "Journey1: cluster B kept its own value (no bleed)" ($stillB -eq 'on-B')

        # Server alias -> B, and the alias resolves to the second endpoint.
        Invoke-Com { $sj.Range('D5').Value2 = 'via-alias' } | Out-Null
        Invoke-Com { $sj.Range('D6').Formula = '=NatsUDFPublishNonVolatile("' + $j1Subject + '",D5,"srvB")' } | Out-Null
        $null = Wait-Cell $sj 'D6' { param($t) $t -notmatch '^\(Connect' -and -not [string]::IsNullOrWhiteSpace($t) }
        $viaAlias = Wait-Cell $sj 'B1' { param($t) $t -eq 'via-alias' }
        $aliasNotA = Invoke-Com { $sj.Range('B2').Text }
        Log "  server alias -> B: cellB='$viaAlias' cellA='$aliasNotA'"
        Assert "Journey1: a server-alias publish reached cluster B" ($viaAlias -eq 'via-alias')
        Assert "Journey1: a server alias is not the same endpoint as cluster A" ($aliasNotA -ne 'via-alias')
        Close-Excel $xj
    } else {
        Skip "Journey1 multi-cluster config (needs the built copy and a second endpoint)"
    }

    # ======================================================================
    # J2. Switch Default without restarting Excel: edit NATSExcel.json, call
    #     ConfigReload, and a subsequent blank-host call resolves to the NEW
    #     default. The connection is cached per endpoint, but the endpoint is
    #     resolved per call, so the next call lands on the new endpoint.
    # ======================================================================
    if ($script:builtCopy -and $haveUrl2) {
        Log "  --- journey: switch Default without an Excel restart ---"
        Write-Cfg -Path $cfgFile -Default 'clusterA' -Url $NatsUrl -Url2 $url2 -Admin $false
        $j2Subject = "$SubjectPrefix.j2"
        $xj = New-HiddenExcel
        $wj = $xj.Workbooks.Add(); $sj = $wj.Worksheets.Item(1)
        # A SUB on B lets a blank-host (Default) publish be attributed: before the
        # switch Default is A and B stays silent; after, Default is B.
        Invoke-Com { $sj.Range('A1').Formula = '=RTD("NatsRtd",,"SUB","' + $j2Subject + '","","clusterB")' } | Out-Null
        $null = Wait-Cell $sj 'A1' { param($t) $t -ne '(ConnectData)' -and -not [string]::IsNullOrWhiteSpace($t) }
        Invoke-Com { $sj.Range('C1').Value2 = 'before' } | Out-Null
        Invoke-Com { $sj.Range('C2').Formula = '=NatsUDFPublishNonVolatile("' + $j2Subject + '",C1,"")' } | Out-Null
        $null = Wait-Cell $sj 'C2' { param($t) $t -notmatch '^\(Connect' -and -not [string]::IsNullOrWhiteSpace($t) }
        Start-Sleep -Milliseconds 700
        $before = Invoke-Com { $sj.Range('A1').Text }
        Assert "Journey2: the default (clusterA) publish did NOT reach cluster B" ($before -ne 'before')

        # Edit the file, reload, and the SAME blank-host formula now targets B.
        Write-Cfg -Path $cfgFile -Default 'clusterB' -Url $NatsUrl -Url2 $url2 -Admin $false
        Invoke-Com { $sj.Range('E1').Formula = '=NatsUDFConfigReload()' } | Out-Null
        $reload = Wait-Cell $sj 'E1' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^\(Connect' }
        Log "  reload = '$reload'"
        Assert "Journey2: ConfigReload reports OK" ($reload -match '^OK')
        Invoke-Com { $sj.Range('C1').Value2 = 'after' } | Out-Null
        $after = Wait-Cell $sj 'A1' { param($t) $t -eq 'after' }
        Log "  after the switch, blank-host publish reached cluster B: '$after'"
        Assert "Journey2: after reload the NEW default endpoint is used" ($after -eq 'after')
        Close-Excel $xj
    } else {
        Skip "Journey2 switch default (needs the built copy and a second endpoint)"
    }

    # ======================================================================
    # J3. Delete the config, then reload. The reload cannot read a file that is
    #     gone, but the process is STILL using the config loaded at startup (the
    #     root is cached): the cell says so instead of claiming "using defaults".
    # ======================================================================
    if ($script:builtCopy) {
        Log "  --- journey: delete the config, then reload ---"
        Write-Cfg -Path $cfgFile -Default 'clusterA' -Url $NatsUrl -Url2 $(if ($haveUrl2) { $url2 } else { $NatsUrl }) -Admin $false
        $xj = New-HiddenExcel
        $wj = $xj.Workbooks.Add(); $sj = $wj.Worksheets.Item(1)
        Invoke-Com { $sj.Range('A1').Formula = '=NatsUDFConfigPath()' } | Out-Null
        $pathBefore = Wait-Cell $sj 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
        Assert "Journey3: a config is loaded before the delete" ($pathBefore -eq $cfgFile)
        Remove-Item -LiteralPath $cfgFile -Force
        Invoke-Com { $sj.Range('A2').Formula = '=NatsUDFConfigReload()' } | Out-Null
        $delReload = Wait-Cell $sj 'A2' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^\(Connect' }
        Log "  delete + reload = '$delReload'"
        Assert "Journey3: reload after a delete reports the previous config is still in use" `
            ($delReload -match 'still using the configuration loaded at startup')
        Close-Excel $xj
    } else {
        Skip "Journey3 delete-config reload (needs the built copy)"
    }

    # Restore a valid config for the journeys that follow (J4+ need only an
    # endpoint that works; they pass it explicitly).
    if ($script:builtCopy) {
        Write-Cfg -Path $cfgFile -Default 'clusterA' -Url $NatsUrl -Url2 $(if ($haveUrl2) { $url2 } else { $NatsUrl }) -Admin $false
    }

    # ======================================================================
    # J4. A cell against a dead host is a bounded error, not a hang. Port 1 on
    #     loopback refuses immediately; a READ waits bounded and reports.
    # ======================================================================
    Log "  --- journey: a dead host is a bounded error ---"
    $badEp = 'nats://127.0.0.1:1'
    $xj = New-HiddenExcel
    $wj = $xj.Workbooks.Add(); $sj = $wj.Worksheets.Item(1)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-Com { $sj.Range('A1').Formula = '=NatsUDFRtt("' + $badEp + '")' } | Out-Null
    $deadText = Wait-Cell $sj 'A1' { param($t) $t -match '^Error' } -TimeoutMs 15000
    $deadMs = $sw.ElapsedMilliseconds
    Log "  dead host: '$($deadText.Substring(0,[Math]::Min(80,$deadText.Length)))' after ${deadMs}ms"
    Assert "Journey4: a dead host shows a named Error: (no #N/A, no blank)" ($deadText -match '^Error')
    Assert "Journey4: the dead-host error is bounded (no Excel-thread hang) [${deadMs}ms]" ($deadMs -lt 10000)
    Close-Excel $xj

    # ======================================================================
    # J5. JetStream replay from a cell: a stream seeded from outside the sheet,
    #     read back by sequence and by last-message.
    # ======================================================================
    Log "  --- journey: replay JetStream from cells ---"
    $j5Stream = ($SubjectPrefix -replace '[^A-Za-z0-9]', '_') + '_J5'
    $j5Subject = "$SubjectPrefix.j5"
    $seeded = $false
    if ($haveNatsCli) {
        try {
            # Memory storage keeps the store small; a same-named stale stream from
            # a prior run is removed first (the tags are the only identifier).
            Invoke-NatsCli @('stream', 'rm', $j5Stream, '--server', $NatsUrl, '--force') | Out-Null
            Invoke-NatsCli @('stream', 'add', $j5Stream, '--subjects', $j5Subject, '--storage', 'memory', '--server', $NatsUrl, '--defaults') | Out-Null
            foreach ($v in 'm1', 'm2', 'm3') { Invoke-NatsCli @('pub', $j5Subject, $v, '--server', $NatsUrl) | Out-Null }
            $seeded = $true
        } catch { $seeded = $false }
    }
    if ($seeded) {
        $xj = New-HiddenExcel
        $wj = $xj.Workbooks.Add(); $sj = $wj.Worksheets.Item(1)
        Invoke-Com { $sj.Range('A1').Formula = '=NatsUDFJSGetMessage("' + $j5Stream + '",1,"' + $NatsUrl + '")' } | Out-Null
        Invoke-Com { $sj.Range('A2').Formula = '=NatsUDFJSGetMessage("' + $j5Stream + '",3,"' + $NatsUrl + '")' } | Out-Null
        Invoke-Com { $sj.Range('A3').Formula = '=NatsUDFJSGetLastMessage("' + $j5Stream + '","' + $j5Subject + '","' + $NatsUrl + '")' } | Out-Null
        $seq1 = Wait-Cell $sj 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^\(Connect' }
        $seq3 = Wait-Cell $sj 'A2' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^\(Connect' }
        $last = Wait-Cell $sj 'A3' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^\(Connect' }
        Log "  replay seq1='$seq1' seq3='$seq3' last='$last'"
        Assert "Journey5: replay by sequence 1 returns the first message" ($seq1 -eq 'm1')
        Assert "Journey5: replay by sequence 3 returns the third message" ($seq3 -eq 'm3')
        Assert "Journey5: last-message returns the latest (m3)" ($last -eq 'm3')

        # The two JetStream RTD kinds on this same stream: MSG polls the last
        # message matching the filter, JS binds an EXISTING durable consumer (so the
        # CLI creates one first) and consumes it.
        $j5Consumer = 'rtde2e'
        Invoke-NatsCli @('consumer', 'add', $j5Stream, $j5Consumer, '--pull',
            '--deliver', 'all', '--ack', 'none', '--filter', $j5Subject,
            '--server', $NatsUrl, '--defaults') | Out-Null

        Invoke-Com { $sj.Range('B1').Formula = '=RTD(' + $rtdQuoted + '"MSG","' + $j5Stream + '","' + $j5Subject + '","' + $NatsUrl + '")' } | Out-Null
        $msgCell = Wait-Cell $sj 'B1' { param($t) $t -eq 'm3' } -TimeoutMs 20000
        Log "  RTD MSG cell = '$msgCell'"
        Assert "Journey5: the RTD MSG cell shows the last message (m3)" ($msgCell -eq 'm3')

        Invoke-Com { $sj.Range('B2').Formula = '=RTD(' + $rtdQuoted + '"JS","' + $j5Stream + '","' + $j5Consumer + '","' + $j5Subject + '","' + $NatsUrl + '")' } | Out-Null
        $jsCell = Wait-Cell $sj 'B2' { param($t) $t -eq 'm3' } -TimeoutMs 25000
        Log "  RTD JS cell = '$jsCell'"
        Assert "Journey5: the RTD JS cell consumed the durable consumer (m3)" ($jsCell -eq 'm3')

        Close-Excel $xj
        Invoke-NatsCli @('stream', 'rm', $j5Stream, '--server', $NatsUrl, '--force') | Out-Null
    } else {
        Skip "Journey5 JetStream replay (nats CLI unavailable)"
    }

    # ======================================================================
    # J6. KV as a small config store in the sheet: put a JSON range into a key,
    #     read it back with JSONToMatrix(NatsUDFKVGet(...)) and round-trip.
    # ======================================================================
    Log "  --- journey: a JSON range in Key/Value ---"
    $j6Bucket = ($SubjectPrefix -replace '[^A-Za-z0-9]', '_') + '_J6'
    $seeded = $false
    if ($haveNatsCli) {
        try {
            Invoke-NatsCli @('kv', 'del', $j6Bucket, '--server', $NatsUrl, '--force') | Out-Null
            Invoke-NatsCli @('kv', 'add', $j6Bucket, '--server', $NatsUrl, '--history', '5') | Out-Null
            $seeded = $true
        } catch { $seeded = $false }
    }
    if ($seeded) {
        $xj = New-HiddenExcel
        $wj = $xj.Workbooks.Add(); $sj = $wj.Worksheets.Item(1)
        Invoke-Com { $sj.Range('A1').Value2 = 10 } | Out-Null
        Invoke-Com { $sj.Range('B1').Value2 = 20 } | Out-Null
        Invoke-Com { $sj.Range('A2').Value2 = 30 } | Out-Null
        Invoke-Com { $sj.Range('B2').Value2 = 40 } | Out-Null
        Invoke-Com { $sj.Range('D1').Formula = '=NatsUDFKVPutJSON("' + $j6Bucket + '","cfg",A1:B2,"' + $NatsUrl + '")' } | Out-Null
        # The revision cell is a write RESULT: re-enter the formula until it comes
        # back (the first evaluation can still observe the async dispatch in flight).
        $rev = ''
        for ($attempt = 1; $attempt -le 4 -and $rev -notmatch '^\d+$'; $attempt++) {
            $rev = Wait-Cell $sj 'D1' { param($t) $t -match '^\d+$' } -TimeoutMs 4000
            if ($rev -notmatch '^\d+$') {
                Start-Sleep -Milliseconds 250
                Invoke-Com { $sj.Range('D1').Formula = '=NatsUDFKVPutJSON("' + $j6Bucket + '","cfg",A1:B2,"' + $NatsUrl + '")' } | Out-Null
            }
        }
        Assert "Journey6: the JSON range was written to KV (a revision came back)" ($rev -match '^\d+$')
        # Expand the stored JSON back into the sheet: JSONToMatrix over the KV
        # read, indexed cell by cell. (A multi-cell FormulaArray registration is
        # not reliable across Excel builds under COM; INDEX is.)
        $j6Read = 'NatsUDFKVGet("' + $j6Bucket + '","cfg","' + $NatsUrl + '")'
        # The reads are async and the FIRST one can still observe the key as absent
        # (returning the "(null)" marker, which JSONToMatrix then rejects with
        # "invalid JSON: Unexpected character: ("), while the next three read the
        # value - a run-to-run flake. Re-enter the formulas until the whole range
        # round-trips, the same retry idiom the rest of this suite uses.
        $rt = @()
        for ($attempt = 1; $attempt -le 4; $attempt++) {
            $rt = @()
            foreach ($spec in @('D3,1,1', 'E3,1,2', 'D4,2,1', 'E4,2,2')) {
                $parts = $spec.Split(',')
                $addr = $parts[0]; $r = $parts[1]; $c = $parts[2]
                Invoke-Com { $sj.Range($addr).Formula = '=INDEX(NatsUDFJSONToMatrix(' + $j6Read + '),' + $r + ',' + $c + ')' } | Out-Null
                $rt += (Wait-Cell $sj $addr { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^\(Connect|^#' })
            }
            if (($rt[0] -match '^10') -and ($rt[1] -match '^20') -and ($rt[2] -match '^30') -and ($rt[3] -match '^40')) { break }
            Start-Sleep -Milliseconds 250
        }
        Log "  KV JSON round trip = $($rt -join '|')"
        Assert "Journey6: the JSON range round-tripped through KV and back" `
            (($rt[0] -match '^10') -and ($rt[1] -match '^20') -and ($rt[2] -match '^30') -and ($rt[3] -match '^40'))

        # The three KV RTD kinds on this same bucket: KVGET polls the key, KVKEYS
        # lists it, KV watches the entry subject and follows a revision.
        Invoke-Com { $sj.Range('E1').Formula = '=RTD(' + $rtdQuoted + '"KVGET","' + $j6Bucket + '","cfg","' + $NatsUrl + '")' } | Out-Null
        $kvgetCell = Wait-Cell $sj 'E1' { param($t) $t -match '^\[\[10' } -TimeoutMs 20000
        Log "  RTD KVGET cell = '$kvgetCell'"
        Assert "Journey6: the RTD KVGET cell shows the stored JSON" ($kvgetCell -match '^\[\[10')

        Invoke-Com { $sj.Range('E2').Formula = '=RTD(' + $rtdQuoted + '"KVKEYS","' + $j6Bucket + '","' + $NatsUrl + '")' } | Out-Null
        $kvkeysCell = Wait-Cell $sj 'E2' { param($t) $t -match 'cfg' } -TimeoutMs 20000
        Log "  RTD KVKEYS cell = '$kvkeysCell'"
        Assert "Journey6: the RTD KVKEYS cell lists the key" ($kvkeysCell -match 'cfg')

        # The KV kind is a WATCH: a Core subscription on the entry subject, so it
        # sees revisions published AFTER it goes live (Core is at-most-once, and the
        # current value is not replayed). Publish until the cell follows, the same
        # retry idiom the rest of this suite uses. H1 holds the value as TEXT: a
        # cell that has held a text-valued formula rejects a numeric assignment with
        # a COM InvalidCastException in this Excel build.
        Invoke-Com { $sj.Range('E3').Formula = '=RTD(' + $rtdQuoted + '"KV","' + $j6Bucket + '","cfg","' + $NatsUrl + '")' } | Out-Null
        $kvFollow = ''
        for ($attempt = 1; $attempt -le 6 -and $kvFollow -notmatch '^9'; $attempt++) {
            Invoke-Com { $sj.Range('H1').Value2 = ('9' + $attempt) } | Out-Null
            Invoke-Com { $sj.Range('H2').Formula = '=NatsUDFKVPutNonVolatile("' + $j6Bucket + '","cfg",H1,"' + $NatsUrl + '")' } | Out-Null
            $kvFollow = Wait-Cell $sj 'E3' { param($t) $t -match '^9\d$' } -TimeoutMs 4000
        }
        Log "  RTD KV watch cell = '$kvFollow'"
        Assert "Journey6: the RTD KV watch follows a revision" ($kvFollow -match '^9\d$')

        # The OBJ kind watches an object's METADATA (never its body), so it has the
        # same push semantics: create the bucket, register the cell, then put.
        $j6Obj = ($SubjectPrefix -replace '[^A-Za-z0-9]', '_') + '_J6O'
        $objFile = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-e2e-object.txt'
        try {
            [System.IO.File]::WriteAllText($objFile, 'e2e-object', [System.Text.Encoding]::ASCII)
            Invoke-NatsCli @('object', 'del', $j6Obj, '-f', '--server', $NatsUrl) | Out-Null
            Invoke-NatsCli @('object', 'add', $j6Obj, '--server', $NatsUrl) | Out-Null
            Invoke-Com { $sj.Range('F1').Formula = '=RTD(' + $rtdQuoted + '"OBJ","' + $j6Obj + '","cfg.txt","' + $NatsUrl + '")' } | Out-Null
            $objCell = ''
            for ($attempt = 1; $attempt -le 6 -and $objCell -notmatch 'cfg\.txt'; $attempt++) {
                Invoke-NatsCli @('object', 'put', $j6Obj, $objFile, '--name', 'cfg.txt', '--server', $NatsUrl) | Out-Null
                $objCell = Wait-Cell $sj 'F1' { param($t) $t -match 'cfg\.txt' } -TimeoutMs 4000
            }
            Log "  RTD OBJ cell = '$objCell'"
            Assert "Journey6: the RTD OBJ cell shows the object metadata" ($objCell -match 'cfg\.txt')
            Invoke-NatsCli @('object', 'del', $j6Obj, '-f', '--server', $NatsUrl) | Out-Null
        } catch {
            Skip "Journey6 RTD OBJ cell (nats CLI object support unavailable)"
        }
        Remove-Item $objFile -Force -ErrorAction SilentlyContinue

        Close-Excel $xj
        Invoke-NatsCli @('kv', 'del', $j6Bucket, '--server', $NatsUrl, '--force') | Out-Null
    } else {
        Skip "Journey6 KV JSON round trip (nats CLI unavailable)"
    }

    # ======================================================================
    # J7. A sheet with ~20 live SUB cells stays responsive: distinct subjects,
    #     all updated, and a recalculation returns within a generous bound.
    # ======================================================================
    Log "  --- journey: a scaled sheet (~20 live cells) ---"
    $xj = New-HiddenExcel
    $wj = $xj.Workbooks.Add(); $sj = $wj.Worksheets.Item(1)
    $cellCount = 20
    $sw7 = [System.Diagnostics.Stopwatch]::StartNew()
    # Register all the SUB cells and their NonVolatile publisher in one pass, then
    # prove each cell live by re-firing a fresh value until it shows (at-most-once
    # Core delivery: a publish can precede the async subscription - a race, not a
    # defect). A (ConnectData) placeholder only clears once a value arrives, so
    # there is no separate readiness step to poll for.
    for ($i = 1; $i -le $cellCount; $i++) {
        $subject = "$SubjectPrefix.j7.$i"
        $subCell = "A$i"
        $valueCell = "C$i"
        $pubCell = "D$i"
        $subFormula = '=RTD("NatsRtd",,"SUB","' + $subject + '","","' + $NatsUrl + '")'
        $pubFormula = '=NatsUDFPublishNonVolatile("' + $subject + '",' + $valueCell + ',"' + $NatsUrl + '")'
        Invoke-Com { $sj.Range($subCell).Formula = $subFormula } | Out-Null
        Invoke-Com { $sj.Range($valueCell).Value2 = "v$i-1" } | Out-Null
        Invoke-Com { $sj.Range($pubCell).Formula = $pubFormula } | Out-Null
    }
    $live = 0
    for ($i = 1; $i -le $cellCount; $i++) {
        $valueCell = "C$i"
        $matched = $false
        for ($attempt = 1; $attempt -le 6 -and -not $matched; $attempt++) {
            $want = "v$i-$attempt"
            Invoke-Com { $sj.Range($valueCell).Value2 = $want } | Out-Null
            $t = Wait-Cell $sj "A$i" { param($t) $t -eq $want } -TimeoutMs 1200
            if ($t -eq $want) { $matched = $true; $live++ }
        }
    }
    # A full recalculation of the whole sheet must return within a generous bound.
    $swRecalc = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-Com { $xj.CalculateFullRebuild() } | Out-Null
    $null = Invoke-Com { $sj.Range('A1').Text }
    $recalcMs = $swRecalc.ElapsedMilliseconds
    $j7Ms = $sw7.ElapsedMilliseconds
    Log "  scaled sheet: $live/$cellCount cells live in ${j7Ms}ms; full recalc ${recalcMs}ms"
    Assert "Journey7: all $cellCount live cells received their value" ($live -eq $cellCount)
    Assert "Journey7: the sheet recalculates within the bound [${recalcMs}ms]" ($recalcMs -lt 20000)
    Close-Excel $xj

    # ======================================================================
    # J8. The #N/A pending marker of an async write. Under -AsyncWrites a write
    #     cell returns Excel's pending marker first; a connection to an
    #     unreachable host makes that window observable. WorksheetFunction.IsNA
    #     is locale-independent, so the marker is checked as a predicate, not by
    #     scraping the text (which reads "#N/A" in an English locale only).
    # ======================================================================
    if ($AsyncWrites) {
        Log "  --- journey: the #N/A pending marker under -AsyncWrites ---"
        $xj = New-HiddenExcel
        $wj = $xj.Workbooks.Add(); $sj = $wj.Worksheets.Item(1)
        Invoke-Com { $sj.Range('A1').Value2 = 'pending' } | Out-Null
        # A request against a refused loopback port: the write is dispatched
        # asynchronously and the cell shows #N/A until the connect fails.
        Invoke-Com { $sj.Range('A2').Formula = '=NatsUDFRequest("test.natsexcel.pending","ping",500,"' + $badEp + '")' } | Out-Null
        $pendingSeen = $false
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($sw.ElapsedMilliseconds -lt 8000 -and -not $pendingSeen) {
            try { if ($xj.WorksheetFunction.IsNA($sj.Range('A2'))) { $pendingSeen = $true } } catch { }
            Start-Sleep -Milliseconds 40
        }
        $pendingMs = $sw.ElapsedMilliseconds
        $resolved = Wait-Cell $sj 'A2' { param($t) $t -match '^Error' } -TimeoutMs 15000
        Log "  pending marker seen=$pendingSeen after ${pendingMs}ms; resolved='$($resolved.Substring(0,[Math]::Min(50,$resolved.Length)))'"
        Assert "Journey8: the async write showed Excel's pending marker (#N/A)" $pendingSeen
        Assert "Journey8: the pending write later resolved to a named Error:" ($resolved -match '^Error')
        Close-Excel $xj
    } else {
        Skip "Journey8 async pending marker (needs -AsyncWrites)"
    }

    Stop-SecondServer
} else {
    Skip "user journeys (-SkipJourneys)"
}

# ---------------------------------------------------------------------------
# 5. Excel restart: a fresh instance must load the XLL and work again
# ---------------------------------------------------------------------------
$excel = New-HiddenExcel
$wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
Invoke-Com { $ws.Range('A1').Formula = '=NatsUDFVersion()' } | Out-Null
$restart = Wait-Cell $ws 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^#' }
Log "  after 'restart' version=$restart"
Assert "Excel restart: the XLL reloads and answers" (-not [string]::IsNullOrWhiteSpace($restart) -and $restart -notmatch '^#')
if (-not $KeepExcelOpen) { Close-Excel $excel } else { $excel.Visible = $true }

# ---------------------------------------------------------------------------
# 6. Cluster failover: kill the node the add-in is on, expect recovery. LAST,
#    because it is destructive: it leaves the URL pointing at a dead node.
# ---------------------------------------------------------------------------
if ($isLocal -and -not $SkipNodeKill -and -not $readOnly) {
    $port = ([Uri]$NatsUrl).Port
    $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
    $ownerPid = if ($conn) { $conn.OwningProcess } else { $null }
    $proc = if ($ownerPid) { Get-Process -Id $ownerPid -ErrorAction SilentlyContinue } else { $null }

    # NEVER kill a server this run did not start: the journey is destructive by
    # design ("it leaves the URL pointing at a dead node"), and a stranger
    # answering the port may be a developer's own server - killing it is exactly
    # what AGENTS.md lesson 3 forbids. Test-OwnedTestServer proves ownership from
    # the command line naming a nats.conf under OUR temp work root.
    if ($proc -and $proc.ProcessName -eq 'nats-server' -and -not (Test-OwnedTestServer $port)) {
        Log "  port $port is answered by pid $ownerPid, which this run did not start (not a disposable test node)"
        Skip "Failover: data still flows after the node died"
    }
    elseif ($proc -and $proc.ProcessName -eq 'nats-server') {
        $excel = New-HiddenExcel
        $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
        Invoke-Com { $ws.Range('A1').Formula = '=RTD("NatsRtd",,"SUB","' + $pub + '","","' + $NatsUrl + '")' } | Out-Null
        $ready = Wait-Cell $ws 'A1' { param($t) $t -ne '(ConnectData)' -and -not [string]::IsNullOrWhiteSpace($t) }
        Invoke-Com { $ws.Range('A2').Formula = '=NatsUDFConnectionState("' + $NatsUrl + '")' } | Out-Null
        $stateBefore = Wait-Cell $ws 'A2' { param($t) -not [string]::IsNullOrWhiteSpace($t) }

        Log "  killing the connected node (pid $ownerPid) to force failover"
        $killed = $true
        try { Stop-Process -Id $ownerPid -Force -ErrorAction Stop }
        catch {
            # A denied kill (another user, or a hardened/EDR-protected process)
            # must not abort the whole run: report it and skip the journey.
            $killed = $false
            Log ("  could not stop pid " + $ownerPid + ": " + $_.Exception.Message)
        }
        Start-Sleep -Seconds 3

        # Prove that data still flows after the node died: a fresh SUB on a new
        # subject plus a UDF publish must reach the cell. A cell that merely
        # stops updating would pass a "did not crash" check but fail a user.
        # The SUB goes on the wire asynchronously and Core is at-most-once (worse
        # right after a node kill, while the client reconnects), so the publish is
        # retried with a CHANGING payload (a repeated identical one would be
        # suppressed by SkipRepeatedMessages) until the cell shows it.
        $afterSubject = "$SubjectPrefix.fail"
        $failReady = Wait-Cell $ws 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) } -TimeoutMs 20000
        Invoke-Com { $ws.Range('A4').Formula = '=RTD(' + $rtdQuoted + '"SUB","' + $afterSubject + '","","' + $NatsUrl + '")' } | Out-Null
        $failValue = ''
        for ($attempt = 1; $attempt -le 8 -and $failValue -notmatch '^post-failover-\d+$'; $attempt++) {
            Invoke-Com { $ws.Range('A3').Value2 = ('post-failover-' + $attempt) } | Out-Null
            Invoke-Com { $ws.Range('A2').Formula = '=NatsUDFPublishNonVolatile("' + $afterSubject + '",A3,"' + $NatsUrl + '")' } | Out-Null
            $failValue = Wait-Cell $ws 'A4' { param($t) $t -match '^post-failover-\d+$' } -TimeoutMs 6000
        }
        Log "  after-kill RTD cell = $failValue"
        if ($killed) {
            Assert "Failover: data still flows after the node died" ($failValue -match '^post-failover-\d+$')
        } else {
            Skip "Failover: data still flows after the node died"
        }
        Close-Excel $excel
    } else {
        Skip "node kill (no local nats-server on port $port)"
    }
} else {
    Skip "node kill (non-local host or -SkipNodeKill)"
}

Log "=== E2E DONE: fails=$fails, skips=$skips, comRetries=$comRetries ==="
# Never leave the disposable second server behind, even on a failure path.
Stop-SecondServer
if ($fails -gt 0) { exit 1 }
exit 0
