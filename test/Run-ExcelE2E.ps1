#Requires -Version 5.1
<#
    Run-ExcelE2E.ps1 - the Excel end-to-end test for NATSExcel.

    Loads the packed XLL into a hidden Excel and exercises what a user actually
    does: UDF write/read, RTD push, JetStream/KV paths, cluster failover (kill
    the node the add-in is on), the multi-workbook case and an Excel restart.

    Requires: Excel, a built XLL, and a running NATS server (start one with
    test\TestSupport\Start-Cluster.ps1). Local hosts only for the destructive
    steps; the node kill is skipped automatically for a non-local URL.

    Every wait is CONDITION-BASED (a helper polls until the cell changes or the
    timeout expires), so a healthy run finishes as fast as Excel allows and a
    failing one still fails at a known bound. No fixed "sleep 5s and hope".

    Hosts default to loopback. Never commit a real endpoint: pass it as -NatsUrl.

    Examples:
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
      powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1 -NatsUrl "nats://127.0.0.1:4222" -Fast
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
    # Shortens the per-assertion bound (a smoke run on a known-good machine).
    [switch]$Fast
)

$ErrorActionPreference = 'Stop'

if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
$RepoRoot = (Resolve-Path $RepoRoot).Path
if (-not $Xll) { $Xll = Join-Path $RepoRoot 'bin\Release\net48\publish\NATSExcel64-packed.xll' }

# Bound per assertion: enough for a slow COM call, short enough not to waste
# minutes. A -Fast run trusts the machine and uses a tight bound.
$BoundMs = if ($Fast) { 4000 } else { 12000 }
$PollMs = 250

$log = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-e2e.log'
Remove-Item $log -ErrorAction SilentlyContinue
function Log([string]$m) { Write-Host $m; Add-Content -Path $log -Value $m }

$script:fails = 0
$script:comRetries = 0
function Assert([string]$name, $cond) {
    Log ("{0,-52} {1}" -f $name, $(if ($cond) { 'PASS' } else { 'FAIL' }))
    if (-not $cond) { $script:fails++ }
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

$isLocal = $NatsUrl -match '127\.0\.0\.1|localhost|\[::1\]'
$pub = "$SubjectPrefix.core"

Log "=== NATSExcel E2E ==="
Log "xll    : $Xll"
Log "nats   : $NatsUrl  (local=$isLocal, bound=${BoundMs}ms)"
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
Invoke-Com { $ws.Range('A2').Formula = '=NatsUDFPublish("' + $pub + '","hello","' + $NatsUrl + '")' } | Out-Null
Invoke-Com { $ws.Range('A3').Formula = '=NatsUDFSubjectLatest("' + $pub + '","","' + $NatsUrl + '")' } | Out-Null
Invoke-Com { $excel.CalculateFullRebuild() } | Out-Null

$version = Wait-Cell $ws 'A1' { param($t) $t -match '^\d+\.\d+' }
$marker = Wait-Cell $ws 'A2' { param($t) $t -match 'OK|Error|#' }
$latest = Wait-Cell $ws 'A3' { param($t) $t -eq 'hello' -or $t -eq '(null)' }
Log "  version=$version publish=$marker latest=$latest"
Assert "UDF: the XLL answers (packing OK)" ($version -match '^\d+\.\d+')
Assert "UDF: publish returns a marker" ($marker -match 'OK|Error|#')

# RTD SUB: use a DEDICATED subject and a NON-VOLATILE publisher. A volatile
# publish cell (A2 above) is recalculated on every calculation and republishes,
# which would overwrite the RTD cell with a stale value - a test artifact, not a
# product bug. The NonVolatile twin fires only when its argument cell changes.
$rtdSubject = "$SubjectPrefix.rtd"
Invoke-Com { $ws.Range('A5').Formula = '=RTD("NatsRtd",,"SUB","' + $rtdSubject + '","","' + $NatsUrl + '")' } | Out-Null
$rtdReady = Wait-Cell $ws 'A5' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
Invoke-Com { $ws.Range('C1').Value2 = 'rtd-1' } | Out-Null
Invoke-Com { $ws.Range('A6').Formula = '=NatsUDFPublishNonVolatile("' + $rtdSubject + '",C1,"' + $NatsUrl + '")' } | Out-Null
$rtdFirst = Wait-Cell $ws 'A5' { param($t) $t -eq 'rtd-1' }
# Change the argument cell: a NonVolatile write re-runs, pushing the new value.
Invoke-Com { $ws.Range('C1').Value2 = 'rtd-2' } | Out-Null
$rtdAfter = Wait-Cell $ws 'A5' { param($t) $t -eq 'rtd-2' }
Log "  RTD ready=$rtdReady first=$rtdFirst after=$rtdAfter"
Assert "RTD: SUB did not error" ($rtdReady -notmatch '#ERROR')
Assert "RTD: the first push reached the cell" ($rtdFirst -eq 'rtd-1')
Assert "RTD: a later push updated the cell" ($rtdAfter -eq 'rtd-2')

# JetStream + KV: each must return a value or a NAMED error, never a hang/blank.
Invoke-Com { $ws.Range('A8').Formula = '=NatsUDFKVGet("NoSuchBucket","k","' + $NatsUrl + '")' } | Out-Null
Invoke-Com { $ws.Range('A9').Formula = '=NatsUDFJSStreams("' + $NatsUrl + '")' } | Out-Null
$kv = Wait-Cell $ws 'A8' { param($t) -not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch '^#' }
$streams = Wait-Cell $ws 'A9' { param($t) -not [string]::IsNullOrWhiteSpace($t) }
Log "  kv(missing bucket)=$($kv.Substring(0,[Math]::Min(90,$kv.Length)))  jsStreams=$streams"
Assert "KV: a missing bucket is a named error" ($kv -match 'Error|not found')
Assert "JetStream: the streams call answers" (-not [string]::IsNullOrWhiteSpace($streams))

# ---------------------------------------------------------------------------
# 2. Multi-workbook: a copy closing must not disturb the original
# ---------------------------------------------------------------------------
if ($isLocal -and -not $SkipMultiWorkbook) {
    $copy = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-e2e-copy.xlsx'
    Invoke-Com { $wb.SaveCopyAs($copy) } | Out-Null
    $excel2 = New-HiddenExcel
    $wb2 = Invoke-Com { $excel2.Workbooks.Open($copy) }
    $null = Invoke-Com { $wb2.Worksheets.Item(1).Range('A5').Text }   # force a calc
    Invoke-Com { $wb2.Close($false) } | Out-Null
    try { $excel2.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) | Out-Null

    # The RTD subject above is dedicated, so a publish cannot be overwritten by a
    # recalculated volatile cell. Change the NonVolatile publisher's argument.
    Invoke-Com { $ws.Range('C1').Value2 = 'after-copy' } | Out-Null
    $afterCopy = Wait-Cell $ws 'A5' { param($t) $t -eq 'after-copy' }
    Log "  RTD after the copy closed = $afterCopy"
    Assert "Multi-workbook: the original keeps receiving" ($afterCopy -eq 'after-copy')
    Remove-Item $copy -ErrorAction SilentlyContinue
} else {
    Log "  multi-workbook SKIPPED"
}

Close-Excel $excel
Log "  (steps 1-2 done, fails=$fails, comRetries=$comRetries)"

# ---------------------------------------------------------------------------
# 3. Cluster failover: kill the node the add-in is on, expect recovery
# ---------------------------------------------------------------------------
if ($isLocal -and -not $SkipNodeKill) {
    $port = ([Uri]$NatsUrl).Port
    $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
    $ownerPid = if ($conn) { $conn.OwningProcess } else { $null }
    $proc = if ($ownerPid) { Get-Process -Id $ownerPid -ErrorAction SilentlyContinue } else { $null }

    if ($proc -and $proc.ProcessName -eq 'nats-server') {
        $excel = New-HiddenExcel
        $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
        Invoke-Com { $ws.Range('A1').Formula = '=RTD("NatsRtd",,"SUB","' + $pub + '","","' + $NatsUrl + '")' } | Out-Null
        $ready = Wait-Cell $ws 'A1' { param($t) $t -ne '(ConnectData)' -and -not [string]::IsNullOrWhiteSpace($t) }
        Invoke-Com { $ws.Range('A2').Formula = '=NatsUDFConnectionState("' + $NatsUrl + '")' } | Out-Null
        $stateBefore = Wait-Cell $ws 'A2' { param($t) -not [string]::IsNullOrWhiteSpace($t) }

        Log "  killing the connected node (pid $ownerPid) to force failover"
        Stop-Process -Id $ownerPid -Force
        Start-Sleep -Seconds 3

        # Prove that data still flows after the node died: a fresh SUB on a new
        # subject plus a UDF publish must reach the cell. A cell that merely
        # stops updating would pass a "did not crash" check but fail a user.
        $afterSubject = "$SubjectPrefix.fail"
        $failReady = Wait-Cell $ws 'A1' { param($t) -not [string]::IsNullOrWhiteSpace($t) } -TimeoutMs 20000
        Invoke-Com { $ws.Range('A3').Value2 = 'post-failover' } | Out-Null
        Invoke-Com { $ws.Range('A2').Formula = '=NatsUDFPublishNonVolatile("' + $afterSubject + '",A3,"' + $NatsUrl + '")' } | Out-Null
        Invoke-Com { $ws.Range('A4').Formula = '=RTD("NatsRtd",,"SUB","' + $afterSubject + '","","' + $NatsUrl + '")' } | Out-Null
        $failValue = Wait-Cell $ws 'A4' { param($t) $t -eq 'post-failover' } -TimeoutMs 25000
        Log "  after-kill RTD cell = $failValue"
        Assert "Failover: data still flows after the node died" ($failValue -eq 'post-failover')
        Close-Excel $excel
    } else {
        Log "  node kill SKIPPED: no local nats-server on port $port"
    }
} else {
    Log "  node kill SKIPPED (non-local host or -SkipNodeKill)"
}

# ---------------------------------------------------------------------------
# 4. Excel restart: a fresh instance must load the XLL and work again
# ---------------------------------------------------------------------------
$excel = New-HiddenExcel
$wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
Invoke-Com { $ws.Range('A1').Formula = '=NatsUDFVersion()' } | Out-Null
$restart = Wait-Cell $ws 'A1' { param($t) $t -match '^\d+\.\d+' }
Log "  after 'restart' version=$restart"
Assert "Excel restart: the XLL reloads and answers" ($restart -match '^\d+\.\d+')
if (-not $KeepExcelOpen) { Close-Excel $excel } else { $excel.Visible = $true }

Log "=== E2E DONE: fails=$fails, comRetries=$comRetries ==="
if ($fails -gt 0) { exit 1 }
exit 0
