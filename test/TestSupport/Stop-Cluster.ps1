#Requires -Version 5.1
<#
    Stop-Cluster.ps1 - tears down the disposable test cluster started by
    Start-Cluster.ps1. Only stops nats-server processes that live under the test
    work directory, so a developer's own server is never touched.

    Examples:
      powershell -File Stop-Cluster.ps1
      powershell -File Stop-Cluster.ps1 -Remove    # also delete the data dirs
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 5)]
    [int]$Nodes = 3,
    [int]$BasePort = 4222,
    [string]$WorkDir,
    [switch]$Remove
)

$ErrorActionPreference = 'Continue'
function Write-Step([string]$m) { Write-Host ("[cluster] " + $m) }

if (-not $WorkDir) {
    $WorkDir = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-cluster'
}

$stopped = 0

# Is the nats-server on $port one of OURS? A path check alone is unreliable: a
# nats-server started with Start-Process exposes an EMPTY .Path, so the old
# `$proc.Path -like "*$WorkDir*"` never matched and the node survived every
# Stop-Cluster (leaking servers across runs). A nats.conf under OUR work dir that
# declares this port is the reliable signal, and it still refuses to touch a
# developer's own server that happens to listen on the same port.
function Test-Ours([int]$port, $proc) {
    if (-not $proc) { return $false }
    if ($proc.ProcessName -notlike 'nats-server*') { return $false }
    if ($proc.Path -and $proc.Path -like "$WorkDir*") { return $true }
    foreach ($conf in (Get-ChildItem -Path $WorkDir -Recurse -Filter 'nats.conf' -ErrorAction SilentlyContinue)) {
        $m = Select-String -Path $conf.FullName -Pattern '^\s*port\s*:\s*(\d+)' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($m -and [int]$m.Matches[0].Groups[1].Value -eq $port) { return $true }
    }
    return $false
}

for ($i = 0; $i -lt $Nodes; $i++) {
    $port = $BasePort + $i
    $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if (-not $conn) { continue }
    $proc = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
    if (Test-Ours $port $proc) {
        Write-Step "stopping the test node on $port (pid $($proc.Id))"
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        $stopped++
        # Wait for the port to free so a following run can bind it.
        $deadline = (Get-Date).AddSeconds(5)
        while ((Get-Date) -lt $deadline) {
            if (-not (Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)) { break }
            Start-Sleep -Milliseconds 100
        }
    } else {
        Write-Step "port $port is answered by something else; leaving it alone"
    }
}
Write-Step "stopped $stopped node(s)"

if ($Remove -and (Test-Path $WorkDir)) {
    Write-Step "removing $WorkDir"
    Remove-Item -Recurse -Force $WorkDir -ErrorAction SilentlyContinue
}
