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
for ($i = 0; $i -lt $Nodes; $i++) {
    $port = $BasePort + $i
    $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if (-not $conn) { continue }
    $proc = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
    if ($proc -and $proc.ProcessName -eq 'nats-server' -and $proc.Path -like "*$WorkDir*") {
        Write-Step "stopping the test node on $port (pid $($proc.Id))"
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        $stopped++
    } else {
        Write-Step "port $port is answered by something else; leaving it alone"
    }
}
Write-Step "stopped $stopped node(s)"

if ($Remove -and (Test-Path $WorkDir)) {
    Write-Step "removing $WorkDir"
    Remove-Item -Recurse -Force $WorkDir -ErrorAction SilentlyContinue
}
