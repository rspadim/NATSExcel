#Requires -Version 5.1
<#
    Start-Cluster.ps1 - brings up a disposable NATS cluster for the test layers.

    Native nats-server.exe (no Docker, no WSL). Each node gets its own durable
    JetStream store so a restart keeps the streams. The default is 1 node; pass
    -Nodes 3 for the cluster/failover tests.

    Hosts are always loopback: this file must never carry a real endpoint.
    A real server is passed to the test programs as a command-line argument.

    Examples:
      powershell -File Start-Cluster.ps1                 # 1 node on 4222
      powershell -File Start-Cluster.ps1 -Nodes 3        # 4222/4223/4224
      powershell -File Start-Cluster.ps1 -Nodes 3 -JetStream:$false
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 5)]
    [int]$Nodes = 1,
    [int]$BasePort = 4222,
    [int]$BaseClusterPort = 6222,
    [string]$WorkDir,
    [string]$ServerExe,
    [string]$Version = '2.15.1',
    [switch]$JetStream = $true,
    [switch]$Force   # kill an existing test cluster first
)

$ErrorActionPreference = 'Stop'

function Write-Step([string]$m) { Write-Host ("[cluster] " + $m) }

if (-not $WorkDir) {
    $WorkDir = Join-Path ([System.IO.Path]::GetTempPath()) 'natsexcel-cluster'
}

# ---- locate (or fetch) nats-server.exe ------------------------------------
function Resolve-ServerExe {
    param([string]$Explicit, [string]$Version, [string]$WorkDir)
    if ($Explicit) {
        if (-not (Test-Path $Explicit)) { throw "nats-server not found: $Explicit" }
        return (Resolve-Path $Explicit).Path
    }
    $onPath = Get-Command nats-server -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }

    $dir = Join-Path $WorkDir 'server'
    $exe = Join-Path $dir 'nats-server.exe'
    if (Test-Path $exe) { return $exe }

    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $zip = Join-Path $dir 'nats-server.zip'
    $url = "https://github.com/nats-io/nats-server/releases/download/v$Version/nats-server-v$Version-windows-amd64.zip"
    Write-Step "downloading $url"
    Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
    $tmp = Join-Path $dir 'extract'
    if (Test-Path $tmp) { Remove-Item -Recurse -Force $tmp }
    Expand-Archive -Path $zip -DestinationPath $tmp -Force
    $found = Get-ChildItem $tmp -Recurse -Filter 'nats-server.exe' | Select-Object -First 1
    if (-not $found) { throw 'nats-server.exe was not in the archive' }
    Copy-Item $found.FullName $exe -Force
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    Remove-Item -Force $zip -ErrorAction SilentlyContinue
    return $exe
}

# ---- stop a previous run ---------------------------------------------------
function Stop-Previous {
    param([int]$Nodes, [int]$BasePort, [string]$WorkDir)
    for ($i = 0; $i -lt $Nodes; $i++) {
        $port = $BasePort + $i
        $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($conn) {
            $proc = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
            if ($proc -and $proc.Path -and $proc.Path -like "*$WorkDir*") {
                Write-Step "stopping the previous test node on $port (pid $($proc.Id))"
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

if ($Force) { Stop-Previous -Nodes $Nodes -BasePort $BasePort -WorkDir $WorkDir }

$exe = Resolve-ServerExe -Explicit $ServerExe -Version $Version -WorkDir $WorkDir
Write-Step "server: $exe"
& $exe --version

$utf8 = New-Object System.Text.UTF8Encoding($false)
$started = @()

for ($i = 0; $i -lt $Nodes; $i++) {
    $n = $i + 1
    $port = $BasePort + $i
    $cport = $BaseClusterPort + $i
    $nodeDir = Join-Path $WorkDir "n$n"
    New-Item -ItemType Directory -Force -Path $nodeDir | Out-Null

    if (Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue) {
        Write-Step "port $port already answered; leaving it alone"
        continue
    }

    $lines = @(
        "server_name: n$n",
        "port: $port"
    )
    if ($Nodes -gt 1) {
        $others = (0..($Nodes - 1) | Where-Object { $_ -ne $i } |
            ForEach-Object { "nats-route://127.0.0.1:$($BaseClusterPort + $_)" }) -join ', '
        $lines += @(
            'cluster {',
            '  name: NATSEXCEL',
            "  listen: 127.0.0.1:$cport",
            "  routes: [ $others ]",
            '}'
        )
    }
    if ($JetStream) {
        $lines += @('jetstream {', "  store_dir: $((Join-Path $nodeDir 'data') -replace '\\','/')", '}')
    }
    $conf = Join-Path $nodeDir 'nats.conf'
    # Write without a BOM: the NATS parser rejects a leading byte-order mark.
    [System.IO.File]::WriteAllText($conf, (($lines -join "`r`n") + "`r`n"), $utf8)

    Write-Step "starting node $n on $port (JetStream: $([bool]$JetStream))"
    Start-Process -FilePath $exe -ArgumentList '-c', $conf -WorkingDirectory $nodeDir -WindowStyle Hidden
    $started += $port
}

# ---- wait for readiness (a real handshake, not just a socket) --------------
function Test-Nats([int]$port) {
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $client.Connect('127.0.0.1', $port)
        $stream = $client.GetStream()
        $stream.ReadTimeout = 3000
        $buffer = New-Object byte[] 4096
        $null = $stream.Read($buffer, 0, $buffer.Length)   # INFO
        $ping = [System.Text.Encoding]::ASCII.GetBytes("PING`r`n")
        $stream.Write($ping, 0, $ping.Length)
        Start-Sleep -Milliseconds 150
        $n2 = $stream.Read($buffer, 0, $buffer.Length)
        $reply = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $n2)
        $client.Close()
        return ($reply -match 'PONG')
    } catch { return $false }
}

for ($i = 0; $i -lt $Nodes; $i++) {
    $port = $BasePort + $i
    $deadline = (Get-Date).AddSeconds(30)
    $ok = $false
    while ((Get-Date) -lt $deadline) {
        if (Test-Nats $port) { $ok = $true; break }
        Start-Sleep -Milliseconds 500
    }
    if (-not $ok) { throw "node on port $port did not answer within 30s" }
    Write-Step "node on $port is ready"
}

$urls = (0..($Nodes - 1) | ForEach-Object { "nats://127.0.0.1:$($BasePort + $_)" }) -join ','
Write-Step "cluster ready: $urls"
Write-Output $urls
