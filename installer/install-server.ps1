#Requires -Version 5.1
<#
  install-server.ps1 - optional NATS server setup for the NATSExcel installer.

  NATS is a single self-contained nats-server.exe, so the native path needs no
  Docker, WSL or package manager. Docker remains an option for users who prefer
  it.

  Kinds:
    native      download the Windows release zip into -Dir, or reuse a
                nats-server already on PATH; write a starter nats.conf with a
                durable store_dir and start it
    docker      run the official "nats" image (JetStream on, durable volume)
    none        do nothing

  -JetStream (default on) enables JetStream (-js). A durable store_dir is always
  written so a restart does not drop streams (the server default is a temp path).

  Exits 0 on success (or when the port already answers), 3 when the chosen
  method is unavailable on this machine, 1 on a hard failure. The installer
  reports the exit code but never fails the add-in install because of it.

  Examples:
    powershell -ExecutionPolicy Bypass -File install-server.ps1 -Kind native -Port 4222 -Dir C:\NATSExcel\server
    powershell -ExecutionPolicy Bypass -File install-server.ps1 -Kind docker -Port 4222
    powershell -ExecutionPolicy Bypass -File install-server.ps1 -Kind native -NoJetStream
#>
[CmdletBinding()]
param(
  [ValidateSet('native', 'docker', 'none')]
  [string]$Kind = 'none',
  [int]$Port = 4222,
  [string]$Dir = '',
  [string]$ContainerName = 'natsexcel',
  [string]$Image = '',
  [switch]$NoJetStream,
  [string]$Version = '2.15.1'
)

$ErrorActionPreference = 'Stop'
$ExitUnsupported = 3

function Write-Step([string]$m) { Write-Host ("[nats-server] " + $m) }

function Invoke-Native([scriptblock]$block) {
  $prev = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try { & $block 2>&1 | ForEach-Object { $_ } } finally { $ErrorActionPreference = $prev }
}

function Test-Port([int]$p) {
  try {
    $c = New-Object System.Net.Sockets.TcpClient
    $c.Connect('127.0.0.1', $p)
    $c.Close()
    return $true
  } catch { return $false }
}

function Wait-Port([int]$p, [int]$seconds = 40) {
  $deadline = (Get-Date).AddSeconds($seconds)
  while ((Get-Date) -lt $deadline) {
    if (Test-Port $p) { return $true }
    Start-Sleep -Milliseconds 700
  }
  return (Test-Port $p)
}

# A real NATS handshake (read INFO, PING -> PONG), not just a socket connect.
function Test-Nats([string]$h, [int]$p) {
  try {
    $client = New-Object System.Net.Sockets.TcpClient
    $client.Connect($h, $p)
    $stream = $client.GetStream()
    $stream.ReadTimeout = 3000
    $buffer = New-Object byte[] 4096
    $null = $stream.Read($buffer, 0, $buffer.Length)   # INFO
    $ping = [System.Text.Encoding]::ASCII.GetBytes("PING`r`n")
    $stream.Write($ping, 0, $ping.Length)
    Start-Sleep -Milliseconds 200
    $n = $stream.Read($buffer, 0, $buffer.Length)
    $reply = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $n)
    $client.Close()
    return ($reply -match 'PONG')
  } catch { return $false }
}

# ---- native ---------------------------------------------------------------
function Write-StarterConf([string]$path, [string]$storeDir, [int]$p, [bool]$js) {
  $store = $storeDir -replace '\\', '/'
  $lines = @(
    '# NATSExcel starter configuration.',
    '# Example uses only localhost; adjust for your environment.',
    'server_name: natsexcel',
    "port: $p",
    'max_payload: 8MB',
    ''
  )
  if ($js) {
    $lines += @(
      'jetstream {',
      "  store_dir: `"$store`"   # durable: the default is a temp path",
      '}'
    )
  } else {
    $lines += '# jetstream is disabled (started without -js)'
  }
  # Write without a BOM: the NATS config parser rejects a leading byte-order mark.
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), $utf8NoBom)
}

function Native-Up() {
  if (-not $Dir) { Write-Step 'no -Dir given for the native kind.'; exit 1 }
  New-Item -ItemType Directory -Force -Path $Dir | Out-Null
  $exe = Join-Path $Dir 'nats-server.exe'

  if (-not (Test-Path $exe)) {
    # Reuse an existing nats-server on PATH when present.
    $onPath = (Get-Command nats-server -ErrorAction SilentlyContinue)
    if ($onPath) {
      # Copy it next to the config instead of running the PATH copy directly:
      # the Windows service helper always points at "$Dir\nats-server.exe", so
      # a PATH-only install would leave the service with nothing to run.
      Write-Step "using nats-server from PATH: $($onPath.Source)"
      Copy-Item $onPath.Source $exe -Force
      $source = $exe
    } else {
      $zip = Join-Path $Dir 'nats-server.zip'
      $url = "https://github.com/nats-io/nats-server/releases/download/v$Version/nats-server-v$Version-windows-amd64.zip"
      Write-Step "downloading $url"
      try { Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing }
      catch { Write-Step ("download failed: " + $_.Exception.Message); exit 1 }
      $tmp = Join-Path $Dir 'extract'
      if (Test-Path $tmp) { Remove-Item -Recurse -Force $tmp }
      Expand-Archive -Path $zip -DestinationPath $tmp -Force
      $found = Get-ChildItem -Path $tmp -Recurse -Filter 'nats-server.exe' | Select-Object -First 1
      if (-not $found) { Write-Step 'nats-server.exe was not found in the archive'; exit 1 }
      Copy-Item $found.FullName $exe -Force
      Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
      Remove-Item -Force $zip -ErrorAction SilentlyContinue
      $source = $exe
    }
  } else {
    $source = $exe
  }

  if (Test-Port $Port) {
    Write-Step "port $Port is already answering; leaving the existing server in place."
    exit 0
  }

  $conf = Join-Path $Dir 'nats.conf'
  if (-not (Test-Path $conf)) { Write-StarterConf $conf (Join-Path $Dir 'data') $Port (-not $NoJetStream) }

  $serverArgs = @('-c', $conf)
  Write-Step "starting $source on port $Port $($(if ($NoJetStream) { '(JetStream off)' } else { '(JetStream on)' }))"
  Start-Process -FilePath $source -ArgumentList $serverArgs -WorkingDirectory $Dir -WindowStyle Hidden

  if (-not (Wait-Port $Port 30)) { Write-Step "started but port $Port never answered"; exit 1 }
  if (Test-Nats '127.0.0.1' $Port) { Write-Step 'PONG - the server is answering the NATS protocol' }
  Write-Step "done: NATS is serving on localhost:$Port (config: $conf)"
  exit 0
}

# ---- docker ---------------------------------------------------------------
function Docker-Up() {
  $docker = Get-Command docker -ErrorAction SilentlyContinue
  if (-not $docker) { Write-Step 'Docker is not installed.'; exit $ExitUnsupported }
  try {
    $prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    & $docker.Source version --format '{{.Server.Version}}' *> $null
    $ErrorActionPreference = $prev
    if ($LASTEXITCODE -ne 0) { Write-Step 'Docker is installed but its engine is not running.'; exit $ExitUnsupported }
  } catch { Write-Step 'Docker engine is not reachable.'; exit $ExitUnsupported } finally { $ErrorActionPreference = 'Stop' }

  if (Test-Port $Port) { Write-Step "port $Port is already answering; leaving the existing server in place."; exit 0 }

  $img = if ($Image) { $Image } else { 'nats:latest' }
  $jsArgs = if ($NoJetStream) { @() } else { @('-js', '--store_dir', '/data') }
  Invoke-Native { & $docker.Source rm -f $ContainerName } | Out-Null
  Write-Step "starting $img on port $Port (auto-restarts on reboot)"
  $runOut = Invoke-Native { & $docker.Source run -d --name $ContainerName -p "${Port}:4222" --restart unless-stopped $img @jsArgs }
  if ($LASTEXITCODE -ne 0) { Write-Step ("docker run failed: " + (($runOut | Out-String).Trim())); exit 1 }
  if (-not (Wait-Port $Port 60)) { Write-Step "the container started but port $Port never answered"; exit 1 }
  Write-Step "done: $img is serving on localhost:$Port"
  exit 0
}

switch ($Kind) {
  'none'   { Write-Step 'no server requested.'; exit 0 }
  'native' { Native-Up }
  'docker' { Docker-Up }
}
