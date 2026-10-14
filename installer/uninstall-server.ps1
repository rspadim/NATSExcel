#Requires -Version 5.1
<#
  uninstall-server.ps1 - removes a NATS server chosen in the uninstaller.
  Only the server(s) the user ticked are removed; nothing else.

  Kind:
    native-service  stop and delete the Windows service (needs admin)
    native-file     remove the downloaded nats-server.exe + its data dir
    docker          remove the "natsexcel" container

  Exit 0 on success or when there was nothing to do, 1 on a hard failure.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidateSet('native-service', 'native-file', 'docker')]
  [string]$Kind,
  [string]$Dir = 'C:\NATSExcel\server',
  [string]$ContainerName = 'natsexcel',
  [string]$ServiceName = 'NATS'
)

$ErrorActionPreference = 'Continue'
function Write-Step([string]$m) { Write-Host ("[nats-uninstall] " + $m) }

switch ($Kind) {
  'native-service' {
    Write-Step "removing the '$ServiceName' Windows service"
    & sc.exe stop $ServiceName 2>&1 | Out-Null
    & sc.exe delete $ServiceName 2>&1 | Out-Null
    exit 0
  }
  'native-file' {
    Write-Step "stopping any nats-server.exe started from $Dir and removing it"

    # Attribute the process by the PORT it serves, not by $proc.Path: a
    # nats-server started with Start-Process exposes an EMPTY .Path (and an empty
    # Win32_Process.CommandLine / ExecutablePath) on a hardened host, so a
    # path-based match silently finds nothing and Remove-Item then fails on the
    # locked exe.
    #
    # The port MUST come from this folder's own nats.conf: falling back to a
    # default 4222 would stop whatever server the user happens to run on that
    # port, in another folder, which is not ours to kill. No conf -> no port
    # attribution (the path match below still applies).
    $port = 0
    $conf = Join-Path $Dir 'nats.conf'
    if (Test-Path $conf) {
      $m = Select-String -Path $conf -Pattern '^\s*port\s*:\s*(\d+)' -ErrorAction SilentlyContinue | Select-Object -First 1
      if ($m) { $port = [int]$m.Matches[0].Groups[1].Value }
    }

    $stopped = @()
    if ($port -gt 0) {
      $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
      if ($conn) {
        $proc = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
        if ($proc -and $proc.ProcessName -like 'nats-server*') {
          Write-Step "stopping the server from this folder on port $port (pid $($proc.Id))"
          Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
          $stopped += $proc.Id
          # Wait for the port and the file lock to be released before removing.
          $deadline = (Get-Date).AddSeconds(5)
          while ((Get-Date) -lt $deadline) {
            if (-not (Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)) { break }
            Start-Sleep -Milliseconds 100
          }
        }
      }
    }
    else {
      Write-Step "no nats.conf in ${Dir}: not attributing a server by port (a default port could belong to another install)"
    }
    # Secondary: any nats-server whose path really is under $Dir.
    Get-Process nats-server -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -and $_.Path -like "$Dir*" -and $stopped -notcontains $_.Id } |
        ForEach-Object {
          Write-Step "stopping nats-server.exe at $($_.Path) (pid $($_.Id))"
          Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }

    if (Test-Path $Dir) {
      Remove-Item -LiteralPath $Dir -Recurse -Force -ErrorAction SilentlyContinue
      if (Test-Path $Dir) {
        # The removal failing means something still holds the exe (a server we
        # could not attribute, or an editor). Say so instead of exiting 0 as if
        # the folder were gone.
        Write-Step "WARNING: $Dir could not be fully removed (a process may still hold nats-server.exe)"
        exit 1
      }
    }
    Write-Step "removed $Dir"
    exit 0
  }
  'docker' {
    $docker = Get-Command docker -ErrorAction SilentlyContinue
    if (-not $docker) { Write-Step 'Docker not found; nothing to do.'; exit 0 }
    Write-Step "removing container '$ContainerName'"
    & docker rm -f $ContainerName 2>&1 | Write-Host
    exit 0
  }
}
