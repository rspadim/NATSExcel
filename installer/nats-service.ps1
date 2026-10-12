#Requires -Version 5.1
<#
  nats-service.ps1 - install/remove the native nats-server as a Windows service.

  This is the ONLY step of the NATSExcel installer that needs admin: the add-in
  itself installs per-user. The installer never blocks on it; run this by hand
  (elevated) if you want NATS to start at boot.

  Actions:
    Install   create the service (idempotent), pointing at nats-server -c nats.conf
    Start     start it
    Stop      stop it
    Remove    stop and delete it
    Status    print the service state

  Examples (elevated):
    powershell -ExecutionPolicy Bypass -File nats-service.ps1 -Action Install -Dir C:\NATSExcel\server
    powershell -ExecutionPolicy Bypass -File nats-service.ps1 -Action Status
#>
[CmdletBinding()]
param(
  [ValidateSet('Install', 'Start', 'Stop', 'Remove', 'Status')]
  [string]$Action = 'Status',
  [string]$Dir = 'C:\NATSExcel\server',
  [string]$ServiceName = 'NATS'
)

$ErrorActionPreference = 'Stop'

function Is-Admin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

$exe = Join-Path $Dir 'nats-server.exe'
$conf = Join-Path $Dir 'nats.conf'

switch ($Action) {
  'Status' {
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($svc) { Write-Host ("[nats-service] " + $svc.Name + " : " + $svc.Status) }
    else { Write-Host '[nats-service] not installed' }
    exit 0
  }
  'Install' {
    if (-not (Is-Admin)) { Write-Host '[nats-service] this action needs an elevated (Administrator) shell.'; exit 2 }
    if (-not (Test-Path $exe)) { Write-Host "[nats-service] not found: $exe"; exit 1 }
    $binPath = '"' + $exe + '" -c "' + $conf + '"'
    $existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($existing) {
      Write-Host "[nats-service] updating the existing '$ServiceName' service"
      & sc.exe config $ServiceName binPath= $binPath start= auto | Out-Null
    } else {
      Write-Host "[nats-service] creating the '$ServiceName' service"
      & sc.exe create $ServiceName binPath= $binPath start= auto DisplayName= 'NATS Server (NATSExcel)' | Out-Null
    }
    & sc.exe start $ServiceName | Out-Null
    Write-Host "[nats-service] service '$ServiceName' is installed and started"
    exit 0
  }
  'Start' {
    if (-not (Is-Admin)) { Write-Host '[nats-service] this action needs an elevated shell.'; exit 2 }
    & sc.exe start $ServiceName | Out-Null
    Write-Host "[nats-service] start requested"; exit 0
  }
  'Stop' {
    if (-not (Is-Admin)) { Write-Host '[nats-service] this action needs an elevated shell.'; exit 2 }
    & sc.exe stop $ServiceName | Out-Null
    Write-Host "[nats-service] stop requested"; exit 0
  }
  'Remove' {
    if (-not (Is-Admin)) { Write-Host '[nats-service] this action needs an elevated shell.'; exit 2 }
    & sc.exe stop $ServiceName 2>&1 | Out-Null
    & sc.exe delete $ServiceName 2>&1 | Out-Null
    Write-Host "[nats-service] service '$ServiceName' removed"; exit 0
  }
}
