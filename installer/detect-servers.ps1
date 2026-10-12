#Requires -Version 5.1
<#
  detect-servers.ps1 - lists NATS servers found on this machine, so the
  NATSExcel uninstaller can offer to remove the ones it created.

  Prints one line per candidate:  kind|label|present
  (present = 1 when found, 0 otherwise). The uninstaller parses this.

  Kinds:
    native-service  the "NATS" Windows service
    native-file     nats-server.exe under the install folder
    docker          a container named "natsexcel"
#>
[CmdletBinding()]
param(
  [string]$Dir = 'C:\NATSExcel\server',
  [string]$ContainerName = 'natsexcel',
  [string]$ServiceName = 'NATS'
)

$ErrorActionPreference = 'Continue'

function Have-Command([string]$name) {
  return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

# ---- Windows service ------------------------------------------------------
$svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
Write-Output ("native-service|NATS (Windows service '$ServiceName')|" + $(if ($svc) { '1' } else { '0' }))

# ---- nats-server.exe under the install folder -----------------------------
$exe = Join-Path $Dir 'nats-server.exe'
$hasFile = (Test-Path $exe)
Write-Output ("native-file|nats-server.exe (" + $Dir + ")|" + $(if ($hasFile) { '1' } else { '0' }))

# ---- Docker container -----------------------------------------------------
$dockerOk = $false
if (Have-Command docker) {
  try { & docker version --format '{{.Server.Version}}' *> $null; $dockerOk = ($LASTEXITCODE -eq 0) } catch { $dockerOk = $false }
}
$hasDocker = $false
if ($dockerOk) {
  try {
    & docker inspect --format '{{.Name}}' $ContainerName *> $null
    $hasDocker = ($LASTEXITCODE -eq 0)
  } catch { $hasDocker = $false }
}
Write-Output ("docker|NATS (Docker container '" + $ContainerName + "')|" + $(if ($hasDocker) { '1' } else { '0' }))
