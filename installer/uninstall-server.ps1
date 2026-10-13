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
    Write-Step "stopping any nats-server.exe under $Dir and removing it"
    # Scope by path: a bare Get-Process nats-server would stop the user's OWN
    # server running elsewhere.
    Get-Process nats-server -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -and $_.Path -like "$Dir*" } |
        Stop-Process -Force -ErrorAction SilentlyContinue
    if (Test-Path $Dir) { Remove-Item -LiteralPath $Dir -Recurse -Force -ErrorAction SilentlyContinue }
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
