#Requires -Version 5.1
<#
  test-connection.ps1 - quick NATS handshake used by the installer wizard.
  Reads the INFO banner, sends PING and expects PONG. Prints the result and
  exits 0 when the server answers, 1 otherwise. No dialogs: the wizard shows
  the message.

  Examples:
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost localhost:4222
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost host:4222 -Token <token>
#>
[CmdletBinding()]
param(
  [string]$ServerHost = 'localhost',
  [int]$Port = 4222,
  [string]$User = '',
  [string]$Password = '',
  [string]$Token = ''
)

$ErrorActionPreference = 'Stop'

$hostName = $ServerHost
$port = $Port
if ($ServerHost -match '^\s*(?:nats://)?([^,:;\s]+)(?::(\d+))?') {
  $hostName = $Matches[1]
  if ($Matches[2]) { $port = [int]$Matches[2] }
}

try {
  $client = New-Object System.Net.Sockets.TcpClient
  $client.Connect($hostName, $port)
  $stream = $client.GetStream()
  $stream.ReadTimeout = 4000
  $buffer = New-Object byte[] 65536

  # Read the INFO banner (the server sends it immediately on connect).
  $n = $stream.Read($buffer, 0, $buffer.Length)
  $info = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $n)
  $authRequired = ($info -match '"auth_required"\s*:\s*true')

  # Present credentials only when the server asks for them.
  if ($authRequired) {
    $connectLine = '{"verbose":false,"pedantic":false,"lang":"natsexcel-installer","version":"0"'
    if ($Token -ne '') {
      $connectLine += ',"auth_token":"' + $Token + '"'
    } elseif ($User -ne '') {
      $connectLine += ',"user":"' + $User + '"'
      if ($Password -ne '') { $connectLine += ',"pass":"' + $Password + '"' }
    }
    $connectLine += '}'
    $payload = [System.Text.Encoding]::ASCII.GetBytes("CONNECT $connectLine`r`n")
    $stream.Write($payload, 0, $payload.Length)
  }

  $ping = [System.Text.Encoding]::ASCII.GetBytes("PING`r`n")
  $stream.Write($ping, 0, $ping.Length)
  Start-Sleep -Milliseconds 250
  $n2 = $stream.Read($buffer, 0, $buffer.Length)
  $reply = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $n2)
  $client.Close()

  if ($reply -match 'PONG') { Write-Host 'PONG'; exit 0 }
  if ($reply -match 'Authorization Violation') { Write-Host 'AUTH failed: the server rejected the credentials'; exit 1 }
  Write-Host ("unexpected reply: " + ($reply.Trim()))
  exit 1
} catch {
  Write-Host ('no answer: ' + $_.Exception.Message)
  exit 1
}
