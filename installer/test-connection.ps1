#Requires -Version 5.1
<#
  test-connection.ps1 - quick NATS handshake used by the installer wizard.

  Reads the INFO banner, sends PING and expects PONG. No dialogs: the wizard
  shows the message.

  Credentials are presented ONLY when the server asks for them
  (`auth_required` in INFO): -Token for a token, -User/-Password for
  user/password. The seed-based styles (NKey, JWT + seed, a creds file) and mTLS
  cannot be proven here - the server nonce must be SIGNED, and only the NATS
  client library can do that - so the wizard passes -Unprovable for them and the
  answer names `tools\NatsValidate --test`, the complete check.

  A value given on the command line is fine for a token; a password goes through
  -PasswordFile (a private temp file the wizard writes and deletes), so a
  password containing a quote or a space never has to survive argument quoting
  and never shows up in a process command line.

  The socket connect is BOUNDED: an unreachable host must not freeze the wizard
  on the OS SYN timeout.

  Exit codes are STABLE reason codes so a caller can show its own (localized)
  message instead of this script's text: 0 = PONG, 2 = unknown host (DNS),
  3 = timeout/unreachable, 4 = connection refused or other socket error,
  5 = the server rejected the credentials, 6 = unexpected reply, 7 = the
  credential style cannot be proven by a wire-level probe. The human detail is
  written after the first line; it is free-form (may be non-ASCII) and meant for
  the log, not for a UI label - a console redirect re-encodes it and would
  mangle accents.

  Examples:
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost localhost:4222
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost nats://host:4222 -Token <token>
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost localhost:4222 -User me -PasswordFile C:\Temp\pw.txt -Quiet
#>
[CmdletBinding()]
param(
  [string]$ServerHost = 'localhost',
  [int]$Port = 4222,
  [string]$User = '',
  [string]$Password = '',
  # The password read from a private temp file (takes precedence over -Password).
  [string]$PasswordFile = '',
  [string]$Token = '',
  # The wizard already knows a style this wire-level probe cannot prove (NKey,
  # JWT + seed, a creds file, mTLS): it asks for the unprovable answer instead of
  # sending a handshake that would be rejected for the wrong reason.
  [switch]$Unprovable,
  # Suppress the console header (the wizard redirects this output to a file and
  # only reads the exit code).
  [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$ConnectTimeoutMs = 5000
$ReadTimeoutMs = 4000

# Console header (see banner.ps1), unless the caller asked to stay quiet: the
# wizard's "Test connection" button redirects this output to a file and wants
# the exit code, not a banner. Failure to load it never breaks the check.
if (-not $Quiet) {
  try {
    $bannerPath = Join-Path $PSScriptRoot 'banner.ps1'
    if (Test-Path -LiteralPath $bannerPath) {
      . $bannerPath
      Write-NatsExcelBanner -Action ('Checking the NATS connection at ' + $ServerHost)
    }
  } catch { }
}

$hostName = $ServerHost
$port = $Port
if ($ServerHost -match '^\s*(?:nats://)?([^,:;\s]+)(?::(\d+))?') {
  $hostName = $Matches[1]
  if ($Matches[2]) { $port = [int]$Matches[2] }
}

# The file content is used verbatim (no Trim): a password may contain spaces.
if ($PasswordFile -ne '' -and (Test-Path -LiteralPath $PasswordFile)) {
  $Password = [System.IO.File]::ReadAllText($PasswordFile)
}

# A failed Task.Wait surfaces as an AggregateException, whose own message says
# nothing ("Exception calling Wait with 1 argument(s)"). The innermost exception
# carries the reason the user needs.
function Get-ErrorMessage($record) {
  $ex = $record.Exception
  while ($ex -and $ex.InnerException) { $ex = $ex.InnerException }
  if ($ex) { return $ex.Message }
  return $record.Exception.Message
}

# 2 = the NAME does not resolve; 4 = it resolves but nothing accepts the
# connection. The classification uses the SocketError CODE, never the message
# text: the text is localized by Windows, so a pattern match breaks per locale.
function Get-ErrorExitCode($record) {
  $hostNotFound = $false
  $e = $record.Exception
  while ($e -and -not $hostNotFound) {
    $se = $e -as [System.Net.Sockets.SocketException]
    if ($se -and ($se.SocketErrorCode -eq [System.Net.Sockets.SocketError]::HostNotFound -or
                  $se.SocketErrorCode -eq [System.Net.Sockets.SocketError]::NoData)) {
      $hostNotFound = $true
    }
    $e = $e.InnerException
  }
  if ($hostNotFound) { return 2 }
  return 4
}

function Connect-Nats([string]$h, [int]$p) {
  $client = New-Object System.Net.Sockets.TcpClient
  try {
    $connect = $client.ConnectAsync($h, $p)
    if (-not $connect.Wait($ConnectTimeoutMs)) {
      $client.Close()
      Write-Host ("no answer in {0} ms" -f $ConnectTimeoutMs)
      # 3 = no answer in time (unreachable host, firewall, wrong port)
      exit 3
    }
    # Surface a failed connect with the exception the caller classifies.
    $connect.GetAwaiter().GetResult() | Out-Null
    return $client
  } catch {
    $client.Close()
    throw
  }
}

# JSON string escaping for the CONNECT line (a token or password may carry a
# quote or a backslash; an unescaped one would produce an unparsable line).
function Js([string]$s) {
  if ($null -eq $s) { return '' }
  return $s.Replace('\', '\\').Replace('"', '\"')
}

try {
  if ($Unprovable) {
    Write-Host 'this credential style needs the NATS client (a signed challenge or TLS); use tools\NatsValidate --test'
    exit 7
  }

  $client = Connect-Nats $hostName $port
  $stream = $client.GetStream()
  $stream.ReadTimeout = $ReadTimeoutMs
  $buffer = New-Object byte[] 65536

  # Read the INFO banner (the server sends it immediately on connect).
  $n = $stream.Read($buffer, 0, $buffer.Length)
  $info = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $n)
  $authRequired = ($info -match '"auth_required"\s*:\s*true')

  # Present credentials only when the server asks for them.
  if ($authRequired) {
    $connectLine = '{"verbose":false,"pedantic":false,"lang":"natsexcel-installer","version":"0"'
    if ($Token -ne '') {
      $connectLine += ',"auth_token":"' + (Js $Token) + '"'
    } elseif ($User -ne '') {
      $connectLine += ',"user":"' + (Js $User) + '"'
      if ($Password -ne '') { $connectLine += ',"pass":"' + (Js $Password) + '"' }
    }
    $connectLine += '}'
    $payload = [System.Text.Encoding]::UTF8.GetBytes("CONNECT $connectLine`r`n")
    $stream.Write($payload, 0, $payload.Length)
    $stream.Flush()
  }

  $ping = [System.Text.Encoding]::ASCII.GetBytes("PING`r`n")
  $stream.Write($ping, 0, $ping.Length)
  $stream.Flush()
  Start-Sleep -Milliseconds 250
  $n2 = $stream.Read($buffer, 0, $buffer.Length)
  $reply = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $n2)
  $client.Close()

  if ($reply -match 'PONG') { Write-Host 'PONG'; exit 0 }
  if ($reply -match 'Authorization Violation') {
    # 5 = the server answered but rejected the credentials
    Write-Host 'the server rejected the credentials'
    exit 5
  }
  # 6 = something answered, but not with a PONG
  Write-Host ('unexpected reply: ' + $reply.Trim())
  exit 6
} catch {
  Write-Host ('no answer: ' + (Get-ErrorMessage $_))
  exit (Get-ErrorExitCode $_)
}
