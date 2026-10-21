#Requires -Version 5.1
<#
  verify-connection.ps1 - end-of-install check for NATSExcel.

  * Tests that a NATS server answers the protocol (INFO + PING -> PONG).
  * With -OpenExcel, opens a visible Excel workbook that loads the installed XLL
    and shows the connection status with a friendly message.

  Exit codes: 0 = NATS answered (and the sheet opened if asked), 1 = no answer.

  Examples:
    powershell -ExecutionPolicy Bypass -File verify-connection.ps1 -AddInDir "C:\NATSExcel" -ServerHost "localhost:4222"
    powershell -ExecutionPolicy Bypass -File verify-connection.ps1 -AddInDir "C:\NATSExcel" -ServerHost "localhost:4222" -OpenExcel
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$AddInDir,
  [string]$ServerHost = 'localhost',
  [int]$Port = 4222,
  [string]$Token = '',
  [switch]$OpenExcel,
  [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms | Out-Null

# Console header (see banner.ps1): this script runs in a visible window during
# the install, and an unidentified window is what makes a user suspect a virus.
# Failure to load it must never break the check itself.
try {
  $bannerPath = Join-Path $PSScriptRoot 'banner.ps1'
  if (Test-Path -LiteralPath $bannerPath) {
    . $bannerPath
    if (-not $Quiet) {
      Write-NatsExcelBanner -Action 'Installing and checking NATSExcel' -InstallDir $AddInDir
    }
  }
} catch { }

function Show([string]$text, [string]$title = 'NATSExcel') {
  if (-not $Quiet) { [System.Windows.Forms.MessageBox]::Show($text, $title) | Out-Null }
  Write-Host $text
}

# Host may already be "host:port" (or a nats:// URL).
$raw = $ServerHost.Trim()
if ($raw -match '^\s*(?:nats://)?([^,:;\s]+)(?::(\d+))?') {
  $serverHost = $Matches[1]
  $serverPort = $Port
  if ($Matches[2]) { $serverPort = [int]$Matches[2] }
} else {
  $serverHost = $ServerHost
  $serverPort = $Port
}
$target = "$serverHost`:$serverPort"

function Test-Nats([string]$h, [int]$p, [string]$tok) {
  try {
    $client = New-Object System.Net.Sockets.TcpClient
    $client.Connect($h, $p)
    $stream = $client.GetStream()
    $stream.ReadTimeout = 3000
    $buffer = New-Object byte[] 4096
    $info = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $stream.Read($buffer, 0, $buffer.Length))
    # Present the token only when the server asks for credentials.
    if (($info -match '"auth_required"\s*:\s*true') -and ($tok -ne '')) {
      $connect = 'CONNECT {"verbose":false,"pedantic":false,"auth_token":"' + $tok + '"}' + "`r`n"
      $w = [System.Text.Encoding]::ASCII.GetBytes($connect)
      $stream.Write($w, 0, $w.Length)
    }
    $payload = [System.Text.Encoding]::ASCII.GetBytes("PING`r`n")
    $stream.Write($payload, 0, $payload.Length)
    Start-Sleep -Milliseconds 200
    $read = $stream.Read($buffer, 0, $buffer.Length)
    $response = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
    $client.Close()
    return ($response -match 'PONG')
  } catch {
    return $false
  }
}

Write-Host "[verify] testing NATS at $target ..."
$ok = Test-Nats $serverHost $serverPort $Token

if (-not $ok) {
  Show ("Could not reach NATS at $target.`r`n`r`n" +
        "- Check that the server is running (the NATS service, nats-server.exe, or docker ps).`r`n" +
        "- If it runs on another host/port, adjust NATSExcel.json.") 'NATSExcel - connection'
  exit 1
}

Write-Host "[verify] NATS answered at $target"

if (-not $OpenExcel) {
  Show ("All good!`r`n`r`nNATS answered at $target.`r`n" +
        "Open Excel, register the XLL (File > Options > Add-ins > Go... > Browse...) " +
        'and use, for example, =RTD("NatsRtd",,"SUB","test.>") in a cell.') 'NATSExcel - installed'
  exit 0
}

# ---- Open a sample sheet in Excel ------------------------------------------
$xll64 = Join-Path $AddInDir 'NATSExcel64-packed.xll'
$xll32 = Join-Path $AddInDir 'NATSExcel-packed.xll'
$xll = if (Test-Path $xll64) { $xll64 } elseif (Test-Path $xll32) { $xll32 } else { $null }
if (-not $xll) {
  Show ("Could not find the XLL in $AddInDir.") 'NATSExcel'
  exit 0
}

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $registered = $excel.RegisterXLL($xll)
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'NATSExcel'

  $ws.Range('A1').Value2 = 'NATSExcel installed successfully!'
  $ws.Range('A1').Font.Size = 16
  $ws.Range('A1').Font.Bold = $true
  $ws.Range('A2').Value2 = 'NATS connection test - ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  $ws.Range('A2').Font.Italic = $true

  $ws.Range('A4').Value2 = 'Server:'
  $ws.Range('B4').Value2 = $target
  $ws.Range('A5').Value2 = 'Status:'
  # Use the configured default cluster (no explicit host), so the credentials
  # the installer just wrote to NATSExcel.json apply. A raw URL argument would
  # bypass that config and fail on an authenticated server.
  $ws.Range('B5').Formula = '=IF(ISERROR(NatsUDFServerInfo()),"FAILED","OK - NATS answered")'
  $ws.Range('A6').Value2 = 'Server says:'
  $ws.Range('B6').Formula = '=NatsUDFServerInfo()'
  $ws.Range('A7').Value2 = 'Cluster:'
  $ws.Range('B7').Formula = '=NatsUDFClusterName()'

  $ws.Range('A9').Value2 = 'Tempo real (pressione F9 / publique em test.hello):'
  $ws.Range('A10').Value2 = 'RTD SUB:'
  $ws.Range('B10').Formula = '=RTD("NatsRtd",,"SUB","test.>","","")'

  $ws.Range('A12').Value2 = 'Exemplo: =NatsUDFSubjectLatest("test.>")'
  $ws.Range('A4:A6').Font.Bold = $true
  $ws.Range('A7').Font.Bold = $true
  $ws.Columns.Item('A').ColumnWidth = 42
  $ws.Columns.Item('B').ColumnWidth = 40
  $ws.Range('B5').Font.Bold = $true

  $excel.Application.Calculate() | Out-Null
  Write-Host "[verify] sample workbook opened (XLL registered: $registered)"
  exit 0
} catch {
  Show ('Could not open Excel: ' + $_.Exception.Message) 'NATSExcel'
  exit 0
}
