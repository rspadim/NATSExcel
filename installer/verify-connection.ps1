#Requires -Version 5.1
<#
  verify-connection.ps1 - end-of-install check for NATSExcel.

  * Proves the connection with the SAME probe the wizard's "Test connection"
    button uses: tools\NatsValidate.exe --test, i.e. the real NATS.Net client,
    against the configuration the installer just wrote. The check is READ-ONLY
    on purpose - the address the user typed is very often somebody else's
    server: it connects and pings, and it never publishes into a subject,
    creates a stream or a consumer, or writes a key. The verdict is "connected
    AND authenticated" (every credential style, because only the client can
    sign the server nonce); per-command authorization is the user's own
    business, not this check's.
  * With -OpenExcel, opens a visible Excel workbook that loads the installed XLL
    and shows the connection status with a friendly message. The sheet only
    READS (=NatsUDFServerInfo and =NatsUDFClusterName) plus one RTD SUB cell on
    this project's own reserved subject (test.natsexcel.>), so it does not
    publish into a namespace it does not own either.

  Exit codes: 0 = NATS answered (and the sheet opened if asked), 1 = no answer.

  Examples:
    powershell -ExecutionPolicy Bypass -File verify-connection.ps1 -AddInDir "C:\NATSExcel" -ServerHost "localhost:4222"
    powershell -ExecutionPolicy Bypass -File verify-connection.ps1 -AddInDir "C:\NATSExcel" -ServerHost "localhost:4222" -Config "C:\NATSExcel\NATSExcel.json" -OpenExcel
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$AddInDir,
  [string]$ServerHost = 'localhost',
  [int]$Port = 4222,
  [string]$Token = '',
  # The configuration to test with (the installer passes the file it wrote, so
  # every credential style is exercised). Without it <AddInDir>\NATSExcel.json
  # is used; without that either, a minimal draft is built from -ServerHost and
  # -Token in a temp file, so a manual run from the scripts folder still works.
  [string]$Config = '',
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

# ---- The connection probe: READ-ONLY, by construction ----------------------
# The probe is tools\NatsValidate.exe --test: the real NATS.Net client the
# add-in uses, which CONNECTS (the server's INFO banner) and sends a PING, and
# does nothing else. It never publishes a message, never creates a stream,
# consumer or KV bucket and never writes a subject: the endpoint the user typed
# is usually not ours, so a probe that mutated it would be a defect, not a test.
#
# The success verdict is "connected AND authenticated". NOTHING here inspects
# per-command authorization: a restricted user is the user's own responsibility,
# and a probe that demanded a publish permission would reject a perfectly usable
# read-only server.
function Test-NatsWithValidator([string]$configPath) {
  $exe = Join-Path $AddInDir 'tools\NatsValidate.exe'
  if (-not (Test-Path -LiteralPath $exe)) {
    # A manually-cloned scripts folder (no tools\ next to it): try the repo
    # layout, then give up and say so rather than guessing.
    $sibling = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'tools\NatsValidate\bin\Release\net48\NatsValidate.exe'
    if (Test-Path -LiteralPath $sibling) { $exe = $sibling } else { return $null }
  }
  if (-not (Test-Path -LiteralPath $configPath)) { return $null }

  $out = Join-Path ([System.IO.Path]::GetTempPath()) ('natsexcel-verify-' + [Guid]::NewGuid().ToString('N') + '.txt')
  try {
    $argv = @('"' + $exe + '"', '"' + $configPath + '"', '--test', '--timeout', '5000')
    $p = Start-Process -FilePath 'cmd.exe' -ArgumentList ('/c "' + ($argv -join ' ') + ' > "' + $out + '" 2>&1"') `
      -NoNewWindow -Wait -PassThru -ErrorAction Stop
    $code = $p.ExitCode
    $detail = ''
    $tested = ''
    if (Test-Path -LiteralPath $out) {
      foreach ($line in (Get-Content -LiteralPath $out)) {
        $t = $line.Trim()
        # The validator prints `  urls      : nats://host:port` before connecting:
        # that is the endpoint it actually tested (the config's, not -ServerHost).
        if (-not $tested -and $t -match '^urls\s*:\s*(\S+)') { $tested = $Matches[1] }
        if ($t.StartsWith('RESULT:')) { $detail = $t }
      }
    }
    return [pscustomobject]@{ Code = $code; Detail = $detail; Tested = $tested }
  } catch {
    return $null
  } finally {
    Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue
  }
}

# The configuration to test: the file the installer wrote (it carries every
# credential style), else a minimal draft for a manual run.
$configPath = $Config
if (-not $configPath) {
  $installed = Join-Path $AddInDir 'NATSExcel.json'
  if (Test-Path -LiteralPath $installed) { $configPath = $installed }
}
$draftPath = $null
if (-not $configPath) {
  $draftPath = Join-Path ([System.IO.Path]::GetTempPath()) ('natsexcel-verify-' + [Guid]::NewGuid().ToString('N') + '.json')
  $auth = if ($Token -ne '') { '  "Auth": { "token": "' + ($Token -replace '\\', '\\\\' -replace '"', '\"') + '" },' } else { '  "Auth": {},' }
  # A here-string with explicit placeholders: inside an @() array literal the
  # comma binds TIGHTER than '+', and -f binds to the LAST element only, so
  # either shortcut silently produces broken JSON ("nats://" on its own line).
  $template = @'
{
  "Default": "local",
  "Servers": { "srv1": "nats://__TARGET__" },
  "Clusters": { "local": { "urls": ["srv1"] } },
__AUTH__
  "JetStream": { "enabled": true, "defaultTimeoutMs": 5000 }
}
'@
  $draft = $template.Replace('__TARGET__', $target).Replace('__AUTH__', $auth)
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($draftPath, $draft, $utf8NoBom)
  $configPath = $draftPath
}

Write-Host "[verify] testing NATS at $target (read-only: connect + PING) ..."
try {
  $probe = Test-NatsWithValidator $configPath
} finally {
  if ($draftPath) { Remove-Item -LiteralPath $draftPath -Force -ErrorAction SilentlyContinue }
}

# The validator tests the CONFIG's endpoint (that is what the add-in will use),
# which is not necessarily the -ServerHost the caller passed. Compare them with
# the scheme stripped (a "nats://host:port" vs "host:port" difference is not a
# mismatch) and, when they really differ, report the server that was tested
# instead of a target the user never asked about.
function Get-HostPortOf([string]$u) {
  $v = $u.Trim()
  if ($v -match '^\s*nats://(.+)$') { $v = $Matches[1] }
  $v = ($v -split ',')[0]
  return $v.Trim().TrimEnd('/')
}
if ($probe -and $probe.Tested -and ((Get-HostPortOf $probe.Tested) -ne (Get-HostPortOf $target))) {
  Write-Host "[verify] note: the configuration names $($probe.Tested); -ServerHost $target was not used."
  $target = $probe.Tested
}

# $null means the validator could not be run at all: say that instead of
# claiming the server is dead.
if ($null -eq $probe) {
  Show ("Could not run the connection check (tools\NatsValidate.exe was not found next to the add-in).`r`n`r`n" +
        "Reinstall NATSExcel, or run it from an installed folder.") 'NATSExcel - connection'
  exit 1
}

$ok = ($probe.Code -eq 0)

if (-not $ok) {
  # Exit 1/2/3 is a config problem, not a dead server: say which, so a typo in a
  # credential field is never reported as "the server is down".
  if ($probe.Code -ge 1 -and $probe.Code -le 3) {
    Show ("The configuration could not be checked.`r`n`r`n" + $probe.Detail) 'NATSExcel - configuration'
    exit 1
  }
  $detail = ''
  if ($probe.Detail) { $detail = "`r`n`r`n" + $probe.Detail }
  Show ("Could not reach NATS at $target.`r`n`r`n" +
        "- Check that the server is running (the NATS service, nats-server.exe, or docker ps).`r`n" +
        "- If it runs on another host/port, adjust NATSExcel.json." + $detail) 'NATSExcel - connection'
  exit 1
}

Write-Host "[verify] NATS answered at $target"

if (-not $OpenExcel) {
  Show ("All good!`r`n`r`nNATS answered at $target.`r`n" +
        "Open Excel, register the XLL (File > Options > Add-ins > Go... > Browse...) " +
        'and use, for example, =RTD("NatsRtd",,"SUB","test.natsexcel.>") in a cell.') 'NATSExcel - installed'
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

  $ws.Range('A9').Value2 = 'Real time (press F9 / publish on test.natsexcel.> - this project''s own reserved prefix):'
  $ws.Range('A10').Value2 = 'RTD SUB:'
  # A subscription only READS. It stays on the project's own reserved prefix
  # (test.natsexcel.>) so the sample never listens to - let alone publishes
  # into - a namespace that belongs to someone else.
  $ws.Range('B10').Formula = '=RTD("NatsRtd",,"SUB","test.natsexcel.>","","")'

  $ws.Range('A12').Value2 = 'Example: =NatsUDFSubjectLatest("test.natsexcel.>")'
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
