<#
.SYNOPSIS
  Lists the [CustomMessages] keys a wizard language must define, so a human can
  diff a translation against the English source of truth.

.DESCRIPTION
  Reads NATSExcel.iss, collects the keys of the `english.` block (the source of
  truth) and, for the requested language tag, prints the `lang.Key=Value` lines
  the .iss is expected to contain (English values shown as a placeholder).

  If the .iss already has a block for that language, it also reports which keys
  are MISSING and which are EXTRA compared with english, and exits non-zero when
  the key set does not match exactly. This does NOT translate anything.

  Inno Setup falls back to a hardcoded English string when a key is absent, so a
  missing key is a silent bug: run this after any edit to the language blocks.

.EXAMPLE
  powershell -File installer\languages\messages.ps1 -Language german

.EXAMPLE
  powershell -File installer\languages\messages.ps1 -Language japanese -KeysOnly
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$Language,

  [string]$IssPath,

  [switch]$KeysOnly
)

$ErrorActionPreference = 'Stop'

if (-not $IssPath) {
  $IssPath = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'installer\NATSExcel.iss'
}
if (-not (Test-Path -LiteralPath $IssPath)) {
  throw "Cannot find NATSExcel.iss at '$IssPath'. Pass -IssPath."
}

$text  = [System.IO.File]::ReadAllText($IssPath, [System.Text.Encoding]::UTF8)
$lines = $text -split "`r`n|`n"

$inCM = $false
$english = New-Object System.Collections.Specialized.OrderedDictionary
$lang    = New-Object System.Collections.Specialized.OrderedDictionary

foreach ($l in $lines) {
  if ($l -match '^\s*\[CustomMessages\]\s*$') { $inCM = $true; continue }
  if ($l -match '^\s*\[') { $inCM = $false; continue }
  if (-not $inCM) { continue }

  if ($l -match '^english\.([^=]+)=(.*)$') {
    $english[$Matches[1]] = $Matches[2]
  }
  elseif ($l -match ('^' + [regex]::Escape($Language) + '\.([^=]+)=(.*)$')) {
    $lang[$Matches[1]] = $Matches[2]
  }
}

if ($english.Count -eq 0) { throw "No 'english.' keys found in [CustomMessages] of $IssPath." }

Write-Host ("English source keys: {0}" -f $english.Count)
Write-Host ("Language: {0}" -f $Language)
Write-Host ''

if ($KeysOnly) {
  foreach ($k in $english.Keys) { Write-Host ("{0}.{1}" -f $Language, $k) }
}
else {
  Write-Host ("Expected lines for '{0}' ({1}.Key=Value, English shown as a placeholder):" -f $Language, $Language)
  foreach ($k in $english.Keys) { Write-Host ("{0}.{1}={2}" -f $Language, $k, $english[$k]) }
}

if ($lang.Count -eq 0) {
  Write-Host ''
  Write-Host ("No '{0}.' block found in the .iss yet (expected lines above)." -f $Language)
  exit 0
}

$missing = @($english.Keys | Where-Object { -not $lang.Contains($_) })
$extra   = @($lang.Keys    | Where-Object { -not $english.Contains($_) })

Write-Host ''
Write-Host ("Present: {0}/{1}" -f $lang.Count, $english.Count)
if ($missing.Count) { Write-Host ("MISSING ({0}): {1}" -f $missing.Count, ($missing -join ', ')) }
if ($extra.Count)   { Write-Host ("EXTRA ({0}): {1}" -f $extra.Count, ($extra -join ', ')) }
if ($missing.Count -or $extra.Count) { exit 1 }

Write-Host 'Key set matches english exactly.'
exit 0
