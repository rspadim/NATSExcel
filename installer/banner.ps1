<#
  banner.ps1 - the console header NATSExcel prints before its own output.

  Why: the installer runs these helpers in a visible console (setting up a
  server, testing the connection, validating a config, opening the sample
  sheet). A window that appears, prints terse technical lines and vanishes is
  exactly what a user reads as suspicious - the single most common "is this a
  virus?" reaction. A clear, self-identifying header answers "what is this, who
  makes it, where did it come from" before any command runs.

  Usage (dot-source it, then print):
    . "$PSScriptRoot\banner.ps1"
    Write-NatsExcelBanner -Action 'Checking the NATS connection'

  Output is ASCII-only ON PURPOSE: the helpers run in consoles whose codepage is
  the OEM one (437/850/866/932...), where box-drawing or accented characters
  turn into mojibake. ASCII renders identically everywhere.

  The art is skipped in a window narrower than MIN_WIDTH (or with no console at
  all, e.g. output redirected to a file); a compact one-line header is printed
  instead, so a narrow CONSOLE never wraps the art into a mess.
#>

$script:NatsExcelBannerHome = 'https://github.com/rspadim/NATSExcel'
$script:NatsExcelBannerMinWidth = 70

# 5-row "standard" style art spelling "NATS Excel", one string per row.
# Baseline on the last row; lowercase sits on the x-height, ascenders (l) reach
# the first row. Kept as plain data: no font engine, no dependency. The longest
# row is 62 characters, so the art fits MIN_WIDTH with room for the indent.
$script:NatsExcelArt = @(
  ' _   _     _     _____  ____                                _ '
  '| \ | |   / \   |_   _|/ ___|    ___  __  __   ___    ___  | |'
  '|  \| |  / _ \    | |  \___ \   / _ \ \ \/ /  / __|  / _ \ | |'
  '| |\  | / ___ \   | |   ___) | |  __/  >  <  | (__  |  __/ | |'
  '|_| \_|/_/   \_\  |_|  |____/   \___| /_/\_\  \___|  \___| |_|'
)

function Get-NatsExcelVersion {
  param([string]$InstallDir = '')

  # 1) The installed DLL carries ProductVersion (e.g. "v0.1.5+<commit>").
  $candidates = @()
  if ($InstallDir -ne '') { $candidates += (Join-Path $InstallDir 'NATSExcel.dll') }
  try {
    # 2) The helper's own folder and its parent ({app}\scripts -> {app}).
    $candidates += (Join-Path $PSScriptRoot 'NATSExcel.dll')
    $candidates += (Join-Path (Split-Path -Parent $PSScriptRoot) 'NATSExcel.dll')
  } catch { }

  foreach ($dll in $candidates) {
    try {
      if ($dll -and (Test-Path -LiteralPath $dll)) {
        $v = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($dll).ProductVersion
        if ($v) {
          # Local builds report "v0.1.5+<sha>"; the sha is noise in a header.
          $plus = $v.IndexOf('+')
          if ($plus -gt 0) { $v = $v.Substring(0, $plus) }
          return $v.Trim()
        }
      }
    } catch { }
  }
  return ''
}

function Test-BannerWideEnough {
  param([int]$MinWidth = 70)
  try {
    $width = [Console]::WindowWidth
    if (-not $width -or $width -lt 1) { return $false }
    return ($width -ge $MinWidth)
  } catch {
    # A redirected stream throws here: assume compact output.
    return $false
  }
}

function Write-NatsExcelBanner {
  param(
    # Short line describing what THIS run does ("Installing the NATS server").
    [string]$Action = '',
    # Installed folder, when the caller knows it.
    [string]$InstallDir = '',
    # Force the full art even when the window looks narrow.
    [switch]$Force
  )

  $version = Get-NatsExcelVersion -InstallDir $InstallDir
  $tagline = 'NATS client add-in for Excel'
  if ($version) { $tagline = $tagline + '  -  ' + $version }

  $wide = $true
  if (-not $Force) { $wide = Test-BannerWideEnough -MinWidth $script:NatsExcelBannerMinWidth }

  Write-Host ''
  if ($wide) {
    $art = $script:NatsExcelArt
    for ($r = 0; $r -lt $art.Count; $r++) {
      Write-Host ('  ' + $art[$r]) -ForegroundColor Cyan
    }
    Write-Host ''
    Write-Host ('  ' + $tagline) -ForegroundColor Gray
    Write-Host ('  ' + $script:NatsExcelBannerHome) -ForegroundColor DarkGray
  } else {
    Write-Host ('  NATSExcel' + $(if ($version) { " $version" } else { '' }) + ' - NATS client add-in for Excel') -ForegroundColor Cyan
    Write-Host ('  ' + $script:NatsExcelBannerHome) -ForegroundColor DarkGray
  }

  if ($Action -ne '') {
    Write-Host ''
    Write-Host ('  ' + $Action) -ForegroundColor White
  }
  if ($InstallDir -ne '') {
    Write-Host ('  Folder: ' + $InstallDir) -ForegroundColor DarkGray
  }
  Write-Host ''
}
