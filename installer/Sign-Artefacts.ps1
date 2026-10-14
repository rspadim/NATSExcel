#Requires -Version 5.1
<#
    Sign-Artefacts.ps1 - stages a set of release binaries for SignPath code
    signing and (after the signing request) verifies the returned signatures.

    Why a script: the two signing rounds (the packed XLLs + the managed dll, and
    then the installer that CONTAINS them) both need the same staging and the
    same post-check, and the order is the whole point:

      1. build            -> bin\Release\net48\publish\*.xll + NATSExcel.dll
      2. sign these       <-- sign the PAYLOAD first
      3. iscc             -> installer\output\NATSExcelSetup-<v>.exe
      4. sign the .exe    <-- then the package that carries the payload

    Signing AFTER the installer build without signing the payload first would
    ship an installer whose inner binaries are unsigned; signing the payload
    afterwards would not change what was packed. Inside-out, always.

    Usage (staging):
      powershell -File Sign-Artefacts.ps1 -Zip out.zip -Files a.xll,b.dll -Root <repo>

    Usage (verify, after the signed artefact was unpacked over the originals):
      powershell -File Sign-Artefacts.ps1 -Verify a.xll,b.dll
#>
[CmdletBinding(DefaultParameterSetName = 'stage')]
param(
    [Parameter(ParameterSetName = 'stage', Mandatory = $true)]
    [string]$Zip,

    [Parameter(ParameterSetName = 'stage', Mandatory = $true)]
    [string[]]$Files,

    # Repository root, so the zip carries paths relative to it (SignPath's
    # artifact configuration matches on those relative paths).
    [Parameter(ParameterSetName = 'stage')]
    [string]$Root = '.',

    [Parameter(ParameterSetName = 'verify', Mandatory = $true)]
    [string[]]$Verify
)

$ErrorActionPreference = 'Stop'

# `powershell -File script.ps1 -Files a,b,c` hands the parameter ONE string
# ("a,b,c"), not an array: -File passes arguments literally. Accept both the
# native array form (a,b,c from a session) and the CLI form by splitting on
# commas (and keeping any entries that were already separate).
function Expand-List([string[]]$values) {
    $out = @()
    foreach ($v in $values) {
        if ($null -eq $v) { continue }
        foreach ($part in ($v -split ',')) {
            $t = $part.Trim()
            if ($t.Length -gt 0) { $out += $t }
        }
    }
    return $out
}

if ($PSCmdlet.ParameterSetName -eq 'stage') {
    $Files = Expand-List $Files
    $resolved = @()
    foreach ($f in $Files) {
        $path = Join-Path $Root $f
        if (-not (Test-Path -LiteralPath $path)) { throw "file to sign is missing: $f" }
        $resolved += (Resolve-Path -LiteralPath $path).Path
    }

    # A zip keeps the relative layout SignPath's artifact-configuration matches
    # on, and it is what the submit-signing-request action consumes.
    if (Test-Path -LiteralPath $Zip) { Remove-Item -LiteralPath $Zip -Force }
    $stage = Join-Path ([System.IO.Path]::GetTempPath()) ('signstage-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $stage | Out-Null
    try {
        foreach ($p in $resolved) {
            $rel = $p.Substring((Resolve-Path -LiteralPath $Root).Path.Length).TrimStart('\', '/')
            $dest = Join-Path $stage $rel
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
            Copy-Item -LiteralPath $p -Destination $dest -Force
        }
        Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $Zip -Force
    }
    finally {
        Remove-Item -Recurse -Force $stage -ErrorAction SilentlyContinue
    }
    Write-Host "staged $($resolved.Count) file(s) -> $Zip"
    exit 0
}

# ---- verify -----------------------------------------------------------------
# A misconfigured signing step must FAIL the job, not silently ship unsigned
# binaries. Get-AuthenticodeSignature on the returned files is the check.
$Verify = Expand-List $Verify
$failed = @()
foreach ($f in $Verify) {
    if (-not (Test-Path -LiteralPath $f)) { $failed += "$f (missing)"; continue }
    $sig = Get-AuthenticodeSignature -LiteralPath $f
    $subject = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { '(none)' }
    Write-Host ("  {0,-30} {1}  {2}" -f $f, $sig.Status, $subject)
    if ($sig.Status -ne 'Valid') { $failed += "$f ($($sig.Status))" }
}
if ($failed.Count -gt 0) {
    throw "unsigned or invalid signature after signing: " + ($failed -join ', ')
}
Write-Host "all $($Verify.Count) file(s) carry a valid signature"
exit 0
