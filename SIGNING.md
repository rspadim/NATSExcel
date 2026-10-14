# Code signing

Status: **inactive**. NATSExcel releases are **not signed** and nothing in the CI
signs them. The SignPath-based pipeline that once existed was removed (it is
preserved in commit `0edaf06` for reference); when a token exists it will be
applied **manually** first (outside the pipeline), and only wired into the CI
after that manual flow is proven.

## What gets signed (when a token exists)

| Artifact | Why |
| --- | --- |
| `NATSExcel-packed.xll`, `NATSExcel64-packed.xll` | the add-ins Excel loads; an unsigned XLL is a classic SmartScreen/AV false positive |
| `NATSExcel.dll` | the managed assembly shipped next to the XLLs (manual installs use it); every binary a user receives should be verifiable |
| `NATSExcelSetup-<version>.exe` (Inno Setup) | the installer is the first file a new user runs and the #1 false-positive target |

Order matters: sign the **payload first** (both XLLs and the dll), then build the
installer, then sign the installer. An installer built from unsigned binaries
keeps the warning no matter what is signed afterwards.

## Manual procedure (the current plan)

1. Produce the release payload (the packed XLLs, `NATSExcel.dll` and the installer) -
   the `Build & Release` artifacts are enough.
2. Sign the payload, rebuild the installer from the signed files, sign the installer.
3. Upload the signed files to the release.

## Service

Free Windows code signing for open-source projects is being pursued through two
programs; neither is active for this project yet, so the landing page does not
advertise signing.

- **Necessary Code Signing** (<https://sign.necessary.nu/>) - the certificate is
  issued to *Necessary Innovations AB (Sweden)*, so a signed binary shows that
  publisher; identity and address are verified before a token is granted, and the
  service signs established projects (the application covers the sibling add-ins
  RedisExcel and NATSExcel). Signing is three `osslsigncode` steps, which is what
  makes it usable from CI:

  ```
  osslsigncode extract-data -in app.exe -out tosign.bin
  curl -X POST -H "Authorization: Bearer $SIGNING_TOKEN" \
    --data-binary @tosign.bin https://sign.necessary.nu/windows/sign -o signed.bin
  osslsigncode attach-signature -sigin signed.bin -in app.exe -out app-signed.exe
  ```

- **SignPath Foundation** (<https://signpath.org/>) - the other free lane; the
  certificate is issued to the foundation, and it additionally requires a published
  code-signing policy on the project's home page plus a manually approved signing
  request per release.

### When a token exists

1. Sign manually first, in the release order (payload, then installer).
2. Verify every file (`Get-AuthenticodeSignature`, or `osslsigncode verify` on what
   the service returns): the status must be `Valid` and the publisher must be the
   signing service's, not a local certificate.
3. Only then wire it into the CI as a guarded step - a fork or a PR has no token and
   must not fail - and keep the post-signing verification as the gate: a
   misconfigured signing step must fail the job, never ship a half-signed release.

## Never

No script in this repository disables or reconfigures antivirus or Defender
(that is MITRE T1562.001 and any EDR will - correctly - block it). Test tooling
may only **read** the AV state.
