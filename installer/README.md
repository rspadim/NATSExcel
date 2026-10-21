# NATSExcel installer

A small [Inno Setup 6](https://jrsoftware.org/isinfo.php) installer that copies
the add-in into one folder, writes a starter `NATSExcel.json`, and keeps both
XLLs (removing the one that does not match your Office bitness). It runs
**per-user** (no admin), because the add-in registers its own COM ProgId under
`HKCU` when Excel loads it. The only admin step is the *optional* Windows
service registration, which is isolated in `nats-service.ps1`.

## What it does

1. Copies `NATSExcel.dll` + the packed XLLs + `NLog.config` into a folder you
   pick (default `C:\NATSExcel`).
2. Detects the Office bitness and keeps only the matching XLL
   (`NATSExcel64-packed.xll` for 64-bit, `NATSExcel-packed.xll` for 32-bit).
   It then **unregisters an older copy**: any Excel auto-load entry
   (`HKCU\Software\Microsoft\Office\<ver>\Excel\Options\OPEN`/`OPENn`) that points
   at a *different* NATSExcel XLL is deleted, and so is a dead "Add-in Manager"
   label left by a renamed or moved copy (the second, dead `NATSExcel` line in
   File > Options > Add-ins). Other add-ins' entries are never touched and this
   install's own entry is written first, so an upgrade into another folder no
   longer leaves **two** add-ins loaded (duplicated functions, two configs, one
   of them showing `#CONFIG`).
   The wizard is localized in **11 languages** (简体中文, English, 繁體中文,
   Português, Español, Русский, Français, 한국어, Deutsch, 日本語 and Português
   (Brasil)), every one with the same 71 wizard strings. The Excel-facing
   function names and cell markers stay in **English** on purpose (see the main
   README "Languages" section): a formula must return the same text whatever the
   user's locale.
3. Asks for the **connection**: the server URL (a comma-separated list for a
   cluster) and the credential style - none, token, user/password, NKey,
   JWT + seed, creds file or mTLS - showing only the fields the chosen style
   needs. Any of them can be edited in `NATSExcel.json` afterwards. The page has
   a **Test connection** button: it probes what is in the fields at that moment
   (no page change, no file written), says up front that it waits **up to 5
   seconds**, and shows the result above the button in green or red. The reason
   is the **validator's own report**, appended after the localized sentence (the
   probe runs `tools\NatsValidate.exe --test`; its exit code picks the sentence:
   config invalid, connection/credentials refused, or unexpected failure). A
   `cmd.exe` redirect re-encodes the console text into the OEM codepage, so the
   **localized sentence never comes from that text** — only the appended detail
   does, and it is ASCII on purpose.
   **The probe is read-only**: it connects and pings (the real NATS.Net client),
   and it never publishes a message, creates a stream/consumer/KV bucket or
   writes a subject. The address the user typed is often somebody else's
   server, so a probe that mutated it would be a defect. The verdict is
   "connected **and** authenticated" — per-command authorization is the user's
   own business, not the probe's.
4. **Optional: installs a server.** NATS runs natively on Windows, so the
   default is *download one small `nats-server.exe`* and start it with JetStream
   on. Docker is offered as an alternative. The native path writes a starter
   `nats.conf` with a **durable** `jetstream { store_dir: ... }` (the server
   default is a temp path).
5. **Optional: registers NATS as a Windows service** (needs admin; skipped with
   a clear message if the wizard is not elevated).
6. **End-of-install check.** It runs the same `tools\NatsValidate.exe --test`
   against the configuration that was just written (read-only: connect + PING,
   never a publish) and, if you ticked the box, opens a visible Excel workbook
   (XLL loaded) that only **reads** (`=NatsUDFServerInfo`,
   `=NatsUDFClusterName`) plus one `SUB` cell on this project's own reserved
   subject (`test.natsexcel.>`).
7. The uninstaller asks whether to keep your `NATSExcel.json` and offers to
   remove a NATS server it created (nothing is removed unless you tick it).

## Silent / "files only" mode

No wizard, no Excel prompt - just copy the files and use the given URL:

```powershell
NATSExcelSetup-1.0.0.exe /VERYSILENT /SUPPRESSMSGBOXES /DIR="C:\NATSExcel" /URL="nats://localhost:4222" /TOKEN="secret" /SERVER=none /OPENEXCEL=0
```

`/URL`, `/TOKEN`, `/SERVER` (`none`|`native`|`docker`) and `/OPENEXCEL`
(`0`|`1`) are optional (defaults: `nats://localhost:4222`, empty, `none`, `0`).
The other credential styles use `/AUTH=<none|token|userpass|nkey|jwt|creds|mtls>`
plus the matching field(s): `/USER`, `/PASSWORD`, `/NKEYSEED`, `/JWT`, `/CREDS`,
`/TLSCERT`, `/TLSCERTPASSWORD`. There is no `/MODE`: the wizard has no
simple/advanced split (the Credentials page already shows only the fields of the
chosen style). The installer is a normal
`.exe`, so you can also install by hand: copy the `.xll`, `NATSExcel.dll`,
`NLog.config` and `NATSExcel.json` into one folder and register the XLL from
Excel.

## Helper scripts

The installer ships these under `C:\NATSExcel\` (all usable on their own):

- `scripts\install-server.ps1 -Kind native -Port 4222 -Dir C:\NATSExcel\server`
  - brings up a server: `native` (download/reuse `nats-server.exe` + starter
  `nats.conf`) or `docker`. Exit `3` means the method is unavailable here.
  `-NoJetStream` starts without `-js`; `-Image nats:2-alpine` pins a tag.
- `scripts\nats-service.ps1 -Action Install|Start|Stop|Remove|Status -Dir ...`
  - the Windows service (the only admin step).
- `tools\NatsValidate.exe "C:\NATSExcel\NATSExcel.json" --test` - the connection
  check the wizard's button and the end-of-install step run: the **same
  NATS.Net client the add-in uses**, so every credential style (token,
  user/password, NKey seed, JWT + seed, creds file, mTLS) is exercised for real.
  `--test` only **connects and pings** - it never publishes, creates a stream or
  writes a subject, and it does not test per-command permissions (the verdict is
  "connected + authenticated"). Exit codes: `0` OK, `1` config problems, `2`
  unreadable, `3` dangling cluster/server reference, `4` connection failed.
  Without `--test` it validates the file and needs no server.
- `scripts\verify-connection.ps1 -AddInDir "C:\NATSExcel" -ServerHost localhost:4222 [-Config <json>] [-OpenExcel]`
  - runs the same read-only `NatsValidate --test` check against `-Config` (or
  `<AddInDir>\NATSExcel.json`, or a draft it builds in a temp file) and, with
  `-OpenExcel`, opens the sample sheet, which only reads
  (`=NatsUDFServerInfo`, `=NatsUDFClusterName`) plus one `SUB` cell on the
  project's own `test.natsexcel.>` prefix. Exit `0` = answered, `1` = no answer
  (or the validator could not be run).
- `scripts\validate-config.ps1` (or double-click `Validate-Config.cmd`) - checks
  `NATSExcel.json` against the shipped `NATSExcel.schema.json` (the same schema
  the add-in embeds) and prints an OK/FAIL report.
- `scripts\detect-servers.ps1` / `scripts\uninstall-server.ps1` - used by the
  uninstaller.
- `scripts\banner.ps1` - the console header the helpers above print (5-row ASCII
  wordmark, the installed version from the DLL's `ProductVersion`, and the
  project URL). ASCII only, because the console renders in an OEM codepage; it
  falls back to a compact one-line header in a window narrower than 70 columns
  or when the output is redirected. Dot-sourced through a try/catch, so a missing
  file degrades to a plain header instead of breaking a helper.

## Uninstall

Run **Start Menu > Uninstall NATSExcel** (or Apps & features). It removes the
files, the Excel auto-load entry and the `NatsRtd` COM ProgId. It then:

- offers to remove any NATS server it finds (the `NATS` Windows service,
  `nats-server.exe` under the install folder, the `natsexcel` Docker container)
  - **nothing is removed unless you tick it** (default: keep everything);
- asks whether to keep your `NATSExcel.json`.

## Build it yourself

```powershell
# One-time: install Inno Setup 6 (winget or choco)
choco install innosetup -y

# Build the add-in first (the installer packages bin\Release\net48\publish)
dotnet build NATSExcel.sln -c Release

# Compile the installer (version is injected; use the tag without the 'v')
iscc "/DAppVersion=1.0.0" "installer\NATSExcel.iss"
# -> installer\output\NATSExcelSetup-1.0.0.exe
```

CI builds the installer on every run (so a broken `.iss` fails fast) and attaches
`NATSExcelSetup-<version>.exe` to the GitHub release on a `v*` tag.

## Notes / limits

- The XLL is **not code-signed**, so Windows SmartScreen and some antivirus
  products may warn; the installer cannot fix that without a code-signing
  certificate.
- Excel must be **closed** while installing/updating (the `.XLL` file is locked
  while loaded).
- Docker is **never required**: it is only offered as an alternative to the
  native server.
- The add-in folder is only the *first* place searched for `NATSExcel.json`; it
  still falls back to `%USERPROFILE%`, the Excel folder and `C:\Windows`.
