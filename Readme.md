# NATSExcel

Excel add-in (XLL) that brings **[NATS](https://nats.io/)** into Excel through
RTD (Real-Time Data) and UDF (User-Defined Functions): type a formula, NATS data
lands in the cell and — for the streaming functions — stays live.

It covers **Core NATS** (pub/sub, request/reply, queue groups, headers) and
**JetStream** (streams, consumers, replay, Key/Value Store, Object Store).
Backed by NATS.Net, NLog and Newtonsoft.Json, and targeting .NET Framework 4.8.

- Per-function catalog (frozen): [`docs/PUBLIC-SURFACE.md`](docs/PUBLIC-SURFACE.md)
- Architecture and the full function list: [`DESIGN.md`](DESIGN.md)
- Contributor/agent guidance: [`AGENTS.md`](AGENTS.md)
- Release history: [`CHANGELOG.md`](CHANGELOG.md)

> **Status: v0.1.0, first release.** The add-in builds, packs and loads in real
> Excel, exposes **91 `NatsUDF*` worksheet functions + 6 `NatsRTD*` status
> helpers** and an RTD server with **9 commands** (`SUB`, `KV`, `KVGET`,
> `KVKEYS`, `REQ`, `JS`, `MSG`, `OBJ`, `RTDCLUSTER`). The installer, the config
> validator, the unit/smoke/liveness/load suites and the Excel end-to-end script
> all exist (see [`CHANGELOG.md`](CHANGELOG.md) and [`PLAN.md`](PLAN.md) §8/§9).
> Sibling project: [RedisExcel](https://github.com/rspadim/RedisExcel) — same
> idea, built on NATS; not a code port, because the NATS client is async-only.

---

## 1. Install

### One-click install (recommended)

1. Open the **[latest release](https://github.com/rspadim/NATSExcel/releases/latest)**
   and, under **Assets**, click **`NATSExcelSetup-<version>.exe`** (for example
   `NATSExcelSetup-0.1.0.exe`).
2. Run it and follow the wizard. It is a **per-user install** (no admin needed),
   into `C:\NATSExcel` by default:
   - it copies the add-in and **keeps only the XLL that matches your Office
     bitness** — `NATSExcel64-packed.xll` for 64-bit Office,
     `NATSExcel-packed.xll` for 32-bit Office;
   - it asks whether to **install a NATS server** (native `nats-server.exe`
     download, recommended; Docker is offered instead) — see
     [§2](#2-no-server-yet-run-one-on-windows);
   - it asks for the **connection** (a URL, or a comma-separated list for a
     cluster) and the **credential style**: none, token, user/password, NKey
     seed, JWT + seed, a creds file, or mTLS. Only the fields for the style you
     picked are shown; the other styles are edited in `NATSExcel.json` later;
   - it writes a starter `NATSExcel.json` and, at the end, **tests the
     connection** and can **open a sample workbook** (see
     [§5](#5-see-it-work)).
   - the wizard is available in **English, Español and Português (Brasil)**.
3. Open Excel. The installer registers the XLL so Excel loads it automatically.
   If it does not load, enable it by hand:
   `File > Options > Add-ins > Go... > Browse...` and pick the `.xll` in
   `C:\NATSExcel`.

> **SmartScreen / antivirus.** The XLL is **not code-signed**, so Windows may
> show a SmartScreen warning: click **More info > Run anyway**. Some antivirus
> products may also warn (see [§7](#7-build-from-source--test)).
>
> The add-in is a `.xll` that Excel locks while it is loaded, so **close Excel**
> while installing or updating.

### Manual install

Keep the add-in files together in one dedicated folder, for example
`C:\NATSExcel\`:

1. From the latest release, download **`NATSExcel.dll`**, the **packed XLL** for
   your Office bitness and **`NLog.config`** (the release assets are
   `NATSExcel.dll`, `NATSExcel-packed.xll`, `NATSExcel64-packed.xll`,
   `NLog.config`, `NATSExcel.json` and the installer).
2. Copy [`NATSExcel.json`](NATSExcel.json) into the same folder and edit it for
   your server (see [§3](#3-connect-excel-to-your-server)).
3. In Excel: `File > Options > Add-ins > Go... > Browse...` and pick the `.xll`.

The add-in searches for `NATSExcel.json` in this order (**first existing file
wins**), as implemented in `Config.cs` `CandidatePaths`:

1. the **add-in folder** (the folder the `.xll` was loaded from),
2. **`%USERPROFILE%`**,
3. the **Excel folder** (the folder of the running Excel process), then
4. **`C:\Windows`**.

If no file is found, the add-in runs with built-in defaults
(`nats://localhost:4222`). Comments (`//` and `/* */`) are allowed in the file;
a wrongly-typed value is dropped on its own and the rest of the file is kept.
`=NatsUDFConfigPath()` tells you which file is actually in use (it prints
`(defaults)` when none was loaded).

---

## 2. No server yet? Run one on Windows

NATS needs **no Docker on Windows**: the server is a single self-contained
`nats-server.exe`, and **JetStream is the same binary** — enabled with `-js`.
Pick one option.

### Option A (recommended): winget

The `NATSAuthors.NATSServer` package is published in the Windows Package Manager
community repository (verified: `winget search NATSServer` lists
`NATSAuthors.NATSServer`):

```powershell
winget install NATSAuthors.NATSServer
nats-server --version
nats-server -js --store_dir C:\NATSExcel\server\data
```

### Option B: the release zip (no package manager)

Download the Windows release from the official
**[nats-server releases page](https://github.com/nats-io/nats-server/releases)**
(pick the `...-windows-amd64.zip` of the version you want), unzip it, and run it
with JetStream on and a durable store directory. The zip name and the pinned
version this project uses (`2.15.1`) match
`https://github.com/nats-io/nats-server/releases/download/v<version>/nats-server-v<version>-windows-amd64.zip`.

```powershell
# After downloading and expanding the zip into .\nats-server-v<version>-windows-amd64\
.\nats-server-v<version>-windows-amd64\nats-server.exe -js --store_dir C:\NATSExcel\server\data
```

The installer's *native* option is exactly this download, and it writes a
starter `nats.conf` with a durable `jetstream { store_dir: ... }` for you.

### Option C: run it at boot (Windows service)

Registering the server as a Windows service is the **only** step that needs
admin. The installer ships a helper for it
(`installer/nats-service.ps1`, also copied to `C:\NATSExcel\scripts\`); run it
from an elevated shell:

```powershell
# Equivalent to what the wizard's "Register NATS as a Windows service" option runs:
powershell -ExecutionPolicy Bypass -File installer\nats-service.ps1 -Action Install -Dir C:\NATSExcel\server
```

The helper creates a service named `NATS` (via `sc.exe create NATS binPath= ...
start= auto`) that runs `nats-server.exe -c nats.conf`, and is idempotent. If
you prefer to do it yourself, point the service at the same `nats-server.exe -c
nats.conf` pair; the helper's script is the reference for the exact arguments.

### Option D: Docker (optional)

```powershell
docker run -d --name natsexcel -p 4222:4222 --restart unless-stopped nats:latest -js --store_dir /data
```

This is the same `nats:latest` image and arguments the installer's Docker option
uses. Docker is **never required**; it is offered as an alternative.

All options listen on `localhost:4222` (`127.0.0.1:4222`) by default.

> **JetStream storage matters.** The server's default `store_dir` is a
> *temporary* path, so **streams are lost on restart** unless you point it at a
> durable directory. The native paths above do that (`--store_dir`, or the
> starter `nats.conf` the installer writes, which contains
> `jetstream { store_dir: ... }`; the Docker option passes `--store_dir /data`).
> A server started **without** `-js`
> still accepts connections and Core pub/sub works, but every JetStream, Key/Value
> and Object Store call returns
> `Error: JetStream is not enabled on this server`.

---

## 3. Connect Excel to your server

`NATSExcel.json` drives everything. Its shape has four moving parts:

- **`Default`** — a plain *name* string, not an object: the endpoint a formula
  uses when it passes none. It may be a `Clusters` name, a `Servers` name or a
  raw `nats://` URL.
- **`Servers`** — named raw URLs, the primitives a cluster reuses.
- **`Clusters`** — the unit a formula names; each has a `urls` list (URLs or
  `Servers` aliases, mixed) plus optional `timeout`, `auth`, `reconnect`,
  `jetstream` and `expectCluster`.
- **`Auth` / `Reconnect`** — shared top-level blocks a cluster **overrides field
  by field**.

A minimal file (a cluster named `local`, no auth):

```json
{
  "Default": "local",
  "Clusters": { "local": { "urls": ["nats://localhost:4222"] } }
}
```

A cluster list with a token (the token sits on the cluster; a cluster's `auth`
overrides the shared top-level `Auth` block field by field):

```json
{
  "Default": "feed",

  "Servers": {
    "n1": "nats://n1.example:4222",
    "n2": "nats://n2.example:4222"
  },

  "Clusters": {
    "feed": {
      "urls": ["n1", "n2"],
      "auth": { "token": "REPLACE_WITH_YOUR_TOKEN" }
    }
  },

  "Reconnect": { "maxAttempts": 60, "waitMs": 2000, "jitterMs": 100 }
}
```

Notes:

- One URL is enough for a cluster — the client discovers the rest. Commas
  separate a URL list.
- **Credentials are always masked** in error cells, RTD `#ERROR` cells, status
  cells and log lines (`MaskNatsEndpoint` / `MaskedExceptionRenderer`).
- A cluster's `auth`, `timeout` and `reconnect` **merge over the shared
  top-level blocks field by field**, so you can set a shared default and change
  only what differs per cluster.
- Never commit a real endpoint or credential. See [`DESIGN.md`](DESIGN.md) §8 for
  every key, and [`NATSExcel.schema.json`](NATSExcel.schema.json) for the schema
  the add-in embeds.

The optional last argument of every UDF/RTD call is a **cluster reference**: a
`Clusters` name, a `Servers` alias, or a raw URL. A non-empty name that does not
resolve is a **`#CONFIG`** error — never a silent connection to `localhost`.

---

## 4. Examples

The catalog is frozen and large — this is a taste. The complete list with exact
arguments, volatility, write twins and the admin gate is
[`docs/PUBLIC-SURFACE.md`](docs/PUBLIC-SURFACE.md) (**91 `NatsUDF*`** + **6
`NatsRTD*`** + **9 RTD commands**).

Inside an Excel formula a literal `"` is doubled (`""`), not backslash-escaped.

### Publish (Core NATS)

```excel
=NatsUDFPublish("orders.new", "{""id"":1}")                    send a message
=NatsUDFPublishNonVolatile("orders.new", "{""id"":1}")         sent once, not per recalc
=NatsUDFPublishJSON("orders.new", A1:C20)                      publish a range as JSON
=NatsUDFPublishWithHeaders("orders.new", "go", A1:B3)          message + header range
```

Writes are **volatile** and re-run on every recalculation; the additive
`...NonVolatile` twin runs once, when its arguments change.

### Read / live

```excel
=RTD("NatsRtd",,"SUB","orders.>","","")        live push (wildcards: * one token, > one or more)
=RTD("NatsRtd",,"SUB","orders.>","workers","feed")   with a queue group and a cluster
=NatsUDFSubjectLatest("orders.>")             last value, cached by a UDF
=NatsUDFSubjectLatestJSON("orders.>")         same, JSON payload
=NatsUDFSubscriptionCount()                   how many subscriptions are live
```

### Request / reply

```excel
=NatsUDFRequest("orders.create", "{""id"":2}", 2000)   reply, or TIMEOUT / No Responders
=NatsUDFRequestJSON("orders.create", A1:C20, 2000)
```

### JetStream

```excel
=NatsUDFJSPublish("orders.new", "{""id"":9}")   durable publish; returns the sequence
=NatsUDFJSGetMessage("ORDERS", 1)               replay by sequence
=NatsUDFJSGetLastMessage("ORDERS", "orders.>")  last message matching a filter
=NatsUDFJSFetch("ORDERS", "batch", 10)          pull a batch into a range
=NatsUDFJSStreamInfo("ORDERS")                  messages / bytes / first / last sequence
=NatsUDFJSStreams()                             list streams
```

Stream/consumer admin (`NatsUDFStreamCreate`, `NatsUDFStreamUpdate`,
`NatsUDFStreamDelete`, `NatsUDFConsumer*`, `NatsUDFJSPurge`,
`NatsUDFJSDeleteMessage`) is **gated by `AllowAdminWrites`** (default `false`);
with the gate closed it returns
`Error: admin writes are disabled (set "AllowAdminWrites": true)`. This is so an
accidental `F9` cannot purge a production stream.

### Key/Value Store

```excel
=NatsUDFKVPut("prices", "BTC", "65000")          upsert
=NatsUDFKVPutJSON("prices", "BTC", A1:C1)        upsert from a range
=NatsUDFKVGet("prices", "BTC")                   read one key
=NatsUDFKVGetMultiple("prices", A1:A5)           read many keys into a range
=NatsUDFKVCreate("prices", "ETH", "3200")        only if absent
=NatsUDFKVUpdate("prices", "BTC", "66000", 3)    compare-and-swap by revision
=NatsUDFKVKeys("prices")                         list keys
=NatsUDFKVHistory("prices", "BTC")               revisions
=RTD("NatsRtd",,"KV","prices","BTC")             live KV watch (push)
```

`KVDelete` and `KVPurge` are admin-gated.

### Object Store

```excel
=NatsUDFObjectPut("reports", "q1.txt", "hello")                 store text
=NatsUDFObjectGet("reports", "q1.txt")                          read it back (text)
=NatsUDFObjectInfo("reports", "q1.txt")                         digest, size, chunks
=NatsUDFObjectList("reports")                                   list objects
```

### Cluster and topology

```excel
=NatsUDFClusterName()      the cluster the server announces
=NatsUDFClusterNodes()     nodes the client can see
=NatsUDFServerInfo()       server id / version / connect URLs
=NatsUDFRtt()              round-trip time
```

---

## 5. See it work

### Validate the config

The recommended tool is the C# validator, `tools/NatsValidate`: it checks the
file and, with `--test`, connects **with the same NATS.Net client the add-in
uses**, so every credential style is really exercised. From a checkout:

```powershell
dotnet run --project tools\NatsValidate -c Release -- NATSExcel.json
dotnet run --project tools\NatsValidate -c Release -- NATSExcel.json --test
```

Exit codes: `0` OK, `1` config problems, `2` unreadable, `3` dangling
cluster/server reference, `4` connection failed.

The installer also ships a lighter, schema-only checker: double-click
**`Validate-Config.cmd`** (in `C:\NATSExcel\`) or run
`scripts\validate-config.ps1`. It validates the file against the same
`NATSExcel.schema.json` the add-in embeds. Its optional wire-level connection
test only covers token and user/password — use `tools/NatsValidate` for the rest
(NKey, JWT + seed, creds file, mTLS).

Inside Excel you can also check the config directly:

```excel
=NatsUDFConfigValidate()   status of the config file in use
=NatsUDFConfigSchema()     the embedded schema
=NatsUDFConfigReload()     re-read NATSExcel.json without restarting Excel
```

### Open the sample workbook

The installer's last page has **"Open a sample NATSExcel workbook at the end"**
(also `/OPENEXCEL=1` in silent mode). It launches a visible workbook with the XLL
registered, a status check and a live `SUB` cell so you can see data arrive. You
can reproduce it later from the installer's helpers:

```powershell
powershell -ExecutionPolicy Bypass -File C:\NATSExcel\scripts\verify-connection.ps1 -AddInDir C:\NATSExcel -ServerHost localhost:4222 -OpenExcel
```

### Update check

```excel
=NatsUDFUpdateAvailable()
```

Returns `TRUE` when a newer NATSExcel release is known to exist. It **never
blocks and never touches the network on the Excel thread** — it asks for a
background refresh (rate-limited internally) and immediately reports the cached
answer. The background check is controlled by the `"UpdateCheck"` key in
`NATSExcel.json` (default `true`).

---

## 6. Troubleshooting

| Symptom | Cause | Fix |
| --- | --- | --- |
| **`#N/D`** in an `=RTD(...)` cell | The RTD server was not created: the `NatsRtd` COM ProgId is not registered, most often because **Excel is running elevated** — elevated processes ignore per-user COM classes registered under `HKCU`. | Close Excel, load the XLL through the normal (non-elevated) path so `AddIn.AutoOpen` registers `NatsRtd` under `HKCU`, then reopen. On a test machine where Excel must run elevated, register the ProgId under `HKLM\Software\Classes` (admin) pointing at the packed XLL. |
| **`#N/A`** in a write cell | That is the **pending** marker of an asynchronous write (`ExcelAsyncUtil.Observe`): the cell shows `#N/A` and later the real value. | No action — it resolves when the write completes. Only investigate if it stays pending and `=NatsUDFConnectionCount()` is `0` (no connection). |
| **`#CONFIG`** | A formula's host argument is a `Clusters`/`Servers` name that does not resolve — the add-in refuses to fall back to `localhost` silently. | Fix the name in the formula, or add the name to `NATSExcel.json`. Check with `=NatsUDFConfigValidate()`. |
| **`Error: JetStream is not enabled on this server`** | The server was started **without `-js`**. Core pub/sub still works; every JetStream/KV/Object call fails. | Restart the server with `-js` **and** a durable `--store_dir` (see [§2](#2-no-server-yet-run-one-on-windows)). |
| **`No Responders`** | A request/reply reached the server, but **nobody is subscribed** to the reply subject. | Make sure a responder subscribes to that subject before you send the request. This is a frozen marker, not a bug. |
| **`TIMEOUT`** | A responder exists but did not answer within `timeoutMs`. | Raise the `timeoutMs` argument, or check the responder. |
| **A cell does not update / looks stale** | Two deliberate behaviors: real-time pushes are **coalesced** (`CoalesceRealtimeUpdates` + `ConflationMs`, default **200 ms** — the latest value wins per window), and `SkipRepeatedMessages` **drops a repeated payload** (compared by ordinal text, so `"007"` and `"7"` differ). | Lower `ConflationMs` or set `CoalesceRealtimeUpdates` to `false` for a fast feed; set `SkipRepeatedMessages` to `false` if you need every identical value. Note: **wildcard subjects are never deduplicated**. |
| **`Error: admin writes are disabled (set "AllowAdminWrites": true)`** | You called a destructive/admin function with the gate closed (the default). | Set `"AllowAdminWrites": true` in `NATSExcel.json` only if you really want recalculation to be able to mutate server state. |
| **`#VALUE!`** | Excel could not use what the function returned. The add-in renders its own failures as `Error: ...`, `#CONFIG` or `#ERROR: ...`, so a bare `#VALUE!` usually points at Excel's argument coercion rather than at NATS. | First prove the add-in loaded: `=NatsUDFVersion()` (a `0.1.0`-style string) and `=NatsUDFConfigPath()`. Then re-check the cell's arguments. |

More detail in the add-in's own log: `=NatsUDFLogPath()` prints it
(`NATSExcel.log` next to the add-in, rolling at 100 MB, 5 archives; it uses the
credential-masking renderer).

---

## 7. Build from source / test

The repository has a solution file:

```powershell
dotnet build NATSExcel.sln -c Release
```

Outputs land in `bin\Release\net48\publish\`: **`NATSExcel-packed.xll`** (32-bit),
**`NATSExcel64-packed.xll`** (64-bit) and `NATSExcel.dll`, plus the dependency
assemblies packed inside the XLLs. The XLL names come from the `.dna` file name
(`NATSExcel.dna`), so if that file is missing or renamed Excel-DNA falls back to
`NATSExcel-AddIn-packed.xll` and the installer/release expectations break.

The add-in targets **.NET Framework 4.8** with `<LangVersion>9.0</LangVersion>`
plus a small `IsExternalInit` shim, because the NATS.Net option objects use
init-only setters.

### Tests

```powershell
# Unit tests (no NATS, no Excel)
dotnet test test\NATSExcel.Tests\NATSExcel.Tests.csproj -c Release

# Start a disposable local cluster (native nats-server, JetStream on, durable store;
# add -Nodes 3 for the cluster/failover tests). No Docker required.
powershell -ExecutionPolicy Bypass -File test\TestSupport\Start-Cluster.ps1

# Excel end-to-end: loads the packed XLL into a hidden Excel and drives it end to end
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

The E2E script needs Excel, a built XLL and a running server. It uses a
**copy-based add-in folder** (no installer, no COM registration, nothing to
click), retries transient COM rejections and prints a `COM retries` summary,
and writes its workbook to `%TEMP%`. It currently passes **11/11**. Useful
parameters: `-NatsUrl` (default `nats://127.0.0.1:4222`), `-Xll`,
`-SkipNodeKill` (auto-skipped for a non-local host), `-SkipMultiWorkbook`,
`-KeepExcelOpen`, `-Fast`.

Other real layers (see [`test/README.md`](test/README.md)):

```powershell
# Load test against the subscription/broadcast path:
#   url [seconds] [publishers] [listeners] [pattern] [queue]
# A publisher count of 0 means listen-only.
dotnet run --project test\LoadTests -c Release -- "nats://127.0.0.1:4222" 10 2 1

# Liveness / fault test (destructive steps: local or disposable server only)
dotnet run --project test\LivenessTests -c Release -- "nats://127.0.0.1:4222"
```

> **Excel COM automation is inherently flaky, and antivirus/EDR widens the
> window.** Expect a retry count (`RPC_E_CALL_REJECTED`); a non-zero total is
> normal. Misleading errors (`InvalidCastException`, COM hangs) usually mean
> "run it again", not "the add-in is broken". The suites print the active
> antivirus state. Also remember: elevated Excel ignores the per-user `NatsRtd`
> COM registration, so `=RTD("NatsRtd", ...)` can show `#N/D` there.

### Build the installer

```powershell
# Requires Inno Setup 6 (https://jrsoftware.org/isinfo.php) on PATH as iscc.
# The version is injected; use the tag without the leading 'v'.
iscc "/DAppVersion=0.1.0" "installer\NATSExcel.iss"
# -> installer\output\NATSExcelSetup-0.1.0.exe
```

CI builds the installer on every run and attaches
`NATSExcelSetup-<version>.exe` (plus the XLLs, `NATSExcel.dll`, `NLog.config`
and `NATSExcel.json`) to the GitHub release on a `v*` tag, using the matching
`CHANGELOG.md` section as the release notes.

See [`installer/README.md`](installer/README.md) for the installer's silent
parameters (`/URL`, `/AUTH=`, `/TOKEN=`, `/SERVER=`, `/OPENEXCEL=`, ...) and
[`AGENTS.md`](AGENTS.md) for the build/test details.

---

## 8. License

- **NATS.Net** (the client library) is licensed under **Apache-2.0**.
- **NATSExcel's own license is not finalized yet.** There is no `LICENSE` file in
  the repository at this point, so do not assume a grant of rights beyond what
  the repository itself states.
