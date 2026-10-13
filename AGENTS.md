# AGENTS.md

Guidance for AI coding agents (and humans) working on NATSExcel.

## What this project is

Excel add-in (XLL) written in C# / .NET Framework 4.8 with Excel-DNA that
integrates **NATS** with Excel, sibling to the **RedisExcel** project:

- **UDFs** — worksheet functions (`NatsUDF*`, `NatsRTD*` status helpers).
- **RTD server** — streaming/polling topics via `=RTD("NatsRtd",, ...)`.
- Coverage: **Core NATS** (pub/sub, request/reply, queue groups, headers) and
  **JetStream** (streams, consumers, replay, Key/Value, Object Store).
- Backends: **NATS.Net** (v3), NLog, Newtonsoft.Json.

Read `DESIGN.md` for the architecture and the full function catalog. It is the
contract; this file is the working guidance.

> **Implementation status (read this first).** The add-in is now largely
> implemented: the packed XLL loads in real Excel and answers UDFs (the P0 gate
> passed), the full `NatsUDF*` catalog (**89** functions) plus the six `NatsRTD*`
> status helpers exist, the RTD server runs, the Excel E2E passes 11/11 and a
> load test measures ~204k msg/s. The unit, load and E2E test layers are real;
> the smoke and liveness layers, the solution file and the release cut are still
> open (see `PLAN.md` §8 and §9). The `## Layout` section below splits what
> exists from what is still planned.

## Golden rules

1. **English only** — code, comments, identifiers, log messages, docs and commit
   messages are written in English.
2. **Never change the public Excel surface** — function names, argument order,
   `[ExcelArgument]`/`[ExcelFunction]` attributes, RTD commands (the frozen set
   is `SUB`, `KV`, `KVGET`, `KVKEYS`, `REQ`, `JS`, `MSG`, `OBJ`, `RTDCLUSTER`;
   the grammar tests in `test\NATSExcel.Tests\NatsRtdTopicGrammarTests.cs` pin
   the exact command list and error text) and the ProgId `NatsRtd`. Users'
   spreadsheets depend on every one of them. New additive functions (e.g. the
   `...NonVolatile` write twins) are fine; renames, argument changes and
   removals are not.
3. Keep the `NATSExcel.json` shape backward compatible. The shipped sample and
   `NATSExcel.schema.json` define the real shape: `Default` (a plain **name**
   string — a `Clusters` name, a `Servers` name or a raw `nats://` URL),
   `Servers`, `Clusters` (each with `urls` plus optional `timeout`, `auth`,
   `reconnect`, `jetstream`, `expectCluster`), a shared top-level `Auth` block
   and a shared top-level `Reconnect` block, `RTD` (delivery tuning only — there
   is no `UDF` block and no RTD/UDF server split), `JetStream`, plus the flat
   toggles `AllowAdminWrites`, `SkipRepeatedMessages`, `CoalesceRealtimeUpdates`,
   `ConflationMs`, `PublishDedupCacheSize`, `SyncWrite`, `AsyncWrites`,
   `UpdateCheck`. A cluster's `auth` merges OVER the shared `Auth` field by
   field. (An earlier draft nested the settings under `Default` with a top-level
   `Auth`/`UDF`; `Config.cs` accepts a nested `Default` only as a back-compat
   fallback and never reads a `UDF` block.)
4. All NATS access goes through `NatsRuntime`
   (`NatsConnectionManager` / `NatsSubscriptionManager`); never create a
   `NatsConnection` anywhere else. This is the only place the process-wide
   managers are built, shut down and recreated after a same-process reload.
5. Subscription listeners are registered/removed as `IDisposable` tokens. Never
   tear down a shared client subscription while other listeners remain — it is
   ref-counted per `(endpoint, subject, queueGroup)`.
6. Any code running inside a `System.Timers.Timer` callback or a NATS client
   callback must catch and log exceptions. The shared `CreateTimer` helper (in
   `NatsRtd.cs`) and the subscription callback wrapper do this; keep it that way.
7. State touched by both the Excel thread and NATS/timer threads must be
   thread-safe (`ConcurrentDictionary`, `Interlocked`, or a private lock).
8. Keep `net48`. Use **`<LangVersion>9.0</LangVersion>`** plus a small
   `IsExternalInit` shim: the NATS.Net option objects use init-only setters
   (C# 9), and `net48` does not ship that type. Verified by a spike (publish +
   `await foreach` round-trip on `net48`); see `DESIGN.md` §13.1 and §13.6.
9. **No private data in tracked files** — no employer or client company names,
   no company-owned domains or internal hostnames, no customer names, no
   credentials and no local machine paths. This applies to **everything** that
   lands in the repository: prose, code comments, XML doc, log messages, commit
   messages, docs, sample configs, test workbooks and benchmark arguments. Only
   `localhost`, `127.0.0.1` and reserved documentation domains (e.g.
   `nats.example`) may appear; examples use `nats://localhost:4222` and relative
   paths. A company name in a shipped doc or a sample config is a leak — grep
   the whole tree before every commit and again before every release:

   ```powershell
   git grep -n -i -E '<company-name>|<company-domain>'   # must return nothing
   ```
10. **Never block the Excel thread unbounded.** NATS .NET is async-only; every
    blocking wait goes through the bounded helper or `ExcelAsyncUtil.Observe`.
    Never call `.GetAwaiter().GetResult()` directly on the Excel thread.
11. **NKey seed: the field is text *or* a `.nk` file path — read the file
    first.** `nkeySeed` accepts either the seed text itself (an `SU...` string)
    or a path to a `.nk` file. The NATS.Net client wants the seed, not the
    path: handing it a path fails with `Invalid base32 character` (the path is
    not base32). Read a readable file and pass its trimmed contents; otherwise
    pass the value through. **Implemented** in `AppConfig.ResolveSeed`, called
    from `NatsConnectionManager.BuildAuth`. The installer wizard labels this
    field exactly this way ("the text, or a path to the .nk file"). The
    reference implementation lives in `tools/NatsValidate/ConnectionTester.cs`
    (`ReadSeed`); keep `ResolveSeed` in step with it.

## Layout

**The add-in is largely implemented: it compiles, packs, loads in real Excel and
answers the whole UDF catalog.** The RTD server, the runtime and the config
layer are code. What remains planned is the smoke/liveness test layers, the
solution file and a few audit items (`PLAN.md` §8/§9). Do not describe a planned
file as if it exists.

### Exists today (ground truth)

| File | Responsibility |
| --- | --- |
| `AddIn.cs` | `IExcelAddIn` entry point (`AutoOpen`/`AutoClose`): registers `MaskedExceptionRenderer`, the `NatsRtd` COM server, the runtime/RTD/UDF reload resets; `AutoClose` shuts the runtime down. |
| `Config.cs` | `AppConfig`: loads `NATSExcel.json` once (`Lazy`), comment-aware parse (`StripComments`), `Default`/`Servers`/`Clusters` resolution (`ResolveEndpointCore`), per-cluster `ResolveSettings`, `MaskNatsEndpoint`, `ResolveSeed`, `Sanitize`, `Reload`, `Version`. |
| `ConfigSchema.cs` | Embedded JSON-Schema walk (`ParseWithComments`/`Clean`) with per-value drop; wired into the load path. |
| `NatsUdf.cs` | The seed diagnostics: `NatsUDFVersion`, `NatsUDFConfigPath`, `NatsUDFConnectionCount`, `NatsUDFLogPath`. |
| `NatsConnectionManager.cs` | One cached `NatsConnection` per resolved endpoint; JetStream context cache; server-info snapshot; bounded `WaitBounded*`; eviction/down-sweep; bounded shutdown. |
| `NatsSubscriptionManager.cs` | Ref-counted Core subscriptions (one client sub per `(endpoint, subject, queueGroup)`, N listeners, wildcard-aware, origin tags); per-key epochs; dedup. |
| `NatsRuntime.cs` | Process-wide singleton wiring connections + subscriptions; shutdown tombstone + reload reset. |
| `NatsRtd.cs` | The `NatsRtd` `ExcelRtdServer`: topic registry, poll/flush timers, backoff, all 9 topic kinds; plus the six `NatsRTD*` status helpers. |
| `NatsRtdTopic.cs` | RTD topic grammar/parser (`NatsRtdTopicSpec.Parse`), no Excel/COM dependency. |
| `NatsUdfCore.cs` | Core NATS `[ExcelFunction]`s (publish, request/reply, subject-latest, unsubscribe) + twins. |
| `NatsUdfJetStream.cs` | JetStream functions (publish-with-ack, get/fetch, stream/consumer info and admin). |
| `NatsUdfKv.cs` | Key/Value functions (get/put/create/update/delete/keys/history/watch). |
| `NatsUdfObjectStore.cs` | Object Store functions (put/get/info/list/delete/watch). |
| `NatsUdfStatus.cs` | Status/diagnostics + cluster/server topology functions and the config functions. |
| `NatsUdfCommon.cs` | Shared validation, error strings, markers, admin/JetStream gates, ranges, write-mode helpers. |
| `NatsUdfAsync.cs` | Async dispatch via Excel-DNA `ExcelAsyncUtil.Observe`; per-endpoint FIFO queue; bounded sync wait; caller refusal. |
| `NatsJson.cs` | Matrix ⇄ JSON conversions (`MatrixToJson`/`JsonToMatrix`, internal helpers reached by the `...JSON` functions). |
| `NatsResultFormatter.cs` | Value formatting sent to Excel (ordinal `SameMessage`, `DecodePayload`, 32767 bound). |
| `TickGate.cs` | Non-blocking reentrancy gate for timer callbacks; `StripedLocks`. |
| `Conflation.cs` | Real-time delivery conflation window (`ConflationMs`): latest value wins per window. |
| `MaskedExceptionRenderer.cs` | NLog `${maskedException}` layout renderer: masks credentials in every logged exception. |
| `UpdateCheck.cs` | Non-blocking GitHub release check (background task). |
| `IsExternalInit.cs` | C# 9 shim so net48 accepts init-only setters. |
| `NATSExcel.csproj` | Add-in project: net48, `LangVersion 9.0`, packs the XLLs, copies `NATSExcel.dll` into `publish\`. |
| `NATSExcel.dna` | Excel-DNA packing manifest (29 dependency DLLs); its **file name** sets the packed XLL name. |
| `NATSExcel.json` / `NATSExcel.schema.json` | Sample config + embedded JSON Schema (validated by `ConfigSchema.cs` and `NatsUDFConfigValidate`). |
| `NLog.config` | Logging to `NATSExcel.log`, rolling at 100 MB, 5 archives; uses `${maskedException:format=toString}` (not `${exception:...}`). |
| `installer/` | Inno Setup script (`NATSExcel.iss`) + helper PowerShell + README: per-user, native-first installer. |
| `tools/NatsValidate/` | C# console validator that runs today (see `## Tests`). |
| `test/` | The unit, load and E2E layers + the cluster harness + the copy-based add-in folder helper (see `## Tests`). |

### Planned (from `PLAN.md`; none of these exist yet)

| Planned item | Responsibility (per the plan) |
| --- | --- |
| `NATSExcel.sln` | Solution file indexing the add-in and test projects (PLAN F0.1); the project file is the build input until it lands. |
| `test\SmokeTests\` | Real-`nats-server -js` smoke layer (PLAN V8.3). |
| `test\LivenessTests\` | Adversarial liveness/fault layer with a heartbeat watchdog (PLAN V8.5). |
| `docs\PUBLIC-SURFACE.md` | The per-function signature catalog (PLAN S3.1). |
| `test\...\PublicSurfaceParityTests.cs` + baseline | The public-surface baseline test (PLAN S7.1/S7.2) and the unbounded-blocking scan (V7.3). |
| `test\Run-ExcelE2E-Matrix.ps1` / `Compare-ExcelE2E.ps1` / `Set-Antivirus.ps1` | The AV matrix/compare helpers (PLAN V8.7); only read-only `TestSupport\Get-AntivirusState.ps1` exists today. |

When a planned item lands, move its row up into "Exists today"; `DESIGN.md` §5.1
and this layout must stay in sync (PLAN S7.4).

## Runtime model (why it is the way it is)

> This is the runtime design and the reasoning behind it. It is what the code
> does today (the runtime milestone landed); the few items still not
> implemented are called out inline. When in doubt, read the code.

- Excel may create more than one RTD server per process. Connections and
  subscriptions are **process-wide** (`NatsRuntime`); topics are per RTD server
  instance. Same model as RedisExcel.
- **One `NatsConnection` per resolved endpoint**, not per host/pool: the NATS
  client multiplexes everything and reconnects internally, so the RedisExcel
  per-pool cache and idle-down machinery are intentionally absent. The client
  holds the URL list and handles cluster discovery.
- **Async-only client.** Every NATS call is a `Task`. RTD polled topics read on
  the poll timer; RTD subscription topics push from the client callback; UDF
  calls that must return in-cell run on a worker with a bounded wait; UDF calls
  that may be asynchronous use `ExcelAsyncUtil.Observe` (pending `#N/A`, then the
  value). This is the RedisExcel `AsyncWrites` path promoted to the default.
- **Subjects, not channels.** NATS wildcards are part of the subject
  (`orders.*` one token, `orders.>` one or more); there is no separate pattern
  command. Publishing requires a concrete subject.
- **Queue groups** load-balance a subject across subscribers; a listener may
  declare one. They are part of the subscription identity
  `(endpoint, subject, queueGroup)`, so a plain subscriber and a queue-group
  subscriber on the same subject are **different** physical subscriptions.
- **JetStream is server-side.** Files are the same binary; JetStream is enabled
  with `-js` (or `jetstream {}`). When the server has it disabled, `JS`/`KV`/
  `OBJ` topics and functions return a clear `#ERROR`/`Error:` naming JetStream,
  never a hang or a silent empty result. Note the `store_dir` default is a temp
  path — operators must point it at durable storage (see `DESIGN.md` §11).
- Core Pub/Sub delivery is at-most-once and unordered across overlapping
  messages (the client dispatches callbacks on the thread pool), exactly like
  RedisExcel's StackExchange.Redis caveat: order-sensitive feeds must carry a
  sequence field and compare it in the formula. JetStream adds sequence numbers
  and revisions natively — prefer them for ordering.
- `SkipRepeatedMessages` compares payloads by **ordinal text** (so `"007"` and
  `"7"` differ) before decoding, for literal subjects and polled reads. Wildcard
  subjects are never deduplicated (channels/subjects interleave).
- Real-time pushes are coalesced by default (`CoalesceRealtimeUpdates`) and
  governed by the explicit `ConflationMs` window: the latest value wins per
  window. This is the main lever against flicker and the recalc cascade.
- Writes are configurable (`SyncWrite`): `sync` awaits the JetStream ack /
  flush; `fireforget` (default) issues result-agnostic publishes without
  awaiting the ack and returns a marker. `ChannelPublishIfChanged`-style dedup
  remembers the last payload only after a confirmed delivery in `sync` mode.
- **Credential masking is mandatory.** `MaskNatsEndpoint` redacts URL userinfo,
  `password=`/`user=`/`token=`, JWTs, NKey seeds, creds-file paths and TLS cert
  paths in every error cell, RTD `#ERROR`, status cell and log line.
  `MaskedExceptionRenderer` (NLog `${maskedException:...}`) applies the same
  masking to logged exceptions; `NLog.config` must keep using it rather than
  `${exception:...}`.
- Admin/destructive operations (`Stream*`, `Consumer*`, `Purge`,
  `DeleteMessage`, KV/Object delete) are gated by `AllowAdminWrites`
  (default `false`) so an accidental recalculation cannot mutate production
  state.
- Values and identifiers (subjects, buckets, keys, stream/consumer names)
  written to NATS always use the invariant culture, regardless of Excel locale.
- The config file **is** read once per Excel process at startup, but
  `NatsUDFConfigReload` re-reads it without a restart. `Config.cs`
  `CandidatePaths` searches: add-in folder, user profile (`%USERPROFILE%`),
  `C:\Windows` (first existing wins). `DESIGN.md` also lists an **Excel folder**
  between the profile and `C:\Windows`; it is designed but not implemented.
  Comments are accepted (string-aware `StripComments`). A wrong-typed value is
  dropped individually by `ConfigSchema.cs` (the schema walk), so the rest of the
  file is kept; only a structural (not-JSON) failure falls back to defaults.

## Build

```powershell
dotnet build NATSExcel.csproj -c Release
```

There is **no `NATSExcel.sln` yet** (PLAN F0.1 will add it); until then build
the project directly. When the solution lands, `dotnet build NATSExcel.sln -c
Release` becomes the command.

Outputs in `bin\Release\net48\publish\`:

- `NATSExcel-packed.xll` (32-bit) and `NATSExcel64-packed.xll` (64-bit),
- `NATSExcel.dll` (placed there by the `CopyNatsExcelDllToPublish` target),
  plus the dependency assemblies packed inside the XLLs.

The XLL names come from the `.dna` file name: `NATSExcel.dna` →
`NATSExcel-packed.xll`. **Trap:** if `NATSExcel.dna` is missing or renamed,
Excel-DNA falls back to `<ProjectName>-AddIn-packed.xll`
(`NATSExcel-AddIn-packed.xll`); the installer, the release asset list
and the E2E script all expect `NATSExcel-packed.xll` and would break.
Verify both names in `publish\` after any `.dna` change.

> **CI (PLAN P10) — present, not yet run on GitHub.** `.github/workflows/build.yml`
> runs on PRs, main pushes and `v*` tags. It builds the add-in with `dotnet build`
> (the project file, since there is no `NATSExcel.sln` yet), runs the unit tests,
> provisions `nats-server -js` and probes port 4222 for `jetstream:true`, builds
> the Inno Setup installer on every run, and on a `v*` tag checks the top
> CHANGELOG heading against the tag, asserts the release assets, uses the
> CHANGELOG section as the release body and publishes the release. The smoke and
> liveness steps are present but skip while those test projects do not exist.

The CI provisions NATS as a single self-contained `nats-server.exe` (no
Memurai/Chocolatey, unlike RedisExcel): download the
`nats-server-<ver>-windows-amd64.zip` from the GitHub releases, start it with
`-js --store_dir <temp>`, and probe port 4222 (the same readiness pattern as
the Redis `PING` wait). Do **not** use the community `nats-server` Chocolatey
package — it is stale (last published 2019). The `jetstream:true` probe up front
makes a "JetStream disabled" run attributable. Docker is never required, on CI
or on a user machine.

## Tests

The real layers are `test\NATSExcel.Tests` (unit), `test\LoadTests` (load) and
`test\Run-ExcelE2E.ps1` (Excel E2E), plus the cluster harness and helpers under
`test\TestSupport\`. The smoke and liveness projects are still `PLAN.md` work and
do **not** exist — the `## Tests` section below marks each layer.

The single reference is `test\README.md`; keep it and this section in sync. Test
projects compile the production sources directly (linked `Compile` items from
`test\TestSupport\ProductionSources.props`), not via `ProjectReference`, so a
new production `.cs` must be added to that props; `ProjectFileParityTests`
fails if a root production file is missing.

To iterate without the installer, use the copy-based helper
`test\TestSupport\New-AddInFolder.ps1`: it copies the build output (the XLL, the
dll, `NLog.config`, the schema and a starter config) into a throwaway folder that
Excel loads via `RegisterXLL` — no wizard, no COM registration, nothing to
click. Prefer it over the installer for a test iteration.

### 0. Config + connection validator (real, runs today)

`tools/NatsValidate` is a C# console app that **exists and works**. It is the
recommended way to check a `NATSExcel.json` and, with `--test`, to prove the
connection: it uses the **same NATS.Net client the add-in uses**
(`NATS.Client.Core` 3.3.0), so every credential style is really exercised
(token, user/password, NKey seed, JWT + seed, creds file, mTLS). The PowerShell
helpers in `installer/` (`validate-config.ps1`, `test-connection.ps1`) only
speak the wire protocol by hand and cover token and user-password; the C# tool
is the complete check.

```powershell
# Schema/structural check only (no server needed)
dotnet run --project tools\NatsValidate -c Release -- NATSExcel.json

# Also connect with the NATS.Net client
dotnet run --project tools\NatsValidate -c Release -- NATSExcel.json --test
```

A bare path argument is the config file; `--config`/`-c`, `--timeout <ms>`,
`--test`/`-t` and `--help` are accepted. With no path it searches this folder,
`%USERPROFILE%`, `..`, then `C:\Windows`. Exit codes: `0` OK, `1` config
problems, `2` unreadable, `3` dangling cluster/server reference, `4` connection
failed.

### 1. Unit tests (no NATS, no Excel) — real (PLAN V8.2)

```powershell
dotnet test test\NATSExcel.Tests\NATSExcel.Tests.csproj -c Release
```

Covers today: `NatsJson` conversions, `AppConfig` load/sanitize/`StripComments`,
the JSON-Schema per-value drop (via the sample), `MaskNatsEndpoint` and
`MaskedExceptionRenderer`, the conflation window, the RTD grammar
(`NatsRtdTopicGrammarTests`), the UDF argument validation/markers/error strings
(`NatsUdfCommonTests`), and a project-file parity test that keeps the unit
project's `Compile` list complete. Growth areas (not yet covered): subscription
identity, connection offline, `WaitBounded`, the async bridge, twin signature
parity, `UpdateCheck`.

`dotnet test` does not build or pack the add-in; CI builds it separately.

### 2. Smoke tests (planned — PLAN V8.3; requires a NATS server with JetStream, no Excel)

`test\SmokeTests` does **not** exist yet. When it lands it will cover the
ref-counted Core subscription behavior (two listeners on one subject,
last-listener teardown, re-subscription, wildcards, queue groups, duplicate
suppression, origin counters, argument validation), request/reply (timeout and
`no responders`), JetStream publish/consume (ack, sequence, durable consumer,
replay, purge), Key/Value (put/get/create CAS by revision/update/delete/
history/watch) and Object Store (put/get/info/digest), plus the real `SyncWrite`
modes and the dedup markers. `tools\NatsValidate --test` already proves the
credential styles the PowerShell wire check cannot (NKey/JWT/creds/mTLS).

### 3. Excel end-to-end — real (PLAN V8.6)

```powershell
# Start a local cluster first (or point at one):
powershell -ExecutionPolicy Bypass -File test\TestSupport\Start-Cluster.ps1
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

Passes 11/11 locally. Uses a hidden Excel instance: loads the packed XLL through
the copy helper (no installer), builds UDF/RTD test sheets, and covers the UDF
round trip, the RTD `SUB` push (two updates), the JetStream/KV named-error
paths, the multi-workbook case, the node-kill failover, and an Excel
restart/reload. `...NonVolatile` single-run and the write modes are exercised
via the publish cells.

Useful parameters: `-RepoRoot <path>`, `-NatsUrl <url>` (default
`nats://127.0.0.1:4222`), `-SubjectPrefix` (default `test.natsexcel`), `-Xll
<path>`, `-SkipNodeKill` (auto-skipped for a non-local host),
`-SkipMultiWorkbook`, `-KeepExcelOpen`, `-Fast`.

Every wait is condition-based; the script retries transient COM rejections and
prints a `COM retries` summary (a retry total is normal — see the Excel
automation lessons below). It saves its workbook to `%TEMP%`, so the host never
lands in the repository. The cluster harness and the sample only reference
`127.0.0.1`/`localhost` and `test.natsexcel.*`.

### 4. Load tests — real (PLAN V8.4; requires NATS)

```powershell
dotnet run --project test\LoadTests -c Release -- manager "nats://127.0.0.1:4222" 10 2 1
```

`manager` exercises the subscription broadcast path; `raw` is the plain
`NATS.Client` baseline. Parameters: url, seconds, publishers, listeners,
subject/pattern and an optional queue group. Reports throughput and allocated
bytes per received message. ~204k msg/s measured on a local manager run.
Read-only stress runs against real servers are allowed, but pass the url only as
a command-line argument (never commit it) and keep the runs short.

### 5. Liveness tests (planned — PLAN V8.5; requires NATS; Docker optional)

`test\LivenessTests` does **not** exist yet. When it lands it will drive the real
managers under adversarial conditions — ThreadPool starvation, subscribe/dispose
churn, server restart/stop-start storms, JetStream consumer reconnects, async
queue/observable bursts, dedup re-join — and enforce liveness bounds with a
dedicated process heartbeat and watchdog thread: a stall past **5 s** fails the
run, delivery/queue must resume within **30 s**, and the async backlog must drain
within **20 s**. `--quick` shortens the attack windows. Only run it against a
disposable/local server: the restart storm is destructive. (The cluster harness
and the E2E node-kill are the real failover coverage today.)

## Excel automation lessons (hard-won, inherited from RedisExcel)

These cost real debugging time — read before writing automation. The E2E script
and cluster harness exist (`test\Run-ExcelE2E.ps1`, `test\TestSupport\`); the
smoke/liveness suites do not. Several lessons describe add-in behavior that now
runs (the `NatsRtd` registration) rather than intent.

1. **Excel COM automation is inherently flaky; antivirus/EDR only adds to it.**
   Even with no antivirus, expect a retry count (`RPC_E_CALL_REJECTED`); a retry
   total is NORMAL. Symptoms are misleading errors (`InvalidCastException`,
   `RPC_E_CALL_REJECTED`, COM calls hanging); the E2E retries with backoff and a
   warm-up phase, so **if a run fails this way, run it again**. Never assume such
   an error means the add-in is broken. Every suite prints the active AV state.
2. **Elevated processes ignore per-user COM classes.** The add-in registers
   `NatsRtd` under `HKCU\Software\Classes` in `AutoOpen` (via
   `ComServer.DllRegisterServer`). If Excel runs elevated, `=RTD("NatsRtd", ...)`
   can show `#N/D`. Workaround for test machines: register under
   `HKLM\Software\Classes` (admin), pointing at the packed XLL.
3. **A server restart is destructive to other users** — only stop/restart a
   disposable/local NATS server; the E2E script skips the restart step for
   non-local hosts.
4. **Never kill `EXCEL.EXE` processes blindly** — the user may have workbooks
   open. The E2E script creates its own hidden instance and quits it.
5. **PowerShell + COM quirks:** cast value types explicitly on the same call
   site (`[string]$value` / `[double]$value`); prefer `Range("A1")` addresses
   over `Cells.Item(r, c)`; pass scriptblock variables explicitly when in doubt.
6. **RTD cells display `#N/D` when the RTD server could not be created** —
   check COM registration before suspecting the add-in logic.
7. **NATS-specific:** a server started without `-js` accepts connections and
   Core pub/sub works, but every JetStream call fails — assert the JetStream
   probe up front and print whether `-js` is on, so a "JS disabled" run is
   attributable rather than looking like an add-in bug. The JetStream
   `store_dir` default is a temp path; the CI container and the installer's
   server option must mount a durable directory or a restart drops the streams.

## Release

The CI workflow exists (`.github/workflows/build.yml`); `CHANGELOG.md` exists
but still holds only an `Unreleased` placeholder — no release has been cut. The
flow: push an annotated `v*` tag; CI builds and publishes the packed XLLs,
`NATSExcel.dll`, `NLog.config`, `NATSExcel.json` and the installer
(`NATSExcelSetup-<version>.exe`) as release assets, using the CHANGELOG section
for the tag as the release body. The version embedded in the NATS connection
`Name` (shown as `dev` for local builds) is the tag. `INFORMATIONALVERSION`
defaults to `dev` in `NATSExcel.csproj`.
