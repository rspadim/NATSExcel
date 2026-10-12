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

## Golden rules

1. **English only** — code, comments, identifiers, log messages, docs and commit
   messages are written in English.
2. **Never change the public Excel surface** — function names, argument order,
   `[ExcelArgument]`/`[ExcelFunction]` attributes, RTD commands (`SUB`, `KV`,
   `KVGET`, `REQ`, `JS`, `MSG`, `OBJ`) and the ProgId `NatsRtd`. Users'
   spreadsheets depend on every one of them. New additive functions (e.g. the
   `...NonVolatile` write twins) are fine; renames, argument changes and
   removals are not.
3. Keep the `NATSExcel.json` shape backward compatible
   (`Nats` / `RTD` / `UDF` / `Servers` / `Auth` / `JetStream`).
4. All NATS access goes through `NatsRuntime`
   (`NatsConnectionManager` / `NatsSubscriptionManager`). Never create a
   `NatsConnection` anywhere else.
5. Subscription listeners are registered/removed as `IDisposable` tokens.
   Never tear down a shared client subscription while other listeners remain —
   it is ref-counted per `(endpoint, subject, queueGroup)`.
6. Any code running inside a `System.Timers.Timer` callback or a NATS client
   callback must catch and log exceptions. The shared `CreateTimer` helper in
   `NatsRtd.cs` and the subscription callback wrapper already do this.
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

## Layout

| File | Responsibility |
| --- | --- |
| `Config.cs` | `AppConfig`: loads `NATSExcel.json` once, safe defaults, endpoint alias resolution, `MaskNatsEndpoint`. |
| `ConfigSchema.cs` | Embedded JSON Schema for `NATSExcel.json`: validates (comments accepted) and drops bad values so one mistake does not discard the file. |
| `NatsConnectionManager.cs` | Single place that creates/caches `NatsConnection` per resolved endpoint and the JetStream context; reconnect observation; bounded wait helper. |
| `NatsSubscriptionManager.cs` | Ref-counted Core subscriptions: one client subscription per `(endpoint, subject, queueGroup)` broadcasting to N listeners; wildcard-aware; survives reconnects. |
| `NatsRuntime.cs` | Process-wide singleton wiring connections + subscriptions; shutdown on add-in unload; same-process reload reset. |
| `NatsRtd.cs` | `IExcelAddIn` + RTD server lifecycle, topic registries, timers, polling; `NatsRTDStatus` functions. |
| `NatsUdfCore.cs` | Core NATS `[ExcelFunction]` implementations (publish, request/reply, subject-latest, unsubscribe) + `...NonVolatile` twins. |
| `NatsUdfJetStream.cs` | JetStream functions (publish-with-ack, get/fetch, stream/consumer info and admin). |
| `NatsUdfKv.cs` | Key/Value functions (get/put/create/update/delete/keys/history/watch). |
| `NatsUdfObjectStore.cs` | Object Store functions (put/get/info/list/delete/watch). |
| `NatsUdfStatus.cs` | Status/diagnostics functions. |
| `NatsUdfAsync.cs` | Optional async write/read dispatch via Excel-DNA `ExcelAsyncUtil.Observe`; per-endpoint FIFO queue; bounded sync wait. |
| `NatsJson.cs` | Matrix ⇄ JSON conversions. |
| `NatsResultFormatter.cs` | Value formatting sent to Excel. |
| `TickGate.cs` | Non-blocking reentrancy gate for timer callbacks; shared `StripedLocks` helper. |
| `Conflation.cs` | Real-time delivery conflation window (`ConflationMs`): latest value wins per window. |
| `UpdateCheck.cs` | Non-blocking GitHub release check (background task). |
| `installer/` | Inno Setup script (`NATSExcel.iss`) + README: per-user helper installer. Built by CI. |

## Runtime model (why it is the way it is)

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
  `password=`/`user=`/`token=`, NKey seeds and creds-file paths in every error
  cell, RTD `#ERROR`, status cell and log line.
- Admin/destructive operations (`Stream*`, `Consumer*`, `Purge`,
  `DeleteMessage`, KV/Object delete) are gated by `AllowAdminWrites`
  (default `false`) so an accidental recalculation cannot mutate production
  state.
- Values and identifiers (subjects, buckets, keys, stream/consumer names)
  written to NATS always use the invariant culture, regardless of Excel locale.
- The config file is read once per Excel process (restart Excel after editing
  it) and searched in: add-in folder, user profile, Excel folder, `C:\Windows`
  (first existing wins). Comments are accepted; a wrong-typed value is dropped
  individually and the rest of the file is kept.

## Build

```powershell
dotnet build NATSExcel.sln -c Release
```

Outputs in `bin\Release\net48\publish\`:

- `NATSExcel-packed.xll` (32-bit) and `NATSExcel64-packed.xll` (64-bit),
- `NATSExcel.dll`, plus the dependency assemblies packed inside the XLLs.

CI (`.github/workflows/build.yml`) runs on pull requests and main pushes; on
`v*` tags it embeds the version in the `ClientName` for diagnostics, verifies
the top CHANGELOG heading matches the tag, asserts every release asset exists,
uses that CHANGELOG section as the release body, and builds the Inno Setup
installer (a broken `.iss` fails the build on every run, not only a tag).

The CI needs a NATS server. Unlike RedisExcel (which installs Memurai via
Chocolatey), the server is a single self-contained `nats-server.exe` that runs
natively on Windows: download the `nats-server-<ver>-windows-amd64.zip` from the
GitHub releases, start it with `-js --store_dir <temp>`, and probe port 4222
(the same readiness pattern as the Redis `PING` wait). Do **not** use the
community `nats-server` Chocolatey package — it is stale (last published 2019).
A `nats-server --jetstream` probe up front makes a "JetStream disabled" run
attributable. Docker is never required, on CI or on a user machine.

## Tests

### 1. Unit tests (no NATS, no Excel)

```powershell
dotnet test test\NATSExcel.Tests\NATSExcel.Tests.csproj -c Release
```

Covers: `NatsJson` conversions and the matrix size/total-cell budget,
`AppConfig` load/sanitize/pruning and the JSON-Schema per-value drop,
`MaskNatsEndpoint` (URL userinfo, `token=`, NKey seed, creds path, quoted
values), endpoint/alias resolution, the conflation window, the subscription
manager (ref-count, wildcard, queue-group identity, duplicate suppression,
origin counters), the `TickGate` reentrancy helper, `UpdateCheck`, the
`...NonVolatile` signature parity (the volatile write set is pinned), the write
mode (`SyncWrite`/`AsyncWrites`) parsing plus async dispatch (per-endpoint
serial order, the synchronous path, the caller refusal), the offline observable
tests, the bounded-wait helper, and a project-file parity test that keeps the
unit project's `Compile` list complete.

The unit, smoke, load and liveness test projects compile the production sources
directly (linked `Compile` items), so a new production `.cs` needed by tests
must be added to their `Compile` lists; a unit parity test pins the unit
project's list. `dotnet test` does not build or pack the add-in; CI builds it
with msbuild.

### 2. Smoke tests (requires a NATS server with JetStream, no Excel)

```powershell
nats-server -js
dotnet run --project test\SmokeTests -c Release -- "nats://127.0.0.1:4222"
```

Covers the ref-counted Core subscription behavior (two listeners on one
subject, last-listener teardown, re-subscription, wildcards, queue groups,
duplicate suppression, origin counters, argument validation), request/reply
(timeout and `no responders`), JetStream publish/consume (ack, sequence,
durable consumer, replay, purge), Key/Value (put/get/create CAS by revision/
update/delete/history/watch) and Object Store (put/get/info/digest), plus the
real `SyncWrite` modes and the dedup markers. CI runs it against a locally
started `nats-server -js` (skips the JetStream section with a printed reason if
the server was started without `-js`).

### 3. Excel end-to-end

```powershell
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

Uses a hidden Excel instance: loads the packed XLL, builds UDF/RTD test sheets,
asserts values against NATS, then covers the multi-workbook case, a server
restart (connection recovery) and the write modes (`sync`/`fireforget`, async
on and off), and asserts that a `...NonVolatile` write runs once (a worksheet
recalculation must not re-send it).

Useful parameters: `-RepoRoot <path>`, `-NatsUrl <url>`, `-SubjectPrefix`,
`-SkipServerRestart` (auto-skipped for non-local hosts), `-NatsCli <command>`,
`-SyncWrite <mode>`, `-AsyncWrites`, `-KeepExcelOpen`.

For local hosts the workbook is saved to `test\NATSExcel.Test.xlsx` (committed
sample); for remote hosts it is saved to `%TEMP%` so the host never lands in the
repository. The committed workbook only references `nats://localhost:4222` and
`test.natsexcel.*`.

### 4. Load tests (requires NATS)

```powershell
dotnet run --project test\LoadTests -c Release -- manager "nats://127.0.0.1:4222" 10 2 1
```

`manager` exercises the subscription broadcast path; `raw` is the plain
`NATS.Client` baseline. Parameters: url, seconds, publishers, listeners,
subject and an optional queue group. Reports throughput, allocated bytes per
received message and GC counts. Read-only stress runs against real servers are
allowed, but pass the url only as a command-line argument (never commit it) and
keep the runs short.

### 5. Liveness tests (requires NATS; Docker optional)

```powershell
dotnet run --project test\LivenessTests -c Release -- "nats://127.0.0.1:4222"
```

Drives the real managers under adversarial conditions — ThreadPool starvation,
subscribe/dispose churn, server restart/stop-start storms, JetStream consumer
reconnects, async queue/observable bursts, dedup re-join — and enforces liveness
bounds with a dedicated process heartbeat and watchdog thread: a stall past
**5 s** fails the run, delivery/queue must resume within **30 s**, and the async
backlog must drain within **20 s**. With `--container <name>` the suite manages
its own disposable `nats -js` container; `--quick` shortens the attack windows.
Only run it against a disposable/local server: the restart storm is destructive.

## Excel automation lessons (hard-won, inherited from RedisExcel)

These cost real debugging time — read before writing automation.

1. **Excel COM automation is inherently flaky; antivirus/EDR only adds to it.**
   Even with no antivirus, expect a retry count (`RPC_E_CALL_REJECTED`); a retry
   total is NORMAL. Symptoms are misleading errors (`InvalidCastException`,
   `RPC_E_CALL_REJECTED`, COM calls hanging); the E2E retries with backoff and a
   warm-up phase, so **if a run fails this way, run it again**. Never assume such
   an error means the add-in is broken. Every suite prints the active AV state.
2. **Elevated processes ignore per-user COM classes.** The add-in registers
   `NatsRtd` under `HKCU\Software\Classes` in `AutoOpen`. If Excel runs
   elevated, `=RTD("NatsRtd", ...)` shows `#N/D`. Workaround for test machines:
   register under `HKLM\Software\Classes` (admin), pointing at the packed XLL.
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

Push an annotated `v*` tag; the CI workflow builds and publishes the packed
XLLs, `NATSExcel.dll`, `NLog.config`, `NATSExcel.json` and the installer
(`NATSExcelSetup-<version>.exe`) as release assets. The version embedded in the
NATS connection `Name` (shown as `dev` for local builds) is the tag.
