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

> **Implementation status (read this first).** The catalog above is the design
> target, not the current build. What actually exists today: the add-in project
> packs and loads, `Config.cs` loads `NATSExcel.json`, and `NatsUdf.cs` exposes
> three diagnostic functions (`NatsUDFVersion`, `NatsUDFConfigPath`,
> `NatsUDFConnectionCount`). There is no RTD server, no NATS connection, no
> test project and no CI yet — everything else is in `PLAN.md`. The `## Layout`
> section below splits what exists from what is planned.

## Golden rules

1. **English only** — code, comments, identifiers, log messages, docs and commit
   messages are written in English.
2. **Never change the public Excel surface** — function names, argument order,
   `[ExcelArgument]`/`[ExcelFunction]` attributes, RTD commands (the plan's
   frozen set is `SUB`, `KV`, `KVGET`, `KVKEYS`, `REQ`, `JS`, `MSG`, `OBJ`;
   `DESIGN.md` §7.9 adds a 9th, `NatsRTDCluster` — the baseline test is the
   authority) and the ProgId `NatsRtd`. Users' spreadsheets depend on every one
   of them. New additive functions (e.g. the `...NonVolatile` write twins) are
   fine; renames, argument changes and removals are not.
3. Keep the `NATSExcel.json` shape backward compatible. The shipped sample and
   `NATSExcel.schema.json` define the real shape: `Default` (with nested `auth`
   and `reconnect`), `RTD`, `UDF`, `Servers`, `Clusters`, `JetStream`, plus the
   flat toggles `AllowAdminWrites`, `SkipRepeatedMessages`,
   `CoalesceRealtimeUpdates`, `ConflationMs`, `PublishDedupCacheSize`,
   `SyncWrite`, `AsyncWrites`, `UpdateCheck`. (An earlier draft called the
   first section `Nats` with a top-level `Auth`; `Config.cs` and the sample do
   not read those.)
4. *(Planned — target of the runtime milestone.)* All NATS access will go
   through `NatsRuntime` (`NatsConnectionManager` / `NatsSubscriptionManager`);
   never create a `NatsConnection` anywhere else. Today only the
   `NatsConnection.cs` stub exists.
5. *(Planned.)* Subscription listeners are registered/removed as `IDisposable`
   tokens. Never tear down a shared client subscription while other listeners
   remain — it is ref-counted per `(endpoint, subject, queueGroup)`.
6. *(Planned.)* Any code running inside a `System.Timers.Timer` callback or a
   NATS client callback must catch and log exceptions. The shared
   `CreateTimer` helper (planned in `NatsRtd.cs`) and the subscription
   callback wrapper must do this.
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
    pass the value through. The installer wizard labels this field exactly this
    way ("the text, or a path to the .nk file"), so the code must honor it.
    The reference implementation lives in `tools/NatsValidate/ConnectionTester.cs`
    (`ReadSeed`); copy that behavior when the add-in auth model lands.

## Layout

**The repository is at an early milestone: the add-in compiles, packs and
loads, config loading and a few diagnostic UDFs work.** The runtime, the RTD
server and almost the entire UDF catalog below are still `PLAN.md` targets, not
code. Do not describe a planned file as if it exists.

### Exists today (ground truth)

| File | Responsibility |
| --- | --- |
| `AddIn.cs` | `IExcelAddIn` entry point (`AutoOpen`/`AutoClose`); currently only logs load/unload. RTD/COM `NatsRtd` registration and runtime reset are planned (PLAN S4.11). |
| `Config.cs` | `AppConfig`: loads `NATSExcel.json` once (`Lazy`), comment-aware parse (`StripComments`), safe defaults (`nats://localhost:4222`), `Servers`/`Clusters` alias resolution (`ResolveUrl`), `Version`. `MaskNatsEndpoint` and the auth model are still planned (PLAN F1.5). |
| `NatsUdf.cs` | The UDF class so far: `NatsUDFVersion`, `NatsUDFConfigPath`, `NatsUDFConnectionCount` (a stub returning `0`). |
| `NatsConnection.cs` | Stub for a raw NATS protocol client (first milestone); the planned NATS.Net runtime replaces it. |
| `IsExternalInit.cs` | C# 9 shim so net48 accepts init-only setters. |
| `NATSExcel.csproj` | Add-in project: net48, `LangVersion 9.0`, packs the XLLs, copies `NATSExcel.dll` into `publish\`. |
| `NATSExcel.dna` | Excel-DNA packing manifest; its **file name** sets the packed XLL name. |
| `NATSExcel.json` / `NATSExcel.schema.json` | Sample config + embedded JSON Schema (loaded but not yet validated by code). |
| `NLog.config` | Logging to `NATSExcel.log`, rolling at 100 MB, 5 archives. |
| `installer/` | Inno Setup script (`NATSExcel.iss`) + helper PowerShell + README: per-user, native-first installer. |
| `tools/NatsValidate/` | C# console validator that runs today (see `## Tests`). |

### Planned (from `PLAN.md`; none of these exist yet)

| Planned file | Responsibility (per the plan) |
| --- | --- |
| `ConfigSchema.cs` | Embedded JSON Schema validation with per-value drop. |
| `NatsConnectionManager.cs` | Single place that creates/caches `NatsConnection` per resolved endpoint + the JetStream context; reconnect observation; bounded wait helper. |
| `NatsSubscriptionManager.cs` | Ref-counted Core subscriptions: one client subscription per `(endpoint, subject, queueGroup)` broadcasting to N listeners; wildcard-aware; survives reconnects. |
| `NatsRuntime.cs` | Process-wide singleton wiring connections + subscriptions; shutdown/reload reset. |
| `NatsRtd.cs` | RTD server lifecycle, topic registries, timers, polling; `NatsRTDStatus` functions. |
| `NatsRtdTopic.cs` | RTD topic grammar/parser, no Excel/COM dependency (PLAN S3.2). |
| `NatsUdfCore.cs` | Core NATS `[ExcelFunction]` implementations (publish, request/reply, subject-latest, unsubscribe) + `...NonVolatile` twins. |
| `NatsUdfJetStream.cs` | JetStream functions (publish-with-ack, get/fetch, stream/consumer info and admin). |
| `NatsUdfKv.cs` | Key/Value functions (get/put/create/update/delete/keys/history/watch). |
| `NatsUdfObjectStore.cs` | Object Store functions (put/get/info/list/delete/watch). |
| `NatsUdfStatus.cs` | Status/diagnostics functions. |
| `NatsUdfCommon.cs` | Shared validation, error strings, markers, admin/JetStream gates (PLAN S3.3). |
| `NatsUdfAsync.cs` | Async dispatch via Excel-DNA `ExcelAsyncUtil.Observe`; per-endpoint FIFO queue; bounded sync wait. |
| `NatsJson.cs` | Matrix ⇄ JSON conversions. |
| `NatsResultFormatter.cs` | Value formatting sent to Excel. |
| `TickGate.cs` | Non-blocking reentrancy gate for timer callbacks; shared `StripedLocks` helper. |
| `Conflation.cs` | Real-time delivery conflation window (`ConflationMs`): latest value wins per window. |
| `UpdateCheck.cs` | Non-blocking GitHub release check (background task). |

When a planned file lands, move its row up into "Exists today"; `DESIGN.md` §5.1
and the two `NatsUdf*` additions above must stay in sync (PLAN S7.4).

## Runtime model (why it is the way it is)

> This is the target design and the reasoning behind it; most of it is **not
> implemented yet** (see the status note and `## Layout` above). It is the
> contract the runtime milestone must satisfy, not a description of current
> behavior.

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
- **Credential masking is mandatory (planned).** `MaskNatsEndpoint` will redact
  URL userinfo, `password=`/`user=`/`token=`, NKey seeds and creds-file paths in
  every error cell, RTD `#ERROR`, status cell and log line. It does not exist in
  `Config.cs` yet; the only masking that runs today is in
  `tools/NatsValidate/ConfigLoader.cs` (`MaskUrl`).
- Admin/destructive operations (`Stream*`, `Consumer*`, `Purge`,
  `DeleteMessage`, KV/Object delete) are gated by `AllowAdminWrites`
  (default `false`) so an accidental recalculation cannot mutate production
  state.
- Values and identifiers (subjects, buckets, keys, stream/consumer names)
  written to NATS always use the invariant culture, regardless of Excel locale.
- The config file **is** read once per Excel process (restart Excel after
  editing it) and, in the current code (`Config.cs` `CandidatePaths`), searched
  in: add-in folder, user profile (`%USERPROFILE%`), `C:\Windows` (first
  existing wins). `DESIGN.md` also lists an **Excel folder** between the profile
  and `C:\Windows`; it is designed but not implemented. Comments are accepted
  (string-aware `StripComments`). The per-value drop on a wrong type is
  **planned** (`ConfigSchema.cs`/`Sanitize`); today a parse error on the first
  existing candidate falls back to defaults for the whole file.

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
(`NATSExcel-AddIn-packed.xll`); the installer, the planned release asset list
and the planned E2E script all expect `NATSExcel-packed.xll` and would break.
Verify both names in `publish\` after any `.dna` change.

> **Planned (PLAN P10), not present yet.** There is no `.github/workflows`,
> no `CHANGELOG.md` and no packed-XLL CI today. The intended CI will run on PRs
> and main pushes; on `v*` tags it will embed the version in the `ClientName`,
> check the top CHANGELOG heading against the tag, assert the release assets,
> use that CHANGELOG section as the release body, and build the Inno Setup
> installer (a broken `.iss` will fail the build on every run, not only a tag).

When the CI lands it must provision a NATS server. Unlike RedisExcel (which
installs Memurai via Chocolatey), the server is a single self-contained
`nats-server.exe` that runs natively on Windows: download the
`nats-server-<ver>-windows-amd64.zip` from the GitHub releases, start it with
`-js --store_dir <temp>`, and probe port 4222 (the same readiness pattern as
the Redis `PING` wait). Do **not** use the community `nats-server` Chocolatey
package — it is stale (last published 2019). A `nats-server --jetstream` probe
up front makes a "JetStream disabled" run attributable. Docker is never
required, on CI or on a user machine.

## Tests

**No test project exists yet.** There is no `test\` directory, no unit/smoke/
load/liveness/E2E project and no `ProductionSources.props`; they are PLAN §V8
work. Do not tell anyone to run `dotnet test` — it would fail to find a project.

### 0. Config + connection validator (real, runs today)

`tools/NatsValidate` is a C# console app that **exists and works**. It is the
recommended way to check a `NATSExcel.json` and, with `--test`, to prove the
connection: it uses the **same NATS.Net client the add-in will use**
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

### 1. Unit tests (no NATS, no Excel) — planned (PLAN V8.2)

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

### 2. Smoke tests (planned — PLAN V8.3; requires a NATS server with JetStream, no Excel)

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

### 3. Excel end-to-end (planned — PLAN V8.6)

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

### 4. Load tests (planned — PLAN V8.4; requires NATS)

```powershell
dotnet run --project test\LoadTests -c Release -- manager "nats://127.0.0.1:4222" 10 2 1
```

`manager` exercises the subscription broadcast path; `raw` is the plain
`NATS.Client` baseline. Parameters: url, seconds, publishers, listeners,
subject and an optional queue group. Reports throughput, allocated bytes per
received message and GC counts. Read-only stress runs against real servers are
allowed, but pass the url only as a command-line argument (never commit it) and
keep the runs short.

### 5. Liveness tests (planned — PLAN V8.5; requires NATS; Docker optional)

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

These cost real debugging time — read before writing automation. The E2E
script and suites they name do **not exist yet** (PLAN §V8); the lessons are
what the planned automation must honor, and several describe intended add-in
behavior (e.g. the `NatsRtd` registration), not code that runs today.

1. **Excel COM automation is inherently flaky; antivirus/EDR only adds to it.**
   Even with no antivirus, expect a retry count (`RPC_E_CALL_REJECTED`); a retry
   total is NORMAL. Symptoms are misleading errors (`InvalidCastException`,
   `RPC_E_CALL_REJECTED`, COM calls hanging); the E2E retries with backoff and a
   warm-up phase, so **if a run fails this way, run it again**. Never assume such
   an error means the add-in is broken. Every suite prints the active AV state.
2. **Elevated processes ignore per-user COM classes.** The add-in is *meant* to
   register `NatsRtd` under `HKCU\Software\Classes` in `AutoOpen` (not wired
   yet: `AutoOpen` only logs). If Excel runs elevated, `=RTD("NatsRtd", ...)`
   shows `#N/D`. Workaround for test machines: register under
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

*(Planned — there is no `.github/workflows` and no `CHANGELOG.md` yet.)* The
intended flow: push an annotated `v*` tag; the CI workflow builds and publishes
the packed XLLs, `NATSExcel.dll`, `NLog.config`, `NATSExcel.json` and the
installer (`NATSExcelSetup-<version>.exe`) as release assets. The version
embedded in the NATS connection `Name` (shown as `dev` for local builds) is the
tag. `INFORMATIONALVERSION` defaults to `dev` in `NATSExcel.csproj`.
