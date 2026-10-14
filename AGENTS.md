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

Read `docs/PUBLIC-SURFACE.md` for the frozen per-function catalog (the contract
users' spreadsheets depend on). The architecture and planning notes (`DESIGN.md`,
`PLAN.md`) are maintainer working copies kept on disk and deliberately not
published; this file is the working guidance.

> **Implementation status (read this first).** The add-in is implemented and
> released (`v0.1.0`, `v0.1.1` tags): the packed XLL loads in real Excel and
> answers UDFs, the full `NatsUDF*` catalog (**91** functions) plus the six
> `NatsRTD*` status helpers exist, and the RTD server runs. All **five** test
> layers are real and pass: unit (366), smoke (168, and 0 skipped on a full local
> run), Excel E2E (57 assertions in the default run; the async-marker journey
> SKIPs unless `-AsyncWrites`),
> load (~204k msg/s) and liveness. The solution file, the CI tag/release flow and
> the installer (11 wizard languages) are in place. `PLAN.md` §8/§9 track the few
> remaining additive items; the `## Layout` section below splits what exists from
> what is still planned.

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
   field. `reconnect.waitMs` is applied as BOTH `ReconnectWaitMin` and
   `ReconnectWaitMax` (the configured number is the wait, not its floor, which
   would let the client double it toward its own 5 s ceiling in
   NATS.Client.Core 3.3.0); `jitterMs` spreads it, and it is added after the clamp
   so it is never swallowed. `maxAttempts` is bounded by the client only when it is
   POSITIVE, so `-1` and `0` both mean unlimited (there is no "no retry at all").
   (An earlier draft nested the settings under `Default` with a top-level
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

**The add-in is implemented and released (`v0.1.0`): it compiles, packs, loads in
real Excel, answers the whole UDF catalog, and all five test layers pass.** A few
items from the audit/plan are still open (`PLAN.md` §8/§9).

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
| `NATSExcel.dna` | Excel-DNA packing manifest (31 dependency DLLs); its **file name** sets the packed XLL name. |
| `NATSExcel.json` / `NATSExcel.schema.json` | Sample config + embedded JSON Schema (validated by `ConfigSchema.cs` and `NatsUDFConfigValidate`). |
| `NLog.config` | Logging to `NATSExcel.log`, rolling at 100 MB, 5 archives; uses `${maskedException:format=toString}` (not `${exception:...}`). |
| `installer/` | Inno Setup script (`NATSExcel.iss`) + helper PowerShell + README: per-user, native-first installer. |
| `tools/NatsValidate/` | C# console validator that runs today (see `## Tests`). |
| `test/` | The unit, load and E2E layers + the cluster harness + the copy-based add-in folder helper (see `## Tests`). |

### Still open (from `PLAN.md`; none of these exist yet)

| Open item | Responsibility (per the plan) |
| --- | --- |
| V8.10 2-node placement error | Assert the named `no suitable peers for placement` error on a 2-node cluster (the 3-node failover half is covered by the E2E). Not reachable through the add-in surface today: `NatsUDFStreamCreate` exposes no `replicas` argument, so the add-in cannot ask for a 2-replica stream. |

V8.8 (the read-only live-feed run) **landed**: `test\Run-ExcelE2E.ps1` takes `-ReadOnly`
and `-AllowRemoteWrites`, skips every write below the first when
`$ReadOnly -or -not $isLocal`, and **refuses a non-local endpoint outright** unless
`-AllowRemoteWrites` is passed explicitly (the endpoint itself stays a CLI argument
and is never committed). It now lives under "Exists today".

The AV matrix helpers (`test\Run-ExcelE2E-Matrix.ps1`, `test\Compare-ExcelE2E.ps1`),
the read-only AV banner (`test\TestSupport\Get-AntivirusState.ps1`, which only
*reads* the state and never disables protection), the `expectCluster` surfacing in
`=NatsUDFConfigValidate()` (the *live* comparison is now real too: the connection
manager logs the mismatch - see `WarnOnClusterMismatch` - and `tools\NatsValidate`
reports it on demand) and the `jetstream.domain` routing have
landed, so they now live under "Exists today".

**The wizard is never asked to, and never does, disable or reconfigure antivirus
software.** No script in this repository touches `Set-MpPreference` or any
protection toggle: a helper that disables real-time protection is indistinguishable
from defense-evasion malware to a scanner and would rightly be blocked, so the
matrix only ever *reads* the AV state for attribution.

When an item lands, move its row up into "Exists today"; `DESIGN.md` §5.1
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
  between the profile and `C:\Windows`; it **is** implemented (the running
  process's main-module directory).
  Comments are accepted (string-aware `StripComments`). A wrong-typed value is
  dropped individually by `ConfigSchema.cs` (the schema walk), so the rest of the
  file is kept; only a structural (not-JSON) failure falls back to defaults.

## Build

```powershell
dotnet build NATSExcel.sln -c Release
```

The solution indexes the add-in, the four test projects and the validator. The
single-project form (`dotnet build NATSExcel.csproj -c Release`) still works if
you only want the add-in.

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

> **CI (PLAN P10).** `.github/workflows/build.yml` runs on PRs, main pushes and
> `v*` tags. It builds the solution, runs the unit tests, downloads the pinned
> `nats-server` zip **and verifies its SHA256**, then runs the smoke and liveness
> suites each in a step that starts its **own** server (the Windows runner kills
> a step's child processes when the step ends, so a server started in an earlier
> step would be gone). It builds the Inno Setup installer on every run, and on a
> `v*` tag it checks the top CHANGELOG heading against the tag, asserts the
> release assets, uses the CHANGELOG section as the release body and publishes
> the release.

The CI provisions NATS as a single self-contained `nats-server.exe` (no
Memurai/Chocolatey, unlike RedisExcel): download the
`nats-server-<ver>-windows-amd64.zip` from the GitHub releases, start it with
`-js --store_dir <temp>`, and probe port 4222 (the same readiness pattern as
the Redis `PING` wait). Do **not** use the community `nats-server` Chocolatey
package — it is stale (last published 2019). The `jetstream:true` probe up front
makes a "JetStream disabled" run attributable. Docker is never required, on CI
or on a user machine.

## Tests

All five layers are real: `test\NATSExcel.Tests` (unit), `test\SmokeTests`
(against a real `nats-server`), `test\LoadTests`, `test\LivenessTests`, and
`test\Run-ExcelE2E.ps1` (Excel), plus the cluster harness and helpers under
`test\TestSupport\`.

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

Covers today: `NatsJson` conversions, `AppConfig` load/sanitize/`StripComments`
(`StripCommentsTests`, `AppConfigPureHelperTests`, `ConfigLoaderToleranceTests`),
endpoint/cluster resolution and the per-cluster auth/timeout/reconnect merge
(`AppConfigResolutionTests`), the JSON-Schema per-value drop (`ConfigSchemaTests`),
`MaskNatsEndpoint` and `MaskedExceptionRenderer` (plus the `NLog.config` wiring that
makes the renderer effective), the conflation window, the RTD grammar
(`NatsRtdTopicGrammarTests`, including the optional-middle positional rule), the
subscription key, `WaitBounded` (`NatsConnectionWaitTests`), the write modes, the
frozen public surface (`PublicSurfaceParityTests` vs `PublicSurface.baseline.txt`),
the source scans that pin what cannot be referenced or driven
(`RtdInvariantTests` for `NatsRtd.cs`, `NoUnboundedBlockingTests`,
`SubjectEpochLeakTests` for the subject-latest epoch/loser path), a few pure-helper
pins (`PureHelperPinTests`: `TickGate`/`StripedLocks`/`UpdateCheck`/the NLog
renderer) and the packaging/parity tests that keep the `Compile` lists, the `.dna`
pack list and the Excel-DNA revision honest (`ProjectFileParityTests`,
`DnaPackagingParityTests`, `ExcelDnaPinTests`).
Growth areas (not yet covered): connection offline, the async bridge
(`NatsWriteObservable`), the RTD kinds the E2E does not create (only `SUB` is
exercised in Excel), `expectCluster` mismatch end to end.

`dotnet test` does not build or pack the add-in; CI builds it separately.

### 2. Smoke tests (real — PLAN V8.3; requires a NATS server, no Excel)

```powershell
powershell -ExecutionPolicy Bypass -File test\TestSupport\Start-Cluster.ps1
dotnet run --project test\SmokeTests -c Release -- "nats://127.0.0.1:4222"
```

Covers the ref-counted Core subscription behavior (two listeners on one subject
share exactly ONE physical subscription - proven on the wire through `/connz`,
which the harness's per-node `http_port` makes possible - plus last-listener
teardown, re-subscription, wildcards, queue groups, duplicate suppression,
argument validation), request/reply (timeout and `no responders`, both markers),
JetStream publish/consume (ack, sequence, durable consumer, replay, purge),
Key/Value (put/get/create/tombstone and a CAS by revision), Object Store
(put/get/digest plus the named size/binary guards), the admin gate on all eleven
gated functions, the dangling-reference `#CONFIG` and gate-ordering rules, the
`NatsUDFSubjectUnsubscribe` lifecycle and the JetStream-disabled path. It also
runs an **auth matrix** that starts a disposable local server per style (no-auth,
token, user/password, NKey) and proves both the success and the bounded failure,
including the `nkeySeed` text-or-`.nk`-file contract, and a **catalog sweep** that
calls every frozen function no other layer touches and requires a usable cell from
each (26 of the 97 had never been invoked anywhere). A full local run reports
**0 skipped**; a check that cannot run prints its reason and skips. An optional
`--read-only` skips every write section. `NATSEXCEL_NATS_SERVER` (a path, as the CI
sets) tells it where the binary is.

### 3. Excel end-to-end — real (PLAN V8.6)

```powershell
# Start a local cluster first (or point at one):
powershell -ExecutionPolicy Bypass -File test\TestSupport\Start-Cluster.ps1
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

Passes 57 assertions locally (the script asserts 53 times, a few inside retry
loops; the async-marker journey SKIPs unless the run is given `-AsyncWrites`).
Uses a hidden Excel instance: loads the packed XLL through
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
dotnet run --project test\LoadTests -c Release -- raw     "nats://127.0.0.1:4222" 10 2 1
```

`manager` exercises the subscription broadcast path; `raw` is the plain
`NATS.Client` baseline. Parameters: url, seconds, publishers, listeners,
subject/pattern and an optional queue group. Reports throughput and allocated
bytes per received message. ~204k msg/s measured on a local manager run.
Read-only stress runs against real servers are allowed, but pass the url only as
a command-line argument (never commit it) and keep the runs short.

### 5. Liveness tests (real — PLAN V8.5; requires NATS; Docker optional)

```powershell
powershell -ExecutionPolicy Bypass -File test\TestSupport\Start-Cluster.ps1
dotnet run --project test\LivenessTests -c Release -- "nats://127.0.0.1:4222" --quick --skip-restart
```

`test\LivenessTests` drives the real
managers under adversarial conditions — ThreadPool starvation, subscribe/dispose
churn, server restart/stop-start storms, async queue/observable bursts and a
resource-stability check — and enforces liveness bounds with a
dedicated process heartbeat and watchdog thread: a stall past **5 s** fails the
run, delivery must resume within **30 s**, and the async backlog must drain
within **20 s**. `--quick` shortens the attack windows; `--skip-restart` skips
only the destructive restart sub-phase (a non-local host skips it automatically).
Only run it against a disposable/local server: the restart storm is destructive.

## Excel automation lessons (hard-won, inherited from RedisExcel)

These cost real debugging time — read before writing automation. The E2E script,
the cluster harness and the smoke/liveness suites all exist
(`test\Run-ExcelE2E.ps1`, `test\TestSupport\`, `test\SmokeTests`,
`test\LivenessTests`). Several lessons describe add-in behavior that now runs
(the `NatsRtd` registration) rather than intent.

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
   And treat a COM call that answers with `$null` or an ARRAY instead of throwing
   as a failure to retry: `Workbooks.Open` did exactly that under load, and
   `if (-not $wb)` is the WRONG test for it - a one-element array is truthy, so the
   code took the success path and the failure surfaced later as "cannot call a
   method on a null-valued expression" (member enumeration over an array holding a
   null). Compare with `$null -eq $wb` and require `-is [System.__ComObject]`, with
   a small retry loop around the open.
6. **RTD cells display `#N/D` when the RTD server could not be created** —
   check COM registration before suspecting the add-in logic.
7. **NATS-specific:** a server started without `-js` accepts connections and
   Core pub/sub works, but every JetStream call fails - assert the JetStream
   probe up front and print whether `-js` is on, so a "JS disabled" run is
   attributable rather than looking like an add-in bug. The JetStream
   `store_dir` default is a temp path; the CI container and the installer's
   server option must mount a durable directory or a restart drops the streams.
8. **Never stop a process by PORT alone - prove it is ours.** A run against a
   localhost URL that some OTHER nats-server already answers used to be treated
   as "our disposable cluster": the failover journey then killed that server (a
   denied kill aborted the run with a raw `Access denied`, and a permitted one
   would have destroyed a developer's instance). `Test-OwnedTestServer` in
   `test\Run-ExcelE2E.ps1` is the proof - the listener must be a `nats-server`
   whose command line names a `nats.conf` under OUR temp work root (a process
   from another user or an elevated shell exposes no command line at all, which
   is itself the signal that it is not ours). No proof, no kill: the journey
   SKIPs with the reason. The same rule applies to the pre-run stale cleanup
   (`Stop-Cluster.ps1` `Test-Ours`, `Start-SecondServer`), and a run that uses a
   server it does not own says so at the top instead of being discovered from a
   denied kill. Pin: `HarnessInvariantTests`.
9. **A crashed run orphans its hidden Excel, which locks the copied XLL.** The
   next run then dies in `New-AddInFolder.ps1` on "the file is being used by
   another process" and, worse, the E2E reported it as an obscure
   "Path is an empty string". The helper now stops a stale host by PROOF (the
   process is invisible AND has our destination XLL loaded) and only REPORTS a
   visible Excel (the user's session, which never has that temp copy loaded).
   Two corollaries for new steps: an `Invoke-Com { ... }` that can answer with
   `$null` instead of throwing (e.g. `Workbooks.Open`) needs its own retry, and
   any step that creates a hidden Excel must quit it on the failure path too -
   `Move-Item`/`Copy-Item` are not the only things a crash leaves behind.
10. **A check that always SKIPs is a check that never ran - treat a permanent
    SKIP as a defect.** The smoke suite's wire-level refcount proof had reported
    "no monitoring port" for its whole life, and it was two independent bugs
    stacked: the harness started no monitoring port AND the probe parsed
    `/connz`'s `subscriptions_list` as a JSON object while the server sends an
    ARRAY (`?subs=1`), so `sawList` stayed false and the check could not pass even
    against a server that exposed `/connz`. Two rules follow: when a suite prints
    a skip reason, make the reason falsifiable (a skip whose text is always the
    same is indistinguishable from a broken probe), and when a layer reports
    `0 skipped` locally, keep it there - new work that needs a skip should say
    why in the same commit.
11. **Never run two Excel-automation suites on the same machine at once.** Each
    suite creates hidden `EXCEL.EXE` instances and quits them when it is done
    (and the RedisExcel sibling suite does the same), so a concurrent run's quit
    kills the instance the other one is driving: the symptom is a COM
    `0x800706BA` "The RPC server is unavailable" in the middle of a journey, plus
    mystery `EXCEL.EXE` processes that carry a window title but have NOT loaded
    our XLL. Those strangers are not ours to kill - the ownership proof (invisible
    AND holding our destination XLL) is what keeps a cleanup from closing the
    other suite's workbook. Before blaming the add-in for a COM death, check
    whether another suite (or the user) is running Excel; if it is, re-run later
    instead of debugging a phantom.

## Release

The CI workflow (`.github/workflows/build.yml`) is live: pushing an annotated
`v*` tag builds the solution, runs the unit/smoke/liveness layers, builds the
installer and publishes the packed XLLs, `NATSExcel.dll`, `NLog.config`,
`NATSExcel.json` and the installer (`NATSExcelSetup-<version>.exe`) as release
assets, using the CHANGELOG section for the tag as the release body. Releases
`v0.1.0` and `v0.1.1` have been cut. The version embedded in the NATS connection
`Name` (shown as `dev` for local builds) is the tag read from
`AssemblyInformationalVersion`; `INFORMATIONALVERSION` defaults to `dev` in
`NATSExcel.csproj`.

The landing page (`docs/`, published to GitHub Pages by
`.github/workflows/pages.yml`) is **version-agnostic by design**: the download
section links `releases/latest`, the version is rendered by a shields.io badge and
the asset table uses a `NATSExcelSetup-<versão>.exe` placeholder, so **cutting a
release needs NO page change**. `pages.yml` only runs when `docs/**` or `mkdocs.yml`
change, so a code/test push does not redeploy it and that is expected. After a
release, the check is just: `/releases/latest` resolves to the new tag and the badge
renders.
