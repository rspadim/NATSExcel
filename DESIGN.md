# NATSExcel — Design

Excel add-in (XLL) that integrates **NATS** with Excel through RTD (Real-Time
Data) and UDF (User Defined Functions), in the same spirit as the sibling
project **RedisExcel**: the user types a formula, NATS data appears in the cell
and stays live.

This is the day-one design contract. The public Excel surface defined here is
frozen before the first release: renames and argument changes break users'
spreadsheets forever (see the golden rules in `AGENTS.md`).

- Status: implemented design contract; the packed XLL loads in real Excel and
  the catalog is exposed (see "As built so far" below for the exact state).
- Target framework: **.NET Framework 4.8** (Excel-DNA), same host as RedisExcel.
- Scope of this document: **Core NATS + JetStream from the start** (streams,
  consumers, Key/Value, Object Store).
- Not a code port of RedisExcel. It is a sibling project: the *architecture
  patterns* are reused, the *code* is new (different client, different protocol,
  different semantics).

---

## 1. Goals

- One add-in that turns NATS subjects into Excel values: live cells (RTD) and
  on-demand reads/writes (UDF).
- Cover **Core NATS** (publish/subscribe, request/reply, queue groups, headers)
  **and JetStream** (streams, consumers, message replay, Key/Value Store,
  Object Store) from v1.
- Never block the Excel thread: NATS .NET is async-only, so the async bridge is
  architecture, not an optional feature (RedisExcel's `AsyncWrites` becomes the
  default path here).
- Configuration is schema-validated and tolerant: one bad value is dropped, the
  rest of the file is kept. Invalid configuration is visible (`#CONFIG`), never
  a silent fallback that points at `localhost`.
- No secrets in cells or logs: passwords, tokens, NKey seeds and creds-file
  paths are always masked.
- One-click Windows installer, CI-built, localized (en/es/pt-BR), like
  RedisExcel.

### As built so far

The design above is the contract; this is what physically exists today. The
add-in is now largely implemented: the packed XLL loads in real Excel and
answers UDFs (the P0 gate passed), the full catalog is exposed, and the Excel
E2E / unit / load layers run.

- **Add-in skeleton**: `NATSExcel.csproj` (net48, `LangVersion 9.0`) with the
  `IsExternalInit` shim, the `NATSExcel.dna` packing manifest, `AddIn.cs`
  (`IExcelAddIn` load/unload plus the `NatsRtd` COM registration and the
  reload resets) and the `Config.cs` / `ConfigSchema.cs` loader that reads
  `NATSExcel.json` once per process with per-value schema drop.
- **Runtime**: `NatsConnectionManager.cs`, `NatsSubscriptionManager.cs`,
  `NatsRuntime.cs`, `NatsUdfAsync.cs`, `Conflation.cs`, `TickGate.cs`,
  `UpdateCheck.cs`, `NatsJson.cs`, `NatsResultFormatter.cs`,
  `MaskedExceptionRenderer.cs`.
- **Excel surface**: `NatsRtd.cs` (the `NatsRtd` RTD server + the six
  `NatsRTD*` status helpers), `NatsRtdTopic.cs` (the Excel-free frozen
  grammar), and the `NatsUdf*` catalog — **91** `NatsUDF*` `[ExcelFunction]`s
  across `NatsUdf.cs` / `NatsUdfCore.cs` / `NatsUdfJetStream.cs` /
  `NatsUdfKv.cs` / `NatsUdfObjectStore.cs` / `NatsUdfStatus.cs`, plus
  `NatsUdfCommon.cs`.
- **Windows installer** (`installer/`): an Inno Setup script that builds and
  runs, per-user, localized **en/es/pt-BR**, with **7** auth styles (none, token,
  user/pass, NKey, JWT, creds file, mTLS) and a native-first "install a NATS
  server" step. A workflow builds it on every CI run; it has not run on GitHub
  yet.
- **Validator tool** (`tools/NatsValidate/`): a standalone C# console app that
  loads and checks `NATSExcel.json` and, with `--test`, connects with the same
  NATS.Net client the add-in uses.
- **Tests**: `test\NATSExcel.Tests` (unit, no NATS/Excel), the cluster harness
  (`test\TestSupport\Start-Cluster.ps1`), the Excel E2E
  (`test\Run-ExcelE2E.ps1`) and the load layer (`test\LoadTests`). The smoke and
  liveness projects are still planned.

The remaining work (the medium/low audit findings, the smoke/liveness layers,
the solution file and the release cut) is tracked in `PLAN.md`.

## 2. Non-goals (v1)

- NATS accounts/operator-mode administration (users, JWTs, accounts) from
  Excel. The add-in *connects* with credentials; it does not manage the trust
  chain.
- Server-side stream/consumer design advice: the add-in creates assets only
  through explicit admin functions (section 7.3), never implicitly.
- A separate watchdog service. Everything runs inside the Excel process.
- Leaf-node / super-cluster topology management (the client just connects to a
  list of URLs; the server-side topology is the operator's concern).
- Sharded/partitioned helper packages (Orbit-style sugar); core client only.

## 3. Why the design differs from RedisExcel

| Concern | RedisExcel | NATSExcel |
| --- | --- | --- |
| Connection | Pool of `ConnectionMultiplexer` per host/pool | **One `NatsConnection`** multiplexes everything; reconnect is built in |
| Client API | Sync available (sync-over-async inside) | **Async-only** (`NATS.Net` v3) → mandatory async bridge |
| Patterns | `SUB` / `PSUB` | Subjects with native wildcards `*` and `>`, plus **queue groups** |
| Request/reply | Not available | **First-class** (Inbox + reply subject) → new UDF |
| Delivery | Pub/Sub is fire-and-forget (lost with no subscriber) | Core is at-most-once; **JetStream** adds persistence, ack, replay |
| Key/value | `GET`/`SET`/`HGET`/`HSET` | **JetStream KV** (buckets, revisions, CAS, watch, history) |
| Files | none | **Object Store** (chunked objects, digest) |
| Auth | `password=`/`user=` in the connection string | user/pass, **token**, **NKey seed**, **creds file**, TLS |
| Discovery | manual hosts + client-side failover (REDIS DESIGN v2.0) | multiple `nats://` URLs; cluster gossip handled by the client |
| Dedup | `SkipRepeatedMessages` (ordinal payload compare) | same idea for Core; JetStream carries **sequence/revision** natively |

Consequences that shape the code:

1. **Async bridge is core.** Every NATS call is `Task`-based. UDF reads and
   writes go through a bounded sync helper (section 5.3) or
   `ExcelAsyncUtil.Observe`; RTD is the naturally async path.
2. **No per-host connection pools.** The manager caches one connection per
   resolved *endpoint set* (a lone URL list is one connection). The RedisExcel
   eviction/idle-down machinery is mostly unnecessary; a single reconnect loop
   replaces it.
3. **JetStream needs its own error vocabulary.** "JS not enabled on the
   server", "stream not found", "wrong last sequence", "no responders" must
   surface as distinct, actionable cell errors.

## 4. Concepts

| Term | Meaning |
| --- | --- |
| Subject | Dot-separated address; `*` = one token, `>` = one or more (subscribe only; publish needs a concrete subject). |
| Endpoint | The connection target: one or more `nats://` URLs plus auth/TLS options; may be named by an alias. |
| Subscription (logical) | `(endpoint, subject, queueGroup?)` as seen by the user. |
| Subscription (physical) | What the client holds; ref-counted and shared by every listener on the same logical subscription. |
| Stream | JetStream durable storage bound to subject(s). |
| Consumer | A JetStream cursor over a stream (durable or ephemeral, push or pull). |
| Bucket (KV) | A JetStream stream (`KV_<name>`) with key/value semantics and revisions. |
| Bucket (Object) | A JetStream stream (`OBJ_<name>`) storing chunked files. |

Endpoint resolution for any host argument: a `Clusters` name, then a `Servers`
alias, then the raw `nats://` URL / URL list. No implicit fallback to a hidden
default for a name that does not resolve (see §8.1).

## 5. Architecture

### 5.1 File layout

| File | Responsibility |
| --- | --- |
| `NatsRuntime.cs` | Process-wide singleton: one `NatsConnection` per endpoint, plus the subscription manager. Shutdown fence and same-process reload, mirroring `RedisRuntime`. |
| `NatsConnectionManager.cs` | Creates/caches connections by resolved endpoint; connection state, reconnect observation, jetstream context cache, bounded wait helper. |
| `NatsSubscriptionManager.cs` | Ref-counted Core subscriptions: one client subscription per `(endpoint, subject, queueGroup)`, N listeners; `IDisposable` tokens; wildcard-aware; origin tags (`RTD`/`UDF`). |
| `AddIn.cs` | `IExcelAddIn` entry point (`AutoOpen`/`AutoClose`): registers the `MaskedExceptionRenderer` and the `NatsRtd` COM server, runs the reload resets, and shuts the runtime down. |
| `NatsRtd.cs` | The `NatsRtd` `ExcelRtdServer`: topic registry, poll timer, Excel flush timer, backoff, wildcard topics, JetStream/KV topic kinds; plus the six `NatsRTD*` status helpers. |
| `NatsRtdTopic.cs` | RTD topic/command parser (`NatsRtdTopicSpec.Parse`); deliberately free of Excel/COM so it is unit-testable (`RtdTopicGrammarTests`). |
| `NatsUdfCore.cs` | Core NATS UDFs: publish, request/reply, subject-latest, unsubscribe. |
| `NatsUdfJetStream.cs` | JetStream UDFs: publish-with-ack, get/fetch, stream/consumer info, admin. |
| `NatsUdfKv.cs` | Key/Value UDFs: get/put/create/update/delete/keys/history/watch. |
| `NatsUdfObjectStore.cs` | Object Store UDFs: put/get/info/list/watch/buckets/delete. |
| `NatsUdf.cs` | The seed diagnostics (`NatsUDFVersion`, `NatsUDFConfigPath`, `NatsUDFConnectionCount`, `NatsUDFLogPath`). |
| `NatsUdfStatus.cs` | Status, diagnostics and cluster/server topology (`NatsUDFServerInfo`, `NatsUDFStats`, `NatsUDFRtt`, `NatsUDFCluster*`, config functions, ...). |
| `NatsUdfCommon.cs` | Shared UDF primitives: argument validation (incl. wildcard rejection on publish), endpoint resolution, `Fail`/`VerboseError`/`ErrorHint` and the marker constants. |
| `NatsUdfAsync.cs` | Shared async bridge: `ExcelAsyncUtil.Observe` wrapper + per-endpoint FIFO write queue + bounded sync wait. |
| `NatsJson.cs` | Matrix ⇄ JSON conversions (`MatrixToJson` / `JsonToMatrix`, internal helpers). |
| `NatsResultFormatter.cs` | Value shaping for cells (JSON objects, KV entries, object metadata). |
| `Conflation.cs` | Real-time conflation window (`ConflationMs`): latest value wins per window. |
| `TickGate.cs` | Non-blocking reentrancy gate for timers + `StripedLocks`. |
| `Config.cs` / `ConfigSchema.cs` | Config load (once per process) + embedded JSON Schema, tolerant validation, endpoint/auth masking. |
| `MaskedExceptionRenderer.cs` | NLog `${maskedException}` layout renderer: masks credentials in every logged exception. |
| `UpdateCheck.cs` | Background GitHub release check. |
| `NatsExcel.dna` | Excel-DNA packing manifest. |
| `installer/` | Inno Setup script + helper PowerShell + README. |

### 5.2 Packages

Explicit packages (not the `NATS.Net` meta package) so the packed XLL stays
lean; the set is pinned in `NATSExcel.csproj` (all at `3.3.0`):

- `NATS.Client.Core` — connection, pub/sub, request/reply.
- `NATS.Client.JetStream` — streams, consumers, message get/fetch.
- `NATS.Client.KeyValueStore` — KV buckets.
- `NATS.Client.ObjectStore` — object buckets.
- (transitive, packed into the XLL: `NATS.Client.Abstractions`, `NATS.NKeys`,
  `Microsoft.Extensions.Logging.Abstractions`, `Microsoft.Bcl.AsyncInterfaces`,
  and the rest of the 31-DLL set enumerated in `NATSExcel.dna`.)

Serialization: the add-in uses **Newtonsoft.Json** for its own matrix ⇄ JSON and
KV/object metadata; the NATS.Net packages pull `System.Text.Json` transitively
(both are packed — see §13 decision 6). Keep the closure minimal and pinned, like
RedisExcel's deliberate list.

### 5.3 Async bridge (the central design decision)

`NATS.Net` exposes only `Task`/`IAsyncEnumerable`. Excel UDFs are synchronous:
they must return a value before returning. Three paths, chosen per call kind:

1. **RTD topics** — push model, naturally async. Polled topics run on the poll
   timer (never on the Excel thread); subscription topics push from the client's
   callback. This is the primary path for live data.
2. **UDF reads/writes that must return in-cell** — run the async operation on a
   worker and wait with a **bounded** helper (`WaitBounded(task, timeout + 500ms)`,
   the RedisExcel pattern). Never `GetAwaiter().GetResult()` directly on the
   Excel thread without a worker and a bound.
3. **UDF calls that can be asynchronous** (`AsyncWrites`-style, and any read
   the user opts into) — `ExcelAsyncUtil.Observe` with a custom observable: the
   cell shows the pending marker (`#N/A`) and later the real value/error. Reuse
   `NatsUdfAsync` (RedisExcel's `RedisWriteObservable`).

Rules derived from this:

- Every `await` in library code uses `ConfigureAwait(false)`.
- A worker is never held per pending item: the per-endpoint FIFO queue delivers
  results via continuations, not one thread per item.
- The Excel thread never blocks longer than the configured response timeout.

## 6. RTD server

ProgID: **`NatsRtd`**. Usage: `=RTD("NatsRtd", , <command>, ...)`.

| Command | Arguments | Behavior |
| --- | --- | --- |
| `SUB` | `<subject> [queueGroup] [host]` | Subscribe (wildcards allowed). Push on every message. |
| `KV` | `<bucket> <key> [host]` | Watch a KV key; push on every revision change. |
| `KVGET` | `<bucket> <key> [host]` | Polled KV read (one read per poll tick). |
| `KVKEYS` | `<bucket> [host]` | Polled key list. |
| `REQ` | `<subject> <payload> <everyMs> [host]` | Issue a request/reply every tick; push the reply. |
| `JS` | `<stream> <consumer> [filter] [host]` | Consume a JetStream consumer (push); push each message. |
| `MSG` | `<stream> <filter> [host]` | Polled "last message matching filter". |
| `OBJ` | `<bucket> <object> [host]` | Watch object metadata; push on change. |
| `RTDCLUSTER` | `[what] [host]` | Live cluster-health cell (`what` = `cluster`, `nodes`, `leader`, `lag`, or blank/`all`); push on change. Tier 2 (§7.8): needs the system account. |

Notes:

- **Nine commands**: the eight above plus `RTDCLUSTER`. The parser rejects
  anything else by name (see §7.8 for the cluster tier this command reads) and
  reports the frozen command list `[SUB, KV, KVGET, KVKEYS, REQ, JS, MSG, OBJ,
  RTDCLUSTER]`; the `RTDCLUSTER what` vocabulary is a separate kind-specific
  check (`RTDCLUSTER expects what to be one of [cluster, nodes, leader, lag]`).
- **No `PSUB`.** NATS wildcards are part of the subject (`orders.*`,
  `orders.>`); a pattern subscription is just `SUB orders.>`.
- A `SUB`/`KV`/`JS` topic that fails to subscribe shows `#ERROR` and is retried
  with 1s..30s backoff, capped per tick (RedisExcel behavior).
- Polled kinds (`KVGET`, `KVKEYS`, `MSG`, `REQ`, `RTDCLUSTER`) retry on every tick.
- If JetStream is not enabled on the server, `JS`/`KV`/`OBJ` topics show a
  clear `#ERROR: JetStream is not enabled on this server`.

## 7. UDF catalog (the "final version")

Conventions mirror RedisExcel exactly:

- Prefix `NatsUDF`. Reads and status are **volatile** (`IsVolatile = true`).
- Every write has an additive **`...NonVolatile`** twin (identical arguments,
  thin delegation) so a recalculation does not re-send it.
- Argument validation failures return a terse, stable `Error: ...`.
- Runtime errors name the identifiers acted on and mask credentials.

### 7.1 Core NATS

| Function | Notes |
| --- | --- |
| `NatsUDFPublish(subject, message, [host])` | + `NonVolatile`. |
| `NatsUDFPublishJSON(subject, range, [host])` | matrix → JSON. + `NonVolatile`. |
| `NatsUDFPublishIfChanged(subject, message, [host])` | dedup on last payload (bounded LRU). + `NonVolatile`. |
| `NatsUDFPublishWithHeaders(subject, message, headers, [host])` | `headers` = key/value range. + `NonVolatile`. |
| `NatsUDFRequest(subject, message, [timeoutMs], [host])` | returns reply, or a `TIMEOUT`/`No Responders` marker. + `NonVolatile`. |
| `NatsUDFRequestJSON(subject, range, [timeoutMs], [host])` | matrix → JSON request. + `NonVolatile`. |
| `NatsUDFSubjectLatest(subject, [queueGroup], [host])` | subscribe + cache last value (the `ChannelLatest` analogue). |
| `NatsUDFSubjectLatestJSON(subject, [queueGroup], [host])` | same, JSON payload. |
| `NatsUDFSubjectUnsubscribe(subject, [host])` | explicit lifecycle. + `NonVolatile`. |
| `NatsUDFSubscriptionCount()` | active subscriptions. |

### 7.2 JetStream — streams, messages, consumers

| Function | Notes |
| --- | --- |
| `NatsUDFJSPublish(subject, message, [host])` | publish with ack; returns sequence. + `NonVolatile`. |
| `NatsUDFJSPublishJSON(subject, range, [host])` | + `NonVolatile`. |
| `NatsUDFJSGetMessage(stream, seq, [host])` | replay by sequence. |
| `NatsUDFJSGetLastMessage(stream, filter, [host])` | last message matching filter. |
| `NatsUDFJSFetch(stream, consumer, count, [timeoutMs], [host])` | pull batch → matrix. |
| `NatsUDFJSStreamInfo(stream, [host])` | messages/bytes/first/last seq. |
| `NatsUDFJSStreams([host])` | list streams. |
| `NatsUDFJSConsumerInfo(stream, consumer, [host])` | pending/redelivered/lag. |
| `NatsUDFJSConsumers(stream, [host])` | list consumers. |
| `NatsUDFJSDeleteMessage(stream, seq, [host])` | + `NonVolatile`; destructive. |
| `NatsUDFJSPurge(stream, [filter], [host])` | + `NonVolatile`; destructive. |

### 7.3 JetStream — admin (opt-in)

| Function | Notes |
| --- | --- |
| `NatsUDFStreamCreate(name, subjects, [storage], [retention], [host])` | + `NonVolatile`. |
| `NatsUDFStreamUpdate(...)` / `NatsUDFStreamDelete(name, [host])` | delete is destructive. |
| `NatsUDFConsumerCreate(stream, name, [filter], [deliverPolicy], [host])` | durable consumer. |
| `NatsUDFConsumerDelete(stream, name, [host])` / `NatsUDFConsumerPause(...)` | destructive. |

Admin and destructive functions are **gated by config** `AllowAdminWrites`
(default `false`): with it off they return
`Error: admin writes are disabled (set "AllowAdminWrites": true)`. Explicit and
reversible, so an accidental F9 never purges a production stream.

### 7.4 Key/Value Store

| Function | Notes |
| --- | --- |
| `NatsUDFKVGet(bucket, key, [host])` | missing key → marker. |
| `NatsUDFKVGetMultiple(bucket, keys, [columns], [host])` | matrix. |
| `NatsUDFKVKeys(bucket, [pattern], [host])` | key list. |
| `NatsUDFKVStatus(bucket, key, [host])` | revision, created/updated, operation. |
| `NatsUDFKVHistory(bucket, key, [host])` | revisions with values. |
| `NatsUDFKVPut(bucket, key, value, [host])` | upsert. + `NonVolatile`. |
| `NatsUDFKVPutJSON(bucket, key, range, [host])` | + `NonVolatile`. |
| `NatsUDFKVCreate(bucket, key, value, [host])` | only if absent (rev 0). + `NonVolatile`. |
| `NatsUDFKVUpdate(bucket, key, value, revision, [host])` | CAS by revision. + `NonVolatile`. |
| `NatsUDFKVDelete(bucket, key, [host])` / `NatsUDFKVPurge(bucket, key, [host])` | + `NonVolatile`. |
| `NatsUDFKVBuckets([host])` | list buckets. |

### 7.5 Object Store

| Function | Notes |
| --- | --- |
| `NatsUDFObjectPut(bucket, name, value, [description], [host])` | + `NonVolatile`. |
| `NatsUDFObjectGet(bucket, name, [host])` | text/bytes; oversized → `Error:`. |
| `NatsUDFObjectInfo(bucket, name, [host])` | digest, size, chunks, metadata. |
| `NatsUDFObjectList(bucket, [host])` | list. |
| `NatsUDFObjectDelete(bucket, name, [host])` | + `NonVolatile`. |
| `NatsUDFObjectWatch(bucket, [host])` | watch metadata (live). |
| `NatsUDFObjectBuckets([host])` | list buckets. |

### 7.6 Status, diagnostics, config, JSON

`NatsUDFVersion`, `NatsUDFConfigPath`, `NatsUDFLogPath`,
`NatsUDFConnectionCount`, `NatsUDFServerInfo` (server id/version/connect_urls),
`NatsUDFStats` (in/out msgs/bytes), `NatsUDFRtt`, `NatsRTDConnectionCount`,
`NatsRTDTopicCount`, `NatsRTDSubscriptionCount`, `NatsRTDSubjectCount`,
`NatsRTDDefaultHost`, `NatsUDFUpdateAvailable`,
`NatsUDFConfigSchema`, `NatsUDFConfigValidate`, `NatsUDFConfigReload`.

**One config-validation function, `NatsUDFConfigValidate`** — it is the only
config check. (An earlier draft carried a `NatsCheckConfigFile` name inherited
from the sibling project; that name is **removed**. The prefix is `NatsUDF` for
every `[ExcelFunction]` without exception.) `NatsUDFConfigReload` is the only
standalone non-volatile consumer function here (§7.9).

**Matrix ⇄ JSON, `NatsUDFMatrixToJSON` / `NatsUDFJSONToMatrix`** — the standalone
pure conversions, both non-volatile (a conversion of its inputs has no reason to
re-run on every recalculation) and server-free. They wrap the same `NatsJson`
helpers the `...JSON` read/write functions use, so the shape is identical: a 1x1
range is a scalar, one row is a flat array, anything else an array of row arrays,
and a JSON `null` becomes the optional `nullValue` (blank by default). A failed
conversion returns `Error: ...` (an invalid JSON gives the parser's message).

### 7.7 `...NonVolatile` write twins (the source of truth)

Each row below is the additive twin of the write base named on the left; the
arguments are **identical** and the body is a thin delegation to the base with
`IsVolatile = false`. This list is the public-surface baseline test's source of
truth for the write set (25 twins); the base lives in the section shown.

| Base write (§) | Twin |
| --- | --- |
| `NatsUDFPublish` (§7.1) | `NatsUDFPublishNonVolatile` |
| `NatsUDFPublishJSON` (§7.1) | `NatsUDFPublishJSONNonVolatile` |
| `NatsUDFPublishIfChanged` (§7.1) | `NatsUDFPublishIfChangedNonVolatile` |
| `NatsUDFPublishWithHeaders` (§7.1) | `NatsUDFPublishWithHeadersNonVolatile` |
| `NatsUDFRequest` (§7.1) | `NatsUDFRequestNonVolatile` |
| `NatsUDFRequestJSON` (§7.1) | `NatsUDFRequestJSONNonVolatile` |
| `NatsUDFSubjectUnsubscribe` (§7.1) | `NatsUDFSubjectUnsubscribeNonVolatile` |
| `NatsUDFJSPublish` (§7.2) | `NatsUDFJSPublishNonVolatile` |
| `NatsUDFJSPublishJSON` (§7.2) | `NatsUDFJSPublishJSONNonVolatile` |
| `NatsUDFJSDeleteMessage` (§7.2) | `NatsUDFJSDeleteMessageNonVolatile` |
| `NatsUDFJSPurge` (§7.2) | `NatsUDFJSPurgeNonVolatile` |
| `NatsUDFStreamCreate` (§7.3) | `NatsUDFStreamCreateNonVolatile` |
| `NatsUDFStreamUpdate` (§7.3) | `NatsUDFStreamUpdateNonVolatile` |
| `NatsUDFStreamDelete` (§7.3) | `NatsUDFStreamDeleteNonVolatile` |
| `NatsUDFConsumerCreate` (§7.3) | `NatsUDFConsumerCreateNonVolatile` |
| `NatsUDFConsumerDelete` (§7.3) | `NatsUDFConsumerDeleteNonVolatile` |
| `NatsUDFConsumerPause` (§7.3) | `NatsUDFConsumerPauseNonVolatile` |
| `NatsUDFKVPut` (§7.4) | `NatsUDFKVPutNonVolatile` |
| `NatsUDFKVPutJSON` (§7.4) | `NatsUDFKVPutJSONNonVolatile` |
| `NatsUDFKVCreate` (§7.4) | `NatsUDFKVCreateNonVolatile` |
| `NatsUDFKVUpdate` (§7.4) | `NatsUDFKVUpdateNonVolatile` |
| `NatsUDFKVDelete` (§7.4) | `NatsUDFKVDeleteNonVolatile` |
| `NatsUDFKVPurge` (§7.4) | `NatsUDFKVPurgeNonVolatile` |
| `NatsUDFObjectPut` (§7.5) | `NatsUDFObjectPutNonVolatile` |
| `NatsUDFObjectDelete` (§7.5) | `NatsUDFObjectDeleteNonVolatile` |

### 7.8 Cluster and server topology UDFs

All read-only. They answer "is the cluster healthy, who is the leader, which
node is lagging" from a cell, which is exactly the kind of thing a spreadsheet
is good at. Two tiers, because the server decides what it will tell a client:

**Tier 1 — no special privilege (works on the data connection).**

| Function | Returns | Notes |
| --- | --- | --- |
| `NatsUDFServerInfo([host])` | JSON | Already in §7.6; adds `server_name`, `cluster`, `connect_urls` when present. |
| `NatsUDFClusterName([host])` | text | The cluster name the **server announces** (`INatsServerInfo.Cluster`); blank for a standalone server. |
| `NatsUDFClusterNodes([host])` | matrix N×3 (`server_name`, `cluster`, `connect_url`) | Built from the client's own discovery (`ClientConnectUrls`), so it needs no account privileges. What the client can *see*, not the full membership. |
| `NatsUDFConnectionState([host])` | text | `CONNECTED` / `RECONNECTING` / `CLOSED` / `LAME_DUCK`, plus the currently selected URL (masked). |
| `NatsUDFRtt([host])` | number (ms) | Already in §7.6; a per-node latency probe when the client has several. |

**Tier 2 — requires the system account (the full membership view).**

| Function | Returns | Notes |
| --- | --- | --- |
| `NatsUDFClusterInfo([host])` | JSON | `$SYS.REQ.SERVER.PING` summary: node list, versions, roles. |
| `NatsUDFClusterHealth([host])` | JSON | Per-node: `server_name`, `mem`, `cpu`, `connections`, `slow_consumers`, `jetstream` state. |
| `NatsUDFJSRaftInfo(stream, [host])` | matrix N×4 (`peer`, `current`, `active`, `lag`) | The stream's Raft group: leader + replicas + lag, i.e. the table from §8.1's verification. |
| `NatsUDFJSStreamPlacement(stream, [host])` | text | Where the Raft group lives; the source of the `no suitable peers for placement` case. |
| `RTDCLUSTER([what], [host])` | RTD command | A live cluster-health cell (`what` = `cluster`, `nodes`, `leader`, `lag`); pushes on change. |

Notes:

- Tier 2 goes through the NATS **system account** (`$SYS`); if the endpoint is
  not authorized for it the call returns a named error
  (`Error: the system account is not available on this endpoint (cluster
  queries need a system account)`), never an empty matrix that looks healthy.
- `NatsUDFJSRaftInfo` works with the data account (the stream's `cluster` block
  is part of `stream info`), so it is the most broadly usable cluster function;
  the read is the same one §8.1 already verified.
- **Read-only by construction**: no cluster function creates, joins, removes or
  reconfigures anything. Cluster formation stays a server-side concern (§8.1).

### 7.9 Catalog size

The catalog is counted mechanically from the tables, not from memory. Two
families:

- **UDF functions** (`NatsUDF*`, all in `NatsUdf*.cs`): **66 base** + the
  **25** `...NonVolatile` write twins (§7.7) = **91** `[ExcelFunction]`.
- **RTD status helpers** (`NatsRTD*`, in `NatsRtd.cs`): **6**
  (`ConnectionCount`, `TopicCount`, `SubscriptionCount`, `SubjectCount`,
  `DefaultHost`, `Cluster`). They carry `[ExcelFunction]` too, but live with the
  RTD server rather than in a `NatsUdf*.cs` source.

So the add-in exposes **91 UDF functions + 6 RTD status helpers** and **9 RTD
commands** (§6).

Base-function split (the 66 `NatsUDF*` base):

| Group | Count |
| --- | --- |
| Write bases (each with a `...NonVolatile` twin) | **25** |
| Reads/status (volatile, twin-less) | **38** |
| Standalone non-volatile, twin-less (`ConfigReload`, `MatrixToJSON`, `JSONToMatrix`) | **3** |
| **Base total** | **66** |

Volatile/non-volatile: **63 volatile** (25 write bases + 38 reads/status) +
**28 non-volatile** (25 twins + 3 standalone) = 91. Functions
gated by `AllowAdminWrites`: **11** (§7.3 plus `JSDeleteMessage`/`JSPurge`,
`KVDelete`/`KVPurge`, `ObjectDelete`).

The count is verified mechanically against the code: **91** `[ExcelFunction]`
`NatsUDF*` (66 base + 25 twins) + **6** `NatsRTD*` = **97** functions. Earlier
drafts of this document said "95 / 88 / 66 / roughly 70" and a 61/28 volatile
split; those numbers were wrong and are superseded by the figures above.

## 8. Configuration

File name: **`NATSExcel.json`**. Search order (first existing wins): add-in
folder → `%USERPROFILE%` → the Excel folder (the running process's main-module
directory) → `C:\Windows` (`AppConfig.CandidatePaths`).
Comments (`//`, `/* */`) accepted; validated against an embedded JSON Schema;
a wrong-typed value is dropped individually, the rest of the file is kept.

Layout principle: **the user picks a cluster, not a host.** `Clusters` is the
unit a formula names; `Servers` is the primitive it reuses. `Default` is a
plain *name* string (a `Clusters` name, a `Servers` name or a raw `nats://`
URL) — it holds no connection settings. Shared connection settings live in the
top-level `Auth` / `Reconnect` blocks and are overridden per cluster; `RTD`
holds only the delivery tuning (there is no separate RTD/UDF server split).

```json
{
  "Default": "local",

  "Servers": {
    "n1": "nats://n1.example:4222",
    "n2": "nats://n2.example:4222",
    "n3": "nats://n3.example:4222",
    "dev": "nats://localhost:4222"
  },

  "Clusters": {
    "feed": {
      "urls": ["n1", "n2", "n3"],
      "auth": { "token": null },
      "jetstream": { "domain": "feed" }
    },
    "local": { "urls": ["dev"] },
    "vault": {
      "urls": ["nats://vault-a.example:4222", "nats://vault-b.example:4222"],
      "timeout": 2000,
      "auth": { "user": "excel", "password": null, "credsFile": null },
      "reconnect": { "maxAttempts": 120 }
    }
  },

  "Auth": {
    "user": null, "password": null, "token": null,
    "jwt": null, "nkeySeed": null, "credsFile": null,
    "tls": false, "tlsCert": null, "tlsCertPassword": null
  },

  "Reconnect": { "maxAttempts": 60, "waitMs": 2000, "jitterMs": 100 },

  "RTD": {
    "NatsUpdateRateMs": 1000,
    "ExcelUpdateRateMs": 100,
    "MessageCounterThreshold": 10000,
    "NatsUpdateStyle": "Automatic"
  },

  "JetStream": { "enabled": true, "defaultTimeoutMs": 5000 },

  "AllowAdminWrites": false,
  "SkipRepeatedMessages": true,
  "CoalesceRealtimeUpdates": true,
  "ConflationMs": 200,
  "PublishDedupCacheSize": 10000,
  "SyncWrite": "fireforget",
  "AsyncWrites": true,
  "UpdateCheck": true
}
```

This is the shape the shipped `NATSExcel.json` and `NATSExcel.schema.json`
define. (An earlier draft nested the settings under `Default` and had a per-
section `UDF`/`RTD.cluster`; `Config.cs` accepts a nested `Default` only as a
back-compat fallback and never reads a `UDF` block.)

### 8.1 Cluster resolution (naming a server or a connection string)

The optional host argument of every UDF/RTD call is a **cluster reference**,
resolved in the same spirit as RedisExcel's `ResolveHostCore`
(`Servers` alias → raw string), extended with `Clusters` at the top
(`AppConfig.ResolveEndpointCore`):

| Order | Input | Resolves to |
| --- | --- | --- |
| 1 | a `Clusters` name (`"feed"`) | the cluster's URL list (+ its auth/JetStream override) |
| 2 | a `Servers` alias (`"n1"`) | that server's URL (one connection) |
| 3 | a raw URL or URL list (`"nats://h:4222"`, `"nats://a,b"`) | used as-is |
| 4 | blank / omitted | the configured `Default` name (resolved through the same three steps), then the built-in `nats://localhost:4222` |

Rules:

- Matching is by **exact name first, then case-insensitive** (RedisExcel
  `ResolveHostCore` behavior), so a config with two spellings keeps its own
  priority.
- **No implicit fallback** to a hidden default when a non-empty name does not
  resolve: a dangling reference is a `#CONFIG` error, never a silent connection
  to `localhost` (the failure mode RedisExcel v1 had and v2 removed). A
  misconfigured `Default` *is* allowed to fall back to the built-in URL (it must
  not break every unauthored call); a dangling name the user typed is not.
- A reference that resolves in **both** `Clusters` and `Servers` resolves to the
  cluster (cluster wins). *(Designed, **not implemented**: no validation warning
  is emitted for the ambiguity today. Tracked in `PLAN.md` §9.)*
- A `Clusters` entry's `urls` accept `Servers` aliases **or** inline URLs, mixed.
- `auth`, `reconnect` and `timeout` may be set per cluster; each falls back to
  the shared top-level block, then to the built-in default. Per the shipped
  schema a cluster may also carry a `jetstream.domain` (the only cluster
  identifier NATS itself uses, to route JetStream API calls between clusters);
  `Config.cs` records it for diagnostics but nothing consumes it yet.
- `jetstream.domain` is the only cluster identifier NATS itself uses (to route
  JetStream API calls between clusters); it is exposed because a multi-cluster
  JetStream deployment needs it.

So all of these are valid in one spreadsheet, with nothing renamed:

```
=RTD("NatsRtd", , "SUB", "md.example.live.spot.>", , "feed")     // cluster
=RTD("NatsRtd", , "SUB", "md.example.live.spot.>", , "n1")       // server
=RTD("NatsRtd", , "SUB", "md.example.live.spot.>", , "nats://n1.example:4222")
=NatsUDFKVGet("prices", "BTC", "feed")
```

`SyncWrite` keeps the RedisExcel meaning adapted to NATS: `sync` = wait for the
JetStream ack / flush; `fireforget` = `PublishAsync` without awaiting the ack
(result-agnostic writes only).

**Authentication model (all NATS credential types supported).** The `auth`
block (under `Default` and optionally under each cluster) maps 1:1 to
`NATS.Client.Core.NatsAuthOpts`, so every credential style NATS accepts is
configurable:

| Type | Config fields | Notes |
| --- | --- | --- |
| No auth | (all null) | server with no `authorization` block |
| User / password | `user` + `password` | per-user; server may store a bcrypt hash |
| Token | `token` | single server-wide shared secret |
| NKey (seed) | `nkeySeed` | client signs the server nonce; nothing secret on the wire |
| JWT + seed | `jwt` + `nkeySeed` | operator mode (decentralized auth) |
| Credentials file | `credsFile` | a `.creds` file (JWT + seed in one file) |
| mTLS | `tls` + `tlsCert` + `tlsCertPassword` | the client certificate *is* the credential; `tlsCert` is the certificate path and `tlsCertPassword` its password |
| URL userinfo | `nats://user:pass@host` | also accepted; merged into `auth` |

These nine fields are exactly the set the shipped `NATSExcel.schema.json`
accepts under `auth` (under both the top-level shared `Auth` and each
`Clusters` entry). Precedence, highest first: `credsFile` → `jwt`+`nkeySeed` →
`nkeySeed` → `token` → `user`+`password` → URL userinfo → no auth. `password`
requires `user`; `jwt` requires `nkeySeed`; a combination the client rejects is
a `#CONFIG` error, not a silent fallback. TLS uses the platform certificate
store by default; the optional `tlsCert` (a `.pfx`/`.p12` bundle path) /
`tlsCertPassword` pair configures a client certificate for mTLS. A named
cluster's `auth` **merges over** the shared top-level `Auth` (replace per
field), so a shared token plus a per-cluster user/pass both work: a cluster that
declares only a `token` keeps the shared `user`/`password` for the fields it
did not set.

**Per-cluster auth is wired to the connection.** A formula is resolved to a URL
(list) before dispatch, so the connection manager receives the *resolved* URL,
not the name the user typed. `AppConfig.ResolveSettings(endpoint)` therefore
maps that URL back to the cluster that owns it (by name first, then by matching
the URL against each cluster's `urls`), and `NatsConnectionManager.BuildOptions`
applies the merged auth / timeout / TLS (mTLS `tlsCert` + `tlsCertPassword`)
for the endpoint it is about to open. Without that mapping a cluster that
declares its own token would silently connect with the shared auth — the
per-cluster block would be dead config.

**NKey seed: text *or* a `.nk` file path — read the file first.** `nkeySeed`
accepts either the seed text itself (an `SU...` string) **or** a path to a
`.nk` file that holds it. The NATS.Net client wants the *seed*, not the path:
handing it a path fails with `Invalid base32 character` (a filesystem path is
not base32). The loader must read a readable file and pass its trimmed
contents, otherwise pass the value through. The installer wizard labels the
field exactly this way ("the text, or a path to the .nk file"), so the code has
to honor it. The add-in implements this in `AppConfig.ResolveSeed` (`SU` prefix
→ use as-is; else a readable file → read and trim; else pass through), called
from `NatsConnectionManager.BuildAuth`; `tools/NatsValidate/ConnectionTester.cs`
(`ReadSeed`) is the reference the add-in copied.

**Masking** (`MaskNatsEndpoint`) covers every credential above: URL userinfo
(`nats://user:pass@host`), `password=`/`user=`/`token=`, a JWT (`eyJ...`), an
NKey seed (`SU...`), and a creds-file path (and a TLS client cert path). Every
error cell, RTD `#ERROR`, status cell and log line goes through it. A **cluster
name is not a secret** and is printed as-is; only the credentials it carries are
masked.

**Logged exceptions are masked too — do not revert this.** A NATS client
exception message can embed the connection URL (or a token) verbatim, so
logging the exception object with NLog's stock `${exception:...}` layout would
write the secret to `NATSExcel.log`, violating the "mask every log line" rule.
The add-in ships `MaskedExceptionRenderer` (a `[LayoutRenderer("maskedException")]`
registered in `AddIn.AutoOpen`), which flattens the exception chain through
NLog's `ExceptionLayoutRenderer` and runs the result through
`AppConfig.MaskNatsEndpoint`. `NLog.config` therefore uses
`${maskedException:format=toString}`, **not** `${exception:...}`. An unknown
renderer renders as an empty string, so an older config never throws; but a
future edit that switches back to `${exception:...}` re-opens the leak.

### 8.2 Transport, TLS and WebSocket (measured)

The add-in speaks the NATS protocol over **TCP by default** (`nats://host:4222`).
WebSocket (`ws://` / `wss://`) is supported for the cases TCP cannot reach — a
browser, a corporate HTTP proxy, a firewall that only allows 80/443 — and is
configured by simply using a `ws://` URL; there is no extra config key.

Transport credentials are per endpoint. Because the dispatcher resolves a
formula's cluster reference to a concrete URL *before* the connection exists,
`NatsConnectionManager.BuildOptions` looks the URL back up through
`AppConfig.ResolveSettings` and applies that cluster's merged auth, timeout and
TLS options (`TlsOpts` with `CertBundleFile`/`CertBundleFilePassword` for mTLS)
to the `NatsOpts` it builds (see §8.1). A raw `nats://user:pass@host` URL is
also accepted and masked like any other credential.

**Wire cost, measured (not assumed).** A proxy counted the real bytes on the
socket for the same publish+subscribe workload, using the actual market-data
payload shape (`{"venue":"example",...,"seq":1}`, 306 B):

| Payload | TCP | WebSocket | Difference |
| --- | --- | --- | --- |
| 64 B | 238.00 B/msg | 238.27 B/msg | +0.1% |
| 306 B (real feed) | 326.00 B/msg | 326.37 B/msg | +0.1% |
| 2048 B | 2068.00 B/msg | 2070.23 B/msg | +0.1% |

The NATS protocol framing is a ~20 B constant per message on TCP; the WebSocket
frame header plus the mandatory client mask adds a few bytes. **The two are
effectively equal in bandwidth** — WebSocket is about reachability, not savings.
(The server's `/connz` counters do NOT show this: `in_bytes` counts payload
only, so it reported both as identical. The measurement above is real socket
bytes.)

**permessage-deflate (WebSocket compression): not usable in this add-in.** Three
verified facts:

1. The server supports it, but only when explicitly enabled:
   `websocket { port: 8090, compression: true }`. Verified: with that on, the
   handshake answers `Sec-WebSocket-Extensions: permessage-deflate;
   server_no_context_takeover; client_no_context_takeover`; without it, no
   extension is negotiated.
2. The NATS .NET client does **not** implement compression — `WebSocketConnection`
   just calls `ClientWebSocket.SendAsync`/`ReceiveAsync`. It only exposes
   `NatsWebSocketOpts.ConfigureClientWebSocketOptions`, handing the caller the
   BCL's `ClientWebSocketOptions`. The only door to deflate there is
   **`DangerousDeflateOptions`, which exists only on .NET 7+**.
3. Excel-DNA targets .NET Framework 4.8 (see §13), and .NET Framework 4.8's
   `ClientWebSocketOptions` has **no** `DangerousDeflateOptions` (verified by
   reflection: only `UseDefaultCredentials`, `Credentials`, `Proxy`,
   `ClientCertificates`, `Cookies`, `KeepAliveInterval`). Setting the
   `Sec-WebSocket-Extensions` header by hand connects, but the runtime cannot
   compress or decompress frames, so it is never negotiated in practice.

Even if it worked, the gain would be modest and partly illusory: measured
single-message deflate on this payload shape saves **~27-30%** (209 B → 153 B;
312 B → 220 B), and the server negotiates `no_context_takeover`, so there is no
dictionary reuse across messages (the 90%+ figure only appears when the *same*
bytes repeat, a case `SkipRepeatedMessages`/`ConflationMs` already removes). The
server also skips compression below `wsCompressThreshold = 64` bytes.

**What actually moves the needle on bandwidth** (all already in the design):
`ConflationMs` (deliver fewer messages), `SkipRepeatedMessages` (drop identical
payloads), a smaller payload (the feed repeats `venue`/`environment`/
`marketType` on every message — ~30% of the body), and preferring TCP. TLS adds
a small constant per record in the opposite direction.

### 8.3 Clusters and JetStream topology
The add-in **consumes** a cluster; it does not form one. Cluster formation is a
server-side concern (a `cluster {}` block in each `nats.conf`). The add-in's job
is to be a *good NATS citizen*: trust discovery, expose what the server
advertises, and never model topology the server already models.

**Discovery-first (the NATS way).** The client exposes real cluster facts from
`INatsServerInfo`, so the add-in does not need to be told the topology:

| Field | Meaning | Used for |
| --- | --- | --- |
| `Cluster` | the cluster name the **server itself** announces | `NatsUDFClusterName()`, and naming a connection in the UI |
| `ClientConnectUrls` | the other nodes the server advertises | `NatsUDFClusterNodes()`, and the client's own reconnect pool |
| `ClusterDynamic` | membership can change at runtime | deciding whether the node list is a snapshot |
| `LameDuckMode` | the node is draining for a restart | a health signal a cell can show |
| `JetStreamAvailable` | JetStream is on for this server | the JS-disabled error path |

Consequence for config: **one URL is enough.** Point at any node of a healthy
cluster and the client discovers the rest. Listing two or three is a *resilience*
choice (survive the listed node being down at connect time), not a requirement.
This is the opposite of the Redis world where the client had to be taught the
shape; here the server teaches it.

- **A URL list, not one URL** (optional but recommended). `Clusters.<name>.urls`
  and `Servers.<alias>` accept a comma list:
  `nats://n1:4222,nats://n2:4222,nats://n3:4222`.
- **Reconnection knobs** (per cluster, optional). *(Designed, **not implemented**:
  the shipped schema accepts a per-cluster `reconnect` block
  (`maxAttempts`/`waitMs`/`jitterMs`) and `Config.cs` parses and merges it, but
  the connection manager does not map it onto `NatsOpts` yet — only the shared
  `timeout` becomes `ConnectTimeout`. Tracked in `PLAN.md` §9.)* The intended
  mapping is `maxAttempts` to `MaxReconnectRetry`, `waitMs` to
  `ReconnectWaitMin`/`Max`, `jitterMs` to `ReconnectJitter`.
- **`NoRandomize`**: NATS shuffles the URL list by default so many clients do not
  all hit node 1. Leave it off (the default); expose it only for deterministic
  diagnostics.

**JetStream in a cluster:** a stream's `num_replicas` decides its durability.
R1 = one node (lost with it). **R3 tolerates one node down** — verified: an R3
stream survived losing a node and kept serving from another. The add-in does
**not** choose replicas; the stream admin does, and `NatsUDFStreamCreate` exposes
`replicas`/`storage` explicitly. Two facts to surface:

- **A cluster too small rejects replicas it cannot place.** R3 on a 2-node
  cluster fails with `no suitable peers for placement (10005)`; that text must
  reach the cell **named**, not as a generic failure.
- **`jetstream.domain`** routes JetStream API calls between clusters; expose it
  per cluster for multi-JetStream deployments.

**Cluster identity: local label vs announced name.** The name under `Clusters`
is a **local label** the user chooses — it is what formulas say
(`"feed"`), and it never has to match the name the server announces. The two are
deliberately different concepts:

- `Clusters.<name>` — *your* handle for a set of URLs (and its auth/JetStream
  override). The auth override is the full `auth` field set from §8 (user/
  password, token, `nkeySeed`, `jwt`+`nkeySeed`, `credsFile`, and the mTLS
  `tls`/`tlsCert`/`tlsCertPassword` triple), merged field-by-field over
  `Default.auth`.
- `NatsUDFClusterName()` / `INatsServerInfo.Cluster` — *the server's* announced
  cluster name (blank for a standalone server).

To catch a config that points at the wrong cluster, a cluster may **optionally**
assert the name:

```json
"Clusters": {
  "feed": { "urls": ["n1", "n2", "n3"], "expectCluster": "FEED" }
}
```

Relationship between the two in practice (this is the intuitive default):

| Config label | `expectCluster` | Server announces | Result |
| --- | --- | --- | --- |
| `feed` | not set | `FEED` | OK, silent (label is local) |
| `feed` | `"FEED"` | `FEED` | OK, silent |
| `feed` | `"FEED"` | `OTHER` | **Warning** (not an error): the label resolves and connects, but the assertion fails — the cell/log says the endpoint is in cluster `OTHER`, not `FEED` |
| any | not set | (blank, standalone) | OK, silent — no cluster to assert |

On the common case the label is enough, so **matching the server's name is not
required** — `"feed"` for a cluster the server calls `FEED` is perfectly normal
and needs no `expectCluster`. The assertion is only for the operator who wants a
guard rail.

The check is **opt-in and non-fatal**: no `expectCluster` means no noise on the
common case (a standalone server announces no cluster), and a mismatch never
blocks a connection — it only makes the divergence visible where the config is
reviewed (`=NatsUDFConfigValidate()` output and the log).

Server-side reference (for the README, not the add-in): each node gets a
`cluster { name: <same>, listen: <route-port>, routes: [ <the others> ] }` block
plus its own `jetstream { store_dir: <durable, per node> }`; the `server_name`
must be unique per node and the cluster `name` must match. A one-line
`nats-server -c nats.conf -t` validates the file before starting it.

## 9. Concurrency and thread-safety rules

- Shared state touched by the Excel thread and client callbacks is thread-safe
  (`ConcurrentDictionary`, `Interlocked`, or a private lock).
- Timer callbacks (`System.Timers.Timer`) and NATS callbacks catch and log
  every exception; the `TickGate` prevents tick reentrancy.
- The subscription lifecycle uses per-key epochs so an unsubscribe racing a
  subscribe cannot install a ghost listener (the RedisExcel `ChannelLatest`
  pattern).
- Connections are created only in `NatsConnectionManager`; the shared wait
  helper is the only bounded-blocking primitive.

## 10. Testing strategy (5 layers, mirroring RedisExcel)

1. **Unit** — no NATS, no Excel. Config load/schema/pruning, masking, JSON
   conversions, the busy/refresh queue, arg parsing, `NonVolatile` signature
   parity, `TickGate`, conflation window, endpoint resolution. **Implemented**
   (`test\NATSExcel.Tests`, run with `dotnet test`).
2. **Smoke** — real `nats-server -js`, no Excel. Subscription ref-counting,
   wildcards, queue groups, request/reply, JetStream publish/consume, KV CRUD +
   CAS, object put/get, `SyncWrite` modes, dedup markers. *Planned.*
3. **Excel E2E** — hidden Excel instance, packed XLL, UDF/RTD assertions,
   multi-workbook case, node kill (failover), Excel restart, `...NonVolatile`
   single-run. **Implemented** (`test\Run-ExcelE2E.ps1`; 11-pass local run).
4. **Load** — `manager` vs raw `NATS.Client` baseline: throughput, allocated
   bytes/message, GC. **Implemented** (`test\LoadTests`; ~204k msg/s measured on
   a local manager run).
5. **Liveness** — adversarial: threadpool starvation, subscribe/dispose churn,
   server restart storms, JetStream consumer reconnects; enforced stall/resume
   bounds with a heartbeat watchdog. *Planned.*

A **cluster harness** (`test\TestSupport\Start-Cluster.ps1` / `Stop-Cluster.ps1`)
brings up a disposable 1- or 3-node `nats-server` for layers 2-5 and the E2E
script; `tools/NatsValidate` is the standalone config/connection checker.

CI (`.github/workflows/build.yml`) runs on `windows-latest`: it builds the
packed XLL, runs the unit layer, provisions `nats-server -js` (a single binary;
no package manager needed, unlike the Redis Memurai step) with a JetStream
probe, and builds the installer on every run. The smoke and liveness steps each
start their own server and run for real (`test\SmokeTests` / `test\LivenessTests`);
a `v*` tag adds the CHANGELOG check, the release-asset assertion and the GitHub
release.

## 11. Installer and CI

**NATS runs natively on Windows — Docker is not required.** The server is a
single self-contained `nats-server.exe` (no dependency runtime, no service
manager), which is a real advantage over the Redis picture in RedisExcel (where
Windows means Docker, WSL or Memurai). Server provisioning therefore has a
native-first default; Docker stays an *optional* convenience for users who
already run it.

Native install options (README "Run a server on Windows" guide, mirroring
RedisExcel's beginner guide):

1. **winget** (simplest): `winget install NATSAuthors.NATSServer`.
2. **GitHub release zip** (works everywhere, no package manager):
   download `nats-server-<ver>-windows-amd64.zip`, unzip, run `nats-server.exe`.
   This is also what CI uses.
3. **Windows service** (run at boot, optional): `nats-server` ships no installer;
   register it with `sc.exe create NATS binPath= "<path>\nats-server.exe -c
   <path>\nats.conf" start= auto`.
4. **Docker** (optional): `docker run -p 4222:4222 nats -js` — offered, not
   assumed.

- **Inno Setup**, localized en/es/pt-BR, per-user, default folder
  `C:\NATSExcel\`: copies the matching XLL and writes a starter
  `NATSExcel.json`. The optional "install a server" step is **native-first**: it
  downloads the `windows-amd64` zip into the install folder (or uses an existing
  `nats-server` on `PATH`), writes a starter `nats.conf` with
  `jetstream { store_dir: ... }` pointing at a **durable** directory, and can
  register/start the Windows service. Docker is offered only as a fallback
  option. The installer never requires admin except for the service step.
- CI (`build.yml`) on `windows-latest`: build the add-in with `dotnet build`,
  run the unit tests, **download the `nats-server-<ver>-windows-amd64.zip`** and
  run it with `-js --store_dir <temp>` (no Chocolatey/package manager, since the
  community `nats-server` choco package is stale, last published 2019), probe
  port 4222 and require `jetstream:true`, build the installer on every run, and
  on a `v*` tag check the CHANGELOG heading, assert the release assets and
  publish them. The smoke and liveness steps exist but skip while those test
  projects are absent; the unit step is unconditional.
- Release assets: `NATSExcel-packed.xll`, `NATSExcel64-packed.xll`,
  `NATSExcel.dll`, `NLog.config`, `NATSExcel.json`,
  `NATSExcelSetup-<version>.exe`.

**JetStream install note** (for the README): JetStream is the same
`nats-server` binary; enable it with `-js` (or `jetstream {}` in the config).
The one real care is `store_dir`: its default is a temp path, so production
(and the installer's server option) must point it at durable storage. The add-in
detects a server without JetStream and reports it per function, it never fails
silently.

## 12. Versioning and public-surface freeze

- The UDF names, argument order, RTD commands and ProgID `NatsRtd` are frozen
  at v1.0.0. New work is additive (new functions, new `...NonVolatile` pairs).
- The `NATSExcel.json` shape is backward compatible: the top-level `Default`
  (a name), `Servers`, `Clusters`, the shared `Auth` and `Reconnect` blocks,
  `RTD`, `JetStream`, and the flat toggles (`AllowAdminWrites`,
  `SkipRepeatedMessages`, `CoalesceRealtimeUpdates`, `ConflationMs`,
  `PublishDedupCacheSize`, `SyncWrite`, `AsyncWrites`, `UpdateCheck`). An
  earlier draft that nested the connection settings under `Default` is still
  parsed as a back-compat fallback.

## 13. Open decisions

0. **Runtime: stay on .NET Framework 4.8 — DECIDED (verified).** Excel-DNA 1.9.0
   supports `net6.0-windows` through `net10.0-windows` as well as .NET Framework
   4.6.2-4.8.1, so a modern runtime IS possible. It is deliberately not taken:

   | | net48 (chosen) | .NET 6+ |
   | --- | --- | --- |
   | Runtime install on the user's machine | none (a Windows component) | the matching **.NET Desktop Runtime**, per bitness |
   | Add-in isolation | separate AppDomain; coexists with any other add-in | no AppDomain ("weak isolation") |
   | Coexistence | always | **only one .NET core runtime per Excel process** — a version mismatch with another add-in blocks a load (mitigated, not fixed, by `<RollForward>`) |
   | Microsoft support horizon | none announced (as long as Windows) | a new major every year, LTS/STS alternating |
   | Modern C# / BCL | no | yes |

   The Excel-DNA author's own guidance is to target .NET Framework for an add-in
   "distributed outside your organization, or on machines you don't have much
   control over" — exactly this add-in's case. The concrete thing a modern
   runtime would buy is WebSocket permessage-deflate (§8.2), which saves ~27% on
   a payload the add-in already trims via `ConflationMs`/`SkipRepeatedMessages`
   — not worth a mandatory runtime install plus a cross-add-in version hazard.
   Revisit only if a different requirement forces a modern runtime (modern TLS,
   `System.Text.Json`, new C#); deflate then comes along for free.

1. **Language version — RESOLVED (verified).** A spike built `net48` against
   `NATS.Client.*` 3.3.0 and connected to a live server: publish + `await foreach`
   subscribe round-tripped. Confirmed empirically:
   - `IAsyncEnumerable` consumption needs **C# 8+** (`await foreach`), unlike
     RedisExcel's pinned 7.3.
   - The client's option objects use **init-only setters** (`NatsOpts { Url = ... }`),
     which need **C# 9**, not 8 (C# 8 fails with error CS8400).
   - `net48` also requires a small `IsExternalInit` shim (`namespace
     System.Runtime.CompilerServices { internal static class IsExternalInit { } }`),
     because init-only setters need that type and .NET Framework does not ship it.
   Proposal: **`net48` with `<LangVersion>9.0</LangVersion>` + the `IsExternalInit`
   shim.** `LangVersion` only sets the compiler dialect; it does not retarget the
   runtime, so `net48` stays.

Reviewer note (do not lose this): the `.dna` `<Reference>` list shipped complete
(31 dependencies, each `Pack="true"`); see decision 6.
2. **Config key names — DECIDED.** Use NATS-neutral names:
   `RTD.NatsUpdateRateMs` and `RTD.NatsUpdateStyle` (not `RedisUpdateRateMs` /
   `ExcelUpdateStyle`). The connection target is named by the top-level `Default`
   (a `Clusters`/`Servers` name or a raw URL), with per-endpoint settings in
   `Clusters`/`Servers` and shared `Auth`/`Reconnect`; the earlier `Nats.url` /
   `Nats.host` top-level pair was dropped. Corrected in §8; see `PLAN.md`
   decision D5.
3. **Object Store in a cell — RESOLVED (implemented).** `NatsUDFObjectGet`
   rejects a value over **1 MiB** (`MaxObjectCellBytes = 1_048_576`) or a binary
   object with a named `Error:`; a text cell is bounded by
   `MaxCellTextLength = 32767`.
4. **License.** NATS.Net is Apache-2.0; pick a license for NATSExcel before the
   first public release (RedisExcel ships none).
5. **Admin gate default — DECIDED.** `AllowAdminWrites` defaults to `false`; the
   destructive functions return the exact gate error until it is set true.
6. **Packed dependency set — RESOLVED (the P0 gate passed).** The `net48` build
   against `NATS.Client.*` 3.3.0 (Core + JetStream + KeyValueStore +
   ObjectStore) restores exactly the set enumerated in `NATSExcel.dna` (31 DLLs,
   each `Pack="true"`): `NATS.Client.Abstractions`, `NATS.NKeys`,
   `Microsoft.Bcl.AsyncInterfaces`, `Microsoft.Bcl.Memory`,
   `Microsoft.Extensions.Logging.Abstractions`, `Microsoft.Extensions.Primitives`,
   `Microsoft.VisualStudio.Threading`/`.Validation`, `Nerdbank.Streams`,
   `System.Buffers`, `System.Diagnostics.DiagnosticSource`, `System.IO.Pipelines`,
   `System.Memory`, `System.Numerics.Vectors`,
   `System.Runtime.CompilerServices.Unsafe`, `System.Text.Json`,
   `System.Text.Encodings.Web`, `System.Threading.Channels`,
   `System.Threading.Tasks.Extensions`, `System.ValueTuple`,
   `Microsoft.Win32.Registry`, `System.ComponentModel.Annotations`,
   `System.Security.AccessControl`, `System.Security.Principal.Windows`, plus the
   Excel-DNA/Newtonsoft/NLog set. `System.Text.Json` and its transitive closure
   sit next to the Newtonsoft dependency (both are packed). The packed XLL loads
   in real Excel and answers UDFs, so the list is demonstrably complete;
   `DnaPackagingParityTests` (PLAN V8.2) is the tripwire against future drift.
