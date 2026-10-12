# NATSExcel — Design

Excel add-in (XLL) that integrates **NATS** with Excel through RTD (Real-Time
Data) and UDF (User Defined Functions), in the same spirit as the sibling
project **RedisExcel**: the user types a formula, NATS data appears in the cell
and stays live.

This is the day-one design contract. The public Excel surface defined here is
frozen before the first release: renames and argument changes break users'
spreadsheets forever (see the golden rules in `AGENTS.md`).

- Status: design draft for discussion.
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

Endpoint resolution for any host argument: a `Servers` alias, then the raw
`nats://` URL / URL list. No implicit fallback to a hidden default.

## 5. Architecture

### 5.1 File layout

| File | Responsibility |
| --- | --- |
| `NatsRuntime.cs` | Process-wide singleton: one `NatsConnection` per endpoint, plus the subscription manager. Shutdown fence and same-process reload, mirroring `RedisRuntime`. |
| `NatsConnectionManager.cs` | Creates/caches connections by resolved endpoint; connection state, reconnect observation, jetstream context cache, bounded wait helper. |
| `NatsSubscriptionManager.cs` | Ref-counted Core subscriptions: one client subscription per `(endpoint, subject, queueGroup)`, N listeners; `IDisposable` tokens; wildcard-aware; origin tags (`RTD`/`UDF`). |
| `NatsRtd.cs` | `IExcelAddIn` (`AutoOpen`/`AutoClose`) + `ExcelRtdServer`: topic registry, poll timer, Excel flush timer, backoff, wildcard topics, JetStream/KV topic kinds. |
| `NatsUdfCore.cs` | Core NATS UDFs: publish, request/reply, subject-latest, unsubscribe. |
| `NatsUdfJetStream.cs` | JetStream UDFs: publish-with-ack, get/fetch, stream/consumer info, admin. |
| `NatsUdfKv.cs` | Key/Value UDFs: get/put/create/update/delete/keys/history/watch. |
| `NatsUdfObjectStore.cs` | Object Store UDFs: put/get/list/watch/links. |
| `NatsUdfStatus.cs` | Status and diagnostics (`NatsUDFConnectionCount`, `NatsUDFServerInfo`, `NatsUDFRtt`, ...). |
| `NatsUdfAsync.cs` | Shared async bridge: `ExcelAsyncUtil.Observe` wrapper + per-endpoint FIFO write queue + bounded sync wait. |
| `NatsJson.cs` | Matrix ⇄ JSON conversions (RedisExcel's `ExcelJson`, unchanged in spirit). |
| `NatsResultFormatter.cs` | Value shaping for cells (JSON objects, KV entries, object metadata). |
| `Conflation.cs` | Real-time conflation window (`ConflationMs`): latest value wins per window. |
| `TickGate.cs` | Non-blocking reentrancy gate for timers + `StripedLocks`. |
| `Config.cs` / `ConfigSchema.cs` | Config load (once per process) + embedded JSON Schema, tolerant validation, endpoint/auth masking. |
| `UpdateCheck.cs` | Background GitHub release check. |
| `NatsExcel.dna` | Excel-DNA packing manifest. |
| `installer/` | Inno Setup script + helper PowerShell + README. |

### 5.2 Packages

Explicit packages (not the `NATS.Net` meta package) so the packed XLL stays
lean; the exact set is confirmed at P0:

- `NATS.Client.Core` — connection, pub/sub, request/reply.
- `NATS.Client.JetStream` — streams, consumers, message get/fetch.
- `NATS.Client.KeyValueStore` — KV buckets.
- `NATS.Client.ObjectStore` — object buckets.
- (transitive: `NATS.Client.Abstractions`, `Microsoft.Extensions.Logging.Abstractions`, `Microsoft.Bcl.AsyncInterfaces`.)

Serialization: the add-in already depends on **Newtonsoft.Json**. Rather than
pull `System.Text.Json` (extra DLLs in the XLL), implement a small
`INatsSerializerRegistry` backed by Newtonsoft and use it for typed publishes.
Verify the packed dependency set against the `.dna` at P0 and keep it minimal,
like RedisExcel's deliberate list.

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

Notes:

- **No `PSUB`.** NATS wildcards are part of the subject (`orders.*`,
  `orders.>`); a pattern subscription is just `SUB orders.>`.
- A `SUB`/`KV`/`JS` topic that fails to subscribe shows `#ERROR` and is retried
  with 1s..30s backoff, capped per tick (RedisExcel behavior).
- Polled kinds (`KVGET`, `KVKEYS`, `MSG`) retry on every tick.
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

`NatsUDFConnectionCount`, `NatsUDFServerInfo` (server id/version/connect_urls),
`NatsUDFStats` (in/out msgs/bytes), `NatsUDFRtt`, `NatsRTDConnectionCount`,
`NatsRTDTopicCount`, `NatsRTDSubscriptionCount`, `NatsRTDSubjectCount`,
`NatsRTDDefaultHost`, `NatsUDFUpdateAvailable`,
`NatsUDFMatrixToJSON` / `NatsUDFJSONToMatrix` (pure, non-volatile),
`NatsCheckConfigFile`, `NatsUDFConfigSchema`, `NatsUDFConfigPath`,
`NatsUDFConfigValidate`, `NatsUDFConfigReload`.

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
| `NatsRTDCluster([what], [host])` | RTD | A live cluster-heath cell (`what` = `leader`, `lag`, `nodes`); pushes on change. |

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

**95 `[ExcelFunction]` in total** once the cluster tier is included: 88 from
§7.1–§7.7 plus **7** new reads (`ClusterName`, `ClusterNodes`,
`ConnectionState`, `ClusterInfo`, `ClusterHealth`, `JSRaftInfo`,
`JSStreamPlacement`). Of the 88, 60 are volatile (25 volatile write bases + 35
reads/status), 25 carry a `...NonVolatile` twin, and 3 are standalone
non-volatile (`MatrixToJSON`, `JSONToMatrix`, `ConfigReload`); the 7 cluster
reads are volatile and have no twins (reads do not need them). RTD commands
become **9** with `NatsRTDCluster`. The pinned count is set by the
public-surface baseline test, which is the authority; an earlier draft said
"roughly 70".

## 8. Configuration

File name: **`NATSExcel.json`**. Search order identical to RedisExcel (add-in
folder → `%USERPROFILE%` → Excel folder → `C:\Windows`; first existing wins).
Comments (`//`, `/* */`) accepted; validated against an embedded JSON Schema;
a wrong-typed value is dropped individually.

Layout principle: **the user picks a cluster, not a host.** `Clusters` is the
unit a formula names; `Servers` is the primitive it reuses. `Default` holds
process-wide connection settings; `RTD`/`UDF` keep only their own tuning and an
optional cluster override.

```json
{
  "Default": {
    "cluster": "local",
    "timeout": 1000,
    "auth": { "token": null, "user": null, "password": null },
    "reconnect": { "maxAttempts": 60, "waitMs": 2000, "jitterMs": 100 }
  },

  "RTD": {
    "NatsUpdateRateMs": 1000,
    "ExcelUpdateRateMs": 100,
    "MessageCounterThreshold": 10000,
    "NatsUpdateStyle": "Automatic",
    "cluster": "feed"
  },

  "UDF": { "timeout": 1000 },

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
      "urls": ["nats://10.0.0.7:4222", "nats://10.0.0.8:4222"],
      "auth": { "user": "excel", "password": null, "credsFile": null }
    }
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

### 8.1 Cluster resolution (naming a server or a connection string)

The optional host argument of every UDF/RTD call is a **cluster reference**,
resolved in the same spirit as RedisExcel's `ResolveHostCore`
(`Servers` alias → raw string), extended with `Clusters` at the top:

| Order | Input | Resolves to |
| --- | --- | --- |
| 1 | a `Clusters` name (`"feed"`) | the cluster's URL list (+ its auth/JetStream override) |
| 2 | a `Servers` alias (`"n1"`) | that server's URL (one connection) |
| 3 | a raw URL or URL list (`"nats://h:4222"`, `"nats://a,b"`) | used as-is |
| 4 | blank / omitted | `RTD.cluster` / `UDF.cluster` → `Default.cluster` → the built-in `nats://localhost:4222` |

Rules:

- Matching is by **exact name first, then case-insensitive** (RedisExcel
  `ResolveHostCore` behavior), so a config with two spellings keeps its own
  priority.
- **No implicit fallback** to a hidden default when a non-empty name does not
  resolve: a dangling reference is a `#CONFIG` error, never a silent connection
  to `localhost` (the failure mode RedisExcel v1 had and v2 removed).
- A reference that resolves in **both** `Clusters` and `Servers` triggers a
  **validation warning** (cluster wins), so the ambiguity is visible.
- A `Clusters` entry's `urls` accept `Servers` aliases **or** inline URLs, mixed.
- `auth`, `jetstream` and `reconnect` may be set per cluster; each falls back to
  the `Default` block, then to the built-in default.
- `jetstream.domain` is the only cluster identifier NATS itself uses (to route
  JetStream API calls between clusters); it is exposed because a multi-cluster
  JetStream deployment needs it.

So all of these are valid in one spreadsheet, with nothing renamed:

```
=RTD("NatsRtd", , "SUB", "md.binance.live.spot.>", , "feed")     // cluster
=RTD("NatsRtd", , "SUB", "md.binance.live.spot.>", , "n1")       // server
=RTD("NatsRtd", , "SUB", "md.binance.live.spot.>", , "nats://n1.example:4222")
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
| mTLS | `tls` + a client certificate (see below) | the certificate *is* the credential |
| URL userinfo | `nats://user:pass@host` | also accepted; merged into `auth` |

Precedence, highest first: `credsFile` → `jwt`+`nkeySeed` → `token` →
`user`+`password` → URL userinfo → no auth. `password` requires `user`; `jwt`
requires `nkeySeed`; a combination the client rejects is a `#CONFIG` error, not
a silent fallback. TLS uses the platform certificate store by default; an
optional `auth.tlsCert` (path) / `auth.tlsCertPassword` pair configures a client
certificate for mTLS. A cluster's `auth` **merges over** the `Default.auth`
(replace per field), so a shared token in `Default` plus a per-cluster user/pass
both work.

**Masking** (`MaskNatsEndpoint`) covers every credential above: URL userinfo
(`nats://user:pass@host`), `password=`/`user=`/`token=`, a JWT (`eyJ...`), an
NKey seed (`SU...`), and a creds-file path (and a TLS client cert path). Every
error cell, RTD `#ERROR`, status cell and log line goes through it. A **cluster
name is not a secret** and is printed as-is; only the credentials it carries are
masked.

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
- **Reconnection knobs** (per cluster, optional), mapped to `NatsOpts`:
  `maxReconnectAttempts` (`MaxReconnectRetry`), `reconnectWaitMs`
  (`ReconnectWaitMin`/`Max`), `reconnectJitterMs` (`ReconnectJitter`),
  `connectTimeoutMs` (`ConnectTimeout`). Defaults are fine; expose them because a
  cell that silently stops updating is the failure mode we care about.
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
  override).
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
reviewed (`=NatsCheckConfigFile()` output and the log).

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
   parity, `TickGate`, conflation window, endpoint resolution.
2. **Smoke** — real `nats-server -js`, no Excel. Subscription ref-counting,
   wildcards, queue groups, request/reply, JetStream publish/consume, KV CRUD +
   CAS, object put/get, `SyncWrite` modes, dedup markers.
3. **Excel E2E** — hidden Excel instance, packed XLL, UDF/RTD assertions,
   multi-workbook case, connection blip (restart the server), write modes and
   async pending marker.
4. **Load** — `manager` vs raw `NATS.Client` baseline: throughput, allocated
   bytes/message, GC.
5. **Liveness** — adversarial: threadpool starvation, subscribe/dispose churn,
   server restart storms, JetStream consumer reconnects; enforced stall/resume
   bounds with a heartbeat watchdog.

CI runs `nats-server -js` (a single binary; no package manager needed, unlike
the Redis Memurai step) plus unit → smoke → liveness(quick).

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
- CI (`build.yml`) on `windows-latest`: build with msbuild, unit tests,
  **download the `nats-server-<ver>-windows-amd64.zip`** and run it with
  `-js --store_dir <temp>` (no Chocolatey/package manager — the community
  `nats-server` choco package is stale, last published 2019), probe port 4222,
  run smoke and liveness(quick), build the installer, and on a `v*` tag assert
  the release assets and publish them.
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
- The `NATSExcel.json` shape is backward compatible (`Nats`/`RTD`/`UDF`/
  `Servers`/`Auth`/`JetStream`).

## 13. Open decisions

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

Reviewer note (do not lose this): the `.dna` `<Reference>` list is currently not
complete — see decision 6.
2. **Config key names — DECIDED.** Use NATS-neutral names:
   `RTD.NatsUpdateRateMs` and `RTD.NatsUpdateStyle` (not `RedisUpdateRateMs` /
   `ExcelUpdateStyle`); `Nats.url` (with `Nats.host` accepted as a documented
   alias). v1 has no users, so there is no compatibility burden — the clean name
   is chosen now. Corrected in §8; see `PLAN.md` decision D5.
3. **Object Store in a cell.** Objects can be large; a string return must have a
   hard size bound (error past it), and binary objects cannot be a text cell.
   Confirm the bound and the behavior for binary.
4. **License.** NATS.Net is Apache-2.0; pick a license for NATSExcel before the
   first public release (RedisExcel ships none).
5. **Admin gate default.** `AllowAdminWrites` default `false` (proposed) vs
   `true` ("do everything"). Destructive functions argue for `false`.
6. **Packed dependency set (verified, needs review).** A `net48` build against
   `NATS.Client.*` 3.3.0 (Core + JetStream + KeyValueStore + ObjectStore) restores
   this runtime dependency set: `NATS.Client.Abstractions`, `NATS.NKeys`,
   `Microsoft.Bcl.AsyncInterfaces`, `Microsoft.Bcl.Memory`,
   `Microsoft.Extensions.Logging.Abstractions`, `Microsoft.Extensions.Primitives`,
   `Nerdbank.Streams`, `System.Buffers`, `System.Diagnostics.DiagnosticSource`,
   `System.IO.Pipelines`, `System.Memory`, `System.Numerics.Vectors`,
   `System.Runtime.CompilerServices.Unsafe`, `System.Text.Json`,
   `System.Text.Encodings.Web`, `System.Threading.Channels`,
   `System.Threading.Tasks.Extensions`, `System.ValueTuple`, plus
   `Microsoft.Win32.Registry`, `System.ComponentModel.Annotations`,
   `System.Security.AccessControl`, `System.Security.Principal.Windows`
   (`Microsoft.VisualStudio.Threading`/`.Validation` appeared on the local .NET 9
   SDK restore; the CI .NET Framework path may differ). `System.Text.Json` and its
   transitive closure are a noticeable addition next to the existing Newtonsoft
   dependency — confirm the exact set on the CI toolchain (msbuild/VS), decide
   whether to keep both JSON stacks or standardize, then enumerate every one of
   them in `NATSExcel.dna` with `Pack="true"` (an unpacked DLL is a runtime
   `FileNotFoundException` inside the XLL).
