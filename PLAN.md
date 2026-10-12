# NATSExcel — Implementation Plan

Fine-grained build plan, ready to execute. Companion to `DESIGN.md` (the
architecture contract) and `AGENTS.md` (working guidance). Nothing here is
implemented yet.

Three workstreams run in parallel once the **P0 gate** is green:

- **F — Foundation / runtime core**: scaffolding, packing, config, connections,
  subscriptions, async bridge.
- **S — Excel surface**: RTD server + the UDF catalog + formatting.
- **V — Verification / delivery**: the 5 test layers, installer, CI, docs.

Consolidated from three independent plans; interface names were reconciled
(section 1.4) where the three drafts disagreed.

---

## 0. Decisions taken up front

| # | Decision | Rationale |
| --- | --- | --- |
| D1 | Test projects **compile the production sources directly** (linked `Compile`), not `ProjectReference`. | A `ProjectReference` would run the Excel-DNA packing on every test run. But instead of four hand-maintained lists (RedisExcel's drift risk, worse with ~20 files), the item group lives once in `test\TestSupport\ProductionSources.props`, imported by all four test projects. The parity test asserts coverage **and** that all four import it. |
| D2 | `NatsRtd.cs` is the only intentional test exclusion (needs Excel); covered by E2E. | Same as RedisExcel. |
| D3 | The native `nats-server.exe` is a first-class test dependency via `NativeServerHarness` (no Docker/WSL/Chocolatey). | It is a single self-contained binary; the restart storm and CI both use it. |
| D4 | JetStream-disabled runs are **attributable, never a false green**: `JetStreamProbe` classifies `ENABLED`/`DISABLED`/`UNREACHABLE`, every suite prints a banner, and `--require-jetstream` turns a skip into a failure in CI. | A mis-started CI server must not silently skip JS/KV/OBJ. |
| D5 | Config keys are NATS-neutral: `NatsUpdateRateMs` / `NatsUpdateStyle` (not `RedisUpdateRateMs`). | `DESIGN.md` §8 copied the RedisExcel key name; corrected there. v1 has no users, so the clean name is free now. |
| D6 | A **remote/real server is read-only by default**. Pointing any suite at a non-local endpoint runs subscribe/read/request paths only; every write, admin and destructive step is skipped, and the endpoint arrives **only as a CLI argument** (never committed). | A live feed server is not a test sandbox: a stray publish or purge is a production incident, and a private host in the repo is a leak (golden rule 9). Mirrors RedisExcel's `-RealChannel`/`-SkipClientKill` discipline, strengthened. |
| D7 | **Support every NATS credential type** in `Auth` (user/password, token, NKey seed, JWT+seed, creds file, URL userinfo, mTLS), mapped 1:1 to `NatsAuthOpts`. | A single field set covers all servers we may face (a token server today, operator-mode/mTLS later), and one precedence rule + masking path keeps it safe. |
| D8 | **Clusters are server-side; the client only discovers.** The add-in never forms a cluster, infers topology, or creates cluster assets implicitly. | Verified behaviour: the client auto-discovers peers from `ClientConnectUrls` and reconnects to a survivor. Topology is an operator concern; the add-in treats the cluster as an opaque endpoint set that reports its own shape. |
| D11 | **Discovery-first cluster awareness.** Use the facts the server advertises (`INatsServerInfo.Cluster`, `ClientConnectUrls`, `ClusterDynamic`, `LameDuckMode`, `JetStreamAvailable`) rather than making the user declare topology; **one URL is enough**, a list is only for connect-time resilience. | This is the "don't suffer like Redis" decision: NATS exposes the cluster on the wire, so the add-in must consume it instead of asking the user to model it. It also makes `ClusterName`/`ClusterNodes`/`ConnectionState` free — no `$SYS`, no extra config. |
| D12 | **A `Clusters` key is a local label, not the server's cluster name.** It is what formulas say and never has to match the announced name; an optional `expectCluster` asserts it and warns (non-fatally) on mismatch. | The intuitive default: users name a cluster for themselves (`"feed"`), and the server's announced name (`FEED`) is a different thing. Requiring them to match would be surprising; ignoring divergence entirely would hide a wrong-cluster config. Opt-in assertion is the middle ground. |
| D10 | **The user picks a `Clusters` name, not a host.** Config is `Default` + `Servers` + `Clusters`; the optional host argument resolves cluster name → server alias → raw URL/list → the default cluster. | NATS topology is a first-class concept (unlike RedisExcel's single host), so the natural user unit is a cluster. It keeps the RedisExcel `ResolveHostCore` chain (alias → raw string) and its exact-then-case-insensitive matching, and adds a `#CONFIG` error for a dangling name instead of a silent `localhost`. |
| D9 | **Both delivery targets are validated for install/test: native Windows and Docker.** Native is the default first-class path; Docker is a convenience. | Both verified end-to-end (below). Native avoids a Docker prerequisite on user machines and in CI; Docker is handy for a quick throwaway server. |

---

## 1. Frozen contracts (freeze before any parallel work)

### 1.1 Public surface counts (pinned)

| Set | Count |
| --- | --- |
| `[ExcelFunction]` total | **88** |
| Volatile | **60** (25 write bases + 35 reads/status) |
| Writes with a `...NonVolatile` twin | **25** |
| Reads/status without a twin | **35** |
| Standalone non-volatile (no twin) | **3** (`MatrixToJSON`, `JSONToMatrix`, `ConfigReload`) |
| RTD commands | **8** (`SUB`, `KV`, `KVGET`, `KVKEYS`, `REQ`, `JS`, `MSG`, `OBJ`) |
| Functions gated by `AllowAdminWrites` | **11** |

`DESIGN.md` §7.7 says "roughly 70"; the pinned number is **88** and the
baseline test is the authority. The full per-function signature catalog is the
S workstream's `docs/PUBLIC-SURFACE.md`, produced in S0.

### 1.2 RTD command grammar (frozen)

| Command | `topicInfo` (1-based) | Kind | Wildcards |
| --- | --- | --- | --- |
| `SUB` | `[1]subject [2]queueGroup [3]endpoint` | push | yes (`*` 1 token, `>` 1+) |
| `KV` | `[1]bucket [2]key [3]endpoint` | push (watch) | no |
| `KVGET` | `[1]bucket [2]key [3]endpoint` | poll | no |
| `KVKEYS` | `[1]bucket [2]endpoint` | poll | no |
| `REQ` | `[1]subject [2]payload [3]everyMs [4]endpoint` | poll | no |
| `JS` | `[1]stream [2]consumer [3]filter [4]endpoint` | push | filter only |
| `MSG` | `[1]stream [2]filter [3]endpoint` | poll | filter only |
| `OBJ` | `[1]bucket [2]object [3]endpoint` | push (metadata) | no |

Frozen parser rules (pinned by `RtdTopicGrammarTests`):

- Command token: `ToUpperInvariant().Trim()`; unknown →
  `unknown command 'X', expected one of [SUB, KV, KVGET, KVKEYS, REQ, JS, MSG, OBJ]`.
- Only a **non-blank** token past the documented arity is an error.
- Blank subject/bucket/stream/consumer/object → specific `... is required`,
  rejected **before** registration (never a retry storm).
- `REQ everyMs`: invariant integer, clamped to `>= 100`.
- **Optional-middle positional rule**: when the optional middle token is blank
  and a trailing token remains, the trailing token is the endpoint **only if**
  it resolves (a `Servers` alias, or contains `://`); otherwise it is the
  middle value. Documented workaround: spell out all tokens.
- RTD cells are **always scalar**; lists/structures are compact JSON text.

### 1.3 Markers and error vocabulary (frozen strings)

Markers: `(ConnectData)`, `(subscribed)`, `(no value)`, `(null)`, `No change`,
`OK`, `OK (fire and forget)`, `OK (fire and forget: all)`, `TIMEOUT`,
`No Responders`, `#CONFIG`.

UDF errors: `Error: <masked context>: <masked cause> | <hint>`.
RTD errors: `#ERROR: ConnectData: <masked>` and
`#ERROR: JetStream is not enabled on this server`.

Named causes: `JetStream is not enabled on this server` · `stream 'x' not found` ·
`consumer 'x' not found on stream 'y'` · `wrong last revision for key 'k'
(expected N)` · `admin writes are disabled (set "AllowAdminWrites": true)`.

Argument validation (terse, stable, `Error: ` prefix) enumerates: missing
subject/bucket/key/object/stream/consumer/range, bad range shape, mismatched
cells, non-integer/out-of-range numeric, non-text endpoint, invalid `timeoutMs`/
`count`/`revision`, `Excel error cells are not valid arguments`, `A multi-cell
range is not a valid scalar argument`.

Bounds: `MaxCellTextLength = 32767`, `MaxObjectCellBytes = 1_048_576`.

### 1.4 Reconciled interface contracts

**`AppConfig`** (`Config.cs`):

```csharp
public enum EnumExcelUpdateStyle { Timer, Realtime, Automatic }

public class NatsSection   { string url; int timeout; }          // url; legacy alias "host"
public class RTDConfig     { int NatsUpdateRateMs; int ExcelUpdateRateMs;
                             long MessageCounterThreshold; EnumExcelUpdateStyle NatsUpdateStyle; }
public class UDFConfig     { int timeout; }
public class AuthConfig    { string user, password, token, jwt, nkeySeed, credsFile;
                             bool tls; string tlsCert, tlsCertPassword; }
public class JetStreamConfig { bool enabled; int defaultTimeoutMs; }
public class ConfigRoot    { NatsSection Nats; RTDConfig RTD; UDFConfig UDF;
                             Dictionary<string,string> Servers; AuthConfig Auth;
                             JetStreamConfig JetStream; bool AllowAdminWrites;
                             bool SkipRepeatedMessages; bool CoalesceRealtimeUpdates;
                             int? ConflationMs; int PublishDedupCacheSize;
                             string SyncWrite; bool AsyncWrites; bool UpdateCheck; }

public static class AppConfig {
  public  const  string ConfigFileName = "NATSExcel.json";
  public  const  string DefaultUrl     = "nats://localhost:4222";
  public  static ConfigRoot Current { get; }                 // Lazy, once per process
  public  static string ResolveEndpoint(string endpoint, string defaultEndpoint);
  public  static string ResolveRtdEndpoint(string endpoint);
  public  static string ResolveUdfEndpoint(string endpoint);
  public  static string MaskNatsEndpoint(string text);
  public  static string SyncWrite { get; }
  public  static bool   AsyncWrites { get; }
  public  static bool   AllowAdminWrites { get; }
  public  static bool   JetStreamEnabled { get; }
  public  static int    JetStreamTimeoutMs { get; }
  public  static string ValidateConfiguration(string path = null);
  public  static string JsonSchemaText { get; }
  internal static ConfigRoot CurrentOverrideForTests;
  internal static ConfigRoot ParseTolerant(string json, Action<Exception> onError, IList<ConfigIssue> issues);
  internal static ConfigRoot LoadFromPaths(IEnumerable<string> paths);
  internal static ConfigRoot Sanitize(ConfigRoot config);
  internal static string ResolveEndpointCore(string endpoint, string defaultEndpoint, IDictionary<string,string> servers);
  internal static string NormalizeSyncWrite(string mode);
  internal static IEnumerable<string> CandidatePaths(string addInDirectory, string excelDirectory);
}
```

**`NatsConnectionManager`** (`NatsConnectionManager.cs`):

```csharp
public enum NatsPool { RtdData, RtdSub, Udf }

public sealed class NatsConnectionManager {
  public NatsConnection  GetConnection(string endpoint);            // cached, non-blocking
  public object          GetJetStream(string endpoint);             // INatsJSContext, cached
  public bool            IsJetStreamEnabled(string endpoint);       // cached, re-probed on reconnect
  public bool            TryGetServerInfo(string endpoint, out NatsServerSnapshot info);
  public int  LiveConnectionCount();
  public int  LiveRtdConnectionCount();
  public int  LiveUdfConnectionCount();
  public void SetEvictionProtection(Func<string,bool> isEndpointProtected);
  public void Shutdown();

  internal static bool WaitBounded(Task task, int timeoutMs);
  internal static T    WaitBoundedResult<T>(Task<T> task, int timeoutMs);
  internal static int  ResponseTimeoutMs(NatsPool pool);            // configured timeout + 500ms, clamped
}
```

`WaitBounded*` are the **only** bounded-blocking primitives in the codebase.
Shutdown fence: any call after `Shutdown()` throws
`InvalidOperationException("NatsConnectionManager is shutting down")`.

**`NatsSubscriptionManager`** (`NatsSubscriptionManager.cs`):

```csharp
public sealed class NatsMessage {
  public string Subject { get; }  public string Reply { get; }
  public string Data { get; }     public byte[] DataBytes { get; }
  public IReadOnlyDictionary<string,string[]> Headers { get; }
  public ulong Sequence { get; }  public DateTime? Timestamp { get; }
}

public sealed class NatsSubscriptionManager : IDisposable {
  public IDisposable Subscribe(string endpoint, string subject, string queueGroup,
                               Action<NatsMessage> onMessage, string origin = null);
  public int  SubscriptionCount { get; }
  public int  ListenerCount { get; }
  public int  ListenerCountWithOrigin(string origin);   // "RTD" / "UDF"
  public int  SubjectCountWithOrigin(string origin);
  public bool HasActiveSubscribers(string endpoint);
  public void Dispose();
  internal static string MakeKey(string endpoint, string subject, string queueGroup);
  internal static bool? SkipRepeatedOverrideForTests;
  internal static event Action<string,string,string> ListenerJoined;
}
```

**`NatsRuntime`**, **`NatsUdfAsync`**, **`UpdateCheck`**: as in the F plan
(`NatsRuntime.Connections/.Subscriptions/.Shutdown()/.ResetAfterAddInReload()`/
`ResolveEndpointForDispatch(object)`; `NatsUdfAsync.Run(...)/Enqueue/SafeMessage/
ResetAfterAddInReload`; `UpdateCheck.Start/EnsureFresh/IsUpdateAvailable/IsNewer/
NormalizeTag`).

---

## 2. Phases

### P0 — Scaffolding and the packing gate (F, with V scaffolding)

| ID | Task | Files | Depends | Acceptance |
| --- | --- | --- | --- | --- |
| F0.1 | Solution + add-in csproj: `net48`, `LangVersion 9.0`, Release `DebugType none`, `InformationalVersion dev`, `Compile Remove="test\**"` | `NATSExcel.sln`, `NATSExcel.csproj` | — | Release build zero errors/warnings |
| F0.2 | `IsExternalInit` shim | `IsExternalInit.cs` | F0.1 | `NatsOpts { Url = ... }` compiles |
| F0.3 | Packages: NATS.Client.Core/JetStream/KeyValueStore/ObjectStore (3.3.0), NLog 6.0.0-rc3, Newtonsoft 13.0.3, NJsonSchema 10.9.0, ExcelDna.* 1.9.0-beta2 | `NATSExcel.csproj` | F0.1 | Restore on the CI toolchain; output DLL set enumerated and recorded |
| F0.4 | `NATSExcel.dna` with **every** shipped DLL `Pack="true"` | `NATSExcel.dna` | F0.3 | `DnaPackagingParityTests` green both directions |
| F0.5 | `NLog.config` → `NATSExcel.log` + rolling | `NLog.config` | F0.1 | Log line on load |
| F0.6 | `.gitignore` | `.gitignore` | — | `git status` clean after Release build |
| F0.7 | Sample `NATSExcel.json` + `NATSExcel.schema.json` (embedded resource) | both | — | Sample validates; resource name fixed |
| F0.8 | `CopyNatsExcelDllToPublish` target | `NATSExcel.csproj` | F0.1 | `publish\` has dll + both XLLs |
| F0.9 | **P0 GATE**: temporary `ScaffoldProbe.cs` exercising NLog + Newtonsoft + NJsonSchema + NATS.Client | `ScaffoldProbe.cs` | F0.1–F0.8 | **Packed XLL loads in real (hidden) Excel and the probe UDF returns a value — no `FileNotFoundException`.** Deleted in F1.6 |
| V0.1 | Four test project skeletons + `ProductionSources.props` + parity test | `test\*\*.csproj`, `test\TestSupport\ProductionSources.props`, `test\NATSExcel.Tests\ProjectFileParityTests.cs` | F0.1 | All 5 build; parity green |

**Nothing proceeds until F0.9 passes.** This is where the ~20-DLL packing risk
is confirmed or refuted.

### P1 — Config (F)

| ID | Task | Files | Acceptance |
| --- | --- | --- | --- |
| F1.1 | `AppConfig` model + `Lazy` load + defaults | `Config.cs` | `AppConfigTests.Sanitize_*` |
| F1.2 | Candidate-path discovery (add-in folder first) | `Config.cs` | `CandidatePaths_*`, malformed-first stops |
| F1.3 | Tolerant parse (schema walk, per-value drop, invariant, case-merge) | `ConfigSchema.cs` | `ConfigSchemaTests.*` |
| F1.4 | `Sanitize` (rates/timeouts, conflation, SyncWrite, style) | `Config.cs` | `Sanitize_*` |
| F1.5 | **`MaskNatsEndpoint`**: userinfo, `password=`/`user=`, `token=`, **JWT (`eyJ...`)**, `SU...` seed, creds-file path, TLS cert path | `Config.cs` | `NatsMaskingTests` one case per form |
| F1.5b | **Auth model + precedence + `NatsAuthOpts` mapping** (D7): user/password, token, nkeySeed, jwt+seed, credsFile, URL userinfo, mTLS (`tls`, `tlsCert`, `tlsCertPassword`); invalid combinations → `#CONFIG` | `Config.cs` | `AuthConfigTests`: each type builds the expected `NatsAuthOpts`; illegal combos rejected |
| F1.5c | **Cluster/URL-list support** (D8, D10, D11): `Default`/`Servers`/`Clusters` model; comma-separated URLs; **discovery via `INatsServerInfo`** (`Cluster`, `ClientConnectUrls`, `ClusterDynamic`, `LameDuckMode`, `JetStreamAvailable`); reconnect knobs mapped to `NatsOpts` (`maxReconnectAttempts`→`MaxReconnectRetry`, `reconnectWaitMs`→`ReconnectWaitMin/Max`, `reconnectJitterMs`→`ReconnectJitter`, `connectTimeoutMs`→`ConnectTimeout`) | `Config.cs`, `NatsConnectionManager.cs` | `ClusterConfigTests`: one URL reaches a 3-node cluster and discovers peers; node loss → reconnect; `Cluster`/`ClientConnectUrls` surfaced |
| F1.5d | **Cluster reference resolution** (D10): name→`Clusters`, name→`Servers`, raw URL/list, then `Default.cluster`; exact-then-case-insensitive; dangling name → `#CONFIG`; name in both sections → validation warning; per-cluster `auth`/`jetstream`/`reconnect` merge over `Default` | `Config.cs` | `ClusterResolutionTests`: all four orders; ambiguity warning; no implicit localhost fallback |
| F1.5e | **Cluster identity** (§8.3): the `Clusters` key is a **local label** (never required to match the server); optional `expectCluster` asserts the announced `INatsServerInfo.Cluster` and reports a **non-fatal warning** on mismatch (surfaced in `=NatsCheckConfigFile()` and the log) | `Config.cs`, `NatsConnectionManager.cs` | `ClusterIdentityTests`: label-only is silent; matching `expectCluster` is silent; mismatch warns and still connects |
| F1.6 | Endpoint alias resolution + auth model; delete `ScaffoldProbe.cs` | `Config.cs` | `ResolveEndpointCore_*`; no implicit localhost |
| F1.7 | Culture-invariance for identifiers/values | `Config.cs` | `NatsValueLocaleTests` (de-DE) |

### P2 — Runtime core (F)

| ID | Task | Files | Acceptance |
| --- | --- | --- | --- |
| F2.1 | `TickGate` + `StripedLocks` | `TickGate.cs` | `TickGateTests` |
| F2.2 | `Conflation` (sanitize/resolve/push-immediately/is-due) | `Conflation.cs` | `ConflationTests` |
| F2.3 | `NatsMessage` shape + `NatsResultFormatter` scaffolding | `NatsConnectionManager.cs`, `NatsResultFormatter.cs` | message accessors tested |
| F2.4 | `NatsEndpoint` parse (URL list, auth merge, `Key` secret-bearing never logged, `ToMaskedString`) | `NatsConnectionManager.cs` | `NatsEndpointTests` |
| F2.5 | `GetConnection` (one per endpoint, `Lazy`, reconnect observation masked, shutdown fence) | `NatsConnectionManager.cs` | `ConnectionManagerTests` |
| F2.6 | `GetJetStream` cache + `IsJetStreamEnabled` gate | `NatsConnectionManager.cs` | cached; JS-off yields the exact error, never a hang |
| F2.7 | `WaitBounded*` + `ResponseTimeoutMs` | `NatsConnectionManager.cs` | `WaitBoundedTests` |
| F2.8 | Live counters + eviction/down-sweep (connected never evicted; subscribers protected) | `NatsConnectionManager.cs` | `Eviction_*` |
| F2.9 | `Shutdown()` idempotent, bounded | `NatsConnectionManager.cs` | `Shutdown_IsIdempotent` |
| F2.10 | Ref-counted `Subscribe` (one wire sub per `(endpoint, subject, queueGroup)`, origin tags) | `NatsSubscriptionManager.cs` | `SubscriptionManagerTests` |
| F2.11 | Per-key **epochs** (no ghost listeners) | `NatsSubscriptionManager.cs` | `SubscriptionEpochTests` |
| F2.12 | Consume loop + fan-out + dedup (ordinal, literal-only; wildcards never deduped) | `NatsSubscriptionManager.cs` | `SubscriptionDedupTests` |
| F2.13 | Reconnect resilience (auto-resubscribe or explicit on `Reconnected`) | `NatsSubscriptionManager.cs` | smoke restart test |
| F2.14 | `Dispose()` | `NatsSubscriptionManager.cs` | `Dispose_ReleasesEveryPhysicalSubscription` |
| F2.15 | `NatsRuntime` singleton + tombstones + reload reset | `NatsRuntime.cs` | `NatsRuntimeTests` |
| F2.16 | `NatsUdfAsync.Run` (sync passthrough / `Observe` + per-endpoint FIFO) | `NatsUdfAsync.cs` | `NatsUdfAsyncTests` |
| F2.17 | `NatsWriteObservable` (one-shot, continuation delivery, no-op dispose) | `NatsUdfAsync.cs` | `NatsUdfObserveTests` |
| F2.18 | FIFO chain + idle release + reload reset | `NatsUdfAsync.cs` | order/drain tests |
| F2.19 | `UpdateCheck` (`IsNewer`/`NormalizeTag`) | `UpdateCheck.cs` | `UpdateCheckTests` |

### P3 — Shared surface primitives (S)

| ID | Task | Files | Acceptance |
| --- | --- | --- | --- |
| S3.1 | `docs/PUBLIC-SURFACE.md`: catalog, grammar, markers, errors, counts | `docs/PUBLIC-SURFACE.md` | counts match §1.1 |
| S3.2 | `NatsRtdTopic.cs`: `NatsRtdKind`, `NatsRtdTopicSpec.Parse` — **no Excel/COM dependency** | `NatsRtdTopic.cs` | `RtdTopicGrammarTests` |
| S3.3 | `NatsUdfCommon.cs`: validation, `RequireSubject` (rejects wildcards on publish), endpoint resolve, `Fail`/`VerboseError`/`ErrorHint`, `ShortForLog`/`BoundContext`, F&F helpers, pair-range flatten, markers as consts | `NatsUdfCommon.cs` | `UdfArgumentValidationTests`, error strings pinned |
| S3.4 | `NatsResultFormatter.cs`: ordinal `SameMessage`, `DecodePayload`, 32767 bound, per-family JSON/matrix formatters (deterministic key order) | `NatsResultFormatter.cs` | `NatsResultFormatterTests` |
| S3.5 | `NatsJson.cs`: `MatrixToJSON`/`JSONToMatrix` (non-volatile, no twins) | `NatsJson.cs` | `NatsJsonTests` |
| S3.6 | Admin gate helper + JetStream-disabled guard helper | `NatsUdfCommon.cs` | `AdminGateTests` (11 gated), JS-off exact text |

### P4 — RTD server (S)

| ID | Task | Acceptance |
| --- | --- | --- |
| S4.1 | `TopicData` (lock, polled two-phase, subscribe lifecycle, disconnect state) | mirrors RedisExcel |
| S4.2 | `ConnectData` (parse, register polled/subscribed, reject malformed, JS probe, `(ConnectData)`, `#ERROR`) | all 8 kinds; malformed never registered |
| S4.3 | `SUB` (wildcards, queue group identity, RTD origin) | two cells share one physical sub |
| S4.4 | `KV` watch (no implicit bucket creation) | push per revision |
| S4.5 | `KVGET`/`KVKEYS` polled | `(no value)` |
| S4.6 | `REQ` with `everyMs >= 100` | no storm |
| S4.7 | `JS` (existing consumer only, explicit ack, dispose on disconnect) | missing consumer → exact error |
| S4.8 | `MSG` polled | filter `>` works |
| S4.9 | `OBJ` metadata watch (never the body) | push on change |
| S4.10 | `DisconnectData` idempotent + race-safe | — |
| S4.11 | `AddIn` `AutoOpen`/`AutoClose` (runtime reset, RTD reset, ThreadPool floor, non-fatal ComServer) | reload leaves no stranded registry |
| S4.12 | `NatsRtd` `ServerStart`/`ServerTerminate`/`Instances`/`ResetAfterAddInReload` | highest-ordinal survivor promoted |
| S4.13 | Timers via `CreateTimer` + `TickGate` (poll, Excel flush, 1s counter) | overrun skipped |
| S4.14 | Poll tick (group by endpoint, bounded parallelism, two-phase skip/commit) | slow endpoint isolated |
| S4.15 | Backoff 1s→2s→…→30s, capped per tick, `(subscribed)` on recovery | schedule unit-pinned |
| S4.16 | Delivery policy (`DeliverToTopic`, `SendToExcelIfDirty(conflationMs)`, dedup) | failed push stays dirty |
| S4.17 | `NatsRtdStatus` (RTD counters, default endpoint, rates) guarded | neutral after teardown |
| S4.18 | Masking audit of RTD log lines + `#ERROR` + status cells | redaction asserted |

### P5 — UDF families (S)

Core (`NatsUdfCore.cs`) — S5.1..S5.9: publish / publishJSON / publishIfChanged
(LRU + `ListenerJoined` marker clear) / publishWithHeaders / request /
requestJSON / subjectLatest(+JSON) / subjectUnsubscribe / subscriptionCount —
each write with its twin (except `SubjectUnsubscribe` never fire-and-forget).

JetStream + admin (`NatsUdfJetStream.cs`) — S5.10..S5.18: jsPublish(+JSON, ack →
sequence) / getMessage / getLastMessage / fetch (N×4) / streamInfo / streams /
consumerInfo / consumers / deleteMessage + purge (gated) / streamCreate/Update/
Delete (gated, immutable fields rejected) / consumerCreate/Delete/Pause (gated).

KV (`NatsUdfKv.cs`) — S5.19..S5.24: get (`(no value)`) / getMultiple (N×1|N×2) /
keys / buckets / status (JSON) / history (N×4) / put(+JSON) / create (rev 0) /
update (CAS) / delete + purge (gated).

Object (`NatsUdfObjectStore.cs`) — S5.25..S5.28: put (+description, digest) /
get (**1 MiB + binary bound**, resolves DESIGN open decision #3) / info / list /
watch (snapshot; live = RTD `OBJ`) / buckets / delete (gated).

### P6 — Status / config / JSON UDFs (S)

S6.1 `ConnectionCount`/`ServerInfo`/`Stats`/`Rtt` (guarded) · S6.2
`UpdateAvailable` (non-blocking) · S6.3 `CheckConfigFile`/`ConfigSchema`/
`ConfigPath`/`ConfigValidate` · S6.4 `ConfigReload` (non-volatile; does **not**
tear down connections; RTD rates apply at next RTD start).

### P7 — Surface freeze tests (S+V)

| ID | Task | Acceptance |
| --- | --- | --- |
| S7.1 | `PublicSurfaceParityTests` + `PublicSurface.baseline.txt` | any rename/reorder fails |
| S7.2 | `NonVolatileTwinTests` (25 pairs; counts 60/25/35/3) | delegation identical offline |
| V7.3 | `NoUnboundedBlockingTests`: source scan fails on `.GetAwaiter().GetResult()`, `.Result`, `.Wait(` outside the bounded helper | zero hits |
| S7.4 | Update `DESIGN.md` §5.1 and `AGENTS.md` layout for `NatsRtdTopic.cs`, `NatsUdfCommon.cs` | tables match tree |

### P8 — Test layers (V)

| ID | Layer | Files | Acceptance |
| --- | --- | --- | --- |
| V8.1 | Harness + probe + banner + AV helper + seams | `test\TestSupport\*`, `test\Start-NatsServer.ps1` | server starts/stops; banner one line |
| V8.2 | **Unit** (no NATS/Excel): config, schema, masking, json, conflation, subscription identity, connection offline, TickGate, WaitBounded, async bridge, observe doubles, twins, arg validation, JS vocabulary, admin gate, SyncWrite, parity tests, DNA parity, schema parity, AV banner | `test\NATSExcel.Tests\*` | all green |
| V8.3 | **Smoke** (real `nats-server -js`): connection identity, ref-count, wildcards, queue groups, request/reply, dedup + concurrent dedup, JetStream pub/consume/replay/purge/fetch, KV CRUD+CAS+history+watch, Object put/get/info/digest, SyncWrite modes, update check, JS-disabled branch, **auth matrix** (token / user+password / NKey seed / creds file / mTLS against a locally started server per type) | `test\SmokeTests\Program.cs` | every skip prints a reason; `--require-jetstream` fails on JS-off; each auth type connects |
| V8.4 | **Load**: `manager` vs `raw` (+ `js`) | `test\LoadTests\Program.cs` | throughput/alloc/GC + banner |
| V8.5 | **Liveness**: heartbeat/watchdog, bounds 5s/30s/20s, churn, starvation, **native restart storm**, backlogs, dedup re-join, `--quick/--skip-restart/--matrix/--soak` | `test\LivenessTests\Program.cs` | stall fails |
| V8.10 | **Cluster test profile** (D8): start a 3-node `TESTCLUSTER` from `test\TestSupport`, assert R3 survives a node kill and the client reconnects; assert R3 on a 2-node cluster yields the named `no suitable peers for placement` error | `test\TestSupport\cluster\*.conf`, `test\LivenessTests\Program.cs` | R3 stream readable from a survivor; named error surfaced |
| V8.6 | **Excel E2E**: scaffold + COM retry + AV banner; UDF sheet; RTD sheet; multi-workbook; connection blip; write modes + async pending marker + NonVolatile-once; workbook sanitize; params | `test\Run-ExcelE2E.ps1` | ALL PASS all 4 modes |
| V8.8 | **Read-only live-feed run** (D6): point the UDF/RTD read paths at an external endpoint to observe real traffic, never writing. A `-ReadOnly` switch (default ON for any non-local endpoint) (a) skips every write/admin/destructive section, (b) refuses publish-like calls at the harness level, and (c) writes nothing to a local server. Accepts a subject/pattern filter so a `>` firehose can be scoped down before `SubjectLatest`/RTD. | `test\Run-ExcelE2E.ps1`, `test\SmokeTests\Program.cs` (`--read-only`) | feed observed end-to-end; zero writes; endpoint only in argv |
| V8.9 | **External-endpoint bookkeeping**: the run prints the endpoint **masked**, keeps the workbook in `%TEMP%` (never `test\NATSExcel.Test.xlsx`), and the committed sample keeps `nats://localhost:4222` + `test.natsexcel.*` only. A pre-commit/pre-release grep (golden rule 9) fails on any company name/domain. | `test\Run-ExcelE2E.ps1`, `.github\workflows\build.yml` | `git grep` returns nothing; no private host / company name in tracked files |
| V8.7 | Matrix + compare + AV helpers | `test\Run-ExcelE2E-Matrix.ps1`, `test\Compare-ExcelE2E.ps1`, `test\Set-Antivirus.ps1` | side-by-side table |

### P9 — Installer (V, native-first)

| ID | Task | Acceptance |
| --- | --- | --- |
| V9.1 | `install-server.ps1`: `native` (download zip or reuse PATH) / `docker` / `none`; writes `nats.conf` with **durable** `store_dir`; exit 3 = unavailable, never fails install | native server answers 4222 |
| V9.2 | `nats-service.ps1`: `sc.exe create NATS binPath= "... -c ..." start= auto` (the only admin step), separate from the per-user install | elevated install/start/remove |
| V9.3 | `detect-servers.ps1` / `uninstall-server.ps1` (opt-in removal, default keep) | only ticked items removed |
| V9.4 | `test-connection.ps1` (real NATS protocol `INFO`/`PING`→`PONG`, optional token/user/pass) | works on JS and non-JS |
| V9.5 | `verify-connection.ps1` (server info + publish/latest round trip, optional `-OpenExcel`) | sample publishes |
| V9.6 | `validate-config.ps1` + `Validate-Config.cmd` (same embedded schema) | OK/FAIL report |
| V9.7 | `NATSExcel.iss`: Inno Setup 6, per-user, `C:\NATSExcel`, Office-bitness XLL, en/es/pt-BR, server page native-first, tasks, end-of-install test | `iscc` compiles every run |
| V9.8 | Silent mode: `/VERYSILENT /DIR /URL /USER /PASSWORD /TOKEN /SERVER /MODE /OPENEXCEL` | files-only install |
| V9.9 | `installer\README.md` | matches the wizard |

### P10 — CI, release, docs (V)

| ID | Task | Acceptance |
| --- | --- | --- |
| V10.1 | `build.yml` build job: checkout, nuget restore, msbuild (dev), unit tests | green |
| V10.2 | nats-server step: pinned zip + SHA256, `-js --store_dir`, 4222 probe, `--jetstream` probe | probe prints enabled |
| V10.3 | Smoke (retry 2×, 3s) + liveness quick (retry 3×, 5s) | real regression fails all attempts |
| V10.4 | Tag guards: CHANGELOG heading anchored match | red on mismatch |
| V10.5 | Release build with tag `InformationalVersion`; Inno Setup on every run; asset-existence assertion; `release` job (checkout first, CHANGELOG body) | release published |
| V10.6 | `Readme.md` (installer + manual + beginner native-Windows guide: winget / release zip / `sc.exe` service / optional Docker) | accurate |
| V10.7 | `CHANGELOG.md`, `test\README.md` | match CI |

---

## 3. Dependency graph and parallel tracks

```
F0.1 ─┬─ F0.2 ─ F0.3 ─ F0.4 ─┐
      ├─ F0.5 ─ F0.6 ─ F0.7 ─┤
      └─ F0.8 ───────────────┴─ F0.9  == P0 GATE (XLL loads in Excel admission)
V0.1 ────────────────────────┘

After the gate:
  Track F:  P1 ─ P2 ─────────────────────────────┐
  Track S:  P3 ─┬─ P4 ─┬─ P5 ─┐                  │
                └─ P7 ◄────────┴─ P6 ─┘          │
  Track V:  V8.1 ─ V8.2 (needs P1/P2 seams)      │
                  V8.3 (needs P4/P5)             │
                  V8.4/V8.5 (needs P2)           │
                  V8.6 (needs P4/P5 + built XLL) │
                  P9 (needs F0.4 assets) ─ P10 ◄─┘
```

- **Track F** is on the critical path for everything; F2.7 (`WaitBounded`) and
  F2.15 (`NatsRuntime`) unblock all of Track S.
- **Track S** P3 can start once F0.3/F1.5/F1.6 land (formatters and grammar are
  Excel-free and unit-testable immediately).
- **Track V** V8.1 (harness) and V8.2 (unit) start as soon as F1/F2 seams exist;
  they test Track S incrementally.

---

## 4. Hand-off points

| # | Frozen by | Consumed by | Payload |
| --- | --- | --- | --- |
| H1 | F0.3/F0.4 | S, V | Package versions, `.dna` Reference list, `publish\` asset names |
| H2 | F1.5/F1.6 | S, V | `AppConfig` surface + `MaskNatsEndpoint` |
| H3 | F2.5/F2.6/F2.7 | S, V | `GetConnection`/`GetJetStream`/`IsJetStreamEnabled`/`WaitBounded*` |
| H4 | F2.10 | S, V | `Subscribe` token + `NatsMessage` |
| H5 | F2.15 | S | `NatsRuntime.ResolveEndpointForDispatch` |
| H6 | F2.16 | S | `NatsUdfAsync.Run`/`Enqueue`/`SafeMessage` |
| H7 | F2.2/F2.3 | S, V | `Conflation`, `NatsResultFormatter` |
| H8 | F0.7 | S, V | Embedded `NATSExcel.schema.json` resource name |
| H9 | S3.1–S3.6 | V | `PUBLIC-SURFACE.md`, signatures, grammar, markers, errors, bounds, 11 gated names |
| H10 | V8.1 | S, V | `NativeServerHarness`, `JetStreamProbe`, `RunBanner`, `ProductionSources.props` |

---

## 5. Consolidated risk register (highest first)

| # | Risk | Impact | Mitigation |
| --- | --- | --- | --- |
| R1 | `.dna` pack list incomplete → `FileNotFoundException` inside the XLL | Add-in looks broken only in Excel | **F0.9 is a hard gate**; enumerate from actual output; `DnaPackagingParityTests` fails the build on drift |
| R2 | CI toolchain divergence for the net48 restore (msbuild/VS vs dotnet SDK — `System.Text.Json`, `Microsoft.VisualStudio.Threading`) | Ships a broken XLL while tests pass | Standardize `nuget restore`+`msbuild` for the shipped build; pin versions; verify the `.dna` list on the CI toolchain at P0 |
| R3 | Blocking the Excel thread on the async-only client | Frozen Excel | `WaitBounded` is the only bounded primitive; writes via `Observe`; RTD reads on the timer; `NoUnboundedBlockingTests` tripwire; liveness bounds |
| R4 | Large surface → Compile-list drift / expensive matrix | Untested production file | Central `ProductionSources.props` + extended parity test asserting coverage and the four imports |
| R5 | Ghost/duplicate listeners | Blank or double-pushed cells | Per-key epochs + install re-check under lock; `SubscriptionEpochTests` |
| R6 | JetStream disabled → hang or silent empty | Unattributable failures | Cached probe per endpoint, re-probed on reconnect; exact error text; `--require-jetstream` in CI |
| R7 | Credential leak (`SU...` seed, creds path, userinfo, `token=`) | Security incident | `MaskNatsEndpoint` on every log line/cell/`#ERROR`/status; `NatsMaskingTests` per form |
| R8 | Reconnect not auto-resubscribing | Permanently deaf cells | Confirm NATS.Net auto-resubscribe at P0; explicit re-subscribe on `Reconnected` otherwise; smoke restart test |
| R9 | JetStream consumer lifecycle in an RTD cell (work-queue split) | Lost messages | RTD `JS` binds only to an existing consumer and acks; documented one-cell-per-consumer; `JSFetch` documents cursor movement |
| R10 | Binary/large objects | OOM / oversized cell | `ObjectGet` rejects > 1 MiB or binary with a named error; 32767 cell bound |
| R11 | Conflation starvation | Stale-looking cell | Single source of truth in `Conflation.Resolve`; failed push stays dirty; default window 200 ms |
| R12 | `REQ` topic floods the server | Self-inflicted load | `everyMs >= 100` clamp + `TickGate` + per-topic last-request clock |
| R13 | RTD positional ambiguity | Wrong subscription | Frozen optional-middle rule + grammar tests + documented workaround |
| R14 | Excel E2E flakiness (COM + AV/EDR) | Red run misread as a defect | Keep warmup/retry/AV banner/COM-retry summary; document that a retry total is normal |
| R15 | Installer admin/service step vs per-user install | Install fails or silently skips | Isolate the service in `nats-service.ps1` with an elevation prompt; add-in install never depends on it |
| R16 | JetStream `store_dir` temp default | Streams lost on restart | Installer writes a durable `store_dir`; CI uses an explicit temp; README calls it out |
| R17 | Dual JSON stack bloat | Larger XLL / more packing risk | Newtonsoft-backed serializer registry; keep the closure minimal and pinned |
| R18 | `UpdateCheck` repo slug / `SyncWrite`/dedup marker semantics diverge | No update alerts; stale cells | Hermetic test seams; ordinal compare, literal-only dedup |

---

## 6. Definition of done

**Per workstream**

- **F**: `publish\` has both packed XLLs + the dll; the P0 gate passed; `.dna`
  parity green; config loads once with tolerant per-value drop and never
  silently falls back to `localhost`; masking covers every credential form; one
  connection per endpoint with a cached JetStream context and an idempotent
  bounded shutdown; subscriptions ref-count correctly with no ghost listeners;
  `NatsRuntime` has a shutdown tombstone and reload reset; the async bridge is a
  pure passthrough when off and an `Observe` FIFO when on; unit suite green with
  no NATS and no Excel.
- **S**: all 88 functions with frozen signatures; all 25 twins pass parity; all 8
  RTD commands behave per §1.2 with malformed topics rejected before
  registration; every marker/error string matches literally; the 11 gated
  functions return the exact gate error and are the only gated ones; zero
  unbounded blocking; masking asserted on UDF/RTD/status paths.
- **V**: five layers green (unit/smoke/load/liveness/E2E) with attributable
  banners; installer compiles every run, native-first, per-user, with the
  service step isolated; CI mirrors the RedisExcel retry policy; docs current.

**`v1.0.0` acceptance gate**

1. `CHANGELOG.md` top heading is exactly `## v1.0.0` (no `(unreleased)`).
2. Unit + smoke + liveness(quick) pass in CI on the tagged commit; the
   `--jetstream` probe reported enabled.
3. All six release assets exist and are attached: `NATSExcel-packed.xll`,
   `NATSExcel64-packed.xll`, `NATSExcel.dll`, `NLog.config`, `NATSExcel.json`,
   `NATSExcelSetup-1.0.0.exe`.
4. `.dna` parity green on the CI toolchain (no unpacked dependency).
5. The release body is the `## v1.0.0` CHANGELOG section.
6. The public surface is documented as immutable (UDF names/argument order, RTD
   commands, ProgID `NatsRtd`, `NATSExcel.json` shape); future work is additive.
7. `Readme.md` beginner guide, `test\README.md` and `installer\README.md` are
   current and contain no private data.

---

## 8. Verified platform tests (executed)

These were run for real on a Windows machine to de-risk the install and test
story before implementation. All used the `nats-server` v2.15.1 Windows binary
and the `nats` CLI v0.5.0; nothing from these runs is committed.

| Scenario | Result |
| --- | --- |
| **Native Windows server**, `-js --store_dir <dir>` | Started, port open, `INFO`/`PING`→`PONG` |
| Core pub/sub (native) | `sub` received the published payload |
| JetStream (native) | Stream created (file storage), message stored, `Messages: 1`; KV bucket created, `put`/`get` returned the value |
| **Docker server**, `docker run -p 4224:4222 nats:latest -js --store_dir /data` | Container up, Core pub/sub round-tripped, stream created |
| Docker image version | `nats-server v2.15.0` (the `nats:latest` image) — slightly behind the native zip |
| **JetStream disabled** (native, started without `-js`) | Core pub/sub still works; JetStream call returns `code=503 err_code=10076 description=jetstream not enabled` — the exact signal the add-in must map to `JetStream is not enabled on this server` |
| **3-node cluster** (`TESTCLUSTER`, R3 stream) | R3 created; leader elected (`n2`); replicas reported `current: true` |
| **Cluster node loss** | Killed the node on the first port; the R3 stream kept its messages (`messages: 3`) and was still readable from a surviving node; the dead node showed `current: false` with `lag` |
| **R3 on a 2-node cluster** | Rejected: `no suitable peers for placement (10005)` — a named error the cell must show |

Practical conclusions folded into `DESIGN.md` §8.1: the client needs only a URL
list (it auto-discovers peers), and JetStream durability is the stream admin's
`num_replicas` choice. The add-in must surface placement failures by name.

## 9. Open items to close before/at P0

1. **License**: NATS.Net is Apache-2.0; pick a license for NATSExcel before the
   first public release (RedisExcel ships none).
2. **`UpdateCheck` repo slug** (`rspadim/NATSExcel`) — confirm.
3. **`AllowAdminWrites` default `false`** — confirm (destructive functions argue
   for `false`).
4. **Object cell bound** `1 MiB` — confirm (S7/§1.3).
5. **Live-feed test endpoint**: a remote read-only server is used to validate the
   read paths against real traffic (D6/V8.8). The endpoint stays **out of the
   repository** — command-line argument only — and every write is disabled for a
   non-local endpoint. Two things to confirm with the feed owner: whether it
   needs **auth** (user/password/token/creds), and whether a **subject filter**
   is expected (a `>` firehose should be scoped down before `SubjectLatest`/RTD).
   Do not commit the endpoint, its domain or any credential.
