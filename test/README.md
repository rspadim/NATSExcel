# Test layers

Five layers, mirroring the sibling RedisExcel project. `NatsRtd.cs` is the only
production file that cannot be compiled outside Excel (it needs the RTD/COM
surface), so it is covered by the Excel E2E only.

`test\TestSupport\Run-Tests.ps1` runs the layers that need no Excel **in parallel**
(each on its own port and work dir) and prints a summary table:

```powershell
powershell -ExecutionPolicy Bypass -File test\TestSupport\Run-Tests.ps1                  # unit + smoke + liveness
powershell -ExecutionPolicy Bypass -File test\TestSupport\Run-Tests.ps1 -WithE2E          # + the Excel end-to-end
powershell -ExecutionPolicy Bypass -File test\TestSupport\Run-Tests.ps1 -WithLoad         # + the load benchmark
powershell -ExecutionPolicy Bypass -File test\TestSupport\Run-Tests.ps1 -LivenessFull     # full, DESTRUCTIVE liveness
```

`-NoBuild` reuses the last build, `-Sequential` runs one layer at a time (least
machine pressure), `-KeepCluster` leaves the servers up, and `-ServerExe` (or
`NATSEXCEL_NATS_SERVER`) points at `nats-server.exe`.

## 0. Config + connection validator (no Excel)

`tools/NatsValidate` validates a `NATSExcel.json` and, with `--test`, connects
with the **same NATS.Net client the add-in uses**, so every credential style is
exercised (token, user/password, NKey, JWT + seed, creds file, mTLS). The
PowerShell helpers only speak the wire protocol, so they cover token and
user/password only.

```powershell
dotnet run --project tools\NatsValidate -c Release -- NATSExcel.json
dotnet run --project tools\NatsValidate -c Release -- NATSExcel.json --test
```

Exit codes: `0` OK, `1` config problems, `2` unreadable, `3` dangling
cluster/server reference, `4` connection failed.

## 1. Unit tests (no NATS server, no Excel)

```powershell
dotnet test test\NATSExcel.Tests\NATSExcel.Tests.csproj -c Release
```

The project compiles the production sources directly (linked `Compile` items
from `test\TestSupport\ProductionSources.props`) instead of a
`ProjectReference`, so the Excel-DNA packing step does not run on every test.
A parity test fails when a root production file is missing from that list.

Coverage is broad: 24+ test files / 360+ tests cover the config constants, the
shipped sample shape, the JSON-schema per-value drop, `StripComments`, the
masking renderer, conflation, the JSON ⇄ matrix conversions, the RTD grammar,
the UDF argument vocabulary and markers, the range helpers, the write modes, the
connection-wait bounds, the subscription identity, and the DNA/public-surface
parity, the culture invariance of every value written to the wire
(`NatsValueLocaleTests`, on a real de-DE thread), the async bridge
(`NatsWriteObservableTests`) and the harness safety contracts
(`HarnessInvariantTests`). Source scans pin what cannot be behavioural
(`RtdInvariantTests`, `NoUnboundedBlockingTests`, `SubjectEpochLeakTests`).
Growth areas: the RTD kinds under `-AsyncWrites`, the JetStream
cluster-replica paths, and a live-feed read-only run (V8.8).

## 2. Cluster and server harness (no Excel)

`test\TestSupport\Start-Cluster.ps1` brings up a disposable NATS cluster on
localhost (1 or 3 nodes, JetStream on, a durable store per node) and leaves it
running for the other layers. `test\TestSupport\Stop-Cluster.ps1` tears it down.
Both work with a locally installed `nats-server.exe` (no Docker required) and
download it from the official release if it is missing.

**No private endpoint is ever stored here**: hosts are `127.0.0.1` and the
ports are the defaults. A real server is passed as a command-line argument and
never committed.

Each node also gets a monitoring port (`http_port: <client port>+4000`) when it is
free, which is what makes the smoke suite's wire-level refcount proof real.

## 3. Smoke tests (requires the cluster, no Excel)

```powershell
dotnet run --project test\SmokeTests -c Release -- "nats://127.0.0.1:4222" --require-jetstream
```

Of the 97 frozen functions, the smoke layer drives the ones that need no Excel:
the ref-counted Core subscriptions (two listeners on one subject share exactly ONE
physical subscription - proven on the wire through `/connz` - and the subscription
is dropped when the last listener goes), wildcards, queue groups, request/reply
(the answer, `no responders` and the timeout), consecutive-duplicate suppression,
the JetStream publish/ack/replay/purge paths and the named missing-stream error,
Key/Value CRUD/CAS/tombstone, Object Store put/get/digest plus the named
size/binary guards, the frozen admin-gate text for all eleven gated functions, the
JetStream-disabled server, the `jetstream.domain` routing, a disposable server per
auth style (no-auth, token, user/password, NKey - including the `nkeySeed`
text-or-`.nk`-file contract), the dangling-reference `#CONFIG` and gate-ordering
rules, the `NatsUDFSubjectUnsubscribe` lifecycle, and a **catalog sweep** that calls
every remaining frozen function once and requires a usable cell from each.

It starts its own disposable servers on free ports for the pieces that need a
different server shape, and reports `0 skipped` on a full local run. `--read-only`
skips every write section; `NATSEXCEL_NATS_SERVER` (a path) tells it where
`nats-server.exe` is, and a check that cannot run prints its reason and SKIPs.

## 4. Excel end-to-end (requires Excel + the cluster)

```powershell
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1
```

Loads the packed XLL into a hidden Excel, and covers what a user actually does:

- UDF write + read round trip (Core NATS).
- RTD `SUB` push: a published value reaches the cell, and a value longer than the
  RTD cell limit is truncated with a visible marker at 255 characters.
- The other RTD kinds: `MSG`, `JS`, `KVGET`, `KVKEYS`, `KV` (watch) and `OBJ`
  (metadata watch), plus the frozen unknown-command and `RTDCLUSTER` vocabulary
  errors, the unresolvable-reference error, the `RTDCLUSTER cluster/nodes` cells
  and the six `NatsRTD*` status helpers.
- JetStream and KV paths (each must return a named error or a value, never a hang).
- **Cluster failover**: kill the node the add-in is connected to and confirm the
  cell recovers through the surviving nodes.
- **Multi-workbook**: open a copy, close it, confirm the original keeps receiving.
- **Excel restart**: close and reopen Excel, confirm a live cell resumes.
- Write modes and the `...NonVolatile` single-run rule.

Useful parameters: `-RepoRoot`, `-NatsUrl` (default `nats://127.0.0.1:4222`),
`-SubjectPrefix`, `-SkipNodeKill` (auto-skipped for a non-local host),
`-KeepExcelOpen`, `-ReadOnly`, `-AllowRemoteWrites`.

> **Start the cluster with `Start-Cluster.ps1` first.** The failover step kills the
> node answering `-NatsUrl`, and it only does that when it can PROVE the listener
> is one this repository started (a `nats-server` whose command line names a
> `nats.conf` under `%TEMP%\natsexcel-cluster`). A server somebody else started -
> your own long-running instance, for example - is never killed: the journey
> reports `SKIP` with the reason, and the run prints a warning at the top saying
> the endpoint was not started by the harness. The same proof gates the pre-run
> cleanup, so a busy port is left alone instead of being taken over.

A crashed run can leave its hidden Excel alive, holding the copied XLL; the next
run then fails to overwrite it. `New-AddInFolder.ps1` clears exactly that case -
it stops a stale instance only when the process is invisible AND has the
destination XLL loaded, and merely reports a VISIBLE Excel (your own session) -
so a fresh run needs no manual cleanup.

Excel COM automation is inherently flaky (antivirus/EDR widens the window): the
script retries transient rejections and prints a `COM retries` summary. A retry
count is normal; a real defect fails every attempt. See `AGENTS.md` for the
hard-won lessons.

## 5. Load test (requires the cluster)

```powershell
dotnet run --project test\LoadTests -c Release -- "nats://127.0.0.1:4222" 10 2 1
```

Parameters: url, seconds, publishers, listeners, pattern. Reports throughput and
allocated bytes per message. Read-only stress runs against a real server are
allowed, but pass the url only as a command-line argument (never commit it) and
keep the runs short.

## 6. Liveness / fault test (requires the cluster; Docker optional)

```powershell
dotnet run --project test\LivenessTests -c Release -- "nats://127.0.0.1:4222"
```

Drives the real managers under adversarial conditions — subscribe/dispose churn,
ThreadPool starvation, **node kill and restart storms**, async queue bursts — and
enforces liveness bounds with a heartbeat watchdog: a stall past **5 s** fails
the run, delivery must resume within **30 s**.

> Only run the destructive layers against a disposable/local cluster: the node
> kill and the restart are destructive.
