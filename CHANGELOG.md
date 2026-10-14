# Changelog

All notable changes to this project are documented in this file. The format follows the spirit of
[Keep a Changelog](https://keepachangelog.com/), and the project uses annotated `v*` tags.

## v0.1.3

Fixes found by a read-only review of the PUBLISHED `v0.1.2` artifacts and of the
release batch itself - two of them only visible in the shipped bytes. **No change to
the public Excel surface** (the same 91 `NatsUDF*` functions and 6 `NatsRTD*`
helpers) and **no change to the `NATSExcel.json` shape**, so upgrading is a drop-in
replacement.

### Fixed

- **The installer's `FileVersion` was empty**: `installer\NATSExcel.iss` set
  `AppVersion` but no `VersionInfoVersion`, so Explorer's Details tab showed a blank
  "File version" and inventory/AV tooling that keys on it read `0.0.0.0` (the
  ProductVersion was correct). The version-info directives are now set - through a
  dedicated NUMERIC define, because Inno rejects `VersionInfoVersion`/`ProductVersion`
  values like `dev` or a `-rc1` suffix, which is exactly what a branch build passes -
  and the CI hands it the tag's number or `0.0.0.0`. Verified by compiling all three
  cases with ISCC: a tag build reports `0.1.2`, a branch build `0.0.0.0`, and no
  defines at all still compiles. The published `v0.1.2` asset still carries the blank
  field; the next release ships the fix.
- **`test\TestSupport\Run-Tests.ps1 -WithE2E` silently SKIPPED the failover
  journey**: the runner starts each layer's cluster under its own temp root
  (`%TEMP%\natsexcel-tests`), which the E2E's ownership proof did not recognize, so
  it reported "not a disposable test node" and skipped. The proof accepts that root
  now, so the documented `-WithE2E` path really runs the journey.
- **`"maxAttempts": 0` does not mean "no retry"**: NATS.Client.Core 3.3.0 bounds the
  retries only when the value is positive, so `0` and `-1` both leave its unlimited
  default (its own documentation says "-1, unlimited"). The code comment, the schema
  descriptions, a test name and the validator's comment claimed otherwise;
  `tools\NatsValidate` now uses `MaxReconnectRetry = 1` so a validator really fails
  fast, and the schema says what `0` does.
- **The "10 s ceiling" was wrong**: `NatsOpts.Default.ReconnectWaitMax` is 5 s in
  NATS.Client.Core 3.3.0, so the pre-fix wait sequence was 2, 4, 5, 5 s - corrected
  in the code comment, `AGENTS.md` and this file, and the schema's `waitMs` now
  documents what `0` does (the client's own schedule).
- Tests only: the `RTDCLUSTER cluster/nodes` predicates accepted Excel's own `#N/D`
  and `#VALUE!` markers as "the cell answers", and the catalog sweep's
  `NatsUDFSubjectLatestJSON` entry accepted the `(null)` marker. The predicates now
  reject every `#`-prefixed marker, and the sweep warms that read and asserts a real
  value (`1x1 value`).

- **The log file was written to the OFFICE folder, not next to the add-in** -
  while `Readme.md`, the three site pages, `installer\NATSExcel.iss` (11
  translations) and the uninstaller all promised "next to the add-in", and
  `=NatsUDFLogPath()` reported that same folder. NLog resolves a relative
  `fileName` against `AppDomain.BaseDirectory`, which in Excel is the Office
  install directory. `AddIn.AutoOpen` now points the file target at the folder
  `=NatsUDFLogPath()` reports, so all four agree; an unwritable add-in folder keeps
  NLog's default instead of costing the user the log (`LogPathTests` pins the
  wiring, the fallback and the probe-file cleanup).
- **`Validate-Config.cmd` was broken as installed**: the installer puts it in
  `{app}` but every `.ps1` in `{app}\scripts`, and the `.cmd` called
  `%~dp0validate-config.ps1` - so the documented double-click in `C:\NATSExcel\`
  failed with "file not found". It now calls `%~dp0scripts\validate-config.ps1`.
- **Dead links on all three published pages**: the "Architecture and design"
  (`DESIGN.md`) and "Development plan" (`PLAN.md`) bullets pointed at files that
  are not in the repository (they are maintainer notes kept on disk and
  gitignored), so every click 404'd. The bullets are replaced by a
  `CHANGELOG.md` link, and `docs\PUBLIC-SURFACE.md` now names the maintainer notes
  instead of citing them as files a reader could open.
- **`installer\verify-connection.ps1` showed Portuguese message boxes** (and a
  Portuguese sample sheet: "NATSExcel instalado com sucesso!", `"FALHOU"`,
  "Servidor:") to every user, whatever the wizard language - against golden rule 1
  and the "the Excel-facing surface stays English" promise. All user-visible text is
  English now.
- **Docs that said things the code does not do**: `MessageCounterThreshold` was
  documented as "Automatic switches to timer delivery" but is never read (the
  schema and the sample now mark it reserved and say `Automatic` behaves as
  `Realtime`); `ConflationMs` "default 200 ms" is the sample's value (with the key
  absent the window is the Excel tick); `#N/D` is the localized rendering of
  `#N/A` (the docs quote both now); the stale "unit (366)", "asserts 53 times",
  "11/11", "Status: v0.1.0", the "~204k msg/s" figure, the growth-area lists that
  have since landed (async bridge, RTD kinds, `expectCluster`, V8.8's `-ReadOnly`),
  the installer README's "token only" wizard description, and the RedisExcel
  sibling prose ("built on NATS" / "sent to Redis/NATS") are corrected.

### Added

- `UnpinnedClaimsTests`: the round-5 audit listed five claims that were TRUE but had
  nothing keeping them true, and each would have silently become a lie. They are
  pinned now - the config search order (the profile before `C:\Windows`, every
  candidate absolute), the NATS client defaults the docs quote (2 s initial wait, 5 s
  ceiling, 100 ms jitter, unlimited retries) plus the "one configured wait is applied
  as both bounds" mapping, the installer's 11 languages all carrying the same 63-key
  set (the manual `languages\messages.ps1` check, automated), the "the releases are
  not code-signed yet" statement (adding `signtool`/`SignPath`/`jsign` to the
  pipeline fails the test until `SIGNING.md` and the three pages are updated), and
  the E2E's assertion count in `AGENTS.md` - the number that had already drifted
  twice, and was wrong again when the test was written (53 call sites, not 59).
## v0.1.2

Four rounds of audit fixes plus the test coverage that pins them. **No change to the
public Excel surface** (the same 91 `NatsUDF*` functions and 6 `NatsRTD*` helpers)
and **no change to the `NATSExcel.json` shape**, so upgrading is a drop-in
replacement of the XLL/installer.

### Fixed

- **A cluster that declares more than one URL never mapped back to its cluster**, so
  every per-cluster setting was silently dropped for exactly the clusters that need
  them most: `auth`, `timeout`, `reconnect`, `jetstream.domain` and the
  `expectCluster` check. The connection manager receives a cluster's resolved
  endpoint, which is its comma-joined URL list, and the lookup compared that whole
  string against a single URL - so it matched nothing and the shared block won
  (wrong credentials, default JetStream subjects). The lookup is now the exact
  inverse of the projection that produces it; a single URL keeps its previous
  behaviour, including the "two clusters share a URL, so do not guess" rule.
- **`"tlsCert": ""` turned a shared `"tls": true` off** for that cluster, dropping the
  connection to opportunistic/plaintext. A certificate implies TLS only when it
  carries one; a present-but-empty token is now no declaration at all.
- **The RTD cluster cell did not mask what it displayed.** Every `RTDCLUSTER` value is
  server-supplied (the cluster name, the advertised connect URLs - which embed a
  userinfo when the server was configured with one - and the raw
  `$SYS.REQ.SERVER.PING` reply), and all of it bypassed the masking the UDF side
  applies to the same values. It is masked now, before the dedup comparison, so the
  deduped value is what was delivered.
- **A config reload could strand a live timer pair.** `ApplyConfig` re-checked the
  teardown tombstone only before swapping the RTD timers, and `DisposeTimer` nulls
  the field before disposing - so a `Teardown` starting in that window disposed
  nothing and the new pair kept running on an RTD server that was no longer
  registered, unreachable by any later teardown and rooting the instance. The
  tombstone is re-read after the pair is started and the pair is unwound.
- **A losing `NatsUDFSubjectLatest` registrant could delete the winner's value.** Its
  verdict was re-sampled after the install lock was released, so a removal plus a
  fresh registrant in between made it delete a LIVE listener's cached value and the
  cell fell back to `(null)` until the next message. The verdict is now decided
  under the install lock, and the cached value is dropped only when nothing is
  registered for that key.
- **`expectCluster` is now really checked.** `=NatsUDFConfigValidate()` said the
  assertion was "checked against the live server when a formula connects", but only
  `tools\NatsValidate` ever compared it. The connection manager logs a bounded,
  once-per-endpoint warning naming the expected and the announced cluster; a
  standalone server announcing no cluster is the most likely mismatch and is
  reported as such.
- **The RTD `OBJ` cell could never show an object's metadata.** Its watch subject
  is `$O.<bucket>.M.<base64url(object name)>`, and the add-in stripped the base64
  `=` padding while every writer keeps it (NATS.Net's own object context and the
  `nats` CLI both publish `ZW1wdHktMThjMzUyMWI=` for a 14-character name). The cell
  therefore subscribed to a subject nothing is ever published to and sat at
  `(ConnectData)` forever - and it looked correct only for names whose length is a
  multiple of 3, where base64 has no padding at all. Found by the new end-to-end
  OBJ cell; pinned by a source scan and by that cell.
- **The smoke suite's wire-level refcount proof could never run**: it parsed
  `/connz`'s `subscriptions_list` as a JSON object, while the server sends an ARRAY
  of subject names, so the probe always answered "no monitoring port" and the check
  reported `SKIP` even against a server that exposed `/connz`. With the parse fixed
  (and the harness's monitoring port above) the suite now proves at the wire that
  two listeners share exactly ONE physical subscription, and that the subscription
  is dropped when the last listener goes - zero skips in the smoke layer.

### Changed

- **The Excel E2E never kills a server it cannot prove it started.** The failover
  journey stops the node answering the endpoint, and it used to take whatever
  process held the port - a developer's own long-running `nats-server` would have
  been killed (a denied kill aborted the run with a raw "Access denied" instead).
  Ownership is now proven from the command line naming a `nats.conf` under the
  harness's temp work root; without that proof the journey reports `SKIP` with the
  reason, a denied kill no longer aborts the run, and a run using somebody else's
  server says so at the top. The same proof gates the pre-run cleanup, and the
  pin is `HarnessInvariantTests`.
- **`New-AddInFolder.ps1` clears the stale hidden Excel a crashed run leaves
  behind** (the instance that keeps the copied XLL loaded, which made every later
  run fail on "being used by another process"). It stops such an instance only when
  it is invisible AND has the destination XLL loaded, and merely reports a VISIBLE
  Excel - the user's own session. The E2E's multi-workbook step also retries
  `Workbooks.Open` (which can answer with `$null` instead of throwing) and quits its
  hidden instance on the failure path.
- **`reconnect.waitMs` is now the whole wait**, not just its floor: the value is
  applied as both `ReconnectWaitMin` and `ReconnectWaitMax`, so `"waitMs": 2000`
  waits 2s per attempt instead of doubling toward the client's own 5s ceiling
  (`jitterMs` still spreads it). The schema says so.
- **The test harness exposes a monitoring port per node** (`http_port: <client
  port>+4000`), so the smoke suite's wire-level refcount proof really runs; it is
  added only when that port is free. `test\TestSupport\Run-Tests.ps1` also gained
  `-WithLoad` (the load benchmark is opt-in locally, always run in CI) and the
  E2E now exercises the RTD kinds beyond `SUB`.

### Added

- Tests: the multi-URL cluster lookups, the blank/null `tlsCert` cases, the RTD
  optional-middle positional rule, the RTD cluster masking and timer-sweep
  invariants, the subject-latest loser path, `TickGate`/`StripedLocks`,
  `NLog.config` keeping the masking renderer, `UpdateCheck`'s semver and "never nag
  a dev build" rules, the Excel-DNA revision pin, `HarnessInvariantTests` (the
  harness safety contracts), the culture-invariance suite (`NatsValueLocaleTests`,
  on a real de-DE thread) and the async bridge (`NatsWriteObservableTests`: the
  work never runs on the subscribing thread, a duplicate subscribe never queues the
  write twice, a throwing write becomes a named error cell, and one endpoint's
  writes stay FIFO).
- The **smoke catalog sweep** calls every frozen function that no other layer
  touched (26 of the 97 had never been invoked anywhere) and requires a usable cell
  from each, so a broken binding, a wrong default or a throw on the empty case can
  no longer reach a user unseen.
- **RTD runtime coverage beyond `SUB`**: the E2E now creates `MSG`, `JS`, `KVGET`,
  `KVKEYS`, `KV` and `OBJ` cells, asserts the frozen unknown-command and
  RTDCLUSTER-vocabulary errors, the unresolvable-reference error, the
  `RTDCLUSTER cluster/nodes` cells, the six `NatsRTD*` helpers and the 255-character
  runtime bound.

## v0.1.1

Fixes the first release's installer, CI and runtime bug fixes found by the expanded
test layers. No change to the public Excel surface or the `NATSExcel.json` shape.

### Added

- **The installer wizard now ships 11 languages** — Chinese Simplified, English,
  Chinese Traditional, Português, Español, Русский, Français, 한국어, Deutsch,
  日本語 and Português (Brasil) — each with the full set of 63 wizard strings, so no
  page falls back to a half-translated state. The two Chinese language files are
  bundled under `installer/languages/` (Inno Setup does not ship them);
  `installer/languages/messages.ps1` reports a missing or extra key for a tag. The
  **Excel-facing surface stays English on purpose** (function names, markers and
  errors), so a formula returns the same text whatever the user's locale.

### Fixed

- **A cluster's `auth`/`reconnect` block erased the shared block.** Applying the
  per-cluster override wrote every field unconditionally, including the `null`/`0`
  the cluster had not declared, so a cluster that set only `token` silently dropped
  the shared `user`/`password`. The override now replaces **field by field** — a
  field the cluster did not declare keeps the shared value.
- **A zero-byte Core payload read as a missing subject.** `NatsUDFSubjectLatest`
  cached the `(null)` "nothing received yet" marker for a message that arrived with
  an empty body, collapsing the two into one cell. An empty payload now caches the
  empty string.
- **An unreachable endpoint misreported as "JetStream is not enabled".**
  `NatsUDFJetStream` and `NatsUDFObjectStore` (like KV) now branch on
  `HasConnection`: a server that answered without JetStream keeps the frozen
  config message, while an endpoint that never answered names the connectivity
  problem.
- **The installer wrote a config its own schema rejects.** `ConfigJson()` emitted the old nested
  shape (`"Default": { "cluster": … }` plus a `"UDF"` block), so a user's first
  `=NatsUDFConfigValidate()` reported an error on the file the wizard had just written. It now
  emits the shipped shape (`"Default"` as a name, top-level `Auth`/`Reconnect`). Verified: a real
  silent install produces a file that `validate-config.ps1` accepts (`OK`).
- **The native server download was unverified.** `install-server.ps1` (and the CI) now pin the
  SHA256 of the release zip and refuse to extract or run a mismatched archive. Verified: the
  correct digest passes; a wrong one discards the file and exits 1.
- **The uninstaller could stop the user's own server.** The `native-file` removal matched any
  `nats-server` process by name; it is now scoped to the install directory.
- **The uninstaller deleted a shared ProgId.** `NatsRtd` is per-user, not per-install: removing it
  broke a second NATSExcel copy's `=RTD` cells. It is removed only when no other NATSExcel XLL is
  registered.
- **`NatsUDFConfigReload` claimed "using defaults" while a config was still in use.** It now says
  the previously loaded configuration is still active (`AppConfig.HasLoadedFile`).
- **Silent mode ignored `/SERVICE`.** `/SERVICE=1` now registers the Windows service (still needing
  elevation, still never failing the add-in install).
- **The CI started the NATS server in a separate step, so it was killed before the tests ran.**
  The Windows runner runs each step in a job object with kill-on-close. The download is now its own
  step (exporting `NATSEXCEL_NATS_SERVER`), and the smoke and liveness steps each start their own
  server through `test\TestSupport\Start-Cluster.ps1`.
- **Documentation:** the README was rewritten for v0.1.0 users; `AGENTS.md` and `PLAN.md` no longer
  claim the solution, smoke and liveness layers are missing; a `LICENSE` (MIT) was added; and a
  double-encoded UTF-8 sequence in a few source comments was corrected.

## v0.1.0

First release. NATSExcel is an Excel add-in (XLL) that integrates NATS with Excel: worksheet
functions (UDFs) for Core NATS and JetStream, plus an RTD server for streaming topics.

### Added

- **Add-in skeleton and packaging.** `NATSExcel.csproj` (net48, C# 9 via the `IsExternalInit`
  shim) and the `NATSExcel.dna` packing manifest. The `.dna` file name is what makes Excel-DNA
  emit `NATSExcel-packed.xll` / `NATSExcel64-packed.xll`, and it packs the full net48 dependency
  closure (ExcelDna, NATS.Client.Core/JetStream/KeyValueStore/ObjectStore, Newtonsoft.Json,
  NLog). A build target copies `NATSExcel.dll` next to the packed XLLs, and `AddIn.cs` wires the
  `AutoOpen`/`AutoClose` lifecycle.
- **Configuration layer.** `AppConfig` loads `NATSExcel.json` once and supports comment-tolerant
  parsing (`StripComments`), endpoint resolution over `Default` / `Servers` / `Clusters`, a shared
  top-level `Auth` and `Reconnect` block with per-cluster overrides, and a JSON-Schema walk
  (`ConfigSchema.cs`) that drops a single wrongly-typed value instead of reverting the whole file.
  NKey `seed_or_path` is resolved by reading a `.nk` file first. `NatsUDFConfigReload` re-reads the
  file without restarting Excel.
- **Runtime layer.** `NatsRuntime`, `NatsConnectionManager` and `NatsSubscriptionManager`: one
  cached `NatsConnection` per resolved endpoint, ref-counted Core subscriptions keyed by
  `(endpoint, subject, queueGroup)`, an async bridge via `ExcelAsyncUtil.Observe`, bounded waits,
  and a process-wide shutdown tombstone with a same-process reload reset.
- **RTD server.** The `NatsRtd` `ExcelRtdServer` with nine topic kinds (`SUB`, `KV`, `KVGET`,
  `KVKEYS`, `REQ`, `JS`, `MSG`, `OBJ`, `RTDCLUSTER`), a poll timer and an Excel flush timer behind
  a reentrancy gate, subscribe backoff, real-time conflation, and a frozen, Excel-free topic
  grammar (`NatsRtdTopic.cs`). Six `NatsRTD*` status helpers accompany the server.
- **UDF catalog.** The full `NatsUDF*` surface across Core NATS, JetStream, Key/Value, Object
  Store and status/diagnostics, including the `...NonVolatile` write twins, and the two pure
  conversions `NatsUDFMatrixToJSON` / `NatsUDFJSONToMatrix`. The catalog exposes the base functions
  plus the write twins and six `NatsRTD*` status helpers; identifiers and values are written with
  the invariant culture, and admin/destructive operations are gated by `AllowAdminWrites`
  (default off).
- **Installer.** An Inno Setup script (`installer/NATSExcel.iss`), per-user and native-first, with
  an Authentication page covering seven auth styles (none, token, user/password, NKey, JWT + seed,
  creds file, mTLS), silent parameters for each, and en/es/pt-BR languages. The native server
  helper writes `nats.conf` without a BOM (the NATS parser rejects one), and a service helper is
  provided for the isolated-service path.
- **Validator tool.** `tools/NatsValidate` validates a `NATSExcel.json` and, with `--test`,
  connects with the same NATS.Net client the add-in uses, so every credential style is exercised
  for real (the PowerShell helpers only speak the wire protocol).
- **Test layers.** The five-layer model: the config/connection validator, the xUnit unit suite
  (config, schema drop, masking, the RTD grammar, the UDF argument vocabulary, the JSON helpers,
  conflation, project-file and `.dna` parity, and an unbounded-blocking source scan), the native
  cluster harness, the Excel end-to-end script, the load test, and the smoke and liveness suites
  (ref-counted subscriptions, wildcards, queue groups, request/reply, JetStream/KV/Object Store,
  the admin gate and the JetStream-disabled path; plus adversarial liveness with a heartbeat
  watchdog).
- **Credential-safe logging.** A `maskedException` NLog layout renderer masks credentials in
  every logged exception (outer and inner), and `MaskNatsEndpoint` redacts userinfo, `password=`,
  JWTs, NKey seeds, creds-file paths and TLS cert paths in error cells, RTD `#ERROR`, status cells
  and log lines.
- **CI.** `.github/workflows/build.yml`: a Release build, unit tests, native `nats-server -js`
  provisioning with a 4222/`jetstream:true` probe, the smoke and liveness suites, an Inno Setup
  installer build on every run, tag guards (CHANGELOG heading, release-asset existence), and a
  release job that uses the CHANGELOG section as the release body.

### Changed

- **Config shape.** `Default` is a plain endpoint name (a `Clusters` name, a `Servers` name or a
  raw URL); endpoints live under `Servers`/`Clusters`; the shared connection defaults moved to
  top-level `Auth`/`Reconnect` blocks that a cluster can override. The earlier nested-`Default`
  shape still loads as a fallback.
- **Transport unified.** Removed the RTD/UDF server split inherited from RedisExcel: one
  `NatsConnection` per endpoint multiplexes reads, subscriptions and JetStream, so a second pool
  bought nothing. The config, loader and schema were aligned to the single-pool model.
- **Per-cluster auth applied.** A cluster's `auth`, `timeout` and `reconnect` settings are now
  actually resolved and applied (matching the resolved URL back to its cluster), instead of only
  the shared block being read.
- **Reconnect policy wired.** The `maxAttempts`/`waitMs`/`jitterMs` knobs now reach `NatsOpts`
  (`MaxReconnectRetry`/`ReconnectWaitMin`/`ReconnectJitter`) rather than being inert config.

### Fixed

- **`NatsUDFRequest` timeout marker.** A request with no responder returned the raw client error
  instead of the frozen `TIMEOUT` marker (`NatsNoReplyException` was unclassified). Found by the
  smoke suite.
- **`NatsUDFObjectGet` missing-object name.** Fetching an absent object rendered `object '' not
  found` with an empty name because the read path passed `null` instead of the object name. Found
  by the smoke suite.
- **`NatsUDFLogPath` on a byte-loaded XLL.** It returned `#VALUE!` because `Assembly.Location` is
  empty inside Excel; it now falls back to the Excel-DNA `XllPath` and never throws.
- **Connection vs JetStream errors (Key/Value).** A KV call against an unreachable
  endpoint reported "JetStream is not enabled", sending users to the wrong problem;
  the KV probe now distinguishes "nobody answered" from "the server has no
  JetStream". (The JetStream and Object Store probes were completed in v0.1.1.)
- **Install and service bugs.** The uninstaller removed every `OPEN` entry containing
  `NATSExcel`, stripping other installs from Excel's add-ins list; it now removes only this
  install's XLL. The native server helper now copies `nats-server.exe` into the target folder so
  the service path has a real binary.
- **Bounded teardown and shutdown races.** Dispose loops run under a total budget instead of a
  per-connection bound (a many-endpoint workbook could freeze Excel for minutes), and
  `GetJetStream` re-checks the shutdown fence after the lazy factory materializes.
- **Excel error cells in a range.** A keys/subjects range containing an Excel error cell was
  silently dropped (rows shifted, streams could be created with fewer subjects than shown); both
  readers now reject it with the frozen error text.
- **`PublishIfChanged` dedup.** A fire-and-forget publish that was never confirmed no longer
  records a dedup marker, which used to starve the feed permanently.
- **Reload and config edge cases.** `ResetAfterAddInReload` removes only the instances it tore
  down; a dangling `Default` falls back to the built-in URL; the `RTD.cluster`/`UDF.cluster`
  overrides are read and declared in the schema.
