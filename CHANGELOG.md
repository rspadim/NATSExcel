# Changelog

All notable changes to this project are documented in this file. The format follows the spirit of
[Keep a Changelog](https://keepachangelog.com/), and the project uses annotated `v*` tags.

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
