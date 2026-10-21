# NATSExcel — frozen public Excel surface

This is the per-function catalog of the add-in's public Excel surface. It is
**frozen at the first release**: the names, argument order, argument
optionality/defaults, volatility and RTD grammar below may only change by
**addition** (a new optional trailing argument, a new function). Renames,
argument reorders, argument removals and volatility flips are breaking changes
to users' spreadsheets and must not happen (AGENTS golden rule 2).

Every row here was read from the source (`NatsUdf*.cs`, `NatsUdf.cs`,
`NatsRtd.cs`), not from memory. The machine-checked authority is
`test\NATSExcel.Tests\PublicSurfaceParityTests.cs` against
`test\NATSExcel.Tests\PublicSurface.baseline.txt`, which is generated from the
built assembly; this document must agree with it.

## Conventions

- **Arguments** are `name` (required) or `[name]` (optional; every optional
  argument defaults to blank/`null`). Order is the call order in Excel.
- **Volatile** = the `[ExcelFunction(..., IsVolatile = true)]` flag. Volatile
  functions recalc on every calculation; non-volatile ones run once per
  entry/argument change. Reads/status are volatile; writes have a volatile base
  plus an explicit `...NonVolatile` twin.
- **Twin** names the additive `...NonVolatile` write twin: identical arguments,
  thin delegation to the base, `IsVolatile = false`. A twin is a separate
  `[ExcelFunction]` and is listed as its own row.
- **Admin-gated** = the function applies the `AllowAdminWrites` gate (config,
  default `false`); with it off it returns
  `Error: admin writes are disabled (set "AllowAdminWrites": true)`. The gate is
  applied by the shared prologue *after* argument validation and endpoint
  resolution. A twin delegating to a gated base is gated in effect too.

## Counts (code truth)

| Set | Count |
| --- | --- |
| `[ExcelFunction]` total | **97** |
| `NatsUDF*` | **91** |
| `NatsRTD*` (RTD status helpers) | **6** |
| `...NonVolatile` twins | **25** |
| Volatile | **69** |
| Non-volatile | **28** (25 twins + `ConfigReload` + the two JSON conversions) |
| Functions gated by `AllowAdminWrites` (bases) | **11** |
| RTD commands | **9** |

> **Doc drift, documented not hidden.** The maintainer design notes (§7.9) and plan (§1.1) - kept on disk, deliberately not part of this repository - said
> "89 `NatsUDF*` + 6 RTD = 95". That count left the two standalone
> `[ExcelFunction]` conversions (`NatsUDFMatrixToJSON`, `NatsUDFJSONToMatrix`)
> out of the `NatsUDF*` total even though they carry the `NatsUDF` prefix. The
> real, reflection-derived totals are **91** and **97**; those are what the
> parity test pins.

## Core NATS — `NatsUdfCore.cs` (17)

| Function | Arguments | Volatile | Twin | Admin-gated |
| --- | --- | --- | --- | --- |
| `NatsUDFPublish` | `subject`, `message`, `[optionalHost]` | yes | `NatsUDFPublishNonVolatile` | no |
| `NatsUDFPublishNonVolatile` | `subject`, `message`, `[optionalHost]` | no | — | no |
| `NatsUDFPublishJSON` | `subject`, `range`, `[optionalHost]` | yes | `NatsUDFPublishJSONNonVolatile` | no |
| `NatsUDFPublishJSONNonVolatile` | `subject`, `range`, `[optionalHost]` | no | — | no |
| `NatsUDFPublishIfChanged` | `subject`, `message`, `[optionalHost]` | yes | `NatsUDFPublishIfChangedNonVolatile` | no |
| `NatsUDFPublishIfChangedNonVolatile` | `subject`, `message`, `[optionalHost]` | no | — | no |
| `NatsUDFPublishWithHeaders` | `subject`, `message`, `headers`, `[optionalHost]` | yes | `NatsUDFPublishWithHeadersNonVolatile` | no |
| `NatsUDFPublishWithHeadersNonVolatile` | `subject`, `message`, `headers`, `[optionalHost]` | no | — | no |
| `NatsUDFRequest` | `subject`, `message`, `[timeoutMs]`, `[optionalHost]` | yes | `NatsUDFRequestNonVolatile` | no |
| `NatsUDFRequestNonVolatile` | `subject`, `message`, `[timeoutMs]`, `[optionalHost]` | no | — | no |
| `NatsUDFRequestJSON` | `subject`, `range`, `[timeoutMs]`, `[optionalHost]` | yes | `NatsUDFRequestJSONNonVolatile` | no |
| `NatsUDFRequestJSONNonVolatile` | `subject`, `range`, `[timeoutMs]`, `[optionalHost]` | no | — | no |
| `NatsUDFSubjectLatest` | `subject`, `[queueGroup]`, `[optionalHost]` | yes | — | no |
| `NatsUDFSubjectLatestJSON` | `subject`, `[queueGroup]`, `[optionalHost]` | yes | — | no |
| `NatsUDFSubjectUnsubscribe` | `subject`, `[optionalHost]` | yes | `NatsUDFSubjectUnsubscribeNonVolatile` | no |
| `NatsUDFSubjectUnsubscribeNonVolatile` | `subject`, `[optionalHost]` | no | — | no |
| `NatsUDFSubscriptionCount` | — | yes | — | no |

## JetStream — messages, streams, consumers — `NatsUdfJetStream.cs` (11)

| Function | Arguments | Volatile | Twin | Admin-gated |
| --- | --- | --- | --- | --- |
| `NatsUDFJSGetMessage` | `stream`, `seq`, `[host]` | yes | — | no |
| `NatsUDFJSGetLastMessage` | `stream`, `filter`, `[host]` | yes | — | no |
| `NatsUDFJSFetch` | `stream`, `consumer`, `count`, `[timeoutMs]`, `[host]` | yes | — | no |
| `NatsUDFJSStreamInfo` | `stream`, `[host]` | yes | — | no |
| `NatsUDFJSStreams` | `[host]` | yes | — | no |
| `NatsUDFJSConsumerInfo` | `stream`, `consumer`, `[host]` | yes | — | no |
| `NatsUDFJSConsumers` | `stream`, `[host]` | yes | — | no |
| `NatsUDFJSPublish` | `subject`, `message`, `[host]` | yes | `NatsUDFJSPublishNonVolatile` | no |
| `NatsUDFJSPublishNonVolatile` | `subject`, `message`, `[host]` | no | — | no |
| `NatsUDFJSPublishJSON` | `subject`, `range`, `[host]` | yes | `NatsUDFJSPublishJSONNonVolatile` | no |
| `NatsUDFJSPublishJSONNonVolatile` | `subject`, `range`, `[host]` | no | — | no |

## JetStream admin (opt-in) — `NatsUdfJetStream.cs` (16)

| Function | Arguments | Volatile | Twin | Admin-gated |
| --- | --- | --- | --- | --- |
| `NatsUDFJSDeleteMessage` | `stream`, `seq`, `[host]` | yes | `NatsUDFJSDeleteMessageNonVolatile` | yes |
| `NatsUDFJSDeleteMessageNonVolatile` | `stream`, `seq`, `[host]` | no | — | yes |
| `NatsUDFJSPurge` | `stream`, `[filter]`, `[host]` | yes | `NatsUDFJSPurgeNonVolatile` | yes |
| `NatsUDFJSPurgeNonVolatile` | `stream`, `[filter]`, `[host]` | no | — | yes |
| `NatsUDFStreamCreate` | `name`, `subjects`, `[storage]`, `[retention]`, `[host]` | yes | `NatsUDFStreamCreateNonVolatile` | yes |
| `NatsUDFStreamCreateNonVolatile` | `name`, `subjects`, `[storage]`, `[retention]`, `[host]` | no | — | yes |
| `NatsUDFStreamUpdate` | `name`, `subjects`, `[storage]`, `[retention]`, `[host]` | yes | `NatsUDFStreamUpdateNonVolatile` | yes |
| `NatsUDFStreamUpdateNonVolatile` | `name`, `subjects`, `[storage]`, `[retention]`, `[host]` | no | — | yes |
| `NatsUDFStreamDelete` | `name`, `[host]` | yes | `NatsUDFStreamDeleteNonVolatile` | yes |
| `NatsUDFStreamDeleteNonVolatile` | `name`, `[host]` | no | — | yes |
| `NatsUDFConsumerCreate` | `stream`, `name`, `[filter]`, `[deliverPolicy]`, `[host]` | yes | `NatsUDFConsumerCreateNonVolatile` | yes |
| `NatsUDFConsumerCreateNonVolatile` | `stream`, `name`, `[filter]`, `[deliverPolicy]`, `[host]` | no | — | yes |
| `NatsUDFConsumerDelete` | `stream`, `name`, `[host]` | yes | `NatsUDFConsumerDeleteNonVolatile` | yes |
| `NatsUDFConsumerDeleteNonVolatile` | `stream`, `name`, `[host]` | no | — | yes |
| `NatsUDFConsumerPause` | `stream`, `name`, `[pauseMs]`, `[host]` | yes | `NatsUDFConsumerPauseNonVolatile` | yes |
| `NatsUDFConsumerPauseNonVolatile` | `stream`, `name`, `[pauseMs]`, `[host]` | no | — | yes |

## Key/Value Store — `NatsUdfKv.cs` (18)

| Function | Arguments | Volatile | Twin | Admin-gated |
| --- | --- | --- | --- | --- |
| `NatsUDFKVGet` | `bucket`, `key`, `[host]` | yes | — | no |
| `NatsUDFKVGetMultiple` | `bucket`, `keys`, `[columns]`, `[host]` | yes | — | no |
| `NatsUDFKVKeys` | `bucket`, `[pattern]`, `[host]` | yes | — | no |
| `NatsUDFKVStatus` | `bucket`, `key`, `[host]` | yes | — | no |
| `NatsUDFKVHistory` | `bucket`, `key`, `[host]` | yes | — | no |
| `NatsUDFKVBuckets` | `[host]` | yes | — | no |
| `NatsUDFKVPut` | `bucket`, `key`, `value`, `[host]` | yes | `NatsUDFKVPutNonVolatile` | no |
| `NatsUDFKVPutNonVolatile` | `bucket`, `key`, `value`, `[host]` | no | — | no |
| `NatsUDFKVPutJSON` | `bucket`, `key`, `range`, `[host]` | yes | `NatsUDFKVPutJSONNonVolatile` | no |
| `NatsUDFKVPutJSONNonVolatile` | `bucket`, `key`, `range`, `[host]` | no | — | no |
| `NatsUDFKVCreate` | `bucket`, `key`, `value`, `[host]` | yes | `NatsUDFKVCreateNonVolatile` | no |
| `NatsUDFKVCreateNonVolatile` | `bucket`, `key`, `value`, `[host]` | no | — | no |
| `NatsUDFKVUpdate` | `bucket`, `key`, `value`, `revision`, `[host]` | yes | `NatsUDFKVUpdateNonVolatile` | no |
| `NatsUDFKVUpdateNonVolatile` | `bucket`, `key`, `value`, `revision`, `[host]` | no | — | no |
| `NatsUDFKVDelete` | `bucket`, `key`, `[host]` | yes | `NatsUDFKVDeleteNonVolatile` | yes |
| `NatsUDFKVDeleteNonVolatile` | `bucket`, `key`, `[host]` | no | — | yes |
| `NatsUDFKVPurge` | `bucket`, `key`, `[host]` | yes | `NatsUDFKVPurgeNonVolatile` | yes |
| `NatsUDFKVPurgeNonVolatile` | `bucket`, `key`, `[host]` | no | — | yes |

## Object Store — `NatsUdfObjectStore.cs` (9)

| Function | Arguments | Volatile | Twin | Admin-gated |
| --- | --- | --- | --- | --- |
| `NatsUDFObjectGet` | `bucket`, `name`, `[host]` | yes | — | no |
| `NatsUDFObjectInfo` | `bucket`, `name`, `[host]` | yes | — | no |
| `NatsUDFObjectList` | `bucket`, `[host]` | yes | — | no |
| `NatsUDFObjectWatch` | `bucket`, `[host]` | yes | — | no |
| `NatsUDFObjectBuckets` | `[host]` | yes | — | no |
| `NatsUDFObjectPut` | `bucket`, `name`, `value`, `[description]`, `[host]` | yes | `NatsUDFObjectPutNonVolatile` | no |
| `NatsUDFObjectPutNonVolatile` | `bucket`, `name`, `value`, `[description]`, `[host]` | no | — | no |
| `NatsUDFObjectDelete` | `bucket`, `name`, `[host]` | yes | `NatsUDFObjectDeleteNonVolatile` | yes |
| `NatsUDFObjectDeleteNonVolatile` | `bucket`, `name`, `[host]` | no | — | yes |

## Status, diagnostics and config — `NatsUdf.cs` + `NatsUdfStatus.cs` (18)

`NatsUDFVersion` returns one line - `NATSExcel <ver> | xll=<path> | config=<path
or "(defaults)"> | connections=<n>` - with the version from the release tag (a
local build reports `dev`), the **full path of the loaded XLL** (two installed
copies report two different paths, which is how a duplicated install becomes
visible from a cell), the config file in use and the live connection count. It
needs no server and never throws.

| Function | Source | Arguments | Volatile | Admin-gated |
| --- | --- | --- | --- | --- |
| `NatsUDFVersion` | `NatsUdf.cs` | — | yes | no |
| `NatsUDFConfigPath` | `NatsUdf.cs` | — | yes | no |
| `NatsUDFConnectionCount` | `NatsUdf.cs` | — | yes | no |
| `NatsUDFLogPath` | `NatsUdf.cs` | — | yes | no |
| `NatsUDFServerInfo` | `NatsUdfStatus.cs` | `[host]` | yes | no |
| `NatsUDFStats` | `NatsUdfStatus.cs` | `[host]` | yes | no |
| `NatsUDFRtt` | `NatsUdfStatus.cs` | `[host]` | yes | no |
| `NatsUDFUpdateAvailable` | `NatsUdfStatus.cs` | — | yes | no |
| `NatsUDFClusterName` | `NatsUdfStatus.cs` | `[host]` | yes | no |
| `NatsUDFClusterNodes` | `NatsUdfStatus.cs` | `[host]` | yes | no |
| `NatsUDFConnectionState` | `NatsUdfStatus.cs` | `[host]` | yes | no |
| `NatsUDFClusterInfo` | `NatsUdfStatus.cs` | `[host]` | yes | no |
| `NatsUDFClusterHealth` | `NatsUdfStatus.cs` | `[host]` | yes | no |
| `NatsUDFJSRaftInfo` | `NatsUdfStatus.cs` | `stream`, `[host]` | yes | no |
| `NatsUDFJSStreamPlacement` | `NatsUdfStatus.cs` | `stream`, `[host]` | yes | no |
| `NatsUDFConfigSchema` | `NatsUdfStatus.cs` | — | yes | no |
| `NatsUDFConfigValidate` | `NatsUdfStatus.cs` | `[path]` | yes | no |
| `NatsUDFConfigReload` | `NatsUdfStatus.cs` | — | no | no |

## Matrix ⇄ JSON conversions — `NatsUdf.cs` (2)

Both are server-free, non-volatile pure conversions over the same `NatsJson`
helpers the `...JSON` functions use (the maintainer design notes, section 7.6; those notes are kept on disk and are not part of this repository).

| Function | Arguments | Volatile | Admin-gated |
| --- | --- | --- | --- |
| `NatsUDFMatrixToJSON` | `range` | no | no |
| `NatsUDFJSONToMatrix` | `json`, `[nullValue]` | no | no |

## RTD status helpers — `NatsRtd.cs` (6)

All volatile, none has a twin, none is admin-gated. They carry
`[ExcelFunction]` like the UDFs but live with the RTD server (maintainer design notes, section 7.9; not part of this repository).

| Function | Arguments | Volatile | Admin-gated |
| --- | --- | --- | --- |
| `NatsRTDConnectionCount` | — | yes | no |
| `NatsRTDTopicCount` | — | yes | no |
| `NatsRTDSubscriptionCount` | — | yes | no |
| `NatsRTDSubjectCount` | — | yes | no |
| `NatsRTDDefaultHost` | — | yes | no |
| `NatsRTDCluster` | — | yes | no |

## RTD server and commands

The RTD server is the COM class `NATSExcel.NatsRtd` with the frozen ProgId
**`NatsRtd`**, used as `=RTD("NatsRtd", , <command>, ...)`. The nine frozen
commands (maintainer plan, section 1.2; not part of this repository; pinned by
`test\NATSExcel.Tests\NatsRtdTopicGrammarTests.cs`):

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
| `RTDCLUSTER` | `[1]what [2]endpoint` | poll | no |

The unknown-command error text is frozen:
`unknown command 'X', expected one of [SUB, KV, KVGET, KVKEYS, REQ, JS, MSG, OBJ, RTDCLUSTER]`.
