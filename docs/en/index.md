---
title: NATSExcel
description: Excel add-in (XLL) that brings NATS into Excel through RTD and user-defined functions
---

<p align="center">
  <img class="nx-logo-light" src="../assets/logo.svg" alt="NATSExcel" width="420">
  <img class="nx-logo-dark" src="../assets/logo-dark.svg" alt="NATSExcel" width="420">
</p>

**Excel add-in (XLL) that brings [NATS](https://nats.io/) into Excel** through RTD
(Real-Time Data) and UDF (User-Defined Functions): type a formula, NATS data lands
in the cell, and — for the streaming functions — stays live.

It covers **Core NATS** (pub/sub, request/reply, queue groups, headers) and
**JetStream** (streams, consumers, replay, Key/Value Store, Object Store). Built
with Excel-DNA on .NET Framework 4.8, backed by NATS.Net, NLog and Newtonsoft.Json.

Windows builds are **code-signed**; signing is provided free of charge to
open-source projects by the [SignPath Foundation](https://signpath.org/).

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/rspadim/NATSExcel/blob/main/LICENSE)
[![Latest release](https://img.shields.io/github/v/release/rspadim/NATSExcel?sort=semver)](https://github.com/rspadim/NATSExcel/releases/latest)

## Requirements

- **Windows** with Microsoft Office (Excel), 64-bit or 32-bit.
- **.NET Framework 4.8**, which ships with Windows 10 and Windows 11.
- A **NATS server** to connect to. The installer can start one natively on
  Windows; Docker is optional and never required.

## What it does

| Area | Coverage |
| --- | --- |
| Core NATS | publish, request/reply, subject-latest, unsubscribe, queue groups, headers |
| JetStream | publish-with-ack, get/fetch/replay, streams and consumers (info + admin) |
| Key/Value | get/put/create/update/delete/keys/history/watch |
| Object Store | put/get/info/list/watch/delete |
| Status | version, config path, connection count, server info, RTT, cluster topology |
| Live cells | `=RTD("NatsRtd", ...)` topics: `SUB`, `KV`, `KVGET`, `KVKEYS`, `REQ`, `JS`, `MSG`, `OBJ`, `RTDCLUSTER` |

## Download

Grab the latest release from the **[releases page](https://github.com/rspadim/NATSExcel/releases/latest)**:

| File | Purpose |
| --- | --- |
| `NATSExcelSetup-<version>.exe` | One-click installer (per-user, native-first; wizard in 11 languages) |
| `NATSExcel64-packed.xll` | 64-bit Office add-in |
| `NATSExcel-packed.xll` | 32-bit Office add-in |
| `NATSExcel.dll` | Managed add-in assembly (shipped next to the XLL) |
| `NLog.config` | Logging configuration |
| `NATSExcel.json` | Starter configuration |

Releases are code-signed; the signing service is provided free of charge to
open-source projects by the [SignPath Foundation](https://signpath.org/).

## Install

1. Run `NATSExcelSetup-<version>.exe`. It copies the add-in, detects your Office
   bitness, writes a starter `NATSExcel.json`, and (optionally) starts a NATS
   server for you.
2. Open Excel. The add-in registers itself, so the XLL loads automatically. If it
   does not, enable it by hand in **File → Options → Add-ins → Manage: Excel
   Add-ins → Browse…**.
3. Connect: edit `NATSExcel.json` (or use the installer wizard). The default is
   `nats://localhost:4222`.

No NATS server yet? The installer can run one natively on Windows (a single
`nats-server.exe`, JetStream included). Docker is optional, never required.

## See it work

```excel
=NatsUDFVersion()                                        ' the add-in answers
=NatsUDFPublish("orders.new","hello")                    ' Core publish
=NatsUDFSubjectLatest("orders.new")                      ' read the latest value
=NatsUDFKVPut("config","rate","1.23")                    ' Key/Value write
=RTD("NatsRtd",,"SUB","orders.>")                        ' a live cell
```

## Documentation

- **Function catalog (frozen):** [PUBLIC-SURFACE.md](../PUBLIC-SURFACE.md)
- **Architecture and design:** [DESIGN.md](https://github.com/rspadim/NATSExcel/blob/main/DESIGN.md)
- **Getting started guide:** [Readme.md](https://github.com/rspadim/NATSExcel/blob/main/Readme.md)
- **Development plan:** [PLAN.md](https://github.com/rspadim/NATSExcel/blob/main/PLAN.md)
- **Tests:** [test/README.md](https://github.com/rspadim/NATSExcel/blob/main/test/README.md)

## Privacy

NATSExcel collects no user data. The only request it makes on its own is a single
background call to the GitHub releases API, used to check whether a newer version
exists; the result is cached and failures (offline, rate limit) are ignored. It
writes a local log file (`NATSExcel.log`) next to its own files. NATS traffic goes
only to the server(s) you configure. The update check can be turned off with
`UpdateCheck: false` in `NATSExcel.json`.

## Support and contributing

Questions, bug reports and pull requests are welcome at
**[github.com/rspadim/NATSExcel](https://github.com/rspadim/NATSExcel/issues)**.
When filing a bug, please include your add-in version (`=NatsUDFVersion()`), your
Office bitness, and whether the server had JetStream enabled.

## Translations

This site is in Portuguese (Brazil), English and Spanish. **Translations are very
welcome** — it is one of the easiest ways to help the project, even if you do not
know C#:

1. Open an _issue_ at
   **[github.com/rspadim/NATSExcel/issues](https://github.com/rspadim/NATSExcel/issues)**
   saying which language you want to add, so nobody duplicates the work.
2. Open a _pull request_ with a new `docs/<locale>/` folder (for example
   `docs/de/`), copying `docs/en/index.md` and translating the text.
3. Add the matching `locale` to the `plugins: i18n: languages:` list in
   `mkdocs.yml` (two or three lines).

The **function catalog** and the rest of the technical documentation do not need
translating: they are the add-in's English reference.

Other ways to help: report bugs with your add-in version and Office bitness,
improve the docs, test on different Excel versions, and suggest usage examples.

## License

NATSExcel is released under the **MIT License** — see
[LICENSE](https://github.com/rspadim/NATSExcel/blob/main/LICENSE).
The bundled NATS.Net client is Apache-2.0.
