# NATSExcel

Excel add-in (XLL) that brings **NATS** into Excel, through RTD (Real-Time
Data) and UDF (User Defined Functions): type a formula, NATS data lands in the
cell and stays live.

Covers **Core NATS** (pub/sub, request/reply, queue groups, headers) and
**JetStream** (streams, consumers, replay, Key/Value Store, Object Store).

> Sibling project to [RedisExcel](https://github.com/rspadim/RedisExcel). Same
> idea, built on NATS. Not a code port: the NATS client is async-only, so the
> async bridge is architecture here, not an option.

- Design and catalog: [`DESIGN.md`](DESIGN.md)
- Implementation plan: [`PLAN.md`](PLAN.md)
- Working guidance for contributors/agents: [`AGENTS.md`](AGENTS.md)

> Status: design phase. The code lives in this repo; the docs above are the
> contract.

---

## Installing

### One-click install (recommended)

1. Go to the **[latest release](https://github.com/rspadim/NATSExcel/releases/latest)**
   and, under **Assets**, click **`NATSExcelSetup-x.y.z.exe`**.
2. Run it and follow the wizard:
   - it copies the add-in to `C:\NATSExcel` and keeps only the XLL that matches
     your Office bitness;
   - ask it to **install a server** if you do not have one yet - NATS runs
     natively on Windows, so it just downloads one small `nats-server.exe`
     (Docker is offered too);
   - it writes a starter `NATSExcel.json` and, at the end, tests the connection
     and can open a sample workbook.
3. Open Excel and (if the add-in did not load itself) enable it:
   `File > Options > Add-ins > Go... > Browse...` and pick
   `NATSExcel64-packed.xll` (or `NATSExcel-packed.xll` for 32-bit Office).

> Windows may show a SmartScreen warning because the XLL is not code-signed:
> click **More info > Run anyway**.

The wizard is available in **English, Español and Português (Brasil)**.

### Manual install

Keep the add-in files in one dedicated folder, for example `C:\NATSExcel\`:

1. Download the latest release (`.xll`, `.dll`) and `NLog.config`
   ([Download Releases](https://github.com/rspadim/NATSExcel/releases)).
2. Download [`NATSExcel.json`](NATSExcel.json) and put it in the same folder
   (the add-in folder is the first place searched for it).
3. In Excel: `File > Options > Add-ins > Go... > Browse...` and pick the `.xll`
   matching your Office bitness.

The config is searched in this order (first existing wins): the add-in folder,
`%USERPROFILE%`, the Excel folder, `C:\Windows`.

> **Updates:** `=NatsUDFUpdateAvailable()` shows TRUE when a newer release
> exists. Toggle with `"UpdateCheck": true|false` in `NATSExcel.json` (default
> on, runs in the background and never blocks Excel).

---

## No server yet? Run NATS on Windows (beginner guide)

NATS needs no Docker on Windows: the server is a single self-contained
`nats-server.exe`, and **JetStream is the same binary** (enabled with `-js`).
Pick one option:

### Option A (recommended): winget

```powershell
winget install NATSAuthors.NATSServer
nats-server --version
```

### Option B: the release zip (no package manager)

```powershell
# Download the Windows release, unzip, run it with JetStream on.
# Adjust the version to the current one.
Invoke-WebRequest 'https://github.com/nats-io/nats-server/releases/download/v2.15.1/nats-server-v2.15.1-windows-amd64.zip' -OutFile nats.zip
Expand-Archive nats.zip -DestinationPath .
.\nats-server-v2.15.1-windows-amd64\nats-server.exe -js --store_dir C:\NATSExcel\server\data
```

The installer uses exactly this path.

### Option C: run it at boot (Windows service)

The server ships no installer; register it as a service (admin):

```powershell
sc.exe create NATS binPath= "C:\NATSExcel\server\nats-server.exe -c C:\NATSExcel\server\nats.conf" start= auto
sc.exe start NATS
```

The installer's "register as a Windows service" option runs the equivalent
(`installer\nats-service.ps1`).

### Option D: Docker (optional)

```powershell
docker run -d --name natsexcel -p 4222:4222 --restart unless-stopped nats:latest -js --store_dir /data
```

All options listen on `localhost:4222` by default. Verify with a NATS client
(`nats server info`), or just connect from Excel.

> **JetStream storage:** the server's default `store_dir` is a *temporary* path.
> Point it at a durable directory (the installer writes
> `jetstream { store_dir: ... }` into `nats.conf` for you), or streams are lost
> on restart.

### Point Excel at another server or cluster

Put the URL in `NATSExcel.json` (a `Clusters` entry) or pass it as the last
argument of a formula:

```excel
=RTD("NatsRtd",,"SUB","md.binance.live.spot.>",,"feed")     // a named cluster
=NatsUDFSubjectLatest("orders.>",,"nats://localhost:4222")   // a raw URL
```

A **cluster** is a comma-separated URL list - **one URL is enough**, the client
discovers the rest. Auth is per cluster: token, user/password, NKey, JWT+seed,
a creds file or mTLS. Credentials are always masked in cells and logs.

---

## Examples

### Publish from Excel

```excel
=NatsUDFPublish("orders.new", "{\"id\":1}")           // send a message
=NatsUDFPublishNonVolatile("orders.new", "once")      // sent once, not per recalc
=NatsUDFPublishJSON("range_of_data", A1:C20)          // publish a matrix as JSON
=NatsUDFPublishWithHeaders("orders.new", "go", A1:B3) // message + headers
```

### Subscribe / read from Excel

```excel
=RTD("NatsRtd",,"SUB","md.binance.live.spot.>")        // live push (wildcards)
=RTD("NatsRtd",,"KV","prices","BTC")                   // live KV watch
=NatsUDFSubjectLatest("prices.BTC")                    // last value, by UDF
=NatsUDFRequest("orders.create", "{\"id\":2}", 2000)   // request/reply
=NatsUDFKVGet("prices", "BTC")                         // JetStream KV read
=NatsUDFJSONToMatrix(NatsUDFSubjectLatest("range_of_data"))
```

### JetStream and topology

```excel
=NatsUDFJSPublish("events.orders", "{\"id\":9}")   // durable publish (ack)
=NatsUDFJSGetMessage("ORDERS", 1)                  // replay by sequence
=NatsUDFJSRaftInfo("ORDERS")                       // leader + replicas + lag
=NatsUDFClusterName()                              // the server's cluster name
=NatsUDFClusterNodes()                             // nodes the client sees
```

The full catalog (~95 functions plus 9 RTD commands) is in
[`DESIGN.md`](DESIGN.md) §7.

---

## Configuration

`NATSExcel.json` drives everything. Minimal file:

```json
{ "Default": { "cluster": "local" },
  "Clusters": { "local": { "urls": ["nats://localhost:4222"] } } }
```

For a cluster with a token, and the real-time window tuned:

```json
{
  "Default": { "cluster": "feed",
               "auth": { "token": null } },
  "Clusters": {
    "feed": { "urls": ["nats://n1:4222", "nats://n2:4222", "nats://n3:4222"],
              "auth": { "token": "@@NATS_TOKEN@@" } }
  },
  "ConflationMs": 200
}
```

Never commit a real endpoint or credential. See `DESIGN.md` §8 for every key.

The installer ships `Validate-Config.cmd` (or `scripts\validate-config.ps1`),
which checks the file against the same JSON Schema the add-in embeds.

---

## Building

```powershell
dotnet build NATSExcel.sln -c Release
```

Outputs in `bin\Release\net48\publish\`: `NATSExcel-packed.xll` (32-bit),
`NATSExcel64-packed.xll` (64-bit) and `NATSExcel.dll`.

The add-in targets **.NET Framework 4.8** with `<LangVersion>9.0</LangVersion>`
plus a small `IsExternalInit` shim, because the NATS .NET client uses init-only
setters.

See [`AGENTS.md`](AGENTS.md) for the test layers and
[`installer/README.md`](installer/README.md) for the installer.

---

## License

NATS.Net is Apache-2.0. NATSExcel's own license is not finalized yet.
