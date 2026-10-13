using ExcelDna.Integration;
using NATS.Client.Core;
using NATS.Client.JetStream;
using NATS.Client.JetStream.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace NATSExcel
{
    /// <summary>
    /// Status, diagnostics and configuration worksheet functions
    /// (DESIGN.md section 7.6) plus the cluster/server topology surface
    /// (DESIGN.md section 7.8).
    ///
    /// <para><b>Two tiers.</b> Tier 1 reads the <c>INatsServerInfo</c> the client
    /// already holds (server name/version/cluster, the discovered connect URLs,
    /// the connection state): no extra round trip and no account privilege.
    /// Tier 2 reads the NATS system account
    /// (<c>$SYS.REQ.SERVER.PING</c>): if the endpoint is not authorized for it
    /// the call returns the NAMED error
    /// <c>Error: the system account is not available on this endpoint (cluster
    /// queries need a system account)</c>, never an empty matrix that looks
    /// healthy. <c>NatsUDFJSRaftInfo</c>/<c>NatsUDFJSStreamPlacement</c> read the
    /// stream's own <c>cluster</c> block, so they work with the data account
    /// (DESIGN.md section 7.8).</para>
    ///
    /// <para><b>Guarded.</b> Every function is wrapped in
    /// <see cref="Guarded{T}"/>: <see cref="NatsRuntime"/> deliberately refuses to
    /// resurrect its managers after <c>Shutdown</c>/<c>ServerTerminate</c>, so an
    /// unguarded read would surface <see cref="InvalidOperationException"/> in a
    /// cell during Excel teardown or a same-process reload. A teardown therefore
    /// returns a neutral value (0, a marker, a single-cell matrix), never a
    /// thrown exception. A genuinely unreachable endpoint is reported as the
    /// usual masked <c>Error: ...</c> text.</para>
    ///
    /// <para><b>No unbounded blocking.</b> Nothing here calls a NATS operation on
    /// the Excel thread unbounded: the only blocking primitive used is
    /// <see cref="NatsConnectionManager.WaitBoundedResult{T}"/> /
    /// <see cref="NatsConnectionManager.WaitBounded"/>.</para>
    ///
    /// <para><b>Already defined elsewhere.</b> <c>NatsUDFVersion</c> and
    /// <c>NatsUDFConfigPath</c> live in <c>NatsUdf.cs</c> (the seed diagnostics)
    /// and <c>NatsUDFConnectionCount</c> there already reports the live connection
    /// count. They are deliberately NOT redefined here: a duplicate
    /// <c>[ExcelFunction]</c> name would collide.</para>
    /// </summary>
    public static class NatsUdfStatus
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>A cluster query without the system account is a named cause,
        /// not an empty result (PLAN section 1.3 / DESIGN.md section 7.8).</summary>
        private const string NoSystemAccount =
            "the system account is not available on this endpoint (cluster queries need a system account)";

        /// <summary>The <c>$SYS</c> cluster discovery/inspection subject.</summary>
        private const string SysPingSubject = "$SYS.REQ.SERVER.PING";

        /// <summary>Defensive cap on the nodes a single PING drain accumulates.</summary>
        private const int MaxSysNodes = 64;

        /// <summary>A neutral cell for "no live connection yet".</summary>
        private const string NoConnection = "(no connection)";

        /// <summary>A server that has not answered INFO yet.</summary>
        private const string Connecting = "(connecting)";

        // The client keeps NatsConnection.GetStats() internal in 3.3.0 (verified
        // against the shipped assembly: NatsStats is public, the accessor is not).
        // The single reflected accessor below is the only way to surface the
        // connection counters through the public API; if it ever disappears the
        // add-in degrades to a "no counters" answer instead of faulting a cell.
        private static readonly MethodInfo StatsMethod = ResolveStatsMethod();

        // ==================================================================
        // section 7.6 - server / stats / rtt / update
        // ==================================================================

        /// <summary>
        /// The connected server's advertised facts as compact JSON
        /// (<c>server_id</c>, <c>server_name</c>, <c>version</c>, <c>go</c>,
        /// <c>host</c>, <c>port</c>, <c>cluster</c>, <c>connect_urls</c>,
        /// <c>max_payload</c>, <c>proto</c>, <c>jetstream</c>). Every URL is
        /// masked. <c>(connecting)</c> while no INFO has arrived yet.
        /// </summary>
        [ExcelFunction(Description = "Returns the connected NATS server's info as JSON", IsVolatile = true)]
        public static object NatsUDFServerInfo(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<object>(() => ServerInfoBody(host), NoConnection);
        }

        /// <summary>
        /// The connection's message/byte counters as compact JSON
        /// (<c>in_msgs</c>, <c>out_msgs</c>, <c>in_bytes</c>, <c>out_bytes</c>).
        /// <c>(connecting)</c> while no client is live.
        /// </summary>
        [ExcelFunction(Description = "Returns the NATS connection counters as JSON", IsVolatile = true)]
        public static object NatsUDFStats(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<object>(() => StatsBody(host), NoConnection);
        }

        /// <summary>
        /// A bounded round-trip time to the server in milliseconds (a real
        /// <c>PING</c>/<c>PONG</c>, not a cached value). A failed probe is the
        /// usual masked error text; a teardown returns <c>-1</c> (no measurement).
        /// </summary>
        [ExcelFunction(Description = "Returns the round-trip time to the NATS server in milliseconds", IsVolatile = true)]
        public static object NatsUDFRtt(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<object>(() => RttBody(host), -1);
        }

        /// <summary>
        /// TRUE when the background release check knows of a newer version. It
        /// never blocks and never touches the network on the Excel thread: it
        /// asks for a refresh (rate-limited internally) and immediately reports
        /// the cached answer.
        /// </summary>
        [ExcelFunction(Description = "TRUE when a newer NATSExcel release is known to exist", IsVolatile = true)]
        public static object NatsUDFUpdateAvailable()
        {
            return Guarded<object>(() =>
            {
                try
                {
                    UpdateCheck.EnsureFresh(TimeSpan.FromHours(6));
                }
                catch (Exception ex)
                {
                    // A refresh request must never fault the cell.
                    logger.Debug(ex, "NatsUDFUpdateAvailable: scheduling the update check failed");
                }

                return UpdateCheck.IsUpdateAvailable();
            }, false);
        }

        // ==================================================================
        // section 7.8 Tier 1 - topology the client already knows (no $SYS)
        // ==================================================================

        /// <summary>
        /// The cluster name the server announces, or <c>(standalone)</c> for a
        /// server that reports none. This is the server's cluster, not the local
        /// <c>Clusters</c> label a formula uses (DESIGN.md D12/D14).
        /// </summary>
        [ExcelFunction(Description = "Returns the cluster name the server announces", IsVolatile = true)]
        public static object NatsUDFClusterName(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<string>(() => ClusterNameBody(host), NoConnection);
        }

        /// <summary>
        /// The nodes the CLIENT can see as an N x 3 matrix
        /// (<c>[0] server_name</c>, <c>[1] cluster</c>, <c>[2] connect_url</c>),
        /// built from the client's own discovery
        /// (<c>INatsServerInfo.ClientConnectUrls</c>): what the client can see,
        /// not the full cluster membership (that needs the system account). Every
        /// URL is masked.
        /// </summary>
        [ExcelFunction(Description = "Lists the cluster nodes the client can see as an Nx3 matrix", IsVolatile = true)]
        public static object NatsUDFClusterNodes(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<object>(() => ClusterNodesBody(host), SingleCell(NoConnection));
        }

        /// <summary>
        /// The connection state (<c>CONNECTED</c> / <c>RECONNECTING</c> /
        /// <c>CLOSED</c> / <c>LAME_DUCK</c>) followed by the masked endpoint the
        /// client is configured for.
        /// </summary>
        [ExcelFunction(Description = "Returns the NATS connection state and the masked endpoint", IsVolatile = true)]
        public static object NatsUDFConnectionState(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<string>(() => ConnectionStateBody(host), "CLOSED");
        }

        // ==================================================================
        // section 7.8 Tier 2 - system account ($SYS)
        // ==================================================================

        /// <summary>
        /// The cluster summary from <c>$SYS.REQ.SERVER.PING</c> as compact JSON
        /// (<c>server_count</c> plus a <c>nodes</c> array of
        /// <c>server_name</c>/<c>server_id</c>/<c>version</c>/<c>go</c>/
        /// <c>cluster</c>/<c>jetstream</c>). Without the system account it is the
        /// named error, never an empty result.
        /// </summary>
        [ExcelFunction(Description = "Returns the cluster summary (system account) as JSON", IsVolatile = true)]
        public static object NatsUDFClusterInfo(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<object>(() => ClusterInfoBody(host), NoConnection);
        }

        /// <summary>
        /// Per-node health from <c>$SYS.REQ.SERVER.PING</c> as compact JSON
        /// (<c>server_count</c> plus a <c>nodes</c> array of <c>server_name</c>,
        /// <c>mem</c>, <c>cpu</c>, <c>connections</c>, <c>slow_consumers</c> and
        /// the <c>jetstream</c> state). Without the system account it is the named
        /// error, never an empty result.
        /// </summary>
        [ExcelFunction(Description = "Returns per-node cluster health (system account) as JSON", IsVolatile = true)]
        public static object NatsUDFClusterHealth(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<object>(() => ClusterHealthBody(host), NoConnection);
        }

        /// <summary>
        /// A stream's Raft group as an N x 4 matrix:
        /// <c>[0] peer</c>, <c>[1] current</c> (TRUE for the leader),
        /// <c>[2] active</c> (milliseconds since the peer was last seen),
        /// <c>[3] lag</c> (uncommitted operations behind the leader). This reads
        /// the stream's own <c>cluster</c> block, so it works with the data
        /// account (DESIGN.md section 7.8) - and needs JetStream enabled.
        /// </summary>
        [ExcelFunction(Description = "Returns a stream's Raft group as an Nx4 matrix", IsVolatile = true)]
        public static object NatsUDFJSRaftInfo(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<object>(() => RaftInfoBody(stream, host), SingleCell(NoConnection));
        }

        /// <summary>
        /// Where a stream's Raft group is placed (the stream's
        /// <c>placement.cluster</c>/<c>tags</c>), or
        /// <c>(no placement constraint)</c>. This is the source of the
        /// <c>no suitable peers for placement</c> case (DESIGN.md section 7.8).
        /// </summary>
        [ExcelFunction(Description = "Returns a stream's Raft placement as text", IsVolatile = true)]
        public static object NatsUDFJSStreamPlacement(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return Guarded<object>(() => StreamPlacementBody(stream, host), NoConnection);
        }

        // ==================================================================
        // section 7.6 - configuration
        // ==================================================================

        /// <summary>
        /// The embedded JSON Schema text that backs
        /// <see cref="NatsUDFConfigValidate"/>. Volatile: the text only changes
        /// with the add-in, but a volatile cell is re-evaluated on every
        /// calculation, so the returned schema is never a stale cached value.
        /// </summary>
        [ExcelFunction(Description = "Returns the embedded NATSExcel.json JSON Schema", IsVolatile = true)]
        public static object NatsUDFConfigSchema()
        {
            return Guarded<string>(() =>
            {
                string schema = ConfigSchema.SchemaText;
                if (string.IsNullOrWhiteSpace(schema))
                    return "(schema not embedded)";
                return NatsResultFormatter.BoundCellText(schema);
            }, "(schema not embedded)");
        }

        /// <summary>
        /// Validates a <c>NATSExcel.json</c> (the process file when
        /// <paramref name="path"/> is omitted) against the embedded schema and
        /// returns <c>OK (...)</c>, or an <c>Error:</c> summary with one line per
        /// problem. Comments (<c>//</c>, <c>/* */</c>) are accepted; this is the
        /// only config check (DESIGN.md section 7.6, D13).
        /// </summary>
        [ExcelFunction(Description = "Validates a NATSExcel.json against the embedded schema", IsVolatile = true)]
        public static object NatsUDFConfigValidate(
            [ExcelArgument(Name = "path", Description = "Optional path to a NATSExcel.json")] object path = null)
        {
            return Guarded<object>(() =>
            {
                string file = null;
                if (!IsBlank(path))
                {
                    if (!(path is string text))
                        return NatsUdfCommon.Fail("the path must be text");
                    file = text.Trim();
                }

                return NatsResultFormatter.BoundCellText(ConfigSchema.Validate(file));
            }, "(validation unavailable)");
        }

        /// <summary>
        /// Re-reads the configuration file and installs the freshly parsed
        /// configuration for subsequent calls. Non-volatile (a recalculation must
        /// not re-read the disk), and it deliberately does <b>not</b> tear down
        /// connections: endpoints are resolved per call, so the new values apply
        /// to the next call while live subscriptions keep running (DESIGN.md
        /// section 7.6, PLAN S6.4). A file that does not validate is reported and
        /// not installed.
        /// </summary>
        /// <remarks>
        /// A file that does not validate is reported and not installed. Reload
        /// deliberately does not tear down connections: endpoints are resolved per
        /// call, so the new values apply to the next call while live subscriptions
        /// keep running.
        /// </remarks>
        [ExcelFunction(Description = "Re-reads NATSExcel.json without restarting Excel")]
        public static object NatsUDFConfigReload()
        {
            return Guarded<string>(() =>
            {
                string path = AppConfig.ConfigPath;
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    // No file to read. If one was loaded before, the process is
                    // STILL using it (the root is cached); say so instead of
                    // claiming "using defaults", which would be a lie.
                    return AppConfig.HasLoadedFile
                        ? "OK (no " + AppConfig.ConfigFileName + " found; still using the configuration loaded at startup)"
                        : "OK (no " + AppConfig.ConfigFileName + " found; using defaults)";
                }

                string validation = ConfigSchema.Validate(path);
                if (validation != null && validation.StartsWith(NatsUdfCommon.ErrorPrefix, StringComparison.Ordinal))
                    return NatsResultFormatter.BoundCellText(validation + " (not reloaded)");

                try
                {
                    AppConfig.Reload();
                    // The RTD server froze its tuning at ServerStart; re-apply it
                    // so an edited rate/conflation/endpoint takes effect without an
                    // Excel restart (the "reload without restart" contract). The
                    // RTD layer registers the hook (it is excluded from the unit
                    // project, so this layer cannot reference it directly).
                    try { NatsRuntime.ConfigReloadHook?.Invoke(); }
                    catch (Exception hookEx) { logger.Warn(hookEx, "NatsUDFConfigReload: the RTD refresh hook failed"); }
                    logger.Info("NatsUDFConfigReload: reloaded " + path);
                    return "OK (reloaded " + path + ")";
                }
                catch (Exception ex)
                {
                    return NatsUdfCommon.VerboseError(
                        "config reload", NatsUdfAsync.SafeMessage(ex), "fix the JSON and retry");
                }
            }, "(reload unavailable)");
        }

        // ==================================================================
        // Bodies
        // ==================================================================

        private static object ServerInfoBody(object host)
        {
            if (!TryResolve(host, out string endpoint, out string error))
                return error;

            INatsServerInfo info = LiveServerInfo(endpoint);
            if (info == null)
                return Connecting;

            var urls = new List<string>();
            if (info.ClientConnectUrls != null)
                foreach (string url in info.ClientConnectUrls)
                    urls.Add(Mask(url));

            return NatsResultFormatter.BoundCellText(NatsResultFormatter.FormatServerInfoJson(
                info.Id,
                info.Name,
                info.Version,
                info.GoVersion,
                info.Host,
                info.Port,
                info.Cluster,
                urls,
                info.MaxPayload,
                (int)info.ProtocolVersion,
                info.JetStreamAvailable));
        }

        private static object StatsBody(object host)
        {
            if (!TryResolve(host, out string endpoint, out string error))
                return error;

            NatsConnection connection = MaterializedConnection(endpoint);
            if (connection == null)
                return Connecting;

            NatsStats stats;
            if (!TryReadStats(connection, out stats))
                return "(no counters)";

            return NatsResultFormatter.BoundCellText(NatsResultFormatter.FormatStatsJson(
                stats.ReceivedMessages,
                stats.SentMessages,
                stats.ReceivedBytes,
                stats.SentBytes));
        }

        private static object RttBody(object host)
        {
            if (!TryResolve(host, out string endpoint, out string error))
                return error;

            int bound = ResponseTimeoutMs();
            try
            {
                NatsConnection connection = NatsRuntime.Connections.GetConnection(endpoint);
                TimeSpan rtt = NatsConnectionManager.WaitBoundedResult(
                    connection.PingAsync(CancellationToken.None).AsTask(), bound);
                return rtt.TotalMilliseconds;
            }
            catch (TimeoutException)
            {
                return NatsUdfCommon.VerboseError(
                    "rtt",
                    "no PONG within " + bound.ToString(CultureInfo.InvariantCulture) + "ms",
                    NatsUdfCommon.ErrorHint("timed out"));
            }
            catch (Exception ex)
            {
                return NatsUdfCommon.VerboseError("rtt", NatsUdfAsync.SafeMessage(ex), null);
            }
        }

        private static string ClusterNameBody(object host)
        {
            if (!TryResolve(host, out string endpoint, out string error))
                return error;

            INatsServerInfo info = LiveServerInfo(endpoint);
            if (info == null)
                return Connecting;
            return string.IsNullOrWhiteSpace(info.Cluster) ? "(standalone)" : info.Cluster;
        }

        private static object ClusterNodesBody(object host)
        {
            if (!TryResolve(host, out string endpoint, out string error))
                return error;

            INatsServerInfo info = LiveServerInfo(endpoint);
            if (info == null)
                return SingleCell(Connecting);

            string name = info.Name ?? string.Empty;
            string cluster = info.Cluster ?? string.Empty;
            var rows = new List<string[]>
            {
                new[] { name, cluster, Mask(endpoint) }
            };

            if (info.ClientConnectUrls != null)
            {
                foreach (string url in info.ClientConnectUrls)
                {
                    if (string.IsNullOrWhiteSpace(url))
                        continue;
                    rows.Add(new[] { name, cluster, Mask(url) });
                }
            }

            return NatsResultFormatter.FormatRowsMatrix(rows);
        }

        private static string ConnectionStateBody(object host)
        {
            if (!TryResolve(host, out string endpoint, out string error))
                return error;

            NatsConnection connection = MaterializedConnection(endpoint);
            if (connection == null)
                return "CLOSED " + Mask(endpoint);

            string state;
            try
            {
                switch (connection.ConnectionState)
                {
                    case NatsConnectionState.Open:
                        state = "CONNECTED";
                        break;
                    case NatsConnectionState.Connecting:
                    case NatsConnectionState.Reconnecting:
                        state = "RECONNECTING";
                        break;
                    default:
                        state = "CLOSED";
                        break;
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "NatsUDFConnectionState: reading the connection state failed");
                return "CLOSED " + Mask(endpoint);
            }

            INatsServerInfo info = LiveServerInfo(endpoint);
            if (info != null && info.LameDuckMode)
                state = "LAME_DUCK";

            return state + " " + Mask(endpoint);
        }

        private static object ClusterInfoBody(object host)
        {
            if (!TryResolve(host, out string endpoint, out string error))
                return error;

            List<JObject> replies;
            if (!TrySysPing(endpoint, out replies, out error))
                return error;
            if (replies.Count == 0)
                return NatsUdfCommon.Fail(NoSystemAccount);

            var nodes = new JArray();
            foreach (JObject reply in replies)
            {
                JObject server = reply["server"] as JObject;
                var node = new JObject();
                node["server_name"] = Token(server?["name"]);
                node["server_id"] = Token(server?["id"]);
                node["version"] = Token(server?["version"]);
                node["go"] = Token(server?["go"]);
                node["cluster"] = Token(server?["cluster"]);
                node["jetstream"] = Token(server?["jetstream"]);
                nodes.Add(node);
            }

            var result = new JObject();
            result["server_count"] = replies.Count;
            result["nodes"] = nodes;
            return NatsResultFormatter.BoundCellText(Compact(result));
        }

        private static object ClusterHealthBody(object host)
        {
            if (!TryResolve(host, out string endpoint, out string error))
                return error;

            List<JObject> replies;
            if (!TrySysPing(endpoint, out replies, out error))
                return error;
            if (replies.Count == 0)
                return NatsUdfCommon.Fail(NoSystemAccount);

            var nodes = new JArray();
            foreach (JObject reply in replies)
            {
                JObject server = reply["server"] as JObject;
                JObject statsz = reply["statsz"] as JObject;
                var node = new JObject();
                node["server_name"] = Token(server?["name"]);
                node["mem"] = Token(statsz?["mem"]);
                node["cpu"] = Token(statsz?["cpu"]);
                node["connections"] = Token(statsz?["connections"]);
                node["slow_consumers"] = Token(statsz?["slow_consumers"]);
                // The JetStream block is passed through verbatim: its shape is the
                // server's, and guessing at it would be a lie in a cell.
                node["jetstream"] = statsz?["jetstream"] ?? JValue.CreateNull();
                nodes.Add(node);
            }

            var result = new JObject();
            result["server_count"] = replies.Count;
            result["nodes"] = nodes;
            return NatsResultFormatter.BoundCellText(Compact(result));
        }

        private static object RaftInfoBody(object stream, object host)
        {
            string streamError = NatsUdfCommon.RequireStream(stream);
            if (streamError != null)
                return streamError;
            if (!TryResolve(host, admin: false, requireJetStream: true, out string endpoint, out string error))
                return error;

            string name = TextOf(stream);
            INatsJSContext context;
            if (!TryGetJetStream(endpoint, out context, out error))
                return error;

            try
            {
                INatsJSStream handle = NatsConnectionManager.WaitBoundedResult(
                    context.GetStreamAsync(name).AsTask(), JsTimeoutMs());
                ClusterInfo cluster = handle.Info?.Cluster;
                if (cluster == null || cluster.Replicas == null || cluster.Replicas.Count == 0)
                {
                    return NatsUdfCommon.VerboseError(
                        "JetStream raft info",
                        "stream '" + (name ?? string.Empty) + "' has no Raft group (a standalone server)",
                        "a Raft view needs a JetStream cluster with more than one replica");
                }

                var rows = new List<string[]>(cluster.Replicas.Count);
                foreach (PeerInfo peer in cluster.Replicas)
                {
                    if (peer == null)
                        continue;
                    rows.Add(new[]
                    {
                        peer.Name ?? string.Empty,
                        peer.Current ? "true" : "false",
                        peer.Active.TotalMilliseconds.ToString("0.###", CultureInfo.InvariantCulture),
                        peer.Lag.ToString(CultureInfo.InvariantCulture)
                    });
                }

                return NatsResultFormatter.FormatRowsMatrix(rows);
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream raft info", ex, name);
            }
        }

        private static object StreamPlacementBody(object stream, object host)
        {
            string streamError = NatsUdfCommon.RequireStream(stream);
            if (streamError != null)
                return streamError;
            if (!TryResolve(host, admin: false, requireJetStream: true, out string endpoint, out string error))
                return error;

            string name = TextOf(stream);
            INatsJSContext context;
            if (!TryGetJetStream(endpoint, out context, out error))
                return error;

            try
            {
                INatsJSStream handle = NatsConnectionManager.WaitBoundedResult(
                    context.GetStreamAsync(name).AsTask(), JsTimeoutMs());
                Placement placement = handle.Info?.Config?.Placement;
                if (placement == null
                    || (string.IsNullOrWhiteSpace(placement.Cluster)
                        && (placement.Tags == null || placement.Tags.Count == 0)))
                {
                    return "(no placement constraint)";
                }

                string tags = placement.Tags == null || placement.Tags.Count == 0
                    ? "(any)"
                    : string.Join(",", placement.Tags);
                string cluster = string.IsNullOrWhiteSpace(placement.Cluster) ? "(any)" : placement.Cluster;
                return "cluster=" + cluster + " tags=" + tags;
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream stream placement", ex, name);
            }
        }

        // ==================================================================
        // Gates, bounded waits and error shaping
        // ==================================================================

        /// <summary>Resolves the optional <c>[host]</c> to a concrete endpoint.</summary>
        private static bool TryResolve(object host, out string endpoint, out string error)
        {
            return TryResolve(host, admin: false, requireJetStream: false, out endpoint, out error);
        }

        /// <summary>
        /// Endpoint resolution plus the optional admin and JetStream gates. The
        /// JetStream gate is the config flag plus the live per-endpoint probe, so
        /// a server started without <c>-js</c> gets the exact
        /// "JetStream is not enabled on this server" text instead of a hang.
        /// </summary>
        private static bool TryResolve(object host, bool admin, bool requireJetStream, out string endpoint, out string error)
        {
            endpoint = null;
            error = null;

            try
            {
                endpoint = NatsRuntime.ResolveEndpointForDispatch(host);
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("endpoint", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }

            if (endpoint == null)
            {
                // A non-text host is a bad argument; an unresolvable text host is
                // a dangling reference, reported as #CONFIG instead of silently
                // connecting to localhost (DESIGN.md section 8.1).
                error = IsBlank(host) || host is string
                    ? NatsUdfCommon.ConfigError
                    : NatsUdfCommon.Fail(NatsUdfCommon.ErrorHostNotText);
                return false;
            }

            if (admin)
            {
                string gate = NatsUdfCommon.RequireAdminWrites(AppConfig.AllowAdminWrites);
                if (gate != null)
                {
                    error = gate;
                    return false;
                }
            }

            if (!requireJetStream)
                return true;

            string jetStream = NatsUdfCommon.RequireJetStream(AppConfig.Current.JetStreamEnabled);
            if (jetStream != null)
            {
                error = jetStream;
                return false;
            }

            try
            {
                NatsConnection connection = NatsRuntime.Connections.GetConnection(endpoint);
                if (connection.ServerInfo == null)
                    NatsConnectionManager.WaitBounded(connection.PingAsync(CancellationToken.None).AsTask(), JsTimeoutMs());
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("connection", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }

            if (!NatsRuntime.Connections.IsJetStreamEnabled(endpoint))
            {
                // Two different causes must not share one message (same rule as
                // the JetStream/KV/Object Store probes): a server that answered
                // but has no JetStream is a config problem, while an endpoint
                // that never answered is a connectivity problem. Saying
                // "JetStream is not enabled" for an unreachable host sends the
                // user looking in the wrong place.
                if (!NatsRuntime.Connections.HasConnection(endpoint))
                    error = NatsUdfCommon.VerboseError(
                        "connection", "no server answered at the endpoint",
                        "check that the server is running and that the URL/token are right");
                else
                    error = NatsUdfCommon.RequireJetStream(false); // exact frozen text
                return false;
            }

            return true;
        }

        private static bool TryGetJetStream(string endpoint, out INatsJSContext context, out string error)
        {
            context = null;
            error = null;
            try
            {
                context = NatsRuntime.Connections.GetJetStream(endpoint);
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("JetStream", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }

            if (context == null)
            {
                error = NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Drains <c>$SYS.REQ.SERVER.PING</c> for up to the response budget and
        /// returns one parsed reply per answering server. An endpoint without the
        /// system account answers nothing, which the caller turns into the NAMED
        /// error <see cref="NoSystemAccount"/> - never an empty, healthy-looking
        /// result.
        /// </summary>
        private static bool TrySysPing(string endpoint, out List<JObject> replies, out string error)
        {
            replies = new List<JObject>();
            error = null;

            int bound = ResponseTimeoutMs();
            var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(bound));
            try
            {
                NatsConnection connection = NatsRuntime.Connections.GetConnection(endpoint);
                var opts = new NatsSubOpts
                {
                    Timeout = TimeSpan.FromMilliseconds(bound),
                    ThrowIfNoResponders = true
                };

                Task<List<JObject>> task = Task.Run(async () =>
                {
                    var nodes = new List<JObject>();
                    await foreach (NatsMsg<string> reply in connection.RequestManyAsync<string, string>(
                        subject: SysPingSubject,
                        data: string.Empty,
                        replyOpts: opts,
                        cancellationToken: cts.Token).ConfigureAwait(false))
                    {
                        if (string.IsNullOrWhiteSpace(reply.Data))
                            continue;
                        try
                        {
                            nodes.Add(JObject.Parse(reply.Data));
                        }
                        catch (JsonException)
                        {
                            // A non-JSON reply is not a node: skip it rather than
                            // fail the whole query.
                        }

                        if (nodes.Count >= MaxSysNodes)
                            break;
                    }

                    return nodes;
                });

                List<JObject> collected = NatsConnectionManager.WaitBoundedResult(task, bound + 500);
                replies = collected ?? new List<JObject>();
                return true;
            }
            catch (NatsNoRespondersException)
            {
                // The documented "no system account" answer.
                return true;
            }
            catch (NatsNoReplyException)
            {
                return true;
            }
            catch (TimeoutException)
            {
                error = NatsUdfCommon.VerboseError(
                    "cluster query",
                    "no answer within " + bound.ToString(CultureInfo.InvariantCulture) + "ms",
                    NatsUdfCommon.ErrorHint("timed out"));
                return false;
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("cluster query", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }
            finally
            {
                try
                {
                    cts.Cancel();
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "NatsUdfStatus: cancelling the cluster query failed");
                }
            }
        }

        private static string MapJsError(string what, Exception ex, string stream)
        {
            if (ex is TimeoutException)
            {
                return NatsUdfCommon.VerboseError(
                    what,
                    "timed out after " + JsTimeoutMs().ToString(CultureInfo.InvariantCulture) + "ms",
                    NatsUdfCommon.ErrorHint("timed out"));
            }

            if (TryApiError(ex, out ApiError api))
            {
                string cause = ApiCause(api);
                string lower = cause.ToLowerInvariant();
                if (lower.Contains("jetstream") && lower.Contains("not enabled"))
                    return NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
                if (api.Code == 404 && stream != null)
                    return NatsUdfCommon.VerboseError(what, NatsUdfCommon.StreamNotFound(stream), NatsUdfCommon.ErrorHint("stream not found"));
                return NatsUdfCommon.VerboseError(what, cause, NatsUdfCommon.ErrorHint(cause));
            }

            string message = NatsUdfAsync.SafeMessage(ex);
            string text = (message ?? string.Empty).ToLowerInvariant();
            if (text.Contains("jetstream") && text.Contains("not enabled"))
                return NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
            if (stream != null && text.Contains("stream") && text.Contains("not found"))
                return NatsUdfCommon.VerboseError(what, NatsUdfCommon.StreamNotFound(stream), NatsUdfCommon.ErrorHint("stream not found"));
            return NatsUdfCommon.VerboseError(what, message, NatsUdfCommon.ErrorHint(message));
        }

        /// <summary>Finds the server's structured error in an (aggregated) exception chain.</summary>
        private static bool TryApiError(Exception ex, out ApiError api)
        {
            api = null;
            if (ex == null)
                return false;
            if (ex is NatsJSApiException apiException && apiException.Error != null)
            {
                api = apiException.Error;
                return true;
            }
            if (ex is AggregateException aggregate)
            {
                foreach (Exception inner in aggregate.Flatten().InnerExceptions)
                {
                    if (TryApiError(inner, out api))
                        return true;
                }
            }
            return TryApiError(ex.InnerException, out api);
        }

        private static string ApiCause(ApiError error)
        {
            string description = error.Description;
            string code = "code " + error.Code.ToString(CultureInfo.InvariantCulture)
                + ", err_code " + error.ErrCode.ToString(CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(description) ? code : description + " (" + code + ")";
        }

        // ==================================================================
        // Client access
        // ==================================================================

        /// <summary>
        /// The live <see cref="INatsServerInfo"/> for an endpoint, or null while
        /// the connection has not received INFO yet. It never forces a connect and
        /// never blocks: the monitor reads whatever the client already holds.
        /// </summary>
        private static INatsServerInfo LiveServerInfo(string endpoint)
        {
            try
            {
                NatsConnection connection = MaterializedConnection(endpoint);
                return connection?.ServerInfo;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "NatsUDFStatus: server info unavailable");
                return null;
            }
        }

        /// <summary>
        /// The cached client for an endpoint. A status read must not block on a
        /// socket, and <see cref="NatsConnectionManager.GetConnection"/> only
        /// allocates the client (the socket opens behind the first await), so this
        /// is cheap and safe.
        /// </summary>
        private static NatsConnection MaterializedConnection(string endpoint)
        {
            try
            {
                return NatsRuntime.Connections.GetConnection(endpoint);
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "NatsUDFStatus: connection unavailable");
                return null;
            }
        }

        /// <summary>
        /// Reads the client's counters through the single reflected accessor (the
        /// accessor is internal in 3.3.0; see <see cref="StatsMethod"/>). False
        /// when the member is gone or the read fails, so the caller reports
        /// "no counters" instead of faulting a cell.
        /// </summary>
        private static bool TryReadStats(NatsConnection connection, out NatsStats stats)
        {
            stats = default;
            try
            {
                if (StatsMethod == null || connection == null)
                    return false;
                object value = StatsMethod.Invoke(connection, null);
                if (!(value is NatsStats read))
                    return false;
                stats = read;
                return true;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "NatsUDFStatus: reading the client counters failed");
                return false;
            }
        }

        private static MethodInfo ResolveStatsMethod()
        {
            try
            {
                return typeof(NatsConnection).GetMethod(
                    "GetStats", BindingFlags.Instance | BindingFlags.NonPublic);
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "NatsUDFStatus: the client counters accessor was not found");
                return null;
            }
        }

        // ==================================================================
        // Guards, bounds and small helpers
        // ==================================================================

        /// <summary>
        /// Runs a status body, replacing a teardown/reload exception (the runtime
        /// refuses to resurrect its managers) with the neutral fallback so an
        /// exception is never surfaced into a worksheet cell. The fallback keeps
        /// the same type as the normal result, so the cell stays well formed.
        /// </summary>
        private static T Guarded<T>(Func<T> read, T fallback)
        {
            try
            {
                return read();
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "NatsUDFStatus: a status read failed (returning the neutral value)");
                return fallback;
            }
        }

        private static int ResponseTimeoutMs()
        {
            try
            {
                return NatsConnectionManager.ResponseTimeoutMs(NatsPool.Udf);
            }
            catch
            {
                return 1000;
            }
        }

        private static int JsTimeoutMs()
        {
            int configured;
            try
            {
                configured = AppConfig.Current.JetStreamTimeoutMs;
            }
            catch
            {
                configured = 0;
            }

            int floor = ResponseTimeoutMs();
            return configured < floor ? floor : configured;
        }

        /// <summary>The masked form of an endpoint or URL, safe for a cell.</summary>
        private static string Mask(string text)
        {
            return AppConfig.MaskNatsEndpoint(text);
        }

        /// <summary>A 1x1 matrix holding a single value (a matrix function must
        /// never return a blank range).</summary>
        private static object[,] SingleCell(object value)
        {
            var matrix = new object[1, 1];
            matrix[0, 0] = value;
            return matrix;
        }

        /// <summary>A deep copy of a JSON token, or JSON null when absent.</summary>
        private static JToken Token(JToken value)
        {
            return value == null || value.Type == JTokenType.Undefined
                ? JValue.CreateNull()
                : value.DeepClone();
        }

        private static string Compact(JObject value)
        {
            return value.ToString(Formatting.None);
        }

        /// <summary>The invariant text of a cell value, trimmed; null when it cannot be read.</summary>
        private static string TextOf(object value)
        {
            try
            {
                return NatsUdfCommon.ToNatsString(value)?.Trim();
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        private static bool IsBlank(object value)
        {
            if (value == null || value is ExcelMissing || value is ExcelEmpty)
                return true;
            return value is string text && string.IsNullOrWhiteSpace(text);
        }
    }
}
