using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using NATS.Client.Core;
using NATS.Client.JetStream;
using NATS.Net;
using NLog;

namespace NATSExcel
{
    /// <summary>
    /// Connection family a call belongs to. Kept from the RedisExcel shape so the
    /// status functions keep their meaning; on NATS a single connection per
    /// endpoint multiplexes every family (DESIGN.md Â§3), so this only selects the
    /// response-timeout budget today.
    /// </summary>
    public enum NatsPool
    {
        RtdData,
        RtdSub,
        Udf
    }

    /// <summary>
    /// Immutable copy of the facts the server advertised for one connection
    /// (<see cref="INatsServerInfo"/>). Callers get a snapshot instead of the live
    /// object so a reconnect that replaces the server info cannot make a cell read
    /// a half-updated instance. <see cref="Empty"/> is returned when there is no
    /// server info yet (never connected) so callers never have to null-check.
    /// </summary>
    public sealed class NatsServerSnapshot
    {
        /// <summary>The "no server info yet" value; <see cref="HasInfo"/> is false.</summary>
        public static readonly NatsServerSnapshot Empty = new NatsServerSnapshot();

        private NatsServerSnapshot()
        {
            ConnectUrls = new string[0];
        }

        internal NatsServerSnapshot(
            string name, string version, string cluster,
            string[] connectUrls, bool jetStreamAvailable, bool lameDuckMode)
        {
            Name = name;
            Version = version;
            Cluster = cluster;
            ConnectUrls = connectUrls ?? new string[0];
            JetStreamAvailable = jetStreamAvailable;
            LameDuckMode = lameDuckMode;
            HasInfo = true;
        }

        /// <summary>True when this came from a real server INFO; false for <see cref="Empty"/>.</summary>
        public bool HasInfo { get; }

        /// <summary>The server's own name (<c>server_name</c>), may be blank.</summary>
        public string Name { get; }

        /// <summary>The nats-server version string.</summary>
        public string Version { get; }

        /// <summary>The cluster name the server announces; blank for a standalone server.</summary>
        public string Cluster { get; }

        /// <summary>The other nodes the server advertises (its discovery pool).</summary>
        public IReadOnlyList<string> ConnectUrls { get; }

        /// <summary>Whether JetStream is enabled on this server.</summary>
        public bool JetStreamAvailable { get; }

        /// <summary>Whether the node is draining for a restart.</summary>
        public bool LameDuckMode { get; }
    }

    /// <summary>
    /// The single place that creates, caches and tears down NATS connections.
    /// There is one <see cref="NatsConnection"/> per resolved endpoint: the NATS
    /// client multiplexes every subject and queue group over a single socket and
    /// reconnects internally, so the RedisExcel per-pool cache is deliberately
    /// absent (DESIGN.md Â§3). What remains is a bounded cache: an endpoint that is
    /// not live and that no subscription is using may be evicted once the cache is
    /// oversized or after it has been down for the whole TTL.
    ///
    /// <para><b>No connect happens on the Excel thread.</b> <see cref="GetConnection"/>
    /// only constructs the client object (a cheap allocation); the NATS client
    /// connects on the first await. A fire-and-forget warm-up on the thread pool
    /// starts that connection so <see cref="IsJetStreamEnabled"/> and
    /// <see cref="GetServerInfo"/> can answer as soon as INFO arrives, but the
    /// caller never waits for it.</para>
    ///
    /// <para><b>Why the bounded wait exists.</b> NATS.Net is async-only, while an
    /// Excel UDF must return a value before returning. <see cref="WaitBounded"/>
    /// and <see cref="WaitBoundedResult{T}"/> are the only bounded-blocking
    /// primitives in the codebase; they cap how long the Excel thread can ever be
    /// held and are the sole intended users of a direct task wait.</para>
    ///
    /// <para><b>Why the shutdown fence is checked before creating.</b> Once
    /// <see cref="Shutdown"/> starts it clears the caches; a create that raced it
    /// could publish a fresh entry into an already-cleared cache and leak a live
    /// socket. The flag is therefore re-checked after the Lazy materializes so
    /// such a connection is disposed and the caller fails instead of receiving a
    /// dying client.</para>
    /// </summary>
    public sealed class NatsConnectionManager : IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // One client per resolved endpoint. The key is the endpoint string as the
        // config resolver produced it; it may contain URL credentials and is
        // therefore never logged directly - every log line uses MaskedEndpoint.
        private readonly ConcurrentDictionary<string, Lazy<NatsConnection>> _connections =
            new ConcurrentDictionary<string, Lazy<NatsConnection>>(StringComparer.Ordinal);

        // JetStream contexts are thin wrappers around a connection; cached so a
        // volatile UDF does not allocate one per call. Dropped together with the
        // connection it points at (eviction or shutdown).
        private readonly ConcurrentDictionary<string, Lazy<INatsJSContext>> _jetStream =
            new ConcurrentDictionary<string, Lazy<INatsJSContext>>(StringComparer.Ordinal);

        // Set at the start of Shutdown and never cleared: a connect that races
        // shutdown must not leave an un-disposed connection behind.
        private volatile bool _shutdown;

        // One-owner gate for Shutdown: concurrent callers return immediately.
        private int _shutdownStarted;

        // A worksheet argument can produce arbitrary endpoint strings, and every
        // unreachable endpoint keeps a client that is trying to reconnect. Past
        // the cap, non-live entries are evicted; a live connection is never
        // evicted, and neither is an endpoint a subscription is using.
        private const int MaxCachedConnections = 512;

        // An endpoint that has stayed non-live for this long is dropped so a
        // session full of mistyped hosts cannot grow forever. Live endpoints are
        // never touched. Driven from the call path, never a timer.
        private const int DownEndpointTtlMs = 300000; // 5 minutes
        private const int DownSweepIntervalMs = 60000;
        private readonly ConcurrentDictionary<string, long> _downSinceTicks =
            new ConcurrentDictionary<string, long>(StringComparer.Ordinal);
        private long _lastDownSweepTicks;

        // Live-consumer protection: an endpoint with an active subscription is
        // never evicted (disposing its connection would silently kill the topic).
        private Func<string, bool> _evictionProtection;

        // Bounded teardown: Shutdown must not wait on the Excel thread forever.
        private const int ConnectionDisposeTimeoutMs = 2000;

        /// <summary>Total budget for a whole shutdown loop. Beyond it the remaining
        /// connections are left to the process exit, so closing Excel cannot block
        /// for minutes when a workbook referenced many endpoints.</summary>
        private const int ShutdownBudgetMs = 5000;

        // Bounds for the response timeout (PLAN 1.4: configured timeout + 500ms,
        // clamped to a sane floor/ceiling).
        private const int ResponseTimeoutMarginMs = 500;
        private const int MinResponseTimeoutMs = 250;
        private const int MaxResponseTimeoutMs = 60000;

        // Bounds for the first-connect timeout (the background warm-up only).
        private const int MinConnectTimeoutMs = 250;
        private const int MaxConnectTimeoutMs = 30000;

        /// <summary>
        /// The masked form of an endpoint, safe for logs and cells. The cache key
        /// itself may carry credentials and must never be logged; always log this.
        /// </summary>
        public static string MaskedEndpoint(string endpoint)
        {
            return AppConfig.MaskNatsEndpoint(endpoint);
        }

        /// <summary>
        /// Registers the live-subscriber veto used by eviction and the down sweep.
        /// A host this predicate reports as protected is never evicted, even when
        /// it sits past the cache cap or has been down for the whole TTL.
        /// </summary>
        public void SetEvictionProtection(Func<string, bool> isEndpointProtected)
        {
            _evictionProtection = isEndpointProtected;
        }

        /// <summary>
        /// Returns the cached connection for a resolved endpoint, creating it on
        /// first use. Construction never connects (and therefore never blocks): the
        /// NATS client opens the socket on the first await, and a background warm-up
        /// on the thread pool starts that. Throws
        /// <see cref="InvalidOperationException"/> once <see cref="Shutdown"/> ran.
        /// </summary>
        public NatsConnection GetConnection(string endpoint)
        {
            // Shutdown fence before creating: a cached connection is already being
            // closed by Shutdown, so it must never be handed out once the flag is set.
            if (_shutdown)
                throw ShuttingDown();
            if (string.IsNullOrWhiteSpace(endpoint))
                throw new ArgumentException("endpoint is required", nameof(endpoint));

            string key = endpoint.Trim();
            // Lazy (ExecutionAndPublication) makes concurrent callers share one
            // client even though GetOrAdd may invoke the factory more than once.
            var lazy = _connections.GetOrAdd(key, k => new Lazy<NatsConnection>(
                () => CreateConnection(k), LazyThreadSafetyMode.ExecutionAndPublication));

            NatsConnection connection;
            try
            {
                connection = lazy.Value;
            }
            catch
            {
                // The factory failed (bad endpoint). Drop the placeholder so the
                // next call retries instead of caching the failure forever.
                // Conditional by value identity: never removes an entry a concurrent
                // caller has just replaced.
                TryRemoveEntry(key, lazy);
                throw;
            }

            if (_shutdown)
            {
                // Shutdown raced this create: dispose the fresh client and fail the
                // caller instead of returning a dying connection.
                DisposeConnection(connection, key, "created during shutdown");
                _connections.TryRemove(key, out _);
                _jetStream.TryRemove(key, out _);
                throw ShuttingDown();
            }

            if (IsLive(connection))
                _downSinceTicks.TryRemove(key, out _); // healthy: clear the down clock
            if (_connections.Count > MaxCachedConnections)
                EvictIdleConnections(key);
            SweepDownEndpointsIfDue();
            return connection;
        }

        /// <summary>
        /// Returns the cached JetStream context for a resolved endpoint, creating
        /// it (and the connection it wraps) on first use. Cheap: the context is a
        /// thin wrapper, no extra socket.
        /// </summary>
        public INatsJSContext GetJetStream(string endpoint)
        {
            if (_shutdown)
                throw ShuttingDown();
            if (string.IsNullOrWhiteSpace(endpoint))
                throw new ArgumentException("endpoint is required", nameof(endpoint));

            string key = endpoint.Trim();
            var lazy = _jetStream.GetOrAdd(key, k => new Lazy<INatsJSContext>(
                () => GetConnection(k).CreateJetStreamContext(),
                LazyThreadSafetyMode.ExecutionAndPublication));
            try
            {
                var context = lazy.Value;
                // Post-materialization re-check, mirroring GetConnection: a
                // Shutdown that ran while the factory was building would hand back
                // a context bound to a disposed connection.
                if (_shutdown)
                {
                    _jetStream.TryRemove(key, out _);
                    throw ShuttingDown();
                }
                return context;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch
            {
                _jetStream.TryRemove(key, out _);
                throw;
            }
        }

        /// <summary>
        /// Whether the server behind the endpoint has JetStream enabled. Reads the
        /// live <see cref="INatsServerInfo"/> so a reconnect re-probes for free;
        /// false while there is no server info yet and never throws. Callers turn
        /// false into the exact "JetStream is not enabled on this server" error
        /// instead of issuing a call that would hang.
        /// </summary>
        public bool IsJetStreamEnabled(string endpoint)
        {
            if (_shutdown || !TryGetMaterializedConnection(endpoint, out var connection))
                return false;
            try
            {
                var info = connection.ServerInfo;
                return info != null && info.JetStreamAvailable;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "GetServerInfo: JetStream probe failed endpoint=" + MaskedEndpoint(endpoint));
                return false;
            }
        }

        /// <summary>
        /// True when the endpoint has a materialized connection whose server INFO
        /// has arrived. Lets a caller tell "the server never answered" apart from
        /// "the server answered and has no JetStream", so the two produce
        /// different errors instead of one misleading one.
        /// </summary>
        public bool HasConnection(string endpoint)
        {
            if (_shutdown || !TryGetMaterializedConnection(endpoint, out var connection))
                return false;
            try
            {
                return connection.ServerInfo != null;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// Snapshot of the advertised server facts, or
        /// <see cref="NatsServerSnapshot.Empty"/> when the endpoint has no
        /// materialized connection or has not received INFO yet. Never null and
        /// never throws.
        /// </summary>
        public NatsServerSnapshot GetServerInfo(string endpoint)
        {
            if (_shutdown || !TryGetMaterializedConnection(endpoint, out var connection))
                return NatsServerSnapshot.Empty;
            return SnapshotFrom(connection, endpoint);
        }

        /// <summary>
        /// <see cref="GetServerInfo"/> in the frozen PLAN Â§1.4 shape: false when no
        /// server info is available.
        /// </summary>
        public bool TryGetServerInfo(string endpoint, out NatsServerSnapshot info)
        {
            info = GetServerInfo(endpoint);
            return info.HasInfo;
        }

        /// <summary>
        /// Number of materialized (Lazy-created) connections across the manager;
        /// zero while the shutdown fence is set. A cached entry whose Lazy never ran
        /// holds nothing and is not counted, and a connection Shutdown is currently
        /// disposing is covered by the fence (the client's own disposed flag is
        /// internal), so the count never includes one that is being torn down.
        /// </summary>
        public int LiveConnectionCount()
        {
            return _shutdown ? 0 : CountMaterialized(_connections);
        }

        /// <summary>
        /// Live connections for the RTD family. A single connection per endpoint
        /// multiplexes RTD and UDF traffic (DESIGN.md Â§3), so this currently equals
        /// <see cref="LiveConnectionCount"/>; it exists for the RTD status surface.
        /// </summary>
        public int LiveRtdConnectionCount()
        {
            return _shutdown ? 0 : CountMaterialized(_connections);
        }

        /// <summary>
        /// Live connections for the UDF family. See
        /// <see cref="LiveRtdConnectionCount"/>: one connection per endpoint serves
        /// every family, so this equals <see cref="LiveConnectionCount"/>.
        /// </summary>
        public int LiveUdfConnectionCount()
        {
            return _shutdown ? 0 : CountMaterialized(_connections);
        }

        /// <summary>
        /// True when the task completed (success, fault or cancellation) within the
        /// bound; false on timeout. This is one of the two bounded-blocking
        /// primitives of the codebase - do not call <c>Task.Wait</c> anywhere else.
        /// </summary>
        // The synchronous waits below are the whole point of this helper: an Excel
        // UDF must return in-cell, so the blocking wait is deliberate and bounded.
        // VSTHRD002 is suppressed here and only here (PLAN V7.3 scans for unbounded
        // waits elsewhere).
#pragma warning disable VSTHRD002 // Deliberate bounded wait; see the XML doc above.
        internal static bool WaitBounded(Task task, int timeoutMs)
        {
            if (task == null)
                return true;
            if (timeoutMs < 0)
                timeoutMs = 0;
            try
            {
                return task.Wait(timeoutMs);
            }
            catch (AggregateException)
            {
                // Completed with a fault or cancellation: that is still "completed".
                return true;
            }
        }

        /// <summary>
        /// Waits for a result with the same bound as <see cref="WaitBounded"/>:
        /// throws <see cref="TimeoutException"/> on timeout, otherwise returns the
        /// value or rethrows the task's own exception (unwrapped).
        /// </summary>
        internal static T WaitBoundedResult<T>(Task<T> task, int timeoutMs)
        {
            if (task == null)
                throw new ArgumentNullException(nameof(task));
            if (timeoutMs < 0)
                timeoutMs = 0;
            try
            {
                if (!task.Wait(timeoutMs))
                    throw new TimeoutException("the NATS operation did not complete within " + timeoutMs + "ms");
            }
            catch (AggregateException)
            {
                // Faulted/cancelled within the bound; fall through to surface the
                // real exception below.
            }

            // Only reached when the task completed: safe (and intended) direct read.
            return task.GetAwaiter().GetResult();
        }

#pragma warning restore VSTHRD002

        /// <summary>
        /// The response timeout for a pool: the configured timeout plus a margin for
        /// connection setup, clamped to a sane floor/ceiling so no formula can hold
        /// the Excel thread unbounded.
        /// </summary>
        /// <remarks>
        /// PLAN Â§1.4 gives the UDF section its own timeout, but <c>Config.cs</c>
        /// does not expose it yet (F1.4); both pools therefore read
        /// <c>Default.timeout</c> until that field lands.
        /// </remarks>
        internal static int ResponseTimeoutMs(NatsPool pool)
        {
            int configured;
            try
            {
                configured = AppConfig.Current.TimeoutMs;
            }
            catch
            {
                configured = 1000;
            }

            if (configured < 0)
                configured = 0;
            long total = (long)configured + ResponseTimeoutMarginMs;
            if (total < MinResponseTimeoutMs)
                total = MinResponseTimeoutMs;
            if (total > MaxResponseTimeoutMs)
                total = MaxResponseTimeoutMs;
            return (int)total;
        }

        /// <summary>
        /// Disposes every connection with a bounded wait and clears the caches.
        /// Idempotent and one-owner: concurrent callers after the first return
        /// immediately.
        /// </summary>
        public void Shutdown()
        {
            if (Interlocked.Exchange(ref _shutdownStarted, 1) != 0)
                return; // another caller already owns the teardown

            // Set before anything is closed/cleared so a racing create either sees
            // the flag (and refuses) or is caught by the post-create re-check.
            _shutdown = true;
            _downSinceTicks.Clear();

            // A TOTAL budget, not a per-connection one. Disposing N connections
            // with a 2 s per-connection bound could otherwise hold the Excel
            // thread (this runs from AutoClose) for minutes with many endpoints.
            // The remaining time is passed to each dispose; the loop bails out
            // when the budget is spent and leaves the rest to the process exit.
            var budget = System.Diagnostics.Stopwatch.StartNew();
            foreach (var kv in _connections)
            {
                var lazy = kv.Value;
                if (!lazy.IsValueCreated)
                {
                    // Nothing materialized. Forcing Lazy.Value could block behind an
                    // in-flight factory; the racing create re-checks _shutdown and
                    // disposes its own connection, so there is nothing to close here.
                    _connections.TryRemove(kv.Key, out _);
                    continue;
                }

                int remaining = ShutdownBudgetMs - (int)budget.ElapsedMilliseconds;
                if (remaining <= 0)
                {
                    logger.Warn("Shutdown: the total budget elapsed; " + _connections.Count
                        + " connection(s) left to the process exit");
                    break;
                }

                try
                {
                    DisposeConnection(lazy.Value, kv.Key, "Shutdown", remaining);
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "Shutdown: error disposing connection endpoint=" + MaskedEndpoint(kv.Key));
                }

                _connections.TryRemove(kv.Key, out _);
            }

            _connections.Clear();
            _jetStream.Clear();
            logger.Info("NatsConnectionManager: shutdown complete");
        }

        /// <summary>Delegates to the idempotent <see cref="Shutdown"/>.</summary>
        public void Dispose()
        {
            Shutdown();
        }

        private NatsConnection CreateConnection(string endpoint)
        {
            if (_shutdown)
                throw ShuttingDown();

            var options = BuildOptions(endpoint);
            var connection = new NatsConnection(options);
            AttachObservers(connection, endpoint);
            StartWarmUp(connection, endpoint);
            logger.Info("NatsConnect: opening connection to " + MaskedEndpoint(endpoint));
            return connection;
        }

        /// <summary>
        /// Builds the client options: the resolved URL list, the connection name
        /// (the add-in version, shown as "dev" for local builds), the mapped auth
        /// and the bounded first-connect timeout.
        /// </summary>
        private static NatsOpts BuildOptions(string endpoint)
        {
            var config = AppConfig.Current;
            var auth = config.Auth;

            int connectTimeoutMs = Clamp(
                config.TimeoutMs <= 0 ? 1000 : config.TimeoutMs,
                MinConnectTimeoutMs, MaxConnectTimeoutMs);

            var options = NatsOpts.Default with
            {
                Url = endpoint,
                Name = "NATSExcel " + AppConfig.Version,
                AuthOpts = BuildAuth(auth),
                ConnectTimeout = TimeSpan.FromMilliseconds(connectTimeoutMs),
            };

            // mTLS is a config flag today (Config.cs reads only "tls"; the cert
            // fields of DESIGN.md Â§8 are planned, PLAN F1.5b). Require makes the
            // intent explicit instead of the opportunistic Auto mode.
            if (auth != null && auth.Tls)
                options = options with { TlsOpts = NatsTlsOpts.Default with { Mode = TlsMode.Require } };

            return options;
        }

        /// <summary>
        /// Maps <see cref="AppConfig.Auth"/> onto <c>NatsAuthOpts</c> with the
        /// DESIGN.md Â§8 precedence: creds file, then JWT + seed, then seed, then
        /// token, then user + password. The NKey seed may be the seed text or a path
        /// to a <c>.nk</c> file; <see cref="AppConfig.ResolveSeed"/> reads the file
        /// (golden rule 11).
        /// </summary>
        private static NatsAuthOpts BuildAuth(AppConfig.Auth auth)
        {
            if (auth == null)
                return new NatsAuthOpts();

            if (!string.IsNullOrWhiteSpace(auth.CredsFile))
                return new NatsAuthOpts { CredsFile = auth.CredsFile.Trim() };
            if (!string.IsNullOrWhiteSpace(auth.Jwt))
                return new NatsAuthOpts { Jwt = auth.Jwt.Trim(), Seed = AppConfig.ResolveSeed(auth.NkeySeed) };
            if (!string.IsNullOrWhiteSpace(auth.NkeySeed))
                return new NatsAuthOpts { Seed = AppConfig.ResolveSeed(auth.NkeySeed) };
            if (!string.IsNullOrWhiteSpace(auth.Token))
                return new NatsAuthOpts { Token = auth.Token.Trim() };
            if (!string.IsNullOrWhiteSpace(auth.User) || !string.IsNullOrWhiteSpace(auth.Password))
                return new NatsAuthOpts { Username = auth.User, Password = auth.Password };
            return new NatsAuthOpts();
        }

        /// <summary>
        /// Starts the initial connect on the thread pool. This is what keeps the
        /// Excel thread free: the caller got the client object and moved on, and the
        /// socket is opened behind it. Failures are logged at Info/Debug (an offline
        /// endpoint is not an error at this layer) and the connection stays cached so
        /// the next real await retries.
        /// </summary>
        private static void StartWarmUp(NatsConnection connection, string endpoint)
        {
            string masked = MaskedEndpoint(endpoint);
            _ = Task.Run(async () =>
            {
                try
                {
                    await connection.ConnectAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    try
                    {
                        logger.Info("NatsConnect: initial connection to " + masked + " not available: "
                                    + AppConfig.MaskNatsEndpoint(ex.Message));
                    }
                    catch
                    {
                        // Never let a logging failure cross a client callback.
                    }
                }
            });
        }

        /// <summary>
        /// Subscribes to the client's reconnect events (the client exposes them, so
        /// the fallback state-poll path is unnecessary) and logs state changes with
        /// the endpoint masked. Every handler catches its own exceptions: it runs on
        /// the client's event task.
        /// </summary>
        private static void AttachObservers(NatsConnection connection, string endpoint)
        {
            string masked = MaskedEndpoint(endpoint);

            connection.ConnectionOpened += (sender, args) =>
            {
                try
                {
                    // Logged at Info: this covers both the initial connect and every
                    // successful reconnect (the client re-runs this event each time).
                    logger.Info("NatsConnect: connected " + masked + ArgSuffix(args));
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "NatsConnect: ConnectionOpened handler error endpoint=" + masked);
                }

                return default;
            };

            connection.ConnectionDisconnected += (sender, args) =>
            {
                try
                {
                    logger.Info("NatsConnect: disconnected " + masked + ArgSuffix(args)
                                + " (the client reconnects in the background)");
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "NatsConnect: ConnectionDisconnected handler error endpoint=" + masked);
                }

                return default;
            };

            connection.ReconnectFailed += (sender, args) =>
            {
                try
                {
                    logger.Debug("NatsConnect: reconnect attempt failed " + masked + ArgSuffix(args));
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "NatsConnect: ReconnectFailed handler error endpoint=" + masked);
                }

                return default;
            };

            connection.LameDuckModeActivated += (sender, args) =>
            {
                try
                {
                    logger.Info("NatsConnect: server entered lame duck mode " + masked + ArgSuffix(args));
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "NatsConnect: LameDuckMode handler error endpoint=" + masked);
                }

                return default;
            };

            connection.ServerError += (sender, args) =>
            {
                try
                {
                    logger.Debug("NatsConnect: server error " + masked + ": "
                                 + AppConfig.MaskNatsEndpoint(args != null ? args.Error : null));
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "NatsConnect: ServerError handler error endpoint=" + masked);
                }

                return default;
            };
        }

        private static string ArgSuffix(NatsEventArgs args)
        {
            if (args == null || string.IsNullOrEmpty(args.Message))
                return string.Empty;
            return " (" + AppConfig.MaskNatsEndpoint(args.Message) + ")";
        }

        /// <summary>
        /// True while the client is connected or actively trying to be: such an
        /// endpoint is never evicted.
        /// </summary>
        private static bool IsLive(NatsConnection connection)
        {
            try
            {
                var state = connection.ConnectionState;
                return state == NatsConnectionState.Open
                       || state == NatsConnectionState.Connecting
                       || state == NatsConnectionState.Reconnecting;
            }
            catch
            {
                return false;
            }
        }

        private bool TryGetMaterializedConnection(string endpoint, out NatsConnection connection)
        {
            connection = null;
            if (_shutdown || string.IsNullOrWhiteSpace(endpoint))
                return false;
            if (!_connections.TryGetValue(endpoint.Trim(), out var lazy) || !lazy.IsValueCreated)
                return false; // never force a connect just to answer a status question
            try
            {
                connection = lazy.Value;
                return true;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "GetServerInfo: connection not materialized endpoint=" + MaskedEndpoint(endpoint));
                return false;
            }
        }

        private static NatsServerSnapshot SnapshotFrom(NatsConnection connection, string endpoint)
        {
            try
            {
                var info = connection.ServerInfo;
                if (info == null)
                    return NatsServerSnapshot.Empty;
                string[] urls = info.ClientConnectUrls == null
                    ? new string[0]
                    : (string[])info.ClientConnectUrls.Clone();
                return new NatsServerSnapshot(
                    info.Name, info.Version, info.Cluster, urls,
                    info.JetStreamAvailable, info.LameDuckMode);
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "GetServerInfo: server info unavailable endpoint=" + MaskedEndpoint(endpoint));
                return NatsServerSnapshot.Empty;
            }
        }

        private static int CountMaterialized(ConcurrentDictionary<string, Lazy<NatsConnection>> dictionary)
        {
            int total = 0;
            foreach (var kv in dictionary)
            {
                if (kv.Value.IsValueCreated)
                    total++;
            }
            return total;
        }

        private bool IsProtected(string endpoint)
        {
            var protection = _evictionProtection;
            if (protection == null)
                return false;
            try
            {
                return protection(endpoint);
            }
            catch (Exception ex)
            {
                // A misbehaving veto must not break the call path or force an eviction.
                logger.Debug(ex, "eviction protection callback failed endpoint=" + MaskedEndpoint(endpoint));
                return false;
            }
        }

        /// <summary>
        /// Evicts non-live, unprotected endpoints once the cache is past the cap.
        /// A connected/reconnecting endpoint and an endpoint with a live
        /// subscription are never touched.
        /// </summary>
        private void EvictIdleConnections(string keepEndpoint)
        {
            if (_connections.Count <= MaxCachedConnections)
                return;

            // A total budget: this runs on the Excel thread (GetConnection), and a
            // per-connection dispose bound multiplied by many over-cap endpoints
            // would freeze a formula evaluation for minutes.
            var budget = System.Diagnostics.Stopwatch.StartNew();
            foreach (var kv in _connections)
            {
                if (_connections.Count <= MaxCachedConnections)
                    break;
                if (budget.ElapsedMilliseconds > ShutdownBudgetMs)
                {
                    logger.Warn("GetConnection: eviction budget elapsed; cache stays over the cap until the next call");
                    break;
                }
                if (string.Equals(kv.Key, keepEndpoint, StringComparison.Ordinal))
                    continue;

                if (!kv.Value.IsValueCreated)
                {
                    // A never-materialized entry holds nothing; reclaim it so failed
                    // endpoint entries cannot sit past the cap forever.
                    TryRemoveEntry(kv.Key, kv.Value);
                    continue;
                }

                NatsConnection connection;
                try
                {
                    connection = kv.Value.Value; // IsValueCreated: cannot block
                }
                catch
                {
                    continue; // failed placeholder: the retry path replaces it
                }

                if (IsLive(connection))
                    continue;
                if (IsProtected(kv.Key))
                    continue; // live subscribers: never orphan their connection

                if (TryRemoveEntry(kv.Key, kv.Value))
                {
                    DisposeConnection(connection, kv.Key, "evicted (cache over cap)");
                    _jetStream.TryRemove(kv.Key, out _);
                    _downSinceTicks.TryRemove(kv.Key, out _);
                    logger.Info("GetConnection: evicted idle connection endpoint=" + MaskedEndpoint(kv.Key)
                                + " (over " + MaxCachedConnections + " entries)");
                }            }
        }

        /// <summary>
        /// Drops endpoints that have stayed non-live for the whole TTL, at most once
        /// per sweep interval and driven by the call path (never a timer). Live
        /// endpoints and protected endpoints are skipped.
        /// </summary>
        private void SweepDownEndpointsIfDue()
        {
            long now = DateTime.UtcNow.Ticks;
            long last = Interlocked.Read(ref _lastDownSweepTicks);
            if (last != 0 && now - last < DownSweepIntervalMs * TimeSpan.TicksPerMillisecond)
                return;
            if (Interlocked.CompareExchange(ref _lastDownSweepTicks, now, last) != last)
                return; // another caller is already sweeping
            if (_shutdown)
                return;

            long ttlTicks = DownEndpointTtlMs * TimeSpan.TicksPerMillisecond;
            var budget = System.Diagnostics.Stopwatch.StartNew();
            foreach (var kv in _connections)
            {
                if (_shutdown)
                    return;
                if (budget.ElapsedMilliseconds > ShutdownBudgetMs)
                {
                    // Same total-budget reasoning as the eviction path: a sweep is
                    // opportunistic work on the Excel thread, never a minute-long stall.
                    logger.Warn("GetConnection: down-sweep budget elapsed; deferring the rest");
                    return;
                }
                if (!kv.Value.IsValueCreated)
                    continue; // the eviction path reclaims those

                NatsConnection connection;
                try
                {
                    connection = kv.Value.Value; // IsValueCreated: cannot block
                }
                catch
                {
                    continue;
                }

                if (IsLive(connection))
                {
                    _downSinceTicks.TryRemove(kv.Key, out _); // healthy again
                    continue;
                }
                if (IsProtected(kv.Key))
                    continue; // live subscribers: never orphan their connection

                long since = _downSinceTicks.GetOrAdd(kv.Key, now);
                if (now - since < ttlTicks)
                    continue; // down, but not long enough yet

                if (TryRemoveEntry(kv.Key, kv.Value))
                {
                    DisposeConnection(connection, kv.Key, "dropped after being down");
                    _jetStream.TryRemove(kv.Key, out _);
                    _downSinceTicks.TryRemove(kv.Key, out _);
                    logger.Info("GetConnection: dropped idle connection endpoint=" + MaskedEndpoint(kv.Key)
                                + " (down for over " + (DownEndpointTtlMs / 1000) + "s)");
                }
            }
        }

        /// <summary>Conditional remove by value identity: never drops an entry a
        /// concurrent caller has just replaced with a fresh Lazy.</summary>
        private bool TryRemoveEntry(string endpoint, Lazy<NatsConnection> lazy)
        {
            return ((ICollection<KeyValuePair<string, Lazy<NatsConnection>>>)_connections).Remove(
                new KeyValuePair<string, Lazy<NatsConnection>>(endpoint, lazy));
        }

        private static void DisposeConnection(NatsConnection connection, string endpoint, string context) =>
            DisposeConnection(connection, endpoint, context, ConnectionDisposeTimeoutMs);

        /// <summary>
        /// Dispose one connection with a bound no larger than
        /// <paramref name="timeoutMs"/>. The caller uses this to keep a whole
        /// loop (shutdown, eviction, down-sweep) inside a total budget instead of
        /// multiplying a per-connection bound by the number of endpoints.
        /// </summary>
        private static void DisposeConnection(NatsConnection connection, string endpoint, string context, int timeoutMs)
        {
            if (connection == null)
                return;
            if (timeoutMs <= 0)
                return; // no budget left: skip rather than block
            string masked = MaskedEndpoint(endpoint);
            try
            {
                var task = connection.DisposeAsync().AsTask();
                if (!WaitBounded(task, timeoutMs))
                {
                    logger.Debug("NatsConnectionManager: " + context + " dispose timed out endpoint=" + masked);
                }
                else if (task.IsFaulted && task.Exception != null)
                {
                    // Reading Exception marks it observed, so a faulted teardown is
                    // never an unobserved task exception.
                    logger.Debug(task.Exception, "NatsConnectionManager: " + context + " dispose faulted endpoint=" + masked);
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "NatsConnectionManager: " + context + " dispose error endpoint=" + masked);
            }
        }

        private static int Clamp(int value, int min, int max)
        {
            if (value < min)
                return min;
            if (value > max)
                return max;
            return value;
        }

        private static InvalidOperationException ShuttingDown()
        {
            return new InvalidOperationException("NatsConnectionManager is shutting down");
        }
    }
}



