using ExcelDna.Integration;
using NLog;
using System;

namespace NATSExcel
{
    /// <summary>
    /// Shared add-in state (PLAN §1.4 / F2.15). Connections and subscriptions are
    /// expensive resources: they are created once per Excel process and reused by
    /// every RTD server (Excel may instantiate more than one) and by UDF calls.
    /// All NATS access goes through here (AGENTS golden rule 4), so this is the
    /// single place the process-wide <see cref="Connections"/> and
    /// <see cref="Subscriptions"/> managers are created, shut down and recreated
    /// after a same-process add-in reload. Mirrors <c>RedisRuntime</c>.
    /// <para>
    /// Both fields are <c>volatile</c> and published in an order that keeps them
    /// consistent: a reader that observes a non-null <c>_connections</c> is
    /// guaranteed to observe a non-null <c>_subscriptions</c>.
    /// </para>
    /// </summary>
    internal static class NatsRuntime
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly object Sync = new object();

        private static volatile NatsConnectionManager _connections;
        private static volatile NatsSubscriptionManager _subscriptions;

        // Set at the start of Shutdown, inside the lock, and cleared only by
        // ResetAfterAddInReload (a same-process add-in reload): after an
        // explicit Shutdown (for example AutoClose) no code path may
        // recreate the managers, no matter how many RTD timer ticks or UDF
        // calls race with the teardown.
        private static volatile bool _shutdown;

        private static readonly Func<NatsConnectionManager> ReadConnections = () => _connections;
        private static readonly Func<NatsSubscriptionManager> ReadSubscriptions = () => _subscriptions;

        public static NatsConnectionManager Connections
        {
            get
            {
                EnsureInitialized();
                return EnsureManager(_connections, ReadConnections);
            }
        }

        public static NatsSubscriptionManager Subscriptions
        {
            get
            {
                EnsureInitialized();
                return EnsureManager(_subscriptions, ReadSubscriptions);
            }
        }

        /// <summary>
        /// Shared null-handling tail of the manager getters: a null value
        /// means the initializer may have raced with Shutdown, so re-run it
        /// (it now throws instead of recreating the managers) and re-read;
        /// the final throw covers a Shutdown that completed in between.
        /// <paramref name="reread"/> reads the volatile field again without
        /// passing it by ref (forbidden for volatile fields).
        /// </summary>
        private static T EnsureManager<T>(T value, Func<T> reread) where T : class
        {
            if (value == null)
            {
                // Shutdown raced with this access: EnsureInitialized now
                // throws instead of recreating the managers.
                EnsureInitialized();
                value = reread();
            }
            return value ?? throw new InvalidOperationException("NatsRuntime is shutting down");
        }

        private static void EnsureInitialized()
        {
            // Checked before creation: once Shutdown has started, callers (timer
            // ticks included) must fail instead of resurrecting disposed managers.
            if (_shutdown)
                throw new InvalidOperationException("NatsRuntime is shutting down");
            if (_connections != null && _subscriptions != null)
                return;
            lock (Sync)
            {
                if (_shutdown)
                    throw new InvalidOperationException("NatsRuntime is shutting down");
                if (_connections != null && _subscriptions != null)
                    return;
                var connections = new NatsConnectionManager();
                var subscriptions = new NatsSubscriptionManager(connections);
                // Eviction must never orphan a live subscription: the
                // subscription manager vetoes evicting endpoints with listeners.
                connections.SetEvictionProtection(subscriptions.HasActiveSubscribers);
                // The NATS client re-subscribes its subjects internally after a
                // reconnect, so no explicit resubscribe wiring is required here.
                // Publish _subscriptions first: a reader that observes a non-null
                // _connections must also observe a fully initialized _subscriptions.
                _subscriptions = subscriptions;
                _connections = connections;
            }
        }

        /// <summary>
        /// Resolves a formula's optional host argument to a concrete endpoint (or
        /// a comma-joined URL list) for UDF dispatch (PLAN §1.4 H5): the frozen
        /// bridge between the Excel-facing argument and
        /// <see cref="AppConfig.ResolveUdfEndpoint"/>. A null/omitted/empty cell,
        /// or a blank text value, means "use the UDF default cluster". A
        /// non-blank reference that does not resolve returns null so the async
        /// dispatcher can surface a <c>#CONFIG</c> error instead of connecting
        /// to a hidden host. Never throws: an invalid argument type also yields
        /// null and is reported by the dispatched Core body.
        /// </summary>
        public static string ResolveEndpointForDispatch(object optionalHost)
        {
            try
            {
                if (optionalHost == null || optionalHost is ExcelMissing || optionalHost is ExcelEmpty)
                    return AppConfig.ResolveUdfEndpoint(null);
                if (!(optionalHost is string host))
                    return null; // non-text endpoint: the Core body reports the type error
                return AppConfig.ResolveUdfEndpoint(string.IsNullOrWhiteSpace(host) ? null : host);
            }
            catch (Exception ex)
            {
                // Resolution must never take down the dispatcher; log (masked by
                // the caller's discipline) and let the Core body report it.
                logger.Debug(ex, "ResolveEndpointForDispatch: endpoint resolution failed");
                return null;
            }
        }

        public static void Shutdown()
        {
            lock (Sync)
            {
                // The flag is set BEFORE the managers are disposed/nulled (and
                // is cleared only by ResetAfterAddInReload), so a racing
                // EnsureInitialized either observes it and throws, or waits
                // on the lock and then observes it.
                _shutdown = true;
                try
                {
                    _subscriptions?.Dispose();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Shutdown: error disposing subscriptions");
                }
                try
                {
                    _connections?.Shutdown();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Shutdown: error disposing connections");
                }
                _subscriptions = null;
                _connections = null;
            }
        }

        /// <summary>
        /// Registered by the RTD layer (which the unit test project excludes,
        /// because it needs Excel/COM) so a config reload can re-apply the RTD
        /// delivery tuning without the UDF layer referencing <c>NatsRtd</c>.
        /// </summary>
        internal static Action ConfigReloadHook { get; set; }

        /// <summary>
        /// Clears the shutdown tombstone so a surviving AppDomain (add-in
        /// unloaded and reloaded without an Excel restart) can lazily recreate
        /// the managers on the next access. No-op when the runtime was not
        /// shut down; thread-safe under the same lock used by
        /// <see cref="Shutdown"/>.
        /// </summary>
        public static void ResetAfterAddInReload()
        {
            lock (Sync)
            {
                if (!_shutdown)
                    return;
                // Managers were disposed and nulled by Shutdown; clear the
                // fields again for clarity, then drop the tombstone last so a
                // racing EnsureInitialized never observes a cleared flag with
                // stale managers.
                _subscriptions = null;
                _connections = null;
                _shutdown = false;
            }
        }
    }
}
