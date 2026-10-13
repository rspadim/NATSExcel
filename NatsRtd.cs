using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using NATS.Client.Core;
using NATS.Client.JetStream;
using NATS.Client.JetStream.Models;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static ExcelDna.Integration.Rtd.ExcelRtdServer;

namespace NATSExcel
{
    /// <summary>
    /// Per-topic RTD state (PLAN S4.1), ported from the sibling RedisExcel
    /// <c>TopicData</c> design and adapted to the async-only NATS client.
    ///
    /// <para>
    /// Threading: three independent producers reach a topic — the poll timer
    /// (worker thread), the NATS subscription callback / JetStream consume loop
    /// (thread-pool threads) and the Excel flush timer. <see cref="_sync"/>
    /// serializes the state update and the Excel push per topic, so an immediate
    /// push and a tick flush can never interleave and overwrite a newer value
    /// with an older one. <see cref="_subscribeSync"/> serializes the
    /// install/detach handoff so a subscription created while the topic is being
    /// disconnected is never leaked.
    /// </para>
    /// <para>
    /// Read-while-disposing race: <see cref="Disconnected"/> is volatile and is
    /// set <em>before</em> a teardown disposes the subscription, and every
    /// mutator re-checks it <em>inside</em> the lock. A callback that was already
    /// past its first check can therefore never publish into an Excel topic that
    /// the server has dropped, nor retain a reference to a torn-down registry.
    /// </para>
    /// </summary>
    internal sealed class TopicData
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private readonly object _sync = new object();
        private readonly object _subscribeSync = new object();
        private string _lastValue;

        /// <summary>Maximum payload length embedded in ToString()/log lines.</summary>
        private const int LastValueLogLimit = 64;

        // Not dirty initially: the "(ConnectData)" placeholder returned by
        // ConnectData must stay visible until the first real value arrives (poll
        // result or subscription message); flushing the initial null _lastValue
        // on the first Excel tick would blank the cell.
        private bool _dirty;

        // Utc ticks of the last value pushed to Excel; with a conflation window
        // this gates the next push so the latest value is delivered at most once
        // per window (0 = never pushed yet, which always pushes).
        private long _lastPushTicks;

        // A persistently failing Excel push stays dirty and is retried every
        // tick; log the first failure at Error and the rest at Debug so a stuck
        // topic cannot flood the log (cleared by a successful push).
        private bool _updateFailureLogged;

        public TopicData(Topic topic, NatsRtdTopicSpec spec, string endpoint, bool polled)
        {
            Topic = topic;
            Spec = spec;
            Endpoint = endpoint;
            Polled = polled;
        }

        public Topic Topic { get; }

        public NatsRtdTopicSpec Spec { get; }

        public NatsRtdKind Kind { get { return Spec.Kind; } }

        /// <summary>The resolved endpoint (already credential-bearing: never log it unmasked).</summary>
        public string Endpoint { get; }

        /// <summary>True for the polled kinds (KVGET, KVKEYS, MSG, REQ, RTDCLUSTER).</summary>
        public bool Polled { get; }

        /// <summary>
        /// The registered subscription (Core/KeyValue/ObjectStore) or, for
        /// <c>JS</c>, the token that cancels the consume loop. Null while a
        /// subscription is awaiting its backoff retry.
        /// </summary>
        public IDisposable Subscription { get; private set; }

        /// <summary>
        /// Cached JetStream stream handle for the polled JS kinds. The handle is
        /// a thin wrapper that only carries the stream name, so caching it saves
        /// one <c>STREAM.INFO</c> round trip per tick; it is dropped on an error
        /// and re-resolved on the next tick.
        /// </summary>
        public INatsJSStream JsStream { get; set; }

        /// <summary>
        /// Failed Subscribe attempts for this topic: the first failure is logged
        /// at Warning with the exception, later retries at Debug with the message
        /// only.
        /// </summary>
        public int SubscribeAttempts;

        private long _nextSubscribeAttemptUtcTicks;

        /// <summary>
        /// Earliest UTC time for the next Subscribe/JS-bind retry after a
        /// failure: the delay doubles from 1s up to a 30s cap (1s, 2s, 4s, ...).
        /// Reset on success. Backed by long ticks so the timer tick and the Excel
        /// thread can read/write it without tearing.
        /// </summary>
        public DateTime NextSubscribeAttemptUtc
        {
            get { return new DateTime(Interlocked.Read(ref _nextSubscribeAttemptUtcTicks), DateTimeKind.Utc); }
            set { Interlocked.Exchange(ref _nextSubscribeAttemptUtcTicks, value.Ticks); }
        }

        private int _subscribeGate;

        /// <summary>
        /// TRUE when the caller acquired the exclusive right to run Subscribe for
        /// this topic. Guards against ConnectData's initial attempt and the timer
        /// retry subscribing the same topic twice (which would leak one
        /// registration).
        /// </summary>
        public bool TryBeginSubscribe() { return Interlocked.CompareExchange(ref _subscribeGate, 1, 0) == 0; }

        public void EndSubscribe() { Interlocked.Exchange(ref _subscribeGate, 0); }

        /// <summary>
        /// Advances the 1s→30s backoff schedule and returns the delay to wait.
        /// </summary>
        public int NextSubscribeBackoff()
        {
            int attempt = Interlocked.Increment(ref SubscribeAttempts);
            // 1s, 2s, 4s, 8s, 16s, then 30s for every later attempt.
            int seconds = Math.Min(30, 1 << Math.Min(attempt - 1, 5));
            NextSubscribeAttemptUtc = DateTime.UtcNow.AddSeconds(seconds);
            return seconds;
        }

        /// <summary>Clears the backoff so a later failure starts again at 1s instead of jumping to the 30s cap.</summary>
        public void ResetSubscribeBackoff()
        {
            NextSubscribeAttemptUtc = default(DateTime);
            Interlocked.Exchange(ref SubscribeAttempts, 0);
        }

        /// <summary>
        /// Installs the subscription unless the topic was disconnected meanwhile.
        /// Returns false when disconnected: the caller must dispose the fresh
        /// token.
        /// </summary>
        public bool InstallSubscription(IDisposable subscription)
        {
            lock (_subscribeSync)
            {
                if (Disconnected)
                    return false;
                Subscription = subscription;
                return true;
            }
        }

        /// <summary>
        /// Marks the topic disconnected and atomically detaches (and returns) the
        /// installed subscription, so a concurrent install can never be lost.
        /// Returns null when there is no subscription (for example a SUB topic
        /// whose initial Subscribe failed and is awaiting a retry).
        /// </summary>
        public IDisposable DetachSubscription()
        {
            lock (_subscribeSync)
            {
                Disconnected = true;
                var subscription = Subscription;
                Subscription = null;
                return subscription;
            }
        }

        /// <summary>Set when the RTD server removed this topic; callbacks and flushes must stop publishing.</summary>
        public volatile bool Disconnected;

        public string LastValue { get { lock (_sync) return _lastValue; } }

        public bool Dirty { get { lock (_sync) return _dirty; } }

        // ------------------------------------------------------------------
        // Polled two-phase commit (compare, then commit once ACCEPTED)
        // ------------------------------------------------------------------

        private byte[] _lastPolledBytes;
        private bool _hasLastPolledBytes;
        private string _lastPolledText;
        private bool _hasLastPolledText;

        /// <summary>
        /// TRUE when the polled byte payload changed since the previous tick.
        /// The comparison is ORDINAL on the raw bytes
        /// (<see cref="NatsResultFormatter.SameMessage"/>), so <c>"007"</c> and
        /// <c>"7"</c> are different values and two distinct invalid UTF-8
        /// payloads never compare equal.
        /// </summary>
        public bool HasChangedPolledBytes(byte[] value)
        {
            lock (_sync)
            {
                if (_hasLastPolledBytes && NatsResultFormatter.SameMessage(_lastPolledBytes, value))
                    return false;
                return true;
            }
        }

        /// <summary>Stores the polled payload as seen; call once it was accepted for delivery.</summary>
        public void CommitPolledBytes(byte[] value)
        {
            lock (_sync)
            {
                _lastPolledBytes = value;
                _hasLastPolledBytes = true;
            }
        }

        /// <summary>
        /// Text counterpart of <see cref="HasChangedPolledBytes"/> for the polled
        /// kinds that produce text (KVKEYS, REQ, RTDCLUSTER). Ordinal comparison:
        /// the RTD duplicate rule is ordinal text, never culture-aware.
        /// </summary>
        public bool HasChangedPolledText(string value)
        {
            lock (_sync)
            {
                if (_hasLastPolledText && string.Equals(_lastPolledText, value, StringComparison.Ordinal))
                    return false;
                return true;
            }
        }

        /// <summary>Stores the polled text as seen; call once it was accepted for delivery.</summary>
        public void CommitPolledText(string value)
        {
            lock (_sync)
            {
                _lastPolledText = value;
                _hasLastPolledText = true;
            }
        }

        /// <summary>
        /// Forgets BOTH polled dedup markers. Called when an error cell is pushed:
        /// the error is not the value a later successful read is compared against,
        /// and the two markers are independent (bytes for KVGET/MSG, text for
        /// KVKEYS/REQ/RTDCLUSTER). Without this, a blip followed by a recovery to
        /// the payload the cell already held before the blip compares equal and is
        /// never pushed, so the stale <c>#ERROR:</c> stays on screen until the
        /// value itself changes.
        /// </summary>
        public void InvalidatePolledMarkers()
        {
            lock (_sync)
            {
                _hasLastPolledBytes = false;
                _hasLastPolledText = false;
            }
        }

        // REQ is issued at most once per everyMs (the parser clamps it to >=100),
        // regardless of the poll tick rate: the interval is a MINIMUM spacing.
        private long _nextRequestDueTicks;

        /// <summary>TRUE when the REQ interval elapsed (or the request was never issued).</summary>
        public bool IsRequestDue(long nowTicks)
        {
            long due = Interlocked.Read(ref _nextRequestDueTicks);
            return due == 0 || nowTicks >= due;
        }

        /// <summary>Schedules the next REQ issue time.</summary>
        public void MarkRequestIssued(int everyMs, long nowTicks)
        {
            long interval = Math.Max(NatsRtdTopicSpec.MinEveryMs, everyMs) * TimeSpan.TicksPerMillisecond;
            Interlocked.Exchange(ref _nextRequestDueTicks, nowTicks + interval);
        }

        /// <summary>
        /// Marks the topic dirty when a value already arrived while ConnectData
        /// was running: Excel-DNA drops pushes for topics that are not active
        /// yet, so the next flush tick must re-publish the cached value once the
        /// topic is active (UpdateValue is a no-op when the value is unchanged).
        /// </summary>
        public void MarkDirtyAfterConnect()
        {
            lock (_sync)
            {
                if (_lastValue != null)
                    _dirty = true;
            }
        }

        // NOTE (upstream risk, Excel-DNA 1.9.0-beta2): a worker thread calling
        // Topic.UpdateValue takes ExcelRtdServer._updateLock and then
        // RtdUpdateSynchronization._lockObject, while the Excel main thread
        // inside ProcessUpdateNotifications takes them in the opposite order
        // (ABBA). The probability per update is low and no freeze was ever
        // observed in the sibling project; the known mitigation (batching
        // UpdateValue on the main thread via ExcelAsyncUtil.QueueAsMacro) would
        // rework the whole delivery path, so this stays tracked upstream.
        public void UpdateAndSendToExcel(string data)
        {
            if (Disconnected)
                return;
            lock (_sync)
            {
                // Defense in depth: the topic may have been disconnected while
                // this call waited for the lock.
                if (Disconnected)
                    return;
                // The state update and the Excel push are serialized per topic,
                // so an immediate push and a tick flush can never interleave. A
                // failed push stays dirty so the next tick retries it.
                _lastValue = data;
                try
                {
                    Topic.UpdateValue(data);
                    // Intentionally keep the topic dirty: the push may have been
                    // dropped because Excel-DNA had not activated the topic yet
                    // (the ConnectData window), and a value arriving in that
                    // window would then never reach the cell until it changes.
                    // The next Excel tick re-pushes the CURRENT value once the
                    // topic is active (UpdateValue is a no-op when unchanged).
                    _dirty = true;
                    _lastPushTicks = DateTime.UtcNow.Ticks;
                }
                catch
                {
                    _dirty = true;
                    throw;
                }
            }
        }

        public void UpdateOnly(string data)
        {
            if (Disconnected)
                return;
            lock (_sync)
            {
                _lastValue = data;
                _dirty = true;
            }
        }

        /// <summary>
        /// Flushes the latest value to Excel when the topic is dirty AND the
        /// conflation window (if any) elapsed; otherwise it stays dirty and a
        /// later tick pushes the then-current value (latest wins).
        /// </summary>
        public void SendToExcelIfDirty(int conflationMs)
        {
            if (Disconnected)
                return;
            Exception updateError = null;
            lock (_sync)
            {
                if (!_dirty)
                    return;
                // Defense in depth: the topic may have been disconnected while
                // this tick waited for the lock.
                if (Disconnected)
                    return;
                long nowTicks = DateTime.UtcNow.Ticks;
                if (!Conflation.IsDue(_lastPushTicks, conflationMs, nowTicks))
                    return; // inside the window: keep the newest value pending
                try
                {
                    // Push while holding the lock: a concurrent state update
                    // cannot interleave and overwrite the newer value. On failure
                    // stay dirty so the next Excel tick retries.
                    Topic.UpdateValue(_lastValue);
                    _dirty = false;
                    _lastPushTicks = nowTicks;
                    _updateFailureLogged = false;
                }
                catch (Exception ex)
                {
                    // _dirty stays true; log outside the lock so a slow log write
                    // cannot delay other producers of this topic.
                    updateError = ex;
                }
            }
            if (updateError != null)
            {
                // Throttled: a permanently failing push would otherwise log an
                // Error every Excel tick (default ~100ms).
                if (_updateFailureLogged)
                {
                    if (logger.IsDebugEnabled)
                        logger.Debug(updateError, "SendToExcelIfDirty: update still failing");
                }
                else
                {
                    _updateFailureLogged = true;
                    logger.Error(updateError, "SendToExcelIfDirty: update failed");
                }
            }
        }

        public override string ToString()
        {
            return "TopicId=" + Topic.TopicId
                + ", Kind=" + Spec.Command
                + ", endpoint=" + AppConfig.MaskNatsEndpoint(Endpoint)
                + ", subject=" + NatsUdfCommon.ShortForLog(NatsRtd.SubjectOf(Spec), LastValueLogLimit)
                + ", LastValue=" + NatsUdfCommon.ShortForLog(LastValue, LastValueLogLimit)
                + ", Dirty=" + Dirty;
        }
    }

    /// <summary>
    /// The NATS RTD server (ProgId <c>NatsRtd</c>), ported from RedisExcel's
    /// <c>RedisRtd</c>: <see cref="ExcelRtdServer"/> lifecycle, a per-instance
    /// topic registry, a poll timer + an Excel flush timer, 1s→30s subscribe
    /// backoff and conflation on delivery.
    ///
    /// <para>
    /// This file is the Excel/COM surface (DESIGN §5.1): it is deliberately the
    /// one root source excluded from the unit/smoke/load/liveness suites (PLAN
    /// D2) because <c>ExcelRtdServer</c> needs Excel. The topic grammar lives in
    /// <c>NatsRtdTopic.cs</c>, which is Excel-free and unit-tested. The
    /// <c>NatsRTD*</c> status helpers live here too (DESIGN §5.1 and §7.9); do
    /// not move them into a new root file without adding that file to
    /// <c>test\TestSupport\ProductionSources.props</c>, or the parity test fails.
    /// </para>
    /// <para>
    /// <b>The Excel thread is never blocked.</b> Polled reads run on the poll
    /// timer, pushes come from the NATS callback or the timers and the JetStream
    /// bind runs on a background task; <see cref="ConnectData"/> only parses,
    /// validates and registers.
    /// </para>
    /// </summary>
    [ComVisible(true)]
    [ProgId("NatsRtd")]
    public class NatsRtd : ExcelRtdServer
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly ConcurrentDictionary<NatsRtd, byte> Instances =
            new ConcurrentDictionary<NatsRtd, byte>();
        private static readonly object InstancesSync = new object();
        private static volatile NatsRtd Instance;

        // Monotonic start ordinal per RTD server instance: when the last-started
        // server terminates, the status/default surface falls back to the
        // highest-ordinal survivor instead of an arbitrary one.
        private static long _nextStartOrdinal;
        private long _startOrdinal;

        private readonly ConcurrentDictionary<int, TopicData> _polledTopics =
            new ConcurrentDictionary<int, TopicData>();
        private readonly ConcurrentDictionary<int, TopicData> _subscribedTopics =
            new ConcurrentDictionary<int, TopicData>();

        /// <summary>
        /// Poll interval for the polled kinds, in milliseconds, when the config
        /// value is missing/invalid (<c>RTD.NatsUpdateRateMs</c>, DESIGN §8).
        /// </summary>
        private const int DefaultNatsUpdateRateMs = 1000;

        /// <summary>Excel flush interval fallback, in milliseconds (<c>RTD.ExcelUpdateRateMs</c>).</summary>
        private const int DefaultExcelUpdateRateMs = 100;

        /// <summary>The <c>RTD.NatsUpdateStyle</c> value that disables immediate pushes.</summary>
        private const string StyleTimer = "Timer";

        private double _excelUpdateRateMs = DefaultExcelUpdateRateMs;
        private double _natsUpdateRateMs = DefaultNatsUpdateRateMs;

        // Realtime style: "Timer" disables immediate pushes (every value waits
        // for the Excel tick); "Realtime"/"Automatic" keep them, subject to the
        // resolved conflation window. The Automatic threshold (msgs/sec) is not
        // applied here: this server has no message counter, so Automatic behaves
        // as Realtime (see ServerStart).
        private bool _realTimeUpdates = true;
        private bool _skipRepeatedMessages = true;
        private int _conflationMs;
        private string _defaultEndpoint;

        private System.Timers.Timer _natsTimer;
        private System.Timers.Timer _excelTimer;

        // Reentrancy gates: a tick is skipped while the matching timer callback is still running.
        private readonly TickGate _natsTickGate = new TickGate();
        private readonly TickGate _excelTickGate = new TickGate();

        // Failed SUB/KV/OBJ registrations are retried with backoff; cap the
        // number of retries per poll tick so a large set of unreachable topics
        // cannot stall the tick.
        private const int PendingSubscribeRetriesPerTick = 4;

        // ------------------------------------------------------------------
        // Status surface (backs NatsRtdStatus; every decision is made here so
        // the guarded [ExcelFunction] wrappers stay trivial)
        // ------------------------------------------------------------------

        public static int ConnectionsCount()
        {
            return NatsRuntime.Connections.LiveRtdConnectionCount();
        }

        /// <summary>
        /// Number of active Core NATS listeners registered by the RTD layer.
        /// RTD-only (UDF listeners are excluded; they use a separate origin tag).
        /// </summary>
        public static int SubscriptionsCount()
        {
            return NatsRuntime.Subscriptions.ListenerCountWithOrigin("RTD");
        }

        /// <summary>Number of subjects with at least one RTD listener (RTD origin only).</summary>
        public static int SubjectsCount()
        {
            return NatsRuntime.Subscriptions.SubjectCountWithOrigin("RTD");
        }

        public static int TopicsCount()
        {
            int total = 0;
            foreach (var instance in Instances.Keys)
                total += instance._polledTopics.Count + instance._subscribedTopics.Count;
            return total;
        }

        /// <summary>The default RTD endpoint, masked; falls back to the config value when no server is live.</summary>
        public static string DefaultEndpoint()
        {
            var instance = Instance;
            string endpoint = instance != null && !string.IsNullOrWhiteSpace(instance._defaultEndpoint)
                ? instance._defaultEndpoint
                : AppConfig.ResolveRtdEndpoint(null);
            return AppConfig.MaskNatsEndpoint(endpoint);
        }

        public static double ExcelUpdateRate() { return Instance?._excelUpdateRateMs ?? 0; }

        /// <summary>
        /// Re-applies the RTD delivery tuning (rates, conflation, style, default
        /// endpoint) to every live RTD server after a
        /// <c>=NatsUDFConfigReload()</c>. Before this, the tuning was frozen at
        /// <see cref="ServerStart"/>: editing <c>RTD.*</c>, <c>ConflationMs</c>,
        /// <c>CoalesceRealtimeUpdates</c> or the default endpoint in
        /// <c>NATSExcel.json</c> had no effect until Excel restarted, the one
        /// place where the "reload without restart" contract did not hold.
        /// Registered as <see cref="NatsRuntime.ConfigReloadHook"/> from
        /// <see cref="ServerStart"/> so the UDF layer (which cannot reference this
        /// Excel-dependent file) can trigger it.
        /// </summary>
        internal static void RefreshFromConfig()
        {
            NatsRtd[] servers;
            lock (InstancesSync)
                servers = Instances.Keys.ToArray();
            foreach (var server in servers)
            {
                try
                {
                    server.ApplyConfig();
                }
                catch (Exception ex)
                {
                    // A timer callback must never let an exception escape.
                    logger.Error(ex, "RefreshFromConfig: failed for one RTD server");
                }
            }
        }

        /// <summary>Re-reads the RTD tuning from the current config and restarts
        /// the timers only when an interval actually changed.</summary>
        private void ApplyConfig()
        {
            var config = AppConfig.Current;
            bool intervalsChanged = false;

            _skipRepeatedMessages = config.SkipRepeatedMessages;
            _conflationMs = Conflation.Resolve(
                config.ConflationMs, config.CoalesceRealtimeUpdates, (int)_excelUpdateRateMs);
            _realTimeUpdates = !string.Equals(
                config.NatsUpdateStyle, StyleTimer, StringComparison.OrdinalIgnoreCase);
            _defaultEndpoint = AppConfig.ResolveRtdEndpoint(null);

            double natsRate = config.NatsUpdateRateMs > 0 ? config.NatsUpdateRateMs : DefaultNatsUpdateRateMs;
            double excelRate = config.ExcelUpdateRateMs > 0 ? config.ExcelUpdateRateMs : DefaultExcelUpdateRateMs;
            if (Math.Abs(natsRate - _natsUpdateRateMs) > 0.5 || Math.Abs(excelRate - _excelUpdateRateMs) > 0.5)
            {
                _natsUpdateRateMs = natsRate;
                _excelUpdateRateMs = excelRate;
                intervalsChanged = true;
            }

            if (intervalsChanged)
            {
                // Recreate both timers so the new cadence takes effect. Stop and
                // dispose the old pair before starting the new one so no tick
                // overlaps (the TickGate already drops a re-entrant tick).
                DisposeTimer(ref _natsTimer);
                DisposeTimer(ref _excelTimer);
                _natsTimer = CreateTimer(_natsUpdateRateMs, "nats", OnNatsTick);
                _excelTimer = CreateTimer(_excelUpdateRateMs, "excel", OnExcelTick);
                _natsTimer.Start();
                _excelTimer.Start();
            }

            logger.Info("RefreshFromConfig: endpoint=" + AppConfig.MaskNatsEndpoint(_defaultEndpoint)
                + ", excelRate=" + _excelUpdateRateMs + "ms"
                + ", natsRate=" + _natsUpdateRateMs + "ms"
                + ", conflationMs=" + _conflationMs
                + ", realtime=" + _realTimeUpdates
                + ", skipRepeated=" + _skipRepeatedMessages
                + ", intervalsChanged=" + intervalsChanged);
        }

        /// <summary>The configured RTD cluster reference (the <c>RTD.cluster</c>
        /// override, else the <c>Default</c> name): the endpoint a blank RTD formula
        /// argument uses (DESIGN §8.1).</summary>
        public static string ClusterName()
        {
            var cfg = AppConfig.Current;
            string name = !string.IsNullOrWhiteSpace(cfg.RtdCluster) ? cfg.RtdCluster : cfg.Default;
            return AppConfig.MaskNatsEndpoint(name);
        }

        // ------------------------------------------------------------------
        // Lifecycle
        // ------------------------------------------------------------------

        protected override bool ServerStart()
        {
            logger.Info("ServerStart: starting RTD server");
            lock (InstancesSync)
            {
                Instance = this;
                _startOrdinal = Interlocked.Increment(ref _nextStartOrdinal);
                Instances[this] = 0;
            }

            var config = AppConfig.Current;
            _skipRepeatedMessages = config.SkipRepeatedMessages;
            _natsUpdateRateMs = config.NatsUpdateRateMs > 0
                ? config.NatsUpdateRateMs
                : DefaultNatsUpdateRateMs;
            _excelUpdateRateMs = config.ExcelUpdateRateMs > 0
                ? config.ExcelUpdateRateMs
                : DefaultExcelUpdateRateMs;
            _conflationMs = Conflation.Resolve(
                config.ConflationMs,
                config.CoalesceRealtimeUpdates,
                (int)_excelUpdateRateMs);
            _defaultEndpoint = AppConfig.ResolveRtdEndpoint(null);
            // "Timer" style pushes nothing immediately; "Realtime" and
            // "Automatic" do. Automatic's message-rate threshold is not applied:
            // this server keeps no per-second counter, so Automatic stays
            // realtime (the conflation window still bounds the push rate).
            _realTimeUpdates = !string.Equals(
                config.NatsUpdateStyle, StyleTimer, StringComparison.OrdinalIgnoreCase);

            logger.Info(
                "ServerStart: config endpoint=" + AppConfig.MaskNatsEndpoint(_defaultEndpoint)
                + ", excelRate=" + _excelUpdateRateMs + "ms"
                + ", natsRate=" + _natsUpdateRateMs + "ms"
                + ", conflationMs=" + _conflationMs
                + ", realtime=" + _realTimeUpdates
                + ", skipRepeated=" + _skipRepeatedMessages);

            _natsTimer = CreateTimer(_natsUpdateRateMs, "nats", OnNatsTick);
            _excelTimer = CreateTimer(_excelUpdateRateMs, "excel", OnExcelTick);
            _natsTimer.Start();
            _excelTimer.Start();
            // Let =NatsUDFConfigReload() re-apply this tuning on a live server.
            NatsRuntime.ConfigReloadHook = RefreshFromConfig;
            return true;
        }

        protected override void ServerTerminate()
        {
            logger.Info("ServerTerminate");
            Teardown();
            // Remove only THIS server; a surviving instance (last-started
            // pointer) is promoted by HighestOrdinalInstance.
            // ResetAfterAddInReload also removes only the instances it tore
            // down, so a server started during its teardown survives.
            lock (InstancesSync)
            {
                Instances.TryRemove(this, out _);
                if (ReferenceEquals(Instance, this))
                    Instance = HighestOrdinalInstance();
            }
        }

        /// <summary>
        /// Stops and disposes this server's timers and drops its topics and
        /// registry entry. Idempotent: called by <see cref="ServerTerminate"/>
        /// and, for any instance Excel did not release, by
        /// <see cref="ResetAfterAddInReload"/> on a same-process reload — a
        /// running <see cref="System.Timers.Timer"/> roots its target, so a
        /// stranded server would otherwise survive the reload and keep polling
        /// the freshly reset runtime.
        /// </summary>
        private void Teardown()
        {
            // Mark every topic before disposing the timers and clearing the
            // registries: despite its name, Timer.Dispose does not wait for an
            // Elapsed callback that is already running, so only the Disconnected
            // flags guarantee that such a tick stops publishing and never touches
            // the live registries after teardown starts.
            foreach (var td in _polledTopics.Values.Concat(_subscribedTopics.Values))
                td.Disconnected = true;

            DisposeTimer(ref _natsTimer);
            DisposeTimer(ref _excelTimer);

            foreach (var td in _subscribedTopics.Values)
            {
                try
                {
                    // DetachSubscription tolerates a null Subscription (a SUB/KV/
                    // OBJ topic whose registration failed and is awaiting a retry)
                    // and wins any race with a retry installing a fresh token.
                    td.DetachSubscription()?.Dispose();
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "ServerTerminate: error disposing subscription");
                }
            }
            _subscribedTopics.Clear();
            _polledTopics.Clear();
        }

        /// <summary>
        /// Tears down every live RTD server. Called from <c>AddIn.AutoOpen</c> on
        /// a same-process reload so a server Excel did not terminate does not
        /// keep its timers (and topic registry) alive across the reload.
        /// </summary>
        internal static void ResetAfterAddInReload()
        {
            // Snapshot under the same lock ServerStart/ServerTerminate mutate the
            // map with, so a concurrent start/terminate cannot be lost or promote
            // a stale instance. Teardown runs outside the lock: it never takes
            // InstancesSync itself.
            NatsRtd[] servers;
            lock (InstancesSync)
            {
                servers = Instances.Keys.ToArray();
            }
            foreach (var server in servers)
                server.Teardown();
            // Remove only the instances that were in the snapshot, never a blanket
            // Clear(): Excel's COM can start a NEW RTD server (ServerStart) during
            // the Teardown window above. Such a server registers itself after this
            // method took its snapshot; clearing the whole map would drop it while
            // its timers keep running, and the next reload - which would find it no
            // longer in the map - could never tear those timers down (a running
            // System.Timers.Timer roots its target), resurrecting exactly the
            // stranded-timer failure this method exists to prevent. The snapshot's
            // instances are already stopped, so removing them by key is safe; the
            // concurrently started server survives and stays reachable for the next
            // reload.
            lock (InstancesSync)
            {
                foreach (var server in servers)
                    Instances.TryRemove(server, out _);
                if (Instance == null || servers.Contains(Instance))
                    Instance = HighestOrdinalInstance();
            }
        }

        /// <summary>
        /// Survivor with the highest start ordinal; called under InstancesSync so
        /// a terminating server never promotes an older instance by accident.
        /// </summary>
        private static NatsRtd HighestOrdinalInstance()
        {
            NatsRtd best = null;
            foreach (var candidate in Instances.Keys)
            {
                if (best == null || candidate._startOrdinal > best._startOrdinal)
                    best = candidate;
            }
            return best;
        }

        private static System.Timers.Timer CreateTimer(double intervalMs, string name, Action action)
        {
            var timer = new System.Timers.Timer(intervalMs) { AutoReset = true };
            timer.Elapsed += (sender, args) =>
            {
                // System.Timers.Timer swallows unhandled exceptions; without this
                // catch the failure is completely silent (one of the original bug
                // symptoms in the sibling project).
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Timer(" + name + "): unhandled error");
                }
            };
            return timer;
        }

        private static void DisposeTimer(ref System.Timers.Timer timer)
        {
            var toDispose = timer;
            timer = null;
            if (toDispose == null)
                return;
            try
            {
                toDispose.Stop();
                toDispose.Dispose();
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "DisposeTimer: error");
            }
        }

        // ------------------------------------------------------------------
        // Topic grammar -> NATS mapping (pure; also used by TopicData.ToString)
        // ------------------------------------------------------------------

        /// <summary>
        /// The subject/identifier an RTD topic reads: the Core subject for
        /// <c>SUB</c>, the KeyValue watch subject <c>$KV.&lt;bucket&gt;.&lt;key&gt;</c>
        /// for <c>KV</c>, the Object Store metadata subject
        /// <c>$O.&lt;bucket&gt;.M.&lt;name&gt;</c> for <c>OBJ</c> and the stream name
        /// for the remaining kinds. Diagnostic only (logs); never used to
        /// dispatch.
        /// </summary>
        internal static string SubjectOf(NatsRtdTopicSpec spec)
        {
            if (spec == null)
                return null;
            switch (spec.Kind)
            {
                case NatsRtdKind.Sub:
                    return spec.Arg1;
                case NatsRtdKind.Kv:
                case NatsRtdKind.KvGet:
                    return "$KV." + spec.Arg1 + "." + spec.Arg2;
                case NatsRtdKind.KvKeys:
                    return "KV_" + spec.Arg1;
                case NatsRtdKind.Obj:
                    return ObjMetaSubject(spec.Arg1, spec.Arg2);
                case NatsRtdKind.Js:
                case NatsRtdKind.Msg:
                    return spec.Arg1;
                case NatsRtdKind.Req:
                    return spec.Arg1;
                default:
                    return spec.Endpoint;
            }
        }

        /// <summary>Object Store metadata subject: <c>$O.&lt;bucket&gt;.M.&lt;base64url(name)&gt;</c>.</summary>
        private static string ObjMetaSubject(string bucket, string name)
        {
            return "$O." + bucket + ".M." + Base64UrlEncode(name ?? string.Empty);
        }

        /// <summary>
        /// Base64url WITHOUT padding, as the Object Store meta subjects use it.
        /// Implemented locally because <c>System.Buffers.Text.Base64UrlEncoder</c>
        /// is not part of the net48 BCL surface.
        /// </summary>
        private static string Base64UrlEncode(string value)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        /// <summary>The KeyValue stream that holds a bucket.</summary>
        private static string KvStream(string bucket) { return "KV_" + bucket; }

        /// <summary>The JetStream kinds, i.e. every kind but plain Core <c>SUB</c> and the Tier-1 cluster cell.</summary>
        private static bool IsJetStreamKind(NatsRtdKind kind)
        {
            switch (kind)
            {
                case NatsRtdKind.Kv:
                case NatsRtdKind.KvGet:
                case NatsRtdKind.KvKeys:
                case NatsRtdKind.Js:
                case NatsRtdKind.Msg:
                case NatsRtdKind.Obj:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>The polled kinds, read on the poll timer instead of pushed.</summary>
        private static bool IsPolledKind(NatsRtdKind kind)
        {
            switch (kind)
            {
                case NatsRtdKind.KvGet:
                case NatsRtdKind.KvKeys:
                case NatsRtdKind.Msg:
                case NatsRtdKind.Req:
                case NatsRtdKind.Cluster:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>The <c>RTDCLUSTER what</c> vocabulary (DESIGN §6/§7.8).</summary>
        private static readonly string[] ClusterWhatValues = { "", "all", "cluster", "nodes", "leader", "lag" };

        // ------------------------------------------------------------------
        // ConnectData / DisconnectData
        // ------------------------------------------------------------------

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            try
            {
                if (_subscribedTopics.TryGetValue(topic.TopicId, out var existingSub))
                {
                    logger.Warn("ConnectData: TopicId=" + topic.TopicId + " already exists in subscriptions");
                    return existingSub.LastValue ?? NatsUdfCommon.ConnectData;
                }
                if (_polledTopics.TryGetValue(topic.TopicId, out var existingPoll))
                {
                    logger.Warn("ConnectData: TopicId=" + topic.TopicId + " already exists in polled topics");
                    return existingPoll.LastValue ?? NatsUdfCommon.ConnectData;
                }

                // The grammar is owned by NatsRtdTopicSpec: this method only
                // consumes the parsed spec, so the frozen command set has exactly
                // one implementation.
                var spec = NatsRtdTopicSpec.Parse(null, topicInfo);
                logger.Info(
                    "ConnectData: command=" + NatsUdfCommon.ShortForLog(spec.Command)
                    + ", arg1=" + NatsUdfCommon.ShortForLog(spec.Arg1)
                    + ", arg2=" + NatsUdfCommon.ShortForLog(spec.Arg2)
                    + ", endpoint=" + AppConfig.MaskNatsEndpoint(spec.Endpoint)
                    + ", TopicId=" + topic.TopicId);

                // A malformed topic is rejected BEFORE registration: registering
                // it would retry a request that can never succeed on every tick,
                // with a full stack trace per attempt.
                if (!spec.IsValid)
                    return ErrorCell("ConnectData", spec.Error);

                string extraError = ValidateKindArguments(spec);
                if (extraError != null)
                    return ErrorCell("ConnectData", extraError);

                // A bounded registry: a runaway sheet (or a hostile workbook)
                // could otherwise register an unbounded number of topics, each with
                // a live subscription or a poll entry, and degrade Excel silently.
                // 10,000 is far above any real workbook and gives a named error
                // instead of unbounded growth.
                const int MaxTopics = 10000;
                if (_polledTopics.Count + _subscribedTopics.Count >= MaxTopics
                    && !_polledTopics.ContainsKey(topic.TopicId)
                    && !_subscribedTopics.ContainsKey(topic.TopicId))
                    return ErrorCell("ConnectData",
                        "too many RTD topics (max " + MaxTopics + "); remove some =RTD cells");

                // Resolve the cluster reference through the RTD section (blank
                // falls back to RTD.cluster -> Default.cluster -> the built-in
                // default). A NON-BLANK reference that resolves to nothing is a
                // #CONFIG-class error, never a silent localhost connection.
                string endpoint = AppConfig.ResolveRtdEndpoint(spec.Endpoint);
                if (string.IsNullOrWhiteSpace(endpoint))
                    return ErrorCell("ConnectData",
                        "'" + spec.Endpoint + "' does not resolve to an endpoint (check Servers/Clusters in NATSExcel.json)");

                // Materialize (or reuse) the client: construction is a cheap
                // allocation and starts the background warm-up. This never
                // connects inline, so the Excel thread stays free.
                var connections = NatsRuntime.Connections;
                connections.GetConnection(endpoint);

                bool polled = IsPolledKind(spec.Kind);
                var td = new TopicData(topic, spec, endpoint, polled);

                // JetStream gate (DESIGN §6): when the server INFO is available
                // AND reports JetStream off, the exact error is returned right
                // away. The topic is still registered so the poll tick (or the
                // JS bind retry) recovers for free once a reconnect changes the
                // answer - and keeps reporting the exact text while it does not.
                string jsError = null;
                if (IsJetStreamKind(spec.Kind))
                {
                    var info = connections.GetServerInfo(endpoint);
                    // IsJetStreamEnabled is false while no INFO has arrived yet;
                    // HasInfo tells "no answer yet" (defer to the tick) apart from
                    // "the server says JetStream is off" (report now).
                    if (info.HasInfo && !connections.IsJetStreamEnabled(endpoint))
                        jsError = NatsUdfCommon.ErrorJetStreamDisabled;
                }

                if (polled)
                {
                    _polledTopics[topic.TopicId] = td;
                }
                else
                {
                    // Register BEFORE subscribing: a transient Subscribe failure
                    // keeps the topic (Subscription == null) so the poll tick
                    // retries it, and the cell starts with the #ERROR text below
                    // until the first message overwrites it.
                    _subscribedTopics[topic.TopicId] = td;
                    var subscribeError = TrySubscribe(td, out _);
                    if (subscribeError != null)
                        return ErrorCell("ConnectData", subscribeError);
                }

                logger.Info("ConnectData: accepted " + td);

                if (jsError != null)
                    return ErrorCell(jsError);

                // A subscription message can arrive between Subscribe and Excel
                // activating the topic; Excel-DNA drops pushes for topics that
                // are not active yet, so return the value cached by that push
                // instead of the sentinel (which would overwrite it permanently
                // in non-coalesced realtime mode). LastValue is read under the
                // topic lock, so this is safe against concurrent updates. A value
                // that raced the activation also marks the topic dirty, so the
                // first flush tick re-publishes it once the topic is active
                // (UpdateValue is a no-op when the value is unchanged).
                string initialValue = td.LastValue;
                td.MarkDirtyAfterConnect();
                return initialValue ?? NatsUdfCommon.ConnectData;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "ConnectData: error");
                return ErrorCell("ConnectData", ex.Message);
            }
        }

        /// <summary>
        /// Kind-specific argument validation the config-free parser cannot do:
        /// the <c>RTDCLUSTER what</c> vocabulary. Returns the error text, or null.
        /// </summary>
        private static string ValidateKindArguments(NatsRtdTopicSpec spec)
        {
            if (spec.Kind != NatsRtdKind.Cluster)
                return null;
            string what = (spec.Arg1 ?? string.Empty).Trim().ToLowerInvariant();
            foreach (var candidate in ClusterWhatValues)
                if (string.Equals(candidate, what, StringComparison.Ordinal))
                    return null;
            return "RTDCLUSTER expects what to be one of [cluster, nodes, leader, lag]";
        }

        protected override void DisconnectData(Topic topic)
        {
            try
            {
                if (_subscribedTopics.TryRemove(topic.TopicId, out var sub))
                {
                    // Mark first so in-flight callbacks stop publishing before the
                    // shared subscription is torn down. DetachSubscription also
                    // tolerates a null Subscription (failed initial Subscribe,
                    // retry pending) and wins any race with a retry that is
                    // installing a fresh token.
                    sub.DetachSubscription()?.Dispose();
                    logger.Info("DisconnectData: removed subscription " + sub);
                }
                else if (_polledTopics.TryRemove(topic.TopicId, out var polled))
                {
                    // In-flight poll ticks may still hold this topic; stop them
                    // from pushing values to a disconnected Excel topic.
                    polled.Disconnected = true;
                    logger.Info("DisconnectData: removed polled topic " + polled);
                }
                else
                {
                    // Excel calls DisconnectData even for topics whose ConnectData
                    // was rejected (Excel-DNA registers them anyway), so this is
                    // expected for malformed topics and not an error.
                    logger.Debug("DisconnectData: unknown TopicId=" + topic.TopicId
                        + " (expected for topics rejected at ConnectData)");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "DisconnectData: error");
            }
        }

        // ------------------------------------------------------------------
        // Subscribing (Core + JetStream)
        // ------------------------------------------------------------------

        /// <summary>
        /// Installs the kind's subscription. Core/KeyValue/ObjectStore kinds
        /// register a ref-counted Core listener through the shared
        /// <see cref="NatsSubscriptionManager"/> (which also performs the
        /// <c>SkipRepeatedMessages</c> suppression for LITERAL subjects, comparing
        /// the raw bytes ordinally; a wildcard subject is never deduplicated).
        /// <c>JS</c> starts a background bind+consume loop and returns its
        /// cancellation token, because the consumer lookup is a round trip and
        /// must never run on the Excel thread.
        /// </summary>
        private IDisposable Subscribe(TopicData td)
        {
            switch (td.Kind)
            {
                case NatsRtdKind.Sub:
                    return SubscribeCore(td, td.Spec.Arg1, td.Spec.Arg2);

                case NatsRtdKind.Kv:
                    // A KV watch is a plain Core subscription on the bucket's
                    // entry subject: the server republishes every revision there.
                    return SubscribeCore(td, "$KV." + td.Spec.Arg1 + "." + td.Spec.Arg2, null);

                case NatsRtdKind.Obj:
                    // Object Store metadata is the only OBJ payload ever read:
                    // the object BODY is never fetched by this cell.
                    return SubscribeCore(td, ObjMetadataFilter(td.Spec.Arg1, td.Spec.Arg2), null);

                case NatsRtdKind.Js:
                    return StartJsConsumeLoop(td);

                default:
                    throw new Exception("kind " + td.Spec.Command + " is not a subscribed kind");
            }
        }

        /// <summary>
        /// The OBJ watch subject. A blank object name watches the whole bucket's
        /// metadata (<c>$O.&lt;bucket&gt;.M.&gt;</c>); the parser requires an
        /// object for OBJ today, so the wildcard branch is defensive.
        /// </summary>
        private static string ObjMetadataFilter(string bucket, string objectName)
        {
            return string.IsNullOrWhiteSpace(objectName)
                ? "$O." + bucket + ".M.>"
                : ObjMetaSubject(bucket, objectName);
        }

        private IDisposable SubscribeCore(TopicData td, string subject, string queueGroup)
        {
            long topicId = td.Topic.TopicId;
            var subscription = NatsRuntime.Subscriptions.Subscribe(
                td.Endpoint,
                subject,
                queueGroup,
                message => OnCoreMessage(td, message),
                "RTD");
            logger.Info("Subscribe: subscribed TopicId=" + topicId
                + ", subject=" + subject
                + ", queue=" + (string.IsNullOrWhiteSpace(queueGroup) ? "<none>" : queueGroup));
            return subscription;
        }

        /// <summary>
        /// Core/KeyValue/ObjectStore message callback. The ref-counted manager
        /// already suppressed a repeated literal-subject payload (ordinal bytes);
        /// nothing is deduplicated again here, and a wildcard subject is never
        /// suppressed anywhere in the path.
        /// </summary>
        private void OnCoreMessage(TopicData td, NatsMessage message)
        {
            if (td.Disconnected)
                return;
            if (logger.IsTraceEnabled)
                logger.Trace("OnCoreMessage: TopicId=" + td.Topic.TopicId
                    + ", subject=" + message.Subject
                    + ", value=" + NatsUdfCommon.ShortForLog(message.Data));

            DeliverToTopic(td, ValueOrTombstone(td, message.Data));
        }

        /// <summary>
        /// A KeyValue tombstone (DEL/PURGE) arrives as a zero-length payload; it
        /// must read as "(no value)", not as a blank cell. Every other kind maps
        /// an empty payload to the empty string, which is a real published value.
        /// </summary>
        private static string ValueOrTombstone(TopicData td, string data)
        {
            if (td.Kind == NatsRtdKind.Kv && string.IsNullOrEmpty(data))
                return NatsUdfCommon.NoValue;
            return data;
        }

        /// <summary>
        /// Attempts to subscribe a registered topic (idempotent). Returns the
        /// failure message, or null on success/already-subscribed. The first
        /// failure is logged at Warning with the exception; later attempts log at
        /// Debug with the message only, so a permanently unreachable endpoint
        /// does not flood the log. Each failure schedules the next retry (1s, 2s,
        /// 4s ... capped at 30s, see <see cref="TopicData.NextSubscribeAttemptUtc"/>).
        /// </summary>
        private string TrySubscribe(TopicData td, out bool attempted)
        {
            attempted = false;
            if (td.Disconnected || td.Subscription != null)
                return null;
            // Only one thread may run Subscribe for a topic at a time: both the
            // Excel thread (ConnectData) and the poll tick (retry) can get here.
            if (!td.TryBeginSubscribe())
                return null;
            attempted = true;
            try
            {
                if (td.Disconnected || td.Subscription != null)
                    return null;
                var subscription = Subscribe(td);
                // Subscribe succeeded: clear the retry backoff so a later failure
                // starts again at 1s instead of jumping to the 30s cap.
                td.ResetSubscribeBackoff();
                if (!td.InstallSubscription(subscription))
                {
                    // DisconnectData/ServerTerminate raced with the install: the
                    // topic was removed while subscribing, so tear the fresh
                    // registration down instead of leaking it.
                    try
                    {
                        subscription.Dispose();
                    }
                    catch (Exception ex)
                    {
                        logger.Debug(ex, "TrySubscribe: error disposing raced subscription, TopicId=" + td.Topic.TopicId);
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                int seconds = td.NextSubscribeBackoff();
                if (td.SubscribeAttempts == 1)
                    logger.Warn(ex, "Subscribe failed, retry in " + seconds + "s: " + td);
                else if (logger.IsDebugEnabled)
                    logger.Debug("Subscribe retry failed (attempt " + td.SubscribeAttempts + "), retry in "
                        + seconds + "s: " + td + ": " + AppConfig.MaskNatsEndpoint(ex.Message));
                return ex.Message;
            }
            finally
            {
                td.EndSubscribe();
            }
        }

        /// <summary>
        /// Retries subscribed topics whose registration failed at ConnectData
        /// (Subscription is null), at most
        /// <see cref="PendingSubscribeRetriesPerTick"/> per tick and only once
        /// each topic's backoff elapsed. A successful retry lets messages flow
        /// into the cell, which still shows the initial "#ERROR: ConnectData: ..."
        /// text until the first message overwrites it.
        /// </summary>
        private void RetryPendingSubscriptions()
        {
            var now = DateTime.UtcNow;
            int retried = 0;
            foreach (var td in _subscribedTopics.Values)
            {
                if (retried >= PendingSubscribeRetriesPerTick)
                    break;
                // TrySubscribe re-checks Disconnected/Subscription itself and
                // reports whether a real attempt happened, so only the backoff
                // gate remains here.
                if (td.NextSubscribeAttemptUtc > now)
                    continue;
                TrySubscribe(td, out bool attempted);
                if (attempted)
                {
                    retried++;
                    // A recovered subscription must stop showing the initial
                    // "#ERROR: ConnectData..." text on a quiet subject: push a
                    // neutral placeholder that the next tick publishes.
                    if (td.Subscription != null && td.LastValue == null)
                        td.UpdateOnly(NatsUdfCommon.Subscribed);
                }
            }
        }

        // ------------------------------------------------------------------
        // JetStream (JS) bind + consume loop
        // ------------------------------------------------------------------

        /// <summary>
        /// Starts the background bind+consume loop for a <c>JS</c> topic and
        /// returns the token that stops it. Returns immediately (no round trip on
        /// the Excel thread); the loop itself reports a failure as the cell's
        /// <c>#ERROR</c> text and retries with the 1s→30s backoff.
        /// </summary>
        private IDisposable StartJsConsumeLoop(TopicData td)
        {
            var cts = new CancellationTokenSource();
            var token = cts.Token;
            _ = Task.Run(async () =>
            {
                // The loop body catches per-iteration failures itself; this
                // wrapper is the last resort, so a failure can never become an
                // unobserved task and no exception can cross the pool.
                try
                {
                    await JsConsumeLoopAsync(td, token).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "JsConsumeLoop: unhandled error " + td);
                }
            });
            return new SubscriptionToken(cts);
        }

        /// <summary>
        /// Binds to an EXISTING consumer (never creates one, PLAN R9) and pushes
        /// each message. A bind or consume failure shows an <c>#ERROR</c> text and
        /// is retried with the capped backoff while the topic stays connected.
        /// </summary>
        private async Task JsConsumeLoopAsync(TopicData td, CancellationToken ct)
        {
            while (!td.Disconnected && !ct.IsCancellationRequested)
            {
                try
                {
                    var connections = NatsRuntime.Connections;

                    // Re-check the JetStream gate: a reconnect may have changed it.
                    var info = connections.GetServerInfo(td.Endpoint);
                    if (info.HasInfo && !connections.IsJetStreamEnabled(td.Endpoint))
                        throw new InvalidOperationException(NatsUdfCommon.ErrorJetStreamDisabled);

                    var js = connections.GetJetStream(td.Endpoint);

                    // The consumer lookup is the authoritative "does it exist"
                    // probe: JS binds to existing consumers only, so a missing
                    // one surfaces the server's own named cause.
                    var consumer = await js.GetConsumerAsync(td.Spec.Arg1, td.Spec.Arg2, ct).ConfigureAwait(false);

                    // A push-based consumer must be consumed through the push
                    // handle (which subscribes to its deliver subject); the pull
                    // handle throws "Consumer is push based" when enumerated.
                    bool pushBased = !string.IsNullOrWhiteSpace(consumer.Info?.Config?.DeliverSubject);
                    if (pushBased)
                        consumer = await js.GetPushConsumerAsync(td.Spec.Arg1, td.Spec.Arg2, ct).ConfigureAwait(false);

                    if (td.Disconnected || ct.IsCancellationRequested)
                        return;

                    td.ResetSubscribeBackoff();
                    // Recovered: clear the "#ERROR" text left by a previous
                    // failure so a quiet consumer does not keep showing it.
                    td.UpdateOnly(NatsUdfCommon.Subscribed);

                    await foreach (var msg in consumer.ConsumeAsync<byte[]>(serializer: null, opts: null, cancellationToken: ct))
                    {
                        if (td.Disconnected || ct.IsCancellationRequested)
                            return;
                        string failure = null;
                        try
                        {
                            // A serialization error is reported per message and the
                            // loop continues; it must never kill the subscription.
                            msg.EnsureSuccess();
                        }
                        catch (Exception ex)
                        {
                            failure = ex.Message;
                        }

                        if (failure != null)
                        {
                            logger.Error("JsConsumeLoop: message error TopicId=" + td.Topic.TopicId
                                + ": " + AppConfig.MaskNatsEndpoint(failure));
                            continue;
                        }

                        var payload = msg.Data;
                        DeliverToTopic(td, payload == null ? NatsUdfCommon.NullValue : NatsResultFormatter.DecodePayload(payload));
                        await TryAckAsync(msg, td, ct).ConfigureAwait(false);
                    }

                    // The consume channel completed without cancellation: the
                    // consumer/stream ended, so fall through to the retry path.
                }
                catch (OperationCanceledException)
                {
                    return;
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "JsConsumeLoop: TopicId=" + td.Topic.TopicId
                        + ", stream=" + NatsUdfCommon.ShortForLog(td.Spec.Arg1)
                        + ", consumer=" + NatsUdfCommon.ShortForLog(td.Spec.Arg2));
                    if (!td.Disconnected && !ct.IsCancellationRequested)
                        td.UpdateOnly(ErrorCell(ErrorContext(td), ex.Message));
                }

                if (td.Disconnected || ct.IsCancellationRequested)
                    return;

                int seconds = td.NextSubscribeBackoff();
                try
                {
                    await Task.Delay(TimeSpan.FromSeconds(seconds), ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Explicit ack (PLAN R9). Skipped when the consumer has no reply
        /// subject (an <c>ack_policy: none</c> consumer), and never fatal.
        /// </summary>
        private static async Task TryAckAsync(INatsJSMsg<byte[]> msg, TopicData td, CancellationToken ct)
        {
            if (msg == null || string.IsNullOrEmpty(msg.ReplyTo))
                return;
            try
            {
                await msg.AckAsync(cancellationToken: ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "JsConsumeLoop: ack failed TopicId=" + td.Topic.TopicId);
            }
        }

        /// <summary>The cancel token returned by <see cref="StartJsConsumeLoop"/> as an IDisposable.</summary>
        private sealed class SubscriptionToken : IDisposable
        {
            private CancellationTokenSource _cts;

            public SubscriptionToken(CancellationTokenSource cts)
            {
                _cts = cts;
            }

            public void Dispose()
            {
                var cts = Interlocked.Exchange(ref _cts, null);
                if (cts == null)
                    return;
                try
                {
                    cts.Cancel();
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "JsConsumeLoop: cancelling the consume loop failed");
                }
                finally
                {
                    cts.Dispose();
                }
            }
        }

        // ------------------------------------------------------------------
        // Timers
        // ------------------------------------------------------------------

        /// <summary>
        /// Reentrancy prologue shared by the two timers: a tick that fires while
        /// the previous one still runs is skipped (never queued). Callers must
        /// release the gate in a finally.
        /// </summary>
        private static bool TryBeginTick(TickGate gate, string name)
        {
            if (gate.TryEnter())
                return true;
            if (logger.IsDebugEnabled)
                logger.Debug(name + ": previous tick still running, skipping");
            return false;
        }

        private void OnNatsTick()
        {
            if (!TryBeginTick(_natsTickGate, "OnNatsTick"))
                return;
            try
            {
                // Failed SUB/KV/OBJ registrations are retried even when this
                // server instance has no polled topics.
                RetryPendingSubscriptions();
                if (_polledTopics.IsEmpty)
                    return;

                // Group by endpoint so one slow/unreachable endpoint cannot delay
                // the others in the same tick; the per-group try/catch keeps the
                // endpoints isolated and the gate is held until every group
                // completed, so ticks never overlap.
                var endpointGroups = _polledTopics.Values.GroupBy(t => t.Endpoint).ToList();
                if (logger.IsDebugEnabled)
                    logger.Debug("OnNatsTick: polling " + _polledTopics.Count + " topic(s) across "
                        + endpointGroups.Count + " endpoint(s), rate=" + _natsUpdateRateMs + "ms");

                Parallel.ForEach(
                    endpointGroups,
                    new ParallelOptions { MaxDegreeOfParallelism = 4 },
                    group =>
                    {
                        try
                        {
                            PollEndpoint(group.Key, group.ToList());
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex, "OnNatsTick: endpoint=" + AppConfig.MaskNatsEndpoint(group.Key));
                        }
                    });
            }
            finally
            {
                _natsTickGate.Exit();
            }
        }

        private void OnExcelTick()
        {
            if (!TryBeginTick(_excelTickGate, "OnExcelTick"))
                return;
            try
            {
                if (_polledTopics.IsEmpty && _subscribedTopics.IsEmpty)
                    return;
                // Dirty values are normally empty when the conflation window is
                // off (updates are sent immediately); flushing every tick is cheap
                // and guarantees delivery of the values still pending inside a
                // window (latest wins).
                foreach (var td in _polledTopics.Values.Concat(_subscribedTopics.Values))
                    td.SendToExcelIfDirty(_conflationMs);
            }
            finally
            {
                _excelTickGate.Exit();
            }
        }

        // ------------------------------------------------------------------
        // Polling
        // ------------------------------------------------------------------

        /// <summary>Per-endpoint poll budget: the configured timeout plus a margin, clamped by the manager.</summary>
        private static int ResponseTimeoutMs()
        {
            return NatsConnectionManager.ResponseTimeoutMs(NatsPool.RtdData);
        }

        /// <summary>
        /// Polls every topic of one endpoint sequentially (NATS has no MGET-style
        /// batch, and the shared client multiplexes the calls anyway). Each topic
        /// is isolated: one failing read never aborts the endpoint's tick.
        /// </summary>
        private void PollEndpoint(string endpoint, List<TopicData> topics)
        {
            var connections = NatsRuntime.Connections;
            bool jetStreamReady = false;

            // The exact JetStream-off text is produced HERE for the polled kinds
            // (DESIGN §6): letting the JS call fail would surface the client's own
            // message instead, and the frozen text must be stable. HasInfo
            // separates "no INFO yet" (poll normally) from "the server says off".
            var info = connections.GetServerInfo(endpoint);
            bool jetStreamOff = info.HasInfo && !connections.IsJetStreamEnabled(endpoint);

            foreach (var td in topics)
            {
                if (td.Disconnected)
                    continue;
                try
                {
                    if (IsJetStreamKind(td.Kind) && jetStreamOff)
                    {
                        const string disabled = "#ERROR: " + NatsUdfCommon.ErrorJetStreamDisabled;
                        if (_skipRepeatedMessages && !td.HasChangedPolledText(disabled))
                            continue;
                        PushPolled(td, disabled, () => td.CommitPolledText(disabled));
                        continue;
                    }

                    // The JetStream context is a thin wrapper over the connection
                    // (no extra socket), created lazily once per endpoint tick.
                    if (IsJetStreamKind(td.Kind) && !jetStreamReady)
                    {
                        connections.GetJetStream(endpoint);
                        jetStreamReady = true;
                    }
                    PollTopic(td, connections);
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "PollEndpoint: TopicId=" + td.Topic.TopicId
                        + ", kind=" + td.Spec.Command
                        + ", endpoint=" + AppConfig.MaskNatsEndpoint(endpoint));
                    if (!td.Disconnected)
                    {
                        // Route the failure through the same ordinal text dedup as
                        // a successful read, so a persistently failing topic does
                        // not push an identical error into Excel every tick.
                        string error = ErrorCell(ErrorContext(td), ex.Message);
                        if (_skipRepeatedMessages && !td.HasChangedPolledText(error))
                            continue;
                        // The error is not the value the next successful read is
                        // deduped against: invalidate both markers so a recovery to
                        // the payload the cell held before the blip is pushed (the
                        // bytes and text markers are independent, and KVGET/MSG
                        // consult only the bytes one).
                        td.InvalidatePolledMarkers();
                        PushPolled(td, error, () => td.CommitPolledText(error));
                    }
                }
            }
        }

        /// <summary>Dispatches one polled read.</summary>
        private void PollTopic(TopicData td, NatsConnectionManager connections)
        {
            switch (td.Kind)
            {
                case NatsRtdKind.KvGet:
                    PollKvGet(td, connections);
                    break;
                case NatsRtdKind.KvKeys:
                    PollKvKeys(td, connections);
                    break;
                case NatsRtdKind.Msg:
                    PollMsg(td, connections);
                    break;
                case NatsRtdKind.Req:
                    PollReq(td, connections);
                    break;
                case NatsRtdKind.Cluster:
                    PollCluster(td, connections);
                    break;
            }
        }

        private void PollKvGet(TopicData td, NatsConnectionManager connections)
        {
            string bucket = td.Spec.Arg1;
            string key = td.Spec.Arg2;
            var stream = ResolveStream(td, connections, KvStream(bucket));

            byte[] data = null;
            try
            {
                var response = AwaitBounded(
                    stream.GetAsync(
                        new StreamMsgGetRequest { LastBySubj = "$KV." + bucket + "." + key },
                        CancellationToken.None));
                data = MessageBytes(response);
            }
            catch (NatsJSApiException api) when (api.Error != null && api.Error.Code == 404)
            {
                // Missing key: a named marker, never an empty cell.
                data = null;
            }

            string value = data == null
                ? NatsUdfCommon.NoValue
                : (data.Length == 0 ? NatsUdfCommon.NoValue : NatsResultFormatter.DecodePayload(data));

            // Two-phase: publish, then commit the value that was accepted. A
            // failed Excel push stays dirty and is retried by the flush tick.
            if (_skipRepeatedMessages && !td.HasChangedPolledBytes(data))
                return;
            PushPolled(td, value, () => td.CommitPolledBytes(data));
        }

        private void PollKvKeys(TopicData td, NatsConnectionManager connections)
        {
            string bucket = td.Spec.Arg1;
            var js = connections.GetJetStream(td.Endpoint);

            var handle = AwaitBounded(js.GetStreamAsync(
                KvStream(bucket),
                new StreamInfoRequest { SubjectsFilter = "$KV." + bucket + ".>" },
                CancellationToken.None));

            string prefix = "$KV." + bucket + ".";
            var keys = new List<string>();
            var subjects = handle?.Info?.State?.Subjects;
            if (subjects != null)
            {
                foreach (var subject in subjects.Keys)
                {
                    if (subject == null)
                        continue;
                    keys.Add(subject.StartsWith(prefix, StringComparison.Ordinal)
                        ? subject.Substring(prefix.Length)
                        : subject);
                }
            }
            keys.Sort(StringComparer.Ordinal);
            string text = NatsResultFormatter.FormatKeysJson(keys);

            if (_skipRepeatedMessages && !td.HasChangedPolledText(text))
                return;
            PushPolled(td, text, () => td.CommitPolledText(text));
        }

        private void PollMsg(TopicData td, NatsConnectionManager connections)
        {
            string streamName = td.Spec.Arg1;
            string filter = td.Spec.Arg2;

            StreamMsgGetRequest request;
            if (string.IsNullOrWhiteSpace(filter))
            {
                // No filter: "the last message in the stream". Resolved from the
                // stream state so the server never sees a request with neither
                // seq nor last_by_subj (which it rejects).
                var handle = ResolveStream(td, connections, streamName);
                ulong lastSeq = handle?.Info?.State?.LastSeq ?? 0;
                if (lastSeq == 0)
                {
                    const string empty = NatsUdfCommon.NoValue;
                    if (_skipRepeatedMessages && !td.HasChangedPolledText(empty))
                        return;
                    PushPolled(td, empty, () => td.CommitPolledText(empty));
                    return;
                }
                request = new StreamMsgGetRequest { Seq = lastSeq };
            }
            else
            {
                request = new StreamMsgGetRequest { LastBySubj = filter };
            }

            var stream = ResolveStream(td, connections, streamName);

            byte[] data = null;
            try
            {
                var response = AwaitBounded(stream.GetAsync(request, CancellationToken.None));
                data = MessageBytes(response);
            }
            catch (NatsJSApiException api) when (api.Error != null && api.Error.Code == 404)
            {
                data = null;
            }

            // MSG reports an empty stored payload as "" (a real published value);
            // only a MISSING message is the "(no value)" marker.
            string value = data == null
                ? NatsUdfCommon.NoValue
                : NatsResultFormatter.DecodePayload(data);
            if (_skipRepeatedMessages && !td.HasChangedPolledBytes(data))
                return;
            PushPolled(td, value, () => td.CommitPolledBytes(data));
        }

        private void PollReq(TopicData td, NatsConnectionManager connections)
        {
            long now = DateTime.UtcNow.Ticks;
            if (!td.IsRequestDue(now))
                return;
            td.MarkRequestIssued(td.Spec.EveryMs, now);

            string subject = td.Spec.Arg1;
            // A blank payload publishes an EMPTY body (the client's serializer
            // rejects null, and "no payload" is what an empty REQ cell means).
            string payload = td.Spec.Arg2 ?? string.Empty;
            int timeoutMs = ResponseTimeoutMs();

            string value;
            try
            {
                var connection = connections.GetConnection(td.Endpoint);
                var reply = AwaitBounded(connection.RequestAsync<string, string>(
                    subject: subject,
                    data: payload,
                    replyOpts: new NatsSubOpts
                    {
                        Timeout = TimeSpan.FromMilliseconds(timeoutMs),
                        // A named "No Responders" beats a silent timeout.
                        ThrowIfNoResponders = true,
                    },
                    cancellationToken: CancellationToken.None));
                string text = reply.Data;
                value = text ?? NatsUdfCommon.NoValue;
            }
            catch (NatsNoRespondersException)
            {
                value = NatsUdfCommon.NoResponders;
            }
            catch (NatsNoReplyException)
            {
                value = NatsUdfCommon.Timeout;
            }
            catch (TimeoutException)
            {
                value = NatsUdfCommon.Timeout;
            }

            if (_skipRepeatedMessages && !td.HasChangedPolledText(value))
                return;
            PushPolled(td, value, () => td.CommitPolledText(value));
        }

        /// <summary>
        /// The live cluster cell. Tier 1 (<c>cluster</c>, <c>nodes</c>) reads the
        /// server INFO the client already holds: no extra round trip and no
        /// account privilege. <c>leader</c> is Tier 2 and goes through the NATS
        /// system account (<c>$SYS.REQ.SERVER.PING</c>); an endpoint without it
        /// reports the named error instead of an empty cell that looks healthy.
        /// </summary>
        private void PollCluster(TopicData td, NatsConnectionManager connections)
        {
            string what = (td.Spec.Arg1 ?? string.Empty).Trim().ToLowerInvariant();
            var info = connections.GetServerInfo(td.Endpoint);

            string value;
            switch (what)
            {
                case "cluster":
                    value = info.HasInfo
                        ? (string.IsNullOrWhiteSpace(info.Cluster) ? "(standalone)" : info.Cluster)
                        : "(connecting)";
                    break;

                case "nodes":
                    value = info.HasInfo ? NodesText(info) : "(connecting)";
                    break;

                case "leader":
                    value = ClusterLeader(td, connections);
                    break;

                case "lag":
                    // Deliberately not invented here: a per-node raft lag needs
                    // the stream's own Raft group, which an RTD cell without a
                    // stream argument cannot address.
                    value = "lag needs the per-stream raft view: use NatsUDFJSRaftInfo(stream)";
                    break;

                default: // blank = "all"
                    value = "cluster=" + (info.HasInfo
                            ? (string.IsNullOrWhiteSpace(info.Cluster) ? "(standalone)" : info.Cluster)
                            : "?")
                        + " nodes=" + (info.HasInfo ? NodesText(info) : "?")
                        + " leader=" + ClusterLeader(td, connections);
                    break;
            }

            if (_skipRepeatedMessages && !td.HasChangedPolledText(value))
                return;
            PushPolled(td, value, () => td.CommitPolledText(value));
        }

        /// <summary>
        /// The nodes the CLIENT can see (its discovery pool plus the initially
        /// seeded server name). This is what the connected client sees, not the
        /// full cluster membership: the full view needs the system account.
        /// </summary>
        private static string NodesText(NatsServerSnapshot info)
        {
            var nodes = new List<string>();
            foreach (var url in info.ConnectUrls)
                if (!string.IsNullOrWhiteSpace(url))
                    nodes.Add(url);

            string text = nodes.Count > 0
                ? string.Join(", ", nodes)
                : "(none)";
            if (!string.IsNullOrWhiteSpace(info.Name))
                text = info.Name + " [" + text + "]";
            return text;
        }

        /// <summary>
        /// Tier 2: the first server that answers <c>$SYS.REQ.SERVER.PING</c>. On a
        /// single-node (or single-leader-per-view) deployment that IS the node in
        /// charge; for a multi-node cluster it is the first responder, which is
        /// what a cell can honestly report without correlating every reply.
        /// </summary>
        private string ClusterLeader(TopicData td, NatsConnectionManager connections)
        {
            const string noSystemAccount =
                "the system account is not available on this endpoint (cluster queries need a system account)";
            int timeoutMs = ResponseTimeoutMs();
            try
            {
                var connection = connections.GetConnection(td.Endpoint);
                var reply = AwaitBounded(connection.RequestAsync<string, string>(
                    subject: "$SYS.REQ.SERVER.PING",
                    data: string.Empty,
                    replyOpts: new NatsSubOpts
                    {
                        Timeout = TimeSpan.FromMilliseconds(timeoutMs),
                        ThrowIfNoResponders = true,
                    },
                    cancellationToken: CancellationToken.None));

                string body = reply.Data;
                if (string.IsNullOrWhiteSpace(body))
                    return noSystemAccount;

                // The documented PING reply nests the node identity under
                // "server" ({ "server": { "name": ..., "id": ... }, "statsz": ... }).
                // An unrecognized shape is reported verbatim rather than guessed at.
                try
                {
                    var json = JObject.Parse(body);
                    string name = (string)json["server"]?["name"];
                    if (!string.IsNullOrWhiteSpace(name))
                        return name;
                    string id = (string)json["server"]?["id"];
                    if (!string.IsNullOrWhiteSpace(id))
                        return id;
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "ClusterLeader: the $SYS ping reply was not JSON");
                }

                return NatsResultFormatter.BoundCellText(body);
            }
            catch (NatsNoRespondersException)
            {
                return noSystemAccount;
            }
            catch (NatsNoReplyException)
            {
                return "the system account did not answer within the response timeout";
            }
            catch (TimeoutException)
            {
                return "the system account did not answer within the response timeout";
            }
            catch (Exception ex)
            {
                return AppConfig.MaskNatsEndpoint(ex.Message);
            }
        }

        /// <summary>
        /// The raw payload of a <c>STREAM.MSG.GET</c> response, or null when the
        /// response carries no message. An EMPTY payload is preserved (length 0):
        /// only the Key/Value path reinterprets it as a tombstone.
        /// </summary>
        private static byte[] MessageBytes(StreamMsgGetResponse response)
        {
            if (response == null || response.Message == null)
                return null;
            var data = response.Message.Data;
            return data.IsEmpty ? Array.Empty<byte>() : data.ToArray();
        }

        /// <summary>Resolves (and caches) the JetStream stream handle for a polled kind.
        /// The handle only carries the stream name, so it stays valid across
        /// ticks; an error clears it so the next tick re-resolves from scratch.
        /// </summary>
        private INatsJSStream ResolveStream(TopicData td, NatsConnectionManager connections, string streamName)
        {
            var cached = td.JsStream;
            if (cached != null)
                return cached;
            try
            {
                var js = connections.GetJetStream(td.Endpoint);
                var handle = AwaitBounded(js.GetStreamAsync(streamName, null, CancellationToken.None));
                td.JsStream = handle;
                return handle;
            }
            catch
            {
                td.JsStream = null;
                throw;
            }
        }

        /// <summary>
        /// Runs a NATS task with a hard bound. The poll timer runs on a worker
        /// thread (never the Excel thread), and a bounded wait is the only
        /// blocking primitive in the codebase (the sole intended consumer of
        /// direct task waits, with the UDF path's helper).
        /// </summary>
        private static T AwaitBounded<T>(ValueTask<T> task)
        {
            return NatsConnectionManager.WaitBoundedResult(task.AsTask(), ResponseTimeoutMs());
        }

        /// <summary>Publishes a polled value and commits it once accepted.</summary>
        private void PushPolled(TopicData td, string value, Action commit)
        {
            if (td.Disconnected)
                return;
            if (logger.IsTraceEnabled)
                logger.Trace("PushPolled: TopicId=" + td.Topic.TopicId
                    + ", kind=" + td.Spec.Command
                    + ", value=" + NatsUdfCommon.ShortForLog(value));
            DeliverToTopic(td, value);
            commit?.Invoke();
        }

        // ------------------------------------------------------------------
        // Delivery
        // ------------------------------------------------------------------

        /// <summary>
        /// Single owner of the realtime-vs-coalesced delivery policy: a realtime
        /// push only when the resolved conflation window is off, otherwise the
        /// value waits for the Excel tick (dirty flag), which conflates to the
        /// newest value per window. The window is also bounded by
        /// <see cref="NatsResultFormatter.BoundCellText"/> so an oversized payload
        /// can never break the cell.
        /// </summary>
        private void DeliverToTopic(TopicData td, string value)
        {
            value = NatsResultFormatter.BoundCellText(value);
            if (_realTimeUpdates && Conflation.PushImmediately(_conflationMs))
                td.UpdateAndSendToExcel(value);
            else
                td.UpdateOnly(value);
        }

        // ------------------------------------------------------------------
        // Formatting helpers
        // ------------------------------------------------------------------

        /// <summary>
        /// The frozen RTD error cell: <c>#ERROR: &lt;context&gt;: &lt;masked cause&gt;</c>
        /// (PLAN §1.3). Every part is masked, so a server error naming a
        /// credential-bearing URL can never leak it, and the text is bounded to
        /// the Excel cell limit.
        /// </summary>
        private static string ErrorCell(string context, string message)
        {
            string cause = AppConfig.MaskNatsEndpoint(message ?? "unknown error");
            string text = string.IsNullOrWhiteSpace(context)
                ? "#ERROR: " + cause
                : "#ERROR: " + context + ": " + cause;
            return NatsResultFormatter.BoundCellText(text);
        }

        /// <summary><c>#ERROR: &lt;masked cause&gt;</c> without a context.</summary>
        private static string ErrorCell(string message)
        {
            return ErrorCell(null, message);
        }

        /// <summary>Short, human context for a polled/subscribe failure cell (identifiers only, never credentials).</summary>
        private static string ErrorContext(TopicData td)
        {
            switch (td.Kind)
            {
                case NatsRtdKind.KvGet:
                case NatsRtdKind.Kv:
                    return td.Spec.Command + " " + NatsUdfCommon.ShortForLog(td.Spec.Arg1) + "/" + NatsUdfCommon.ShortForLog(td.Spec.Arg2);
                case NatsRtdKind.KvKeys:
                    return "KVKEYS " + NatsUdfCommon.ShortForLog(td.Spec.Arg1);
                case NatsRtdKind.Req:
                    return "REQ " + NatsUdfCommon.ShortForLog(td.Spec.Arg1);
                case NatsRtdKind.Js:
                case NatsRtdKind.Msg:
                    return td.Spec.Command + " " + NatsUdfCommon.ShortForLog(td.Spec.Arg1);
                case NatsRtdKind.Obj:
                    return "OBJ " + NatsUdfCommon.ShortForLog(td.Spec.Arg1);
                case NatsRtdKind.Cluster:
                    return "RTDCLUSTER";
                default:
                    return td.Spec.Command;
            }
        }
    }

    /// <summary>
    /// The six <c>NatsRTD*</c> status helpers (DESIGN §7.6 / §7.9). They live in
    /// this file because DESIGN §5.1 places them with the RTD server, and
    /// because a new root source would have to be registered in
    /// <c>test\TestSupport\ProductionSources.props</c> (the parity test fails
    /// otherwise).
    ///
    /// <para>
    /// Every helper reports a neutral value instead of throwing after
    /// <c>ServerTerminate</c>/<c>NatsRuntime.Shutdown</c>: the runtime refuses to
    /// resurrect the managers, so an unguarded access would surface
    /// <see cref="InvalidOperationException"/> in the cell during Excel
    /// teardown/reload.
    /// </para>
    /// </summary>
    public static class NatsRtdStatus
    {
        /// <summary>Runs the status read, replacing a teardown/reload exception
        /// with the neutral value the cell used to get from its own catch.</summary>
        private static T Guarded<T>(Func<T> read, T fallback)
        {
            try
            {
                return read();
            }
            catch
            {
                return fallback;
            }
        }

        [ExcelFunction(Description = "Returns the number of active NATS connections used by the RTD server.", IsVolatile = true)]
        public static int NatsRTDConnectionCount()
        {
            return Guarded(() => NatsRtd.ConnectionsCount(), 0);
        }

        [ExcelFunction(Description = "Returns the total number of active Excel RTD topics.", IsVolatile = true)]
        public static int NatsRTDTopicCount()
        {
            return Guarded(() => NatsRtd.TopicsCount(), 0);
        }

        [ExcelFunction(Description = "Returns the number of active NATS subscriptions registered by the RTD server.", IsVolatile = true)]
        public static int NatsRTDSubscriptionCount()
        {
            return Guarded(() => NatsRtd.SubscriptionsCount(), 0);
        }

        [ExcelFunction(Description = "Returns the number of NATS subjects with RTD subscriptions.", IsVolatile = true)]
        public static int NatsRTDSubjectCount()
        {
            return Guarded(() => NatsRtd.SubjectsCount(), 0);
        }

        [ExcelFunction(Description = "Returns the default NATS endpoint used by the RTD server (credentials masked).", IsVolatile = true)]
        public static string NatsRTDDefaultHost()
        {
            return Guarded(() => NatsRtd.DefaultEndpoint() ?? "(none)", "(none)");
        }

        [ExcelFunction(Description = "Returns the configured RTD cluster reference (or the process default).", IsVolatile = true)]
        public static string NatsRTDCluster()
        {
            return Guarded(() => NatsRtd.ClusterName() ?? "(none)", "(none)");
        }
    }
}
