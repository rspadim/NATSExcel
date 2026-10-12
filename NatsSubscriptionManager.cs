using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using NATS.Client.Core;
using NLog;

namespace NATSExcel
{
    /// <summary>
    /// An immutable snapshot of one received Core NATS message (frozen shape,
    /// PLAN.md §1.4). A single instance is fanned out to every listener of a
    /// physical subscription, so it is value-like and never mutated afterwards.
    /// <para>
    /// <see cref="DataBytes"/> is the raw payload the server delivered (already
    /// an owned copy from the client's byte[] deserializer) and <see cref="Data"/>
    /// is its lenient UTF-8 decoding (<see cref="NatsResultFormatter.DecodePayload"/>:
    /// invalid sequences become U+FFFD, a null payload stays null, an empty payload
    /// becomes ""). Keep the bytes around for the second consumer: the ordinal
    /// duplicate comparison must never decode (<c>"007"</c> != <c>"7"</c>).
    /// </para>
    /// <para>
    /// <see cref="Sequence"/> and <see cref="Timestamp"/> are 0/null for plain
    /// Core pub/sub, which carries neither. When the server supplies the
    /// <c>Nats-Sequence</c> / <c>Nats-Time-Stamp</c> headers (JetStream payloads
    /// delivered over a Core subscription) they are surfaced here; an
    /// unparseable value is reported as absent rather than as a wrong number.
    /// </para>
    /// </summary>
    public sealed class NatsMessage
    {
        /// <summary>Concrete subject the message was published on (not the
        /// pattern the cell subscribed with: <c>orders.*</c> delivers
        /// <c>orders.new</c>).</summary>
        public string Subject { get; }

        /// <summary>Reply subject, or null when the message has none.</summary>
        public string Reply { get; }

        /// <summary>Lenient UTF-8 decoding of <see cref="DataBytes"/>.</summary>
        public string Data { get; }

        /// <summary>Raw payload bytes (never decoded for comparisons).</summary>
        public byte[] DataBytes { get; }

        /// <summary>
        /// Request headers, one entry per name with all of its values; an empty
        /// dictionary when the message carries none. Name lookup is
        /// case-insensitive (the NATS header spec is case-insensitive) while the
        /// original spelling is preserved.
        /// </summary>
        public IReadOnlyDictionary<string, string[]> Headers { get; }

        /// <summary>JetStream sequence, or 0 when absent (plain Core).</summary>
        public ulong Sequence { get; }

        /// <summary>Server timestamp (UTC), or null when absent.</summary>
        public DateTime? Timestamp { get; }

        public NatsMessage(
            string subject,
            string reply,
            string data,
            byte[] dataBytes,
            IReadOnlyDictionary<string, string[]> headers,
            ulong sequence,
            DateTime? timestamp)
        {
            Subject = subject;
            Reply = reply;
            Data = data;
            DataBytes = dataBytes;
            Headers = headers ?? NatsSubscriptionManager.NoHeaders;
            Sequence = sequence;
            Timestamp = timestamp;
        }
    }

    /// <summary>
    /// Ref-counted Core NATS subscriptions (PLAN.md F2.10–F2.14), ported from the
    /// sibling RedisExcel <c>RedisSubscriptionManager</c> design.
    ///
    /// One physical NATS wire subscription exists per
    /// <c>(endpoint, subject, queueGroup)</c> — a plain subscriber and a
    /// queue-group subscriber on the same subject are different physical
    /// subscriptions (DESIGN.md §5.1) — broadcasting to every registered
    /// listener. The first listener starts it, later listeners only join, and
    /// the last listener tears it down. The optional <c>origin</c> tag
    /// ("RTD"/"UDF") is diagnostic only and never affects delivery or identity.
    ///
    /// Wildcards are part of the subject (<c>*</c> one token, <c>&gt;</c> one or
    /// more); there is no separate pattern command.
    ///
    /// Delivery is asynchronous. <see cref="Subscribe"/> never blocks the Excel
    /// thread and never performs the wire SUBSCRIBE inline: it registers the
    /// listener and starts a per-state consume loop on the thread pool. The loop
    /// runs the network SUBSCRIBE outside the lifecycle lock and installs the
    /// returned <c>INatsSub&lt;byte[]&gt;</c> under it only after re-checking the
    /// per-key epoch, so a teardown racing the subscribe can never leave a ghost
    /// listener. Callers therefore must not assume the server has seen the SUB by
    /// the time <see cref="Subscribe"/> returns; that is inherent to the
    /// async-only client and to Core's at-most-once delivery.
    ///
    /// Reconnect: verified against NATS.Net 3.3.0 — the client re-issues every
    /// live SUB on reconnect (<c>SubscriptionManager.WriteReconnectCommandsAsync</c>),
    /// so a connection drop does NOT end the consume loop and needs no action
    /// here. The loop re-issues its own SUBSCRIBE (with a capped backoff) only
    /// when the wire subscription genuinely ends while this manager is still
    /// alive and the state is still current.
    ///
    /// Duplicate suppression: price-style feeds republish unchanged values
    /// constantly; when <c>SkipRepeatedMessages</c> is on, an identical
    /// CONSECUTIVE payload is compared ORDINALLY on the raw bytes (no decode,
    /// so <c>"007"</c> != <c>"7"</c>) and skipped. Only literal subjects are
    /// deduplicated: a subject containing <c>*</c> or <c>&gt;</c> interleaves
    /// different concrete subjects and is never suppressed.
    ///
    /// Every listener exception is caught and logged; a throwing consumer can
    /// never kill the consume loop.
    /// </summary>
    public sealed class NatsSubscriptionManager : IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>Empty header set shared by headerless messages.</summary>
        internal static readonly IReadOnlyDictionary<string, string[]> NoHeaders =
            new Dictionary<string, string[]>(0, StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Raised after a listener was successfully registered (first listener or
        /// joiner, any origin): (endpoint, subject, queueGroup). The queue group
        /// is null when there is none. Consumer-side dedup markers are cleared on
        /// join so a returning listener still receives the next payload even when
        /// its value did not change. Handler exceptions are caught and logged;
        /// they never fail the Subscribe call.
        /// </summary>
        internal static event Action<string, string, string> ListenerJoined;

        /// <summary>Test-only override of <c>SkipRepeatedMessages</c> (null = use
        /// the process configuration). The production sources are compiled
        /// directly into the test projects, so this seam pins the
        /// duplicate-suppression branch deterministically instead of depending on
        /// a machine-local NATSExcel.json. Must be set before the manager is
        /// constructed (the flag is captured per instance).</summary>
#pragma warning disable 0649 // assigned only by the linked test sources
        internal static bool? SkipRepeatedOverrideForTests;
#pragma warning restore 0649

        private static bool SkipRepeatedEnabled =>
            SkipRepeatedOverrideForTests ?? AppConfig.SkipRepeatedMessages;

        /// <summary>Effective duplicate-suppression flag after applying the test
        /// override; test-only, so a suite can branch deterministically.</summary>
        internal static bool SkipRepeatedEnabledForTests => SkipRepeatedEnabled;

        // Per-key stripes serializing the state lifecycle handoff (start/stop of
        // the consume loop and the install/removal of the wire subscription)
        // ACROSS state generations for the same registry key: a joiner arriving
        // while the previous generation is being torn down must not start a loop
        // for a dying state, and a removal must not race a fresh install. The
        // client assigns a fresh SID to each wire subscription, so key-level
        // serialization of the decision is enough; the two generations' SUB/UNSUB
        // commands never collide on the server.
        private static readonly StripedLocks NetworkLockStripes = new StripedLocks(64);

        /// <summary>Bounded backoff before the loop re-issues its SUBSCRIBE after
        /// a genuine wire-subscription end (mirrors the RTD 1s→30s schedule).</summary>
        private const int ReconnectBackoffMinMs = 1000;
        private const int ReconnectBackoffMaxMs = 30000;

        private static object NetworkGate(string key)
        {
            return NetworkLockStripes.For(key);
        }

        private readonly NatsConnectionManager _connections;
        private readonly ConcurrentDictionary<string, SubscriptionState> _states =
            new ConcurrentDictionary<string, SubscriptionState>();
        private readonly bool _skipRepeated;
        private long _nextListenerId;
        private volatile bool _disposed;

        public NatsSubscriptionManager(NatsConnectionManager connections)
        {
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _skipRepeated = SkipRepeatedEnabled;
        }

        /// <summary>Total number of physical subscription states
        /// (one per (endpoint, subject, queueGroup) with at least one listener).</summary>
        public int SubscriptionCount => _states.Count;

        /// <summary>Total number of registered listeners across all subscriptions.</summary>
        public int ListenerCount
        {
            get
            {
                int total = 0;
                foreach (var state in _states.Values)
                    total += state.Listeners.Count;
                return total;
            }
        }

        /// <summary>
        /// Number of registered listeners with the given caller-defined origin tag
        /// (for example "RTD" or "UDF"). Comparison is ordinal (case-sensitive);
        /// a null origin matches untagged listeners.
        /// </summary>
        public int ListenerCountWithOrigin(string origin)
        {
            int total = 0;
            foreach (var state in _states.Values)
            {
                foreach (var listener in state.Listeners.Values)
                {
                    if (listener != null && string.Equals(listener.Origin, origin, StringComparison.Ordinal))
                        total++;
                }
            }
            return total;
        }

        /// <summary>
        /// Number of subscription states that have at least one listener with the
        /// given caller-defined origin tag ("RTD"/"UDF"). Comparison is ordinal;
        /// a null origin matches untagged listeners.
        /// </summary>
        public int SubjectCountWithOrigin(string origin)
        {
            int total = 0;
            foreach (var state in _states.Values)
            {
                if (HasListenerWithOrigin(state, origin))
                    total++;
            }
            return total;
        }

        /// <summary>
        /// Whether any live subscription state for the endpoint still has
        /// listeners. Consumed by the connection manager: evicting/shutting down
        /// a connection with subscribers would silence them permanently.
        /// </summary>
        public bool HasActiveSubscribers(string endpoint)
        {
            foreach (var state in _states.Values)
            {
                if (string.Equals(state.Endpoint, endpoint, StringComparison.Ordinal) && !state.Listeners.IsEmpty)
                    return true;
            }
            return false;
        }

        private static bool HasListenerWithOrigin(SubscriptionState state, string origin)
        {
            foreach (var listener in state.Listeners.Values)
            {
                if (listener != null && string.Equals(listener.Origin, origin, StringComparison.Ordinal))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Registers a listener for <paramref name="subject"/> on
        /// <paramref name="endpoint"/>. <paramref name="endpoint"/> must already
        /// be resolved (AppConfig.ResolveRtdEndpoint/ResolveUdfEndpoint); it is an
        /// opaque registry key, never re-resolved here. A blank
        /// <paramref name="queueGroup"/> means no queue group.
        /// <para>
        /// The optional <paramref name="origin"/> is a caller-defined tag
        /// ("RTD"/"UDF") used only by the origin-scoped counters. Disposing the
        /// returned token removes this listener and, when it was the last one,
        /// stops the wire subscription; disposing is always safe and idempotent.
        /// </para>
        /// </summary>
        public IDisposable Subscribe(
            string endpoint,
            string subject,
            string queueGroup,
            Action<NatsMessage> onMessage,
            string origin = null)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(NatsSubscriptionManager));
            if (string.IsNullOrWhiteSpace(endpoint))
                throw new ArgumentException("endpoint is required", nameof(endpoint));
            if (string.IsNullOrWhiteSpace(subject))
                throw new ArgumentException("subject is required", nameof(subject));
            if (onMessage == null)
                throw new ArgumentNullException(nameof(onMessage));

            string normalizedQueue = string.IsNullOrWhiteSpace(queueGroup) ? null : queueGroup;
            string key = MakeKey(endpoint, subject, normalizedQueue);

            SubscriptionState state;
            long id;
            int disposedSpins = 0;

            while (true)
            {
                if (_disposed)
                    throw new ObjectDisposedException(nameof(NatsSubscriptionManager));

                state = _states.GetOrAdd(
                    key, _ => new SubscriptionState(key, endpoint, subject, normalizedQueue));

                lock (state.Sync)
                {
                    if (state.Disposed)
                    {
                        // The remover marks the state disposed BEFORE removing the
                        // mapping. Wait bounded for that removal; if it never comes
                        // (scheduler pause, or the remover lost a restore race),
                        // evict the stale mapping ourselves so this can never spin
                        // forever on the Excel thread.
                        if (++disposedSpins > 500)
                        {
                            RemoveStateEntry(state.Key, state);
                            if (_states.TryGetValue(key, out var stillCurrent) && ReferenceEquals(stillCurrent, state))
                                throw new InvalidOperationException("a disposed subscription state could not be recreated");
                        }
                        else
                        {
                            Thread.Yield();
                        }

                        continue; // releases the monitor and retries with a fresh state
                    }

                    bool wasActive = !state.Listeners.IsEmpty;
                    id = Interlocked.Increment(ref _nextListenerId);
                    state.Listeners[id] = new Listener(onMessage, origin);
                    state.RebuildSnapshot();

                    // A listener joining an already-active state never saw the
                    // payload currently held by the shared dedup marker, so clear
                    // it: the next publish - even an identical repeat - must be
                    // fanned out to the joiner (and everyone else) instead of
                    // being suppressed and leaving the new cell blank.
                    if (wasActive)
                        state.ResetDedup();
                }

                // Start the consume loop (and therefore the wire SUBSCRIBE) for the
                // first listener. The actual network call happens on the loop's
                // own task, off the Excel thread.
                EnsureLoop(state);

                if (_disposed)
                {
                    // Dispose() completed while we were subscribing: the manager is
                    // shutting down, so fail with the documented contract.
                    RollbackListener(state, id);
                    throw new ObjectDisposedException(nameof(NatsSubscriptionManager));
                }

                break;
            }

            logger.Debug(
                "Subscribe: endpoint=" + AppConfig.MaskNatsEndpoint(endpoint) +
                ", subject=" + subject +
                ", queue=" + (normalizedQueue ?? "<none>") +
                ", origin=" + (origin ?? "<null>") +
                ", listeners=" + state.Listeners.Count);

            RaiseListenerJoined(endpoint, subject, normalizedQueue);
            return new Registration(this, state, id);
        }

        /// <summary>
        /// Fans out a join notification with per-handler isolation: a failing
        /// consumer cleanup must not break an already successful Subscribe.
        /// </summary>
        private static void RaiseListenerJoined(string endpoint, string subject, string queueGroup)
        {
            var handlers = ListenerJoined;
            if (handlers == null)
                return;

            foreach (Action<string, string, string> handler in handlers.GetInvocationList())
            {
                try
                {
                    handler(endpoint, subject, queueGroup);
                }
                catch (Exception ex)
                {
                    logger.Error(
                        ex,
                        "ListenerJoined handler failed: endpoint=" + AppConfig.MaskNatsEndpoint(endpoint) +
                        ", subject=" + subject + ", queue=" + (queueGroup ?? "<none>"));
                }
            }
        }

        // ------------------------------------------------------------------
        // Consume loop (wire handoff + fan-out)
        // ------------------------------------------------------------------

        /// <summary>
        /// Starts the state's consume loop when it is not already running.
        /// Synchronous and cheap (a Task.Run): never blocks the Excel thread.
        /// Serialized per registry key with the striped network gate so a joiner
        /// racing a teardown cannot start a loop for a dying state.
        /// </summary>
        private void EnsureLoop(SubscriptionState state)
        {
            CancellationTokenSource cts = null;
            long epoch = 0;

            lock (NetworkGate(state.Key))
            {
                lock (state.Sync)
                {
                    if (_disposed || state.Disposed || state.LoopRunning)
                        return;

                    state.LoopRunning = true;
                    epoch = ++state.Epoch;
                    cts = new CancellationTokenSource();
                    state.LoopCts = cts;
                }
            }

            Task.Run(() => ConsumeLoopAsync(state, epoch, cts));
        }

        /// <summary>
        /// Owns the wire subscription for one state generation: it SUBSCRIBEs,
        /// reads the message channel, and (only when the wire subscription ends
        /// while the state is still current) re-issues the SUBSCRIBE after a
        /// capped backoff. Never throws - a dead connection or a throwing
        /// consumer must not take down the loop silently.
        /// </summary>
        private async Task ConsumeLoopAsync(SubscriptionState state, long epoch, CancellationTokenSource cts)
        {
            var ct = cts.Token;

            try
            {
                await ConsumeLoopCoreAsync(state, epoch, ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                logger.Error(
                    ex,
                    "ConsumeLoop: unhandled endpoint=" + AppConfig.MaskNatsEndpoint(state.Endpoint) +
                    ", subject=" + state.Subject);
            }
            finally
            {
                // Release the loop slot so a later listener can restart the loop
                // if this one ever exits unexpectedly. The restart covers the rare
                // race where a joiner's EnsureLoop observed LoopRunning == true
                // just before this exit.
                bool restart = false;
                lock (NetworkGate(state.Key))
                {
                    lock (state.Sync)
                    {
                        if (state.LoopCts == cts)
                        {
                            state.LoopRunning = false;
                            state.LoopCts = null;
                            restart = !_disposed && !state.Disposed && !state.Listeners.IsEmpty;
                        }
                    }
                }

                try
                {
                    cts.Dispose();
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "ConsumeLoop: releasing the loop token failed");
                }

                if (restart)
                    EnsureLoop(state);
            }
        }

        /// <summary>
        /// The loop body: SUBSCRIBE, read, and (only when the wire subscription
        /// ends while the state is still current) re-issue after a capped backoff.
        /// </summary>
        private async Task ConsumeLoopCoreAsync(SubscriptionState state, long epoch, CancellationToken ct)
        {
            int attempt = 0;

            while (!ct.IsCancellationRequested && !_disposed && !state.Disposed)
            {
                INatsSub<byte[]> sub = null;
                try
                {
                    var connection = _connections.GetConnection(state.Endpoint);
                    if (connection == null)
                        throw new InvalidOperationException(
                            "no NATS connection for endpoint " + AppConfig.MaskNatsEndpoint(state.Endpoint));

                    // The network subscribe deliberately runs OUTSIDE the lifecycle
                    // lock; byte[] keeps decoding under this manager's control.
                    sub = await connection
                        .SubscribeCoreAsync<byte[]>(state.Subject, state.QueueGroup, null, null, ct)
                        .ConfigureAwait(false);

                    // Install re-checks the epoch under both gates: a subscribe
                    // that raced a teardown must not become a ghost listener.
                    bool installed = false;
                    lock (NetworkGate(state.Key))
                    {
                        lock (state.Sync)
                        {
                            if (!_disposed && !state.Disposed && state.Epoch == epoch && !ct.IsCancellationRequested)
                            {
                                state.WireSub = sub;
                                installed = true;
                            }
                        }
                    }

                    if (!installed)
                    {
                        await SafeUnsubscribeAsync(sub).ConfigureAwait(false);
                        return;
                    }

                    attempt = 0;
                    await ReadLoopAsync(state, epoch, sub, ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    // Shutdown / teardown: the token registration already sent
                    // UNSUB and completed the channel.
                    return;
                }
                catch (Exception ex)
                {
                    logger.Error(
                        ex,
                        "ConsumeLoop: endpoint=" + AppConfig.MaskNatsEndpoint(state.Endpoint) +
                        ", subject=" + state.Subject);
                }
                finally
                {
                    lock (state.Sync)
                    {
                        if (ReferenceEquals(state.WireSub, sub))
                            state.WireSub = null;
                    }

                    if (sub != null && !ct.IsCancellationRequested)
                        await SafeUnsubscribeAsync(sub).ConfigureAwait(false);
                }

                // The wire subscription ended without us asking. Only re-issue when
                // this manager and this state generation are still current.
                if (ct.IsCancellationRequested || _disposed || state.Disposed || Volatile.Read(ref state.Epoch) != epoch)
                    return;

                attempt++;
                try
                {
                    await Task.Delay(BackoffMs(attempt), ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Drains the subscription channel. A message that passed the closed check
        /// is dispatched once; once the state is disposed (or the epoch moved) the
        /// loop stops draining, so a removed listener can never receive a payload
        /// again.
        /// </summary>
        private async Task ReadLoopAsync(SubscriptionState state, long epoch, INatsSub<byte[]> sub, CancellationToken ct)
        {
            var reader = sub.Msgs;

            while (!ct.IsCancellationRequested)
            {
                bool more;
                try
                {
                    more = await reader.WaitToReadAsync(ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return;
                }

                if (!more)
                    return; // channel completed: the wire subscription ended

                NatsMsg<byte[]> msg;
                while (reader.TryRead(out msg))
                {
                    if (ct.IsCancellationRequested || state.Disposed || _disposed)
                        return;
                    if (Volatile.Read(ref state.Epoch) != epoch)
                        return;

                    Dispatch(state, msg);
                }
            }
        }

        /// <summary>
        /// Duplicate suppression (ordinal bytes, literal subjects only) followed
        /// by the fan-out. Each listener is isolated: one throwing consumer never
        /// breaks the loop or its peers.
        /// </summary>
        private void Dispatch(SubscriptionState state, NatsMsg<byte[]> msg)
        {
            byte[] bytes = msg.Data;

            if (_skipRepeated && !state.Wildcard)
            {
                // Compare raw bytes BEFORE decoding: ordinal text semantics
                // ("007" != "7") and no lossy U+FFFD mapping of invalid payloads.
                // Best-effort under (unlikely) concurrency: two racing copies may
                // both fan out, which is harmless.
                lock (state.DedupSync)
                {
                    if (state.HasLastPayload && NatsResultFormatter.SameMessage(state.LastPayload, bytes))
                        return;
                    state.LastPayload = bytes;
                    state.HasLastPayload = true;
                }
            }

            var listeners = state.Snapshot;
            if (listeners.Length == 0)
                return;

            var message = BuildMessage(msg, bytes);
            for (int i = 0; i < listeners.Length; i++)
            {
                try
                {
                    listeners[i](message);
                }
                catch (Exception ex)
                {
                    logger.Error(
                        ex,
                        "HandleMessage: listener error endpoint=" + AppConfig.MaskNatsEndpoint(state.Endpoint) +
                        ", subject=" + state.Subject);
                }
            }
        }

        /// <summary>Projects a client message onto the frozen <see cref="NatsMessage"/> shape.</summary>
        private static NatsMessage BuildMessage(NatsMsg<byte[]> msg, byte[] bytes)
        {
            string reply = string.IsNullOrEmpty(msg.ReplyTo) ? null : msg.ReplyTo;
            ulong sequence = 0;
            DateTime? timestamp = null;
            IReadOnlyDictionary<string, string[]> headers = NoHeaders;

            var rawHeaders = msg.Headers;
            if (rawHeaders != null && rawHeaders.Count > 0)
            {
                var map = new Dictionary<string, string[]>(rawHeaders.Count, StringComparer.OrdinalIgnoreCase);
                foreach (var kv in rawHeaders)
                    map[kv.Key] = kv.Value.ToArray();
                headers = map;
                TryReadMetadata(map, out sequence, out timestamp);
            }

            string data = NatsResultFormatter.DecodePayload(bytes);
            return new NatsMessage(msg.Subject, reply, data, bytes, headers, sequence, timestamp);
        }

        /// <summary>
        /// Best-effort JetStream metadata carried in headers
        /// (<c>Nats-Sequence</c>, <c>Nats-Time-Stamp</c>). Absent or unparseable
        /// values are reported as 0/null; plain Core pub/sub has neither.
        /// </summary>
        private static void TryReadMetadata(
            IReadOnlyDictionary<string, string[]> headers, out ulong sequence, out DateTime? timestamp)
        {
            sequence = 0;
            timestamp = null;

            if (headers.TryGetValue("Nats-Sequence", out var seqValues)
                && seqValues != null && seqValues.Length > 0
                && ulong.TryParse(seqValues[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedSequence))
            {
                sequence = parsedSequence;
            }

            if (headers.TryGetValue("Nats-Time-Stamp", out var tsValues)
                && tsValues != null && tsValues.Length > 0
                && DateTime.TryParse(
                    tsValues[0],
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                    out var parsedTimestamp))
            {
                timestamp = parsedTimestamp;
            }
        }

        /// <summary>Fire-and-log teardown of a wire subscription.</summary>
        private static async Task SafeUnsubscribeAsync(INatsSub<byte[]> sub)
        {
            if (sub == null)
                return;

            try
            {
                await sub.UnsubscribeAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "SafeUnsubscribe: wire subscription teardown failed");
            }
        }

        /// <summary>1s, 2s, 4s … capped at 30s (mirrors the RTD backoff schedule).</summary>
        private static int BackoffMs(int attempt)
        {
            if (attempt <= 1)
                return ReconnectBackoffMinMs;

            long ms = ReconnectBackoffMinMs;
            for (int i = 1; i < attempt && ms < ReconnectBackoffMaxMs; i++)
                ms *= 2;
            return (int)Math.Min(ms, ReconnectBackoffMaxMs);
        }

        // ------------------------------------------------------------------
        // Listener removal, state release and dispose
        // ------------------------------------------------------------------

        /// <summary>
        /// Removes a single listener. When it was the last one the state is marked
        /// disposed and released; concurrent joiners keep the subscription alive.
        /// </summary>
        private void Remove(SubscriptionState state, long id)
        {
            if (!DetachListener(state, id))
                return;

            ReleaseState(state);
            logger.Debug(
                "Remove: endpoint=" + AppConfig.MaskNatsEndpoint(state.Endpoint) +
                ", subject=" + state.Subject +
                ", queue=" + (state.QueueGroup ?? "<none>") + " unsubscribed");
        }

        /// <summary>
        /// Removes a single listener after a failed/lost Subscribe. Only when this
        /// was the last listener is the state released.
        /// </summary>
        private void RollbackListener(SubscriptionState state, long id)
        {
            if (DetachListener(state, id))
                ReleaseState(state);
        }

        /// <summary>Removes the listener from the state; returns true when it was
        /// the last one (the state is then marked disposed).</summary>
        private static bool DetachListener(SubscriptionState state, long id)
        {
            bool lastListener;
            lock (state.Sync)
            {
                if (state.Disposed || !state.Listeners.TryRemove(id, out _))
                    return false;
                state.RebuildSnapshot();
                lastListener = state.Listeners.IsEmpty;
                if (lastListener)
                    state.Disposed = true;
            }
            return lastListener;
        }

        /// <summary>
        /// Drops the state's registry entry and stops its consume loop: the shared
        /// tail of <see cref="Remove"/> and <see cref="RollbackListener"/>, called
        /// only once the state's last listener is gone.
        /// </summary>
        private void ReleaseState(SubscriptionState state)
        {
            RemoveStateEntry(state.Key, state);
            StopLoop(state);
        }

        /// <summary>
        /// Stops the state's consume loop and invalidates its epoch so an
        /// in-flight install can never complete. Non-blocking: it only cancels the
        /// loop's token (the client's cancellation registration sends UNSUB and
        /// completes the channel), so it is safe to call from the Excel thread.
        /// Serialized per key with the striped network gate.
        /// </summary>
        private void StopLoop(SubscriptionState state)
        {
            CancellationTokenSource cts = null;

            lock (NetworkGate(state.Key))
            {
                lock (state.Sync)
                {
                    state.Epoch++; // invalidate any in-flight install for this generation
                    if (!state.LoopRunning && state.LoopCts == null)
                        return;

                    state.LoopRunning = false;
                    cts = state.LoopCts;
                    state.LoopCts = null;
                }
            }

            if (cts == null)
                return;

            try
            {
                cts.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // Already torn down.
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "StopLoop: cancelling the consume loop failed");
            }
        }

        /// <summary>
        /// Removes the registry entry only if it still maps to this exact state.
        /// .NET Framework has no atomic value-checked TryRemove, so the mapping is
        /// re-checked and a raced fresh entry that got evicted is restored with a
        /// bounded retry; a live state stays reachable.
        /// </summary>
        private void RemoveStateEntry(string key, SubscriptionState state)
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                if (!_states.TryGetValue(key, out var current))
                    return; // already gone (Dispose cleared it or a remover won)
                if (!ReferenceEquals(current, state))
                    return; // a fresh live state owns the key: leave it alone
                if (_states.TryRemove(key, out var removed))
                {
                    if (ReferenceEquals(removed, state))
                        return; // our state was removed cleanly

                    // The mapping changed between the check and the remove: put the
                    // evicted live entry back (retry if the restore races).
                    if (_states.TryAdd(key, removed))
                        return;
                }
            }

            logger.Debug("RemoveStateEntry: gave up restoring a raced entry");
        }

        /// <summary>
        /// Stops every consume loop and disposes the wire subscriptions. Cancels
        /// tokens only: it never blocks the calling (Excel) thread. Idempotent, and
        /// a loop that exits during shutdown can never re-add a payload (its
        /// generation is invalidated and <see cref="_disposed"/> is set first).
        /// </summary>
        public void Dispose()
        {
            _disposed = true;

            foreach (var state in _states.Values)
            {
                lock (state.Sync)
                {
                    state.Disposed = true;
                }

                StopLoop(state);
            }

            _states.Clear();
        }

        /// <summary>
        /// Collision-safe composite registry key for a physical subscription.
        /// Every part is length-prefixed with an ordinal separator, so an endpoint
        /// or subject that itself contains the separator (or a colon) cannot
        /// collide with another combination. Format:
        /// <c>{endpoint.Length}:{endpoint}\u0001{subject.Length}:{subject}\u0001{queue.Length}:{queue}</c>.
        /// A null/blank queue group is the empty part, keeping a plain subscriber
        /// distinct from a queue-group subscriber on the same subject.
        /// </summary>
        internal static string MakeKey(string endpoint, string subject, string queueGroup)
        {
            string ep = endpoint ?? string.Empty;
            string subj = subject ?? string.Empty;
            string q = queueGroup ?? string.Empty;

            return ep.Length.ToString(CultureInfo.InvariantCulture) + ":" + ep + "\u0001"
                 + subj.Length.ToString(CultureInfo.InvariantCulture) + ":" + subj + "\u0001"
                 + q.Length.ToString(CultureInfo.InvariantCulture) + ":" + q;
        }

        /// <summary>
        /// One registered listener plus its caller-defined origin tag ("RTD",
        /// "UDF", ...). The tag is never used for delivery, only for counting.
        /// </summary>
        private sealed class Listener
        {
            public readonly Action<NatsMessage> Callback;
            public readonly string Origin;

            public Listener(Action<NatsMessage> callback, string origin)
            {
                Callback = callback;
                Origin = origin;
            }
        }

        /// <summary>
        /// One physical subscription generation: its listeners, its lifecycle
        /// epoch, its consume loop and its shared dedup marker.
        /// </summary>
        private sealed class SubscriptionState
        {
            private static readonly Action<NatsMessage>[] EmptyListeners = new Action<NatsMessage>[0];

            public readonly string Key;
            public readonly string Endpoint;
            public readonly string Subject;
            public readonly string QueueGroup;
            public readonly bool Wildcard;
            public readonly ConcurrentDictionary<long, Listener> Listeners =
                new ConcurrentDictionary<long, Listener>();
            public readonly object Sync = new object();
            public readonly object DedupSync = new object();

            /// <summary>Latched before the registry entry is removed; read without
            /// the lock by the consume loop's per-message closed check.</summary>
            public volatile bool Disposed;

            /// <summary>Generation counter; bumped on loop start and on stop, so an
            /// install or fan-out from a stale generation is rejected.</summary>
            public long Epoch;

            public bool LoopRunning;
            public CancellationTokenSource LoopCts;
            public INatsSub<byte[]> WireSub;
            public byte[] LastPayload;
            public bool HasLastPayload;

            private volatile Action<NatsMessage>[] _snapshot = EmptyListeners;

            public SubscriptionState(string key, string endpoint, string subject, string queueGroup)
            {
                Key = key;
                Endpoint = endpoint;
                Subject = subject;
                QueueGroup = queueGroup;
                // Wildcards interleave different concrete subjects, so a pattern
                // subscription must never be deduplicated.
                Wildcard = subject.IndexOf('*') >= 0 || subject.IndexOf('>') >= 0;
            }

            /// <summary>Copy-on-write snapshot of the listener callbacks: the
            /// consume loop reads it without locks and without per-message
            /// allocations. Rebuilt under Sync whenever a listener changes.</summary>
            public Action<NatsMessage>[] Snapshot => _snapshot;

            public void RebuildSnapshot()
            {
                var snapshot = new Action<NatsMessage>[Listeners.Count];
                int index = 0;
                foreach (var kvp in Listeners)
                    snapshot[index++] = kvp.Value.Callback;
                _snapshot = snapshot;
            }

            /// <summary>
            /// Clears the shared duplicate-suppression marker so the next payload
            /// is treated as new for the whole subscription. Called when a listener
            /// joins an already-active state: the joiner never saw the suppressed
            /// payload, so an identical republish must be fanned out.
            /// </summary>
            public void ResetDedup()
            {
                lock (DedupSync)
                {
                    HasLastPayload = false;
                    LastPayload = null;
                }
            }
        }

        /// <summary>The disposable token returned by <see cref="Subscribe"/>.</summary>
        private sealed class Registration : IDisposable
        {
            private NatsSubscriptionManager _owner;
            private readonly SubscriptionState _state;
            private readonly long _id;

            public Registration(NatsSubscriptionManager owner, SubscriptionState state, long id)
            {
                _owner = owner;
                _state = state;
                _id = id;
            }

            public void Dispose()
            {
                var owner = Interlocked.Exchange(ref _owner, null);
                if (owner != null)
                    owner.Remove(_state, _id);
            }
        }
    }
}
