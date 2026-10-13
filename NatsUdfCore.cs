using ExcelDna.Integration;
using NLog;
using NATS.Client.Core;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace NATSExcel
{
    /// <summary>
    /// Core NATS user defined functions (DESIGN.md §7.1): publish,
    /// publish-JSON, publish-if-changed, publish-with-headers, request/reply,
    /// subject-latest, unsubscribe and the UDF subscription counter. Every write
    /// base has an additive <c>...NonVolatile</c> twin (DESIGN.md §7.7) so a
    /// worksheet recalculation does not re-send it.
    ///
    /// <para><b>Every write goes through <see cref="NatsUdfAsync.Run"/>.</b> That
    /// is the single entry point that decides between the synchronous
    /// passthrough (<c>AsyncWrites</c> off, the work runs on the Excel thread)
    /// and the per-endpoint FIFO queue with Excel-DNA's pending marker
    /// (<c>AsyncWrites</c> on). Reads that must return a value in-cell block
    /// only through <see cref="NatsConnectionManager.WaitBounded"/> /
    /// <see cref="NatsConnectionManager.WaitBoundedResult{T}"/>; the codebase
    /// never calls <c>.GetAwaiter().GetResult()</c> on the Excel thread
    /// (golden rule 10).</para>
    ///
    /// <para><b>Argument validation runs before any NATS access</b> and returns
    /// the terse, stable <c>Error: ...</c> text from <see cref="NatsUdfCommon"/>.
    /// A dangling cluster/server reference resolves to null (never a silent
    /// <c>localhost</c>) and becomes the <c>#CONFIG</c> error. Runtime failures
    /// go through <see cref="NatsUdfCommon.VerboseError"/>, which masks every
    /// credential.</para>
    ///
    /// <para>No NATS access is created here: connections and subscriptions come
    /// from <see cref="NatsRuntime"/> (golden rule 4).</para>
    /// </summary>
    public static class NatsUdfCore
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>Origin tag for the UDF-owned listeners (counted by
        /// <see cref="NatsUDFSubscriptionCount"/> and the RTD status surface).</summary>
        private const string OriginTag = "UDF";

        /// <summary>Slack added to a caller-supplied request timeout for the
        /// bounded wait, so the client's own timeout fires first and reports the
        /// real cause instead of our wait abandoning a live reply.</summary>
        private const int ResponseMarginMs = 500;

        /// <summary>Ceiling for a caller-supplied request timeout bound, so no
        /// formula can hold a thread arbitrarily long.</summary>
        private const int MaxRequestBoundMs = 60000;

        // ==================================================================
        // Construction
        // ==================================================================

        static NatsUdfCore()
        {
            // A listener that joins (initially or after a reconnect retry) never
            // saw the payload currently remembered by the PublishIfChanged dedup
            // marker. Clearing the marker on join keeps a returning listener from
            // being starved until the payload happens to change (PLAN S5.6,
            // mirrored from the sibling project's ListenerJoined handler).
            try
            {
                NatsSubscriptionManager.ListenerJoined += HandleListenerJoined;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "NatsUdfCore: subscribing to ListenerJoined failed");
            }
        }

        // ==================================================================
        // PublishIfChanged dedup state (bounded LRU, per endpoint/subject)
        // ==================================================================

        // Per-key stripes serialize the check -> publish -> store sequence without
        // a lock per key. The ListenerJoined purge takes the same stripe so it can
        // never interleave with an in-flight publish.
        private static readonly StripedLocks PublishLockStripes = new StripedLocks(64);

        /// <summary>
        /// Last payload successfully published per <c>(endpoint, subject)</c>.
        /// Bounded LRU keyed by <see cref="DedupKey"/>. Not readonly: the unit
        /// tests swap in a deterministic capacity.
        /// </summary>
        private static DedupCache _lastPublished = new DedupCache(CapacityFromConfig());

        private static object PublishLock(string key)
        {
            return PublishLockStripes.For(key);
        }

        private static int CapacityFromConfig()
        {
            try
            {
                int capacity = AppConfig.PublishDedupCacheSize;
                return capacity > 0 ? capacity : 10000;
            }
            catch
            {
                return 10000;
            }
        }

        /// <summary>Dedup marker cache, exposed internally so the listener-join
        /// clearing rules can be unit tested offline.</summary>
        internal static DedupCache LastPublishedForTests => _lastPublished;

        /// <summary>Test-only: replaces the shared cache with a deterministic
        /// capacity so the seeds used by the listener tests cannot be evicted by
        /// a small PublishDedupCacheSize from the machine's config file.</summary>
        internal static void ResetDedupCacheForTests(int capacity)
        {
            _lastPublished = new DedupCache(capacity > 0 ? capacity : 1);
        }

        /// <summary>
        /// Collision-safe dedup key for an <c>(endpoint, subject)</c> pair: the
        /// endpoint length is prefixed so an endpoint or subject that itself
        /// contains the separator cannot collide with another combination.
        /// Format: <c>{endpoint.Length}:{endpoint}\u0001{subject}</c>.
        /// </summary>
        internal static string DedupKey(string endpoint, string subject)
        {
            string ep = endpoint ?? string.Empty;
            return ep.Length.ToString(CultureInfo.InvariantCulture) + ":" + ep + "\u0001" + (subject ?? string.Empty);
        }

        /// <summary>Splits a <see cref="DedupKey"/> back into its parts; false for
        /// a foreign/malformed key so it is skipped instead of misattributed.</summary>
        internal static bool TryParseDedupKey(string key, out string endpoint, out string subject)
        {
            endpoint = null;
            subject = null;
            if (string.IsNullOrEmpty(key))
                return false;

            int colon = key.IndexOf(':');
            if (colon <= 0)
                return false;
            if (!int.TryParse(key.Substring(0, colon), NumberStyles.None, CultureInfo.InvariantCulture, out int length))
                return false;

            int start = colon + 1;
            if (length < 0 || start + length >= key.Length)
                return false;
            if (key[start + length] != '\u0001')
                return false;

            endpoint = key.Substring(start, length);
            subject = key.Substring(start + length + 1);
            return true;
        }

        /// <summary>
        /// Entry point for <see cref="NatsSubscriptionManager.ListenerJoined"/>:
        /// clears the publish dedup markers a newly joined listener invalidated.
        /// Never throws into Subscribe.
        /// </summary>
        internal static void HandleListenerJoined(string endpoint, string subject, string queueGroup)
        {
            try
            {
                ClearDedupForListener(endpoint, subject);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "NatsUdfCore.HandleListenerJoined: endpoint="
                    + NatsConnectionManager.MaskedEndpoint(endpoint) + ", subject=" + subject);
            }
        }

        /// <summary>
        /// Clears the publish dedup markers invalidated by a listener join: the
        /// exact <c>(endpoint, subject)</c> key for a literal subscription, or
        /// every marker of the same endpoint for a wildcard (a pattern
        /// interleaves concrete subjects, so an over-clear - one extra publish -
        /// is the safe direction).
        /// </summary>
        internal static void ClearDedupForListener(string endpoint, string subject)
        {
            string[] keys = _lastPublished.SnapshotKeys();
            if (keys.Length == 0)
                return;

            bool wildcard = subject != null && (subject.IndexOf('*') >= 0 || subject.IndexOf('>') >= 0);
            for (int i = 0; i < keys.Length; i++)
            {
                if (!TryParseDedupKey(keys[i], out var keyEndpoint, out var keySubject))
                    continue;
                if (!string.Equals(keyEndpoint, endpoint, StringComparison.Ordinal))
                    continue;
                if (!wildcard && !string.Equals(keySubject, subject, StringComparison.Ordinal))
                    continue;

                lock (PublishLock(keys[i]))
                    _lastPublished.Remove(keys[i]);
            }
        }

        // ==================================================================
        // SubjectLatest lifecycle state (ChannelLatest analogue)
        // ==================================================================

        /// <summary>One installed subject-latest listener plus the gate that
        /// serializes a callback against <c>Close</c>: a callback that already
        /// passed the closed check must not re-add a stale payload after the
        /// listener was removed.</summary>
        private sealed class SubjectListener
        {
            public IDisposable Token;
            public readonly object Sync = new object();
            private volatile bool _closed;

            public bool IsClosed => _closed;

            public void Close()
            {
                _closed = true;
            }
        }

        private static readonly ConcurrentDictionary<string, SubjectListener> _subjectListeners =
            new ConcurrentDictionary<string, SubjectListener>(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, string> _latestMessages =
            new ConcurrentDictionary<string, string>(StringComparer.Ordinal);

        // Per-key stripes for the SubjectLatest lifecycle (subscribe install vs
        // unsubscribe/reset). The lifecycle lock is NEVER held across the
        // subscribe registration, and the epoch counter (bumped by every
        // unsubscribe/reset) lets an install detect that a removal raced it and
        // decline instead of resurrecting a ghost listener.
        private static readonly StripedLocks LifecycleLockStripes = new StripedLocks(64);
        private static readonly ConcurrentDictionary<string, long> _subjectEpochs =
            new ConcurrentDictionary<string, long>(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, int> _subscribeWatchers =
            new ConcurrentDictionary<string, int>(StringComparer.Ordinal);
        private static long _subjectEpochBumps;

        private static object LifecycleLock(string key)
        {
            return LifecycleLockStripes.For(key);
        }

        private static long SubjectEpoch(string key)
        {
            return _subjectEpochs.TryGetValue(key, out var epoch) ? epoch : 0;
        }

        /// <summary>
        /// Keys with a SubjectLatest registration in flight. The epoch sweep must
        /// not prune one of these: deleting the entry would erase the evidence
        /// that an unsubscribe raced the registration, and since a brand-new
        /// key's epoch is 0 (== a pruned, absent entry), the install re-check
        /// would wrongly pass and leak a ghost listener.
        /// </summary>
        private static void BeginSubscribeWatch(string key)
        {
            _subscribeWatchers.AddOrUpdate(key, 1, (_, count) => count + 1);
        }

        private static void EndSubscribeWatch(string key)
        {
            while (_subscribeWatchers.TryGetValue(key, out var count))
            {
                if (count <= 1)
                {
                    if (_subscribeWatchers.TryRemove(key, out _))
                        return;
                }
                else if (_subscribeWatchers.TryUpdate(key, count - 1, count))
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Bumps the lifecycle epoch for a key. The bounded global prune (one
        /// sweep per 1024 bumps) drops epochs of keys with no live listener and
        /// no in-flight registration, so the map cannot grow without bound while
        /// distinct keys each stay at epoch 1.
        /// </summary>
        private static long BumpSubjectEpoch(string key)
        {
            long epoch = _subjectEpochs.AddOrUpdate(key, 1, (_, current) => current + 1);
            if ((Interlocked.Increment(ref _subjectEpochBumps) & 1023) == 0)
            {
                foreach (var kv in _subjectEpochs)
                {
                    if (_subscribeWatchers.ContainsKey(kv.Key))
                        continue;
                    if (!_subjectListeners.ContainsKey(kv.Key))
                        _subjectEpochs.TryRemove(kv.Key, out _);
                }
            }
            return epoch;
        }

        /// <summary>
        /// Drops the process-wide SubjectLatest state (listeners, cached values
        /// and dedup markers). A same-process add-in reload leaves the previous
        /// session's subscription tokens pointing at a dead runtime; without this
        /// a cached listener/value would make the reloaded add-in answer from
        /// stale state without ever re-subscribing.
        /// </summary>
        internal static void ResetAfterAddInReload()
        {
            foreach (var kv in _subjectListeners)
            {
                SubjectListener listener;
                lock (LifecycleLock(kv.Key))
                {
                    BumpSubjectEpoch(kv.Key);
                    if (!_subjectListeners.TryRemove(kv.Key, out listener))
                        continue;
                }
                try
                {
                    lock (listener.Sync)
                    {
                        listener.Close();
                        _latestMessages.TryRemove(kv.Key, out _);
                    }
                    listener.Token?.Dispose();
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "NatsUdfCore.ResetAfterAddInReload: disposing a listener failed");
                }
            }

            _latestMessages.Clear();

            foreach (string key in _lastPublished.SnapshotKeys())
            {
                lock (PublishLock(key))
                    _lastPublished.Remove(key);
            }

            _subjectEpochs.Clear();
            _subscribeWatchers.Clear();
        }

        // ==================================================================
        // Core NATS publishes (volatile bases + NonVolatile twins)
        // ==================================================================

        [ExcelFunction(Description = "Publishes a message to a NATS subject", IsVolatile = true)]
        public static object NatsUDFPublish(
            [ExcelArgument(Description = "NATS subject to publish to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUdfAsync.Run("NatsUDFPublish", optionalHost,
                new object[] { subject, message, optionalHost },
                () => PublishBody(subject, message, optionalHost));
        }

        [ExcelFunction(Description = "Publishes a message to a NATS subject; runs once per entry/argument change (non-volatile)")]
        public static object NatsUDFPublishNonVolatile(
            [ExcelArgument(Description = "NATS subject to publish to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUDFPublish(subject, message, optionalHost);
        }

        [ExcelFunction(Description = "Publishes an Excel range to a NATS subject as JSON", IsVolatile = true)]
        public static object NatsUDFPublishJSON(
            [ExcelArgument(Description = "NATS subject to publish to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Excel range to serialize as JSON")] object[,] range,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUdfAsync.Run("NatsUDFPublishJSON", optionalHost,
                new object[] { subject, range, optionalHost },
                () => PublishJSONBody(subject, range, optionalHost));
        }

        [ExcelFunction(Description = "Publishes an Excel range to a NATS subject as JSON; runs once per entry/argument change (non-volatile)")]
        public static object NatsUDFPublishJSONNonVolatile(
            [ExcelArgument(Description = "NATS subject to publish to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Excel range to serialize as JSON")] object[,] range,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUDFPublishJSON(subject, range, optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message to a NATS subject unless it is unchanged since the last delivery", IsVolatile = true)]
        public static object NatsUDFPublishIfChanged(
            [ExcelArgument(Description = "NATS subject to publish to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUdfAsync.Run("NatsUDFPublishIfChanged", optionalHost,
                new object[] { subject, message, optionalHost },
                () => PublishIfChangedBody(subject, message, optionalHost));
        }

        [ExcelFunction(Description = "Publishes a message to a NATS subject unless it is unchanged since the last delivery; runs once per entry/argument change (non-volatile)")]
        public static object NatsUDFPublishIfChangedNonVolatile(
            [ExcelArgument(Description = "NATS subject to publish to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUDFPublishIfChanged(subject, message, optionalHost);
        }

        [ExcelFunction(Description = "Publishes a message with key/value headers to a NATS subject", IsVolatile = true)]
        public static object NatsUDFPublishWithHeaders(
            [ExcelArgument(Description = "NATS subject to publish to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Two-column or two-row range of header name/value pairs")] object[,] headers,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUdfAsync.Run("NatsUDFPublishWithHeaders", optionalHost,
                new object[] { subject, message, headers, optionalHost },
                () => PublishWithHeadersBody(subject, message, headers, optionalHost));
        }

        [ExcelFunction(Description = "Publishes a message with key/value headers to a NATS subject; runs once per entry/argument change (non-volatile)")]
        public static object NatsUDFPublishWithHeadersNonVolatile(
            [ExcelArgument(Description = "NATS subject to publish to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Message content to publish")] object message,
            [ExcelArgument(Description = "Two-column or two-row range of header name/value pairs")] object[,] headers,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUDFPublishWithHeaders(subject, message, headers, optionalHost);
        }

        // ==================================================================
        // Core NATS request/reply (volatile bases + NonVolatile twins)
        // ==================================================================

        [ExcelFunction(Description = "Sends a request to a NATS subject and returns the reply text", IsVolatile = true)]
        public static object NatsUDFRequest(
            [ExcelArgument(Description = "NATS subject to send the request to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Request payload")] object message,
            [ExcelArgument(Description = "Optional reply timeout in milliseconds")] object timeoutMs = null,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUdfAsync.Run("NatsUDFRequest", optionalHost,
                new object[] { subject, message, timeoutMs, optionalHost },
                () => RequestBody(subject, message, timeoutMs, optionalHost));
        }

        [ExcelFunction(Description = "Sends a request to a NATS subject and returns the reply text; runs once per entry/argument change (non-volatile)")]
        public static object NatsUDFRequestNonVolatile(
            [ExcelArgument(Description = "NATS subject to send the request to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Request payload")] object message,
            [ExcelArgument(Description = "Optional reply timeout in milliseconds")] object timeoutMs = null,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUDFRequest(subject, message, timeoutMs, optionalHost);
        }

        [ExcelFunction(Description = "Sends an Excel range as a JSON request to a NATS subject and returns the reply text", IsVolatile = true)]
        public static object NatsUDFRequestJSON(
            [ExcelArgument(Description = "NATS subject to send the request to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Excel range to serialize as the JSON request")] object[,] range,
            [ExcelArgument(Description = "Optional reply timeout in milliseconds")] object timeoutMs = null,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUdfAsync.Run("NatsUDFRequestJSON", optionalHost,
                new object[] { subject, range, timeoutMs, optionalHost },
                () => RequestJSONBody(subject, range, timeoutMs, optionalHost));
        }

        [ExcelFunction(Description = "Sends an Excel range as a JSON request to a NATS subject and returns the reply text; runs once per entry/argument change (non-volatile)")]
        public static object NatsUDFRequestJSONNonVolatile(
            [ExcelArgument(Description = "NATS subject to send the request to (concrete, no wildcards)")] object subject,
            [ExcelArgument(Description = "Excel range to serialize as the JSON request")] object[,] range,
            [ExcelArgument(Description = "Optional reply timeout in milliseconds")] object timeoutMs = null,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUDFRequestJSON(subject, range, timeoutMs, optionalHost);
        }

        // ==================================================================
        // SubjectLatest / unsubscribe / count
        // ==================================================================

        [ExcelFunction(Description = "Reads the latest message seen on a NATS subject (subscribes on first use)", IsVolatile = true)]
        public static object NatsUDFSubjectLatest(
            [ExcelArgument(Description = "NATS subject to read (wildcards allowed)")] object subject,
            [ExcelArgument(Description = "Optional queue group for the subscription")] object queueGroup = null,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            string error;
            string latest = SubjectLatestText(subject, queueGroup, optionalHost, out error);
            return error ?? latest ?? NatsUdfCommon.NullValue;
        }

        [ExcelFunction(Description = "Reads the latest message seen on a NATS subject and parses it as JSON into cells", IsVolatile = true)]
        public static object[,] NatsUDFSubjectLatestJSON(
            [ExcelArgument(Description = "NATS subject to read (wildcards allowed)")] object subject,
            [ExcelArgument(Description = "Optional queue group for the subscription")] object queueGroup = null,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            string error;
            string latest = SubjectLatestText(subject, queueGroup, optionalHost, out error);
            if (error != null)
                return SingleCell(error);
            if (string.IsNullOrEmpty(latest) || latest == NatsUdfCommon.NullValue)
                return SingleCell(NatsUdfCommon.NullValue);
            return NatsJson.JsonToMatrix(latest);
        }

        [ExcelFunction(Description = "Unsubscribes a NATS subject previously subscribed with NatsUDFSubjectLatest", IsVolatile = true)]
        public static object NatsUDFSubjectUnsubscribe(
            [ExcelArgument(Description = "NATS subject to unsubscribe (wildcards allowed)")] object subject,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUdfAsync.Run("NatsUDFSubjectUnsubscribe", optionalHost,
                new object[] { subject, optionalHost },
                () => SubjectUnsubscribeBody(subject, optionalHost));
        }

        [ExcelFunction(Description = "Unsubscribes a NATS subject previously subscribed with NatsUDFSubjectLatest; runs once per entry/argument change (non-volatile)")]
        public static object NatsUDFSubjectUnsubscribeNonVolatile(
            [ExcelArgument(Description = "NATS subject to unsubscribe (wildcards allowed)")] object subject,
            [ExcelArgument(Description = "Optional cluster name, server alias or nats:// URL")] object optionalHost = null)
        {
            return NatsUDFSubjectUnsubscribe(subject, optionalHost);
        }

        [ExcelFunction(Description = "Returns the number of active UDF NATS subscriptions", IsVolatile = true)]
        public static double NatsUDFSubscriptionCount()
        {
            // Like the other status helpers, a shutdown/reload race returns the
            // neutral value instead of throwing into Excel.
            try
            {
                double count = NatsRuntime.Subscriptions.ListenerCountWithOrigin(OriginTag);
                if (logger.IsDebugEnabled)
                    logger.Debug("NatsUDFSubscriptionCount: listeners=" + count.ToString(CultureInfo.InvariantCulture));
                return count;
            }
            catch
            {
                return 0;
            }
        }

        // ==================================================================
        // Write bodies (run inside NatsUdfAsync.Run's work delegate)
        // ==================================================================

        private static object PublishBody(object subject, object message, object optionalHost)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: false);
            if (error != null)
                return error;

            string subjectText = NatsUdfCommon.ToNatsString(subject);
            if (!TryPayload(message, out var payload, out error))
                return error;

            string endpoint = ResolveEndpoint(optionalHost, out error);
            if (error != null)
                return error;

            return Send(endpoint, subjectText, payload, null, "publish subject='" + subjectText + "'", replyDependent: false);
        }

        private static object PublishJSONBody(object subject, object[,] range, object optionalHost)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: false);
            if (error != null)
                return error;

            string subjectText = NatsUdfCommon.ToNatsString(subject);
            error = NatsUdfCommon.RequireRange(range);
            if (error != null)
                return error;

            string json = NatsJson.MatrixToJson(range);
            if (json == null || json.StartsWith(NatsUdfCommon.ErrorPrefix, StringComparison.Ordinal))
                return json;

            string endpoint = ResolveEndpoint(optionalHost, out error);
            if (error != null)
                return error;

            return Send(endpoint, subjectText, json, null, "publishJSON subject='" + subjectText + "'", replyDependent: false);
        }

        private static object PublishIfChangedBody(object subject, object message, object optionalHost)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: false);
            if (error != null)
                return error;

            string subjectText = NatsUdfCommon.ToNatsString(subject);
            if (!TryPayload(message, out var payload, out error))
                return error;

            string endpoint = ResolveEndpoint(optionalHost, out error);
            if (error != null)
                return error;

            string context = "publishIfChanged subject='" + subjectText + "'";
            string key = DedupKey(endpoint, subjectText);

            // Check -> publish -> store is serialized per key so concurrent
            // recalculations cannot invert each other, and it shares the stripe
            // with the ListenerJoined purge.
            lock (PublishLock(key))
            {
                if (_lastPublished.TryGet(key, out var previous)
                    && string.Equals(previous, payload, StringComparison.Ordinal))
                {
                    if (logger.IsDebugEnabled)
                        logger.Debug(context + " unchanged endpoint=" + NatsConnectionManager.MaskedEndpoint(endpoint));
                    return NatsUdfCommon.NoChange;
                }

                object result = Send(endpoint, subjectText, payload, null, context, replyDependent: false);

                // Remember the payload only after a CONFIRMED delivery. Under
                // SyncWrite=fireforget (the shipped default) the result is a
                // fire-and-forget marker: the bytes were handed to the client's
                // writer task, not acknowledged by the server, and a late failure
                // is only logged (Send -> Observe). Storing it as dedup state
                // would make every later recalculation of an unchanged value
                // answer "No change" forever, so a publish that was actually
                // dropped (server down) could never be retried and the feed would
                // starve. DESIGN.md 8.1: the marker is remembered only after a
                // confirmed delivery (the awaited sync "OK"). A fire-and-forget
                // write therefore leaves the cached marker untouched (a prior
                // confirmed value must keep deduplicating) and only a real
                // failure clears it.
                if (IsWriteSuccess(result) && !IsFireAndForgetResult(result))
                    _lastPublished.Set(key, payload);
                else if (!IsWriteSuccess(result))
                    _lastPublished.Remove(key);

                return result;
            }
        }

        private static object PublishWithHeadersBody(object subject, object message, object[,] headers, object optionalHost)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: false);
            if (error != null)
                return error;

            string subjectText = NatsUdfCommon.ToNatsString(subject);
            if (!TryPayload(message, out var payload, out error))
                return error;

            error = NatsUdfCommon.RequireRange(headers, "headers");
            if (error != null)
                return error;

            NatsHeaders natsHeaders;
            try
            {
                natsHeaders = BuildHeaders(NatsUdfCommon.FlattenPairRange(headers));
            }
            catch (ArgumentException ex)
            {
                return NatsUdfCommon.Fail(ex.Message);
            }

            string endpoint = ResolveEndpoint(optionalHost, out error);
            if (error != null)
                return error;

            return Send(endpoint, subjectText, payload, natsHeaders,
                "publishWithHeaders subject='" + subjectText + "'", replyDependent: false);
        }

        private static object RequestBody(object subject, object message, object timeoutMs, object optionalHost)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: false);
            if (error != null)
                return error;

            string subjectText = NatsUdfCommon.ToNatsString(subject);
            if (!TryPayload(message, out var payload, out error))
                return error;

            return RequestValidated(subjectText, payload, timeoutMs, optionalHost);
        }

        private static object RequestJSONBody(object subject, object[,] range, object timeoutMs, object optionalHost)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: false);
            if (error != null)
                return error;

            string subjectText = NatsUdfCommon.ToNatsString(subject);
            error = NatsUdfCommon.RequireRange(range);
            if (error != null)
                return error;

            string json = NatsJson.MatrixToJson(range);
            if (json == null || json.StartsWith(NatsUdfCommon.ErrorPrefix, StringComparison.Ordinal))
                return json;

            return RequestValidated(subjectText, json, timeoutMs, optionalHost);
        }

        /// <summary>
        /// Shared request/reply tail: subject and payload are already validated.
        /// Under <c>fireforget-all</c> the request is a plain Core publish that
        /// is not awaited and returns the ": all" marker (there is no reply to
        /// report); otherwise the call waits for the single reply through the
        /// bounded helper and maps a timeout / no-responders sentinel to its
        /// frozen marker.
        /// </summary>
        private static object RequestValidated(string subjectText, string payload, object timeoutMs, object optionalHost)
        {
            string context = "request subject='" + subjectText + "'";
            string endpoint = null;
            try
            {
                string error = ParseTimeout(timeoutMs, out int? requestedMs);
                if (error != null)
                    return error;

                endpoint = ResolveEndpoint(optionalHost, out error);
                if (error != null)
                    return error;

                string mode = AppConfig.SyncWrite;
                bool fireAndForget = NatsUdfCommon.ShouldFireAndForget(mode, replyDependent: true);
                var connection = NatsRuntime.Connections.GetConnection(endpoint);

                if (fireAndForget)
                {
                    Observe(connection.PublishAsync(subjectText, payload).AsTask(), context, endpoint, warnOnFailure: true);
                    if (logger.IsDebugEnabled)
                        logger.Debug(context + " sent fire-and-forget endpoint=" + NatsConnectionManager.MaskedEndpoint(endpoint));
                    return NatsUdfCommon.FireAndForgetMarker(mode);
                }

                NatsSubOpts opts = BuildRequestOptions(requestedMs);
                Task<NatsMsg<byte[]>> task = connection
                    .RequestAsync<string, byte[]>(subjectText, payload, null, null, null, null, opts, CancellationToken.None)
                    .AsTask();

                NatsMsg<byte[]> reply;
                try
                {
                    reply = NatsConnectionManager.WaitBoundedResult(task, RequestBoundMs(requestedMs));
                }
                catch (Exception ex)
                {
                    // Observe a request our bounded wait abandoned so its late
                    // failure is never an unobserved task exception.
                    if (!task.IsCompleted)
                        Observe(task, context, endpoint, warnOnFailure: false);

                    string marker = ClassifyRequestFailure(ex);
                    if (marker != null)
                        return marker;
                    return RuntimeError(context, ex);
                }

                if (reply.HasNoResponders)
                    return NatsUdfCommon.NoResponders;

                string text = NatsResultFormatter.DecodePayload(reply.Data) ?? NatsUdfCommon.NullValue;
                if (logger.IsDebugEnabled)
                    logger.Debug(context + " replied bytes=" + (reply.Data == null ? 0 : reply.Data.Length)
                        + " endpoint=" + NatsConnectionManager.MaskedEndpoint(endpoint));
                return NatsResultFormatter.BoundCellText(text);
            }
            catch (Exception ex)
            {
                return RuntimeError(context, ex);
            }
        }

        /// <summary>
        /// "Never fire-and-forget": the local listener removal and the cached
        /// value are touched deterministically regardless of the SyncWrite mode.
        /// </summary>
        private static object SubjectUnsubscribeBody(object subject, object optionalHost)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: true);
            if (error != null)
                return error;

            string subjectText = NatsUdfCommon.ToNatsString(subject);
            string endpoint = ResolveEndpoint(optionalHost, out error);
            if (error != null)
                return error;

            // NatsUDFSubjectUnsubscribe carries no queue group (DESIGN.md §7.1),
            // so it removes the plain listener that NatsUDFSubjectLatest
            // registered for this (endpoint, subject).
            string key = NatsSubscriptionManager.MakeKey(endpoint, subjectText, null);
            SubjectListener removed = null;
            lock (LifecycleLock(key))
            {
                // Bump first: a concurrent SubjectLatest that captured the old
                // epoch declines to install its listener after this removal
                // (otherwise the unsubscribe could report success while a racing
                // listener stays live).
                BumpSubjectEpoch(key);
                _subjectListeners.TryRemove(key, out removed);
            }

            if (removed != null)
            {
                // Close and drop the cached value under the listener lock so a
                // callback that already passed the closed check cannot re-add a
                // pre-unsubscribe payload after the removal.
                lock (removed.Sync)
                {
                    removed.Close();
                    _latestMessages.TryRemove(key, out _);
                }
                try
                {
                    removed.Token?.Dispose();
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "NatsUDFSubjectUnsubscribe: disposing the listener failed");
                }

                // Drop the publish dedup marker with the listener so a rejoining
                // listener receives the next publish even when the payload did
                // not change.
                ClearDedupForListener(endpoint, subjectText);
            }

            if (logger.IsDebugEnabled)
                logger.Debug("NatsUDFSubjectUnsubscribe: subject='" + subjectText
                    + "' endpoint=" + NatsConnectionManager.MaskedEndpoint(endpoint) + " unsubscribed");
            return "Subject '" + subjectText + "' unsubscribed successfully.";
        }

        // ==================================================================
        // SubjectLatest
        // ==================================================================

        /// <summary>
        /// Returns the cached latest payload of a subject (subscribing once if
        /// needed), or null with <paramref name="error"/> set. The marker is
        /// <see cref="NatsUdfCommon.NullValue"/> before the first message.
        /// </summary>
        private static string SubjectLatestText(object subject, object queueGroup, object optionalHost, out string error)
        {
            error = NatsUdfCommon.RequireSubject(subject, allowWildcard: true);
            if (error != null)
                return null;

            string subjectText = NatsUdfCommon.ToNatsString(subject);
            string queue = NormalizeQueueGroup(queueGroup, out error);
            if (error != null)
                return null;
            string endpoint = ResolveEndpoint(optionalHost, out error);
            if (error != null)
                return null;

            try
            {
                string key = NatsSubscriptionManager.MakeKey(endpoint, subjectText, queue);
                EnsureSubjectListener(endpoint, subjectText, queue, key);
                return _latestMessages.TryGetValue(key, out var latest) ? latest : NatsUdfCommon.NullValue;
            }
            catch (Exception ex)
            {
                error = RuntimeError("subjectLatest subject='" + subjectText + "'", ex);
                return null;
            }
        }

        /// <summary>
        /// Registers the subject-latest listener once. The epoch re-check after
        /// the (non-blocking) registration is what prevents a ghost listener
        /// when an unsubscribe/reset raced it.
        /// </summary>
        private static void EnsureSubjectListener(string endpoint, string subject, string queueGroup, string key)
        {
            long epochBefore = 0;
            bool needsSubscribe = false;
            lock (LifecycleLock(key))
            {
                if (!_subjectListeners.ContainsKey(key))
                {
                    epochBefore = SubjectEpoch(key);
                    // Register the in-flight window INSIDE the same lock as the
                    // epoch read: the bounded prune does not hold this lock, so
                    // taking the watch afterwards could let a removal + prune
                    // erase the entry and make the install re-check (epoch == 0)
                    // wrongly pass.
                    BeginSubscribeWatch(key);
                    needsSubscribe = true;
                }
            }

            if (!needsSubscribe)
                return;

            var listener = new SubjectListener();
            try
            {
                listener.Token = NatsRuntime.Subscriptions.Subscribe(
                    endpoint, subject, queueGroup,
                    message =>
                    {
                        lock (listener.Sync)
                        {
                            if (listener.IsClosed)
                                return;
                            _latestMessages[key] = message != null && message.Data != null
                                ? message.Data
                                : NatsUdfCommon.NullValue;
                        }
                    },
                    OriginTag);

                bool installed = false;
                lock (LifecycleLock(key))
                {
                    // Install only when no unsubscribe/reset raced the
                    // registration: an epoch change means the removal already
                    // completed and this listener must not come back alive.
                    if (SubjectEpoch(key) == epochBefore && _subjectListeners.TryAdd(key, listener))
                        installed = true;
                }

                if (!installed)
                {
                    // Lost a race (another thread registered first) or an
                    // unsubscribe/reset won: close before disposing so this
                    // losing listener never writes.
                    lock (listener.Sync)
                        listener.Close();
                    listener.Token?.Dispose();
                    return;
                }

                if (logger.IsDebugEnabled)
                    logger.Debug("NatsUDFSubjectLatest: subscribed subject='" + subject
                        + "' queue=" + (queueGroup ?? "<none>")
                        + " endpoint=" + NatsConnectionManager.MaskedEndpoint(endpoint));
            }
            finally
            {
                EndSubscribeWatch(key);
            }
        }

        // ==================================================================
        // Shared send / request helpers
        // ==================================================================

        /// <summary>
        /// Sends one payload: fire-and-forget per the configured mode, or a
        /// bounded synchronous publish that waits for the client to write it.
        /// </summary>
        private static object Send(
            string endpoint, string subject, string payload, NatsHeaders headers,
            string context, bool replyDependent)
        {
            try
            {
                string mode = AppConfig.SyncWrite;
                bool fireAndForget = NatsUdfCommon.ShouldFireAndForget(mode, replyDependent);
                var connection = NatsRuntime.Connections.GetConnection(endpoint);
                Task task = connection.PublishAsync(subject, payload, headers).AsTask();

                if (fireAndForget)
                {
                    // Do not await: the caller gets the marker immediately and the
                    // client writes the bytes on its own writer task. The
                    // continuation observes a late failure so it is never an
                    // unobserved task exception.
                    Observe(task, context, endpoint, warnOnFailure: true);
                    if (logger.IsDebugEnabled)
                        logger.Debug(context + " sent fire-and-forget endpoint=" + NatsConnectionManager.MaskedEndpoint(endpoint));
                    return NatsUdfCommon.FireAndForgetMarker(mode);
                }

                if (!NatsConnectionManager.WaitBounded(task, NatsConnectionManager.ResponseTimeoutMs(NatsPool.Udf)))
                {
                    Observe(task, context, endpoint, warnOnFailure: false);
                    return NatsUdfCommon.VerboseError(context,
                        "the NATS client did not complete the write within "
                            + NatsConnectionManager.ResponseTimeoutMs(NatsPool.Udf).ToString(CultureInfo.InvariantCulture) + "ms",
                        "check the connection or raise Default.timeout (NATSExcel.json)");
                }

                if (task.IsCanceled)
                    return NatsUdfCommon.VerboseError(context, "the write was cancelled", null);
                if (task.IsFaulted)
                    return RuntimeError(context, task.Exception);

                if (logger.IsDebugEnabled)
                    logger.Debug(context + " published endpoint=" + NatsConnectionManager.MaskedEndpoint(endpoint));
                return NatsUdfCommon.Ok;
            }
            catch (Exception ex)
            {
                return RuntimeError(context, ex);
            }
        }

        private static NatsSubOpts BuildRequestOptions(int? requestedMs)
        {
            // Timeout is init-only, so an explicit bound needs its own initializer.
            if (requestedMs.HasValue)
            {
                return new NatsSubOpts
                {
                    ThrowIfNoResponders = true,
                    Timeout = TimeSpan.FromMilliseconds(requestedMs.Value),
                };
            }
            return new NatsSubOpts { ThrowIfNoResponders = true };
        }

        /// <summary>
        /// Bounded wait for a reply: the caller's timeout plus slack, never less
        /// than the pool response timeout and never above the ceiling.
        /// </summary>
        private static int RequestBoundMs(int? requestedMs)
        {
            int poolBound = NatsConnectionManager.ResponseTimeoutMs(NatsPool.Udf);
            if (!requestedMs.HasValue)
                return poolBound;

            long bound = (long)requestedMs.Value + ResponseMarginMs;
            if (bound < poolBound) bound = poolBound;
            if (bound > MaxRequestBoundMs) bound = MaxRequestBoundMs;
            return (int)bound;
        }

        /// <summary>
        /// Maps a request failure to its frozen marker, or null when it is a real
        /// error the caller must report verbosely. The message check is a
        /// defensive fallback for a client version that surfaces the 503
        /// no-responders sentinel as a plain <see cref="NatsException"/>.
        /// </summary>
        private static string ClassifyRequestFailure(Exception ex)
        {
            if (ex is TimeoutException || ex is NatsTimeoutException || ex is OperationCanceledException)
                return NatsUdfCommon.Timeout;
            // The client throws NatsNoReplyException when nothing replied within
            // the request timeout (see the RTD REQ path, which maps it the same
            // way). Without this the cell showed the raw
            // "Error: request subject='...': No reply received" instead of the
            // frozen TIMEOUT marker.
            if (ex is NatsNoReplyException)
                return NatsUdfCommon.Timeout;
            if (ex is NatsNoRespondersException)
                return NatsUdfCommon.NoResponders;

            string message = ex.Message;
            if (!string.IsNullOrEmpty(message) && message.IndexOf("no responders", StringComparison.OrdinalIgnoreCase) >= 0)
                return NatsUdfCommon.NoResponders;
            if (!string.IsNullOrEmpty(message) && message.IndexOf("no reply", StringComparison.OrdinalIgnoreCase) >= 0)
                return NatsUdfCommon.Timeout;
            return null;
        }

        /// <summary>
        /// Groups a flattened pair range into a header set with one entry per
        /// name (duplicates become multiple values). A fresh instance per publish
        /// is required: <see cref="NatsHeaders"/> is not thread-safe.
        /// </summary>
        internal static NatsHeaders BuildHeaders(List<KeyValuePair<string, string>> pairs)
        {
            var grouped = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var order = new List<string>();
            foreach (var pair in pairs)
            {
                if (!grouped.TryGetValue(pair.Key, out var values))
                {
                    values = new List<string>();
                    grouped[pair.Key] = values;
                    order.Add(pair.Key);
                }
                values.Add(pair.Value);
            }

            var headers = new NatsHeaders(grouped.Count);
            for (int i = 0; i < order.Count; i++)
                headers.Add(order[i], grouped[order[i]].ToArray());
            return headers;
        }

        // ==================================================================
        // Validation / conversion helpers
        // ==================================================================

        /// <summary>
        /// Converts a message argument to the invariant text written to NATS. A
        /// missing cell becomes an empty payload (a legitimate publish); a
        /// multi-cell range or an Excel error cell is the frozen terse error.
        /// </summary>
        private static bool TryPayload(object message, out string payload, out string error)
        {
            payload = null;
            error = null;
            try
            {
                payload = NatsUdfCommon.ToNatsString(message) ?? string.Empty;
                return true;
            }
            catch (ArgumentException ex)
            {
                error = NatsUdfCommon.Fail(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Validates an optional <c>timeoutMs</c> argument: absent, blank or zero
        /// means "use the pool default"; a fraction, non-number, negative or
        /// out-of-range value or an Excel error cell is the terse argument error.
        /// </summary>
        private static string ParseTimeout(object timeoutMs, out int? value)
        {
            value = null;
            if (timeoutMs == null || timeoutMs is ExcelMissing || timeoutMs is ExcelEmpty)
                return null;

            if (!NatsUdfCommon.TryToInt(timeoutMs, out int parsed))
                return NatsUdfCommon.Fail("timeoutMs must be a whole number of milliseconds");
            if (parsed < 0)
                return NatsUdfCommon.Fail("timeoutMs must not be negative");
            if (parsed == 0)
                return null;

            value = parsed;
            return null;
        }

        /// <summary>
        /// Normalizes an optional queue group: blank/omitted means "no queue
        /// group" (a plain subscriber). A multi-cell range or an Excel error cell
        /// is the frozen terse argument error.
        /// </summary>
        private static string NormalizeQueueGroup(object queueGroup, out string error)
        {
            error = null;
            if (queueGroup == null || queueGroup is ExcelMissing || queueGroup is ExcelEmpty)
                return null;
            try
            {
                string text = NatsUdfCommon.ToNatsString(queueGroup);
                return string.IsNullOrWhiteSpace(text) ? null : text;
            }
            catch (ArgumentException ex)
            {
                error = NatsUdfCommon.Fail(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Resolves the optional host argument to a concrete endpoint. A blank or
        /// omitted argument uses the UDF default cluster; a non-text value is the
        /// terse argument error; a named reference that resolves to nothing is
        /// the <c>#CONFIG</c> error, never a silent <c>localhost</c>
        /// (DESIGN.md §8.1).
        /// </summary>
        private static string ResolveEndpoint(object optionalHost, out string error)
        {
            error = null;
            string reference = null;
            string display = "(default)";

            if (optionalHost != null && !(optionalHost is ExcelMissing) && !(optionalHost is ExcelEmpty))
            {
                if (!(optionalHost is string host))
                {
                    error = NatsUdfCommon.Fail("host must be a text value");
                    return null;
                }
                if (!string.IsNullOrWhiteSpace(host))
                {
                    reference = host.Trim();
                    display = reference;
                }
            }

            string endpoint;
            try
            {
                endpoint = AppConfig.ResolveUdfEndpoint(reference);
            }
            catch (Exception ex)
            {
                // Config access must never fault a formula; report it as the
                // usual masked runtime error (a bad file is a #CONFIG-class
                // problem, not a raw exception in a cell).
                error = RuntimeError("host=" + NatsUdfCommon.ShortForLog(display, 128), ex);
                return null;
            }

            if (string.IsNullOrWhiteSpace(endpoint))
            {
                error = NatsUdfCommon.VerboseError(
                    "host=" + NatsUdfCommon.ShortForLog(display, 128),
                    NatsUdfCommon.ConfigError + ": not a known cluster or server",
                    "check Default.cluster, Clusters and Servers");
                return null;
            }
            return endpoint;
        }

        private static bool IsWriteSuccess(object result)
        {
            var text = result as string;
            return text == NatsUdfCommon.Ok
                || text == NatsUdfCommon.FireAndForget
                || text == NatsUdfCommon.FireAndForgetAll;
        }

        /// <summary>
        /// True when the result is one of the fire-and-forget markers, i.e. the
        /// write was accepted by the client but not confirmed by the server.
        /// <see cref="IsWriteSuccess"/> treats those as success so the cell can
        /// report them; PublishIfChanged must NOT turn that into confirmed dedup
        /// state (DESIGN.md 8.1), so it asks this before remembering a payload.
        /// </summary>
        private static bool IsFireAndForgetResult(object result)
        {
            var text = result as string;
            return text == NatsUdfCommon.FireAndForget
                || text == NatsUdfCommon.FireAndForgetAll;
        }

        private static string RuntimeError(string context, Exception ex)
        {
            return RuntimeError(context, NatsUdfAsync.SafeMessage(ex));
        }

        private static string RuntimeError(string context, string cause)
        {
            if (string.IsNullOrWhiteSpace(cause))
                cause = "unknown error";
            return NatsUdfCommon.VerboseError(context, cause, NatsUdfCommon.ErrorHint(cause));
        }

        /// <summary>
        /// Observes a background publish so a late failure is logged instead of
        /// surfacing as an unobserved task exception. Runs synchronously on the
        /// completing thread and never throws.
        /// </summary>
        private static void Observe(Task task, string context, string endpoint, bool warnOnFailure)
        {
            if (task == null)
                return;
            _ = task.ContinueWith(
                completed =>
                {
                    try
                    {
                        if (!completed.IsFaulted)
                            return;
                        string line = "NatsUdfCore: background NATS write failed: " + context
                            + " endpoint=" + NatsConnectionManager.MaskedEndpoint(endpoint);
                        if (warnOnFailure)
                            logger.Error(completed.Exception, line);
                        else
                            logger.Debug(completed.Exception, line);
                    }
                    catch
                    {
                        // Never let a logging failure cross a completion callback.
                    }
                },
                CancellationToken.None,
                TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
        }

        private static object[,] SingleCell(object value)
        {
            var matrix = new object[1, 1];
            matrix[0, 0] = value;
            return matrix;
        }

        // ==================================================================
        // Bounded LRU for PublishIfChanged
        // ==================================================================

        /// <summary>
        /// Bounded least-recently-used cache remembering the last delivered
        /// payload per <c>(endpoint, subject)</c> without growing without limit.
        /// Entries are evicted from the cold end past the configured capacity
        /// (<see cref="AppConfig.PublishDedupCacheSize"/>). Every operation is
        /// serialized internally; the per-key check -> publish -> store sequence
        /// is still protected by <see cref="NatsUdfCore"/>'s striped lock.
        /// </summary>
        internal sealed class DedupCache
        {
            private sealed class Entry
            {
                public readonly string Key;
                public string Value;

                public Entry(string key, string value)
                {
                    Key = key;
                    Value = value;
                }
            }

            private readonly int _capacity;
            private readonly Dictionary<string, LinkedListNode<Entry>> _entries;
            private readonly LinkedList<Entry> _recent = new LinkedList<Entry>();
            private readonly object _sync = new object();

            public DedupCache(int capacity)
            {
                if (capacity <= 0)
                    throw new ArgumentOutOfRangeException(nameof(capacity), "capacity must be positive");
                _capacity = capacity;
                // The cap can be configured very large; only use a small initial
                // hint so a typo cannot preallocate a huge dictionary.
                _entries = new Dictionary<string, LinkedListNode<Entry>>(Math.Min(capacity, 1024), StringComparer.Ordinal);
            }

            public int Count
            {
                get
                {
                    lock (_sync)
                        return _entries.Count;
                }
            }

            /// <summary>Snapshot of the tracked keys, taken under the internal
            /// lock and returned as a plain array so callers can walk it without
            /// holding that lock (the listener-join purge takes per-key stripes
            /// next, and the publish path takes them in the opposite order).</summary>
            public string[] SnapshotKeys()
            {
                lock (_sync)
                {
                    var keys = new string[_entries.Count];
                    _entries.Keys.CopyTo(keys, 0);
                    return keys;
                }
            }

            /// <summary>Looks up the value and marks the entry as most recently used.</summary>
            public bool TryGet(string key, out string value)
            {
                lock (_sync)
                {
                    if (!_entries.TryGetValue(key, out var node))
                    {
                        value = null;
                        return false;
                    }
                    _recent.Remove(node);
                    _recent.AddFirst(node);
                    value = node.Value.Value;
                    return true;
                }
            }

            /// <summary>Adds or updates the value, evicting the least recently
            /// used entry when the cache is over capacity.</summary>
            public void Set(string key, string value)
            {
                lock (_sync)
                {
                    if (_entries.TryGetValue(key, out var existing))
                    {
                        existing.Value.Value = value;
                        _recent.Remove(existing);
                        _recent.AddFirst(existing);
                        return;
                    }

                    var entry = new Entry(key, value);
                    var node = _recent.AddFirst(entry);
                    _entries.Add(key, node);
                    if (_entries.Count > _capacity)
                    {
                        var oldest = _recent.Last;
                        _recent.RemoveLast();
                        _entries.Remove(oldest.Value.Key);
                    }
                }
            }

            /// <summary>Removes the entry if present; returns whether it existed.</summary>
            public bool Remove(string key)
            {
                lock (_sync)
                {
                    if (!_entries.TryGetValue(key, out var node))
                        return false;
                    _entries.Remove(key);
                    _recent.Remove(node);
                    return true;
                }
            }
        }
    }
}
