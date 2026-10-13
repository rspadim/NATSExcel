using NATS.Client.Core;
using NATS.Client.JetStream;
using NATS.Client.KeyValueStore;
using NATS.Client.ObjectStore;
using NATSExcel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NATSExcel.SmokeTests
{
    /// <summary>
    /// Behavior smoke test for the NATSExcel runtime (PLAN V8.3, no Excel
    /// required). It drives the real managers and UDF entry points against a
    /// live <c>nats-server</c>:
    ///
    /// <list type="bullet">
    /// <item>Core subscription ref-counting, wildcards and queue groups.</item>
    /// <item>Request/reply (answer, no-responders, timeout).</item>
    /// <item>Consecutive-duplicate suppression (ordinal, literal-only).</item>
    /// <item>JetStream publish/ack/replay/purge and the named missing-stream error.</item>
    /// <item>Key/Value CRUD, CAS, tombstone and the named missing-bucket error.</item>
    /// <item>Object Store put/get/digest and the named missing-object error.</item>
    /// <item>The frozen admin-gate text, both with the gate closed and open.</item>
    /// <item>The JetStream-disabled path (skipped, with a reason, when the server
    /// has JetStream on).</item>
    /// </list>
    ///
    /// <para>Usage:
    /// <c>dotnet run --project test\SmokeTests -c Release -- "nats://127.0.0.1:4222"</c>.
    /// The URL is a command-line argument only; no real endpoint is ever committed.
    /// Pass <c>--require-jetstream</c> to fail (instead of skip) when the server
    /// runs without <c>-js</c>; pass <c>--read-only</c> to skip every write/admin
    /// section.</para>
    ///
    /// <para>The production sources are compiled directly into this assembly
    /// (linked <c>Compile</c> items), so the internal test seams
    /// (<c>AppConfig.CurrentOverrideForTests</c>,
    /// <c>NatsUdfAsync.AsyncWritesOverrideForTests</c>,
    /// <c>NatsSubscriptionManager.SkipRepeatedOverrideForTests</c>) are visible and
    /// the code under test is the code that ships.</para>
    /// </summary>
    internal static class Program
    {
        private const string DefaultUrl = "nats://127.0.0.1:4222";

        /// <summary>Frozen gate text (NatsUdfCommon.ErrorAdminWritesDisabled + prefix).</summary>
        private const string AdminGateText = "Error: admin writes are disabled (set \"AllowAdminWrites\": true)";

        /// <summary>Frozen JetStream-disabled text (NatsUdfCommon.ErrorJetStreamDisabled + prefix).</summary>
        private const string JetStreamOffText = "Error: JetStream is not enabled on this server";

        private static int _failures;
        private static int _passes;
        private static int _skips;
        private static readonly object Sync = new object();

        private static string _url;
        private static string _prefix;

        private static async Task<int> Main(string[] args)
        {
            string url = args.FirstOrDefault(a => a.Contains("://")) ?? DefaultUrl;
            bool requireJetStream = args.Any(a => string.Equals(a, "--require-jetstream", StringComparison.OrdinalIgnoreCase));
            bool readOnly = args.Any(a => string.Equals(a, "--read-only", StringComparison.OrdinalIgnoreCase));

            _url = url.Trim();
            _prefix = "smoke." + Guid.NewGuid().ToString("N").Substring(0, 10);

            // Deterministic seams: no Excel (sync dispatch), duplicate suppression
            // pinned on (the dedup checks must exercise the suppression branch),
            // and a config that points the blank-host UDF calls at the CLI URL.
            NatsUdfAsync.AsyncWritesOverrideForTests = false;
            NatsSubscriptionManager.SkipRepeatedOverrideForTests = true;
            AppConfig.CurrentOverrideForTests = NewRoot(url, allowAdminWrites: false, syncWrite: "sync");

            Console.WriteLine("NATSExcel smoke test (no Excel)");
            Console.WriteLine("  url        : " + _url);
            Console.WriteLine("  prefix     : " + _prefix);
            Console.WriteLine("  read-only  : " + readOnly);
            Console.WriteLine();

            NatsConnection raw;
            try
            {
                raw = await ConnectAsync(_url).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Console.WriteLine("  connect failed: " + Describe(ex));
                raw = null;
            }

            if (raw == null)
            {
                Check(false, "a NATS server answers at " + _url
                    + " (start one with test\\TestSupport\\Start-Cluster.ps1, or `nats-server -js`)");
                return Summary();
            }

            try
            {
                var info = raw.ServerInfo;
                bool jetStream = info != null && info.JetStreamAvailable;
                Console.WriteLine("  server     : " + (info?.Name ?? "(unknown)") + " " + (info?.Version ?? "?")
                    + ", cluster=" + (string.IsNullOrEmpty(info?.Cluster) ? "(standalone)" : info.Cluster)
                    + ", jetstream=" + jetStream);
                Console.WriteLine();

                await RunCoreChecksAsync(raw).ConfigureAwait(false);
                await RunRequestReplyChecksAsync(raw).ConfigureAwait(false);
                await RunDedupChecksAsync(raw).ConfigureAwait(false);

                if (readOnly)
                {
                    Skip("JetStream checks", "read-only mode");
                    Skip("Key/Value checks", "read-only mode");
                    Skip("Object Store checks", "read-only mode");
                    Skip("admin gate checks", "read-only mode");
                    Skip("JetStream-disabled path", "read-only mode");
                }
                else
                {
                    await RunGateChecksAsync(raw, jetStream).ConfigureAwait(false);
                    if (jetStream)
                    {
                        await RunJetStreamChecksAsync(raw).ConfigureAwait(false);
                        await RunKvChecksAsync(raw).ConfigureAwait(false);
                        await RunObjectChecksAsync(raw).ConfigureAwait(false);
                        Skip("JetStream-disabled path",
                            "the server has JetStream enabled (INFO jetstream:true)");
                    }
                    else
                    {
                        if (requireJetStream)
                            Check(false, "the server has JetStream enabled (--require-jetstream)");
                        else
                            Skip("JetStream checks", "the server runs without -js");
                        Skip("Key/Value checks", "the server runs without -js");
                        Skip("Object Store checks", "the server runs without -js");
                        await RunJetStreamDisabledChecksAsync(raw).ConfigureAwait(false);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("UNEXPECTED: " + Describe(ex));
                Interlocked.Increment(ref _failures);
            }
            finally
            {
                try { await raw.DisposeAsync().ConfigureAwait(false); } catch { }
                try { NatsRuntime.Shutdown(); } catch { }
                AppConfig.CurrentOverrideForTests = null;
                NatsUdfAsync.AsyncWritesOverrideForTests = null;
                NatsSubscriptionManager.SkipRepeatedOverrideForTests = null;
            }

            return Summary();
        }

        // ==================================================================
        // 1-3. Core subscription ref-counting, wildcards, queue groups
        // ==================================================================

        private static async Task RunCoreChecksAsync(NatsConnection raw)
        {
            var subscriptions = NatsRuntime.Subscriptions;

            // ------------------------------------------------------------
            // 1. Ref-counting: two listeners share ONE wire subscription.
            // ------------------------------------------------------------
            string subject = _prefix + ".refcount";
            var receivedA = new List<string>();
            var receivedB = new List<string>();

            int listenersBefore = subscriptions.ListenerCount;
            int statesBefore = subscriptions.SubscriptionCount;

            var tokenA = subscriptions.Subscribe(_url, subject, null,
                m => { lock (Sync) receivedA.Add(Data(m)); }, "SMOKE");
            var tokenB = subscriptions.Subscribe(_url, subject, null,
                m => { lock (Sync) receivedB.Add(Data(m)); }, "SMOKE");

            Check(await WaitUntilAsync(() => subscriptions.ListenerCount == listenersBefore + 2
                                  && subscriptions.SubscriptionCount == statesBefore + 1, 5000).ConfigureAwait(false),
                "refcount: two listeners on one subject share exactly one physical subscription "
                + "(states +1, listeners +2)");

            int connzSubs;
            bool connzOk = TryConnzSubscriptionCount(_url, subject, out connzSubs);
            if (connzOk)
                Check(connzSubs == 1, "refcount: the server reports exactly one subscription for the subject (connz=" + connzSubs + ")");
            else
                Skip("refcount: server-side subscription count",
                    "no monitoring port (start nats-server with -m for /connz); delivery-based proof used");

            await PublishAsync(raw, subject, "msg1").ConfigureAwait(false);
            Check(await WaitUntilAsync(() => Count(receivedA) == 1 && Count(receivedB) == 1, 5000).ConfigureAwait(false),
                "refcount: both listeners received the first message");

            // Disposing one listener must not tear the wire subscription down.
            tokenA.Dispose();
            Check(await WaitUntilAsync(() => subscriptions.ListenerCount == listenersBefore + 1, 5000).ConfigureAwait(false),
                "refcount: listener A removed");
            await PublishAsync(raw, subject, "msg2").ConfigureAwait(false);
            Check(await WaitUntilAsync(() => Count(receivedB) == 2 && Items(receivedB)[1] == "msg2", 5000).ConfigureAwait(false),
                "refcount: B keeps receiving after A left (wire subscription survived)");
            Check(Count(receivedA) == 1, "refcount: the removed listener received nothing after its dispose");

            // The last listener unsubscribes the wire subscription.
            tokenB.Dispose();
            Check(await WaitUntilAsync(() => subscriptions.SubscriptionCount == statesBefore
                                  && subscriptions.ListenerCount == listenersBefore, 5000).ConfigureAwait(false),
                "refcount: counters return to baseline after the last dispose");

            // No delivery reaches a disposed listener, and (with monitoring) the
            // server-side subscription is gone.
            int afterDispose = Count(receivedB);
            await PublishAsync(raw, subject, "msg3").ConfigureAwait(false);
            await PublishAsync(raw, subject, "msg4").ConfigureAwait(false);
            await Task.Delay(400).ConfigureAwait(false);
            Check(Count(receivedB) == afterDispose && Count(receivedA) == 1,
                "refcount: no delivery to either disposed listener after the last dispose");
            if (connzOk && TryConnzSubscriptionCount(_url, subject, out connzSubs))
                Check(connzSubs == 0, "refcount: the server dropped the subject's subscription (connz=" + connzSubs + ")");

            // Re-subscription must work (no stuck state).
            var receivedC = new List<string>();
            var tokenC = subscriptions.Subscribe(_url, subject, null,
                m => { lock (Sync) receivedC.Add(Data(m)); }, "SMOKE");
            Check(await WaitUntilAsync(() => subscriptions.SubscriptionCount == statesBefore + 1, 5000).ConfigureAwait(false),
                "refcount: re-subscription registers a new state");
            await PublishAsync(raw, subject, "msg5").ConfigureAwait(false);
            Check(await WaitUntilAsync(() => Count(receivedC) == 1 && Items(receivedC)[0] == "msg5", 5000).ConfigureAwait(false),
                "refcount: the new listener receives after re-subscription");

            // A double dispose is a safe no-op.
            tokenC.Dispose();
            int statesAfterFirst = subscriptions.SubscriptionCount;
            int listenersAfterFirst = subscriptions.ListenerCount;
            try
            {
                tokenC.Dispose();
                Check(true, "refcount: the second Dispose does not throw");
            }
            catch (Exception ex)
            {
                Check(false, "refcount: the second Dispose does not throw (" + Describe(ex) + ")");
            }
            Check(subscriptions.SubscriptionCount == statesAfterFirst && subscriptions.ListenerCount == listenersAfterFirst,
                "refcount: the second Dispose changes no counters");

            // ------------------------------------------------------------
            // 2. Wildcards (* one token, > one or more) + a literal subscriber.
            // ------------------------------------------------------------
            string wildStar = _prefix + ".wild.*";
            string wildGt = _prefix + ".wild.>";
            string wildOne = _prefix + ".wild.one";
            string wildTwo = _prefix + ".wild.one.two";

            var star = new List<string>();
            var gt = new List<string>();
            var literal = new List<string>();
            var starToken = subscriptions.Subscribe(_url, wildStar, null, m => { lock (Sync) star.Add(m.Subject); }, "SMOKE");
            var gtToken = subscriptions.Subscribe(_url, wildGt, null, m => { lock (Sync) gt.Add(m.Subject); }, "SMOKE");
            var litToken = subscriptions.Subscribe(_url, wildOne, null, m => { lock (Sync) literal.Add(Data(m)); }, "SMOKE");

            Check(await WaitUntilAsync(() => subscriptions.ListenerCount >= listenersBefore + 3, 5000).ConfigureAwait(false),
                "wildcards: all three listeners registered");

            // The wire SUBSCRIBE is asynchronous (at-most-once Core), so warm each
            // subscription up before asserting a delivery count: publish onto a
            // throwaway subject the wildcards match, wait until each listener has
            // seen it, then clear the collected values.
            await WarmUpAsync(raw, _prefix + ".wild.warm", star, gt).ConfigureAwait(false);
            lock (Sync) { star.Clear(); gt.Clear(); literal.Clear(); }

            await PublishAsync(raw, wildOne, "u1").ConfigureAwait(false);
            await PublishAsync(raw, wildTwo, "u2").ConfigureAwait(false);
            Check(await WaitUntilAsync(() => Count(star) == 1 && Count(gt) == 2, 5000).ConfigureAwait(false),
                "wildcards: `*` matched the one-token subject only and `>` matched both "
                + "(star=" + Count(star) + ", gt=" + Count(gt) + ")");
            Check(Count(literal) == 1 && Items(literal)[0] == "u1",
                "wildcards: the literal subscriber is independent (only its own subject)");
            Check(Items(star).All(s => s == wildOne),
                "wildcards: the `*` subject never leaks a two-token subject");

            starToken.Dispose();
            gtToken.Dispose();
            litToken.Dispose();

            // ------------------------------------------------------------
            // 3. Queue groups: a plain and a queue subscriber on the same subject
            //    are DISTINCT physical subscriptions (different SIDs), so the
            //    manager's per-subscription fan-out delivers to every listener of
            //    each. The wire identity is what this asserts; the load-balancing
            //    guarantee is the server's (one delivery per queue-group message
            //    per group), not the manager's.
            // ------------------------------------------------------------
            string queueSubject = _prefix + ".queue";
            const int queueMessages = 100;
            var q1 = new List<string>();
            var q2 = new List<string>();
            var plain = new List<string>();
            var q1Token = subscriptions.Subscribe(_url, queueSubject, "workers", m => { lock (Sync) q1.Add(Data(m)); }, "SMOKE");
            var q2Token = subscriptions.Subscribe(_url, queueSubject, "workers", m => { lock (Sync) q2.Add(Data(m)); }, "SMOKE");
            var plainToken = subscriptions.Subscribe(_url, queueSubject, null, m => { lock (Sync) plain.Add(Data(m)); }, "SMOKE");

            Check(await WaitUntilAsync(() => subscriptions.ListenerCount >= listenersBefore + 3
                                  && subscriptions.SubscriptionCount >= statesBefore + 2, 5000).ConfigureAwait(false),
                "queue: a plain and a queue subscriber on one subject are two distinct physical subscriptions "
                + "(2 states, listeners=" + subscriptions.ListenerCount + ")");

            // The wire SUBSCRIBE is asynchronous (at-most-once Core), so warm each
            // state before asserting a delivery count. q1 and q2 share ONE wire
            // subscription (same (endpoint, subject, queueGroup) key), so a single
            // queue-group delivery proves that state's SUB is live; the plain
            // subscriber is a second state and is warmed separately.
            Check(await WarmUpQueueAsync(raw, queueSubject, q1, q2).ConfigureAwait(false),
                "queue: the queue-group subscription is live on the wire");
            Check(await WarmUpAsync(raw, queueSubject, plain).ConfigureAwait(false),
                "queue: the plain subscription is live on the wire");
            lock (Sync) { q1.Clear(); q2.Clear(); plain.Clear(); }

            for (int i = 0; i < queueMessages; i++)
                await PublishAsync(raw, queueSubject, "q-" + i.ToString(CultureInfo.InvariantCulture)).ConfigureAwait(false);

            bool allDelivered = await WaitUntilAsync(
                () => Count(plain) == queueMessages && Count(q1) == queueMessages && Count(q2) == queueMessages, 15000).ConfigureAwait(false);
            Check(allDelivered,
                "queue: the two queue listeners share the group's single wire subscription, so each saw all "
                + queueMessages + " messages (plain=" + Count(plain) + ", q1=" + Count(q1) + ", q2=" + Count(q2) + ")");
            Check(Count(plain) == queueMessages,
                "queue: the plain subscriber is a separate subscription and received every message ("
                + Count(plain) + ")");

            q1Token.Dispose();
            q2Token.Dispose();
            plainToken.Dispose();
            Check(await WaitUntilAsync(() => subscriptions.SubscriptionCount == statesBefore
                                  && subscriptions.ListenerCount == listenersBefore, 5000).ConfigureAwait(false),
                "queue: all queue/plain subscriptions released");
        }

        // ==================================================================
        // 4. Request/reply
        // ==================================================================

        private static async Task RunRequestReplyChecksAsync(NatsConnection raw)
        {
            string served = _prefix + ".req";
            string nobody = _prefix + ".req.nobody";
            string silent = _prefix + ".req.silent";

            // A responder that answers with the request payload prefixed.
            using (var responder = new Responder(raw, served, msg => "pong:" + msg))
            {
                Check(await responder.WaitReadyAsync(3000).ConfigureAwait(false), "request: responder subscribed");

                object reply = NatsUdfCore.NatsUDFRequest(served, "ping", 2000, null);
                Check(string.Equals(AsString(reply), "pong:ping", StringComparison.Ordinal),
                    "request: an answered request returns the reply text (got '" + AsString(reply) + "')");
            }

            // No responders: nobody subscribed to the subject.
            object noResponders = NatsUdfCore.NatsUDFRequest(nobody, "ping", 2000, null);
            Check(string.Equals(AsString(noResponders), NatsUdfCommon.NoResponders, StringComparison.Ordinal),
                "request: a request to nobody returns the \"" + NatsUdfCommon.NoResponders + "\" marker "
                + "(got '" + AsString(noResponders) + "')");

            // A responder that never answers must yield the TIMEOUT marker.
            using (var silentResponder = new Responder(raw, silent, reply: null))
            {
                Check(await silentResponder.WaitReadyAsync(3000).ConfigureAwait(false), "request: silent responder subscribed");

                // Diagnostic (not an assertion): the raw client's own failure type for
                // the same scenario, so a failing UDF mapping is attributable.
                Console.WriteLine("  request diagnostic: raw client failure type for a 400 ms timed-out request = "
                    + await RawRequestFailureAsync(raw, silent, 400).ConfigureAwait(false));

                var watch = Stopwatch.StartNew();
                object timeout = NatsUdfCore.NatsUDFRequest(silent, "ping", 400, null);
                watch.Stop();
                Check(string.Equals(AsString(timeout), NatsUdfCommon.Timeout, StringComparison.Ordinal),
                    "request: a slow responder that never answers returns \"" + NatsUdfCommon.Timeout + "\" "
                    + "(got '" + AsString(timeout) + "' in " + watch.ElapsedMilliseconds + " ms)");
                Check(watch.ElapsedMilliseconds < 6000,
                    "request: the timeout is bounded (" + watch.ElapsedMilliseconds + " ms)");
            }

            await Task.CompletedTask;
        }

        // ==================================================================
        // 5. Duplicate suppression
        // ==================================================================

        private static async Task RunDedupChecksAsync(NatsConnection raw)
        {
            var subscriptions = NatsRuntime.Subscriptions;
            Check(NatsSubscriptionManager.SkipRepeatedEnabledForTests,
                "dedup: the duplicate-suppression branch is pinned on by the test seam");

            // Literal subject: identical consecutive payloads delivered once.
            string literal = _prefix + ".dedup";
            var received = new List<string>();
            var token = subscriptions.Subscribe(_url, literal, null, m => { lock (Sync) received.Add(Data(m)); }, "SMOKE");
            Check(await WaitUntilAsync(() => subscriptions.SubscriptionCount > 0, 3000).ConfigureAwait(false),
                "dedup: literal listener registered");

            for (int i = 0; i < 4; i++)
                await PublishAsync(raw, literal, "dup").ConfigureAwait(false);
            await PublishAsync(raw, literal, "dup2").ConfigureAwait(false);
            await WaitUntilAsync(() => Count(received) >= 2, 5000).ConfigureAwait(false);
            int stable = WaitForCountStable(() => Count(received), 2, 400, 5000);
            Check(stable >= 2 && stable <= 3,
                "dedup: four identical payloads collapse to one delivery plus the change "
                + "(received " + stable + ", expected 2 for \"dup\"*4 + \"dup2\")");

            // Ordinal compare: "007" and "7" are both delivered.
            string ordinal = _prefix + ".dedup.ordinal";
            var ord = new List<string>();
            var ordToken = subscriptions.Subscribe(_url, ordinal, null, m => { lock (Sync) ord.Add(Data(m)); }, "SMOKE");
            await Task.Delay(150).ConfigureAwait(false);
            await PublishAsync(raw, ordinal, "007").ConfigureAwait(false);
            await PublishAsync(raw, ordinal, "7").ConfigureAwait(false);
            Check(await WaitUntilAsync(() => Count(ord) == 2, 5000).ConfigureAwait(false),
                "dedup: \"007\" and \"7\" are both delivered (ordinal byte compare, got " + Count(ord) + ")");
            Check(Count(ord) == 2 && Items(ord)[0] == "007" && Items(ord)[1] == "7",
                "dedup: the ordinal pair arrives in order and unchanged");

            // A wildcard subject is never suppressed.
            string wildcard = _prefix + ".dedup.*";
            string wildSubject = _prefix + ".dedup.wild";
            var wild = new List<string>();
            var wildToken = subscriptions.Subscribe(_url, wildcard, null, m => { lock (Sync) wild.Add(Data(m)); }, "SMOKE");
            await Task.Delay(150).ConfigureAwait(false);
            await PublishAsync(raw, wildSubject, "same").ConfigureAwait(false);
            await PublishAsync(raw, wildSubject, "same").ConfigureAwait(false);
            await PublishAsync(raw, wildSubject, "same").ConfigureAwait(false);
            Check(await WaitUntilAsync(() => Count(wild) == 3, 5000).ConfigureAwait(false),
                "dedup: a wildcard subject is never suppressed (received " + Count(wild) + " of 3 identical)");

            token.Dispose();
            ordToken.Dispose();
            wildToken.Dispose();
            await Task.Delay(100).ConfigureAwait(false);
        }

        // ==================================================================
        // 9. The admin gate (exact frozen text), closed then open
        // ==================================================================

        private static async Task RunGateChecksAsync(NatsConnection raw, bool jetStream)
        {
            string stream = StreamName();
            string bucket = KvName();
            string objBucket = ObjName();

            // --- Gate closed (default) ------------------------------------
            AppConfig.CurrentOverrideForTests = NewRoot(_url, allowAdminWrites: false, syncWrite: "sync");

            var closed = new List<KeyValuePair<string, object>>
            {
                new KeyValuePair<string, object>("NatsUDFStreamDelete",
                    NatsUdfJetStream.NatsUDFStreamDelete(stream, null)),
                new KeyValuePair<string, object>("NatsUDFStreamCreate",
                    NatsUdfJetStream.NatsUDFStreamCreate(stream, _prefix + ".gate.>", null, null, null)),
                new KeyValuePair<string, object>("NatsUDFJSPurge",
                    NatsUdfJetStream.NatsUDFJSPurge(stream, null, null)),
                new KeyValuePair<string, object>("NatsUDFJSDeleteMessage",
                    NatsUdfJetStream.NatsUDFJSDeleteMessage(stream, 1, null)),
                new KeyValuePair<string, object>("NatsUDFConsumerDelete",
                    NatsUdfJetStream.NatsUDFConsumerDelete(stream, "c", null)),
                new KeyValuePair<string, object>("NatsUDFKVDelete",
                    NatsUdfKv.NatsUDFKVDelete(bucket, "k", null)),
                new KeyValuePair<string, object>("NatsUDFKVPurge",
                    NatsUdfKv.NatsUDFKVPurge(bucket, "k", null)),
                new KeyValuePair<string, object>("NatsUDFObjectDelete",
                    NatsUdfObjectStore.NatsUDFObjectDelete(objBucket, "o", null)),
            };

            bool allExact = true;
            foreach (var call in closed)
            {
                string text = AsString(call.Value);
                if (!string.Equals(text, AdminGateText, StringComparison.Ordinal))
                {
                    allExact = false;
                    Console.WriteLine("  gate mismatch " + call.Key + " -> '" + text + "'");
                }
            }
            Check(allExact, "gate(closed): every admin/destructive function returns the exact frozen text "
                + "\"" + AdminGateText + "\" (" + closed.Count + " calls)");

            // A read (no admin gate) still works while the gate is closed.
            object readWhileClosed = NatsUdfJetStream.NatsUDFJSStreamInfo(stream, null);
            Check(!string.Equals(AsString(readWhileClosed), AdminGateText, StringComparison.Ordinal),
                "gate(closed): a non-admin read is not blocked by the admin gate");

            // --- Gate open ------------------------------------------------
            AppConfig.CurrentOverrideForTests = NewRoot(_url, allowAdminWrites: true, syncWrite: "sync");

            if (!jetStream)
            {
                // The gate opened, but a server without -js refuses the call with the
                // named JetStream error BEFORE any server mutation - so a "proceeds"
                // assertion is only meaningful on a JetStream server.
                Skip("gate(open): the admin call proceeds",
                    "the server runs without -js, so the JetStream gate refuses it first");
            }
            else
            {
                object created = NatsUdfJetStream.NatsUDFStreamCreate(stream, _prefix + ".gate.>", null, null, null);
                Check(string.Equals(AsString(created), NatsUdfCommon.Ok, StringComparison.Ordinal),
                    "gate(open): NatsUDFStreamCreate proceeds (got '" + AsString(created) + "')");

                object deleted = NatsUdfJetStream.NatsUDFStreamDelete(stream, null);
                Check(string.Equals(AsString(deleted), NatsUdfCommon.Ok, StringComparison.Ordinal),
                    "gate(open): NatsUDFStreamDelete proceeds (got '" + AsString(deleted) + "')");

                // The same call with the gate open reaches the server: a second
                // delete of the now-missing stream is a named error, never the gate.
                object reread = NatsUdfJetStream.NatsUDFStreamDelete(stream, null);
                Check(AsString(reread).IndexOf(AdminGateText, StringComparison.Ordinal) < 0,
                    "gate(open): the admin call is no longer refused by the gate (got '" + AsString(reread) + "')");
            }

            await Task.CompletedTask;
        }

        // ==================================================================
        // 6. JetStream: publish/ack, durable consumer, replay, purge, missing stream
        // ==================================================================

        private static async Task RunJetStreamChecksAsync(NatsConnection raw)
        {
            AppConfig.CurrentOverrideForTests = NewRoot(_url, allowAdminWrites: true, syncWrite: "sync");

            string stream = StreamName();
            string consumer = "smoke_c_" + Guid.NewGuid().ToString("N").Substring(0, 8);
            string filter = _prefix + ".js.>";
            string subject = _prefix + ".js.a";

            try
            {
                object created = NatsUdfJetStream.NatsUDFStreamCreate(stream, filter, null, null, null);
                Check(string.Equals(AsString(created), NatsUdfCommon.Ok, StringComparison.Ordinal),
                    "jetstream: stream created (got '" + AsString(created) + "')");

                // Publish-with-ack returns a sequence.
                object seqResult = NatsUdfJetStream.NatsUDFJSPublish(subject, "js-1", null);
                double seq = AsDouble(seqResult);
                Check(seq >= 1, "jetstream: publish-with-ack returns a sequence (got '" + AsString(seqResult) + "')");

                object secondSeq = NatsUdfJetStream.NatsUDFJSPublish(subject, "js-2", null);
                Check(AsDouble(secondSeq) == seq + 1,
                    "jetstream: the second publish advances the sequence (got '" + AsString(secondSeq) + "')");

                // A durable consumer receives.
                object consumerCreated = NatsUdfJetStream.NatsUDFConsumerCreate(stream, consumer, filter, null, null);
                Check(string.Equals(AsString(consumerCreated), NatsUdfCommon.Ok, StringComparison.Ordinal),
                    "jetstream: a durable consumer was created (got '" + AsString(consumerCreated) + "')");

                object fetch = NatsUdfJetStream.NatsUDFJSFetch(stream, consumer, 10, 2000, null);
                string[][] rows = MatrixRows(fetch);
                Check(rows.Length == 2,
                    "jetstream: the durable consumer pulled the two stored messages (got " + rows.Length + " rows)");
                if (rows.Length == 2)
                {
                    Check(rows[0].Length >= 4 && rows[0][3] == "js-1" && rows[1][3] == "js-2",
                        "jetstream: the fetched payloads match in order (got '"
                        + string.Join("','", rows.Select(r => r.Length >= 4 ? r[3] : "?")) + "')");
                }

                // Replay by sequence matches the published payload.
                object replayed = NatsUdfJetStream.NatsUDFJSGetMessage(stream, (long)seq, null);
                Check(string.Equals(AsString(replayed), "js-1", StringComparison.Ordinal),
                    "jetstream: a message replayed by sequence matches (got '" + AsString(replayed) + "')");

                // The named missing-stream error.
                string missing = StreamName();
                object missingResult = NatsUdfJetStream.NatsUDFJSGetMessage(missing, 1, null);
                string missingText = AsString(missingResult);
                Check(missingText.StartsWith(NatsUdfCommon.ErrorPrefix, StringComparison.Ordinal)
                      && missingText.IndexOf("stream '" + missing + "' not found", StringComparison.Ordinal) >= 0,
                    "jetstream: a missing stream returns the named error (got '" + missingText + "')");

                // Purge empties the stream.
                object purged = NatsUdfJetStream.NatsUDFJSPurge(stream, null, null);
                Check(AsDouble(purged) == 2, "jetstream: purge reports 2 purged messages (got '" + AsString(purged) + "')");

                object info = NatsUdfJetStream.NatsUDFJSStreamInfo(stream, null);
                long messages = JsonLong(info, "messages");
                Check(messages == 0, "jetstream: the stream is empty after the purge (messages=" + messages + ")");
            }
            finally
            {
                try { NatsUdfJetStream.NatsUDFStreamDelete(stream, null); } catch { }
            }

            await Task.CompletedTask;
        }

        // ==================================================================
        // 7. Key/Value: CRUD, create/CAS, tombstone, missing bucket
        // ==================================================================

        private static async Task RunKvChecksAsync(NatsConnection raw)
        {
            AppConfig.CurrentOverrideForTests = NewRoot(_url, allowAdminWrites: true, syncWrite: "sync");

            string bucket = KvName();
            string missingBucket = KvName();
            string key = "alpha";

            INatsJSContext js = NatsRuntime.Connections.GetJetStream(_url);
            var kv = new NatsKVContext(js);
            try
            {
                await kv.CreateStoreAsync(bucket, CancellationToken.None).ConfigureAwait(false);

                // put/get round trip.
                object put = NatsUdfKv.NatsUDFKVPut(bucket, key, "v1", null);
                double rev1 = AsDouble(put);
                Check(rev1 >= 1, "kv: put returns a revision (got '" + AsString(put) + "')");

                object got = NatsUdfKv.NatsUDFKVGet(bucket, key, null);
                Check(string.Equals(AsString(got), "v1", StringComparison.Ordinal),
                    "kv: get returns the stored value (got '" + AsString(got) + "')");

                // A missing key is the documented marker.
                object absent = NatsUdfKv.NatsUDFKVGet(bucket, "does-not-exist-" + Guid.NewGuid().ToString("N"), null);
                Check(string.Equals(AsString(absent), NatsUdfCommon.NoValue, StringComparison.Ordinal),
                    "kv: a missing key returns \"" + NatsUdfCommon.NoValue + "\" (got '" + AsString(absent) + "')");

                // create-if-absent on an existing key fails, never overwrites.
                object createExisting = NatsUdfKv.NatsUDFKVCreate(bucket, key, "vX", null);
                string createText = AsString(createExisting);
                Check(createText.IndexOf("already exists", StringComparison.Ordinal) >= 0,
                    "kv: create on an existing key fails (got '" + createText + "')");
                Check(string.Equals(AsString(NatsUdfKv.NatsUDFKVGet(bucket, key, null)), "v1", StringComparison.Ordinal),
                    "kv: the failed create did not overwrite the value");

                // CAS with the wrong revision fails; with the right one it succeeds.
                // (A KV bucket is one backing stream, so revisions are stream-wide:
                // the successful CAS must compare against THIS key's current
                // revision - rev1, read from the put above - not rev1+1 predicted
                // from a later write.)
                object wrongCas = NatsUdfKv.NatsUDFKVUpdate(bucket, key, "v2", 999, null);
                string wrongText = AsString(wrongCas);
                Check(wrongText.IndexOf("wrong last revision", StringComparison.Ordinal) >= 0,
                    "kv: a CAS with the wrong revision fails (got '" + wrongText + "')");
                Check(string.Equals(AsString(NatsUdfKv.NatsUDFKVGet(bucket, key, null)), "v1", StringComparison.Ordinal),
                    "kv: the failed CAS did not overwrite the value");

                object rightCas = NatsUdfKv.NatsUDFKVUpdate(bucket, key, "v2", (long)rev1, null);
                Check(AsDouble(rightCas) > rev1,
                    "kv: a CAS with the current revision succeeds and advances it "
                    + "(got '" + AsString(rightCas) + "' from " + rev1 + ")");
                Check(string.Equals(AsString(NatsUdfKv.NatsUDFKVGet(bucket, key, null)), "v2", StringComparison.Ordinal),
                    "kv: the successful CAS stored the new value");

                // create-if-absent on a new key succeeds.
                string newKey = "beta";
                object createNew = NatsUdfKv.NatsUDFKVCreate(bucket, newKey, "b1", null);
                Check(AsDouble(createNew) >= 1, "kv: create on an absent key returns a revision (got '" + AsString(createNew) + "')");

                // delete leaves the documented tombstone behavior.
                object deleted = NatsUdfKv.NatsUDFKVDelete(bucket, key, null);
                Check(string.Equals(AsString(deleted), NatsUdfCommon.Ok, StringComparison.Ordinal),
                    "kv: delete returns " + NatsUdfCommon.Ok + " (got '" + AsString(deleted) + "')");
                Check(string.Equals(AsString(NatsUdfKv.NatsUDFKVGet(bucket, key, null)), NatsUdfCommon.NoValue, StringComparison.Ordinal),
                    "kv: a deleted key reads as \"" + NatsUdfCommon.NoValue + "\"");

                object history = NatsUdfKv.NatsUDFKVHistory(bucket, key, null);
                string[][] histRows = MatrixRows(history);
                bool hasDel = histRows.Any(r => r.Any(c => string.Equals(c, "DEL", StringComparison.Ordinal)));
                Check(hasDel,
                    "kv: the revision history keeps the DEL tombstone (rows=" + histRows.Length + ")");

                // A missing bucket is a named error and never an implicit create.
                object missingGet = NatsUdfKv.NatsUDFKVGet(missingBucket, "k", null);
                string missingText = AsString(missingGet);
                Check(missingText.IndexOf("bucket '" + missingBucket + "' not found", StringComparison.Ordinal) >= 0,
                    "kv: a missing bucket is a named error, never an implicit create (got '" + missingText + "')");

                object buckets = NatsUdfKv.NatsUDFKVBuckets(null);
                Check(!MatrixContains(buckets, missingBucket),
                    "kv: the missing bucket was not implicitly created");
            }
            finally
            {
                try { await kv.DeleteStoreAsync(bucket, CancellationToken.None).ConfigureAwait(false); } catch { }
            }
        }

        // ==================================================================
        // 8. Object Store: put/get with digest, missing object
        // ==================================================================

        private static async Task RunObjectChecksAsync(NatsConnection raw)
        {
            AppConfig.CurrentOverrideForTests = NewRoot(_url, allowAdminWrites: true, syncWrite: "sync");

            string bucket = ObjName();
            string name = "greeting.txt";

            INatsJSContext js = NatsRuntime.Connections.GetJetStream(_url);
            var obj = new NatsObjContext(js);
            try
            {
                await obj.CreateObjectStoreAsync(bucket, CancellationToken.None).ConfigureAwait(false);

                object put = NatsUdfObjectStore.NatsUDFObjectPut(bucket, name, "hello object", null, null);
                string digest = AsString(put);
                Check(digest != null && digest.StartsWith("SHA-256=", StringComparison.Ordinal),
                    "object: put returns the server digest (got '" + digest + "')");

                object got = NatsUdfObjectStore.NatsUDFObjectGet(bucket, name, null);
                Check(string.Equals(AsString(got), "hello object", StringComparison.Ordinal),
                    "object: get returns the stored bytes (got '" + AsString(got) + "')");

                object info = NatsUdfObjectStore.NatsUDFObjectInfo(bucket, name, null);
                string infoText = AsString(info);
                Check(infoText != null && digest != null && infoText.IndexOf(digest, StringComparison.Ordinal) >= 0,
                    "object: info reports the same digest (got '" + infoText + "')");

                // A missing object is the documented marker: the named
                // "object '<name>' not found" error (DESIGN section 7.5).
                string missingName = "does-not-exist-" + Guid.NewGuid().ToString("N");
                object missing = NatsUdfObjectStore.NatsUDFObjectGet(bucket, missingName, null);
                string missingText = AsString(missing);
                Check(missingText != null
                      && missingText.IndexOf("object '" + missingName + "' not found", StringComparison.Ordinal) >= 0,
                    "object: a missing object is the documented named error (got '" + missingText + "')");
            }
            finally
            {
                try { obj.DeleteObjectStore(bucket, CancellationToken.None); } catch { }
            }

            await Task.CompletedTask;
        }

        // ==================================================================
        // 10. The JetStream-disabled path
        // ==================================================================

        private static async Task RunJetStreamDisabledChecksAsync(NatsConnection raw)
        {
            AppConfig.CurrentOverrideForTests = NewRoot(_url, allowAdminWrites: true, syncWrite: "sync");

            string subject = _prefix + ".nojs.a";
            string bucket = KvName();

            var watch = Stopwatch.StartNew();
            object js = NatsUdfJetStream.NatsUDFJSPublish(subject, "x", null);
            object kv = NatsUdfKv.NatsUDFKVGet(bucket, "k", null);
            watch.Stop();

            Check(string.Equals(AsString(js), JetStreamOffText, StringComparison.Ordinal),
                "js-disabled: NatsUDFJSPublish returns the exact \"" + JetStreamOffText + "\" (got '" + AsString(js) + "')");
            Check(string.Equals(AsString(kv), JetStreamOffText, StringComparison.Ordinal),
                "js-disabled: NatsUDFKVGet returns the exact \"" + JetStreamOffText + "\" (got '" + AsString(kv) + "')");
            Check(watch.ElapsedMilliseconds < 5000,
                "js-disabled: neither call hangs (" + watch.ElapsedMilliseconds + " ms)");

            await Task.CompletedTask;
        }

        // ==================================================================
        // Raw client helpers
        // ==================================================================

        /// <summary>Opens a raw NATS connection (bounded, never hangs).</summary>
        private static async Task<NatsConnection> ConnectAsync(string url)
        {
            var opts = NatsOpts.Default with
            {
                Url = url,
                Name = "NATSExcel SmokeTests",
                ConnectTimeout = TimeSpan.FromSeconds(5),
            };
            var connection = new NatsConnection(opts);
            try
            {
                await WithTimeoutAsync(connection.ConnectAsync().AsTask(), 6000).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Console.WriteLine("  connect failed: " + Describe(ex));
                try { await connection.DisposeAsync().ConfigureAwait(false); } catch { }
                throw;
            }
            return connection;
        }

        private static Task PublishAsync(NatsConnection connection, string subject, string payload)
        {
            return connection.PublishAsync(subject, payload).AsTask();
        }

        /// <summary>
        /// Warms a subscription up so its asynchronous wire SUBSCRIBE is established
        /// before a delivery count is asserted: publishes a throwaway payload on
        /// <paramref name="subject"/> until every observed list grew (Core is
        /// at-most-once, so a publish before the SUB is silently dropped). The
        /// caller clears the lists afterwards.
        /// </summary>
        private static async Task<bool> WarmUpAsync(NatsConnection connection, string subject, params List<string>[] observed)
        {
            int[] baseline = observed.Select(Count).ToArray();
            for (int attempt = 0; attempt < 50; attempt++)
            {
                // A UNIQUE payload per attempt: the manager suppresses identical
                // consecutive payloads on a literal subject, so a constant warm
                // payload would be dropped after the first one.
                await PublishAsync(connection, subject, "warm-" + Guid.NewGuid().ToString("N")).ConfigureAwait(false);
                if (await WaitUntilAsync(
                        () => observed.Select((list, i) => Count(list) > baseline[i]).All(grew => grew), 500).ConfigureAwait(false))
                    return true;
                await Task.Delay(50).ConfigureAwait(false);
            }
            return false;
        }

        /// <summary>
        /// Warms the shared queue-group subscription: a single delivery proves the
        /// one wire subscription both listeners ride on is live (which listener got
        /// it is the server's choice).
        /// </summary>
        private static async Task<bool> WarmUpQueueAsync(NatsConnection connection, string subject, List<string> q1, List<string> q2)
        {
            int baseline = Count(q1) + Count(q2);
            for (int attempt = 0; attempt < 50; attempt++)
            {
                await PublishAsync(connection, subject, "warm-" + Guid.NewGuid().ToString("N")).ConfigureAwait(false);
                if (await WaitUntilAsync(() => Count(q1) + Count(q2) > baseline, 500).ConfigureAwait(false))
                    return true;
                await Task.Delay(50).ConfigureAwait(false);
            }
            return false;
        }

        /// <summary>
        /// Diagnostic: issues the same bounded request the UDF issues, with the raw
        /// client, and reports the exception's type and message. Used to attribute
        /// a failing <c>TIMEOUT</c> mapping to the exact client failure.
        /// </summary>
        private static async Task<string> RawRequestFailureAsync(NatsConnection connection, string subject, int timeoutMs)
        {
            try
            {
                var opts = new NatsSubOpts { ThrowIfNoResponders = true, Timeout = TimeSpan.FromMilliseconds(timeoutMs) };
                await WithTimeoutAsync(
                    connection.RequestAsync<string, byte[]>(subject, "ping", null, null, null, null, opts, CancellationToken.None).AsTask(),
                    6000).ConfigureAwait(false);
                return "(no failure: a reply arrived)";
            }
            catch (Exception ex)
            {
                Exception inner = ex;
                while (inner != null)
                {
                    if (inner.GetType().Name.IndexOf("Timeout", StringComparison.OrdinalIgnoreCase) >= 0
                        || inner.GetType().Name.IndexOf("NoReply", StringComparison.OrdinalIgnoreCase) >= 0)
                        return inner.GetType().FullName + ": " + inner.Message;
                    inner = inner.InnerException;
                }
                return Describe(ex);
            }
        }

        /// <summary>One received Core message: the lenient UTF-8 payload text.</summary>
        private static string Data(NatsMessage message)
        {
            return message?.Data;
        }

        /// <summary>
        /// A background Core subscriber that collects decoded payloads. It is used
        /// as a request/reply responder: <paramref name="reply"/> null means
        /// "receive but never answer" (the TIMEOUT case).
        /// </summary>
        private sealed class Responder : IDisposable
        {
            private readonly CancellationTokenSource _cts = new CancellationTokenSource();
            private readonly TaskCompletionSource<bool> _ready =
                new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly List<string> _seen = new List<string>();

            public Responder(NatsConnection connection, string subject, Func<string, string> reply)
            {
                _ = Task.Run(async () =>
                {
                    try
                    {
                        INatsSub<byte[]> sub = await connection
                            .SubscribeCoreAsync<byte[]>(subject, null, null, null, _cts.Token)
                            .ConfigureAwait(false);
                        _ready.TrySetResult(true);
                        while (!_cts.IsCancellationRequested)
                        {
                            if (!await sub.Msgs.WaitToReadAsync(_cts.Token).ConfigureAwait(false))
                                return;
                            NatsMsg<byte[]> msg;
                            while (sub.Msgs.TryRead(out msg))
                            {
                                string text = Encoding.UTF8.GetString(msg.Data ?? Array.Empty<byte>());
                                lock (_seen) _seen.Add(text);
                                if (reply != null && !string.IsNullOrEmpty(msg.ReplyTo))
                                    await connection.PublishAsync(msg.ReplyTo, reply(text)).ConfigureAwait(false);
                            }
                        }
                    }
                    catch (OperationCanceledException) { }
                    catch (Exception ex) { Console.WriteLine("  responder error: " + Describe(ex)); }
                });
            }

            /// <summary>Completes once the wire SUB is established (bounded).</summary>
            public async Task<bool> WaitReadyAsync(int timeoutMs)
            {
                Task completed = await Task.WhenAny(_ready.Task, Task.Delay(timeoutMs)).ConfigureAwait(false);
                return completed == _ready.Task;
            }

            public void Dispose()
            {
                // Cancel only: the loop exits on the token, and the client sends UNSUB.
                try { _cts.Cancel(); } catch { }
            }
        }

        /// <summary>
        /// Reads the server's own subscription count for a subject from the
        /// monitoring <c>/connz?subs=1</c> endpoint, when one is enabled. Returns
        /// false (and the caller prints a SKIP) when no monitoring port answers or
        /// the response carries no subscription list - the cluster harness and CI
        /// both start a server without <c>-m</c>.
        /// </summary>
        private static bool TryConnzSubscriptionCount(string url, string subject, out int count)
        {
            count = -1;
            string host = "127.0.0.1";
            int port = 4222;
            try
            {
                var parsed = new Uri(url.Split(',')[0].Trim());
                if (!string.IsNullOrEmpty(parsed.Host))
                    host = parsed.Host;
                if (parsed.Port > 0)
                    port = parsed.Port;
            }
            catch
            {
                return false;
            }

            foreach (int monitorPort in new[] { port + 4000, 8222 })
            {
                try
                {
                    using (var client = new WebClient())
                    {
                        client.Headers[HttpRequestHeader.Accept] = "application/json";
                        string body = client.DownloadString(
                            "http://" + host + ":" + monitorPort.ToString(CultureInfo.InvariantCulture) + "/connz?subs=1");
                        JObject root = JObject.Parse(body);
                        int total = 0;
                        bool sawList = false;
                        foreach (JToken connection in root["connections"] ?? new JArray())
                        {
                            JObject subs = connection["subscriptions_list"] as JObject;
                            if (subs == null)
                                continue;
                            sawList = true;
                            foreach (var property in subs.Properties())
                            {
                                if (string.Equals((string)property.Value, subject, StringComparison.Ordinal))
                                    total++;
                            }
                        }
                        if (!sawList)
                            return false; // the server did not include per-subject data
                        count = total;
                        return true;
                    }
                }
                catch
                {
                    // No monitoring port (or not on this candidate): try the next.
                }
            }
            return false;
        }

        // ==================================================================
        // Assertions / primitives
        // ==================================================================

        private static void Check(bool condition, string label)
        {
            if (condition)
                Interlocked.Increment(ref _passes);
            else
                Interlocked.Increment(ref _failures);
            Console.WriteLine((condition ? "PASS " : "FAIL ") + label);
        }

        private static void Skip(string label, string reason)
        {
            Interlocked.Increment(ref _skips);
            Console.WriteLine("SKIP " + label + " (" + reason + ")");
        }

        private static int Summary()
        {
            Console.WriteLine();
            Console.WriteLine("smoke summary: " + _passes + " passed, " + _failures + " failed, " + _skips + " skipped");
            Console.WriteLine(_failures == 0 ? "ALL PASS" : _failures + " FAILURE(S)");
            return _failures == 0 ? 0 : 1;
        }

        private static async Task<bool> WaitUntilAsync(Func<bool> condition, int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (condition()) return true;
                await Task.Delay(25).ConfigureAwait(false);
            }
            return condition();
        }

        /// <summary>
        /// Waits until <paramref name="count"/> stays equal to <paramref name="expected"/>
        /// for a settle window so a value still arriving cannot pass a fixed sleep;
        /// returns the last observed count.
        /// </summary>
        private static int WaitForCountStable(Func<int> count, int expected, int settleMs, int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            long stableSince = -1;
            int last = -1;
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                last = count();
                if (last == expected)
                {
                    if (stableSince < 0) stableSince = sw.ElapsedMilliseconds;
                    if (sw.ElapsedMilliseconds - stableSince >= settleMs)
                        return last;
                }
                else
                {
                    stableSince = -1;
                }
                Thread.Sleep(20);
            }
            return last;
        }

        /// <summary>
        /// Bounded await for a NATS task: throws <see cref="TimeoutException"/> when
        /// the task does not complete in time so a dead endpoint fails fast instead
        /// of hanging. The smoke harness is a console app with no synchronization
        /// context, so waiting on a client task cannot deadlock (the same
        /// justification the production sources use for their thread-pool
        /// continuations).
        /// </summary>
#pragma warning disable VSTHRD003 // No synchronization context in a console harness; see the summary above.
        private static async Task WithTimeoutAsync(Task task, int timeoutMs)
        {
            Task completed = await Task.WhenAny(task, Task.Delay(timeoutMs)).ConfigureAwait(false);
            if (completed != task)
                throw new TimeoutException("the NATS operation did not complete within " + timeoutMs + "ms");
            await task.ConfigureAwait(false);
        }

        private static async Task<T> WithTimeoutAsync<T>(Task<T> task, int timeoutMs)
        {
            Task completed = await Task.WhenAny(task, Task.Delay(timeoutMs)).ConfigureAwait(false);
            if (completed != task)
                throw new TimeoutException("the NATS operation did not complete within " + timeoutMs + "ms");
            return await task.ConfigureAwait(false);
        }
#pragma warning restore VSTHRD003

        private static int Count(List<string> list)
        {
            lock (Sync) return list.Count;
        }

        private static string[] Items(List<string> list)
        {
            lock (Sync) return list.ToArray();
        }

        private static string AsString(object value)
        {
            if (value == null)
                return null;
            if (value is string text)
                return text;
            if (value is double d)
                return d.ToString("G17", CultureInfo.InvariantCulture);
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static double AsDouble(object value)
        {
            if (value is double d)
                return d;
            if (value is string text && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
                return parsed;
            return double.NaN;
        }

        private static long JsonLong(object value, string field)
        {
            var text = value as string;
            if (string.IsNullOrEmpty(text))
                return -1;
            try
            {
                return JObject.Parse(text).Value<long?>(field) ?? -1;
            }
            catch
            {
                return -1;
            }
        }

        private static string[][] MatrixRows(object matrix)
        {
            if (!(matrix is object[,] grid))
                return new string[0][];

            int rows = grid.GetLength(0);
            int cols = grid.GetLength(1);
            var result = new string[rows][];
            for (int r = 0; r < rows; r++)
            {
                var row = new string[cols];
                for (int c = 0; c < cols; c++)
                    row[c] = AsString(grid[r, c]) ?? string.Empty;
                result[r] = row;
            }
            return result;
        }

        private static bool MatrixContains(object matrix, string needle)
        {
            foreach (string[] row in MatrixRows(matrix))
                foreach (string cell in row)
                    if (string.Equals(cell, needle, StringComparison.Ordinal))
                        return true;
            return false;
        }

        private static string Describe(Exception ex)
        {
            if (ex == null)
                return "no exception";
            var parts = new List<string>();
            for (var e = ex; e != null; e = e.InnerException)
                parts.Add(e.GetType().Name + ": " + e.Message);
            return string.Join(" | ", parts);
        }

        private static string StreamName()
        {
            return "smoke_js_" + Guid.NewGuid().ToString("N").Substring(0, 10);
        }

        private static string KvName()
        {
            return "smoke_kv_" + Guid.NewGuid().ToString("N").Substring(0, 10);
        }

        private static string ObjName()
        {
            return "smoke_obj_" + Guid.NewGuid().ToString("N").Substring(0, 10);
        }

        /// <summary>
        /// The test config root: points the blank-host UDF calls at the CLI URL,
        /// pins synchronous dispatch (no Excel) and the requested admin gate.
        /// </summary>
        private static AppConfig.Root NewRoot(string url, bool allowAdminWrites, string syncWrite)
        {
            return AppConfig.Sanitize(new AppConfig.Root
            {
                Default = url,
                AllowAdminWrites = allowAdminWrites,
                SyncWrite = syncWrite,
                AsyncWrites = false,
                JetStreamEnabled = true,
                TimeoutMs = 1000,
                JetStreamTimeoutMs = 8000,
            });
        }
    }
}
