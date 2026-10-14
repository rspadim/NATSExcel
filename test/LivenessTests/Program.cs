using ExcelDna.Integration;
using NATS.Client.Core;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NATSExcel.LivenessTests
{
    /// <summary>
    /// Adversarial liveness (stall) suite for the NATSExcel production sources
    /// (PLAN V8.5), ported from the sibling RedisExcel liveness layer.
    ///
    /// It drives the REAL <see cref="NatsRuntime"/> /
    /// <see cref="NatsConnectionManager"/> / <see cref="NatsSubscriptionManager"/>
    /// against a live NATS server (default <c>nats://127.0.0.1:4222</c>) while
    /// attacking liveness, and it FAILS the run (exit 1) when a bound is crossed:
    ///
    ///   W)  warmup/connect: subscribe, publish, confirm one delivery, settle.
    ///   S)  runtime startup/shutdown/refresh transition (a throwaway manager pair
    ///       plus the process-wide refresh); the shared runtime must survive.
    ///   A0) subscribe/dispose churn on a few hot keys from several threads;
    ///       no listener leak and delivery still flows.
    ///   A)  ThreadPool starvation + heavy subscribe/dispose/publish churn +
    ///       an <see cref="NatsUdfAsync"/> backlog that must drain after release.
    ///   B)  lock contention on every public entry point (GetConnection,
    ///       subscribe/dispose, the counters, the bounded waits, SubjectLatest/
    ///       unsubscribe, the write UDFs) from many threads; no deadlock and the
    ///       counters stay consistent/return to baseline.
    ///   C)  fault injection: stop and restart the local nats-server process
    ///       (found by the listening port; auto-skipped for a non-local host)
    ///       plus a subscribe/dispose burst during the outage; delivery must
    ///       resume within the bound and no exception may escape a callback.
    ///   D)  async queue bursts (FIFO, drain bound) and an IExcelObservable
    ///       burst under starvation; every queued item must complete.
    ///   F)  reconnect storm: a burst of server kill/restart flapping with a
    ///       subscription live; delivery must resume within the bound each time,
    ///       no duplicate listener is left and no callback exception escapes.
    ///   G)  requests during an outage: every request returns a bounded error
    ///       (No Responders / TIMEOUT / an "Error:") while down and succeeds
    ///       again within the resume bound after the restart.
    ///   H)  JetStream during an outage: a stream and durable consumer must
    ///       survive a server restart (a durable store_dir), then a new publish
    ///       and a fetch must both work after recovery.
    ///   I)  async backlog with a flapping server: a burst enqueued while the
    ///       server flaps must drain within the bound with every item accounted
    ///       for (completed or errored), no silently lost item.
    ///   J)  starvation + fault: the ThreadPool starvation phase combined with a
    ///       server restart; the watchdog must never trip and the queue must
    ///       still drain.
    ///   K)  graceful shutdown under load: NatsRuntime.Shutdown() while traffic
    ///       is in flight returns within the bound, the tombstone refuses to
    ///       resurrect the managers, and ResetAfterAddInReload plus a fresh use
    ///       proves recovery.
    ///   E)  resource stability: listener/subscription/connection counts return
    ///       to their warmup baseline (no leak).
    ///
    /// The heartbeat/watchdog pair is the heart of the layer: a dedicated
    /// process thread touches a timestamp every 20 ms and a dedicated watchdog
    /// thread exits non-zero the moment that timestamp goes stale past the stall
    /// bound, so a plain hang FAILS instead of hanging forever.
    ///
    /// Bounds: stall 5 s, delivery resume 30 s, async backlog drain 20 s.
    ///
    /// Usage:
    ///   dotnet run --project test\LivenessTests -c Release -- "nats://127.0.0.1:4222"
    ///   dotnet run --project test\LivenessTests -c Release -- "nats://127.0.0.1:4222" --quick --skip-restart
    ///
    /// Only run the destructive phase C (the server stop/restart) against a
    /// disposable/local server. A real server is passed as an argument only and
    /// is never committed (AGENTS golden rule 9).
    /// </summary>
    internal static class Program
    {
        // ------------------------------------------------------------------
        // Bounds (PLAN V8.5). A crossing fails the run.
        // ------------------------------------------------------------------

        /// <summary>No process stall may exceed this (the watchdog enforces it).</summary>
        private const double StuckBoundMs = 5000;

        /// <summary>Delivery must resume within this after a fault.</summary>
        private const double ResumeBoundMs = 30000;

        /// <summary>A released async backlog must drain within this.</summary>
        private const double QueueDrainBoundMs = 20000;

        /// <summary>Graceful shutdown (<see cref="NatsRuntime.Shutdown"/>) must
        /// return within this. It is a hang detector, not a tight timing bound:
        /// the connection manager's own total teardown budget is 5 s, so 10 s
        /// still fails a shutdown that fell back to an unbounded wait.</summary>
        private const double ShutdownBoundMs = 10000;

        // ------------------------------------------------------------------
        // Command line / environment
        // ------------------------------------------------------------------

        private static string Host = "nats://127.0.0.1:4222";
        private static bool Quick;
        private static bool SkipRestart;
        private static bool SelfTestWatchdog;

        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        private static readonly Stopwatch Clock = Stopwatch.StartNew();

        // The real managers (the code under test).
        private static NatsConnectionManager C;
        private static NatsSubscriptionManager S;
        private static DeliveryProbe Delivery;
        private static ProcessHeartbeat Proc;
        private static StallWatchdog Watchdog;

        // Bookkeeping for the resource-stability phase.
        private static int BaselineConnections = -1;
        private static readonly List<string> TrackedSubjects = new List<string>();

        // Cached server identity for the stop/restart phases. Resolving it costs a
        // PowerShell + CIM round trip (~3 s), so it is resolved once (phase C) and
        // reused; the pid is refreshed after every successful restart.
        private static string ServerExePath;
        private static string ServerCommandLine;
        private static int ServerPid;

        // Failure collection and worst-case trackers. A failure never aborts the
        // run: it is collected and exit 1 is returned at the end; the watchdog can
        // also exit(1) directly on a stalled process.
        private static readonly ConcurrentQueue<string> Failures = new ConcurrentQueue<string>();
        private static readonly ConcurrentDictionary<string, Latency> Lats =
            new ConcurrentDictionary<string, Latency>();
        private static readonly object WorstGate = new object();
        private static readonly Dictionary<string, double> Worst = new Dictionary<string, double>();

        private static long JoinedEvents;
        private static long ExitCodeByWatchdog;

        // First unexpected UDF/worker error seen in phase B (set once).
        private static string FirstBError;

        /// <summary>Records the first phase-B error exactly once (many worker
        /// threads race here).</summary>
        private static void NoteBError(string text)
        {
            if (text == null) return;
            Interlocked.CompareExchange(ref FirstBError, text, null);
        }

        private static int Main(string[] args)
        {
            try
            {
                ParseArgs(args);
            }
            catch (ShowUsageException)
            {
                Console.WriteLine(Usage);
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[FATAL] " + ex.Message);
                Console.WriteLine(Usage);
                return 1;
            }

            Console.WriteLine("[LIVENESS] url=" + Host
                + " mode=" + (Quick ? "quick" : "full")
                + " skipRestart=" + SkipRestart
                + " stuckBoundMs=" + Fmt(StuckBoundMs)
                + " resumeBoundMs=" + Fmt(ResumeBoundMs)
                + " drainBoundMs=" + Fmt(QueueDrainBoundMs));
            Console.WriteLine("[LIVENESS] antivirus=" + DescribeAntivirus());
            Console.WriteLine("[LIVENESS] destructivePhase=" + (IsLocalEndpoint() ? "enabled (local host)" : "SKIPPED (non-local host)")
                + (SkipRestart ? " skipRestart=yes" : ""));

            try
            {
                if (SelfTestWatchdog)
                    RunWatchdogStallSelfTest();
                else
                    RunAll();
            }
            catch (Exception ex)
            {
                Console.WriteLine("[FATAL] " + ex);
                Failures.Enqueue("fatal: " + ex.Message);
            }
            finally
            {
                try { if (Delivery != null) Delivery.Dispose(); } catch { }
                try { if (Watchdog != null) Watchdog.Stop(); } catch { }
                try { if (Proc != null) Proc.Stop(); } catch { }
                try { if (S != null) S.Dispose(); } catch { }
                try { if (C != null) C.Shutdown(); } catch { }
                try { NatsRuntime.Shutdown(); } catch { }
            }

            Console.WriteLine();
            Console.WriteLine("[WORST] procHeartbeatMs=" + Fmt(GetWorst("procHeartbeatMs")));
            Console.WriteLine("[WORST] watchdogMs=" + Fmt(GetWorst("watchdogMs")));
            Console.WriteLine("[WORST] deliveryGapMs=" + Fmt(GetWorst("deliveryGapMs")));
            Console.WriteLine("[WORST] deliveryResumeMs=" + Fmt(GetWorst("deliveryResumeMs")));
            Console.WriteLine("[WORST] queueDrainMs=" + Fmt(GetWorst("queueDrainMs")));
            Console.WriteLine("[WORST] shutdownMs=" + Fmt(GetWorst("shutdownMs")));
            Console.WriteLine("[WORST] poolHeartbeatGapMs=" + Fmt(GetWorst("poolHeartbeatGapMs")));
            Console.WriteLine("[SUMMARY] listenerJoinedEvents=" + Interlocked.Read(ref JoinedEvents));
            Console.WriteLine("[SUMMARY] failed=" + Failures.Count + " elapsedSec=" + Clock.Elapsed.TotalSeconds.ToString("0.0", Inv));
            foreach (var f in Failures) Console.WriteLine("[FAILED] " + f);
            if (Interlocked.Read(ref ExitCodeByWatchdog) != 0 && Failures.Count == 0)
                Failures.Enqueue("watchdog stopped the run (process stall past " + Fmt(StuckBoundMs) + "ms)");
            return Failures.Count == 0 ? 0 : 1;
        }

        private const string Usage =
            "Usage: dotnet run --project test\\LivenessTests -c Release -- [url] [options]\n" +
            "  url             nats:// URL of the server under test\n" +
            "                  (default: nats://127.0.0.1:4222)\n" +
            "  --quick         shortened attack windows for a fast sanity run\n" +
            "  --skip-restart  skip the destructive stop/restart sub-phase of phase C\n" +
            "                  (phase C still runs its subscribe/dispose outage burst)";

        /// <summary>Signals an explicit --help (exit 0, not an error).</summary>
        private sealed class ShowUsageException : Exception { }

        private static void ParseArgs(string[] args)
        {
            var positional = new List<string>();
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                if (string.Equals(a, "--quick", StringComparison.OrdinalIgnoreCase))
                    Quick = true;
                else if (string.Equals(a, "--skip-restart", StringComparison.OrdinalIgnoreCase))
                    SkipRestart = true;
                else if (string.Equals(a, "--selftest-watchdog-stall", StringComparison.OrdinalIgnoreCase))
                    SelfTestWatchdog = true;
                else if (string.Equals(a, "--help", StringComparison.OrdinalIgnoreCase)
                         || string.Equals(a, "-h", StringComparison.OrdinalIgnoreCase))
                    throw new ShowUsageException();
                else if (a.StartsWith("--", StringComparison.Ordinal))
                    throw new ArgumentException("unknown option '" + a + "'");
                else
                    positional.Add(a);
            }
            if (positional.Count > 0) Host = positional[0].Trim();
            if (positional.Count > 1)
                throw new ArgumentException("unexpected extra argument '" + positional[1] + "'");
            if (string.IsNullOrWhiteSpace(Host))
                throw new ArgumentException("the url must not be blank");
        }

        // ------------------------------------------------------------------
        // Run
        // ------------------------------------------------------------------

        private static void RunAll()
        {
            StartProbes();
            SelfCheckWatchdog();
            Warmup();
            PhaseStartupTeardown();
            PhaseA0();
            PhaseA();
            PhaseB();
            PhaseC();
            PhaseD();
            PhaseF();
            PhaseG();
            PhaseH();
            PhaseI();
            PhaseJ();
            PhaseK();
            PhaseResourceStability();
            Console.WriteLine();
            Console.WriteLine("[PHASE] id=done desc=all phases complete t=" + Clock.Elapsed.TotalSeconds.ToString("0.0", Inv));
        }

        /// <summary>
        /// Creates the real managers, the probes and the dedicated heartbeat /
        /// watchdog threads.
        /// </summary>
        private static void StartProbes()
        {
            // Hermetic, deterministic seams (test-only overrides of the linked
            // production sources):
            //  - pin the config so the Default endpoint resolves to the URL under
            //    test and the dedup/conflation knobs are fixed;
            //  - turn async UDF dispatch OFF so a non-Excel caller runs the write
            //    synchronously (the Observe path needs a worksheet caller); the
            //    real per-endpoint queue is still driven directly in phase D;
            //  - disable consecutive-payload dedup so a repeated probe payload is
            //    never suppressed and the delivery assertions are unambiguous.
            AppConfig.CurrentOverrideForTests = BuildLivenessConfig(Host);
            NatsUdfAsync.AsyncWritesOverrideForTests = false;
            NatsSubscriptionManager.SkipRepeatedOverrideForTests = false;

            C = new NatsConnectionManager();
            S = new NatsSubscriptionManager(C);
            C.SetEvictionProtection(S.HasActiveSubscribers);

            NatsSubscriptionManager.ListenerJoined += (h, s, q) => Interlocked.Increment(ref JoinedEvents);

            // Materialize the process-wide runtime once, so the run really
            // exercises the shared instance (AGENTS golden rule 4). The
            // application-function path reads it directly.
            var sharedConnections = NatsRuntime.Connections;
            var sharedSubscriptions = NatsRuntime.Subscriptions;
            if (sharedConnections == null || sharedSubscriptions == null)
                throw new InvalidOperationException("NatsRuntime did not initialize");

            Delivery = new DeliveryProbe();
            Proc = new ProcessHeartbeat("proc-heartbeat", 20);
            Proc.Start();
            Watchdog = new StallWatchdog(Proc, StuckBoundMs);
            Watchdog.Start();
        }

        /// <summary>
        /// Self-check for the watchdog's stall decision: the layer must fail a
        /// stale heartbeat and pass a fresh one. Exercised once at startup so the
        /// "a plain hang fails the run" contract cannot silently rot into an
        /// always-pass (which would defeat the whole layer).
        /// </summary>
        private static void SelfCheckWatchdog()
        {
            bool staleFails = StallWatchdog.IsStalled(StuckBoundMs + 1, StuckBoundMs);
            bool freshPasses = !StallWatchdog.IsStalled(20, StuckBoundMs);
            bool boundaryPasses = !StallWatchdog.IsStalled(StuckBoundMs, StuckBoundMs);
            Check("selfcheck.watchdog-stall-decision", staleFails && freshPasses && boundaryPasses,
                "stale->fail=" + staleFails + " fresh->pass=" + freshPasses + " boundary->pass=" + boundaryPasses);
            Check("selfcheck.watchdog-runtime-armed", ReferenceEquals(Watchdog.Heartbeat, Proc),
                "the watchdog is armed against the live process heartbeat");
        }

        /// <summary>
        /// Standalone end-to-end proof of the watchdog (--selftest-watchdog-stall):
        /// freeze the process heartbeat and assert the watchdog terminates the
        /// process non-zero instead of letting it hang. Used to verify the harness
        /// itself; never part of a normal run.
        /// </summary>
        private static void RunWatchdogStallSelfTest()
        {
            Console.WriteLine("[SELFTEST] probing the watchdog end-to-end (this process must exit 1)");
            Proc = new ProcessHeartbeat("proc-heartbeat", 20);
            Proc.Start();
            Thread.Sleep(100); // let it tick so Ticks > 0
            Proc.Stop();       // now freeze it: the timestamp never advances again
            Watchdog = new StallWatchdog(Proc, StuckBoundMs);
            Watchdog.Start();

            long deadline = Stopwatch.GetTimestamp() + MsToTicks(StuckBoundMs * 3);
            while (Stopwatch.GetTimestamp() < deadline)
                Thread.Sleep(25);

            Failures.Enqueue("selftest: the watchdog did NOT fire within " + Fmt(StuckBoundMs * 3)
                + "ms; the stall exit path is broken");
        }

        /// <summary>Builds the hermetic config: Default resolves to the URL under
        /// test, JetStream on, fire-and-forget writes, dedup off (the manager
        /// override pins that anyway).</summary>
        private static AppConfig.Root BuildLivenessConfig(string url)
        {
            var cfg = new AppConfig.Root
            {
                Default = "local",
                TimeoutMs = 1000,
                JetStreamEnabled = true,
                PublishDedupCacheSize = 10000,
                SkipRepeatedMessages = false,
                CoalesceRealtimeUpdates = true,
            };
            cfg.Servers["local"] = url;
            var cluster = new AppConfig.Cluster { Label = "local" };
            cluster.Urls.Add(url);
            cfg.Clusters["local"] = cluster;
            return AppConfig.Sanitize(cfg);
        }

        // ------------------------------------------------------------ warmup

        private static void Warmup()
        {
            PhaseBanner("W", "warmup/connect");

            var connection = C.GetConnection(Host);
            Observe(connection.ConnectAsync().AsTask());

            Delivery.Start(C, S, Host, "test.natsexcel.live");

            double first = Delivery.ResumeAfter(0, ResumeBoundMs);
            Check("W.first-delivery", first >= 0, "first live delivery ms=" + Fmt(first));
            Metric("W", "firstDeliveryMs", first);
            Metric("W", "subscriptionCount", S.SubscriptionCount);
            Metric("W", "listenerCount", S.ListenerCount);
            Metric("W", "liveConnections", C.LiveConnectionCount());
            Metric("W", "serverVersion", Shorten(C.GetServerInfo(Host).Version ?? "-", 40));

            BaselineConnections = C.LiveConnectionCount();
            NoteWorst("deliveryResumeMs", first);
            DumpLatencies("W");
        }

        // ------------------------------ runtime startup/shutdown/refresh (F2.16)

        /// <summary>
        /// Pins the runtime create/shutdown/refresh transition (PLAN F2.16): a
        /// fresh manager/subscription pair works end-to-end, then is shut down
        /// cleanly, while the process-wide <see cref="NatsRuntime"/> is refreshed
        /// (<see cref="NatsRuntime.ResetAfterAddInReload"/>) and keeps its live
        /// instance. This is the shape of an add-in reload; a same-process reload
        /// must never tear the shared runtime down for the next caller.
        /// </summary>
        private static void PhaseStartupTeardown()
        {
            PhaseBanner("S", "runtime startup/shutdown/refresh transition");
            var sharedConnectionsBefore = NatsRuntime.Connections;
            var sharedSubscriptionsBefore = NatsRuntime.Subscriptions;

            // A test-only manager pair: create, use, tear down. It must not touch
            // the shared runtime instance.
            var throwaway = new NatsConnectionManager();
            var throwawaySubs = new NatsSubscriptionManager(throwaway);
            string subject = "test.natsexcel.s." + Guid.NewGuid().ToString("N");
            int got = 0;
            var token = throwawaySubs.Subscribe(Host, subject, null, _ => Interlocked.Increment(ref got), "S");
            bool delivered = false;
            try
            {
                var connection = throwaway.GetConnection(Host);
                Observe(connection.ConnectAsync().AsTask());
                for (int i = 0; i < 40 && !delivered; i++)
                {
                    Observe(connection.PublishAsync(subject, "s-" + i).AsTask());
                    delivered = WaitUntil(() => Volatile.Read(ref got) > 0, 250) >= 0;
                }
            }
            catch (Exception ex)
            {
                Metric("S", "throwawayError", ShortType(ex));
            }
            Check("S.throwaway-manager-round-trip", delivered,
                "a fresh manager/subscription pair delivered=" + delivered);

            token.Dispose();
            throwawaySubs.Dispose();
            throwaway.Shutdown();
            Metric("S", "throwawayConnectionsAfterShutdown", throwaway.LiveConnectionCount());
            Check("S.throwaway-shutdown-owned", throwaway.LiveConnectionCount() == 0,
                "LiveConnectionCount after Shutdown=" + throwaway.LiveConnectionCount());

            // The refresh is a no-op when the runtime was not shut down; the shared
            // instance must survive it unchanged.
            NatsRuntime.ResetAfterAddInReload();
            bool sameConnections = ReferenceEquals(sharedConnectionsBefore, NatsRuntime.Connections);
            bool sameSubscriptions = ReferenceEquals(sharedSubscriptionsBefore, NatsRuntime.Subscriptions);
            Check("S.shared-runtime-survives-refresh", sameConnections && sameSubscriptions,
                "Connections identity preserved=" + sameConnections
                    + " Subscriptions identity preserved=" + sameSubscriptions);
            DumpLatencies("S");
        }

        // ------------------------------------- A0: churn baseline hot path

        private static void PhaseA0()
        {
            long start = PhaseBanner("A0", "churn baseline hot path (no publish flood)");

            var leftovers = new ConcurrentQueue<IDisposable>();
            var churn = new Worker("A0-churn", 4, (id, n) =>
            {
                string subject = "test.natsexcel.a0.churn." + ((n / 2) % 16);
                long t0 = Stopwatch.GetTimestamp();
                IDisposable token = S.Subscribe(Host, subject, null, NoopMessage, "A0");
                if ((n % 7) == 0) leftovers.Enqueue(token);
                else token.Dispose();
                GetLat("A0.subscribe-dispose").Record(Stopwatch.GetTimestamp() - t0);
            });

            churn.Start();
            // Soak on the observed churn counter, not a fixed window: a healthy
            // run reaches the floor in a fraction of the window and exits early,
            // while a wedged/slow run still runs the whole window (the maximum is
            // unchanged) and then fails the assertions below. The floor is a
            // ~3x margin over the A0.churn-hot-path assertion floor, so the phase
            // still executes tens of thousands of real subscribe/dispose ops.
            WaitUntil(() => churn.Ops >= (Quick ? 20000 : 80000), Quick ? 1500 : 4000);
            churn.StopAndJoin();

            int disposeErrors = 0;
            foreach (var token in leftovers)
            {
                try { token.Dispose(); }
                catch (Exception ex) { disposeErrors++; Metric("A0", "disposeError", ex.GetType().Name); }
            }

            Metric("A0", "churnOps", churn.Ops);
            Metric("A0", "churnErrors", churn.Errors);
            Metric("A0", "churnFirstError", churn.FirstError ?? "-");
            Metric("A0", "listenerCount", S.ListenerCount);
            Metric("A0", "subscriptionCount", S.SubscriptionCount);

            Check("A0.churn-hot-path", churn.Ops >= (Quick ? 200 : 500) && churn.Errors == 0,
                "ops=" + churn.Ops + " errors=" + churn.Errors + " first=" + (churn.FirstError ?? "-"));

            // No listener leak: everything disposed, so only the delivery probe's
            // one listener/subscription may remain.
            long settled = WaitUntil(() => S.ListenerCount == 1 && S.SubscriptionCount == 1,
                Quick ? 5000 : 10000);
            Check("A0.no-listener-leak", settled >= 0,
                "ListenerCount=" + S.ListenerCount + " SubscriptionCount=" + S.SubscriptionCount
                    + " (only the delivery probe must remain) disposeErrors=" + disposeErrors);

            // Delivery still works after the churn.
            long prev = Delivery.Count;
            double resume = Delivery.ResumeAfter(prev, ResumeBoundMs);
            Check("A0.delivery-alive", resume >= 0 && resume <= ResumeBoundMs, "resume=" + Fmt(resume) + "ms");

            NoteWorst("deliveryResumeMs", resume);
            NoteWorst("deliveryGapMs", Delivery.MaxGapMs);
            Metric("A0", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("A0");
        }

        // ------------------------------------------- A: ThreadPool starvation

        private static void PhaseA()
        {
            long start = PhaseBanner("A", "ThreadPool starvation + churn + async backlog");

            var release = StarvePool(out int minW, out int minIo, out int maxW, out int maxIo, out bool setMin, out bool setMax, 500);
            Check("A.pool-guard-applied", setMin && setMax, "SetMin(4,4)=" + setMin + " SetMax(4,4)=" + setMax);
            if (!(setMin && setMax))
            {
                RestorePool(minW, minIo, maxW, maxIo, release);
                return;
            }

            // The pool heartbeat is created AFTER the pool is exhausted, so its
            // first continuation cannot run until release: the gap measured after
            // release spans the whole attack window (positive proof of starvation).
            var pool = new PoolHeartbeat();
            pool.Start();
            pool.Reset();
            long poolTicksAtStarveStart = pool.Ticks;
            int attackMs = Quick ? 2500 : 6000;

            // Backlog enqueued while starved: its continuations need pool threads,
            // so it cannot drain until the pool is released.
            int backlog = Quick ? 60 : 150;
            string queueKey = Host + "|queueA";
            var drained = new CountdownEvent(backlog);
            for (int i = 0; i < backlog; i++)
            {
                Task<object> ignored = NatsUdfAsync.Enqueue(queueKey, () => { drained.Signal(); return (object)null; });
                GC.KeepAlive(ignored);
            }

            long prevDeliveries = Delivery.Count;
            var leftovers = new ConcurrentQueue<IDisposable>();
            var churn = new Worker("A-churn", 4, (id, n) =>
            {
                string subject = "test.natsexcel.a.churn." + ((n / 2) % 16);
                long t0 = Stopwatch.GetTimestamp();
                IDisposable token = S.Subscribe(Host, subject, null, NoopMessage, "A");
                if ((n % 4) == 0) leftovers.Enqueue(token);
                else token.Dispose();
                GetLat("A.subscribe-dispose").Record(Stopwatch.GetTimestamp() - t0);
            });
            var pubs = new Worker("A-publish", 2, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                var connection = C.GetConnection(Host);
                Observe(connection.PublishAsync("test.natsexcel.a.churn." + ((n / 2) % 16), "A-" + id + "-" + n).AsTask());
                GetLat("A.getconnection-publish").Record(Stopwatch.GetTimestamp() - t0);
            });

            churn.Start();
            pubs.Start();
            Thread.Sleep(attackMs);
            churn.StopAndJoin();
            pubs.StopAndJoin();

            long poolTicksDuringAttack = pool.Ticks - poolTicksAtStarveStart;
            int queuePendingAtAttackEnd = drained.CurrentCount;

            // Release the pool: sleepers end, continuations and the heartbeat resume.
            RestorePool(minW, minIo, maxW, maxIo, release);
            WaitUntil(() => pool.Ticks > poolTicksAtStarveStart, 3000);
            double poolGapAfterRelease = pool.MaxGapMs;

            long drainStart = Stopwatch.GetTimestamp();
            bool drainedOk = drained.Wait((int)QueueDrainBoundMs);
            double drainMs = MsSince(drainStart);
            double resumeMs = Delivery.ResumeAfter(prevDeliveries, ResumeBoundMs);

            Metric("A", "churnOps", churn.Ops);
            Metric("A", "churnErrors", churn.Errors);
            Metric("A", "churnFirstError", churn.FirstError ?? "-");
            Metric("A", "publishOps", pubs.Ops);
            Metric("A", "publishErrors", pubs.Errors);
            Metric("A", "poolTicksDuringAttack", poolTicksDuringAttack);
            Metric("A", "poolHeartbeatGapAfterMs", poolGapAfterRelease);
            Metric("A", "queuePendingAtAttackEnd", queuePendingAtAttackEnd);
            Metric("A", "deliveriesDuringAttack", Delivery.Count - prevDeliveries);
            Metric("A", "deliveryResumeMs", resumeMs);
            Metric("A", "queueDrainMs", drainMs);

            // The pool heartbeat cannot tick while the clamped pool runs only the
            // blocked sleepers, so the first post-release gap spans the whole
            // window: that is positive proof the pool really was starved (a counter
            // stuck at 0 could also mean the heartbeat thread died).
            Check("A.pool-was-starved", poolTicksDuringAttack == 0 && queuePendingAtAttackEnd > 0
                    && poolGapAfterRelease >= attackMs - 1000,
                "pool ticks during attack=" + poolTicksDuringAttack
                    + " queued items stuck=" + queuePendingAtAttackEnd
                    + " gap after release=" + Fmt(poolGapAfterRelease) + "ms (window " + attackMs + "ms)");
            Check("A.dedicated-heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
            Check("A.entrypoints-not-stuck", Watchdog.MaxIterMs < StuckBoundMs, "watchdog max iteration=" + Fmt(Watchdog.MaxIterMs) + "ms");
            Check("A.delivery-flows-during-starvation", Delivery.Count > prevDeliveries,
                "deliveries during attack=" + (Delivery.Count - prevDeliveries) + " (inline fan-out keeps flowing)");
            Check("A.queue-drains-within-bound", drainedOk,
                "queueDrainMs=" + Fmt(drainMs) + " (bound " + Fmt(QueueDrainBoundMs) + "ms)");
            Check("A.delivery-resumes-within-bound", resumeMs >= 0 && resumeMs <= ResumeBoundMs,
                "resumeMs=" + Fmt(resumeMs) + " (bound " + Fmt(ResumeBoundMs) + "ms)");
            Check("A.no-unexpected-errors", churn.Errors == 0 && pubs.Errors == 0,
                "churnErrors=" + churn.Errors + " publishErrors=" + pubs.Errors
                    + " first=" + (churn.FirstError ?? pubs.FirstError ?? "-"));

            foreach (var token in leftovers) { try { token.Dispose(); } catch { } }
            pool.Stop();
            NoteWorst("poolHeartbeatGapMs", poolGapAfterRelease);
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            NoteWorst("deliveryResumeMs", resumeMs);
            NoteWorst("queueDrainMs", drainMs);
            NoteWorst("watchdogMs", Watchdog.MaxIterMs);
            Metric("A", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("A");
        }

        // ------------------------------------------------ B: lock contention

        private static void PhaseB()
        {
            long start = PhaseBanner("B", "lock contention on every public entry point");

            int listenersBefore = S.ListenerCount;
            int subscriptionsBefore = S.SubscriptionCount;
            long counterReads = 0;
            long negativeCounters = 0;
            long oversizedCounters = 0;
            long udfErrors = 0;
            long udfTimeouts = 0;
            int workers = Quick ? 1 : 2;
            var leftovers = new ConcurrentQueue<IDisposable>();
            var latestSubjects = new string[4];
            for (int i = 0; i < latestSubjects.Length; i++)
            {
                latestSubjects[i] = "test.natsexcel.b.latest." + i;
                TrackedSubjects.Add(latestSubjects[i]);
            }

            var list = new List<Worker>();

            list.Add(new Worker("B-subscribe", workers, (id, n) =>
            {
                string subject = "test.natsexcel.b.churn." + ((n / 2) % 24);
                long t0 = Stopwatch.GetTimestamp();
                IDisposable token = S.Subscribe(Host, subject, null, NoopMessage, "B");
                if ((n % 5) == 0) leftovers.Enqueue(token);
                else token.Dispose();
                GetLat("B.subscribe-dispose").Record(Stopwatch.GetTimestamp() - t0);
            }));

            list.Add(new Worker("B-counters", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                int subscriptions = S.SubscriptionCount;
                int listeners = S.ListenerCount;
                int originListeners = S.ListenerCountWithOrigin("B");
                int originSubjects = S.SubjectCountWithOrigin("B");
                int live = C.LiveConnectionCount();
                int liveRtd = C.LiveRtdConnectionCount();
                int liveUdf = C.LiveUdfConnectionCount();
                bool connected = C.HasConnection(Host);
                bool jetStream = C.IsJetStreamEnabled(Host);
                GetLat("B.counters").Record(Stopwatch.GetTimestamp() - t0);

                // Consistency: no counter may go negative, and none may run away
                // (an unbounded growth would be a leak, which phase E pins again).
                Interlocked.Increment(ref counterReads);
                if (subscriptions < 0 || listeners < 0 || originListeners < 0 || originSubjects < 0 || live < 0)
                    Interlocked.Increment(ref negativeCounters);
                if (subscriptions > 100000 || listeners > 100000)
                    Interlocked.Increment(ref oversizedCounters);
                GC.KeepAlive(connected); GC.KeepAlive(jetStream); GC.KeepAlive(liveRtd + liveUdf);
            }));

            list.Add(new Worker("B-getconnection", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                var connection = C.GetConnection(Host);
                var context = C.GetJetStream(Host);
                GetLat("B.getconnection-js").Record(Stopwatch.GetTimestamp() - t0);
                GC.KeepAlive(connection);
                GC.KeepAlive(context);
            }));

            list.Add(new Worker("B-publish", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                var connection = C.GetConnection(Host);
                Observe(connection.PublishAsync("test.natsexcel.b.churn." + ((n / 2) % 24), "B-" + n).AsTask());
                GetLat("B.publish").Record(Stopwatch.GetTimestamp() - t0);
            }));

            list.Add(new Worker("B-boundedwait", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                bool completed = NatsConnectionManager.WaitBounded(Task.Delay(1), 2000);
                bool timedOut = NatsConnectionManager.WaitBounded(Task.Delay(1500), 30);
                try
                {
                    NatsConnectionManager.WaitBoundedResult(Task.FromResult(42), 2000);
                }
                catch (Exception ex)
                {
                    Interlocked.Increment(ref udfErrors);
                    NoteBError("WaitBoundedResult: " + ShortType(ex));
                }
                GetLat("B.boundedwait").Record(Stopwatch.GetTimestamp() - t0);
                if (!completed || timedOut)
                {
                    Interlocked.Increment(ref udfErrors);
                    NoteBError("bounded wait wrong: completed=" + completed + " timedOut=" + timedOut);
                }
            }));

            list.Add(new Worker("B-udf-publish", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                object result = NatsUdfCore.NatsUDFPublish("test.natsexcel.b.udf." + (n % 4), "B-" + n, Host);
                GetLat("B.udf.publish").Record(Stopwatch.GetTimestamp() - t0);
                if (IsError(result))
                {
                    if (IsTimeoutText(result)) Interlocked.Increment(ref udfTimeouts);
                    else
                    {
                        Interlocked.Increment(ref udfErrors);
                        NoteBError(Convert.ToString(result, Inv));
                    }
                }
            }));

            list.Add(new Worker("B-udf-pif", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                object result = NatsUdfCore.NatsUDFPublishIfChanged("test.natsexcel.b.pif." + (n % 4), "pif-" + n, Host);
                GetLat("B.udf.publishIfChanged").Record(Stopwatch.GetTimestamp() - t0);
                if (IsError(result) && !IsTimeoutText(result))
                {
                    Interlocked.Increment(ref udfErrors);
                    NoteBError(Convert.ToString(result, Inv));
                }
            }));

            list.Add(new Worker("B-udf-request", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                object result = NatsUdfCore.NatsUDFRequest("test.natsexcel.b.req." + (n % 4), "r-" + n, 250, Host);
                GetLat("B.udf.request").Record(Stopwatch.GetTimestamp() - t0);
                // No responder exists: the frozen marker is "No Responders" or
                // "TIMEOUT"; any other outcome is a real defect.
                string text = Convert.ToString(result, Inv);
                bool expected = text == NatsUdfCommon.NoResponders
                    || text == NatsUdfCommon.Timeout
                    || text == NatsUdfCommon.FireAndForget
                    || text == NatsUdfCommon.FireAndForgetAll;
                if (!expected && IsError(result))
                {
                    if (IsTimeoutText(result)) Interlocked.Increment(ref udfTimeouts);
                    else
                    {
                        Interlocked.Increment(ref udfErrors);
                        NoteBError(text);
                    }
                }
                else if (!expected)
                {
                    Interlocked.Increment(ref udfErrors);
                    NoteBError("unexpected request result: " + Shorten(text ?? "<null>", 160));
                }
            }));

            list.Add(new Worker("B-subjectlatest", workers, (id, n) =>
            {
                string subject = latestSubjects[n % latestSubjects.Length];
                long t0 = Stopwatch.GetTimestamp();
                NatsUdfCore.NatsUDFSubjectLatest(subject, null, Host);
                GetLat("B.subjectlatest").Record(Stopwatch.GetTimestamp() - t0);
                long t1 = Stopwatch.GetTimestamp();
                NatsUdfCore.NatsUDFSubjectUnsubscribe(subject, Host);
                GetLat("B.subjectunsubscribe").Record(Stopwatch.GetTimestamp() - t1);
            }));

            list.Add(new Worker("B-runtime", 1, (id, n) =>
            {
                long t0 = Stopwatch.GetTimestamp();
                var connections = NatsRuntime.Connections;
                var subscriptions = NatsRuntime.Subscriptions;
                double count = NatsUdfCore.NatsUDFSubscriptionCount();
                GetLat("B.runtime").Record(Stopwatch.GetTimestamp() - t0);
                GC.KeepAlive(connections);
                GC.KeepAlive(subscriptions);
                GC.KeepAlive(count);
            }));

            foreach (var w in list) w.Start();
            // Flood on real progress from EVERY public entry point, not a fixed
            // window: the flood stops once each entry point has really been
            // hammered (every worker counts only SUCCESSFUL ops, so a wedged entry
            // point never reaches its floor), while a stuck/slow entry point still
            // consumes the whole window (the same maximum) and then fails
            // B.entrypoints-not-stuck / B.no-unexpected-errors below.
            WaitUntil(() =>
            {
                for (int i = 0; i < list.Count; i++)
                    if (list[i].Ops < (Quick ? 25 : 60)) return false;
                return true;
            }, Quick ? 3000 : 8000);
            foreach (var w in list) w.StopAndJoin();
            foreach (var w in list)
                Console.WriteLine("[METRIC] phase=B worker=" + w.Name + " ops=" + w.Ops + " errors=" + w.Errors
                    + " firstError=" + (w.FirstError ?? "-"));

            long ops = 0, workerErrors = 0;
            for (int i = 0; i < list.Count; i++) { ops += list[i].Ops; workerErrors += list[i].Errors; }

            // Post-flood recovery: a fresh subscription must receive a delivery
            // again once the flood stops. The wire SUBSCRIBE is asynchronous, so
            // the probe re-publishes and keeps the listener alive until one lands
            // (a fresh subscription must succeed, but not necessarily on the very
            // first publish).
            string recoverError = null;
            bool recovered = false;
            string recoverSubject = "test.natsexcel.b.recover." + Guid.NewGuid().ToString("N");
            int recoveredCount = 0;
            long recoveredAt = WaitUntil(() =>
            {
                try
                {
                    var token = S.Subscribe(Host, recoverSubject, null, _ => Interlocked.Increment(ref recoveredCount), "B");
                    try
                    {
                        var connection = C.GetConnection(Host);
                        for (int i = 0; i < 40; i++)
                        {
                            Observe(connection.PublishAsync(recoverSubject, "recover-" + i).AsTask());
                            if (WaitUntil(() => Volatile.Read(ref recoveredCount) > 0, 250) >= 0)
                            {
                                recovered = true;
                                return true;
                            }
                        }
                        return false;
                    }
                    finally { token.Dispose(); }
                }
                catch (Exception ex)
                {
                    recoverError = ShortType(ex);
                    return false;
                }
            }, ResumeBoundMs);

            foreach (var token in leftovers) { try { token.Dispose(); } catch { } }
            foreach (string subject in latestSubjects) { try { NatsUdfCore.NatsUDFSubjectUnsubscribe(subject, Host); } catch { } }

            long settled = WaitUntil(() => S.ListenerCount == listenersBefore && S.SubscriptionCount == subscriptionsBefore,
                Quick ? 8000 : 15000);

            Metric("B", "ops", ops);
            Metric("B", "workerErrors", workerErrors);
            Metric("B", "counterReads", counterReads);
            Metric("B", "negativeCounters", negativeCounters);
            Metric("B", "oversizedCounters", oversizedCounters);
            Metric("B", "udfErrors", udfErrors);
            Metric("B", "udfTimeouts", udfTimeouts);
            Metric("B", "firstUdfError", FirstBError ?? "-");
            Metric("B", "listenerCountAfter", S.ListenerCount);
            Metric("B", "subscriptionCountAfter", S.SubscriptionCount);
            Metric("B", "recoverMs", recoveredAt);
            Metric("B", "procHeartbeatMs", Proc.MaxGapMs);
            Metric("B", "watchdogMaxMs", Watchdog.MaxIterMs);

            Check("B.entrypoints-not-stuck", Watchdog.MaxIterMs < StuckBoundMs && Proc.MaxGapMs < StuckBoundMs,
                "watchdog max=" + Fmt(Watchdog.MaxIterMs) + "ms proc heartbeat=" + Fmt(Proc.MaxGapMs) + "ms");
            Check("B.counters-consistent", negativeCounters == 0 && oversizedCounters == 0,
                "reads=" + counterReads + " negative=" + negativeCounters + " oversized=" + oversizedCounters);
            Check("B.no-unexpected-errors", workerErrors == 0 && udfErrors == 0,
                "workerErrors=" + workerErrors + " udfErrors=" + udfErrors + " udfTimeouts=" + udfTimeouts
                    + " first=" + (FirstBError ?? "-"));
            Check("B.subscribe-recovers-after-flood", recovered,
                "recovery subscribe succeeded once the flood stopped" + (recoverError == null ? "" : " lastError=" + recoverError));
            Check("B.counters-return-to-baseline", settled >= 0,
                "ListenerCount=" + S.ListenerCount + " (baseline " + listenersBefore + ") SubscriptionCount="
                    + S.SubscriptionCount + " (baseline " + subscriptionsBefore + ")");

            NoteWorst("watchdogMs", Watchdog.MaxIterMs);
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            Metric("B", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("B");
        }

        // ------------------------------------- C: fault injection (stop/restart)

        private static void PhaseC()
        {
            long start = PhaseBanner("C", "stop/restart the local nats-server + outage burst");

            int port = ResolvePort(Host);
            if (!IsLocalEndpoint())
            {
                Console.WriteLine("[SKIP] phase=C sub=server-restart reason=non-local host " + Host
                    + " (the destructive stop/restart only runs against a loopback server)");
            }
            else if (SkipRestart)
            {
                Console.WriteLine("[SKIP] phase=C sub=server-restart reason=--skip-restart");
            }
            else if (!ProbeNats(port, 3000))
            {
                Console.WriteLine("[SKIP] phase=C sub=server-restart reason=no NATS listener on port " + port);
            }
            else
            {
                RunServerRestartPhase(port);
            }

            // The outage burst runs in every mode: subscribe/dispose churn under
            // whatever fault ran (or nothing) must never let an exception escape a
            // callback and must leave the counters balanced.
            OutageBurst("C");

            Metric("C", "procHeartbeatMs", Proc.MaxGapMs);
            Metric("C", "liveConnections", C.LiveConnectionCount());
            Check("C.heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            NoteWorst("deliveryGapMs", Delivery.MaxGapMs);
            Metric("C", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("C");
        }

        private static void RunServerRestartPhase(int port)
        {
            if (!EnsureServerIdentity(port))
            {
                Console.WriteLine("[SKIP] phase=C sub=server-restart reason=could not resolve the nats-server process on port " + port);
                return;
            }

            string exePath = ServerExePath;
            string commandLine = ServerCommandLine;
            int pid = ServerPid;
            Metric("C", "serverPid", pid);
            Metric("C", "serverExe", Shorten(exePath, 160));

            long prevDeliveries = Delivery.Count;

            // --- stop ---
            Console.WriteLine("[PHASE] id=C1 desc=stop nats-server pid=" + pid);
            if (!KillProcess(pid))
            {
                Failures.Enqueue("C.restart :: could not stop the nats-server process (pid " + pid + ")");
                return;
            }
            ServerPid = 0; // stale after the kill; refreshed by the restart
            long down = WaitUntil(() => !ProbeNats(port, 500), ResumeBoundMs / 2);
            Metric("C", "serverDownAfterMs", down);
            Check("C.server-stopped", down >= 0, "the port stopped answering within the wait");

            // An async backlog whose work targets the SAME endpoint while the
            // server is unreachable: every item must still complete.
            string queueKey = Host + "|c";
            int backlog = Quick ? 10 : 25;
            var tasks = new Task<object>[backlog];
            for (int i = 0; i < backlog; i++)
            {
                int n = i;
                tasks[n] = NatsUdfAsync.Enqueue(queueKey, () =>
                    NatsUdfCore.NatsUDFPublish("test.natsexcel.c.outage." + n, "c-" + n, Host));
            }

            OutageBurst("C-outage");

            // --- restart ---
            Console.WriteLine("[PHASE] id=C2 desc=restart nats-server");
            long up = RestartServerBounded(exePath, commandLine, port, ResumeBoundMs);
            Metric("C", "serverRestarted", up >= 0);
            Metric("C", "serverUpAfterMs", up);
            Check("C.server-restarted", up >= 0, "server answered again after " + Fmt(up) + "ms");

            double resume = Delivery.ResumeAfter(prevDeliveries, ResumeBoundMs);
            Metric("C", "deliveryResumeMs", resume);
            Check("C.delivery-resumes-within-bound", resume >= 0 && resume <= ResumeBoundMs,
                "resumeMs=" + Fmt(resume) + " (bound " + Fmt(ResumeBoundMs) + "ms)");
            Check("C.delivery-gap-bounded", Delivery.MaxGapMs < ResumeBoundMs,
                "max inter-arrival gap across the outage=" + Fmt(Delivery.MaxGapMs) + "ms (bound " + Fmt(ResumeBoundMs) + "ms)");
            NoteWorst("deliveryResumeMs", resume);

            // The outage backlog must drain: every queued publish completes once
            // the server is back.
            long drainStart = Stopwatch.GetTimestamp();
            // Deliberate bounded wait (a liveness bound, not an Excel-thread wait):
            // the whole point of this phase is to time how long the backlog takes.
#pragma warning disable VSTHRD002
            bool drained = Task.WaitAll(tasks, (int)QueueDrainBoundMs);
#pragma warning restore VSTHRD002
            double drainMs = MsSince(drainStart);
            int completed = 0;
            for (int i = 0; i < tasks.Length; i++) if (tasks[i].IsCompleted) completed++;
            Metric("C", "outageBacklogCompleted", completed + "/" + backlog);
            Metric("C", "outageBacklogDrainMs", drainMs);
            Check("C.outage-backlog-drains", drained && completed == backlog,
                "completed=" + completed + "/" + backlog + " drainMs=" + Fmt(drainMs)
                    + " (bound " + Fmt(QueueDrainBoundMs) + "ms)");
            NoteWorst("queueDrainMs", drainMs);
        }

        /// <summary>Subscribe/dispose churn under the current fault; a callback or
        /// a Subscribe exception must never escape.</summary>
        private static void OutageBurst(string phase)
        {
            var leftovers = new ConcurrentQueue<IDisposable>();
            var churn = new Worker(phase + "-outage-churn", 4, (id, n) =>
            {
                string subject = "test.natsexcel.c.burst." + ((n / 2) % 12);
                long t0 = Stopwatch.GetTimestamp();
                IDisposable token = S.Subscribe(Host, subject, null, ListenNoop, "C");
                if ((n % 3) == 0) leftovers.Enqueue(token);
                else token.Dispose();
                GetLat(phase + ".subscribe-dispose").Record(Stopwatch.GetTimestamp() - t0);
            });
            churn.Start();
            // Soak on the observed churn counter, not a fixed window: the burst
            // stops as soon as it has really churned, while a stuck/erroring burst
            // (ops only counts successes) still runs the whole window (the maximum
            // is unchanged) and then fails the no-exception assertion below.
            WaitUntil(() => churn.Ops >= (Quick ? 5000 : 15000), Quick ? 800 : 2000);
            churn.StopAndJoin();
            foreach (var token in leftovers) { try { token.Dispose(); } catch { } }

            Metric(phase, "burstOps", churn.Ops);
            Metric(phase, "burstErrors", churn.Errors);
            Metric(phase, "burstFirstError", churn.FirstError ?? "-");
            Check(phase + ".outage-burst-no-exception", churn.Errors == 0,
                "ops=" + churn.Ops + " errors=" + churn.Errors + " first=" + (churn.FirstError ?? "-"));
        }

        // ------------------------------------ D: async queue/observable bursts

        private static void PhaseD()
        {
            long start = PhaseBanner("D", "async queue/observable bursts");

            // --- D1: FIFO queue burst ---
            string[] hosts = { Host + "|d1a", Host + "|d1b" };
            const int producersPerHost = 2;
            int itemsPerProducer = Quick ? 60 : 150;
            int total = hosts.Length * producersPerHost * itemsPerProducer;
            var done = new CountdownEvent(total);
            var active = new int[hosts.Length];
            var maxActive = new int[hosts.Length];
            var lastSeq = new Dictionary<string, int>();
            object seqGate = new object();
            int fifoViolations = 0;
            var producers = new List<Thread>();

            for (int h = 0; h < hosts.Length; h++)
            {
                int hi = h;
                string hostKey = hosts[h];
                for (int p = 0; p < producersPerHost; p++)
                {
                    int pi = p;
                    var t = new Thread(() =>
                    {
                        for (int i = 0; i < itemsPerProducer; i++)
                        {
                            int item = i;
                            Task<object> ignored = NatsUdfAsync.Enqueue(hostKey, () =>
                            {
                                int a = Interlocked.Increment(ref active[hi]);
                                UpdateMax(ref maxActive[hi], a);
                                lock (seqGate)
                                {
                                    string key = hi + "|" + pi;
                                    if (lastSeq.TryGetValue(key, out int prev) && prev + 1 != item) fifoViolations++;
                                    lastSeq[key] = item;
                                }
                                Thread.SpinWait(2000);
                                Interlocked.Decrement(ref active[hi]);
                                done.Signal();
                                return (object)item;
                            });
                        }
                    });
                    t.IsBackground = true;
                    t.Name = "D1-producer-" + h + "-" + p;
                    producers.Add(t);
                }
            }

            long drainStart = Stopwatch.GetTimestamp();
            foreach (var t in producers) t.Start();
            bool drainedOk = done.Wait((int)QueueDrainBoundMs);
            foreach (var t in producers) t.Join(10000);
            double drainMs = MsSince(drainStart);
            bool overlap = false;
            for (int i = 0; i < maxActive.Length; i++) if (maxActive[i] > 1) overlap = true;
            bool queuesGone = WaitUntil(() => !NatsUdfAsync.HasQueueForTests(hosts[0]) && !NatsUdfAsync.HasQueueForTests(hosts[1]), 5000) >= 0;

            Metric("D", "items", total);
            Metric("D", "drainMs", drainMs);
            Metric("D", "fifoViolations", fifoViolations);
            Metric("D", "maxConcurrentPerHost", string.Join(",", maxActive));
            Check("D.all-executed", drainedOk, "completed " + (total - done.CurrentCount) + "/" + total
                + " within " + Fmt(QueueDrainBoundMs) + "ms");
            Check("D.fifo-per-producer", fifoViolations == 0, "violations=" + fifoViolations);
            Check("D.no-overlap-per-host", !overlap, "maxConcurrentPerHost=" + string.Join(",", maxActive));
            Check("D.queues-removed-when-idle", queuesGone, "idle queue entries removed");
            NoteWorst("queueDrainMs", drainMs);

            // --- D2: observable burst under ThreadPool starvation ---
            string hostKey2 = Host + "|d2";
            int count = Quick ? 40 : 60;
            var release = StarvePool(out int minW, out int minIo, out int maxW, out int maxIo, out bool setMin, out bool setMax, 300);

            var observers = new RecordingObserver[count];
            for (int i = 0; i < count; i++)
            {
                int tag = i;
                observers[i] = new RecordingObserver();
                new NatsUdfAsync.NatsWriteObservable(hostKey2, () => (object)tag).Subscribe(observers[i]);
            }

            Thread.Sleep(Quick ? 800 : 1500);
            int completedWhileStarved = CountCompleted(observers);
            RestorePool(minW, minIo, maxW, maxIo, release);

            long observableDrainStart = Stopwatch.GetTimestamp();
            long drainedAt = WaitUntil(() => CountCompleted(observers) == count, QueueDrainBoundMs);
            double observableDrainMs = drainedAt < 0 ? -1 : MsSince(observableDrainStart);
            int deliveredWrong = 0;
            for (int i = 0; i < count; i++)
            {
                object[] values = observers[i].Values;
                if (observers[i].CompletedCount != 1 || values.Length != 1 || !(values[0] is int v) || v != i)
                    deliveredWrong++;
            }
            bool queueGone2 = WaitUntil(() => !NatsUdfAsync.HasQueueForTests(hostKey2), 5000) >= 0;

            Metric("D2", "items", count);
            Metric("D2", "completedWhileStarved", completedWhileStarved);
            Metric("D2", "drainMs", observableDrainMs);
            Metric("D2", "wrongDeliveries", deliveredWrong);
            Check("D2.pool-guard-applied", setMin && setMax, "SetMin(4,4)=" + setMin + " SetMax(4,4)=" + setMax);
            Check("D2.every-queued-item-completes", drainedAt >= 0,
                "completed=" + CountCompleted(observers) + "/" + count
                    + " within " + Fmt(QueueDrainBoundMs) + "ms (no lost item)");
            Check("D2.exactly-once-each", deliveredWrong == 0, "wrong deliveries=" + deliveredWrong);
            Check("D2.queue-removed", queueGone2, "idle queue entry removed");
            Check("D2.heartbeat", Proc.MaxGapMs < StuckBoundMs, "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms");
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            NoteWorst("queueDrainMs", observableDrainMs);
            Metric("D", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("D");
        }

        // ------------------------------- F: reconnect storm (server flapping)

        /// <summary>
        /// The failure a user meets when the NATS server is restarted repeatedly:
        /// kill and restart the local server several times in quick succession
        /// while a subscription is live, and assert delivery resumes within the
        /// bound on every cycle, the counters return to their baseline (no
        /// duplicate listener) and a throwing consumer never breaks fan-out.
        /// Destructive: only against a loopback server (skipped with a reason).
        /// </summary>
        private static void PhaseF()
        {
            long start = PhaseBanner("F", "reconnect storm (burst of server kill/restart)");

            if (!CanRestartServer("F"))
            {
                Metric("F", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("F");
                return;
            }

            int port = ResolvePort(Host);
            if (!EnsureServerIdentity(port))
            {
                Console.WriteLine("[SKIP] phase=F sub=server-restart reason=could not resolve the nats-server process on port " + port);
                Metric("F", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("F");
                return;
            }
            string exePath = ServerExePath;
            string commandLine = ServerCommandLine;
            Metric("F", "serverPid", ServerPid);

            int listenersBefore = S.ListenerCount;
            int subscriptionsBefore = S.SubscriptionCount;
            string subject = "test.natsexcel.f.storm." + Guid.NewGuid().ToString("N");
            long goodCount = 0;
            long throwCount = 0;

            // Two listeners share ONE physical subscription. The second
            // deliberately throws: a throwing consumer must never break fan-out
            // to its sibling or the consume loop.
            IDisposable good = S.Subscribe(Host, subject, null,
                _ => Interlocked.Increment(ref goodCount), "F");
            IDisposable bad = S.Subscribe(Host, subject, null,
                _ =>
                {
                    Interlocked.Increment(ref throwCount);
                    throw new InvalidOperationException("F: deliberate listener failure");
                }, "F");

            // A dedicated publisher keeps traffic on the F subject across the
            // storm, so "delivery resumes" is a real observation and the throwing
            // listener really receives messages. Its publish failures during an
            // outage are expected and observed (never unobserved task exceptions).
            int stopPublisher = 0;
            Thread publisher = new Thread(() =>
            {
                long n = 0;
                while (Volatile.Read(ref stopPublisher) == 0)
                {
                    try
                    {
                        var connection = C.GetConnection(Host);
                        Observe(connection.PublishAsync(subject, "f-" + Interlocked.Increment(ref n)).AsTask());
                    }
                    catch
                    {
                        // The endpoint is down: keep the loop alive, retry next tick.
                    }
                    Thread.Sleep(20);
                }
            });
            publisher.IsBackground = true;
            publisher.Name = "F-publisher";
            publisher.Start();

            try
            {
                // Wait for the subscription to be genuinely live (a first
                // delivery) before the first kill: the journey is "a live
                // subscription survives a reconnect storm".
                long warm = WaitUntil(() => Volatile.Read(ref goodCount) > 0, ResumeBoundMs);
                Check("F.subscription-live-before-storm", warm >= 0, "first delivery in " + Fmt(warm) + "ms");

                int cycles = Quick ? 2 : 3;
                int killed = 0;
                int restarted = 0;
                double worstResume = -1;
                long prevGood = goodCount;

                for (int cycle = 0; cycle < cycles; cycle++)
                {
                    // Reuse the cached pid; the restart below refreshes it.
                    int currentPid = CurrentServerPid(port);
                    if (currentPid <= 0 || !KillProcess(currentPid))
                    {
                        Failures.Enqueue("F.storm :: could not stop the server in cycle " + cycle);
                        break;
                    }
                    killed++;
                    WaitUntil(() => !ProbeNats(port, 500), ResumeBoundMs / 2);

                    // A subscribe/dispose burst while the server is down: a
                    // callback or a Subscribe exception must never escape.
                    OutageBurst("F");

                    long up = RestartServerBounded(exePath, commandLine, port, ResumeBoundMs);
                    if (up < 0)
                    {
                        Failures.Enqueue("F.storm :: the server did not come back in cycle " + cycle);
                        break;
                    }
                    restarted++;

                    long at = WaitUntil(() => Volatile.Read(ref goodCount) > prevGood, ResumeBoundMs);
                    prevGood = goodCount;
                    if (at < 0 || at > ResumeBoundMs)
                        Failures.Enqueue("F.cycle-" + cycle + ".delivery-resumes-within-bound :: resumeMs=" + Fmt(at)
                            + " (bound " + Fmt(ResumeBoundMs) + "ms)");
                    if (at > worstResume) worstResume = at;
                    Metric("F", "cycle" + cycle + "ResumeMs", at);
                }

                Metric("F", "killed", killed);
                Metric("F", "restarted", restarted);
                Metric("F", "witnessDeliveries", goodCount);
                Metric("F", "throwingListenerCalls", throwCount);
                Check("F.storm-completed", killed == cycles && restarted == cycles,
                    "killed=" + killed + "/" + cycles + " restarted=" + restarted + "/" + cycles);
                Check("F.callback-exception-isolated", goodCount > 0 && throwCount > 0,
                    "sibling deliveries=" + goodCount + " throwing-listener calls=" + throwCount
                        + " (fan-out survived a throwing consumer)");
                NoteWorst("deliveryResumeMs", worstResume);
            }
            finally
            {
                Volatile.Write(ref stopPublisher, 1);
                publisher.Join(5000);
                try { good.Dispose(); } catch { }
                try { bad.Dispose(); } catch { }
            }

            long settled = WaitUntil(() => S.ListenerCount == listenersBefore && S.SubscriptionCount == subscriptionsBefore,
                Quick ? 8000 : 15000);
            Check("F.no-listener-leak-after-storm", settled >= 0,
                "ListenerCount=" + S.ListenerCount + " (baseline " + listenersBefore + ") SubscriptionCount="
                    + S.SubscriptionCount + " (baseline " + subscriptionsBefore + ")");
            Check("F.heartbeat", Proc.MaxGapMs < StuckBoundMs && Watchdog.MaxIterMs < StuckBoundMs,
                "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms watchdog max=" + Fmt(Watchdog.MaxIterMs) + "ms");
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            NoteWorst("watchdogMs", Watchdog.MaxIterMs);
            Metric("F", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("F");
        }

        // ------------------------------------- G: requests during an outage

        /// <summary>
        /// Fire request/reply calls while the local server is down: every call
        /// must return a bounded error (No Responders / TIMEOUT / an "Error:")
        /// rather than hang, and a request must succeed again within the resume
        /// bound after the restart. Destructive (skipped with a reason).
        /// </summary>
        private static void PhaseG()
        {
            long start = PhaseBanner("G", "requests during a server outage");

            if (!CanRestartServer("G"))
            {
                Metric("G", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("G");
                return;
            }

            int port = ResolvePort(Host);
            if (!EnsureServerIdentity(port))
            {
                Console.WriteLine("[SKIP] phase=G sub=server-restart reason=could not resolve the nats-server process on port " + port);
                Metric("G", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("G");
                return;
            }
            string exePath = ServerExePath;
            string commandLine = ServerCommandLine;

            // Requests must be real request/reply here: fire-and-forget would
            // never exercise the outage path. Restored in the finally.
            AppConfig.Root saved = AppConfig.CurrentOverrideForTests;
            AppConfig.CurrentOverrideForTests = BuildPhaseConfig(Host, allowAdminWrites: false, syncWrite: "sync");
            try
            {
                string subject = "test.natsexcel.g.req." + Guid.NewGuid().ToString("N");

                // Baseline while up with no responder: an exact bounded marker.
                object upBounded = NatsUdfCore.NatsUDFRequest(subject, "g-0", 500, Host);
                Check("G.request-bounded-when-up", IsBoundedRequestError(upBounded),
                    "up result='" + CellText(upBounded) + "'");

                // A live responder proves a real reply round trip, so the
                // post-restart success is a real success and not just a marker.
                var baselineConn = NewRawConnection(Host);
                var baselineResponder = new RawResponder(baselineConn, subject, p => "reply:" + p);
                try
                {
                    bool ready = baselineResponder.WaitReady(Quick ? 10000 : 20000);
                    object live = ready
                        ? NatsUdfCore.NatsUDFRequest(subject, "g-live", 1000, Host)
                        : "Error: the baseline responder did not become ready";
                    Check("G.request-succeeds-when-up",
                        string.Equals(CellText(live), "reply:g-live", StringComparison.Ordinal),
                        "result='" + CellText(live) + "'");
                }
                finally
                {
                    baselineResponder.Dispose();
                    DisposeConnection(baselineConn);
                }

                // --- outage: fire requests while the server is down ---
                if (!KillProcess(CurrentServerPid(port)))
                {
                    Failures.Enqueue("G :: could not stop the nats-server process (pid " + ServerPid + ")");
                    return;
                }
                ServerPid = 0;
                long down = WaitUntil(() => !ProbeNats(port, 500), ResumeBoundMs / 2);
                Check("G.server-stopped", down >= 0, "the port stopped answering after " + Fmt(down) + "ms");

                int n = Quick ? 3 : 6;
                int bounded = 0;
                int unbounded = 0;
                double worst = 0;
                for (int i = 0; i < n; i++)
                {
                    long t0 = Stopwatch.GetTimestamp();
                    object r = NatsUdfCore.NatsUDFRequest(subject, "g-down-" + i, 500, Host);
                    double ms = MsSince(t0);
                    if (ms > worst) worst = ms;
                    if (IsBoundedRequestError(r)) bounded++;
                    else
                    {
                        unbounded++;
                        Failures.Enqueue("G.request-during-outage :: '" + CellText(r) + "' is not a bounded error");
                    }
                }
                Metric("G", "outageRequests", n);
                Metric("G", "boundedErrors", bounded);
                Metric("G", "worstOutageRequestMs", worst);
                Check("G.every-outage-request-bounded", unbounded == 0 && bounded == n,
                    "bounded=" + bounded + "/" + n + " worst=" + Fmt(worst) + "ms");
                Check("G.outage-request-time-bounded", worst <= ShutdownBoundMs,
                    "worst outage request " + Fmt(worst) + "ms (a hang past the bound fails the run)");

                // --- restart: a request must succeed again within the bound ---
                long up = RestartServerBounded(exePath, commandLine, port, ResumeBoundMs);
                Check("G.server-restarted", up >= 0, "the server answered again after " + Fmt(up) + "ms");

                var afterConn = NewRawConnection(Host);
                var afterResponder = new RawResponder(afterConn, subject, p => "reply:" + p);
                try
                {
                    bool ready = afterResponder.WaitReady(Quick ? 10000 : 20000);
                    long resumeAt = WaitUntil(() =>
                        string.Equals(CellText(NatsUdfCore.NatsUDFRequest(subject, "g-after", 1000, Host)),
                            "reply:g-after", StringComparison.Ordinal), ResumeBoundMs);
                    Metric("G", "requestResumeMs", resumeAt);
                    Check("G.request-succeeds-after-restart", ready && resumeAt >= 0,
                        "first successful request after the restart ms=" + Fmt(resumeAt) + " responderReady=" + ready);
                    NoteWorst("deliveryResumeMs", resumeAt);
                }
                finally
                {
                    afterResponder.Dispose();
                    DisposeConnection(afterConn);
                }
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = saved;
            }

            Check("G.heartbeat", Proc.MaxGapMs < StuckBoundMs && Watchdog.MaxIterMs < StuckBoundMs,
                "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms watchdog max=" + Fmt(Watchdog.MaxIterMs) + "ms");
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            Metric("G", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("G");
        }

        // ------------------------------- H: JetStream during an outage

        /// <summary>
        /// The "server restarted, my stream is still there" journey: create a
        /// stream with a durable consumer, publish before the outage, restart the
        /// server, then assert the stored messages survived (a durable
        /// <c>store_dir</c>) and that a new publish and a fetch both work after
        /// recovery. Destructive (skipped with a reason).
        /// </summary>
        private static void PhaseH()
        {
            long start = PhaseBanner("H", "JetStream stream/consumer across a server outage");

            if (!CanRestartServer("H"))
            {
                Metric("H", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("H");
                return;
            }

            int port = ResolvePort(Host);
            if (!EnsureServerIdentity(port))
            {
                Console.WriteLine("[SKIP] phase=H sub=server-restart reason=could not resolve the nats-server process on port " + port);
                Metric("H", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("H");
                return;
            }
            string exePath = ServerExePath;
            string commandLine = ServerCommandLine;

            if (!NatsRuntime.Connections.IsJetStreamEnabled(Host))
            {
                Console.WriteLine("[SKIP] phase=H sub=jetstream reason=the server does not advertise JetStream (start it with -js)");
                Metric("H", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("H");
                return;
            }

            AppConfig.Root saved = AppConfig.CurrentOverrideForTests;
            AppConfig.CurrentOverrideForTests = BuildPhaseConfig(Host, allowAdminWrites: true, syncWrite: "sync");

            string stream = "live_h_" + Guid.NewGuid().ToString("N").Substring(0, 8);
            string consumer = "live_hc_" + Guid.NewGuid().ToString("N").Substring(0, 8);
            string filter = "test.natsexcel.h.>";
            string subject = "test.natsexcel.h.a";

            try
            {
                object created = NatsUdfJetStream.NatsUDFStreamCreate(stream, filter, null, null, Host);
                Check("H.stream-created", string.Equals(CellText(created), NatsUdfCommon.Ok, StringComparison.Ordinal),
                    "create result='" + CellText(created) + "'");
                if (!string.Equals(CellText(created), NatsUdfCommon.Ok, StringComparison.Ordinal))
                    return;

                object consumerResult = NatsUdfJetStream.NatsUDFConsumerCreate(stream, consumer, filter, null, Host);
                Check("H.durable-consumer-created",
                    string.Equals(CellText(consumerResult), NatsUdfCommon.Ok, StringComparison.Ordinal),
                    "consumer result='" + CellText(consumerResult) + "'");

                object seq1 = NatsUdfJetStream.NatsUDFJSPublish(subject, "h-1", Host);
                object seq2 = NatsUdfJetStream.NatsUDFJSPublish(subject, "h-2", Host);
                object seq3 = NatsUdfJetStream.NatsUDFJSPublish(subject, "h-3", Host);
                Check("H.messages-stored-before-outage",
                    AsDouble(seq1) >= 1 && AsDouble(seq3) == AsDouble(seq2) + 1,
                    "seq1=" + CellText(seq1) + " seq2=" + CellText(seq2) + " seq3=" + CellText(seq3));

                // --- outage ---
                if (!KillProcess(CurrentServerPid(port)))
                {
                    Failures.Enqueue("H :: could not stop the nats-server process (pid " + ServerPid + ")");
                    return;
                }
                ServerPid = 0;
                WaitUntil(() => !ProbeNats(port, 500), ResumeBoundMs / 2);

                long up = RestartServerBounded(exePath, commandLine, port, ResumeBoundMs);
                Check("H.server-restarted", up >= 0, "the server answered again after " + Fmt(up) + "ms");

                // (a) the stored messages survived the restart (durable store_dir).
                long survivedAt = WaitUntil(() =>
                    JsonLong(CellText(NatsUdfJetStream.NatsUDFJSStreamInfo(stream, Host)), "messages") == 3,
                    ResumeBoundMs);
                Metric("H", "streamSurvivedAfterMs", survivedAt);
                Check("H.messages-survived-restart", survivedAt >= 0,
                    "stream '" + stream + "' messages after the restart="
                        + JsonLong(CellText(NatsUdfJetStream.NatsUDFJSStreamInfo(stream, Host)), "messages") + " (expected 3)"
                        + " (bounded wait " + Fmt(survivedAt) + "ms)");

                object replayed = NatsUdfJetStream.NatsUDFJSGetMessage(stream, 1L, Host);
                Check("H.replay-after-restart", string.Equals(CellText(replayed), "h-1", StringComparison.Ordinal),
                    "replayed='" + CellText(replayed) + "'");

                // (b) a NEW publish and a fetch both work after recovery.
                long publishAt = WaitUntil(() =>
                {
                    double seq = AsDouble(NatsUdfJetStream.NatsUDFJSPublish(subject, "h-4", Host));
                    return !double.IsNaN(seq) && seq >= 4;
                }, ResumeBoundMs);
                Metric("H", "newPublishAfterMs", publishAt);
                Check("H.new-publish-after-restart", publishAt >= 0,
                    "the first accepted publish after the restart ms=" + Fmt(publishAt));
                NoteWorst("deliveryResumeMs", publishAt);

                // Settle to four stored messages, then one deterministic fetch.
                WaitUntil(() => JsonLong(CellText(NatsUdfJetStream.NatsUDFJSStreamInfo(stream, Host)), "messages") == 4,
                    Quick ? 4000 : 8000);
                object fetch = NatsUdfJetStream.NatsUDFJSFetch(stream, consumer, 10, 3000, Host);
                string[][] rows = MatrixRows(fetch);
                Metric("H", "fetchedRows", rows.Length);
                Check("H.fetch-after-restart", rows.Length == 4,
                    "the durable consumer fetched " + rows.Length + " rows after recovery (expected 4)");
            }
            finally
            {
                try { NatsUdfJetStream.NatsUDFStreamDelete(stream, Host); } catch { }
                AppConfig.CurrentOverrideForTests = saved;
            }

            Check("H.heartbeat", Proc.MaxGapMs < StuckBoundMs && Watchdog.MaxIterMs < StuckBoundMs,
                "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms watchdog max=" + Fmt(Watchdog.MaxIterMs) + "ms");
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            Metric("H", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("H");
        }

        // ------------------- I: async backlog with a flapping server

        /// <summary>
        /// Enqueue a burst of async writes while the server flaps; the queue must
        /// drain within the bound after recovery with every item accounted for
        /// (completed or errored) - no silently lost item. Destructive (skipped
        /// with a reason).
        /// </summary>
        private static void PhaseI()
        {
            long start = PhaseBanner("I", "async backlog with a flapping server");

            if (!CanRestartServer("I"))
            {
                Metric("I", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("I");
                return;
            }

            int port = ResolvePort(Host);
            if (!EnsureServerIdentity(port))
            {
                Console.WriteLine("[SKIP] phase=I sub=server-restart reason=could not resolve the nats-server process on port " + port);
                Metric("I", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("I");
                return;
            }
            string exePath = ServerExePath;
            string commandLine = ServerCommandLine;

            string queueKey = Host + "|i";
            int backlog = Quick ? 40 : 100;
            const int poisonedEvery = 7;
            var gate = new ManualResetEventSlim(false);
            var tasks = new Task<object>[backlog];
            for (int i = 0; i < backlog; i++)
            {
                int n = i;
                tasks[n] = NatsUdfAsync.Enqueue(queueKey, () =>
                {
                    // Every item is held behind the gate, so the whole backlog is
                    // genuinely pending while the server flaps.
                    gate.Wait(TimeSpan.FromSeconds(30));
                    if (n % poisonedEvery == 0)
                        throw new InvalidOperationException("I: deliberate queued-item failure");
                    NatsUdfCore.NatsUDFPublish("test.natsexcel.i." + (n % 4), "i-" + n, Host);
                    return (object)("item-" + n);
                });
            }

            int cycles = Quick ? 2 : 3;
            int killed = 0;
            int restarted = 0;
            try
            {
                for (int cycle = 0; cycle < cycles; cycle++)
                {
                    if (CurrentServerPid(port) <= 0 || !KillProcess(ServerPid))
                    {
                        Failures.Enqueue("I.flap :: could not stop the server in cycle " + cycle);
                        break;
                    }
                    ServerPid = 0;
                    killed++;
                    WaitUntil(() => !ProbeNats(port, 500), ResumeBoundMs / 2);
                    OutageBurst("I");

                    long up = RestartServerBounded(exePath, commandLine, port, ResumeBoundMs);
                    if (up < 0)
                    {
                        Failures.Enqueue("I.flap :: the server did not come back in cycle " + cycle);
                        break;
                    }
                    restarted++;
                }
            }
            finally
            {
                gate.Set();
            }

            long drainStart = Stopwatch.GetTimestamp();
            bool drained;
#pragma warning disable VSTHRD002 // deliberate bounded wait: this phase times the queue drain.
            try { drained = Task.WaitAll(tasks, (int)QueueDrainBoundMs); }
            catch (AggregateException) { drained = true; /* a faulted item is accounted below */ }
#pragma warning restore VSTHRD002
            double drainMs = MsSince(drainStart);

            int completed = 0;
            int errored = 0;
            int faulted = 0;
            int missing = 0;
            for (int i = 0; i < tasks.Length; i++)
            {
                Task<object> task = tasks[i];
                if (!task.IsCompleted) { missing++; continue; }
                if (task.IsFaulted) { faulted++; continue; }
                completed++;
                var text = BoundedResult(task) as string;
                if (text != null && text.StartsWith(NatsUdfCommon.ErrorPrefix, StringComparison.Ordinal))
                    errored++;
            }

            Metric("I", "items", backlog);
            Metric("I", "killed", killed);
            Metric("I", "restarted", restarted);
            Metric("I", "drainMs", drainMs);
            Metric("I", "completed", completed);
            Metric("I", "errored", errored);
            Metric("I", "faulted", faulted);
            Metric("I", "missing", missing);
            Check("I.flap-completed", killed == cycles && restarted == cycles,
                "killed=" + killed + "/" + cycles + " restarted=" + restarted + "/" + cycles);
            Check("I.backlog-drains-within-bound", drained && missing == 0,
                "drained=" + drained + " drainMs=" + Fmt(drainMs) + " missing=" + missing
                    + " (bound " + Fmt(QueueDrainBoundMs) + "ms)");
            Check("I.every-item-accounted-for", completed + faulted == backlog,
                "completed=" + completed + " errored=" + errored + " faulted=" + faulted + " total=" + backlog);
            Check("I.poisoned-item-errored", errored > 0,
                "errored items=" + errored + " (the poisoned item returned Error: text, not a silent loss)");
            Check("I.queue-removed", WaitUntil(() => !NatsUdfAsync.HasQueueForTests(queueKey), 5000) >= 0,
                "the idle queue entry was removed");
            Check("I.heartbeat", Proc.MaxGapMs < StuckBoundMs && Watchdog.MaxIterMs < StuckBoundMs,
                "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms watchdog max=" + Fmt(Watchdog.MaxIterMs) + "ms");
            NoteWorst("queueDrainMs", drainMs);
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            Metric("I", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("I");
        }

        // ------------------- J: starvation combined with a server fault

        /// <summary>
        /// The starvation phase combined with a server restart: the pool is
        /// exhausted, an async backlog is enqueued and the server is killed in the
        /// same window; the watchdog must never trip (no stall past 5 s) and the
        /// queue must still drain once the pool and the server recover.
        /// Destructive (skipped with a reason).
        /// </summary>
        private static void PhaseJ()
        {
            long start = PhaseBanner("J", "ThreadPool starvation combined with a server restart");

            if (!CanRestartServer("J"))
            {
                Metric("J", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("J");
                return;
            }

            int port = ResolvePort(Host);
            if (!EnsureServerIdentity(port))
            {
                Console.WriteLine("[SKIP] phase=J sub=server-restart reason=could not resolve the nats-server process on port " + port);
                Metric("J", "durationSec", MsSince(start) / 1000.0);
                DumpLatencies("J");
                return;
            }
            string exePath = ServerExePath;
            string commandLine = ServerCommandLine;

            int listenersBefore = S.ListenerCount;
            int subscriptionsBefore = S.SubscriptionCount;

            var release = StarvePool(out int minW, out int minIo, out int maxW, out int maxIo, out bool setMin, out bool setMax, 500);
            Check("J.pool-guard-applied", setMin && setMax, "SetMin(4,4)=" + setMin + " SetMax(4,4)=" + setMax);
            if (!(setMin && setMax))
            {
                RestorePool(minW, minIo, maxW, maxIo, release);
                return;
            }

            int attackMs = Quick ? 2500 : 6000;
            int backlog = Quick ? 40 : 100;
            string queueKey = Host + "|j";
            var drained = new CountdownEvent(backlog);
            for (int i = 0; i < backlog; i++)
            {
                Task<object> ignored = NatsUdfAsync.Enqueue(queueKey, () => { drained.Signal(); return (object)null; });
                GC.KeepAlive(ignored);
            }

            long prevDeliveries = Delivery.Count;

            // The fault lands while the pool is exhausted.
            bool killed = KillProcess(CurrentServerPid(port));
            ServerPid = 0;
            Check("J.server-stopped", killed, "the nats-server process was stopped during the starvation window");
            WaitUntil(() => !ProbeNats(port, 500), ResumeBoundMs / 2);

            Thread.Sleep(attackMs);
            int pendingDuringAttack = drained.CurrentCount;

            // Release the pool, then bring the server back.
            RestorePool(minW, minIo, maxW, maxIo, release);
            long up = RestartServerBounded(exePath, commandLine, port, ResumeBoundMs);
            Check("J.server-restarted", up >= 0, "the server answered again after " + Fmt(up) + "ms");

            long drainStart = Stopwatch.GetTimestamp();
            bool drainedOk = drained.Wait((int)QueueDrainBoundMs);
            double drainMs = MsSince(drainStart);
            double resumeMs = Delivery.ResumeAfter(prevDeliveries, ResumeBoundMs);

            Metric("J", "pendingDuringAttack", pendingDuringAttack);
            Metric("J", "queueDrainMs", drainMs);
            Metric("J", "deliveryResumeMs", resumeMs);
            Metric("J", "procHeartbeatMs", Proc.MaxGapMs);
            Metric("J", "watchdogMaxMs", Watchdog.MaxIterMs);
            Check("J.backlog-pending-under-starvation", pendingDuringAttack > 0,
                "queued items stuck while starved=" + pendingDuringAttack + "/" + backlog);
            Check("J.queue-drains-after-recovery", drainedOk,
                "drainMs=" + Fmt(drainMs) + " (bound " + Fmt(QueueDrainBoundMs) + "ms)");
            Check("J.delivery-resumes-within-bound", resumeMs >= 0 && resumeMs <= ResumeBoundMs,
                "resumeMs=" + Fmt(resumeMs) + " (bound " + Fmt(ResumeBoundMs) + "ms)");
            Check("J.watchdog-never-tripped", Watchdog.MaxIterMs < StuckBoundMs && Proc.MaxGapMs < StuckBoundMs,
                "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms watchdog max=" + Fmt(Watchdog.MaxIterMs)
                    + "ms (bound " + Fmt(StuckBoundMs) + "ms)");
            Check("J.no-duplicate-listener",
                WaitUntil(() => S.ListenerCount == listenersBefore && S.SubscriptionCount == subscriptionsBefore, 8000) >= 0,
                "ListenerCount=" + S.ListenerCount + " SubscriptionCount=" + S.SubscriptionCount);
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            NoteWorst("watchdogMs", Watchdog.MaxIterMs);
            NoteWorst("queueDrainMs", drainMs);
            NoteWorst("deliveryResumeMs", resumeMs);
            Metric("J", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("J");
        }

        // ------------------- K: graceful shutdown under load

        /// <summary>
        /// The "Excel closed while a cell was live" case: call
        /// <see cref="NatsRuntime.Shutdown"/> while traffic is in flight on the
        /// shared runtime and assert it returns within the documented budget, the
        /// tombstone refuses to resurrect the managers, and
        /// <see cref="NatsRuntime.ResetAfterAddInReload"/> plus a fresh use proves
        /// recovery.
        /// </summary>
        private static void PhaseK()
        {
            long start = PhaseBanner("K", "graceful shutdown under load (tombstone + reload)");

            // The shared runtime is the production instance the add-in uses; it is
            // independent of the standalone managers the other phases drive.
            var sharedSubs = NatsRuntime.Subscriptions;
            var sharedConns = NatsRuntime.Connections;
            try
            {
                var warm = sharedConns.GetConnection(Host);
                Observe(warm.ConnectAsync().AsTask());
                NatsConnectionManager.WaitBounded(warm.PingAsync().AsTask(), 5000);
            }
            catch (Exception ex)
            {
                Metric("K", "warmupError", ShortType(ex));
            }

            long deliveries = 0;
            long publishErrors = 0;
            bool stop = false;
            IDisposable token = null;
            Thread publisher = null;
            try
            {
                string subject = "test.natsexcel.k.load." + Guid.NewGuid().ToString("N");
                token = sharedSubs.Subscribe(Host, subject, null, _ => Interlocked.Increment(ref deliveries), "K");

                publisher = new Thread(() =>
                {
                    long n = 0;
                    while (!Volatile.Read(ref stop))
                    {
                        try
                        {
                            var connection = sharedConns.GetConnection(Host);
                            Observe(connection.PublishAsync(subject, "k-" + Interlocked.Increment(ref n)).AsTask());
                        }
                        catch
                        {
                            Interlocked.Increment(ref publishErrors);
                        }
                        Thread.Sleep(5);
                    }
                });
                publisher.IsBackground = true;
                publisher.Name = "K-publisher";
                publisher.Start();
                // Soak on the observed in-flight traffic, not a fixed window: the
                // shutdown below must land while a delivery has really flowed.
                // A silent publisher still consumes the whole window (the maximum
                // is unchanged) and then the shutdown bound assertion fails.
                WaitUntil(() => Interlocked.Read(ref deliveries) > 0, Quick ? 500 : 1500);

                // Shut the shared runtime down WHILE traffic is in flight.
                long t0 = Stopwatch.GetTimestamp();
                NatsRuntime.Shutdown();
                double shutdownMs = MsSince(t0);
                Metric("K", "shutdownMs", shutdownMs);
                Check("K.shutdown-bounded", shutdownMs <= ShutdownBoundMs,
                    "NatsRuntime.Shutdown returned in " + Fmt(shutdownMs) + "ms (bound " + Fmt(ShutdownBoundMs) + "ms)");
                NoteWorst("shutdownMs", shutdownMs);

                Volatile.Write(ref stop, true);
                publisher.Join(5000);

                // The managers are NOT resurrected by a later access (the tombstone).
                bool connectionsThrows = ThrowsShuttingDown(() => { var _ = NatsRuntime.Connections; });
                bool subscriptionsThrows = ThrowsShuttingDown(() => { var _ = NatsRuntime.Subscriptions; });
                Check("K.tombstone-blocks-resurrection", connectionsThrows && subscriptionsThrows,
                    "Connections throws=" + connectionsThrows + " Subscriptions throws=" + subscriptionsThrows);

                Check("K.shared-managers-shut-down", sharedConns.LiveConnectionCount() == 0,
                    "shared LiveConnectionCount=" + sharedConns.LiveConnectionCount());

                // A same-process add-in reload clears the tombstone; a fresh use
                // must work (Excel closed the add-in and reopened it).
                NatsRuntime.ResetAfterAddInReload();
                var freshSubs = NatsRuntime.Subscriptions;
                var freshConns = NatsRuntime.Connections;
                Check("K.runtime-recreated-after-reload",
                    freshSubs != null && freshConns != null
                        && !ReferenceEquals(freshSubs, sharedSubs) && !ReferenceEquals(freshConns, sharedConns),
                    "fresh managers were recreated after ResetAfterAddInReload");

                string freshSubject = "test.natsexcel.k.fresh." + Guid.NewGuid().ToString("N");
                int got = 0;
                IDisposable freshToken = freshSubs.Subscribe(Host, freshSubject, null, _ => Interlocked.Increment(ref got), "K");
                try
                {
                    var freshConn = freshConns.GetConnection(Host);
                    Observe(freshConn.ConnectAsync().AsTask());
                    bool delivered = false;
                    for (int i = 0; i < 60 && !delivered; i++)
                    {
                        Observe(freshConn.PublishAsync(freshSubject, "k-fresh-" + i).AsTask());
                        delivered = WaitUntil(() => Volatile.Read(ref got) > 0, 250) >= 0;
                    }
                    Check("K.fresh-use-after-reload", delivered,
                        "a fresh subscription/message round trip after the reload delivered=" + delivered);
                }
                finally
                {
                    try { freshToken.Dispose(); } catch { }
                }
            }
            finally
            {
                Volatile.Write(ref stop, true);
                try { if (publisher != null) publisher.Join(5000); } catch { }
                try { if (token != null) token.Dispose(); } catch { }
            }

            Metric("K", "loadDeliveries", deliveries);
            Metric("K", "loadPublishErrors", publishErrors);
            Metric("K", "procHeartbeatMs", Proc.MaxGapMs);
            Check("K.heartbeat", Proc.MaxGapMs < StuckBoundMs && Watchdog.MaxIterMs < StuckBoundMs,
                "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms watchdog max=" + Fmt(Watchdog.MaxIterMs) + "ms");
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            Metric("K", "durationSec", MsSince(start) / 1000.0);
            DumpLatencies("K");
        }

        // ------------------------------------------ E: resource stability

        private static void PhaseResourceStability()
        {
            Console.WriteLine();
            Console.WriteLine("[PHASE] id=E desc=resource stability t=" + Clock.Elapsed.TotalSeconds.ToString("0.0", Inv));

            long deliveries = Delivery.Count;
            int live = C.LiveConnectionCount();

            // Stop the delivery probe first: its own subscription is not a leak.
            try { Delivery.Dispose(); } catch { }
            foreach (string subject in TrackedSubjects)
            {
                try { NatsUdfCore.NatsUDFSubjectUnsubscribe(subject, Host); } catch { }
            }

            long settled = WaitUntil(() => S.ListenerCount == 0 && S.SubscriptionCount == 0, 15000);
            Metric("E", "listenerCount", S.ListenerCount);
            Metric("E", "subscriptionCount", S.SubscriptionCount);
            Check("E.no-listener-leak", settled >= 0,
                "ListenerCount=" + S.ListenerCount + " SubscriptionCount=" + S.SubscriptionCount);

            int liveRtd = C.LiveRtdConnectionCount();
            int liveUdf = C.LiveUdfConnectionCount();
            Metric("E", "liveConnections", live);
            Metric("E", "baselineConnections", BaselineConnections);
            Check("E.no-connection-leak", live == BaselineConnections,
                "LiveConnectionCount=" + live + " (baseline " + BaselineConnections + ")");
            Check("E.connection-families-consistent", live == liveRtd && live == liveUdf,
                "live=" + live + " rtd=" + liveRtd + " udf=" + liveUdf);

            // The runtime must still be healthy (never shut down by our teardown).
            bool runtimeHealthy = false;
            try { runtimeHealthy = NatsRuntime.Subscriptions != null; }
            catch (Exception ex) { Metric("E", "runtimeError", ShortType(ex)); }
            Check("E.runtime-healthy-after-phases", runtimeHealthy, "NatsRuntime.Subscriptions reachable");

            Check("E.delivery-was-alive", deliveries > 0, "delivery count=" + deliveries);
            Metric("E", "deliveryCount", deliveries);
            Metric("E", "procHeartbeatMs", Proc.MaxGapMs);
            Metric("E", "watchdogMaxMs", Watchdog.MaxIterMs);
            Check("E.heartbeat", Proc.MaxGapMs < StuckBoundMs && Watchdog.MaxIterMs < StuckBoundMs,
                "proc max gap=" + Fmt(Proc.MaxGapMs) + "ms watchdog max=" + Fmt(Watchdog.MaxIterMs) + "ms");
            NoteWorst("procHeartbeatMs", Proc.MaxGapMs);
            NoteWorst("watchdogMs", Watchdog.MaxIterMs);
        }

        // ------------------------------------------------------------------
        // Server-managed phase helpers (F, G, H, I, J)
        // ------------------------------------------------------------------

        /// <summary>
        /// Whether a destructive stop/restart phase may run: a loopback server, no
        /// <c>--skip-restart</c>, and a live NATS listener on the port. Prints the
        /// reason (the "--skip-restart" case is already covered by the run banner).
        /// </summary>
        private static bool CanRestartServer(string phase)
        {
            if (!IsLocalEndpoint())
            {
                Console.WriteLine("[SKIP] phase=" + phase + " sub=server-restart reason=non-local host " + Host
                    + " (the destructive stop/restart only runs against a loopback server)");
                return false;
            }
            if (SkipRestart)
            {
                Console.WriteLine("[SKIP] phase=" + phase + " sub=server-restart reason=--skip-restart");
                return false;
            }
            int port = ResolvePort(Host);
            if (!ProbeNats(port, 3000))
            {
                Console.WriteLine("[SKIP] phase=" + phase + " sub=server-restart reason=no NATS listener on port " + port);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Resolves and caches the local server's pid, executable and command line
        /// once (phase C), so the later stop/restart phases do not each pay a
        /// PowerShell + CIM round trip (~3 s). False when it cannot be resolved.
        /// </summary>
        private static bool EnsureServerIdentity(int port)
        {
            if (!string.IsNullOrEmpty(ServerExePath) && !string.IsNullOrEmpty(ServerCommandLine))
                return true;
            if (!FindServerProcess(port, out string pidText, out string exePath, out string commandLine))
                return false;
            ServerExePath = exePath;
            ServerCommandLine = commandLine;
            ServerPid = int.Parse(pidText, Inv);
            return true;
        }

        /// <summary>The live server pid, re-resolved after an unexpected loss (the
        /// cached pid is stale after a restart) or the cached one on the fast path.</summary>
        private static int CurrentServerPid(int port)
        {
            if (ServerPid > 0 && ProcessAlive(ServerPid))
                return ServerPid;
            if (FindServerProcess(port, out string pidText, out string exePath, out string commandLine))
            {
                ServerExePath = exePath;
                ServerCommandLine = commandLine;
                ServerPid = int.Parse(pidText, Inv);
            }
            return ServerPid;
        }

        private static bool ProcessAlive(int pid)
        {
            try
            {
                using (var p = Process.GetProcessById(pid))
                    return !p.HasExited;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Restarts the local server and waits (bounded) for it to answer again,
        /// retrying the launch: a server killed moments earlier can still hold the
        /// port or its JetStream store lock, so a single launch attempt is not
        /// reliable. Returns the elapsed ms on success, or -1 within the bound.
        /// </summary>
        private static long RestartServerBounded(string exePath, string commandLine, int port, double boundMs)
        {
            long start = Stopwatch.GetTimestamp();
            int attempt = 0;
            while (MsSince(start) < boundMs)
            {
                attempt++;
                if (attempt > 1)
                    Thread.Sleep(500); // let the killed predecessor release the port/store lock

                StartProcess(exePath, commandLine);
                double remaining = boundMs - MsSince(start);
                long up = WaitUntil(() => ProbeNats(port, 1000), Math.Min(5000.0, remaining));
                if (up >= 0)
                {
                    if (attempt > 1)
                        Console.WriteLine("[METRIC] name=serverRestartAttempts value=" + attempt);
                    return (long)MsSince(start);
                }
            }
            Console.WriteLine("[METRIC] name=serverRestartAttempts value=" + attempt + " (gave up within " + Fmt(boundMs) + "ms)");
            return -1;
        }

        /// <summary>
        /// The hermetic config for the server-managed phases: the endpoint under
        /// test, JetStream on, no dedup/conflation surprises, and the write mode
        /// the phase needs (request phases need <c>sync</c> so a request really
        /// waits for a reply; the JetStream phase needs the admin gate open).
        /// </summary>
        private static AppConfig.Root BuildPhaseConfig(string url, bool allowAdminWrites, string syncWrite)
        {
            var cfg = new AppConfig.Root
            {
                Default = "local",
                TimeoutMs = 1000,
                JetStreamTimeoutMs = 8000,
                JetStreamEnabled = true,
                AllowAdminWrites = allowAdminWrites,
                SyncWrite = syncWrite,
                PublishDedupCacheSize = 10000,
                SkipRepeatedMessages = false,
                CoalesceRealtimeUpdates = true,
            };
            cfg.Servers["local"] = url;
            var cluster = new AppConfig.Cluster { Label = "local" };
            cluster.Urls.Add(url);
            cfg.Clusters["local"] = cluster;
            return AppConfig.Sanitize(cfg);
        }

        /// <summary>
        /// True for a request result that is a bounded failure: the frozen
        /// no-responders/timeout markers, or any "Error:" text (a client-side
        /// timeout or connection error). A real reply is not a bounded error.
        /// </summary>
        private static bool IsBoundedRequestError(object result)
        {
            string text = CellText(result);
            if (text == NatsUdfCommon.NoResponders || text == NatsUdfCommon.Timeout)
                return true;
            if (text.StartsWith(NatsUdfCommon.ErrorPrefix, StringComparison.Ordinal))
                return true;
            return IsTimeoutText(text);
        }

        private static string CellText(object value)
        {
            if (value == null) return "";
            if (value is string text) return text;
            return Convert.ToString(value, Inv) ?? "";
        }

        /// <summary>
        /// Reads a completed task's result on a bounded path. The task is known
        /// complete; the access is a direct read under the same bounded-wait
        /// discipline the harness uses everywhere (VSTHRD002 suppressed here and
        /// only here, mirroring <see cref="NatsConnectionManager.WaitBounded"/>).
        /// </summary>
#pragma warning disable VSTHRD002 // deliberate read of a task this phase already bounded-waited on.
        private static T BoundedResult<T>(Task<T> task) { return task.Result; }
#pragma warning restore VSTHRD002

        /// <summary>The numeric value of a cell result, or NaN when it is not a number.</summary>
        private static double AsDouble(object value)
        {
            if (value is double d) return d;
            if (value is float f) return f;
            if (value is int i) return i;
            if (value is long l) return l;
            if (value is string s && double.TryParse(s, NumberStyles.Float, Inv, out double parsed)) return parsed;
            return double.NaN;
        }

        /// <summary>
        /// Reads a long field from the compact JetStream JSON
        /// (<c>NatsResultFormatter.FormatStreamInfoJson</c>), e.g.
        /// <c>"messages":&lt;long&gt;</c>. Returns -1 when the field is absent.
        /// </summary>
        private static long JsonLong(string json, string field)
        {
            if (string.IsNullOrEmpty(json)) return -1;
            string needle = "\"" + field + "\":";
            int at = json.IndexOf(needle, StringComparison.Ordinal);
            if (at < 0) return -1;
            int start = at + needle.Length;
            int end = start;
            while (end < json.Length
                   && (char.IsDigit(json[end]) || json[end] == '-' || json[end] == '+' || json[end] == '.'))
                end++;
            return long.TryParse(json.Substring(start, end - start), NumberStyles.Integer, Inv, out long value)
                ? value
                : -1;
        }

        /// <summary>An NxM matrix (or a scalar) as text rows; a non-matrix yields one row.</summary>
        private static string[][] MatrixRows(object matrix)
        {
            if (matrix is object[,] grid)
            {
                int rows = grid.GetLength(0);
                int cols = grid.GetLength(1);
                var result = new string[rows][];
                for (int r = 0; r < rows; r++)
                {
                    result[r] = new string[cols];
                    for (int c = 0; c < cols; c++)
                        result[r][c] = CellText(grid[r, c]);
                }
                return result;
            }
            if (matrix is string[][] jagged) return jagged;
            return new[] { new[] { CellText(matrix) } };
        }

        /// <summary>True when an access throws the shutdown-tombstone error.</summary>
        private static bool ThrowsShuttingDown(Action access)
        {
            try
            {
                access();
                return false;
            }
            catch (InvalidOperationException ex)
            {
                return ex.Message.IndexOf("shutting down", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>A raw NATS client outside the code under test (a test double).</summary>
        private static NatsConnection NewRawConnection(string url)
        {
            var opts = NatsOpts.Default with
            {
                Url = url,
                Name = "NATSExcel LivenessTests",
                ConnectTimeout = TimeSpan.FromSeconds(5),
            };
            var connection = new NatsConnection(opts);
            NatsConnectionManager.WaitBounded(connection.ConnectAsync().AsTask(), 6000);
            return connection;
        }

        private static void DisposeConnection(NatsConnection connection)
        {
            if (connection == null) return;
            try { NatsConnectionManager.WaitBounded(connection.DisposeAsync().AsTask(), 5000); } catch { }
        }

        /// <summary>
        /// A real request/reply responder on a raw client connection, so a request
        /// success is a genuine round trip rather than a marker. Mirrors the smoke
        /// layer's responder (Core <c>SubscribeCoreAsync</c> + reply to
        /// <c>ReplyTo</c>).
        /// </summary>
        private sealed class RawResponder
        {
            private readonly CancellationTokenSource _cts = new CancellationTokenSource();
            private readonly TaskCompletionSource<bool> _ready =
                new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            public RawResponder(NatsConnection connection, string subject, Func<string, string> reply)
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
                                string text = msg.Data == null ? "" : Encoding.UTF8.GetString(msg.Data);
                                if (reply != null && !string.IsNullOrEmpty(msg.ReplyTo))
                                    await connection.PublishAsync(msg.ReplyTo, reply(text)).ConfigureAwait(false);
                            }
                        }
                    }
                    catch (OperationCanceledException) { }
                    catch (Exception ex) { Console.WriteLine("  [responder] " + ex.GetType().Name); }
                });
            }

            /// <summary>Blocks (bounded) until the wire SUB is established.</summary>
            public bool WaitReady(int timeoutMs)
            {
                long start = Stopwatch.GetTimestamp();
                while (Stopwatch.GetTimestamp() - start < MsToTicks(timeoutMs))
                {
                    if (_ready.Task.IsCompleted) return BoundedResult(_ready.Task);
                    Thread.Sleep(20);
                }
                return false;
            }

            public void Dispose()
            {
                try { _cts.Cancel(); } catch { }
            }
        }

        // ------------------------------------------------------------------
        // Server process management (phase C)
        // ------------------------------------------------------------------

        /// <summary>Find the process listening on the port and its executable and
        /// command line, via a temp PowerShell script (no WMI assembly dependency
        /// and no fragile inline quoting).</summary>
        private static bool FindServerProcess(int port, out string pid, out string exePath, out string commandLine)
        {
            pid = null; exePath = null; commandLine = null;
            string script =
                "$ErrorActionPreference='SilentlyContinue'\r\n" +
                "$c = Get-NetTCPConnection -LocalPort " + port + " -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1\r\n" +
                "if (-not $c) { return }\r\n" +
                "$p = Get-CimInstance Win32_Process -Filter (\"ProcessId=\" + $c.OwningProcess) -ErrorAction SilentlyContinue\r\n" +
                "if (-not $p) { return }\r\n" +
                "Write-Output ($c.OwningProcess.ToString() + [char]9 + [string]$p.ExecutablePath + [char]9 + [string]$p.CommandLine)\r\n";

            string scriptPath = null;
            try
            {
                scriptPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                    "natsexcel-liveness-find-" + Guid.NewGuid().ToString("N") + ".ps1");
                System.IO.File.WriteAllText(scriptPath, script, new UTF8Encoding(false));
                string output = RunProcess("powershell.exe",
                    "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"" + scriptPath + "\"", 20000,
                    out _);
                if (string.IsNullOrEmpty(output)) return false;
                foreach (string rawLine in output.Split('\n'))
                {
                    string line = rawLine.Trim('\r', ' ');
                    if (line.Length == 0) continue;
                    string[] parts = line.Split('\t');
                    if (parts.Length < 3) continue;
                    pid = parts[0].Trim();
                    exePath = parts[1].Trim();
                    commandLine = parts[2].Trim();
                    return pid.Length > 0 && exePath.Length > 0;
                }
                return false;
            }
            finally
            {
                if (scriptPath != null) { try { System.IO.File.Delete(scriptPath); } catch { } }
            }
        }

        private static string RunProcess(string fileName, string arguments, int timeoutMs, out int exitCode)
        {
            exitCode = -1;
            try
            {
                var psi = new ProcessStartInfo(fileName, arguments)
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                };
                using (var p = Process.Start(psi))
                {
                    string stdout = p.StandardOutput.ReadToEnd();
                    p.StandardError.ReadToEnd();
                    if (!p.WaitForExit(timeoutMs)) { try { p.Kill(); } catch { } return null; }
                    exitCode = p.ExitCode;
                    return stdout;
                }
            }
            catch
            {
                return null;
            }
        }

        private static bool KillProcess(int pid)
        {
            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    p.Kill();
                    return p.WaitForExit(15000);
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>Relaunch the server with the exact executable and command line
        /// resolved before it was stopped (never a hand-written command). Launched
        /// detached so the child never inherits our std handles (a child holding
        /// the stdout pipe open would keep a piped run from ever terminating) and
        /// its log output never interleaves with the test output. Returns the new
        /// pid, or 0 when the launch failed; the caller records it so the next
        /// kill needs no PowerShell/CIM re-resolution.</summary>
        private static int StartProcess(string exePath, string commandLine)
        {
            try
            {
                string args = StripExecutable(commandLine);
                var psi = new ProcessStartInfo(exePath, args)
                {
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = System.IO.Path.GetDirectoryName(exePath),
                };
                int pid;
                using (var p = Process.Start(psi))
                {
                    if (p == null) return 0;
                    pid = p.Id;
                }
                ServerPid = pid;
                return pid;
            }
            catch
            {
                return 0;
            }
        }

        private static string StripExecutable(string commandLine)
        {
            if (string.IsNullOrWhiteSpace(commandLine)) return "-c";
            string line = commandLine.Trim();
            if (line.StartsWith("\"", StringComparison.Ordinal))
            {
                int end = line.IndexOf('"', 1);
                if (end < 0) return "-c";
                return line.Substring(end + 1).Trim();
            }
            int space = line.IndexOf(' ');
            if (space < 0) return "-c";
            return line.Substring(space + 1).Trim();
        }

        /// <summary>Raw NATS handshake probe (INFO + PING/PONG) outside the code
        /// under test, so fault injection and measurement never depend on it.</summary>
        private static bool ProbeNats(int port, int timeoutMs)
        {
            try
            {
                using (var client = new TcpClient())
                {
                    IAsyncResult ar = client.BeginConnect("127.0.0.1", port, null, null);
                    if (!ar.AsyncWaitHandle.WaitOne(timeoutMs)) return false;
                    client.EndConnect(ar);
                    using (var stream = client.GetStream())
                    {
                        stream.ReadTimeout = timeoutMs;
                        stream.WriteTimeout = timeoutMs;
                        var buffer = new byte[4096];
                        int n = stream.Read(buffer, 0, buffer.Length);
                        if (n <= 0) return false;
                        byte[] ping = Encoding.ASCII.GetBytes("PING\r\n");
                        stream.Write(ping, 0, ping.Length);
                        Thread.Sleep(120);
                        int n2 = stream.Read(buffer, 0, buffer.Length);
                        string reply = Encoding.ASCII.GetString(buffer, 0, Math.Max(0, n2));
                        return reply.IndexOf("PONG", StringComparison.Ordinal) >= 0;
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        private static bool IsLocalEndpoint()
        {
            string host = ExtractHost(Host);
            if (string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase)) return true;
            return host == "127.0.0.1" || host == "::1";
        }

        private static string ExtractHost(string url)
        {
            try
            {
                string authority = url;
                int scheme = authority.IndexOf("://", StringComparison.Ordinal);
                if (scheme >= 0) authority = authority.Substring(scheme + 3);
                int comma = authority.IndexOf(',');
                if (comma >= 0) authority = authority.Substring(0, comma);
                int slash = authority.IndexOf('/');
                if (slash >= 0) authority = authority.Substring(0, slash);
                int at = authority.LastIndexOf('@');
                if (at >= 0) authority = authority.Substring(at + 1);
                if (authority.StartsWith("[", StringComparison.Ordinal))
                {
                    int close = authority.IndexOf(']');
                    return close > 0 ? authority.Substring(1, close - 1) : authority;
                }
                int colon = authority.LastIndexOf(':');
                return colon >= 0 ? authority.Substring(0, colon) : authority;
            }
            catch
            {
                return url;
            }
        }

        private static int ResolvePort(string url)
        {
            try
            {
                string authority = url;
                int scheme = authority.IndexOf("://", StringComparison.Ordinal);
                if (scheme >= 0) authority = authority.Substring(scheme + 3);
                int comma = authority.IndexOf(',');
                if (comma >= 0) authority = authority.Substring(0, comma);
                int slash = authority.IndexOf('/');
                if (slash >= 0) authority = authority.Substring(0, slash);
                int at = authority.LastIndexOf('@');
                if (at >= 0) authority = authority.Substring(at + 1);
                if (authority.StartsWith("[", StringComparison.Ordinal))
                {
                    int close = authority.IndexOf(']');
                    if (close >= 0 && close + 2 < authority.Length)
                        return int.Parse(authority.Substring(close + 2), Inv);
                    return 4222;
                }
                int colon = authority.LastIndexOf(':');
                if (colon >= 0 && colon + 1 < authority.Length
                    && int.TryParse(authority.Substring(colon + 1), NumberStyles.Integer, Inv, out int port))
                    return port;
                return 4222;
            }
            catch
            {
                return 4222;
            }
        }

        private static string DescribeAntivirus()
        {
            try
            {
                string script = FindRepoFile("test\\TestSupport\\Get-AntivirusState.ps1");
                if (script == null) return "(unavailable: helper not found)";
                var psi = new ProcessStartInfo("powershell.exe",
                    "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"" + script + "\"")
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                };
                using (var p = Process.Start(psi))
                {
                    string stdout = p.StandardOutput.ReadToEnd().Trim();
                    p.StandardError.ReadToEnd();
                    if (!p.WaitForExit(15000)) { try { p.Kill(); } catch { } return "(unavailable: helper timed out)"; }
                    return string.IsNullOrEmpty(stdout) ? "(unknown)" : stdout;
                }
            }
            catch (Exception ex)
            {
                return "(unavailable: " + ShortType(ex) + ")";
            }
        }

        /// <summary>Find a repository-relative file by walking up from the assembly.</summary>
        private static string FindRepoFile(string relative)
        {
            try
            {
                var dir = new System.IO.DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
                for (int i = 0; i < 8 && dir != null; i++)
                {
                    string candidate = System.IO.Path.Combine(dir.FullName, relative);
                    if (System.IO.File.Exists(candidate)) return candidate;
                    dir = dir.Parent;
                }
            }
            catch { }
            return null;
        }

        private static long PhaseBanner(string id, string desc)
        {
            Console.WriteLine();
            Console.WriteLine("[PHASE] id=" + id + " desc=" + desc + " t=" + Clock.Elapsed.TotalSeconds.ToString("0.0", Inv));
            foreach (var kv in Lats) kv.Value.Reset();
            if (Proc != null) Proc.Reset();
            if (Delivery != null) Delivery.ResetMax();
            if (Watchdog != null) Watchdog.ResetMax();
            return Stopwatch.GetTimestamp();
        }

        private static void Metric(string phase, string name, object value)
        {
            string text = value is double d ? Fmt(d) : Convert.ToString(value, Inv);
            Console.WriteLine("[METRIC] phase=" + phase + " name=" + name + " value=" + text);
        }

        private static void Check(string name, bool ok, string detail)
        {
            Console.WriteLine("[ASSERT] name=" + name + " passed=" + ok + " detail=" + detail);
            if (!ok) Failures.Enqueue(name + " :: " + detail);
        }

        private static void DumpLatencies(string phase)
        {
            foreach (var kv in Lats.OrderBy(k => k.Key))
            {
                if (kv.Value.Count == 0) continue;
                Console.WriteLine("[METRIC] phase=" + phase + " lat=" + kv.Key + " maxMs=" + Fmt(kv.Value.MaxMs) + " n=" + kv.Value.Count);
            }
        }

        private static void NoteWorst(string name, double value)
        {
            if (value < 0) return;
            lock (WorstGate)
            {
                if (!Worst.TryGetValue(name, out double current) || value > current) Worst[name] = value;
            }
        }

        private static double GetWorst(string name)
        {
            lock (WorstGate)
            {
                return Worst.TryGetValue(name, out double value) ? value : -1;
            }
        }

        private static Latency GetLat(string name) { return Lats.GetOrAdd(name, _ => new Latency()); }

        private static void NoopMessage(NatsMessage message) { }

        /// <summary>A listener that must never throw, so the "no exception escapes
        /// a callback" property is tested by the channel it is delivered on.</summary>
        private static void ListenNoop(NatsMessage message) { GC.KeepAlive(message); }

        private static bool IsError(object result)
        {
            var text = result as string;
            return text != null && text.StartsWith(NatsUdfCommon.ErrorPrefix, StringComparison.Ordinal);
        }

        /// <summary>
        /// True for a transient timeout-family result. Under the intentional
        /// phase-B flood a request/reply to a responder-less subject may return
        /// either the frozen "TIMEOUT"/"No Responders" marker or the client's
        /// generic "No reply received" (a request that timed out on its inbox
        /// rather than surfacing the 503 sentinel). All three are bounded, not a
        /// hang; the watchdog is what proves liveness, and the E2E layer is what
        /// pins the exact sentinel, so they are counted as timeouts here.
        /// </summary>
        private static bool IsTimeoutText(object result)
        {
            var text = result as string;
            if (text == null) return false;
            return text.IndexOf("timeout", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("did not complete", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("no reply received", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string ShortType(Exception ex) { return ex.GetType().Name + ": " + ex.Message; }

        private static string Shorten(string text, int max)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= max) return text ?? "";
            return text.Substring(0, max) + "...";
        }

        /// <summary>Fire-and-log: observes a task so a late failure is never an
        /// unobserved task exception.</summary>
        private static void Observe(Task task)
        {
            if (task == null) return;
            Task observation = task.ContinueWith(
                t => { var ignored = t.Exception; },
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
            GC.KeepAlive(observation);
        }

        private static void UpdateMax(ref int target, int value)
        {
            int current;
            while ((current = Volatile.Read(ref target)) < value)
            {
                if (Interlocked.CompareExchange(ref target, value, current) == current) return;
            }
        }

        private static void RecordMax(ref long target, long value)
        {
            long current;
            while ((current = Interlocked.Read(ref target)) < value)
            {
                if (Interlocked.CompareExchange(ref target, value, current) == current) return;
            }
        }

        private static long WaitUntil(Func<bool> probe, double timeoutMs)
        {
            long start = Stopwatch.GetTimestamp();
            long deadline = start + MsToTicks(timeoutMs);
            while (Stopwatch.GetTimestamp() < deadline)
            {
                try { if (probe()) return (long)MsSince(start); }
                catch { }
                Thread.Sleep(20);
            }
            return -1;
        }

        private static int CountCompleted(RecordingObserver[] observers)
        {
            int n = 0;
            for (int i = 0; i < observers.Length; i++) if (observers[i].CompletedCount == 1) n++;
            return n;
        }

        private static void GetPool(out int minW, out int minIo, out int maxW, out int maxIo)
        {
            ThreadPool.GetMinThreads(out minW, out minIo);
            ThreadPool.GetMaxThreads(out maxW, out maxIo);
        }

        /// <summary>
        /// Starve the pool: clamp min/max to (4,4) and block enough queued sleepers
        /// to exhaust every worker thread the pool has or may create (more than
        /// <c>2 * ProcessorCount + 8</c>, and never fewer than 32, so a many-core
        /// box is exhausted too). Released by <see cref="RestorePool"/>.
        /// </summary>
        private static ManualResetEventSlim StarvePool(
            out int minW, out int minIo, out int maxW, out int maxIo,
            out bool setMin, out bool setMax, int settleMs)
        {
            GetPool(out minW, out minIo, out maxW, out maxIo);
            setMin = ThreadPool.SetMinThreads(4, 4);
            setMax = ThreadPool.SetMaxThreads(4, 4);
            var release = new ManualResetEventSlim(false);
            int sleepers = Math.Max(32, Environment.ProcessorCount * 2 + 8);
            for (int i = 0; i < sleepers; i++)
                ThreadPool.QueueUserWorkItem(_ => release.Wait(120000));
            // Settle on the real condition the window exists for - the pool has no
            // free worker (every worker thread is blocked on a sleeper) - up to the
            // SAME maximum. A healthy box reaches it almost immediately; a box where
            // the clamp did not take still waits the whole window (the fallback
            // preserves the old behaviour exactly).
            WaitUntil(() =>
            {
                ThreadPool.GetAvailableThreads(out int availableWorkers, out int availableIo);
                GC.KeepAlive(availableIo);
                return availableWorkers == 0;
            }, settleMs);
            return release;
        }

        private static void RestorePool(int minW, int minIo, int maxW, int maxIo, ManualResetEventSlim release)
        {
            release.Set();
            ThreadPool.SetMaxThreads(maxW, maxIo);
            ThreadPool.SetMinThreads(minW, minIo);
        }

        private static double TicksToMs(long ticks) { return ticks * 1000.0 / Stopwatch.Frequency; }
        private static long MsToTicks(double ms) { return (long)(ms * Stopwatch.Frequency / 1000.0); }
        private static double MsSince(long startTicks) { return TicksToMs(Stopwatch.GetTimestamp() - startTicks); }

        private static string Fmt(double ms)
        {
            return ms < 0 ? "timeout" : ms.ToString("0.0", Inv);
        }

        // ------------------------------------------------------------------
        // Monitors
        // ------------------------------------------------------------------

        private sealed class Latency
        {
            private long _count;
            private long _maxTicks;

            public void Record(long ticks)
            {
                Interlocked.Increment(ref _count);
                RecordMax(ref _maxTicks, ticks);
            }

            public void Reset()
            {
                Interlocked.Exchange(ref _count, 0);
                Interlocked.Exchange(ref _maxTicks, 0);
            }

            public long Count { get { return Interlocked.Read(ref _count); } }
            public double MaxMs { get { return TicksToMs(Interlocked.Read(ref _maxTicks)); } }
        }

        private sealed class Worker
        {
            private readonly Thread[] _threads;
            private volatile bool _stop;
            private long _ops;
            private long _errors;
            private string _firstError;

            public readonly string Name;

            public Worker(string name, int count, Action<int, long> body)
            {
                Name = name;
                _threads = new Thread[count];
                for (int i = 0; i < count; i++)
                {
                    int id = i;
                    var t = new Thread(() =>
                    {
                        long n = 0;
                        while (!_stop)
                        {
                            try { body(id, n); Interlocked.Increment(ref _ops); }
                            catch (Exception ex)
                            {
                                Interlocked.Increment(ref _errors);
                                if (Volatile.Read(ref _firstError) == null)
                                    Volatile.Write(ref _firstError, ShortType(ex));
                            }
                            n++;
                        }
                    });
                    t.IsBackground = true;
                    t.Name = name + "-" + i;
                    _threads[i] = t;
                }
            }

            public long Ops { get { return Interlocked.Read(ref _ops); } }
            public long Errors { get { return Interlocked.Read(ref _errors); } }
            public string FirstError { get { return Volatile.Read(ref _firstError); } }
            public void Start() { foreach (var t in _threads) t.Start(); }

            public void StopAndJoin(int timeoutMs = 15000)
            {
                _stop = true;
                foreach (var t in _threads) t.Join(timeoutMs);
            }
        }

        /// <summary>
        /// Dedicated-thread heartbeat: touches <see cref="LastBeatTicks"/> every
        /// interval and records the longest gap. A long gap means the process
        /// itself stalled (not the pool).
        /// </summary>
        private sealed class ProcessHeartbeat
        {
            private readonly string _name;
            private readonly int _intervalMs;
            private Thread _thread;
            private volatile bool _stop;
            private long _maxGapTicks;
            private long _ticks;
            private long _lastBeatTicks;

            public ProcessHeartbeat(string name, int intervalMs)
            {
                _name = name;
                _intervalMs = intervalMs;
                Interlocked.Exchange(ref _lastBeatTicks, Stopwatch.GetTimestamp());
            }

            public void Start()
            {
                _thread = new Thread(() =>
                {
                    long last = Stopwatch.GetTimestamp();
                    while (!_stop)
                    {
                        Thread.Sleep(_intervalMs);
                        long now = Stopwatch.GetTimestamp();
                        Interlocked.Exchange(ref _lastBeatTicks, now);
                        RecordMax(ref _maxGapTicks, now - last);
                        last = now;
                        Interlocked.Increment(ref _ticks);
                    }
                });
                _thread.IsBackground = true;
                _thread.Name = _name;
                _thread.Start();
            }

            public void Stop() { _stop = true; }
            public void Reset() { Interlocked.Exchange(ref _maxGapTicks, 0); }
            public double MaxGapMs { get { return TicksToMs(Interlocked.Read(ref _maxGapTicks)); } }
            public long Ticks { get { return Interlocked.Read(ref _ticks); } }
            public long LastBeatTicks { get { return Interlocked.Read(ref _lastBeatTicks); } }
        }

        /// <summary>
        /// ThreadPool heartbeat: a self-rescheduling continuation. Its gap
        /// measures how long the pool could not run work (the starvation probe).
        /// </summary>
        private sealed class PoolHeartbeat
        {
            private volatile bool _stop;
            private long _maxGapTicks;
            private long _lastTicks;
            private long _ticks;

            public void Start()
            {
                Interlocked.Exchange(ref _lastTicks, Stopwatch.GetTimestamp());
                Schedule();
            }

            public void Stop() { _stop = true; }
            public void Reset() { Interlocked.Exchange(ref _maxGapTicks, 0); }
            public double MaxGapMs { get { return TicksToMs(Interlocked.Read(ref _maxGapTicks)); } }
            public long Ticks { get { return Interlocked.Read(ref _ticks); } }

            private void Schedule()
            {
                Task delay = Task.Delay(25);
                Task continuation = delay.ContinueWith(_ =>
                {
                    if (_stop) return;
                    long now = Stopwatch.GetTimestamp();
                    RecordMax(ref _maxGapTicks, now - Interlocked.Read(ref _lastTicks));
                    Interlocked.Exchange(ref _lastTicks, now);
                    Interlocked.Increment(ref _ticks);
                    Schedule();
                }, TaskScheduler.Default);
                GC.KeepAlive(continuation);
            }
        }

        /// <summary>
        /// The core of the liveness layer: a dedicated thread that fails the run        /// the moment the process heartbeat goes stale past the stall bound. That
        /// turns a plain hang (a deadlock, a wedged callback) into a non-zero exit
        /// instead of an unbounded wait, which is what the bound is for.
        ///
        /// It also hammers the cheap public entry points each iteration and
        /// records the worst iteration time: the "is any entry point stuck behind
        /// a lock" probe. The watchdog thread is dedicated, so ThreadPool
        /// starvation does not inflate its own iteration time.
        /// </summary>
        private sealed class StallWatchdog
        {
            private readonly ProcessHeartbeat _heartbeat;
            private readonly double _stallBoundMs;
            private Thread _thread;
            private volatile bool _stop;
            private long _maxIterTicks;
            private long _errors;

            public StallWatchdog(ProcessHeartbeat heartbeat, double stallBoundMs)
            {
                _heartbeat = heartbeat;
                _stallBoundMs = stallBoundMs;
            }

            public void Start()
            {
                _thread = new Thread(Loop);
                _thread.IsBackground = true;
                _thread.Name = "stall-watchdog";
                _thread.Start();
            }

            public void Stop() { _stop = true; }
            public void ResetMax() { Interlocked.Exchange(ref _maxIterTicks, 0); }
            public double MaxIterMs { get { return TicksToMs(Interlocked.Read(ref _maxIterTicks)); } }
            public long Errors { get { return Interlocked.Read(ref _errors); } }

            public ProcessHeartbeat Heartbeat { get { return _heartbeat; } }

            /// <summary>The stall decision, isolated so a self-check can pin it:
            /// true when the heartbeat has gone stale past the bound.</summary>
            public static bool IsStalled(double sinceBeatMs, double stallBoundMs)
            {
                return sinceBeatMs > stallBoundMs;
            }

            private void Loop()
            {
                // Wait for the heartbeat's first tick so a slow startup is never
                // mistaken for a stall.
                long waited = 0;
                while (!_stop && _heartbeat.Ticks == 0 && waited < (long)(_stallBoundMs * 4))
                {
                    Thread.Sleep(10);
                    waited += 10;
                }

                while (!_stop)
                {
                    long t0 = Stopwatch.GetTimestamp();
                    try
                    {
                        int subscriptions = S.SubscriptionCount;
                        int listeners = S.ListenerCount;
                        int origin = S.ListenerCountWithOrigin("B");
                        int live = C.LiveConnectionCount();
                        int liveUdf = C.LiveUdfConnectionCount();
                        double count = NatsUdfCore.NatsUDFSubscriptionCount();
                        GC.KeepAlive(subscriptions + listeners + origin + live + liveUdf + count);
                    }
                    catch
                    {
                        Interlocked.Increment(ref _errors);
                    }
                    RecordMax(ref _maxIterTicks, Stopwatch.GetTimestamp() - t0);

                    // The stall check: the heartbeat thread must have touched its
                    // timestamp within the bound. A stale timestamp means the
                    // process (or the heartbeat thread itself) stopped making
                    // progress -> a hang, and the run must fail now, not hang.
                    double sinceBeat = TicksToMs(Stopwatch.GetTimestamp() - _heartbeat.LastBeatTicks);
                    if (_heartbeat.Ticks > 0 && IsStalled(sinceBeat, _stallBoundMs))
                    {
                        Interlocked.Exchange(ref ExitCodeByWatchdog, 1);
                        Console.WriteLine("[WATCHDOG] STALL: process heartbeat " + Fmt(sinceBeat)
                            + "ms stale (bound " + Fmt(_stallBoundMs) + "ms) -> failing the run, exit 1");
                        try { Console.Out.Flush(); } catch { }
                        Environment.Exit(1);
                    }

                    Thread.Sleep(25);
                }
            }
        }

        /// <summary>
        /// Live delivery probe: a real runtime subscription on a dedicated subject
        /// fed by a dedicated publisher thread; records inter-arrival gaps,
        /// latency and the resume time after an attack window.
        /// </summary>
        private sealed class DeliveryProbe
        {
            private readonly object _gate = new object();
            private NatsConnectionManager _connections;
            private string _endpoint;
            private string _subject;
            private IDisposable _token;
            private volatile bool _stop;
            private Thread _publisher;
            private long _count;
            private long _lastArrivalTicks;
            private long _maxGapTicks;
            private long _maxLatencyTicks;
            private long _published;
            private long _publishErrors;
            private long _lastPublishTicks;

            public void Start(NatsConnectionManager connections, NatsSubscriptionManager subscriptions, string endpoint, string subject)
            {
                _connections = connections;
                _endpoint = endpoint;
                _subject = subject;
                Interlocked.Exchange(ref _lastArrivalTicks, Stopwatch.GetTimestamp());
                _token = subscriptions.Subscribe(endpoint, subject, null, OnMessage, "LIVENESS");
                _publisher = new Thread(PublishLoop) { IsBackground = true, Name = "delivery-publisher" };
                _publisher.Start();
            }

            public long Count { get { return Interlocked.Read(ref _count); } }
            public long Published { get { return Interlocked.Read(ref _published); } }
            public long PublishErrors { get { return Interlocked.Read(ref _publishErrors); } }
            public void ResetMax() { Interlocked.Exchange(ref _maxGapTicks, 0); Interlocked.Exchange(ref _maxLatencyTicks, 0); }
            public double MaxGapMs { get { return TicksToMs(Interlocked.Read(ref _maxGapTicks)); } }
            public double MaxLatencyMs { get { return TicksToMs(Interlocked.Read(ref _maxLatencyTicks)); } }

            public double ResumeAfter(long previousCount, double timeoutMs)
            {
                long start = Stopwatch.GetTimestamp();
                long deadline = start + MsToTicks(timeoutMs);
                while (Stopwatch.GetTimestamp() < deadline)
                {
                    if (Interlocked.Read(ref _count) > previousCount) return MsSince(start);
                    Thread.Sleep(5);
                }
                return -1;
            }

            public void Dispose()
            {
                _stop = true;
                try { if (_token != null) _token.Dispose(); } catch { }
                if (_publisher != null) _publisher.Join(5000);
            }

            private void OnMessage(NatsMessage message)
            {
                long now = Stopwatch.GetTimestamp();
                lock (_gate)
                {
                    if (Interlocked.Read(ref _count) > 0)
                        RecordMax(ref _maxGapTicks, now - Interlocked.Read(ref _lastArrivalTicks));
                    Interlocked.Exchange(ref _lastArrivalTicks, now);
                    Interlocked.Increment(ref _count);
                    long publishedAt = Interlocked.Read(ref _lastPublishTicks);
                    if (publishedAt > 0) RecordMax(ref _maxLatencyTicks, now - publishedAt);
                }
            }

            private void PublishLoop()
            {
                var connection = _connections.GetConnection(_endpoint);
                long n = 0;
                while (!_stop)
                {
                    try
                    {
                        Interlocked.Exchange(ref _lastPublishTicks, Stopwatch.GetTimestamp());
                        Observe(connection.PublishAsync(_subject, "hb-" + Interlocked.Increment(ref n)).AsTask());
                        Interlocked.Increment(ref _published);
                    }
                    catch
                    {
                        Interlocked.Increment(ref _publishErrors);
                    }
                    Thread.Sleep(25);
                }
            }
        }

        // ------------------------------------------------------------------
        // Observer double (self-contained: this project owns only test\LivenessTests)
        // ------------------------------------------------------------------

        /// <summary>IExcelObserver double recording OnNext values and the
        /// completion count.</summary>
        private sealed class RecordingObserver : IExcelObserver
        {
            private readonly object _gate = new object();
            private readonly List<object> _values = new List<object>();
            private int _completedCount;

            public void OnNext(object value)
            {
                lock (_gate) _values.Add(value);
            }

            public void OnError(Exception exception)
            {
                lock (_gate) _values.Add(exception);
            }

            public void OnCompleted()
            {
                Interlocked.Increment(ref _completedCount);
            }

            public object[] Values
            {
                get { lock (_gate) return _values.ToArray(); }
            }

            public int CompletedCount { get { return Volatile.Read(ref _completedCount); } }
        }
    }
}
