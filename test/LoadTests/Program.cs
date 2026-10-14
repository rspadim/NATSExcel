using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NATS.Client.Core;
using NATSExcel;

namespace NATSExcel.LoadTests
{
    /// <summary>
    /// Load test for the subscription broadcast path, mirroring the sibling
    /// project's load layer. It drives the REAL <see cref="NatsSubscriptionManager"/>
    /// (linked production source) against a live NATS server, and reports the
    /// throughput plus the allocated bytes per received message.
    ///
    /// Usage:
    ///   LoadTests "nats://127.0.0.1:4222" [seconds] [publishers] [listeners] [pattern] [queue]
    ///   LoadTests manager "nats://127.0.0.1:4222" [seconds] [publishers] [listeners]
    ///   LoadTests raw     "nats://127.0.0.1:4222" [seconds] [publishers] [listeners]
    /// `manager` (the default) drives the REAL NatsSubscriptionManager broadcast
    /// path; `raw` is the plain NATS.Client.Core baseline for comparison.
    /// A publisher count of 0 means listen-only (an external generator publishes):
    ///   nats bench pub &lt;subject&gt; --size 300 --msgs 100000 --clients 4
    ///
    /// The run FAILS (exit 1) when the measured throughput is below a conservative
    /// floor (env NATSEXCEL_LOAD_MIN_MSGS_PER_SEC, default 5000) unless listen-only,
    /// so a real regression is visible instead of just printed. Hosts are passed
    /// as arguments only; never commit a real endpoint.
    /// </summary>
    internal static class Program
    {
        private const int DefaultMinMsgsPerSec = 5000;
        // async Main: a console app has no SynchronizationContext, so awaiting is
        // the correct way to wait; the analyzer (VSTHRD002) rejects a blocking
        // GetAwaiter().GetResult() here.
        private static async Task<int> Main(string[] args)
        {
            // Optional leading mode: `manager` (default) or `raw`. Accepting it
            // first keeps the historical `LoadTests "url" ...` form working.
            string mode = "manager";
            int offset = 0;
            if (args.Length > 0 &&
                (string.Equals(args[0], "manager", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(args[0], "raw", StringComparison.OrdinalIgnoreCase)))
            {
                mode = args[0].ToLowerInvariant();
                offset = 1;
            }

            string url = args.Length > offset ? args[offset] : "nats://127.0.0.1:4222";
            int seconds = Arg(args, offset + 1, 10);
            int publishers = Arg(args, offset + 2, 2);
            int listeners = Arg(args, offset + 3, 1);
            string pattern = args.Length > offset + 4 ? args[offset + 4] : null;
            string queue = args.Length > offset + 5 ? args[offset + 5] : null;

            string subject = pattern ?? "load." + Guid.NewGuid().ToString("N");
            Console.WriteLine("NATSExcel load test");
            Console.WriteLine("  mode       : " + mode);
            Console.WriteLine("  url        : " + url);
            Console.WriteLine("  seconds    : " + seconds);
            Console.WriteLine("  publishers : " + publishers + (publishers == 0 ? " (listen-only)" : string.Empty));
            Console.WriteLine("  listeners  : " + listeners);
            Console.WriteLine("  subject    : " + subject + (queue != null ? " (queue " + queue + ")" : string.Empty));
            Console.WriteLine();

            try
            {
                return string.Equals(mode, "raw", StringComparison.Ordinal)
                    ? await RunRawAsync(url, seconds, publishers, listeners, subject, queue).ConfigureAwait(false)
                    : await RunAsync(url, seconds, publishers, listeners, subject, queue).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL: " + Flatten(ex));
                return 1;
            }
        }

        /// <summary>
        /// The throughput floor. A regression below it fails the run; listen-only
        /// (an external publisher) has no floor because this process does not
        /// control the rate. Overridable with NATSEXCEL_LOAD_MIN_MSGS_PER_SEC=0.
        /// </summary>
        private static int MinMsgsPerSec()
        {
            string env = Environment.GetEnvironmentVariable("NATSEXCEL_LOAD_MIN_MSGS_PER_SEC");
            return int.TryParse(env, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : DefaultMinMsgsPerSec;
        }

        private static async Task<int> RunAsync(string url, int seconds, int publishers, int listeners, string subject, string queue)
        {
            var connections = NatsRuntime.Connections;
            var subscriptions = NatsRuntime.Subscriptions;
            string endpoint = url;

            long received = 0;
            long payloadBytes = 0;
            var tokens = new List<IDisposable>();

            for (int i = 0; i < listeners; i++)
            {
                tokens.Add(subscriptions.Subscribe(endpoint, subject, queue, message =>
                {
                    Interlocked.Increment(ref received);
                    if (message.DataBytes != null)
                        Interlocked.Add(ref payloadBytes, message.DataBytes.Length);
                }, "LOAD"));
            }
            Console.WriteLine("  subscribed (" + listeners + " listener(s))");

            var stop = new CancellationTokenSource(TimeSpan.FromSeconds(seconds));
            var publisherTasks = new List<Task>();
            if (publishers > 0)
            {
                var connection = connections.GetConnection(endpoint);
                for (int p = 0; p < publishers; p++)
                {
                    int worker = p;
                    publisherTasks.Add(Task.Run(async () =>
                    {
                        // A UNIQUE payload per message: the subscription manager
                        // suppresses identical consecutive payloads for a literal
                        // subject (SkipRepeatedMessages), so a constant payload
                        // would be dropped after the first delivery and the run
                        // would measure nothing.
                        long n = 0;
                        while (!stop.IsCancellationRequested)
                        {
                            string payload = NewPayload(300, worker, n++);
                            try { await connection.PublishAsync(subject, payload).ConfigureAwait(false); }
                            catch (OperationCanceledException) { return; }
                            catch { return; }
                        }
                    }));
                }
                Console.WriteLine("  publishing with " + publishers + " task(s) (unique payloads)");
            }
            else
            {
                Console.WriteLine("  listen-only: publish externally to " + subject);
            }

            long before = GC.GetTotalMemory(forceFullCollection: false);
            int gen0 = GC.CollectionCount(0), gen1 = GC.CollectionCount(1), gen2 = GC.CollectionCount(2);
            var sw = Stopwatch.StartNew();
            await Task.Delay(TimeSpan.FromSeconds(seconds)).ConfigureAwait(false);
            stop.Cancel();
            try { await Task.WhenAll(publisherTasks).ConfigureAwait(false); } catch { }
            sw.Stop();
            long after = GC.GetTotalMemory(forceFullCollection: false);

            long count = Interlocked.Read(ref received);
            long bytes = Interlocked.Read(ref payloadBytes);
            double secs = sw.Elapsed.TotalSeconds;
            double allocPerMsg = count == 0 ? 0 : (after - before) / (double)count;
            Console.WriteLine();
            Console.WriteLine("RESULTS (manager path)");
            Console.WriteLine("  received      : " + count.ToString("N0", CultureInfo.InvariantCulture) + " messages");
            Console.WriteLine("  payload       : " + (bytes / 1024.0 / 1024.0).ToString("0.0", CultureInfo.InvariantCulture) + " MiB");
            Console.WriteLine("  elapsed       : " + secs.ToString("0.0", CultureInfo.InvariantCulture) + " s");
            Console.WriteLine("  throughput    : " + (count / secs).ToString("N0", CultureInfo.InvariantCulture) + " msg/s");
            Console.WriteLine("  alloc/message : " + allocPerMsg.ToString("0.0", CultureInfo.InvariantCulture) + " bytes");
            Console.WriteLine("  GC            : gen0 " + (GC.CollectionCount(0) - gen0)
                + ", gen1 " + (GC.CollectionCount(1) - gen1)
                + ", gen2 " + (GC.CollectionCount(2) - gen2));

            foreach (var t in tokens) t.Dispose();
            subscriptions.Dispose();
            connections.Shutdown();

            double rate = count / secs;
            Console.WriteLine("  floor         : " + MinMsgsPerSec().ToString("N0", CultureInfo.InvariantCulture) + " msg/s");
            if (publishers == 0)
                return 0;                              // external publisher: no floor
            if (count == 0)
            {
                Console.WriteLine("FAIL: no message was received - the broadcast path delivered nothing");
                return 1;
            }
            int floor = MinMsgsPerSec();
            if (rate < floor)
            {
                Console.WriteLine("FAIL: throughput " + rate.ToString("N0", CultureInfo.InvariantCulture)
                    + " msg/s is below the floor " + floor.ToString("N0", CultureInfo.InvariantCulture) + " msg/s");
                return 1;
            }
            return 0;
        }

        /// <summary>
        /// The plain <c>NATS.Client.Core</c> baseline: one raw subscription per
        /// listener, no <see cref="NatsSubscriptionManager"/>. Same workload and
        /// floor as the manager mode, so the two can be compared.
        /// </summary>
        private static async Task<int> RunRawAsync(string url, int seconds, int publishers, int listeners, string subject, string queue)
        {
            await using var connection = new NatsConnection(new NatsOpts { Url = url });

            long received = 0;
            long payloadBytes = 0;
            var subs = new List<INatsSub<byte[]>>();
            for (int i = 0; i < listeners; i++)
            {
                var sub = await connection.SubscribeCoreAsync<byte[]>(subject, queueGroup: queue).ConfigureAwait(false);
                subs.Add(sub);
                _ = Task.Run(async () =>
                {
                    var reader = sub.Msgs;
                    while (await reader.WaitToReadAsync().ConfigureAwait(false))
                    {
                        while (reader.TryRead(out var msg))
                        {
                            Interlocked.Increment(ref received);
                            if (msg.Data != null)
                                Interlocked.Add(ref payloadBytes, msg.Data.Length);
                        }
                    }
                });
            }
            Console.WriteLine("  subscribed (" + listeners + " raw listener(s))");

            var stop = new CancellationTokenSource(TimeSpan.FromSeconds(seconds));
            var publisherTasks = new List<Task>();
            if (publishers > 0)
            {
                for (int p = 0; p < publishers; p++)
                {
                    int worker = p;
                    publisherTasks.Add(Task.Run(async () =>
                    {
                        long n = 0;
                        while (!stop.IsCancellationRequested)
                        {
                            try { await connection.PublishAsync(subject, NewPayload(300, worker, n++)).ConfigureAwait(false); }
                            catch (OperationCanceledException) { return; }
                            catch { return; }
                        }
                    }));
                }
                Console.WriteLine("  publishing with " + publishers + " task(s) (unique payloads)");
            }
            else
            {
                Console.WriteLine("  listen-only: publish externally to " + subject);
            }

            var sw = Stopwatch.StartNew();
            await Task.Delay(TimeSpan.FromSeconds(seconds)).ConfigureAwait(false);
            stop.Cancel();
            try { await Task.WhenAll(publisherTasks).ConfigureAwait(false); } catch { }
            sw.Stop();

            long count = Interlocked.Read(ref received);
            long bytes = Interlocked.Read(ref payloadBytes);
            double secs = sw.Elapsed.TotalSeconds;
            double rate = count / secs;
            Console.WriteLine();
            Console.WriteLine("RESULTS (raw path)");
            Console.WriteLine("  received      : " + count.ToString("N0", CultureInfo.InvariantCulture) + " messages");
            Console.WriteLine("  payload       : " + (bytes / 1024.0 / 1024.0).ToString("0.0", CultureInfo.InvariantCulture) + " MiB");
            Console.WriteLine("  elapsed       : " + secs.ToString("0.0", CultureInfo.InvariantCulture) + " s");
            Console.WriteLine("  throughput    : " + rate.ToString("N0", CultureInfo.InvariantCulture) + " msg/s");
            Console.WriteLine("  floor         : " + MinMsgsPerSec().ToString("N0", CultureInfo.InvariantCulture) + " msg/s");
            if (publishers == 0)
                return 0;
            if (count == 0)
            {
                Console.WriteLine("FAIL: no message was received");
                return 1;
            }
            int floor = MinMsgsPerSec();
            if (rate < floor)
            {
                Console.WriteLine("FAIL: throughput " + rate.ToString("N0", CultureInfo.InvariantCulture)
                    + " msg/s is below the floor " + floor.ToString("N0", CultureInfo.InvariantCulture) + " msg/s");
                return 1;
            }
            return 0;
        }

        private static string NewPayload(int size, int worker, long seq)
        {
            var sb = new StringBuilder(size + 48);
            sb.Append("{\"symbol\":\"X/Y\",\"bid\":1.0,\"w\":");
            sb.Append(worker.ToString(CultureInfo.InvariantCulture));
            sb.Append(",\"n\":");
            sb.Append(seq.ToString(CultureInfo.InvariantCulture));
            sb.Append(",\"pad\":\"");
            while (sb.Length < size - 3) sb.Append('x');
            sb.Append("\"}");
            return sb.ToString();
        }

        private static int Arg(string[] args, int index, int fallback)
        {
            if (args.Length <= index) return fallback;
            return int.TryParse(args[index], NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : fallback;
        }

        private static string Flatten(Exception ex)
        {
            var parts = new List<string>();
            for (var e = ex; e != null; e = e.InnerException) parts.Add(e.Message);
            return string.Join(" | ", parts);
        }
    }
}
