using System;
using System.Collections;
using System.Reflection;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The per-key subject epoch map detects an unsubscribe racing a
    /// SubjectLatest install. It is dead weight for a key with no live listener,
    /// so it is pruned opportunistically - and the prune must never erase a key
    /// whose subscribe is still in flight (doing so made the install re-check
    /// "epoch == epochBefore" pass as 0 == 0 and leak a ghost listener).
    ///
    /// Adapted from the sibling RedisExcel suite (ChannelEpochLeakTests), which
    /// found exactly this class of bug in the same-shaped code; NATSExcel had the
    /// fix but no test pinning it.
    /// </summary>
    public class SubjectEpochLeakTests
    {
        private static readonly FieldInfo Epochs = typeof(NatsUdfCore)
            .GetField("_subjectEpochs", BindingFlags.Static | BindingFlags.NonPublic);
        private static readonly MethodInfo Bump = typeof(NatsUdfCore)
            .GetMethod("BumpSubjectEpoch", BindingFlags.Static | BindingFlags.NonPublic);
        private static readonly MethodInfo BeginWatch = typeof(NatsUdfCore)
            .GetMethod("BeginSubscribeWatch", BindingFlags.Static | BindingFlags.NonPublic);
        private static readonly MethodInfo EndWatch = typeof(NatsUdfCore)
            .GetMethod("EndSubscribeWatch", BindingFlags.Static | BindingFlags.NonPublic);

        private static IDictionary Dict()
        {
            Assert.NotNull(Epochs);
            return (IDictionary)Epochs.GetValue(null);
        }

        private static void BumpKey(string key)
        {
            Assert.NotNull(Bump);
            Bump.Invoke(null, new object[] { key });
        }

        /// <summary>
        /// One sweep fires per 1024 global bumps, and any 1024 consecutive bumps
        /// contain exactly one multiple of 1024, so this guarantees a sweep.
        /// </summary>
        private static void ForceSweep(string prefix)
        {
            for (int i = 0; i < 1024; i++)
                BumpKey(prefix + ":" + i);
        }

        [Fact]
        public void DistinctSubjectChurn_DoesNotGrowTheEpochMapUnboundedly()
        {
            var dict = Dict();
            dict.Clear();

            const int subjects = 5000;
            for (int i = 0; i < subjects; i++)
                BumpKey("epoch:leak:" + i);

            // The global-bump prune sweeps every 1024 bumps, so after 5000 distinct
            // keys the map must be far below 5000. Allow one sweep's slack plus the
            // current window: a regression to per-key pruning would leave ~5000.
            int count = dict.Count;
            dict.Clear();
            Assert.True(count < 2048,
                "the epoch map must stay bounded, got " + count + " entries for " + subjects + " subjects");
        }

        [Fact]
        public void ActuallyBumpsTheEpochPerKey()
        {
            var dict = Dict();
            dict.Clear();
            string key = "epoch:bump:" + Guid.NewGuid().ToString("N");

            long first = (long)Bump.Invoke(null, new object[] { key });
            long second = (long)Bump.Invoke(null, new object[] { key });
            dict.Clear();

            // The racing-install detection relies on this increment.
            Assert.Equal(1, first);
            Assert.Equal(2, second);
        }

        [Fact]
        public void EpochSweep_DoesNotPruneAnInFlightSubscribe()
        {
            // The sweep used to drop any entry with no live listener, INCLUDING one
            // whose SubjectLatest subscribe had read epochBefore but not installed
            // yet: erasing it made the install re-check (SubjectEpoch == epochBefore)
            // wrongly pass and leak a ghost listener. The in-flight watcher vetoes
            // the prune; this is the regression pin for that fix.
            Assert.NotNull(BeginWatch);
            Assert.NotNull(EndWatch);
            var dict = Dict();
            dict.Clear();

            string key = "epoch:watch:" + Guid.NewGuid().ToString("N");
            BumpKey(key);                                   // a racing unsubscribe/reset
            BeginWatch.Invoke(null, new object[] { key });   // subscribe in flight
            ForceSweep("epoch:watch:sweep");                 // must NOT sweep this key
            bool survived = dict.Contains(key);
            dict.Clear();

            Assert.True(survived, "a key with an in-flight subscribe must survive the epoch sweep");

            // Once the subscribe resolves with no listener, the key is prunable.
            BumpKey(key);
            EndWatch.Invoke(null, new object[] { key });
            ForceSweep("epoch:watch:sweep2");
            bool pruned = !dict.Contains(key);
            dict.Clear();
            Assert.True(pruned, "once the subscribe resolves with no listener, the key is prunable");
        }
        [Fact]
        public void TheLoserPathDecidesUnderTheLockAndNeverDropsALiveListenersValue()
        {
            // A source scan, because the file IS compiled here but this race has no
            // deterministic seam (it needs a concurrent unsubscribe to interleave
            // between the epoch sample and the cache removal). The invariant is
            // structural: the loser's verdict must be taken under the SAME lock that
            // makes the install atomic, and the cached value must never be dropped
            // while a listener for that key is registered.
            string src = System.IO.File.ReadAllText(
                System.IO.Path.Combine(PureHelperPinTests.RepoRoot(), "NatsUdfCore.cs"));

            int start = src.IndexOf(
                "private static void EnsureSubjectListener(", StringComparison.Ordinal);
            Assert.True(start > 0, "EnsureSubjectListener is gone");
            int end = src.IndexOf("private static ", start + 10, StringComparison.Ordinal);
            string method = end > start ? src.Substring(start, end - start) : src.Substring(start);

            // The verdict is recorded inside the install lock...
            int lockAt = method.IndexOf("lock (LifecycleLock(key))", StringComparison.Ordinal);
            int verdict = method.IndexOf("removedByEpoch = true;", StringComparison.Ordinal);
            Assert.True(lockAt > 0 && verdict > lockAt,
                "the loser's verdict must be decided inside the install lock, not re-sampled after it");

            // ...and the epoch compare happens ONCE: a second compare after the lock
            // is the bug this pins (it saw a LATER removal and deleted the fresh
            // winner's value).
            int compares = 0;
            for (int i = method.IndexOf("SubjectEpoch(key) != epochBefore", StringComparison.Ordinal);
                 i >= 0;
                 i = method.IndexOf("SubjectEpoch(key) != epochBefore", i + 1, StringComparison.Ordinal))
                compares++;
            Assert.Equal(1, compares);

            // The removal is gated on "nothing is live for this key".
            int remove = method.IndexOf("_latestMessages.TryRemove(key", StringComparison.Ordinal);
            Assert.True(remove > 0, "the loser no longer cleans its cached value");
            Assert.Contains("!_subjectListeners.ContainsKey(key)", method);
            Assert.True(method.IndexOf("!_subjectListeners.ContainsKey(key)", StringComparison.Ordinal) < remove,
                "the live-listener check must gate the removal");
        }
    }
}
