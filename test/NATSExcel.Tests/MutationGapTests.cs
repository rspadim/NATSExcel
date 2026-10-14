using System;
using System.Collections.Generic;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins production lines whose exact value/branch no other test asserted, so a
    /// silent mutation of the line would not fail the suite. Adapted from the
    /// sibling RedisExcel suite (MutationGapTests).
    /// </summary>
    [Collection("config-seam")]
    public class MutationGapTests
    {
        private static void WithTimeout(int timeoutMs, Action<AppConfig.Root> body)
        {
            var previous = AppConfig.CurrentOverrideForTests;
            try
            {
                var cfg = AppConfig.Sanitize(new AppConfig.Root());
                cfg.TimeoutMs = timeoutMs;
                AppConfig.CurrentOverrideForTests = cfg;
                body(cfg);
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = previous;
            }
        }

        [Fact]
        public void ResponseTimeoutMs_IsTheConfiguredTimeoutPlusTheMargin()
        {
            WithTimeout(7000, _ =>
                // 7000 + the 500 ms response margin.
                Assert.Equal(7500, NatsConnectionManager.ResponseTimeoutMs(NatsPool.Udf)));
        }

        [Fact]
        public void ResponseTimeoutMs_ClampsToTheCeilingInsteadOfOverflowing()
        {
            WithTimeout(int.MaxValue, _ =>
                // (long)configured + 500 would wrap an int; the bound must clamp.
                Assert.Equal(60000, NatsConnectionManager.ResponseTimeoutMs(NatsPool.RtdData)));
        }

        [Fact]
        public void ResponseTimeoutMs_NeverGoesBelowTheFloor()
        {
            WithTimeout(0, _ =>
                Assert.Equal(500, NatsConnectionManager.ResponseTimeoutMs(NatsPool.Udf)));
            // A negative timeout (a hand-edited file) is treated as 0, not as a
            // smaller budget.
            WithTimeout(-5000, _ =>
                Assert.Equal(500, NatsConnectionManager.ResponseTimeoutMs(NatsPool.Udf)));
        }

        [Fact]
        public void ResponseTimeoutMs_IsPositiveForEveryPool()
        {
            // One connection per endpoint multiplexes every family on NATS, so the
            // budget is shared today; this pins that no pool is left at zero (a
            // zero bound would make every call time out instantly).
            WithTimeout(1000, _ =>
            {
                foreach (NatsPool pool in Enum.GetValues(typeof(NatsPool)))
                    Assert.True(NatsConnectionManager.ResponseTimeoutMs(pool) >= 250,
                        "pool " + pool + " has a sub-floor response budget");
            });
        }
    }
}
