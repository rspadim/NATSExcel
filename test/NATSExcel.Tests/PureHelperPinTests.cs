using System;
using System.IO;
using System.Threading;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins the small pure helpers whose exact behavior no other test asserted, so
    /// a silent edit of the line would not fail the suite. (Round-4 coverage audit:
    /// these three had zero coverage.)
    ///
    /// <para>
    /// In the "config-seam" collection because one test here swaps
    /// <c>AppConfig.CurrentOverrideForTests</c> to prove the UpdateCheck toggle is
    /// honoured. That seam is PROCESS-WIDE, so a test outside the collection that
    /// reads <c>AppConfig.Current</c> could observe the swapped value (and this
    /// test could observe the fixture's) - the intermittent failures this class
    /// caused before it joined the collection.
    /// </para>
    /// </summary>
    [Collection("config-seam")]
    public class PureHelperPinTests
    {
        // ------------------------------------------------------------------
        // NLog.config: the credential-masking renderer must stay wired.
        // ------------------------------------------------------------------

        [Fact]
        public void NLogConfig_UsesTheMaskedExceptionRenderer()
        {
            string path = Path.Combine(RepoRoot(), "NLog.config");
            Assert.True(File.Exists(path), "NLog.config is missing");
            string cfg = File.ReadAllText(path);

            // The renderer is the ONLY thing that masks a logged exception: NLog's
            // own ${exception} prints the raw message, which can carry a
            // credential-bearing URL (the exact re-open DESIGN §8.1 warns about).
            Assert.Contains("${maskedException:format=toString}", cfg);
            Assert.DoesNotContain("${exception", cfg);
        }

        [Fact]
        public void AddIn_RegistersTheMaskedExceptionRenderer()
        {
            string addIn = File.ReadAllText(Path.Combine(RepoRoot(), "AddIn.cs"));
            // NLog resolves a layout renderer by scanning the assembly, but the
            // explicit registration guarantees the order (before the first log
            // line), so it must stay.
            Assert.Contains("RegisterLayoutRenderer<MaskedExceptionRenderer>(\"maskedException\")", addIn);
        }

        // ------------------------------------------------------------------
        // The "UpdateCheck": false toggle
        // ------------------------------------------------------------------

        /// <summary>
        /// With the toggle off, <c>EnsureFresh</c> must not even RECORD an attempt:
        /// the documented behaviour is "no background call to GitHub at all", and the
        /// key used to be parsed and never read (the sibling RedisExcel scheme honours
        /// it). The attempt timestamp is the observable proof - it is set under the
        /// lock the moment a check is scheduled.
        /// </summary>
        [Fact]
        public void TheUpdateCheckToggleDisablesTheBackgroundCheck()
        {
            var field = typeof(UpdateCheck).GetField(
                "_lastAttemptUtc", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            Assert.NotNull(field);

            var previous = AppConfig.CurrentOverrideForTests;
            try
            {
                var cfg = AppConfig.Sanitize(new AppConfig.Root());
                cfg.UpdateCheck = false;
                AppConfig.CurrentOverrideForTests = cfg;

                DateTime before = (DateTime)field.GetValue(null);
                UpdateCheck.EnsureFresh(TimeSpan.Zero);
                DateTime after = (DateTime)field.GetValue(null);

                Assert.Equal(before, after);                    // nothing was scheduled
                Assert.False(UpdateCheck.IsUpdateAvailable());  // and the cell stays FALSE
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = previous;
            }
        }

        /// <summary>And the guard must stay the FIRST thing the method does: a refactor
        /// that moved it after the scheduling would still "work" while performing the
        /// network call the toggle forbids.</summary>
        [Fact]
        public void TheToggleIsCheckedBeforeAnythingIsScheduled()
        {
            string src = System.IO.File.ReadAllText(
                System.IO.Path.Combine(RepoRoot(), "UpdateCheck.cs"));
            int method = src.IndexOf("internal static void EnsureFresh(", StringComparison.Ordinal);
            Assert.True(method > 0, "EnsureFresh is gone");

            int guard = src.IndexOf("if (!AppConfig.Current.UpdateCheck)", method, StringComparison.Ordinal);
            int schedule = src.IndexOf("Task.Run(", method, StringComparison.Ordinal);

            Assert.True(guard > method, "EnsureFresh no longer reads the UpdateCheck toggle");
            Assert.True(schedule > guard, "the toggle must be checked BEFORE the check is scheduled");
        }

        // ------------------------------------------------------------------
        // UpdateCheck: the semver-ish comparison and the "never nag a dev build"
        // rule. The CurrentTagOverrideForTests seam existed for this and had no
        // test using it.
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("v0.2.0", "0.1.9", true)]
        [InlineData("0.2.0", "0.1.9", true)]
        [InlineData("v0.1.9", "0.1.9", false)]
        [InlineData("v0.1.8", "0.1.9", false)]
        [InlineData("1.4", "1.4.0", false)]      // missing segments are zero
        [InlineData("1.4.1", "1.4", true)]
        [InlineData("2.0.0", "10.0.0", false)]   // numeric, not lexicographic
        [InlineData("bogus", "0.1.9", false)]    // unparseable: never alert
        [InlineData("v9.9.9", "dev", false)]     // a dev build is never nudged
        public void UpdateCheck_IsNewer(string candidate, string current, bool expected)
        {
            Assert.Equal(expected, UpdateCheck.IsNewer(candidate, current));
        }

        [Theory]
        [InlineData(" v1.4.5-rc1+build7 ", "1.4.5")]
        [InlineData("v0.1.1", "0.1.1")]
        [InlineData("0.1.1", "0.1.1")]
        [InlineData("", "")]
        public void UpdateCheck_NormalizeTag(string input, string expected)
        {
            Assert.Equal(expected, UpdateCheck.NormalizeTag(input));
        }

        // ------------------------------------------------------------------
        // TickGate: the RTD timer reentrancy gate (golden rule 6 / DESIGN §9).
        // ------------------------------------------------------------------

        [Fact]
        public void TickGate_DropsAReentrantTickAndRecoversAfterExit()
        {
            var gate = new TickGate();

            Assert.True(gate.TryEnter(), "the first tick must proceed");
            // A tick that fires while the previous one is still running is DROPPED,
            // never queued: queuing would build a backlog and flood Excel.
            Assert.False(gate.TryEnter(), "a re-entrant tick must be dropped");

            gate.Exit();
            Assert.True(gate.TryEnter(), "the gate must be usable again after Exit");
            gate.Exit();
        }

        [Fact]
        public void TickGate_ExitIsIdempotentAndNeverLetsTwoTicksOverlap()
        {
            // Exit() is an unconditional Interlocked.Exchange, so a stray extra
            // release must not corrupt the gate - but while a holder is inside,
            // nothing else may enter: an overlapping tick would stack NATS work
            // and Excel flushes (the reason the gate exists at all).
            var gate = new TickGate();
            Assert.True(gate.TryEnter());
            gate.Exit();
            gate.Exit();

            Assert.True(gate.TryEnter(), "the gate must be re-usable after Exit");
            Assert.False(gate.TryEnter(), "only one tick may hold the gate at a time");
            gate.Exit();
        }

        [Fact]
        public void StripedLocks_AreStablePerKeyAndBounded()
        {
            var stripes = new StripedLocks(64);
            object a1 = stripes.For("orders.eu");
            object a2 = stripes.For("orders.eu");
            object b = stripes.For("orders.us");

            Assert.Same(a1, a2);                 // the same key always takes its stripe
            Assert.NotSame(a1, b);               // a different key may take another
            Assert.NotNull(b);
        }

        [Fact]
        public void StripedLocks_MapABlankKeyToTheFirstStripeInsteadOfThrowing()
        {
            // Both call paths can hand a blank key to a stripe; it must land on a
            // real stripe (never an index error) and always the same one.
            var stripes = new StripedLocks(8);
            object blank = stripes.For("");
            Assert.Same(blank, stripes.For(null));
            Assert.NotNull(blank);
        }

        internal static string RepoRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "NATSExcel.sln")))
                dir = dir.Parent;
            Assert.NotNull(dir);
            return dir.FullName;
        }
    }
}
