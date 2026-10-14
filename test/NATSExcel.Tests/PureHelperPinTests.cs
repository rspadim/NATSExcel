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
    /// </summary>
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
