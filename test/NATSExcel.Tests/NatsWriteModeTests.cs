using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins the write-mode decision table (PLAN.md §1.3, DESIGN.md §8.1
    /// <c>SyncWrite</c>): the three modes crossed with reply-dependence, plus the
    /// exact fire-and-forget marker per mode. The default <c>fireforget</c> only
    /// fires result-agnostic writes; <c>fireforget-all</c> also fires
    /// reply-dependent ones; <c>sync</c> never fires.
    /// </summary>
    public class NatsWriteModeTests
    {
        // ------------------------------------------------------------------
        // The full matrix: 3 modes x (reply-dependent true/false).
        // ------------------------------------------------------------------

        [Theory]
        // sync: never fire and forget, whatever the reply dependence.
        [InlineData("sync", true, false)]
        [InlineData("sync", false, false)]
        // fireforget (default): only result-agnostic writes.
        [InlineData("fireforget", true, false)]
        [InlineData("fireforget", false, true)]
        // fireforget-all: everything, reply-dependent writes included.
        [InlineData("fireforget-all", true, true)]
        [InlineData("fireforget-all", false, true)]
        public void ShouldFireAndForget_MatchesTheDocumentedMatrix(
            string mode, bool replyDependent, bool expected)
        {
            Assert.Equal(expected, NatsUdfCommon.ShouldFireAndForget(mode, replyDependent));
        }

        [Fact]
        public void ShouldFireAndForget_UnknownOrBlankModeFallsBackToDefault()
        {
            // The default is "fireforget", so it only fires result-agnostic writes.
            Assert.False(NatsUdfCommon.ShouldFireAndForget("bogus", replyDependent: true));
            Assert.True(NatsUdfCommon.ShouldFireAndForget("bogus", replyDependent: false));
            Assert.False(NatsUdfCommon.ShouldFireAndForget(null, replyDependent: true));
            Assert.True(NatsUdfCommon.ShouldFireAndForget("", replyDependent: false));
        }

        [Fact]
        public void ShouldFireAndForget_NormalizesCaseAndWhitespace()
        {
            Assert.False(NatsUdfCommon.ShouldFireAndForget("  SYNC  ", replyDependent: false));
            Assert.True(NatsUdfCommon.ShouldFireAndForget(" FireForget ", replyDependent: false));
            Assert.True(NatsUdfCommon.ShouldFireAndForget(" FireForget-All ", replyDependent: true));
        }

        // ------------------------------------------------------------------
        // The marker per mode.
        // ------------------------------------------------------------------

        [Fact]
        public void FireAndForgetMarker_ReturnsTheExactMarkerPerMode()
        {
            Assert.Equal("OK (fire and forget)", NatsUdfCommon.FireAndForgetMarker("fireforget"));
            Assert.Equal("OK (fire and forget: all)", NatsUdfCommon.FireAndForgetMarker("fireforget-all"));
        }

        [Fact]
        public void FireAndForgetMarker_NonAllModesGetThePlainMarker()
        {
            // sync never fires, but the marker helper still answers the plain
            // marker for it (only fireforget-all gets the ": all" variant).
            Assert.Equal(NatsUdfCommon.FireAndForget, NatsUdfCommon.FireAndForgetMarker("sync"));
            Assert.Equal(NatsUdfCommon.FireAndForget, NatsUdfCommon.FireAndForgetMarker("bogus"));
            Assert.Equal(NatsUdfCommon.FireAndForget, NatsUdfCommon.FireAndForgetMarker(null));
        }

        [Fact]
        public void FireAndForgetMarker_MatchesTheConstants()
        {
            Assert.Equal(NatsUdfCommon.FireAndForget, NatsUdfCommon.FireAndForgetMarker(NatsUdfCommon.SyncWriteFireForget));
            Assert.Equal(NatsUdfCommon.FireAndForgetAll, NatsUdfCommon.FireAndForgetMarker(NatsUdfCommon.SyncWriteFireForgetAll));
        }
    }
}
