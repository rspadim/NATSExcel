using System;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The real-time conflation window (DESIGN.md §8, PLAN F2.2) is pure and
    /// deterministic: the clock is always passed in, so the helpers are tested
    /// without a timer. <c>Sanitize</c> clamps, <c>Resolve</c> applies the
    /// documented precedence, and <c>IsDue</c> never stalls a cell.
    /// </summary>
    public class ConflationTests
    {
        // ------------------------------------------------------------------
        // Sanitize.
        // ------------------------------------------------------------------

        [Fact]
        public void Sanitize_ClampsNegativeToZero()
        {
            Assert.Equal(0, Conflation.Sanitize(-1));
            Assert.Equal(0, Conflation.Sanitize(int.MinValue));
        }

        [Fact]
        public void Sanitize_CapsAtMaxWindow()
        {
            Assert.Equal(Conflation.MaxWindowMs, Conflation.Sanitize(Conflation.MaxWindowMs + 1));
            Assert.Equal(Conflation.MaxWindowMs, Conflation.Sanitize(int.MaxValue));
        }

        [Fact]
        public void Sanitize_KeepsAnInRangeValue()
        {
            Assert.Equal(0, Conflation.Sanitize(0));
            Assert.Equal(200, Conflation.Sanitize(200));
            Assert.Equal(Conflation.MaxWindowMs, Conflation.Sanitize(Conflation.MaxWindowMs));
        }

        [Fact]
        public void MaxWindowMs_IsOneHour()
        {
            Assert.Equal(3600000, Conflation.MaxWindowMs);
        }

        // ------------------------------------------------------------------
        // Resolve precedence.
        // ------------------------------------------------------------------

        [Fact]
        public void Resolve_ExplicitWindowWinsOverTheBoolean()
        {
            Assert.Equal(200, Conflation.Resolve(200, coalesceRealtimeUpdates: false, excelTickMs: 100));
            Assert.Equal(200, Conflation.Resolve(200, coalesceRealtimeUpdates: true, excelTickMs: 100));
        }

        [Fact]
        public void Resolve_ExplicitWindowIsSanitized()
        {
            Assert.Equal(0, Conflation.Resolve(-5, coalesceRealtimeUpdates: true, excelTickMs: 100));
            Assert.Equal(Conflation.MaxWindowMs, Conflation.Resolve(int.MaxValue, false, 100));
        }

        [Fact]
        public void Resolve_NullWindow_UsesTheExcelTickWhenCoalescing()
        {
            Assert.Equal(100, Conflation.Resolve(null, coalesceRealtimeUpdates: true, excelTickMs: 100));
        }

        [Fact]
        public void Resolve_NullWindow_CoalescingOffMeansNoWindow()
        {
            Assert.Equal(0, Conflation.Resolve(null, coalesceRealtimeUpdates: false, excelTickMs: 100));
        }

        [Fact]
        public void Resolve_NullWindow_TickValueIsAlsoSanitized()
        {
            Assert.Equal(0, Conflation.Resolve(null, coalesceRealtimeUpdates: true, excelTickMs: -10));
            Assert.Equal(
                Conflation.MaxWindowMs,
                Conflation.Resolve(null, coalesceRealtimeUpdates: true, excelTickMs: int.MaxValue));
        }

        // ------------------------------------------------------------------
        // PushImmediately.
        // ------------------------------------------------------------------

        [Fact]
        public void PushImmediately_IsTrueForANonPositiveWindow()
        {
            Assert.True(Conflation.PushImmediately(0));
            Assert.True(Conflation.PushImmediately(-1));
            Assert.False(Conflation.PushImmediately(1));
        }

        // ------------------------------------------------------------------
        // IsDue.
        // ------------------------------------------------------------------

        [Fact]
        public void IsDue_WithWindowZero_IsAlwaysDue()
        {
            Assert.True(Conflation.IsDue(0L, 0, 0L));
            Assert.True(Conflation.IsDue(12345L, 0, 12345L));
            Assert.True(Conflation.IsDue(long.MaxValue, -1, long.MaxValue));
        }

        [Fact]
        public void IsDue_WhenNeverPushed_IsDue()
        {
            Assert.True(Conflation.IsDue(0L, 1000, 500L));
        }

        [Fact]
        public void IsDue_WhenTheWindowHasElapsed_IsDue()
        {
            long windowTicks = 1000L * TimeSpan.TicksPerMillisecond;
            Assert.True(Conflation.IsDue(1000L, 1000, 1000L + windowTicks));
            Assert.True(Conflation.IsDue(1000L, 1000, 1000L + windowTicks + 1));
        }

        [Fact]
        public void IsDue_WhenTheWindowHasNotElapsed_IsNotDue()
        {
            long now = 1000L + 999L * TimeSpan.TicksPerMillisecond;
            Assert.False(Conflation.IsDue(1000L, 1000, now));
        }

        [Fact]
        public void IsDue_WhenTheClockGoesBackwards_DoesNotStall()
        {
            // A negative elapsed time must never trap a cell: it is treated as
            // due so the next flush can recover.
            Assert.True(Conflation.IsDue(5000L, 1000, 4000L));
        }
    }
}
