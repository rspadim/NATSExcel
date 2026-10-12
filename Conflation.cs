using System;

namespace NATSExcel
{
    /// <summary>
    /// Real-time delivery conflation window (DESIGN.md §8, PLAN F2.2).
    ///
    /// A fast feed keeps publishing while Excel can only consume a few updates
    /// per second; pushing every message would cause a recalculation cascade and
    /// screen flicker. The window is the single source of truth: the latest value
    /// of a topic is delivered at most once per window.
    ///
    /// Pure and deterministic: the clock is always passed in by the caller, so
    /// the helpers are unit-testable without a timer.
    /// </summary>
    internal static class Conflation
    {
        /// <summary>Maximum window: one hour. A larger configured value is capped.</summary>
        public const int MaxWindowMs = 3600000;

        /// <summary>Clamp a window to [0, <see cref="MaxWindowMs"/>]; a negative
        /// value means "no window" (0).</summary>
        public static int Sanitize(int windowMs)
        {
            if (windowMs < 0) return 0;
            if (windowMs > MaxWindowMs) return MaxWindowMs;
            return windowMs;
        }

        /// <summary>
        /// Resolve the effective window from the configuration. An explicit
        /// <paramref name="conflationMs"/> always wins (sanitized). When it is
        /// absent (null), the legacy boolean decides: coalescing on means one
        /// push per Excel tick (<paramref name="excelTickMs"/>), off means no
        /// window (0 = one push per message).
        /// </summary>
        public static int Resolve(int? conflationMs, bool coalesceRealtimeUpdates, int excelTickMs)
        {
            if (conflationMs.HasValue)
                return Sanitize(conflationMs.Value);
            if (!coalesceRealtimeUpdates)
                return 0;
            return Sanitize(excelTickMs);
        }

        /// <summary>True when a message must be pushed immediately: the window is
        /// off, so nothing is allowed to wait for the next flush.</summary>
        public static bool PushImmediately(int windowMs)
        {
            return windowMs <= 0;
        }

        /// <summary>
        /// True when a push is due: the window has elapsed since the last push.
        /// A window of 0 is always due. A clock that moved backwards (a negative
        /// elapsed time) must not stall a cell forever, so it is treated as due.
        /// </summary>
        public static bool IsDue(long lastPushTicks, int windowMs, long nowTicks)
        {
            if (windowMs <= 0) return true;
            if (lastPushTicks == 0) return true; // never pushed yet
            long elapsed = nowTicks - lastPushTicks;
            if (elapsed < 0) return true; // clock went backwards: do not starve
            return elapsed >= windowMs * TimeSpan.TicksPerMillisecond;
        }
    }
}
