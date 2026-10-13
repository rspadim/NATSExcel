using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// <see cref="NatsConnectionManager.WaitBounded"/> and
    /// <see cref="NatsConnectionManager.WaitBoundedResult{T}"/> are the only two
    /// bounded-blocking primitives in the codebase (AGENTS golden rule 10): an
    /// Excel UDF must return in-cell, so every blocking wait is capped. These are
    /// pure static helpers, so they are tested with synthetic tasks and no server.
    ///
    /// The bound is asserted with generous slack (a bounded wait may return a few
    /// timer ticks late, never early), not to the millisecond.
    /// </summary>
    public class NatsConnectionWaitTests
    {
        private const int BoundMs = 150;

        // ------------------------------------------------------------------
        // WaitBounded (bool: completed within the bound?)
        // ------------------------------------------------------------------

        [Fact]
        public void WaitBounded_CompletedTask_ReturnsTrue()
        {
            Assert.True(NatsConnectionManager.WaitBounded(Task.CompletedTask, 1000));
        }

        [Fact]
        public void WaitBounded_NullTask_ReturnsTrue()
        {
            // "A null task completed with nothing to wait for."
            Assert.True(NatsConnectionManager.WaitBounded(null, 1000));
        }

        [Fact]
        public void WaitBounded_FaultedTask_ReturnsTrue()
        {
            Assert.True(NatsConnectionManager.WaitBounded(
                Task.FromException(new InvalidOperationException("boom")), 1000));
        }

        [Fact]
        public void WaitBounded_CancelledTask_ReturnsTrue()
        {
            Assert.True(NatsConnectionManager.WaitBounded(
                Task.FromCanceled(new CancellationToken(canceled: true)), 1000));
        }

        [Fact]
        public void WaitBounded_TaskThatNeverCompletes_ReturnsFalseAtTheBound()
        {
            var tcs = new TaskCompletionSource<bool>();
            var sw = Stopwatch.StartNew();
            bool completed = NatsConnectionManager.WaitBounded(tcs.Task, BoundMs);
            sw.Stop();

            Assert.False(completed);
            Assert.True(
                sw.ElapsedMilliseconds >= BoundMs - 50,
                "returned before the bound: " + sw.ElapsedMilliseconds + "ms");
            Assert.True(
                sw.ElapsedMilliseconds < 3000,
                "waited far past the bound: " + sw.ElapsedMilliseconds + "ms");
        }

        [Fact]
        public void WaitBounded_ZeroAndNegativeBounds_ClampToZeroAndDoNotBlock()
        {
            var tcs = new TaskCompletionSource<bool>();
            var sw = Stopwatch.StartNew();

            Assert.False(NatsConnectionManager.WaitBounded(tcs.Task, 0));
            Assert.False(NatsConnectionManager.WaitBounded(tcs.Task, -1000));

            // An already-completed task is still reported as completed with a 0 bound.
            Assert.True(NatsConnectionManager.WaitBounded(Task.CompletedTask, 0));

            sw.Stop();
            Assert.True(sw.ElapsedMilliseconds < 2000, "a zero bound blocked for " + sw.ElapsedMilliseconds + "ms");
        }

        // ------------------------------------------------------------------
        // WaitBoundedResult<T> (value or TimeoutException at the bound)
        // ------------------------------------------------------------------

        [Fact]
        public void WaitBoundedResult_CompletedTask_ReturnsTheValue()
        {
            Assert.Equal(42, NatsConnectionManager.WaitBoundedResult(Task.FromResult(42), 1000));
        }

        [Fact]
        public void WaitBoundedResult_TaskCompletingWithinTheBound_ReturnsTheValue()
        {
            Task<int> task = Task.Run(async () =>
            {
                await Task.Delay(30);
                return 7;
            });

            Assert.Equal(7, NatsConnectionManager.WaitBoundedResult(task, 2000));
        }

        [Fact]
        public void WaitBoundedResult_FaultedTask_RethrowsTheOriginalExceptionUnwrapped()
        {
            Task<int> task = Task.FromException<int>(new InvalidOperationException("boom"));

            var ex = Assert.Throws<InvalidOperationException>(
                () => NatsConnectionManager.WaitBoundedResult(task, 1000));
            Assert.Equal("boom", ex.Message);
        }

        [Fact]
        public void WaitBoundedResult_TaskThatNeverCompletes_ThrowsTimeoutAtTheBound()
        {
            var tcs = new TaskCompletionSource<int>();
            var sw = Stopwatch.StartNew();

            Assert.Throws<TimeoutException>(
                () => NatsConnectionManager.WaitBoundedResult(tcs.Task, BoundMs));

            sw.Stop();
            Assert.True(
                sw.ElapsedMilliseconds >= BoundMs - 50,
                "timed out before the bound: " + sw.ElapsedMilliseconds + "ms");
            Assert.True(
                sw.ElapsedMilliseconds < 3000,
                "waited far past the bound: " + sw.ElapsedMilliseconds + "ms");
        }

        [Fact]
        public void WaitBoundedResult_NullTask_Throws()
        {
            Assert.Throws<ArgumentNullException>(
                () => NatsConnectionManager.WaitBoundedResult<int>(null, 100));
        }

        [Fact]
        public void WaitBoundedResult_NegativeBound_ClampsToZero()
        {
            Assert.Equal(5, NatsConnectionManager.WaitBoundedResult(Task.FromResult(5), -1000));

            var tcs = new TaskCompletionSource<int>();
            Assert.Throws<TimeoutException>(
                () => NatsConnectionManager.WaitBoundedResult(tcs.Task, -1000));
        }
    }
}
