using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using ExcelDna.Integration;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Records every observer signal with a completion latch, so a test can assert
    /// on the exact delivery instead of sleeping. Adapted from the sibling
    /// RedisExcel suite's <c>ObserverDoubles</c>.
    /// </summary>
    internal sealed class RecordingObserver : IExcelObserver
    {
        private readonly object _gate = new object();
        private readonly List<object> _values = new List<object>();
        private readonly List<Exception> _errors = new List<Exception>();
        private readonly ManualResetEventSlim _completed = new ManualResetEventSlim(false);

        public void OnNext(object value) { lock (_gate) _values.Add(value); }
        public void OnError(Exception exception) { lock (_gate) _errors.Add(exception); }
        public void OnCompleted() { _completed.Set(); }

        /// <summary>TRUE when the terminal signal arrived within the window.</summary>
        public bool WaitForCompletion(int timeoutMs) { return _completed.Wait(timeoutMs); }

        public object[] Values { get { lock (_gate) return _values.ToArray(); } }
        public Exception[] Errors { get { lock (_gate) return _errors.ToArray(); } }

        /// <summary>The single delivered value, asserting there is exactly one.</summary>
        public object OnlyValue
        {
            get
            {
                var values = Values;
                Assert.True(values.Length == 1, "expected exactly one delivered value, got " + values.Length);
                return values[0];
            }
        }
    }

    /// <summary>
    /// The async write dispatch (PLAN S3.3): <c>NatsWriteObservable</c> is what an
    /// async UDF returns to Excel-DNA, and it carries the guarantees the whole
    /// AsyncWrites design rests on - the work never runs on the Excel thread, a
    /// duplicate subscribe never queues the write twice, an enqueue failure still
    /// reaches the cell as a named error, and one endpoint's writes stay FIFO. The
    /// round-4 audit found this type referenced by no test at all; the liveness
    /// suite exercises it end to end but never asserts these properties directly.
    /// </summary>
    public class NatsWriteObservableTests
    {
        /// <summary>A fresh endpoint per test keeps the per-endpoint FIFO (static
        /// state) from carrying anything between tests.</summary>
        private static string Endpoint(string tag)
        {
            return "test://observable-" + tag + "-" + Guid.NewGuid().ToString("N");
        }

        [Fact]
        public void AWriteDeliversItsValueAndCompletes()
        {
            var observer = new RecordingObserver();
            var observable = new NatsUdfAsync.NatsWriteObservable(Endpoint("value"), () => "OK");

            observable.Subscribe(observer);

            Assert.True(observer.WaitForCompletion(5000), "the cell must get a terminal signal");
            Assert.Equal(new object[] { "OK" }, observer.Values);
            Assert.Empty(observer.Errors);
        }

        [Fact]
        public void AThrowingWriteBecomesANamedErrorCell_NotAnUnhandledException()
        {
            var observer = new RecordingObserver();
            var observable = new NatsUdfAsync.NatsWriteObservable(Endpoint("throw"), () =>
            {
                throw new InvalidOperationException("boom");
            });

            observable.Subscribe(observer);

            Assert.True(observer.WaitForCompletion(5000), "a failure still has to terminate the cell");
            object value = observer.OnlyValue;
            Assert.IsType<string>(value);
            Assert.StartsWith(NatsUdfCommon.ErrorPrefix, (string)value);
            Assert.Contains("boom", (string)value);
            Assert.Empty(observer.Errors);   // the exception is the CELL's error, not the observer's
        }

        [Fact]
        public void TheWorkNeverRunsOnTheSubscribingThread()
        {
            // The whole reason AsyncWrites exists: a reply-dependent write must not
            // block the Excel thread. The continuation carries
            // RunContinuationsAsynchronously, so the work runs on the pool even when
            // the endpoint's queue is brand new (whose Tail is an already-completed
            // sentinel - the case that used to run INLINE).
            int callerThread = Thread.CurrentThread.ManagedThreadId;
            var observer = new RecordingObserver();
            var observable = new NatsUdfAsync.NatsWriteObservable(
                Endpoint("thread"), () => Thread.CurrentThread.ManagedThreadId);

            observable.Subscribe(observer);

            Assert.True(observer.WaitForCompletion(5000));
            Assert.NotEqual(callerThread, (int)observer.OnlyValue);
        }

        [Fact]
        public void SubscribeReturnsImmediatelyEvenWhenTheWriteIsSlow()
        {
            var observer = new RecordingObserver();
            var observable = new NatsUdfAsync.NatsWriteObservable(Endpoint("slow"), () =>
            {
                Thread.Sleep(700);
                return "slow-done";
            });

            var watch = Stopwatch.StartNew();
            observable.Subscribe(observer);
            watch.Stop();

            Assert.True(watch.ElapsedMilliseconds < 300,
                "Subscribe must not wait for the write (took " + watch.ElapsedMilliseconds + " ms)");
            Assert.True(observer.WaitForCompletion(10000));
            Assert.Equal("slow-done", observer.OnlyValue);
        }

        [Fact]
        public void ADuplicateSubscribeRunsTheWriteOnceAndDeliversTheSameValue()
        {
            // Excel-DNA registers one observable per call, so a second Subscribe is a
            // defensive path - but it must never enqueue the write twice (a re-run
            // could double-publish) and must still terminate the second cell.
            int runs = 0;
            var first = new RecordingObserver();
            var second = new RecordingObserver();
            var observable = new NatsUdfAsync.NatsWriteObservable(Endpoint("dup"), () =>
            {
                Interlocked.Increment(ref runs);
                return "once";
            });

            observable.Subscribe(first);
            observable.Subscribe(second);

            Assert.True(first.WaitForCompletion(5000) && second.WaitForCompletion(5000));
            Assert.Equal(1, runs);
            Assert.Equal("once", first.OnlyValue);
            Assert.Equal("once", second.OnlyValue);
        }

        [Fact]
        public void DisposingTheSubscriptionIsANoOpAndTheValueStillArrives()
        {
            // Excel-DNA may dispose the handle while the write is still queued (the
            // cell was recalculated, the sheet was closed). The write was already
            // issued, so it must run to completion: Dispose is deliberately a no-op.
            var observer = new RecordingObserver();
            var observable = new NatsUdfAsync.NatsWriteObservable(Endpoint("dispose"), () => "still-there");

            IDisposable handle = observable.Subscribe(observer);
            handle.Dispose();
            handle.Dispose();   // double dispose must not throw either

            Assert.True(observer.WaitForCompletion(5000));
            Assert.Equal("still-there", observer.OnlyValue);
        }

        [Fact]
        public void OneEndpointsWritesRunInSubmissionOrder()
        {
            // The per-endpoint FIFO is what makes "writes issued from a formula
            // evaluate in formula order" true; the works below would interleave (and
            // finish out of order) if the queue chained on the wrong tail or ran
            // them concurrently.
            string endpoint = Endpoint("fifo");
            var order = new List<int>();
            var observers = new List<RecordingObserver>();

            for (int i = 0; i < 4; i++)
            {
                int index = i;
                var observer = new RecordingObserver();
                observers.Add(observer);
                var observable = new NatsUdfAsync.NatsWriteObservable(endpoint, () =>
                {
                    lock (order) order.Add(index);
                    Thread.Sleep(40);
                    return index;
                });
                observable.Subscribe(observer);
            }

            foreach (var observer in observers)
                Assert.True(observer.WaitForCompletion(10000), "every queued write must complete");

            lock (order)
                Assert.Equal(new[] { 0, 1, 2, 3 }, order.ToArray());
        }
    }
}
