using ExcelDna.Integration;
using NLog;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NATSExcel
{
    /// <summary>
    /// Async dispatch for UDF work, gated by the AsyncWrites configuration key.
    ///
    /// NATS.Net is async-only (DESIGN.md §5.3): every NATS call is a Task, while
    /// an Excel UDF must return a value before returning. This bridge is the
    /// central design decision behind that gap. When AsyncWrites is false it is
    /// a pure synchronous passthrough: the caller's thread runs the work and
    /// sees its value or exception exactly as before. When it is true the work
    /// is handed to a per-endpoint serial queue and Excel-DNA's async machinery
    /// makes the calculation return Excel's pending marker (#N/A) immediately;
    /// the cell is filled with the real value (or the usual "Error: ..." text)
    /// when the queued item finishes.
    ///
    /// Ordering: every work item submitted for an endpoint is appended to that
    /// endpoint's FIFO queue, so writes to one NATS endpoint execute in
    /// submission order. Submission happens in NatsWriteObservable.Subscribe,
    /// which Excel-DNA invokes synchronously on the Excel thread during the
    /// internal RTD ConnectData of the Observe registration, so the queue order
    /// is the formula evaluation order. Different endpoints use different queues
    /// and can run concurrently. Idle queues are removed from the registry as
    /// soon as their last item completes. No thread-pool thread is held while an
    /// item is queued: the result is delivered by a continuation.
    ///
    /// Identity caveats: the calling cell reference is part of the identity, so
    /// inserting/moving rows or columns (coordinates change) makes the next
    /// evaluation a new call and re-issues the write. While the internal RTD
    /// topic stays connected, repeated evaluations with unchanged arguments
    /// return the cached completed value, so the write is not re-issued (this
    /// includes F9 on volatile write cells under AsyncWrites). When Excel
    /// instead detaches the topic (e.g. an unchanged recalculation on the
    /// observer skip path), the next evaluation re-registers the call and a
    /// volatile write is issued again: the dedup is best-effort per connected
    /// registration, not exactly-once across the sheet lifetime. Calls without
    /// a worksheet caller are refused with an Error cell because the identity
    /// would be shared or unstable.
    /// </summary>
    internal static class NatsUdfAsync
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // ExcelAsyncUtil identifies an async call by the (name, parameters)
        // pair. The prefix keeps these keys apart from other add-ins/wrappers
        // that happen to use the same UDF function names.
        private const string AsyncNamePrefix = "NatsUdfAsync:";

        private sealed class EndpointQueue
        {
            public readonly object Gate = new object();

            // Tail of the chained task sequence: the previous item, or a
            // completed sentinel for a brand-new queue.
            public Task<object> Tail = Task.FromResult<object>(null);
        }

        /// <summary>
        /// Bumped by <see cref="ResetAfterAddInReload"/>. A work item captures the
        /// generation when it is enqueued; if the add-in reloads before the item
        /// runs, the stale tail is skipped instead of executing the write again
        /// (Excel re-registers the observer and would otherwise issue it twice).
        /// </summary>
        private static int _generation;

        private static readonly ConcurrentDictionary<string, EndpointQueue> _queues =
            new ConcurrentDictionary<string, EndpointQueue>(StringComparer.Ordinal);

        /// <summary>
        /// Test-only override of <see cref="AppConfig.AsyncWrites"/> (null = use
        /// the process configuration). The offline unit tests pin this to false
        /// (sync-path tests) or true (the offline async cases: caller refusal
        /// and invalid-endpoint fallback) without touching Excel async.
        /// </summary>
#pragma warning disable 0649 // assigned only by the linked unit test sources
        internal static bool? AsyncWritesOverrideForTests;
#pragma warning restore 0649

        private static bool AsyncWritesEnabled => AsyncWritesOverrideForTests ?? AppConfig.AsyncWrites;

        /// <summary>
        /// Frozen entry point used by every NATS write UDF (and any read the
        /// user opts into async). In sync mode (AsyncWrites false) it returns
        /// <c>work()</c> directly on the caller thread - the exact pre-async
        /// behavior. In async mode it returns the object Excel expects for a
        /// pending async call.
        /// </summary>
        /// <param name="identityArgs">The UDF's own arguments, in their declared
        /// order; Excel-DNA uses them as the async call identity. They must be
        /// value-equal on the re-call that follows the completed async call, so
        /// the cached result is returned instead of running the write again.
        /// They must also be values Excel-DNA supports (null, scalars, strings,
        /// DateTime, numerics, enums, ExcelReference/ExcelError/ExcelEmpty/
        /// ExcelMissing and arrays of those). May be null (treated as empty).</param>
        internal static object Run(string functionName, object optionalHost, object[] identityArgs, Func<object> work)
        {
            // Sync mode: pure passthrough, the caller's thread runs the work.
            if (!AsyncWritesEnabled)
                return work();

            // Endpoint resolution is used for queueing only. A non-text host
            // argument, or a named reference that resolves to nothing, yields
            // null: fall back to the sync path so the dispatched Core body can
            // surface its usual "Error: ..." / "#CONFIG" text on the Excel
            // thread instead of hiding the cause behind an async registration.
            string endpoint = NatsRuntime.ResolveEndpointForDispatch(optionalHost);
            if (endpoint == null)
                return work();

            try
            {
                // The async name/parameters identify the registered call for
                // Excel. The name is stable per write function; the parameters
                // combine the calling cell reference (structurally equal on the
                // completed re-call, distinct across cells), the resolved NATS
                // endpoint and the UDF's own arguments (identityArgs) so a
                // changed argument is never served by an unrelated pending
                // call. Excel-DNA matches (name, parameters) by value and
                // returns the cached result on the re-call instead of
                // registering again.
                object caller = CallerIdentityForDedup();
                if (caller == null)
                {
                    // Without a worksheet cell reference the identity is shared
                    // across callers (or unstable between the call and the
                    // delivery re-call), which could skip or duplicate the
                    // write - fail loudly instead. Rare: a call outside
                    // worksheet evaluation (e.g. Application.Run) or a
                    // transient xlfCaller failure.
                    logger.Error($"NatsUdfAsync.Run: no worksheet caller for {functionName}; async write refused");
                    return NatsUdfCommon.Fail("async write needs a worksheet caller (AsyncWrites)");
                }
                object parameters = new object[] { caller, endpoint, identityArgs ?? new object[0] };
                string asyncName = AsyncNamePrefix + (functionName ?? string.Empty);

                // Observe overload with a custom one-shot observable:
                // Excel-DNA creates the observable at registration and invokes
                // its Subscribe synchronously on the Excel thread during the
                // internal RTD ConnectData, so the enqueue there is the single
                // side effect per registered call (the delivery re-call returns
                // the cached value from Excel-DNA's state lookup and never
                // re-subscribes). The queue continuation delivers
                // OnNext/OnCompleted, so no thread-pool thread waits while an
                // item is queued. The ExcelAsyncHandle overload (deferred
                // SetResult) crashed Excel under COM automation in the sibling
                // project, so it stays unused here too.
                object asyncResult = ExcelAsyncUtil.Observe(asyncName, parameters, () => new NatsWriteObservable(endpoint, work));
                if (asyncResult == null)
                {
                    // Excel-DNA returns null when its internal RTD registration
                    // failed. On some failure paths (array caller/xlUncalced)
                    // Subscribe may already have run, so an enqueued write still
                    // completes in order while the cell shows this error
                    // (at-least-once, same as the synchronous path); reporting
                    // the failure here cannot duplicate an enqueue.
                    logger.Error($"NatsUdfAsync.Run: async RTD registration failed for {functionName}");
                    return NatsUdfCommon.Fail("async dispatch failed (RTD registration)");
                }
                return asyncResult;
            }
            catch (Exception ex)
            {
                // A failure before Subscribe cannot duplicate a write; after it
                // the write is already enqueued and still completes in order
                // (at-least-once). Surface the failure instead of silently
                // retrying on the Excel thread.
                logger.Error(ex, $"NatsUdfAsync.Run: async dispatch failed for {functionName}");
                return NatsUdfCommon.Fail("async dispatch failed: " + SafeMessage(ex));
            }
        }

        /// <summary>
        /// Calling cell reference used as part of the async call identity: the
        /// same cell produces a structurally equal ExcelReference on the
        /// completed re-call, while different cells never share an identity.
        /// Returns null outside a worksheet call (or on a transient xlfCaller
        /// failure); Run refuses the async dispatch in that case because the
        /// identity could be shared or unstable.
        /// </summary>
        private static object CallerIdentityForDedup()
        {
            try
            {
                object caller = XlCall.Excel(XlCall.xlfCaller);
                if (caller is ExcelReference)
                    return caller;
            }
            catch
            {
                // xlfCaller is a live C API call and can transiently fail.
            }
            return null;
        }

        /// <summary>
        /// Appends a work item to the endpoint's FIFO queue and returns the task
        /// that completes with the item's result. The continuation chain is
        /// built under the queue's gate, so two submissions for one endpoint can
        /// never race; a stale queue instance (removed while idle) is detected
        /// after taking the gate and the item retries on the live instance. This
        /// guarantees FIFO order within a session; the only exception is
        /// <see cref="ResetAfterAddInReload"/>, which clears the map while a tail
        /// may still be in flight, so an endpoint can briefly run on two queues
        /// across an add-in reload (no write is lost or duplicated - the docs
        /// describe async dedup as best-effort).
        /// </summary>
        /// <summary>
        /// Enqueue a write for its endpoint and return a task that completes with
        /// the result. (The name has no "Async" suffix by contract: the caller is
        /// a synchronous UDF.)
        /// </summary>
        // VSTHRD200: the "Async" suffix is deliberately omitted; this name is the
        // frozen internal contract consumed by NatsWriteObservable (PLAN H6).
#pragma warning disable VSTHRD200 // Enqueue is the frozen internal name; see above.
        internal static Task<object> Enqueue(string endpoint, Func<object> work)
        {
            int generation = Volatile.Read(ref _generation);
            while (true)
            {
                EndpointQueue queue = _queues.GetOrAdd(endpoint, _ => new EndpointQueue());
                lock (queue.Gate)
                {
                    if (!_queues.TryGetValue(endpoint, out var current) || !ReferenceEquals(current, queue))
                        continue; // idle-removal race: retry against the live queue

                    Task<object> next = queue.Tail.ContinueWith(
                        _ => generation == Volatile.Read(ref _generation)
                            ? ExecuteSafely(work)
                            : SkippedByReload(),
                        CancellationToken.None,
                        // RunContinuationsAsynchronously is REQUIRED: a fresh or
                        // idle-released queue's Tail is an already-completed
                        // sentinel (Task.FromResult), and without this flag the
                        // continuation runs INLINE on the subscribing (Excel)
                        // thread - so a reply-dependent KV/JS put would block the
                        // UI thread for up to its response timeout, exactly the
                        // freeze AsyncWrites exists to remove.
                        TaskContinuationOptions.RunContinuationsAsynchronously,
                        TaskScheduler.Default);
                    queue.Tail = next;

                    // Drop the queue once its last item completed and no newer
                    // item claimed it (ReferenceEquals check under the gate).
                    // This continuation is scheduled after the work item, so a
                    // failure here can only leak an idle queue entry: Enqueue
                    // must never throw once the item is scheduled.
                    try
                    {
                        // VSTHRD110: this continuation only drops the idle queue
                        // entry; it is intentionally fire-and-forget.
                        _ = next.ContinueWith(
                            _ => ReleaseQueue(endpoint, queue, next),
                            CancellationToken.None,
                            TaskContinuationOptions.None,
                            TaskScheduler.Default);
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, "NatsUdfAsync.Enqueue: scheduling the idle-queue release failed");
                    }

                    return next;
                }
            }
        }
#pragma warning restore VSTHRD200

        private static void ReleaseQueue(string endpoint, EndpointQueue queue, Task<object> completed)
        {
            lock (queue.Gate)
            {
                // Remove only when this is still the dictionary's queue for the
                // endpoint and nothing newer claimed it (both checks under the
                // gate; a racing Enqueue retries against the live instance).
                if (ReferenceEquals(queue.Tail, completed)
                    && _queues.TryGetValue(endpoint, out var current)
                    && ReferenceEquals(current, queue))
                {
                    _queues.TryRemove(endpoint, out _);
                }
            }
        }

        /// <summary>Drops the per-endpoint queues on an add-in reload: the
        /// previous session's pending continuations must not be carried into the
        /// new runtime (async write identity is per-process). A later Enqueue
        /// builds a fresh queue; any still-running item finishes and its
        /// ReleaseQueue is a no-op (its endpoint no longer maps to that queue). An
        /// endpoint whose old tail is still in flight can therefore run on the old
        /// and the new queue briefly (FIFO is not guaranteed across the reload) -
        /// no write is lost or duplicated, only the strict per-endpoint
        /// serialization is suspended at the reload boundary.</summary>
        internal static void ResetAfterAddInReload()
        {
            _queues.Clear();
            // Invalidate any in-flight tail so it skips its work instead of
            // running it a second time when Excel re-registers the observer.
            Interlocked.Increment(ref _generation);
        }

        /// <summary>The result of a work item skipped because the add-in reloaded
        /// between its enqueue and its run.</summary>
        private static object SkippedByReload()
        {
            return NatsUdfCommon.Fail("superseded by add-in reload (the write was already issued; Excel re-registered the cell)");
        }

        /// <summary>
        /// Runs one queued work item. Exceptions must never surface through the
        /// async handle: they are logged and converted to the same "Error: ..."
        /// text the synchronous UDF bodies return.
        /// </summary>
        private static object ExecuteSafely(Func<object> work)
        {
            try
            {
                return work();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "NatsUdfAsync: queued work item failed");
                return NatsUdfCommon.Fail(SafeMessage(ex));
            }
        }

        /// <summary>
        /// The exception text handed to a cell, defensively: the whole
        /// InnerException chain is flattened (a faulted Task carries an
        /// AggregateException, whose wrapper message "One or more errors
        /// occurred." alone is useless) and the joined text is masked with
        /// <see cref="AppConfig.MaskNatsEndpoint"/> because a NATS exception
        /// message can embed a connection URL (nats://user:password@host) or a
        /// token. A throwing Message getter must never turn the error conversion
        /// itself into a faulted task (the observable would then have no terminal
        /// signal for the cell).
        /// </summary>
        internal static string SafeMessage(Exception ex)
        {
            try
            {
                if (ex == null)
                    return "async write failed";

                // Breadth-first over the chain, with a depth cap: the chain is
                // data we did not build, so a pathological / cyclic graph must
                // not spin here.
                var text = new StringBuilder();
                var pending = new Queue<Exception>();
                pending.Enqueue(ex);
                int depth = 0;
                while (pending.Count > 0 && depth < 16)
                {
                    Exception current = pending.Dequeue();
                    if (current == null)
                        continue;
                    depth++;
                    if (current is AggregateException aggregate)
                    {
                        // Skip the wrapper message: it adds no information over
                        // the inner failures.
                        foreach (Exception inner in aggregate.InnerExceptions)
                            pending.Enqueue(inner);
                        continue;
                    }
                    if (text.Length > 0)
                        text.Append(" -> ");
                    text.Append(current.Message);
                    if (current.InnerException != null)
                        pending.Enqueue(current.InnerException);
                }

                string joined = text.ToString();
                return string.IsNullOrEmpty(joined) ? "async write failed" : AppConfig.MaskNatsEndpoint(joined);
            }
            catch
            {
                return "async write failed";
            }
        }

        /// <summary>Whether a queue currently exists for the endpoint; test-only
        /// (the drain/order tests assert the idle release).</summary>
        internal static bool HasQueueForTests(string endpoint) => _queues.ContainsKey(endpoint);

        /// <summary>
        /// One-shot Excel-DNA observable for an async write. Subscribe runs
        /// synchronously on the Excel thread during the internal RTD ConnectData
        /// (inside the ExcelAsyncUtil.Observe registration), so the per-endpoint
        /// enqueue happens there and the per-endpoint FIFO order is the formula
        /// evaluation order. The queue continuation delivers the result (or the
        /// usual "Error: ..." text) and completion, so no thread-pool thread
        /// waits while the item is queued (DESIGN.md §5.3). Excel cancelling a
        /// calculation or detaching the topic only disposes this subscription;
        /// the returned disposable is a no-op, so a queued or executing write
        /// still runs to completion in per-endpoint order.
        /// A duplicate Subscribe on the same instance never enqueues the write
        /// again and is still delivered the same queued result (defensive path;
        /// Excel-DNA registers one observable per call).
        /// </summary>
        internal sealed class NatsWriteObservable : IExcelObservable
        {
            private static readonly Logger logger = LogManager.GetCurrentClassLogger();

            // Every Subscribe returns this shared instance: Dispose must be a
            // no-op and the object carries no per-subscription state.
            private static readonly IDisposable NoOpDisposable = new NoOpDisposableImpl();

            private sealed class NoOpDisposableImpl : IDisposable
            {
                public void Dispose()
                {
                }
            }

            private readonly string _endpoint;
            private readonly Func<object> _work;

            // One-shot enqueue guard: Excel-DNA never subscribes twice for one
            // registered call, but a duplicate Subscribe must never enqueue the
            // write again (Excel-DNA's own ThreadPoolDelegateObservable threw in
            // that case). The latch publishes the queued write task (never its
            // value) before any delivery, so a duplicate subscription never blocks
            // the calling (Excel) thread: it attaches asynchronously and receives
            // the very same result, including an enqueue failure.
            private int _subscribed;
            private readonly TaskCompletionSource<Task<object>> _queuedReady =
                new TaskCompletionSource<Task<object>>(TaskCreationOptions.RunContinuationsAsynchronously);

            internal NatsWriteObservable(string endpoint, Func<object> work)
            {
                _endpoint = endpoint;
                _work = work;
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                if (Interlocked.CompareExchange(ref _subscribed, 1, 0) == 0)
                {
                    Task<object> queued;
                    try
                    {
                        // Enqueue synchronously HERE: Excel-DNA calls Subscribe
                        // during the RTD ConnectData on the Excel thread, so
                        // same-endpoint writes enter the FIFO in formula
                        // evaluation order. The delivery runs only when the item
                        // reaches the head of the queue.
                        queued = NatsUdfAsync.Enqueue(_endpoint, _work);
                    }
                    catch (Exception ex)
                    {
                        // Enqueue converts work failures to "Error: ..." task
                        // results and never throws once it schedules the item;
                        // this guards truly unexpected failures. Publish the
                        // failure through the same latch so a duplicate Subscribe
                        // gets the identical result instantly instead of a
                        // spin + error.
                        logger.Error(ex, "NatsWriteObservable: enqueue failed");
                        queued = Task.FromResult<object>(NatsUdfCommon.Fail(NatsUdfAsync.SafeMessage(ex)));
                    }
                    _queuedReady.TrySetResult(queued);
                    AttachDelivery(queued, observer);
                    return NoOpDisposable;
                }

                // Duplicate Subscribe for an already-registered call: never
                // enqueue again, and never block the calling (Excel) thread:
                // wait for the first Subscribe to publish the queued task
                // asynchronously and deliver the very same result to this
                // observer.
                logger.Warn("NatsWriteObservable: duplicate Subscribe; the same queued result is delivered");
                // Delivery is dispatched to the thread pool, never run inline on
                // the calling (Excel) thread: the observer is notified
                // asynchronously. VSTHRD110: deliberately fire-and-forget;
                // DeliverWhenReadyAsync observes every failure itself.
                _ = Task.Run(() =>
                {
                    // Discard is deliberate: DeliverWhenReadyAsync observes every
                    // failure itself (VSTHRD110 accepts a discard).
                    _ = DeliverWhenReadyAsync(_queuedReady.Task, observer);
                });
                return NoOpDisposable;
            }

            private static void AttachDelivery(Task<object> queued, IExcelObserver observer)
            {
                // Same rule as above: delivery runs on the thread pool, not on the
                // thread that happens to complete the queued task (which may be the
                // Excel thread for a synchronous first write).
                // VSTHRD110: deliberately fire-and-forget; see DeliverQueuedAsync.
                _ = Task.Run(() =>
                {
                    _ = DeliverQueuedAsync(queued, observer);
                });
            }

            /// <summary>
            /// Waits asynchronously for the first Subscribe to publish the queued
            /// write task, then delivers it to this (duplicate) observer. Runs as
            /// a thread-pool continuation, never on the Excel thread.
            /// </summary>
            // VSTHRD003: the queued/ready tasks are started on the thread pool by
            // Enqueue; this method only ever runs as a thread-pool continuation
            // (ConfigureAwait(false)), so there is no synchronization context to
            // deadlock on.
#pragma warning disable VSTHRD003 // cross-context await is safe off the Excel thread; see above.
            private static async Task DeliverWhenReadyAsync(Task<Task<object>> ready, IExcelObserver observer)
            {
                try
                {
                    Task<object> queued = await ready.ConfigureAwait(false);
                    await DeliverQueuedAsync(queued, observer).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "NatsWriteObservable: waiting for the queued write failed");
                    DeliverAndComplete(observer, NatsUdfCommon.Fail(NatsUdfAsync.SafeMessage(ex)));
                }
            }
#pragma warning restore VSTHRD003

            /// <summary>
            /// Maps the queued task's outcome to the value Excel should show.
            /// ExecuteSafely already converts a work failure into "Error: ..."
            /// text; a cancelled/faulted task is handled defensively. The
            /// <c>await</c> is the non-blocking result read: this method only runs
            /// as a thread-pool continuation, never on the Excel thread.
            /// </summary>
            // VSTHRD003: same as DeliverWhenReadyAsync - the queued task is
            // thread-pool work and this continuation has no synchronization
            // context, so the cross-context await cannot deadlock.
#pragma warning disable VSTHRD003 // cross-context await is safe off the Excel thread; see above.
            private static async Task DeliverQueuedAsync(Task<object> queued, IExcelObserver observer)
            {
                object result;
                try
                {
                    result = await queued.ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    // ExecuteSafely converts a work failure into "Error: ..." text,
                    // so a cancellation here is a defensive terminal signal.
                    result = NatsUdfCommon.Fail("async write was cancelled");
                }
                catch (Exception ex)
                {
                    // Computing the result must never strand a cell: a
                    // pathological fault (an exception whose Message getter
                    // itself throws) would skip the delivery entirely and leave
                    // Excel pending forever.
                    logger.Error(ex, "NatsWriteObservable: computing the async write result failed");
                    result = NatsUdfCommon.Fail(NatsUdfAsync.SafeMessage(ex));
                }
                DeliverAndComplete(observer, result);
            }
#pragma warning restore VSTHRD003

            /// <summary>
            /// Delivers one value and then completes, never throwing into the
            /// queue continuation: Excel may already have detached the topic, and
            /// the continuation must still clear the queue slot. Failures are
            /// logged; completion is attempted even when the delivery failed.
            /// </summary>
            private static void DeliverAndComplete(IExcelObserver observer, object result)
            {
                TryNotify(() => observer.OnNext(result), "NatsWriteObservable: delivering the async write result failed");
                TryNotify(observer.OnCompleted, "NatsWriteObservable: completing the async write observer failed");
            }

            private static void TryNotify(Action action, string failure)
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, failure);
                }
            }
        }
    }
}
