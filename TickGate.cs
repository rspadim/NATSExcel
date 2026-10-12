using System;
using System.Threading;

namespace NATSExcel
{
    /// <summary>
    /// Non-blocking reentrancy gate for timer callbacks: the first caller enters,
    /// concurrent callers are rejected until Exit() is called.
    /// <para>
    /// A skipped tick is the correct behavior: the timer fires again on the next
    /// period anyway, and the data read at that point is newer than anything a
    /// queued (overlapping) callback could have produced. Letting a slow tick
    /// overlap the next one would stack NATS work and Excel flushes on the thread
    /// pool, flood the server and race the RTD topic state. Dropping the tick
    /// keeps at most one callback in flight, bounds the backlog to one period and
    /// never blocks the timer thread. Never queue a rejected tick.
    /// </para>
    /// </summary>
    internal sealed class TickGate
    {
        private int _busy;

        /// <summary>TRUE when the gate was free and is now held by the caller; FALSE means another callback is still running and this tick must be skipped.</summary>
        public bool TryEnter()
        {
            return Interlocked.CompareExchange(ref _busy, 1, 0) == 0;
        }

        /// <summary>
        /// Releases the gate. Only call it after a successful TryEnter by the
        /// current holder: calling Exit while another owner holds the gate would
        /// reopen the section early.
        /// </summary>
        public void Exit()
        {
            Interlocked.Exchange(ref _busy, 0);
        }
    }

    /// <summary>
    /// Fixed set of lock stripes selected by key hash: callers serialize
    /// per-key work without a lock per key. The stripe count is fixed at
    /// construction, so the same key always maps to the same stripe.
    /// </summary>
    internal sealed class StripedLocks
    {
        private readonly object[] _stripes;

        public StripedLocks(int stripeCount)
        {
            if (stripeCount < 1)
                stripeCount = 1;

            _stripes = new object[stripeCount];
            for (int i = 0; i < _stripes.Length; i++)
                _stripes[i] = new object();
        }

        /// <summary>
        /// Stripe for the key: its ordinal hash masked non-negative, modulo the
        /// count. A null or empty key maps to the first stripe instead of throwing.
        /// </summary>
        public object For(string key)
        {
            if (string.IsNullOrEmpty(key))
                return _stripes[0];

            return _stripes[(StringComparer.Ordinal.GetHashCode(key) & 0x7FFFFFFF) % _stripes.Length];
        }
    }
}
