using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// <see cref="NatsSubscriptionManager.MakeKey"/> is the pure, internal identity
    /// function behind the ref-counted Core subscriptions: one physical wire
    /// subscription per <c>(endpoint, subject, queueGroup)</c>, so two distinct
    /// triples must never share a registry key (DESIGN.md §5.1, PLAN F2.10-F2.14).
    ///
    /// The origin counters (<c>ListenerCountWithOrigin</c>/<c>SubjectCountWithOrigin</c>)
    /// have no pure helper to exercise offline: they walk live <c>_states</c>, and
    /// the only way to add state is <c>Subscribe</c>, which starts a consume loop
    /// that opens a socket. Those counters stay covered by the smoke/liveness
    /// layers; this file pins the key contract, which is where the ref-counting
    /// actually lives.
    /// </summary>
    public class NatsSubscriptionKeyTests
    {
        [Fact]
        public void SameInputs_ProduceTheSameKey()
        {
            Assert.Equal(
                NatsSubscriptionManager.MakeKey("nats://h:4222", "orders.*", "workers"),
                NatsSubscriptionManager.MakeKey("nats://h:4222", "orders.*", "workers"));
        }

        [Fact]
        public void Key_FollowsTheDocumentedLengthPrefixedFormat()
        {
            // {endpoint.Length}:{endpoint}\u0001{subject.Length}:{subject}\u0001{queue.Length}:{queue}
            Assert.Equal(
                "2:ep\u00013:sub\u00012:qg",
                NatsSubscriptionManager.MakeKey("ep", "sub", "qg"));
        }

        [Fact]
        public void NullParts_BecomeEmptyParts()
        {
            Assert.Equal("0:\u00010:\u00010:", NatsSubscriptionManager.MakeKey(null, null, null));
        }

        [Fact]
        public void ChangingAnyComponent_ChangesTheKey()
        {
            string baseline = NatsSubscriptionManager.MakeKey("ep", "sub", "qg");
            Assert.NotEqual(baseline, NatsSubscriptionManager.MakeKey("ep2", "sub", "qg"));
            Assert.NotEqual(baseline, NatsSubscriptionManager.MakeKey("ep", "sub2", "qg"));
            Assert.NotEqual(baseline, NatsSubscriptionManager.MakeKey("ep", "sub", "qg2"));
        }

        [Fact]
        public void PlainAndQueueGroupSubscribers_AreDifferentKeys()
        {
            Assert.NotEqual(
                NatsSubscriptionManager.MakeKey("ep", "sub", null),
                NatsSubscriptionManager.MakeKey("ep", "sub", "workers"));
        }

        [Fact]
        public void NullAndEmptyQueueGroup_AreTheSameKey()
        {
            // Subscribe normalizes a blank queue group to null before calling
            // MakeKey, so null and "" must describe the same physical subscription.
            Assert.Equal(
                NatsSubscriptionManager.MakeKey("ep", "sub", null),
                NatsSubscriptionManager.MakeKey("ep", "sub", string.Empty));
        }

        [Fact]
        public void ComponentsContainingTheSeparatorOrColon_CannotCollideWithADifferentSplit()
        {
            // "\u0001" is the field separator and ':' is part of the prefix. The
            // length prefix is what keeps two different triples apart even when a
            // component carries either character (or shifts a boundary).
            Assert.NotEqual(
                NatsSubscriptionManager.MakeKey("a:b", "c", string.Empty),
                NatsSubscriptionManager.MakeKey("a", "b:c", string.Empty));
            Assert.NotEqual(
                NatsSubscriptionManager.MakeKey("a\u0001b", "c", string.Empty),
                NatsSubscriptionManager.MakeKey("a", "b\u0001c", string.Empty));
            Assert.NotEqual(
                NatsSubscriptionManager.MakeKey("ab", "c", string.Empty),
                NatsSubscriptionManager.MakeKey("a", "bc", string.Empty));
        }
    }
}
