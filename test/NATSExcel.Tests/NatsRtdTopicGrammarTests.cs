using System;
using System.Collections.Generic;
using System.Globalization;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The frozen RTD topic grammar (PLAN.md §1.2, DESIGN.md §6): every one of
    /// the nine commands parses to its kind with the documented positional
    /// arguments, malformed topics are rejected before they can be registered,
    /// and <see cref="NatsRtdTopicSpec.Parse"/> never throws.
    /// </summary>
    public class NatsRtdTopicGrammarTests
    {
        /// <summary>
        /// Canonical RTD shape: the command is repeated at <c>topicInfo[0]</c>, so
        /// the parser only reads the value tokens positionally.
        /// </summary>
        private static NatsRtdTopicSpec Parse(string command, params string[] topicInfo)
        {
            return NatsRtdTopicSpec.Parse(command, topicInfo);
        }

        // ------------------------------------------------------------------
        // Every frozen command parses to its kind.
        // ------------------------------------------------------------------

        [Fact]
        public void Sub_ParsesSubjectQueueGroupAndEndpoint()
        {
            var spec = Parse("SUB", "SUB", "orders.*", "workers", "nats://h:4222");
            Assert.Equal(NatsRtdKind.Sub, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("SUB", spec.Command);
            Assert.Equal("orders.*", spec.Arg1);
            Assert.Equal("workers", spec.Arg2);
            Assert.Equal("nats://h:4222", spec.Arg3);
            Assert.Equal("nats://h:4222", spec.Endpoint); // SUB endpoint is position 3
        }

        [Fact]
        public void Kv_ParsesBucketKeyAndEndpoint()
        {
            var spec = Parse("KV", "KV", "bucket", "key", "nats://h:4222");
            Assert.Equal(NatsRtdKind.Kv, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("bucket", spec.Arg1);
            Assert.Equal("key", spec.Arg2);
            Assert.Equal("nats://h:4222", spec.Arg3);
            Assert.Equal("nats://h:4222", spec.Endpoint);
        }

        [Fact]
        public void KvGet_ParsesBucketKeyAndEndpoint()
        {
            var spec = Parse("KVGET", "KVGET", "bucket", "key", "nats://h:4222");
            Assert.Equal(NatsRtdKind.KvGet, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("bucket", spec.Arg1);
            Assert.Equal("key", spec.Arg2);
            Assert.Equal("nats://h:4222", spec.Arg3);
            Assert.Equal("nats://h:4222", spec.Endpoint);
        }

        [Fact]
        public void KvKeys_ParsesBucketAndEndpointAtPosition2()
        {
            var spec = Parse("KVKEYS", "KVKEYS", "bucket", "nats://h:4222");
            Assert.Equal(NatsRtdKind.KvKeys, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("bucket", spec.Arg1);
            Assert.Equal("nats://h:4222", spec.Arg2);
            Assert.Null(spec.Arg3);
            Assert.Equal("nats://h:4222", spec.Endpoint); // KVKEYS endpoint is position 2
        }

        [Fact]
        public void Req_ParsesSubjectPayloadEveryMsAndEndpointAtPosition4()
        {
            var spec = Parse("REQ", "REQ", "svc", "ping", "100", "nats://h:4222");
            Assert.Equal(NatsRtdKind.Req, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("svc", spec.Arg1);
            Assert.Equal("ping", spec.Arg2);
            Assert.Equal("100", spec.Arg3);
            Assert.Equal("nats://h:4222", spec.Arg4);
            Assert.Equal(100, spec.EveryMs);
            Assert.Equal("nats://h:4222", spec.Endpoint); // REQ endpoint is position 4
        }

        [Fact]
        public void Js_ParsesStreamConsumerFilterAndEndpointAtPosition4()
        {
            var spec = Parse("JS", "JS", "stream", "consumer", "filter", "nats://h:4222");
            Assert.Equal(NatsRtdKind.Js, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("stream", spec.Arg1);
            Assert.Equal("consumer", spec.Arg2);
            Assert.Equal("filter", spec.Arg3);
            Assert.Equal("nats://h:4222", spec.Arg4);
            Assert.Equal("nats://h:4222", spec.Endpoint);
        }

        [Fact]
        public void Msg_ParsesStreamFilterAndEndpoint()
        {
            var spec = Parse("MSG", "MSG", "stream", "filter", "nats://h:4222");
            Assert.Equal(NatsRtdKind.Msg, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("stream", spec.Arg1);
            Assert.Equal("filter", spec.Arg2);
            Assert.Equal("nats://h:4222", spec.Arg3);
            Assert.Equal("nats://h:4222", spec.Endpoint);
        }

        [Fact]
        public void Obj_ParsesBucketObjectAndEndpoint()
        {
            var spec = Parse("OBJ", "OBJ", "bucket", "object", "nats://h:4222");
            Assert.Equal(NatsRtdKind.Obj, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("bucket", spec.Arg1);
            Assert.Equal("object", spec.Arg2);
            Assert.Equal("nats://h:4222", spec.Arg3);
            Assert.Equal("nats://h:4222", spec.Endpoint);
        }

        [Fact]
        public void RtdCluster_ParsesWhatAndEndpointAtPosition2()
        {
            var spec = Parse("RTDCLUSTER", "RTDCLUSTER", "leader", "nats://h:4222");
            Assert.Equal(NatsRtdKind.Cluster, spec.Kind);
            Assert.True(spec.IsValid);
            Assert.Equal("leader", spec.Arg1);
            Assert.Equal("nats://h:4222", spec.Arg2);
            Assert.Equal("nats://h:4222", spec.Endpoint); // RTDCLUSTER endpoint is position 2
        }

        [Fact]
        public void CommandToken_IsTrimmedAndUpperCased()
        {
            var spec = Parse("  sub  ", "orders");
            Assert.Equal(NatsRtdKind.Sub, spec.Kind);
            Assert.Equal("SUB", spec.Command);
            Assert.Equal("orders", spec.Arg1);
        }

        [Fact]
        public void ValueArguments_KeepTheirSurroundingWhitespace()
        {
            // Whitespace is significant in NATS: "orders" and "orders " differ.
            var spec = Parse("SUB", " orders ");
            Assert.Equal(NatsRtdKind.Sub, spec.Kind);
            Assert.Equal(" orders ", spec.Arg1);
        }

        // ------------------------------------------------------------------
        // Unknown commands.
        // ------------------------------------------------------------------

        [Fact]
        public void UnknownCommand_ReturnsUnknownKindAndTheFrozenCommandList()
        {
            var spec = Parse("bogus", "bogus");
            Assert.Equal(NatsRtdKind.Unknown, spec.Kind);
            Assert.False(spec.IsValid);
            Assert.Equal("BOGUS", spec.Command);
            Assert.Equal(
                "unknown command 'BOGUS', expected one of [SUB, KV, KVGET, KVKEYS, REQ, JS, MSG, OBJ, RTDCLUSTER]",
                spec.Error);
        }

        // ------------------------------------------------------------------
        // Arity: only a NON-BLANK token past the documented arity is an error.
        // ------------------------------------------------------------------

        [Fact]
        public void NonBlankExtraToken_IsRejectedWithItsTopicPosition()
        {
            var spec = Parse("SUB", "SUB", "orders", "workers", "nats://h", "surprise");
            Assert.False(spec.IsValid);
            Assert.Equal(NatsRtdKind.Unknown, spec.Kind);
            Assert.Equal(
                "SUB accepts at most 3 argument(s); unexpected extra argument 'surprise' at topic position 5",
                spec.Error);
        }

        [Fact]
        public void BlankTrailingToken_IsNotAnError()
        {
            // Excel leaves trailing empty arguments on =RTD(...) calls.
            var spec = Parse("SUB", "SUB", "orders", "", "");
            Assert.True(spec.IsValid);
            Assert.Equal("orders", spec.Arg1);
        }

        // ------------------------------------------------------------------
        // A blank required value is rejected before the extra-argument check.
        // ------------------------------------------------------------------

        [Fact]
        public void BlankRequiredValue_IsRejectedBeforeTheExtraArgumentCheck()
        {
            var spec = Parse("SUB", "SUB", "", "workers", "nats://h", "surprise");
            Assert.False(spec.IsValid);
            Assert.Equal("SUB requires a subject as the second argument", spec.Error);
        }

        public static TheoryData<string, string[], string> BlankRequiredCases =>
            new TheoryData<string, string[], string>
            {
                { "SUB", new[] { "SUB", "" }, "SUB requires a subject as the second argument" },
                { "KV", new[] { "KV", "" }, "KV requires a bucket as the second argument" },
                { "KV", new[] { "KV", "bucket", "" }, "KV requires a key as the third argument" },
                { "KVGET", new[] { "KVGET", "bucket", "" }, "KVGET requires a key as the third argument" },
                { "KVKEYS", new[] { "KVKEYS", "" }, "KVKEYS requires a bucket as the second argument" },
                { "REQ", new[] { "REQ", "" }, "REQ requires a subject as the second argument" },
                { "JS", new[] { "JS", "" }, "JS requires a stream as the second argument" },
                { "JS", new[] { "JS", "stream", "" }, "JS requires a consumer as the third argument" },
                { "MSG", new[] { "MSG", "" }, "MSG requires a stream as the second argument" },
                // The article is deliberately always "a" (RequiredMessage XML
                // doc): the template stays mechanical, so "object" keeps "a".
                { "OBJ", new[] { "OBJ", "bucket", "" }, "OBJ requires a object as the third argument" },
            };

        [Theory]
        [MemberData(nameof(BlankRequiredCases))]
        public void BlankRequiredValue_IsRejectedWithTheSpecificMessage(
            string command, string[] topic, string expected)
        {
            var spec = NatsRtdTopicSpec.Parse(command, topic);
            Assert.False(spec.IsValid);
            Assert.Equal(NatsRtdKind.Unknown, spec.Kind);
            Assert.Equal(expected, spec.Error);
        }

        // ------------------------------------------------------------------
        // REQ everyMs is clamped to >= 100.
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("0", 100)]
        [InlineData("-5", 100)]
        [InlineData("50", 100)]
        [InlineData("100", 100)]
        [InlineData("250", 250)]
        [InlineData("9999999999", int.MaxValue)]
        public void ReqEveryMs_IsClampedToTheMinimumAndExposedAsInvariantText(string raw, int expected)
        {
            var spec = Parse("REQ", "REQ", "svc", "payload", raw);
            Assert.True(spec.IsValid);
            Assert.Equal(expected, spec.EveryMs);
            Assert.Equal(expected.ToString(CultureInfo.InvariantCulture), spec.Arg3);
            Assert.Null(spec.Endpoint); // no endpoint argument supplied
        }

        [Theory]
        [InlineData("")]
        [InlineData("1.5")]
        [InlineData("abc")]
        public void ReqEveryMs_NonInteger_IsRejected(string raw)
        {
            var spec = Parse("REQ", "REQ", "svc", "payload", raw);
            Assert.False(spec.IsValid);
            Assert.Equal("REQ requires everyMs as a positive number of milliseconds", spec.Error);
        }

        // ------------------------------------------------------------------
        // The parser never throws on malformed input.
        // ------------------------------------------------------------------

        public static TheoryData<string, string[]> MalformedInputs =>
            new TheoryData<string, string[]>
            {
                { "", new string[0] },
                { "   ", new[] { "   " } },
                { "*", new[] { "*" } },
                { ">", new[] { ">" } },
                { "orders.*", new[] { "orders.*" } },
                { new string('A', 200000), null },
                { "a.b", new[] { "a.b" } },
                { "SUB", new[] { "SUB", null } },
                { "SUB", new[] { "   " } },
                { "KV", new[] { "KV", null, null } },
            };

        [Theory]
        [MemberData(nameof(MalformedInputs))]
        public void MalformedInput_NeverThrowsAndReturnsAnError(string command, string[] topic)
        {
            NatsRtdTopicSpec spec = null;
            Exception thrown = Record.Exception(() => { spec = NatsRtdTopicSpec.Parse(command, topic); });

            Assert.Null(thrown);
            Assert.NotNull(spec);
            Assert.False(spec.IsValid);
            Assert.False(string.IsNullOrEmpty(spec.Error));
        }

        [Fact]
        public void NullCommandAndNullList_DoesNotThrow()
        {
            NatsRtdTopicSpec spec = null;
            Exception thrown = Record.Exception(() => { spec = NatsRtdTopicSpec.Parse(null, null); });

            Assert.Null(thrown);
            Assert.NotNull(spec);
            Assert.Equal(NatsRtdKind.Unknown, spec.Kind);
            Assert.Equal(
                "unknown command '', expected one of [SUB, KV, KVGET, KVKEYS, REQ, JS, MSG, OBJ, RTDCLUSTER]",
                spec.Error);
        }

        [Fact]
        public void CommandListConstant_MatchesTheErrorText()
        {
            var spec = Parse("nope", "nope");
            Assert.Equal(
                "unknown command 'NOPE', expected one of " + NatsRtdTopicSpec.CommandList,
                spec.Error);
            Assert.Equal("[SUB, KV, KVGET, KVKEYS, REQ, JS, MSG, OBJ, RTDCLUSTER]", NatsRtdTopicSpec.CommandList);
        }

        [Fact]
        public void MinEveryMs_IsHundred()
        {
            Assert.Equal(100, NatsRtdTopicSpec.MinEveryMs);
        }

        // ------------------------------------------------------------------
        // The optional-middle positional rule (ResolveOptionalMiddle). It is the
        // only place where a token's POSITION can change meaning, and the Excel E2E
        // never reaches it: every E2E cell fills the slots explicitly, so a
        // regression here would move a queue group into the endpoint slot (or the
        // reverse) with no layer noticing.
        //
        // Note on the assertions: the parser sizes `slots` to the command's arity
        // and leaves an unsupplied slot null, while a slot the user left as "" stays
        // "". Both are blank to the parser (IsBlank is IsNullOrWhiteSpace), so a
        // "nothing here" assertion is IsNullOrEmpty - not Assert.Null, which would
        // be pinning an incidental difference between two equivalent topics.
        // ------------------------------------------------------------------

        [Fact]
        public void Sub_BlankMiddle_ATokenWithASpaceIsTheQueueGroupNotAnEndpoint()
        {
            // LooksLikeEndpoint: no "://" and it carries whitespace, so the trailing
            // token cannot be an alias - it is the middle (queue group) value.
            var spec = Parse("SUB", "SUB", "orders", "", "my group");

            Assert.True(spec.IsValid);
            Assert.Equal("orders", spec.Arg1);
            Assert.Equal("my group", spec.Arg2);
            Assert.True(string.IsNullOrEmpty(spec.Arg3), "the endpoint slot must be released");
            Assert.True(string.IsNullOrEmpty(spec.Endpoint), "a queue group is not an endpoint");
        }

        [Fact]
        public void Sub_BlankMiddle_ASingleTokenIsTheEndpointNotTheQueueGroup()
        {
            // LooksLikeEndpoint: a single non-blank token with no whitespace is the
            // config-free stand-in for a Servers alias, so it stays the endpoint.
            var spec = Parse("SUB", "SUB", "orders", "", "srvB");

            Assert.True(spec.IsValid);
            Assert.Equal("orders", spec.Arg1);
            Assert.True(string.IsNullOrEmpty(spec.Arg2), "an alias is not a queue group");
            Assert.Equal("srvB", spec.Arg3);
            Assert.Equal("srvB", spec.Endpoint);
        }

        [Fact]
        public void Sub_ABlankOrFilledMiddleNeverMovesAToken()
        {
            // The rule fires only when the middle is blank AND the endpoint is not:
            // SUB(subject, "workers") stays a queue group (a two-value call is not
            // an endpoint argument), and a fully blank tail stays blank.
            var twoValues = Parse("SUB", "SUB", "orders", "workers");
            Assert.Equal("workers", twoValues.Arg2);
            Assert.True(string.IsNullOrEmpty(twoValues.Arg3));

            var blankTail = Parse("SUB", "SUB", "orders", "", "");
            Assert.True(string.IsNullOrEmpty(blankTail.Arg2));
            Assert.True(string.IsNullOrEmpty(blankTail.Arg3));
        }

        [Fact]
        public void Sub_AUrlTokenIsAlwaysTheEndpoint()
        {
            var spec = Parse("SUB", "SUB", "orders", "", "nats://h:4222");
            Assert.True(string.IsNullOrEmpty(spec.Arg2));
            Assert.Equal("nats://h:4222", spec.Arg3);
        }

        [Fact]
        public void Msg_BlankMiddle_UsesTheSameRule()
        {
            // MSG[1]stream [2]filter [3]endpoint: the same helper at (1, 2).
            var spaced = Parse("MSG", "MSG", "stream", "", "my filter");
            Assert.Equal("my filter", spaced.Arg2);
            Assert.True(string.IsNullOrEmpty(spaced.Arg3));

            var alias = Parse("MSG", "MSG", "stream", "", "srvB");
            Assert.True(string.IsNullOrEmpty(alias.Arg2));
            Assert.Equal("srvB", alias.Arg3);
        }

        [Fact]
        public void Js_BlankMiddle_UsesTheSameRuleAtItsOwnSlots()
        {
            // JS[1]stream [2]consumer [3]filter [4]endpoint: the helper runs at
            // (2, 3), so a blank filter with a spaced token keeps it as the filter.
            var spaced = Parse("JS", "JS", "stream", "consumer", "", "my filter");
            Assert.Equal("consumer", spaced.Arg2);
            Assert.Equal("my filter", spaced.Arg3);
            Assert.True(string.IsNullOrEmpty(spaced.Arg4));

            var alias = Parse("JS", "JS", "stream", "consumer", "", "srvB");
            Assert.True(string.IsNullOrEmpty(alias.Arg3));
            Assert.Equal("srvB", alias.Arg4);

            // A FILLED filter is never moved: JS(stream, consumer, filter) keeps the
            // filter in position 3 even though a bare token looks like an alias.
            var filled = Parse("JS", "JS", "stream", "consumer", "orders.>");
            Assert.Equal("orders.>", filled.Arg3);
            Assert.True(string.IsNullOrEmpty(filled.Arg4));
        }
    }
}
