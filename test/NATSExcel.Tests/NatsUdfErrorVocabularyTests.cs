using System;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins the frozen error vocabulary (PLAN.md §1.3): the exact marker/error
    /// constants, the <c>Fail</c> shape, the <c>VerboseError</c> shape
    /// (<c>Error: &lt;context&gt;: &lt;cause&gt; | &lt;hint&gt;</c>), the credential
    /// masking applied to every part, and the named-cause builders. Every value
    /// here is read from <see cref="NatsUdfCommon"/> at test time (never copied
    /// from memory) so a reworded constant fails this suite.
    /// </summary>
    public class NatsUdfErrorVocabularyTests
    {
        // ------------------------------------------------------------------
        // The marker / error constants are the contract.
        // ------------------------------------------------------------------

        [Fact]
        public void MarkerConstants_AreTheExactDocumentedStrings()
        {
            Assert.Equal("(ConnectData)", NatsUdfCommon.ConnectData);
            Assert.Equal("(subscribed)", NatsUdfCommon.Subscribed);
            Assert.Equal("(no value)", NatsUdfCommon.NoValue);
            Assert.Equal("(null)", NatsUdfCommon.NullValue);
            Assert.Equal("No change", NatsUdfCommon.NoChange);
            Assert.Equal("OK", NatsUdfCommon.Ok);
            Assert.Equal("OK (fire and forget)", NatsUdfCommon.FireAndForget);
            Assert.Equal("OK (fire and forget: all)", NatsUdfCommon.FireAndForgetAll);
            Assert.Equal("TIMEOUT", NatsUdfCommon.Timeout);
            Assert.Equal("No Responders", NatsUdfCommon.NoResponders);
            Assert.Equal("#CONFIG", NatsUdfCommon.ConfigError);
        }

        [Fact]
        public void ErrorConstants_AreTheExactDocumentedStrings()
        {
            Assert.Equal("Error: ", NatsUdfCommon.ErrorPrefix);
            Assert.Equal(
                "Excel error cells are not valid arguments",
                NatsUdfCommon.ErrorExcelErrorCell);
            Assert.Equal(
                "A multi-cell range is not a valid scalar argument",
                NatsUdfCommon.ErrorMultiCellScalar);
            Assert.Equal(
                "expected a range with 2 columns or 2 rows",
                NatsUdfCommon.ErrorRangeShape);
            Assert.Equal(
                "keys and values must have the same number of cells",
                NatsUdfCommon.ErrorCellCountMismatch);
            Assert.Equal("no entries to write", NatsUdfCommon.ErrorNoEntries);
            Assert.Equal(
                "JetStream is not enabled on this server",
                NatsUdfCommon.ErrorJetStreamDisabled);
            Assert.Equal(
                "admin writes are disabled (set \"AllowAdminWrites\": true)",
                NatsUdfCommon.ErrorAdminWritesDisabled);
        }

        [Fact]
        public void FrozenBounds_AreTheDocumentedValues()
        {
            Assert.Equal(32767, NatsUdfCommon.MaxCellTextLength);
            Assert.Equal(1_048_576, NatsUdfCommon.MaxObjectCellBytes);
        }

        // ------------------------------------------------------------------
        // Fail: the terse prefix.
        // ------------------------------------------------------------------

        [Fact]
        public void Fail_PrependsTheFrozenPrefix()
        {
            Assert.Equal("Error: x", NatsUdfCommon.Fail("x"));
            Assert.Equal(NatsUdfCommon.ErrorPrefix + "anything", NatsUdfCommon.Fail("anything"));
        }

        [Fact]
        public void Fail_NullMessageBecomesJustThePrefix()
        {
            Assert.Equal(NatsUdfCommon.ErrorPrefix, NatsUdfCommon.Fail(null));
        }

        // ------------------------------------------------------------------
        // VerboseError: the documented shape.
        // ------------------------------------------------------------------

        [Fact]
        public void VerboseError_ProducesContextCauseAndHint()
        {
            Assert.Equal(
                "Error: publish: boom | try again",
                NatsUdfCommon.VerboseError("publish", "boom", "try again"));
        }

        [Fact]
        public void VerboseError_OmitsTheHintWhenBlank()
        {
            Assert.Equal("Error: publish: boom", NatsUdfCommon.VerboseError("publish", "boom", null));
            Assert.Equal("Error: publish: boom", NatsUdfCommon.VerboseError("publish", "boom", "   "));
        }

        [Fact]
        public void VerboseError_SubstitutesPlaceholdersForBlankParts()
        {
            // A blank context or cause must still yield a usable, shaped error.
            Assert.Equal(
                "Error: (no context): unknown error",
                NatsUdfCommon.VerboseError(null, null, null));
            Assert.Equal(
                "Error: (no context): unknown error",
                NatsUdfCommon.VerboseError("", "", ""));
        }

        [Fact]
        public void VerboseError_MasksCredentialsInContextCauseAndHint()
        {
            string text = NatsUdfCommon.VerboseError(
                "nats://exceluser:secretpw@nats.example:4222",
                "auth failed for token=tok3n",
                "see nats://a:b@nats.example");

            // The secrets are gone...
            Assert.DoesNotContain("secretpw", text);
            Assert.DoesNotContain("tok3n", text);
            // ...but the non-secret host survives so the cell stays actionable.
            Assert.Contains("nats.example", text);
            // The shape is unchanged by masking.
            Assert.StartsWith("Error: ", text);
            Assert.Contains(" | ", text);
        }

        [Fact]
        public void VerboseError_MasksURLUserInfoInTheCause()
        {
            string text = NatsUdfCommon.VerboseError(
                "connect", "failed to reach nats://user:p4ss@nats.example:4222", null);

            Assert.DoesNotContain("p4ss", text);
            Assert.Contains("nats.example", text);
            Assert.Contains("****@", text);
        }

        [Fact]
        public void VerboseError_BoundsAPathologicalContext()
        {
            string big = new string('x', 5000);
            string text = NatsUdfCommon.VerboseError(big, "boom", null);
            // The bounded part ends with the documented truncation marker and the
            // whole cell stays far below the Excel limit.
            Assert.Contains("...", text);
            Assert.True(text.Length < 1000, "VerboseError must bound its context");
        }

        // ------------------------------------------------------------------
        // Named-cause builders name their identifiers.
        // ------------------------------------------------------------------

        [Fact]
        public void StreamNotFound_NamesTheStream()
        {
            Assert.Equal("stream 'orders' not found", NatsUdfCommon.StreamNotFound("orders"));
            Assert.Equal("stream '' not found", NatsUdfCommon.StreamNotFound(null));
        }

        [Fact]
        public void ConsumerNotFound_NamesTheConsumerAndStream()
        {
            Assert.Equal(
                "consumer 'worker' not found on stream 'orders'",
                NatsUdfCommon.ConsumerNotFound("worker", "orders"));
            Assert.Equal(
                "consumer '' not found on stream ''",
                NatsUdfCommon.ConsumerNotFound(null, null));
        }

        [Fact]
        public void WrongLastRevision_NamesTheKeyAndExpectedRevision()
        {
            Assert.Equal(
                "wrong last revision for key 'BTC' (expected 7)",
                NatsUdfCommon.WrongLastRevision("BTC", 7));
            Assert.Equal(
                "wrong last revision for key '' (expected 0)",
                NatsUdfCommon.WrongLastRevision(null, 0));
        }

        // ------------------------------------------------------------------
        // The admin / JetStream gates return the named cause with the prefix.
        // ------------------------------------------------------------------

        [Fact]
        public void RequireAdminWrites_ReturnsTheNamedGateErrorWhenClosed()
        {
            Assert.Null(NatsUdfCommon.RequireAdminWrites(true));
            Assert.Equal(
                NatsUdfCommon.ErrorPrefix + NatsUdfCommon.ErrorAdminWritesDisabled,
                NatsUdfCommon.RequireAdminWrites(false));
        }

        [Fact]
        public void RequireJetStream_ReturnsTheNamedGateErrorWhenDisabled()
        {
            Assert.Null(NatsUdfCommon.RequireJetStream(true));
            Assert.Equal(
                NatsUdfCommon.ErrorPrefix + NatsUdfCommon.ErrorJetStreamDisabled,
                NatsUdfCommon.RequireJetStream(false));
        }

        // ------------------------------------------------------------------
        // ErrorHint: the recurrent failure classes.
        // ------------------------------------------------------------------

        [Fact]
        public void ErrorHint_IsNullWhenTheCauseSpeaksForItself()
        {
            Assert.Null(NatsUdfCommon.ErrorHint(null));
            Assert.Null(NatsUdfCommon.ErrorHint("something odd happened"));
        }

        [Fact]
        public void ErrorHint_CoversTheRecurrentClasses()
        {
            Assert.NotNull(NatsUdfCommon.ErrorHint("JetStream is not enabled on this server"));
            Assert.NotNull(NatsUdfCommon.ErrorHint("no suitable peers for placement (10005)"));
            Assert.NotNull(NatsUdfCommon.ErrorHint("No Responders"));
            Assert.NotNull(NatsUdfCommon.ErrorHint("wrong last revision for key 'k' (expected 1)"));
            Assert.NotNull(NatsUdfCommon.ErrorHint("consumer 'c' not found on stream 's'"));
            Assert.NotNull(NatsUdfCommon.ErrorHint("stream 's' not found"));
            Assert.NotNull(NatsUdfCommon.ErrorHint("request timed out"));
        }
    }
}
