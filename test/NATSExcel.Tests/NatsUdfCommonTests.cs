using System;
using System.Globalization;
using System.Threading;
using ExcelDna.Integration;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// <see cref="NatsUdfCommon"/> is the shared, client-agnostic vocabulary:
    /// invariant value conversion, argument validation, the frozen markers and
    /// the write-mode rules (PLAN.md §1.3/§1.4). These tests pin each documented
    /// behavior with no server and no Excel instance.
    /// </summary>
    public class NatsUdfCommonTests
    {
        // ------------------------------------------------------------------
        // Frozen markers (PLAN.md §1.3).
        // ------------------------------------------------------------------

        [Fact]
        public void FrozenMarkers_AreExactlyTheDocumentedStrings()
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
        public void FrozenErrorAndModeConstants_ArePinned()
        {
            Assert.Equal("Error: ", NatsUdfCommon.ErrorPrefix);
            Assert.Equal(32767, NatsUdfCommon.MaxCellTextLength);
            Assert.Equal(1_048_576, NatsUdfCommon.MaxObjectCellBytes);
            Assert.Equal("sync", NatsUdfCommon.SyncWriteSync);
            Assert.Equal("fireforget", NatsUdfCommon.SyncWriteFireForget);
            Assert.Equal("fireforget-all", NatsUdfCommon.SyncWriteFireForgetAll);
        }

        // ------------------------------------------------------------------
        // ToNatsString: invariant text conversion.
        // ------------------------------------------------------------------

        [Fact]
        public void ToNatsString_UsesInvariantDecimalPoint_EvenUnderCommaLocale()
        {
            CultureInfo previous = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                Assert.Equal("1.5", NatsUdfCommon.ToNatsString(1.5));
                Assert.Equal("67000.5", NatsUdfCommon.ToNatsString(67000.5));
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        [Fact]
        public void ToNatsString_BooleanBecomesLowerCaseLiteral()
        {
            Assert.Equal("true", NatsUdfCommon.ToNatsString(true));
            Assert.Equal("false", NatsUdfCommon.ToNatsString(false));
        }

        [Fact]
        public void ToNatsString_NegativeZero_KeepsItsSign()
        {
            // Documented behavior: -0.0 would otherwise format as "0" and lose
            // the sign bit; the code inspects the bits and returns "-0".
            Assert.Equal("-0", NatsUdfCommon.ToNatsString(-0.0));
            Assert.Equal("-0", NatsUdfCommon.ToNatsString(-0.0f));
            Assert.Equal("0", NatsUdfCommon.ToNatsString(0.0));
        }

        [Fact]
        public void ToNatsString_PassesStringsThroughVerbatim()
        {
            Assert.Equal("orders", NatsUdfCommon.ToNatsString("orders"));
            Assert.Equal("  spaced  ", NatsUdfCommon.ToNatsString("  spaced  "));
        }

        [Fact]
        public void ToNatsString_TrulyMissingCellBecomesNull()
        {
            Assert.Null(NatsUdfCommon.ToNatsString(null));
            Assert.Null(NatsUdfCommon.ToNatsString(ExcelMissing.Value));
            Assert.Null(NatsUdfCommon.ToNatsString(ExcelEmpty.Value));
        }

        [Fact]
        public void ToNatsString_RejectsAMultiCellRangeWithTheFrozenMessage()
        {
            var ex = Assert.Throws<ArgumentException>(
                () => NatsUdfCommon.ToNatsString(new object[2, 2]));
            Assert.Equal(NatsUdfCommon.ErrorMultiCellScalar, ex.Message);
        }

        [Fact]
        public void ToNatsString_RejectsAnExcelErrorCellWithTheFrozenMessage()
        {
            var ex = Assert.Throws<ArgumentException>(
                () => NatsUdfCommon.ToNatsString(ExcelError.ExcelErrorValue));
            Assert.Equal(NatsUdfCommon.ErrorExcelErrorCell, ex.Message);
        }

        // ------------------------------------------------------------------
        // TryToLong / TryToInt: integral conversion.
        // ------------------------------------------------------------------

        [Fact]
        public void TryToLong_RejectsMultiCellBlankExcelErrorBoolFractionAndText()
        {
            long value;
            Assert.False(NatsUdfCommon.TryToLong(new object[1, 2], out value)); // multi-cell range
            Assert.False(NatsUdfCommon.TryToLong(null, out value));             // blank
            Assert.False(NatsUdfCommon.TryToLong("", out value));               // blank text
            Assert.False(NatsUdfCommon.TryToLong(ExcelError.ExcelErrorValue, out value));
            Assert.False(NatsUdfCommon.TryToLong(true, out value));             // bool
            Assert.False(NatsUdfCommon.TryToLong(1.5, out value));              // fraction
            Assert.False(NatsUdfCommon.TryToLong(1.5m, out value));             // decimal fraction
            Assert.False(NatsUdfCommon.TryToLong("abc", out value));            // non-numeric text
        }

        [Fact]
        public void TryToLong_AcceptsIntegersAndNumericLookingStrings()
        {
            long value;
            Assert.True(NatsUdfCommon.TryToLong("42", out value));
            Assert.Equal(42, value);
            Assert.True(NatsUdfCommon.TryToLong(" 7 ", out value));
            Assert.Equal(7, value);
            Assert.True(NatsUdfCommon.TryToLong(9.0, out value));
            Assert.Equal(9, value);
            Assert.True(NatsUdfCommon.TryToLong(-3, out value));
            Assert.Equal(-3, value);
        }

        [Fact]
        public void TryToInt_RejectsNonIntegralsAndOutOfRangeValues()
        {
            int value;
            Assert.False(NatsUdfCommon.TryToInt(new object[1, 2], out value));
            Assert.False(NatsUdfCommon.TryToInt("", out value));
            Assert.False(NatsUdfCommon.TryToInt(ExcelError.ExcelErrorNA, out value));
            Assert.False(NatsUdfCommon.TryToInt(false, out value));
            Assert.False(NatsUdfCommon.TryToInt(2.25, out value));
            Assert.False(NatsUdfCommon.TryToInt("nope", out value));
            // A long that does not fit an int must not silently wrap.
            Assert.False(NatsUdfCommon.TryToInt(1L + int.MaxValue, out value));
        }

        [Fact]
        public void TryToInt_AcceptsIntegersAndNumericLookingStrings()
        {
            int value;
            Assert.True(NatsUdfCommon.TryToInt("42", out value));
            Assert.Equal(42, value);
            Assert.True(NatsUdfCommon.TryToInt(5.0, out value));
            Assert.Equal(5, value);
        }

        // ------------------------------------------------------------------
        // RequireSubject: wildcards are subscribe-only.
        // ------------------------------------------------------------------

        [Fact]
        public void RequireSubject_RejectsWildcardsWhenTheyAreNotAllowed()
        {
            Assert.NotNull(NatsUdfCommon.RequireSubject("a.*", allowWildcard: false));
            Assert.NotNull(NatsUdfCommon.RequireSubject("a.>", allowWildcard: false));
        }

        [Fact]
        public void RequireSubject_AllowsAConcreteSubjectWhenWildcardsAreNotAllowed()
        {
            Assert.Null(NatsUdfCommon.RequireSubject("a.b", allowWildcard: false));
        }

        [Fact]
        public void RequireSubject_AllowsWildcardsWhenTheyAreAllowed()
        {
            Assert.Null(NatsUdfCommon.RequireSubject("a.*", allowWildcard: true));
            Assert.Null(NatsUdfCommon.RequireSubject("a.>", allowWildcard: true));
        }

        [Fact]
        public void RequireSubject_DoesNotTreatAPartialTokenAsAWildcard()
        {
            // "a*b" is a literal token, not a wildcard (token-based match).
            Assert.Null(NatsUdfCommon.RequireSubject("a*b", allowWildcard: false));
            Assert.Null(NatsUdfCommon.RequireSubject("orders.>" + "x", allowWildcard: false));
        }

        [Fact]
        public void RequireSubject_RejectsABlankSubject()
        {
            Assert.Equal(
                NatsUdfCommon.ErrorPrefix + "a non-empty subject is required",
                NatsUdfCommon.RequireSubject("", allowWildcard: true));
            Assert.NotNull(NatsUdfCommon.RequireSubject(null, allowWildcard: true));
        }

        // ------------------------------------------------------------------
        // ShouldFireAndForget: the frozen write-mode table.
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("sync", true, false)]
        [InlineData("sync", false, false)]
        [InlineData("fireforget", true, false)]
        [InlineData("fireforget", false, true)]
        [InlineData("fireforget-all", true, true)]
        [InlineData("fireforget-all", false, true)]
        public void ShouldFireAndForget_FollowsTheDocumentedTable(
            string mode, bool replyDependent, bool expected)
        {
            Assert.Equal(expected, NatsUdfCommon.ShouldFireAndForget(mode, replyDependent));
        }

        [Fact]
        public void ShouldFireAndForget_UnknownOrBlankModeFallsBackToFireForgetDefault()
        {
            Assert.False(NatsUdfCommon.ShouldFireAndForget("bogus", replyDependent: true));
            Assert.True(NatsUdfCommon.ShouldFireAndForget("bogus", replyDependent: false));
            Assert.True(NatsUdfCommon.ShouldFireAndForget(null, replyDependent: false));
        }

        [Fact]
        public void ShouldFireAndForget_ModeIsCaseAndWhitespaceInsensitive()
        {
            Assert.False(NatsUdfCommon.ShouldFireAndForget("  SYNC  ", replyDependent: false));
            Assert.True(NatsUdfCommon.ShouldFireAndForget(" FireForget-All ", replyDependent: true));
        }

        [Fact]
        public void FireAndForgetMarker_OnlyFireForgetAllGetsTheAllMarker()
        {
            Assert.Equal(NatsUdfCommon.FireAndForget, NatsUdfCommon.FireAndForgetMarker("fireforget"));
            Assert.Equal(NatsUdfCommon.FireAndForget, NatsUdfCommon.FireAndForgetMarker("sync"));
            Assert.Equal(NatsUdfCommon.FireAndForgetAll, NatsUdfCommon.FireAndForgetMarker("fireforget-all"));
        }
    }
}
