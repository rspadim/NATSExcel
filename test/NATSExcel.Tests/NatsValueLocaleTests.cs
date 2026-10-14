using System;
using System.Globalization;
using System.Threading;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// PLAN F1.7 (culture invariance). Values and identifiers written to NATS must
    /// look the same whatever the user's Excel locale is: a German user's cell
    /// holding 1,5 must publish "1.5" (a JSON number and a NATS payload are not
    /// locale-aware), and a subject must never be rewritten by a decimal separator.
    ///
    /// These run under a REAL de-DE culture on the executing thread, because that is
    /// the only way the assertion can fail: with the invariant culture installed,
    /// "1.5" passes even from a culture-dependent ToString(). The culture is restored
    /// in a finally, so a failure cannot poison the rest of the run.
    /// </summary>
    public class NatsValueLocaleTests
    {
        /// <summary>Runs the body on a German thread and restores the previous
        /// culture even when the body throws.</summary>
        private static void InGerman(Action body)
        {
            CultureInfo previousCulture = Thread.CurrentThread.CurrentCulture;
            CultureInfo previousUi = Thread.CurrentThread.CurrentUICulture;
            try
            {
                var german = new CultureInfo("de-DE");
                Thread.CurrentThread.CurrentCulture = german;
                Thread.CurrentThread.CurrentUICulture = german;
                body();
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previousCulture;
                Thread.CurrentThread.CurrentUICulture = previousUi;
            }
        }

        [Fact]
        public void TheFixtureItselfUsesADecimalComma()
        {
            // Guards the test, not the product: if the machine cannot provide de-DE
            // (or a future .NET formats it differently), the assertions below would
            // silently stop proving anything.
            InGerman(() =>
            {
                var german = new CultureInfo("de-DE");
                Assert.Equal("1,5", 1.5.ToString(german));
                Assert.Equal("1,5", 1.5.ToString());
            });
        }

        [Theory]
        [InlineData(1.5, "1.5")]
        [InlineData(0.1, "0.1")]
        [InlineData(1234.5, "1234.5")]     // no thousands separator either
        [InlineData(-2.25, "-2.25")]
        [InlineData(10.0, "10")]           // an integral double stays an integer
        [InlineData(0.0, "0")]
        public void ToNatsString_FormatsNumbersInvariantly(double value, string expected)
        {
            InGerman(() => Assert.Equal(expected, NatsUdfCommon.ToNatsString(value)));
        }

        [Fact]
        public void ToNatsString_KeepsTheSignOfNegativeZero()
        {
            InGerman(() => Assert.Equal("-0", NatsUdfCommon.ToNatsString(-0.0)));
        }

        [Fact]
        public void ToNatsString_FormatsDatesInIsoRoundTrip()
        {
            var when = new DateTime(2026, 10, 14, 12, 30, 45, 123, DateTimeKind.Utc);
            InGerman(() => Assert.Equal("2026-10-14T12:30:45.1230000Z", NatsUdfCommon.ToNatsString(when)));
        }

        [Fact]
        public void ToNatsString_UsesTheFrozenBooleanVocabulary()
        {
            InGerman(() =>
            {
                Assert.Equal("true", NatsUdfCommon.ToNatsString(true));
                Assert.Equal("false", NatsUdfCommon.ToNatsString(false));
            });
        }

        [Fact]
        public void MatrixToJson_EmitsJsonNumbersWithADot()
        {
            // JSON has no locale: "1,5" would be an invalid document (and a parser
            // would read it as two values).
            InGerman(() =>
            {
                Assert.Equal("[1.5,2]", NatsJson.MatrixToJson(new object[,] { { 1.5, 2.0 } }));
                Assert.Equal("0.1", NatsJson.MatrixToJson(new object[,] { { 0.1 } }));
            });
        }

        [Fact]
        public void IdentifiersAreNeverRewrittenByTheCulture()
        {
            // A subject/bucket/key is a byte sequence, not a number: the validation
            // and the pass-through must not touch it under any culture.
            InGerman(() =>
            {
                Assert.Null(NatsUdfCommon.RequireSubject("orders.eu.1", allowWildcard: false));
                Assert.Null(NatsUdfCommon.RequireBucket("k-v.1"));
                Assert.Null(NatsUdfCommon.RequireStream("s.1"));
                Assert.Equal("orders.eu.1", NatsUdfCommon.ToNatsString("orders.eu.1"));
            });
        }
    }
}
