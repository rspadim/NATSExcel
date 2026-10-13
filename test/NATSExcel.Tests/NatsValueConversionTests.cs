using System;
using System.Globalization;
using System.Threading;
using ExcelDna.Integration;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins the invariant value-conversion rules of
    /// <see cref="NatsUdfCommon.ToNatsString"/> (DESIGN.md §7.6/§8.1): numbers use
    /// the invariant culture regardless of the host locale, booleans become the
    /// lower-case literals, negative zero keeps its sign, a DateTime becomes
    /// ISO-8601 round-trip text and a truly missing cell becomes null.
    /// </summary>
    public class NatsValueConversionTests
    {
        private static readonly CultureInfo CommaLocale = new CultureInfo("de-DE");

        private static void WithCulture(CultureInfo culture, Action action)
        {
            CultureInfo previous = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = culture;
                action();
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        [Fact]
        public void ToNatsString_KeepsTheInvariantDecimalPointUnderCommaLocale()
        {
            WithCulture(CommaLocale, () =>
            {
                Assert.Equal("1.5", NatsUdfCommon.ToNatsString(1.5));
                Assert.Equal("67000.5", NatsUdfCommon.ToNatsString(67000.5));
                // A negative decimal too, to be sure the sign separator is not
                // picked up from the locale.
                Assert.Equal("-0.25", NatsUdfCommon.ToNatsString(-0.25));
            });
        }

        [Fact]
        public void ToNatsString_BooleansAreLowerCaseLiterals()
        {
            Assert.Equal("true", NatsUdfCommon.ToNatsString(true));
            Assert.Equal("false", NatsUdfCommon.ToNatsString(false));
        }

        [Fact]
        public void ToNatsString_NegativeZeroKeepsItsSign()
        {
            Assert.Equal("-0", NatsUdfCommon.ToNatsString(-0.0));
            Assert.Equal("-0", NatsUdfCommon.ToNatsString(-0.0f));
            Assert.Equal("0", NatsUdfCommon.ToNatsString(0.0));
        }

        [Fact]
        public void ToNatsString_DateTimeIsIso8601RoundTripText()
        {
            var value = new DateTime(2026, 10, 13, 12, 34, 56, DateTimeKind.Utc);
            string text = NatsUdfCommon.ToNatsString(value);

            // The round-trip ("o") format round-trips exactly and uses the
            // invariant layout even under a comma-decimal culture.
            DateTime parsed = DateTime.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
            Assert.Equal(value, parsed);
            Assert.Equal(value.ToString("o", CultureInfo.InvariantCulture), text);
        }

        [Fact]
        public void ToNatsString_NullAndMissingCellsAreNull()
        {
            Assert.Null(NatsUdfCommon.ToNatsString(null));
            Assert.Null(NatsUdfCommon.ToNatsString(ExcelMissing.Value));
            Assert.Null(NatsUdfCommon.ToNatsString(ExcelEmpty.Value));
        }

        [Fact]
        public void ToNatsString_StringsPassThroughVerbatim()
        {
            Assert.Equal("orders", NatsUdfCommon.ToNatsString("orders"));
            Assert.Equal("  spaced  ", NatsUdfCommon.ToNatsString("  spaced  "));
        }

        [Fact]
        public void ToNatsString_IntegralDoubleDropsTheFraction()
        {
            Assert.Equal("42", NatsUdfCommon.ToNatsString(42.0));
        }

        [Fact]
        public void ToNatsString_RejectsARangeAndAnExcelErrorCell()
        {
            Assert.Equal(
                NatsUdfCommon.ErrorMultiCellScalar,
                Assert.Throws<ArgumentException>(
                    () => NatsUdfCommon.ToNatsString(new object[2, 2])).Message);
            Assert.Equal(
                NatsUdfCommon.ErrorExcelErrorCell,
                Assert.Throws<ArgumentException>(
                    () => NatsUdfCommon.ToNatsString(ExcelError.ExcelErrorValue)).Message);
        }
    }
}
