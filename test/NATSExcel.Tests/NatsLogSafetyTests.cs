using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins the log-safety bounds (PLAN.md S3.3): <c>ShortForLog</c> caps a value
    /// at 64 chars with a "...(truncated)" suffix by default; <c>BoundContext</c>
    /// caps a diagnostic context at 200 chars with a "..." suffix. A value at or
    /// under the limit is returned unchanged and a null/empty input is handled
    /// without throwing.
    /// </summary>
    public class NatsLogSafetyTests
    {
        [Fact]
        public void ShortForLog_ShortValueIsUnchanged()
        {
            Assert.Equal("orders", NatsUdfCommon.ShortForLog("orders"));
            Assert.Equal(new string('x', 64), NatsUdfCommon.ShortForLog(new string('x', 64)));
        }

        [Fact]
        public void ShortForLog_LongValueIsTruncatedWithTheDocumentedMarker()
        {
            string value = new string('x', 100);
            string result = NatsUdfCommon.ShortForLog(value);

            Assert.Equal(new string('x', 64) + "...(truncated)", result);
        }

        [Fact]
        public void ShortForLog_HonorsACustomLimit()
        {
            Assert.Equal("ab...(truncated)", NatsUdfCommon.ShortForLog("abcdef", 2));
        }

        [Fact]
        public void ShortForLog_NullAndEmptyAreHandled()
        {
            Assert.Null(NatsUdfCommon.ShortForLog(null));
            Assert.Equal("", NatsUdfCommon.ShortForLog(""));
        }

        [Fact]
        public void BoundContext_ShortValueIsUnchanged()
        {
            Assert.Equal("publish", NatsUdfCommon.BoundContext("publish"));
            Assert.Equal(new string('y', 200), NatsUdfCommon.BoundContext(new string('y', 200)));
        }

        [Fact]
        public void BoundContext_LongValueIsTruncatedWithTheDocumentedMarker()
        {
            string value = new string('y', 300);
            string result = NatsUdfCommon.BoundContext(value);

            Assert.Equal(new string('y', 200) + "...", result);
        }

        [Fact]
        public void BoundContext_HonorsACustomLimit()
        {
            Assert.Equal("ab...", NatsUdfCommon.BoundContext("abcdef", 2));
        }

        [Fact]
        public void BoundContext_NullAndEmptyAreHandled()
        {
            Assert.Null(NatsUdfCommon.BoundContext(null));
            Assert.Equal("", NatsUdfCommon.BoundContext(""));
        }
    }
}
