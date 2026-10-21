using System;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The version/diagnostic worksheet function. It answers the first question a
    /// support case asks ("which NATSExcel is running, and from where?"), and it
    /// is the cheap way to spot TWO installed copies: each one reports a
    /// different xll path. It must never throw into Excel, and it must work
    /// without a NATS server (a broken install is exactly when it is used).
    /// </summary>
    public class VersionInfoTests
    {
        [Fact]
        public void BuildInfo_VersionIsTheTagWithoutTheVPrefix()
        {
            // The tag is embedded by the CI (InformationalVersion); the readable
            // version is that tag minus the "v" and any build metadata. A local
            // build reports "dev", so the assertion is on the SHAPE, not on the
            // number - and it must never be blank.
            Assert.False(string.IsNullOrWhiteSpace(BuildInfo.Version));
            Assert.False(string.IsNullOrWhiteSpace(BuildInfo.Tag));
            Assert.StartsWith("NATSExcel " + BuildInfo.Version, NatsUdf.NatsUDFVersion());
            Assert.DoesNotContain("+", BuildInfo.Version);
            Assert.DoesNotContain(" ", BuildInfo.Version);
            Assert.False(BuildInfo.Version.StartsWith("v", StringComparison.OrdinalIgnoreCase)
                && BuildInfo.Version.Length > 1 && char.IsDigit(BuildInfo.Version[1]),
                "the leading 'v' must be stripped: " + BuildInfo.Version);
        }

        [Fact]
        public void NatsUDFVersion_ReportsTheXllPathAndConfigInUse()
        {
            string text = NatsUdf.NatsUDFVersion();

            // xll= is the point of the function: without it two copies are
            // indistinguishable. Under the test runner there is no Excel host,
            // so ExcelDnaUtil.XllPath is unavailable and it reports "(unknown)";
            // inside Excel it carries the loaded file.
            Assert.Contains("| xll=", text, StringComparison.Ordinal);
            Assert.Contains("| config=", text, StringComparison.Ordinal);
            Assert.Contains("| connections=", text, StringComparison.Ordinal);
            Assert.Equal(4, text.Split('|').Length);

            // Single line on purpose: a multi-line volatile cell makes the sheet
            // jump every recalculation.
            Assert.DoesNotContain("\n", text);
            Assert.DoesNotContain("\r", text);
        }

        [Fact]
        public void NatsUDFVersion_DoesNotNeedANatsServer()
        {
            // No connection is opened by the function itself; the counter is a
            // neutral read (0 when nothing connected). An endpoint that is down
            // must not turn this cell into an error.
            Assert.Contains("connections=0", NatsUdf.NatsUDFVersion());

            // It must never throw into the cell, whatever the config situation is:
            // a missing file reports "(defaults)", an unreadable one is still one
            // line of text.
            Assert.DoesNotContain("Error:", NatsUdf.NatsUDFVersion());
        }
    }
}
