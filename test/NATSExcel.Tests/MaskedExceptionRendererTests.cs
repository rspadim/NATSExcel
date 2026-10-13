using System;
using System.IO;
using NLog;
using NLog.Config;
using NLog.LayoutRenderers;
using NLog.Targets;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The log file must never contain a credential. A NATS endpoint can be a
    /// full connection string, and a client exception embeds it verbatim, so the
    /// exception field of every log line goes through the masking layout
    /// renderer. These tests render the real NLog pipeline and read the output.
    /// </summary>
    public class MaskedExceptionRendererTests
    {
        private static string Render(LogEventInfo logEvent)
        {
            var renderer = new MaskedExceptionRenderer { Format = "toString" };
            return renderer.Render(logEvent);
        }

        private static LogEventInfo Event(Exception ex)
        {
            return new LogEventInfo(LogLevel.Error, "test", "failed") { Exception = ex };
        }

        [Fact]
        public void MasksUrlUserInfoInAnExceptionMessage()
        {
            string text = Render(Event(new Exception("connect failed to nats://excel:s3cret@host:4222")));
            Assert.DoesNotContain("s3cret", text);
            Assert.Contains("****@", text);
            // The host stays readable: only the credential is redacted.
            Assert.Contains("host:4222", text);
        }

        [Fact]
        public void MasksAKeyValuePasswordAndAToken()
        {
            string text = Render(Event(new Exception("host=node1:4222,password=topsecret,user=excel")));
            Assert.DoesNotContain("topsecret", text);
            Assert.DoesNotContain("excel", text);
            Assert.Contains("password=****", text);
        }

        [Fact]
        public void MasksAnNkeySeedAndAJwt()
        {
            string seed = "SU" + new string('A', 30);
            string jwt = "eyJhbGciOiJlZDI1NTE5AAAAI" + new string('x', 20);
            string text = Render(Event(new Exception("auth " + seed + " " + jwt)));
            Assert.DoesNotContain(seed, text);
            Assert.DoesNotContain(jwt, text);
        }

        [Fact]
        public void MasksEveryInnerExceptionInTheChain()
        {
            var inner = new Exception("inner nats://u:pw@h1:4222");
            var outer = new Exception("outer nats://u:pw@h2:4222", inner);
            string text = Render(Event(outer));
            Assert.DoesNotContain("u:pw@", text);
            Assert.Contains("h1:4222", text);
            Assert.Contains("h2:4222", text);
        }

        [Fact]
        public void AnExceptionWithoutACredentialIsUnchanged()
        {
            string text = Render(Event(new Exception("timed out after 1000 ms")));
            Assert.Contains("timed out after 1000 ms", text);
        }

        [Fact]
        public void ALogEventWithoutAnExceptionRendersEmpty()
        {
            var logEvent = new LogEventInfo(LogLevel.Info, "test", "no exception");
            Assert.Equal(string.Empty, Render(logEvent));
        }
    }
}
