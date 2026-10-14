using System;
using System.IO;
using NLog;
using NLog.Config;
using NLog.Targets;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The log file must land where every document, the installer's uninstaller and
    /// <c>=NatsUDFLogPath()</c> say it does: NEXT TO THE ADD-IN. NLog resolves a
    /// relative <c>fileName</c> against <c>AppDomain.BaseDirectory</c>, which in Excel
    /// is the OFFICE folder, so the shipped config alone put the log in Program Files
    /// and the three claims disagreed with reality. <c>AddIn.AutoOpen</c> repoints the
    /// target through <see cref="NatsUdf.TryPointLogAt"/>; this pins that wiring,
    /// including the unwritable-folder fallback (a read-only add-in folder must keep
    /// NLog's default rather than cost the user the log entirely).
    /// </summary>
    public class LogPathTests
    {
        /// <summary>Runs the body with a minimal NLog configuration whose file target
        /// is named "file" (the name the shipped NLog.config uses), restoring the
        /// previous configuration afterwards.</summary>
        private static void WithFileTarget(Action<FileTarget> body)
        {
            LoggingConfiguration previous = LogManager.Configuration;
            var config = new LoggingConfiguration();
            var target = new FileTarget("file") { FileName = "NATSExcel.log" };
            config.AddTarget(target);
            LogManager.Configuration = config;
            try
            {
                body(target);
            }
            finally
            {
                LogManager.Configuration = previous;
            }
        }

        [Fact]
        public void TheFileTargetIsPointedAtTheAddInFolder()
        {
            string dir = Path.Combine(Path.GetTempPath(), "natsexcel-logpath-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string logPath = Path.Combine(dir, "NATSExcel.log");
                WithFileTarget(target =>
                {
                    Assert.True(NatsUdf.TryPointLogAt(logPath), "the target must be repointed");
                    Assert.Equal(logPath, target.FileName.ToString());
                });
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        [Fact]
        public void TheLogPathTheUdfReportsIsTheAddInFolderOne()
        {
            // Outside Excel the assembly has a file location, so the reported path is
            // the add-in's own folder - the same folder the target is pointed at.
            string reported = NatsUdf.NatsUDFLogPath();
            Assert.EndsWith("NATSExcel.log", reported, StringComparison.OrdinalIgnoreCase);
            Assert.True(Path.IsPathRooted(reported), "the log path must be absolute, not a relative name");
            Assert.Equal(Path.GetDirectoryName(reported), Path.GetDirectoryName(
                Path.Combine(Path.GetDirectoryName(reported), "NATSExcel.log")));
        }

        [Fact]
        public void AnUnwritableFolderKeepsTheConfiguredDefault()
        {
            // A folder that does not exist stands in for "not writable": the helper
            // must decline (FALSE) and leave the target alone, so the add-in still logs
            // somewhere instead of failing.
            string missing = Path.Combine(Path.GetTempPath(), "natsexcel-missing-" + Guid.NewGuid().ToString("N"));
            WithFileTarget(target =>
            {
                Assert.False(NatsUdf.TryPointLogAt(Path.Combine(missing, "NATSExcel.log")));
                Assert.Equal("NATSExcel.log", target.FileName.ToString());
            });
        }

        [Fact]
        public void ANullOrBlankPathIsRefused()
        {
            WithFileTarget(target =>
            {
                Assert.False(NatsUdf.TryPointLogAt(null));
                Assert.False(NatsUdf.TryPointLogAt(string.Empty));
                Assert.Equal("NATSExcel.log", target.FileName.ToString());
            });
        }

        [Fact]
        public void TheProbeFileIsNotLeftBehind()
        {
            string dir = Path.Combine(Path.GetTempPath(), "natsexcel-logpath-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                WithFileTarget(target => NatsUdf.TryPointLogAt(Path.Combine(dir, "NATSExcel.log")));
                Assert.Empty(Directory.GetFiles(dir, ".natsexcel-write-probe"));
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }
    }
}
