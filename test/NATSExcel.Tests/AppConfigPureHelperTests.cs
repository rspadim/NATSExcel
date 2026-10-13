using System;
using System.IO;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The remaining pure <see cref="AppConfig"/> seams: the write-mode
    /// normalizer, the NKey-seed resolver (AGENTS golden rule 11) and
    /// <see cref="AppConfig.Sanitize"/> (PLAN F1.4). No config fixture is needed:
    /// these are static and pure (the seed test touches only a private temp file).
    /// </summary>
    public class AppConfigPureHelperTests
    {
        // ------------------------------------------------------------------
        // NormalizeSyncWrite.
        // ------------------------------------------------------------------

        [Fact]
        public void NormalizeSyncWrite_RecognizesTheThreeModesCaseInsensitively()
        {
            Assert.Equal("sync", AppConfig.NormalizeSyncWrite("sync"));
            Assert.Equal("sync", AppConfig.NormalizeSyncWrite("SYNC"));
            Assert.Equal("sync", AppConfig.NormalizeSyncWrite("  Sync  "));
            Assert.Equal("fireforget", AppConfig.NormalizeSyncWrite("fireforget"));
            Assert.Equal("fireforget", AppConfig.NormalizeSyncWrite("FIREforget"));
            Assert.Equal("fireforget-all", AppConfig.NormalizeSyncWrite("fireforget-all"));
            Assert.Equal("fireforget-all", AppConfig.NormalizeSyncWrite("FireForget-All"));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("bogus")]
        [InlineData("sink")]
        public void NormalizeSyncWrite_UnknownOrBlank_FallsBackToFireForget(string mode)
        {
            Assert.Equal("fireforget", AppConfig.NormalizeSyncWrite(mode));
        }

        // ------------------------------------------------------------------
        // ResolveSeed.
        // ------------------------------------------------------------------

        [Fact]
        public void ResolveSeed_SeedText_PassesThroughTrimmed()
        {
            const string seed = "SUAMK2J3N4P5Q6R7S8T9V0W1X2Y3Z4A5B6C7D8E9";
            Assert.Equal(seed, AppConfig.ResolveSeed(seed));
            Assert.Equal(seed, AppConfig.ResolveSeed("  " + seed + "  "));
        }

        [Fact]
        public void ResolveSeed_ExistingFilePath_IsReadAndTrimmed()
        {
            string path = Path.Combine(Path.GetTempPath(), "natsexcel-seed-" + Guid.NewGuid().ToString("N") + ".nk");
            try
            {
                File.WriteAllText(path, "  SUFILEBACKEDSEED0123456789ABCDEFGHIJ  \r\n");
                Assert.Equal("SUFILEBACKEDSEED0123456789ABCDEFGHIJ", AppConfig.ResolveSeed(path));
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [Fact]
        public void ResolveSeed_NonExistentPath_PassesThrough()
        {
            string path = Path.Combine(Path.GetTempPath(), "natsexcel-missing-" + Guid.NewGuid().ToString("N") + ".nk");
            Assert.Equal(path, AppConfig.ResolveSeed(path));
        }

        [Fact]
        public void ResolveSeed_Blank_PassesThroughUnchanged()
        {
            Assert.Null(AppConfig.ResolveSeed(null));
            Assert.Equal(string.Empty, AppConfig.ResolveSeed(string.Empty));
            Assert.Equal("   ", AppConfig.ResolveSeed("   "));
        }

        // ------------------------------------------------------------------
        // Sanitize.
        // ------------------------------------------------------------------

        [Fact]
        public void Sanitize_NullRoot_ReturnsAFreshDefaultRoot()
        {
            AppConfig.Root root = null;
            Exception thrown = Record.Exception(() => { root = AppConfig.Sanitize(null); });
            Assert.Null(thrown);
            Assert.NotNull(root);
            Assert.Equal("local", root.Default);
            Assert.Equal(1000, root.TimeoutMs);
        }

        [Fact]
        public void Sanitize_NonPositiveTimeoutAndRates_ResetToDefaults()
        {
            var root = new AppConfig.Root
            {
                TimeoutMs = 0,
                JetStreamTimeoutMs = -5,
                PublishDedupCacheSize = 0,
                NatsUpdateRateMs = -1,
                ExcelUpdateRateMs = 0,
            };

            AppConfig.Sanitize(root);

            Assert.Equal(1000, root.TimeoutMs);
            Assert.Equal(5000, root.JetStreamTimeoutMs);
            Assert.Equal(10000, root.PublishDedupCacheSize);
            Assert.Equal(1000, root.NatsUpdateRateMs);
            Assert.Equal(100, root.ExcelUpdateRateMs);
        }

        [Fact]
        public void Sanitize_PositiveValues_AreKept()
        {
            var root = new AppConfig.Root
            {
                TimeoutMs = 750,
                JetStreamTimeoutMs = 900,
                PublishDedupCacheSize = 50,
                NatsUpdateRateMs = 250,
                ExcelUpdateRateMs = 200,
            };

            AppConfig.Sanitize(root);

            Assert.Equal(750, root.TimeoutMs);
            Assert.Equal(900, root.JetStreamTimeoutMs);
            Assert.Equal(50, root.PublishDedupCacheSize);
            Assert.Equal(250, root.NatsUpdateRateMs);
            Assert.Equal(200, root.ExcelUpdateRateMs);
        }

        [Fact]
        public void Sanitize_BlankDefault_BecomesLocal()
        {
            var root = new AppConfig.Root { Default = "   " };
            AppConfig.Sanitize(root);
            Assert.Equal("local", root.Default);
        }

        [Fact]
        public void Sanitize_ConflationMs_IsClampedByTheConflationWindow()
        {
            Assert.Equal(0, AppConfig.Sanitize(new AppConfig.Root { ConflationMs = -50 }).ConflationMs);
            Assert.Equal(0, AppConfig.Sanitize(new AppConfig.Root { ConflationMs = 0 }).ConflationMs);
            Assert.Equal(250, AppConfig.Sanitize(new AppConfig.Root { ConflationMs = 250 }).ConflationMs);
            Assert.Equal(
                Conflation.MaxWindowMs,
                AppConfig.Sanitize(new AppConfig.Root { ConflationMs = Conflation.MaxWindowMs + 1 }).ConflationMs);
        }

        [Fact]
        public void Sanitize_NullConflationMs_StaysNull()
        {
            Assert.Null(AppConfig.Sanitize(new AppConfig.Root { ConflationMs = null }).ConflationMs);
        }

        [Fact]
        public void Sanitize_SyncWrite_IsNormalized()
        {
            Assert.Equal("fireforget", AppConfig.Sanitize(new AppConfig.Root { SyncWrite = "BOGUS" }).SyncWrite);
            Assert.Equal("sync", AppConfig.Sanitize(new AppConfig.Root { SyncWrite = "SYNC" }).SyncWrite);
            Assert.Equal("fireforget-all", AppConfig.Sanitize(new AppConfig.Root { SyncWrite = "FireForget-All" }).SyncWrite);
        }
    }
}
