using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// A minimal, robust smoke test for the configuration loader. It deliberately
    /// sticks to the parts of the <c>AppConfig</c> surface that are frozen
    /// (PLAN.md §1.4): the file-name and default-URL constants, plus the
    /// comment-tolerant read path. The richer behavior (sanitize, schema
    /// validation, endpoint resolution, masking) belongs to the dedicated unit
    /// tests and is not asserted here.
    /// </summary>
    public class ConfigSmokeTests
    {
        [Fact]
        public void ConfigConstants_ArePinned()
        {
            Assert.Equal("NATSExcel.json", AppConfig.ConfigFileName);
            Assert.Equal("nats://localhost:4222", AppConfig.DefaultUrl);
        }

        [Fact]
        public void ShippedSample_IsCommentTolerantAndParses_WithTheExpectedShape()
        {
            string repoRoot = FindRepoRoot();
            string samplePath = Path.Combine(repoRoot, "NATSExcel.json");
            Assert.True(File.Exists(samplePath), "shipped sample not found at " + samplePath);

            string raw = File.ReadAllText(samplePath);
            Assert.Contains("//", raw); // the sample really does carry comments

            // The loader strips // and /* */ without damaging string contents
            // (so nats:// survives), which is what makes the sample readable.
            string stripped = AppConfig.StripComments(raw);
            Assert.DoesNotContain("//", stripped.Replace("nats://", string.Empty));

            var root = JObject.Parse(stripped);

            // Shape pinned by the shipped sample and NATSExcel.schema.json.
            Assert.NotNull(root["Default"]);
            Assert.Equal("local", (string)root["Default"]["cluster"]);
            Assert.Equal("nats://localhost:4222", (string)root["Servers"]["dev"]);
            Assert.NotNull(root["Clusters"]["local"]);
            Assert.NotNull(root["JetStream"]);
            Assert.Equal("fireforget", (string)root["SyncWrite"]);
            Assert.False((bool)root["AllowAdminWrites"]);
        }

        /// <summary>
        /// Walks up from the test binary until the repository root. There is no
        /// NATSExcel.sln yet (PLAN F0.1), so the root is anchored by the add-in
        /// project file instead.
        /// </summary>
        private static string FindRepoRoot()
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null)
            {
                if (File.Exists(Path.Combine(directory.FullName, "NATSExcel.csproj")))
                    return directory.FullName;
                directory = directory.Parent;
            }
            throw new InvalidOperationException(
                "NATSExcel.csproj was not found above " + AppContext.BaseDirectory);
        }
    }
}
