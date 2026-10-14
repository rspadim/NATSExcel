using System;
using System.Reflection;
using ExcelDna.Integration;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins the Excel-DNA integration build the async write dispatch was audited
    /// against: the lock-order and threading findings behind
    /// <c>NatsWriteObservable</c> / <c>ExcelAsyncUtil.Observe</c> were verified for
    /// this exact revision, and the packed XLL ships it (NATSExcel.dna). A framework
    /// bump must be a deliberate change that re-runs that audit, so version drift
    /// fails loudly here instead of silently changing the delivery semantics.
    ///
    /// Adapted from the sibling RedisExcel suite (ExcelDnaPinTests): the two
    /// add-ins share the same async dispatch design.
    /// </summary>
    public class ExcelDnaPinTests
    {
        [Fact]
        public void ExcelDnaIntegration_InformationalVersion_IsPinned()
        {
            var attribute = (AssemblyInformationalVersionAttribute)Attribute.GetCustomAttribute(
                typeof(ExcelAsyncUtil).Assembly,
                typeof(AssemblyInformationalVersionAttribute));

            Assert.NotNull(attribute);
            Assert.StartsWith("1.9.0.6+c631da4e", attribute.InformationalVersion);
        }

        /// <summary>The add-in must declare the same Excel-DNA package the test
        /// project compile-references: the pack list in the .dna names only the
        /// PACKED dependencies (Excel-DNA itself is provided by the XLL host), so
        /// the csproj is what pins the version the audit was run against.</summary>
        [Fact]
        public void TheAddInDeclaresTheSameExcelDnaVersionAsThePin()
        {
            string csproj = System.IO.File.ReadAllText(
                System.IO.Path.Combine(RepoRoot(), "NATSExcel.csproj"));
            Assert.Contains("Include=\"ExcelDna.Integration\" Version=\"1.9.0-beta2\"", csproj);
        }

        internal static string RepoRoot()
        {
            var dir = new System.IO.DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !System.IO.File.Exists(System.IO.Path.Combine(dir.FullName, "NATSExcel.sln")))
                dir = dir.Parent;
            Assert.NotNull(dir);
            return dir.FullName;
        }
    }
}
