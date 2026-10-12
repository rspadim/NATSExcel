using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Guards the shared production-source list in
    /// <c>test\TestSupport\ProductionSources.props</c> (PLAN D1): the test
    /// projects compile the production sources directly (no ProjectReference),
    /// so a new root production file that is missing from the props would never
    /// be exercised by the unit suite. <c>NatsRtd.cs</c> is the only intentional
    /// exclusion (PLAN D2): its RTD/COM surface needs Excel and is covered by the
    /// E2E script instead.
    ///
    /// This mirrors the RedisExcel.Tests parity test, with two differences: the
    /// item list lives in the imported props (the single source of truth), and
    /// the repository root is anchored by <c>NATSExcel.csproj</c> rather than a
    /// solution file, because the solution does not exist yet (PLAN F0.1).
    /// </summary>
    public class ProjectFileParityTests
    {
        private const string ExcludedFileName = "NatsRtd.cs";

        [Fact]
        public void ProductionSourcesProps_CoversEveryRootProductionSourceExceptNatsRtd()
        {
            string repoRoot = FindRepoRoot();
            string propsPath = Path.Combine(
                repoRoot, "test", "TestSupport", "ProductionSources.props");
            Assert.True(File.Exists(propsPath), "ProductionSources.props not found at " + propsPath);

            var compiled = CollectCompileItems(propsPath);

            // Every top-level production source (the repository root is the only
            // place production .cs files live) must be listed by the props, except
            // the deliberately excluded RTD/COM file.
            List<string> required = Directory.GetFiles(repoRoot, "*.cs", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFullPath)
                .Where(path => !string.Equals(
                    Path.GetFileName(path), ExcludedFileName, StringComparison.OrdinalIgnoreCase))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            List<string> missing = required.Where(path => !compiled.Contains(path)).ToList();
            Assert.True(missing.Count == 0,
                "production sources missing from ProductionSources.props: "
                + string.Join(", ", missing));

            // The exclusion is intentional: listing NatsRtd.cs would pull the
            // RTD/COM surface into the unit suite and mask that gap.
            Assert.DoesNotContain(
                Path.GetFullPath(Path.Combine(repoRoot, ExcludedFileName)), compiled);
        }

        [Fact]
        public void UnitTestProject_ImportsTheSharedProductionSourcesProps()
        {
            string repoRoot = FindRepoRoot();
            string projectPath = Path.Combine(
                repoRoot, "test", "NATSExcel.Tests", "NATSExcel.Tests.csproj");
            Assert.True(File.Exists(projectPath), "test project not found at " + projectPath);

            XNamespace ns = XDocument.Load(projectPath).Root.Name.Namespace;
            bool imports = XDocument.Load(projectPath).Descendants(ns + "Import")
                .Select(e => (string)e.Attribute("Project"))
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Any(p => p.Replace('\\', '/').EndsWith(
                    "TestSupport/ProductionSources.props", StringComparison.OrdinalIgnoreCase));

            Assert.True(imports,
                "NATSExcel.Tests.csproj must Import ..\\TestSupport\\ProductionSources.props (PLAN D1)");
        }

        /// <summary>
        /// Reads the props file and resolves every <c>Compile Include</c> item to
        /// an absolute path, expanding the <c>$(MSBuildThisFileDirectory)</c>
        /// token (which points at the props file's own folder) and the simple
        /// globs the project format allows.
        /// </summary>
        private static HashSet<string> CollectCompileItems(string propsPath)
        {
            string propsDirectory = Path.GetDirectoryName(Path.GetFullPath(propsPath))
                + Path.DirectorySeparatorChar;
            var compiled = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            XDocument props = XDocument.Load(propsPath);
            XNamespace ns = props.Root.Name.Namespace;
            foreach (var compile in props.Descendants(ns + "Compile"))
            {
                var include = (string)compile.Attribute("Include");
                if (string.IsNullOrWhiteSpace(include))
                    continue; // Remove/Update items carry no Include

                string resolved = include.Replace("$(MSBuildThisFileDirectory)", propsDirectory);
                foreach (string file in ExpandCompileItem(propsDirectory, resolved))
                    compiled.Add(file);
            }

            return compiled;
        }

        /// <summary>
        /// Resolves one Compile Include value to absolute paths, expanding the
        /// simple globs the csproj format allows (for example "..\..\*.cs").
        /// </summary>
        private static IEnumerable<string> ExpandCompileItem(string baseDirectory, string include)
        {
            string normalized = include
                .Replace('\\', Path.DirectorySeparatorChar)
                .Replace('/', Path.DirectorySeparatorChar);
            if (normalized.IndexOf('*') < 0 && normalized.IndexOf('?') < 0)
            {
                yield return Path.GetFullPath(Path.Combine(baseDirectory, normalized));
                yield break;
            }

            string directoryPart = Path.GetDirectoryName(normalized);
            bool recursive = directoryPart != null
                && directoryPart.IndexOf("**", StringComparison.Ordinal) >= 0;
            if (recursive)
                directoryPart = directoryPart.Replace("**", ".");
            string searchDirectory = Path.GetFullPath(Path.Combine(
                baseDirectory,
                string.IsNullOrEmpty(directoryPart) ? "." : directoryPart));
            string filePattern = Path.GetFileName(normalized);

            foreach (string file in Directory.GetFiles(
                searchDirectory,
                filePattern,
                recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
            {
                yield return Path.GetFullPath(file);
            }
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
