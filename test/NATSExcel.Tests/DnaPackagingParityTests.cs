using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Guards the Excel-DNA packing list (PLAN F0.4 / risk R1): <c>NATSExcel.dna</c>
    /// must pack every dependency DLL the net48 build emits, because a dependency
    /// left out of the pack is only discovered as a <c>FileNotFoundException</c>
    /// inside Excel at runtime (the packed XLL loads, the UDF then explodes).
    ///
    /// This is the compile-time twin of the Excel E2E load: it proves the list is
    /// complete in both directions against the REAL build output, without Excel.
    ///
    /// A Release build is a hard prerequisite (the same convention as
    /// <see cref="ProjectFileParityTests"/>): a missing <c>bin\Release\net48</c>
    /// FAILS the test rather than skipping, so CI can never go green on a stale or
    /// absent build tree.
    /// </summary>
    public class DnaPackagingParityTests
    {
        /// <summary>The packed XLL name is derived from the .dna FILE name
        /// (risk R19); pin it so a rename cannot silently fall back to
        /// <c>NATSExcel-AddIn-packed.xll</c>.</summary>
        private const string DnaFileName = "NATSExcel.dna";

        [Fact]
        public void Dna_ListsEveryBuildOutputDllAndNoStaleEntry()
        {
            string repoRoot = FindRepoRoot();
            string outputDirectory = Path.Combine(repoRoot, "bin", "Release", "net48");
            Assert.True(Directory.Exists(outputDirectory),
                "run `dotnet build NATSExcel.csproj -c Release` first: the build output "
                + outputDirectory + " does not exist");

            List<string> outputDlls = Directory.GetFiles(outputDirectory, "*.dll",
                    SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var spec = DnaSpec.Load(Path.Combine(repoRoot, DnaFileName));
            HashSet<string> packed = spec.PackedReferences;

            // The add-in assembly itself ships as <ExternalLibrary>, not a packed
            // <Reference>, so it is expected in the output but not in the reference
            // set. Read the .dna rather than assuming: only exclude the exact
            // ExternalLibrary file name(s).
            var selfAssemblies = new HashSet<string>(spec.ExternalLibraries,
                StringComparer.OrdinalIgnoreCase);

            List<string> unpacked = outputDlls
                .Where(name => !selfAssemblies.Contains(name))
                .Where(name => !packed.Contains(name))
                .ToList();
            Assert.True(unpacked.Count == 0,
                "build output DLLs missing from " + DnaFileName + " (add them as "
                + "<Reference Path=\"...\" Pack=\"true\" /> or Excel will throw a "
                + "FileNotFoundException at runtime): " + string.Join(", ", unpacked));

            // Direction 2: a stale <Reference> (a packed name with no output DLL,
            // e.g. after a package downgrade) is also drift.
            List<string> stale = packed
                .Where(name => !outputDlls.Contains(name, StringComparer.OrdinalIgnoreCase))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            Assert.True(stale.Count == 0,
                DnaFileName + " names DLLs that are not in the build output (stale "
                + "entries): " + string.Join(", ", stale));
        }

        [Fact]
        public void Dna_NamesTheAddInAssemblyAsTheExternalLibrary()
        {
            string repoRoot = FindRepoRoot();
            var spec = DnaSpec.Load(Path.Combine(repoRoot, DnaFileName));

            // The COM server (RTD "NatsRtd") hangs off the add-in assembly, so the
            // <ExternalLibrary> must name NATSExcel.dll (risk: an empty/wrong
            // ExternalLibrary silently drops the COM server).
            Assert.Contains("NATSExcel.dll", spec.ExternalLibraries,
                StringComparer.OrdinalIgnoreCase);
        }

        [Fact]
        public void DnaFile_NameMatchesThePackedXllContract()
        {
            string repoRoot = FindRepoRoot();
            // Excel-DNA derives the packed XLL name from the .dna file name
            // (NATSExcel.dna -> NATSExcel-packed.xll). A missing/renamed .dna emits
            // NATSExcel-AddIn-packed.xll instead, which the installer, the release
            // asset list and the E2E script do not expect (risk R19).
            Assert.True(File.Exists(Path.Combine(repoRoot, DnaFileName)),
                DnaFileName + " must exist at the repository root (its file name sets "
                + "the packed XLL name: NATSExcel-packed.xll)");
        }

        [Fact]
        public void AddInAssembly_EmbedsTheConfigSchemaResource()
        {
            string repoRoot = FindRepoRoot();
            string outputDirectory = Path.Combine(repoRoot, "bin", "Release", "net48");
            string assemblyPath = Path.Combine(outputDirectory, "NATSExcel.dll");

            Assert.True(Directory.Exists(outputDirectory),
                "run `dotnet build NATSExcel.csproj -c Release` first: the build output "
                + outputDirectory + " does not exist");
            Assert.True(File.Exists(assemblyPath), "NATSExcel.dll not found at " + assemblyPath);

            // The add-in reads the JSON Schema from its own assembly
            // (ConfigSchema.ReadSchemaText), not from disk. Reflection-only load
            // avoids running any of the add-in's static initializers.
            var assembly = System.Reflection.Assembly.ReflectionOnlyLoadFrom(assemblyPath);

            // MSBuild prefixes the logical name with the root namespace, so
            // accept either the bare name or a namespace-prefixed one, exactly
            // as ConfigSchema does at runtime.
            string[] resourceNames = assembly.GetManifestResourceNames();
            bool embedded = resourceNames.Any(
                name => name.Equals("NATSExcel.schema.json", StringComparison.Ordinal)
                    || name.EndsWith(".NATSExcel.schema.json", StringComparison.Ordinal));
            Assert.True(embedded,
                "the embedded NATSExcel.schema.json resource is missing from "
                + assemblyPath + " (the add-in reads it via ConfigSchema); found: "
                + string.Join(", ", resourceNames));
        }

        /// <summary>
        /// The parsed <c>.dna</c>: the packed <c>&lt;Reference&gt;</c> file names
        /// and the <c>&lt;ExternalLibrary&gt;</c> file names.
        /// </summary>
        private sealed class DnaSpec
        {
            public HashSet<string> PackedReferences { get; private set; }
            public HashSet<string> ExternalLibraries { get; private set; }

            public static DnaSpec Load(string dnaPath)
            {
                Assert.True(File.Exists(dnaPath), DnaFileName + " not found at " + dnaPath);
                XDocument document = XDocument.Load(dnaPath);

                var references = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                // Only entries that are actually packed matter for runtime. A
                // <Reference> without Pack="true" is NOT embedded in the XLL.
                foreach (var reference in document.Descendants("Reference"))
                {
                    string path = (string)reference.Attribute("Path");
                    if (string.IsNullOrWhiteSpace(path))
                        continue;
                    string pack = (string)reference.Attribute("Pack");
                    if (!string.Equals(pack, "true", StringComparison.OrdinalIgnoreCase))
                        continue;

                    references.Add(NormalizeDll(path));
                }

                var externals = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var external in document.Descendants("ExternalLibrary"))
                {
                    string path = (string)external.Attribute("Path");
                    if (!string.IsNullOrWhiteSpace(path))
                        externals.Add(NormalizeDll(path));
                }

                return new DnaSpec
                {
                    PackedReferences = references,
                    ExternalLibraries = externals,
                };
            }

            private static string NormalizeDll(string path)
            {
                // A Path may carry a directory segment; compare on the file name.
                return Path.GetFileName(path.Replace('\\', Path.DirectorySeparatorChar));
            }
        }

        /// <summary>
        /// Walks up from the test binary until the repository root, anchored by
        /// <c>NATSExcel.csproj</c> (there is no solution file yet, PLAN F0.1).
        /// Mirrors <see cref="ProjectFileParityTests"/>.
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
