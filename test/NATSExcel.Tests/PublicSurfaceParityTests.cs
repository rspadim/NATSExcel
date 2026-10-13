using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelDna.Integration;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Freezes the add-in's public Excel surface (PLAN S7.1; AGENTS golden
    /// rule 2, "never change the public Excel surface"). It reflects over every
    /// production type that carries <c>[ExcelFunction]</c> and compares the
    /// result to the committed baseline <c>PublicSurface.baseline.txt</c>, so
    /// ANY rename, argument reorder, added/removed argument, optional/default
    /// change, description edit or volatility change FAILS the test until the
    /// baseline is intentionally refreshed.
    ///
    /// <para>
    /// Two reflection sources are used, deliberately:
    /// </para>
    /// <list type="bullet">
    /// <item>
    /// The <b>shipped add-in assembly</b> (<c>bin\Release\net48\NATSExcel.dll</c>)
    /// is the baseline's source of truth: it holds the COMPLETE surface,
    /// including the six <c>NatsRTD*</c> status helpers that live in
    /// <c>NatsRtd.cs</c>. Like <see cref="DnaPackagingParityTests"/>, a Release
    /// build is a hard prerequisite (a missing assembly fails rather than
    /// skips). The baseline is therefore generated from the real code, never
    /// hand-written.
    /// </item>
    /// <item>
    /// The <b>linked production sources in the test assembly</b> (every
    /// root <c>NatsUdf*.cs</c> is compiled in via
    /// <c>test\TestSupport\ProductionSources.props</c>) are cross-checked
    /// against the shipped assembly. If the two disagree the shipped DLL is
    /// stale: rebuild with <c>dotnet build NATSExcel.csproj -c Release</c>.
    /// <c>NatsRtd.cs</c> is intentionally not linked, which is why the RTD
    /// helpers can only be read from the shipped assembly.
    /// </item>
    /// </list>
    ///
    /// <para>
    /// <b>Refreshing the baseline.</b> A normal run never writes anything. To
    /// regenerate the baseline from the real code (after an intentional,
    /// additive surface change), set the environment variable
    /// <c>NATSEXCEL_REFRESH_PUBLIC_SURFACE=1</c> and run the unit suite; the
    /// test rewrites <c>PublicSurface.baseline.txt</c> and passes. Then review
    /// the diff and commit it on purpose:
    /// </para>
    /// <code>
    /// $env:NATSEXCEL_REFRESH_PUBLIC_SURFACE = '1'
    /// dotnet test test\NATSExcel.Tests\NATSExcel.Tests.csproj -c Release
    /// Remove-Item Env:\NATSEXCEL_REFRESH_PUBLIC_SURFACE
    /// dotnet test test\NATSExcel.Tests\NATSExcel.Tests.csproj -c Release
    /// </code>
    ///
    /// <para>
    /// <b>Counts.</b> The code truth pinned by
    /// <see cref="PublicSurface_CountsAndSplitArePinned"/> is
    /// <b>97</b> <c>[ExcelFunction]</c>s = <b>91</b> <c>NatsUDF*</c> + <b>6</b>
    /// <c>NatsRTD*</c>, with <b>25</b> <c>...NonVolatile</c> write twins (so 66
    /// <c>NatsUDF*</c> bases). <c>DESIGN.md</c> §7.9 and <c>PLAN.md</c> §1.1
    /// still say "89 UDF functions + 6 RTD = 95" because they count the two
    /// standalone conversions (<c>NatsUDFMatrixToJSON</c> /
    /// <c>NatsUDFJSONToMatrix</c>) outside the <c>NatsUDF*</c> total; both are
    /// in fact <c>NatsUDF*</c> <c>[ExcelFunction]</c>s, so the real totals are
    /// 91 and 97. This test pins the real numbers and documents the doc drift
    /// rather than weakening the assertion.
    /// </para>
    /// </summary>
    public class PublicSurfaceParityTests
    {
        private const string BaselineFileName = "PublicSurface.baseline.txt";
        private const string RefreshEnvironmentVariable = "NATSEXCEL_REFRESH_PUBLIC_SURFACE";
        private const string TwinSuffix = "NonVolatile";
        private const string NoDescription = "(no description)";

        // The real, code-derived counts (see the class remarks for the
        // DESIGN/PLAN narrative these supersede).
        private const int ExpectedTotal = 97;
        private const int ExpectedUdf = 91;
        private const int ExpectedRtd = 6;
        private const int ExpectedTwins = 25;

        /// <summary>The six frozen <c>NatsRTD*</c> status helpers (DESIGN §7.9),
        /// pinned by name so a rename cannot slip through the count alone.</summary>
        private static readonly string[] ExpectedRtdNames =
        {
            "NatsRTDConnectionCount",
            "NatsRTDTopicCount",
            "NatsRTDSubscriptionCount",
            "NatsRTDSubjectCount",
            "NatsRTDDefaultHost",
            "NatsRTDCluster",
        };

        [Fact]
        public void PublicSurface_MatchesTheCommittedBaseline()
        {
            List<FunctionEntry> shipped = CollectSurface(ShippedAssembly());

            List<string> actualLines = shipped
                .Select(entry => entry.Line)
                .OrderBy(line => line, StringComparer.Ordinal)
                .ToList();

            string baselinePath = BaselinePath();
            if (IsRefreshRequested())
            {
                WriteBaseline(baselinePath, actualLines);
                return;
            }

            Assert.True(File.Exists(baselinePath),
                "the public-surface baseline is missing: " + baselinePath
                + " (regenerate it with " + RefreshEnvironmentVariable + "=1)");

            List<string> baselineLines = ReadBaseline(baselinePath);
            Assert.True(baselineLines.SequenceEqual(actualLines, StringComparer.Ordinal),
                BuildDiff(baselineLines, actualLines));

            // The linked production sources must agree with the shipped assembly
            // for every NatsUDF* function; a disagreement means the build output
            // is stale relative to the sources.
            Dictionary<string, string> linked = CollectSurface(LinkedAssembly())
                .ToDictionary(entry => entry.Name, entry => entry.Line, StringComparer.Ordinal);

            foreach (FunctionEntry entry in shipped.Where(
                item => item.Name.StartsWith("NatsUDF", StringComparison.Ordinal)))
            {
                Assert.True(linked.TryGetValue(entry.Name, out string linkedLine),
                    "the shipped add-in declares '" + entry.Name
                    + "' but the linked production sources do not. Rebuild the add-in "
                    + "(`dotnet build NATSExcel.csproj -c Release`) and re-run.");

                Assert.True(string.Equals(entry.Line, linkedLine, StringComparison.Ordinal),
                    "the shipped add-in and the linked production sources disagree for '"
                    + entry.Name + "'. Rebuild the add-in "
                    + "(`dotnet build NATSExcel.csproj -c Release`) and re-run.\n"
                    + "  shipped: " + entry.Line + "\n"
                    + "  linked : " + linkedLine);
            }
        }

        [Fact]
        public void PublicSurface_CountsAndSplitArePinned()
        {
            List<FunctionEntry> shipped = CollectSurface(ShippedAssembly());
            Dictionary<string, FunctionEntry> byName = shipped
                .ToDictionary(entry => entry.Name, entry => entry, StringComparer.Ordinal);

            int total = shipped.Count;
            int udf = shipped.Count(entry => entry.Name.StartsWith("NatsUDF", StringComparison.Ordinal));
            int rtd = shipped.Count(entry => entry.Name.StartsWith("NatsRTD", StringComparison.Ordinal));

            List<FunctionEntry> twins = shipped
                .Where(entry => entry.Name.EndsWith(TwinSuffix, StringComparison.Ordinal))
                .ToList();

            Assert.Equal(ExpectedTotal, total);
            Assert.Equal(ExpectedUdf, udf);
            Assert.Equal(ExpectedRtd, rtd);
            Assert.Equal(ExpectedTwins, twins.Count);

            // The split must be exhaustive and disjoint: every function is either
            // a NatsUDF* or a NatsRTD* (nothing else is registered).
            Assert.Equal(total, udf + rtd);

            // Pin the six RTD helper names, not just the count.
            string[] rtdNames = shipped
                .Where(entry => entry.Name.StartsWith("NatsRTD", StringComparison.Ordinal))
                .Select(entry => entry.Name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(
                ExpectedRtdNames.OrderBy(name => name, StringComparer.Ordinal).ToArray(),
                rtdNames);

            // Every twin has a base with an IDENTICAL argument signature, the
            // base is volatile and the twin is not (a twin whose signature
            // differs from its base is a real surface defect, not test noise).
            foreach (FunctionEntry twin in twins)
            {
                string baseName = twin.Name.Substring(0, twin.Name.Length - TwinSuffix.Length);
                Assert.True(byName.ContainsKey(baseName),
                    "the NonVolatile twin '" + twin.Name + "' has no matching base '" + baseName + "'");

                FunctionEntry @base = byName[baseName];
                Assert.True(string.Equals(@base.ArgsSignature, twin.ArgsSignature, StringComparison.Ordinal),
                    "the twin '" + twin.Name + "' has a different argument signature from its base '"
                    + baseName + "'.\n  base : " + @base.ArgsSignature + "\n  twin : " + twin.ArgsSignature);

                Assert.True(@base.IsVolatile, "the base '" + baseName + "' of a NonVolatile twin must be volatile");
                Assert.False(twin.IsVolatile, "the twin '" + twin.Name + "' must not be volatile");
            }
        }

        /// <summary>One <c>[ExcelFunction]</c> as it appears in the baseline.</summary>
        private sealed class FunctionEntry
        {
            public string Name;
            public string Line;
            public string ArgsSignature;
            public bool IsVolatile;

            public static FunctionEntry Create(MethodInfo method, ExcelFunctionAttribute attribute)
            {
                ParameterInfo[] parameters = method.GetParameters();
                string args = parameters.Length == 0
                    ? "(no arguments)"
                    : string.Join("; ", parameters.Select(FormatParameter));

                string line = method.Name
                    + " | volatile=" + (attribute.IsVolatile ? "true" : "false")
                    + " | " + args
                    + " | \"" + Escape(attribute.Description ?? NoDescription) + "\"";

                return new FunctionEntry
                {
                    Name = method.Name,
                    Line = line,
                    ArgsSignature = args,
                    IsVolatile = attribute.IsVolatile,
                };
            }

            private static string FormatParameter(ParameterInfo parameter)
            {
                var argumentAttribute = parameter
                    .GetCustomAttributes(typeof(ExcelArgumentAttribute), false)
                    .OfType<ExcelArgumentAttribute>()
                    .FirstOrDefault();

                string optional = string.Empty;
                if (parameter.IsOptional)
                {
                    optional = parameter.HasDefaultValue
                        ? "=" + FormatDefault(parameter.DefaultValue)
                        : "=?";
                }

                return parameter.Name
                    + ":" + TypeName(parameter.ParameterType)
                    + optional
                    + " \"" + Escape(argumentAttribute?.Description ?? NoDescription) + "\"";
            }

            private static string TypeName(Type type)
            {
                if (type == typeof(object))
                    return "object";
                if (type == typeof(string))
                    return "string";
                if (type == typeof(double))
                    return "double";
                if (type == typeof(int))
                    return "int";
                if (type == typeof(bool))
                    return "bool";
                if (type.IsArray)
                    return TypeName(type.GetElementType()) + "[" + new string(',', type.GetArrayRank() - 1) + "]";
                return type.Name;
            }

            private static string FormatDefault(object value)
            {
                if (value == null)
                    return "null";
                if (value is string text)
                    return "\"" + Escape(text) + "\"";
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "(null)";
            }

            private static string Escape(string value)
            {
                return value
                    .Replace("\\", "\\\\")
                    .Replace("\"", "\\\"")
                    .Replace("\r", "\\r")
                    .Replace("\n", "\\n");
            }
        }

        private static List<FunctionEntry> CollectSurface(Assembly assembly)
        {
            var entries = new List<FunctionEntry>();
            Type attributeType = typeof(ExcelFunctionAttribute);

            foreach (Type type in SafeGetTypes(assembly))
            {
                foreach (MethodInfo method in type.GetMethods(
                    BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly))
                {
                    var attributes = method.GetCustomAttributes(attributeType, false);
                    if (attributes.Length == 0)
                        continue;

                    entries.Add(FunctionEntry.Create(method, (ExcelFunctionAttribute)attributes[0]));
                }
            }

            return entries;
        }

        private static IEnumerable<Type> SafeGetTypes(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                return ex.Types.Where(type => type != null);
            }
        }

        private static Assembly LinkedAssembly()
        {
            return typeof(PublicSurfaceParityTests).Assembly;
        }

        private static Assembly ShippedAssembly()
        {
            string path = Path.Combine(
                FindRepoRoot(), "bin", "Release", "net48", "NATSExcel.dll");
            Assert.True(File.Exists(path),
                "a Release build of the add-in is required to freeze the public surface "
                + "(including the NatsRtd.cs helpers): run "
                + "`dotnet build NATSExcel.csproj -c Release` first. Missing: " + path);
            return Assembly.LoadFrom(path);
        }

        private static bool IsRefreshRequested()
        {
            string value = Environment.GetEnvironmentVariable(RefreshEnvironmentVariable);
            return !string.IsNullOrEmpty(value)
                && (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase));
        }

        private static string BaselinePath()
        {
            return Path.Combine(
                FindRepoRoot(), "test", "NATSExcel.Tests", BaselineFileName);
        }

        private static IEnumerable<string> BaselineHeader()
        {
            yield return "# NATSExcel public Excel surface baseline (PLAN S7.1).";
            yield return "# Frozen at the first release; additive changes only (AGENTS golden rule 2).";
            yield return "# Generated from the built add-in assembly + the linked production sources;";
            yield return "# never hand-write this file. Checked by PublicSurfaceParityTests.";
            yield return "# Refresh deliberately: set " + RefreshEnvironmentVariable + "=1 and run";
            yield return "#   dotnet test test\\NATSExcel.Tests\\NATSExcel.Tests.csproj -c Release";
            yield return "#";
            yield return "# Every non-comment line is:";
            yield return "#   <name> | volatile=<true|false> | <param:type[=default] \"<arg description>\"; ...> | \"<function description>\"";
            yield return "# Sorted ordinal by function name. No machine-specific data.";
        }

        private static void WriteBaseline(string path, List<string> lines)
        {
            string content = string.Join("\n", BaselineHeader().Concat(lines)) + "\n";
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content, new UTF8Encoding(false));
        }

        private static List<string> ReadBaseline(string path)
        {
            return File.ReadAllLines(path)
                .Select(line => line.TrimEnd('\r'))
                .Where(line => line.Length > 0 && !line.StartsWith("#", StringComparison.Ordinal))
                .ToList();
        }

        private static Dictionary<string, string> ToNameMap(IEnumerable<string> lines)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string line in lines)
            {
                int separator = line.IndexOf(" | ", StringComparison.Ordinal);
                string name = separator < 0 ? line : line.Substring(0, separator);
                map[name] = line;
            }
            return map;
        }

        private static string BuildDiff(List<string> baseline, List<string> actual)
        {
            Dictionary<string, string> before = ToNameMap(baseline);
            Dictionary<string, string> after = ToNameMap(actual);

            var builder = new StringBuilder();
            builder.AppendLine("The public Excel surface drifted from " + BaselineFileName + ".");
            builder.AppendLine("If the change is intentional (additive only!), refresh the baseline: set "
                + RefreshEnvironmentVariable + "=1 and re-run this test.");

            foreach (string name in after.Keys.Except(before.Keys).OrderBy(key => key, StringComparer.Ordinal))
                builder.AppendLine("  ADDED   " + after[name]);

            foreach (string name in before.Keys.Except(after.Keys).OrderBy(key => key, StringComparer.Ordinal))
                builder.AppendLine("  REMOVED " + before[name]);

            foreach (string name in after.Keys.Intersect(before.Keys).OrderBy(key => key, StringComparer.Ordinal))
            {
                if (string.Equals(after[name], before[name], StringComparison.Ordinal))
                    continue;
                builder.AppendLine("  CHANGED " + name);
                builder.AppendLine("    baseline: " + before[name]);
                builder.AppendLine("    actual  : " + after[name]);
            }

            return builder.ToString();
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
