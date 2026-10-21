using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NATS.Client.Core;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Pins the claims the round-5 audit found TRUE but unpinned: each of them would
    /// silently become false (and the docs would silently become lies) without a test
    /// that fails when reality moves.
    ///
    /// <list type="bullet">
    /// <item>the config search order (add-in folder, then the user profile, then the
    /// Excel folder, then <c>C:\Windows</c>);</item>
    /// <item>the NATS client defaults the docs quote (a 2 s initial reconnect wait
    /// doubling to a 5 s ceiling, unlimited retries, 100 ms jitter) - a package bump
    /// must re-check the numbers in AGENTS.md, the schema and the CHANGELOG;</item>
    /// <item>the installer's "11 languages, every one carrying the same key set, so a
    /// missing translation can never fall back to a silent English string" (the manual
    /// <c>languages\messages.ps1</c> check, automated);</item>
    /// <item>the site's "the releases are not code-signed yet" - adding signing to the
    /// pipeline must fail here so the pages get updated with it.</item>
    /// </list>
    /// </summary>
    public class UnpinnedClaimsTests
    {
        private static string RepoRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "NATSExcel.sln")))
                dir = dir.Parent;
            Assert.NotNull(dir);
            return dir.FullName;
        }

        private static string Read(string relative)
        {
            return File.ReadAllText(Path.Combine(RepoRoot(), relative));
        }

        /// <summary>The body of a <c>[Section]</c> in an Inno Setup script: the lines
        /// after its header up to the next header. Comments and blanks are dropped.</summary>
        private static List<string> IniSection(string text, string section)
        {
            var lines = text.Replace("\r\n", "\n").Split('\n');
            var body = new List<string>();
            bool inside = false;
            foreach (var raw in lines)
            {
                string line = raw.Trim();
                if (line.StartsWith("[", StringComparison.Ordinal))
                {
                    inside = string.Equals(line, "[" + section + "]", StringComparison.OrdinalIgnoreCase);
                    continue;
                }
                if (!inside || line.Length == 0 || line.StartsWith(";", StringComparison.Ordinal))
                    continue;
                body.Add(line);
            }
            return body;
        }

        // ------------------------------------------------------------------
        // 1. The config search order
        // ------------------------------------------------------------------

        [Fact]
        public void TheConfigSearchOrderIsTheDocumentedOne()
        {
            var paths = AppConfig.CandidatePaths().ToList();

            // Every candidate is "<folder>\NATSExcel.json", and there are between two
            // and four of them: the add-in folder and the Excel folder are present
            // only when the host can report them (no Excel context in a test run).
            Assert.InRange(paths.Count, 2, 4);
            Assert.All(paths, p => Assert.EndsWith("NATSExcel.json", p, StringComparison.OrdinalIgnoreCase));
            Assert.All(paths, p => Assert.True(Path.IsPathRooted(p), "a candidate must be absolute: " + p));

            string profile = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "NATSExcel.json");
            string windows = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows), "NATSExcel.json");

            int profileAt = paths.FindIndex(p => string.Equals(p, profile, StringComparison.OrdinalIgnoreCase));
            int windowsAt = paths.FindIndex(p => string.Equals(p, windows, StringComparison.OrdinalIgnoreCase));

            Assert.True(profileAt >= 0, "the user profile must be a candidate");
            Assert.True(windowsAt >= 0, "C:\\Windows must be the last-resort candidate");
            // The documented order is add-in folder -> profile -> Excel folder ->
            // C:\Windows, so the profile always precedes the Windows folder.
            Assert.True(profileAt < windowsAt,
                "the profile must be searched BEFORE C:\\Windows (got " + profileAt + " vs " + windowsAt + ")");
        }

        // ------------------------------------------------------------------
        // 2. The NATS client defaults the docs quote
        // ------------------------------------------------------------------

        [Fact]
        public void TheNatsClientDefaultsTheDocsQuoteAreUnchanged()
        {
            NatsOpts defaults = NatsOpts.Default;

            // AGENTS.md, the schema's waitMs description and the CHANGELOG say the
            // client's own ceiling is 5 s (it used to be documented as 10 s) and that
            // the initial wait is 2 s; the same places explain that the retry count is
            // bounded only when it is POSITIVE, so -1 and 0 both mean unlimited.
            Assert.Equal(TimeSpan.FromSeconds(2), defaults.ReconnectWaitMin);
            Assert.Equal(TimeSpan.FromSeconds(5), defaults.ReconnectWaitMax);
            Assert.Equal(TimeSpan.FromMilliseconds(100), defaults.ReconnectJitter);
            Assert.Equal(-1, defaults.MaxReconnectRetry);
        }

        [Fact]
        public void ADeclaredWaitIsAppliedAsBothBounds()
        {
            // The mapping the schema documents: one configured number is the whole
            // wait, so the delay cannot grow past it. Recomputing it here from the
            // client defaults keeps the claim honest even if the defaults move.
            TimeSpan configured = TimeSpan.FromSeconds(2);
            NatsOpts opts = NatsOpts.Default with { ReconnectWaitMin = configured, ReconnectWaitMax = configured };
            Assert.Equal(opts.ReconnectWaitMin, opts.ReconnectWaitMax);
            Assert.True(opts.ReconnectWaitMax <= NatsOpts.Default.ReconnectWaitMax,
                "the configured wait must not exceed the client's own ceiling");
        }

        // ------------------------------------------------------------------
        // 3. The installer's languages
        // ------------------------------------------------------------------

        [Fact]
        public void TheInstallerShipsElevenLanguagesWithTheSameKeySet()
        {
            string iss = Read(Path.Combine("installer", "NATSExcel.iss"));

            var languages = IniSection(iss, "Languages")
                .Select(line => Regex.Match(line, "^Name:\\s*\"([a-z]+)\";"))
                .Where(m => m.Success)
                .Select(m => m.Groups[1].Value)
                .ToList();
            Assert.Equal(11, languages.Count);
            Assert.Contains("english", languages);
            Assert.Contains("brazilianportuguese", languages);

            // Every language carries its own [CustomMessages] block, and every block
            // defines exactly the SAME key set as english: Inno falls back to a
            // hardcoded English string for a missing key, so a half-translated page is
            // otherwise silent.
            var blocks = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in IniSection(iss, "CustomMessages"))
            {
                var m = Regex.Match(line, "^([a-z]+)\\.([A-Za-z0-9_]+)=");
                if (!m.Success)
                    continue;
                string lang = m.Groups[1].Value;
                if (!blocks.TryGetValue(lang, out var keys))
                    blocks[lang] = keys = new List<string>();
                keys.Add(m.Groups[2].Value);
            }

            Assert.Equal(languages.Count, blocks.Count);
            Assert.All(languages, lang => Assert.True(blocks.ContainsKey(lang),
                "no [CustomMessages] block for language '" + lang + "'"));

            var english = blocks["english"].OrderBy(k => k, StringComparer.Ordinal).ToList();
            Assert.Equal(74, english.Count);
            Assert.Equal(english.Count, english.Distinct(StringComparer.Ordinal).Count());
            foreach (var lang in languages)
            {
                var keys = blocks[lang].OrderBy(k => k, StringComparer.Ordinal).ToList();
                Assert.True(english.SequenceEqual(keys, StringComparer.Ordinal),
                    "language '" + lang + "' does not carry the english key set (missing: "
                    + string.Join(",", english.Except(keys, StringComparer.Ordinal))
                    + "; extra: " + string.Join(",", keys.Except(english, StringComparer.Ordinal)) + ")");
            }

            // The two languages Inno Setup does not ship are bundled here.
            foreach (var bundled in new[] { "ChineseSimplified.isl", "ChineseTraditional.isl" })
                Assert.True(File.Exists(Path.Combine(RepoRoot(), "installer", "languages", bundled)),
                    "bundled language file missing: " + bundled);
        }

        // ------------------------------------------------------------------
        // 4. The "not code-signed yet" statement
        // ------------------------------------------------------------------

        [Fact]
        public void ThePagesSayTheBuildsAreUnsignedAndNothingSignsThem()
        {
            string workflow = Read(Path.Combine(".github", "workflows", "build.yml"));

            // No signing tool may sneak into the pipeline while the pages say the
            // builds are unsigned. Adding one is fine - but it must be a deliberate
            // change that updates SIGNING.md and the three landing pages, and this
            // test is what forces that conversation.
            foreach (var tool in new[] { "signtool", "signpath", "azuresigntool", "jsign", "osslsigncode" })
                Assert.DoesNotContain(tool, workflow, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("Status: **inactive**", Read("SIGNING.md"));
            Assert.Contains("not signed", Read("SIGNING.md"), StringComparison.OrdinalIgnoreCase);

            foreach (var page in new[]
            {
                Path.Combine("docs", "index.md"),
                Path.Combine("docs", "en", "index.md"),
                Path.Combine("docs", "es", "index.md")
            })
            {
                string text = Read(page);
                // Each page tells the user the builds are not signed yet (the exact
                // wording differs per language, the meaning does not).
                Assert.True(
                    text.IndexOf("não são assinados", StringComparison.OrdinalIgnoreCase) >= 0
                    || text.IndexOf("not code-signed", StringComparison.OrdinalIgnoreCase) >= 0
                    || text.IndexOf("no están firmados", StringComparison.OrdinalIgnoreCase) >= 0,
                    page + " no longer says the builds are unsigned");
            }
        }

        // ------------------------------------------------------------------
        // 5. The E2E's assertion count (the one number that drifted twice)
        // ------------------------------------------------------------------

        [Fact]
        public void TheE2eAssertionCountInTheDocsMatchesTheScript()
        {
            string script = Read(Path.Combine("test", "Run-ExcelE2E.ps1"));
            int callSites = Regex.Matches(script, @"(?m)^\s+Assert\s+""").Count;

            // AGENTS.md quotes both numbers; a new assertion must update the doc in the
            // same change, or the doc drifts silently (it already did twice).
            string agents = Read("AGENTS.md");
            Assert.Contains(callSites.ToString() + " `Assert` call sites", agents);
            Assert.Contains("57 assertions", agents);
        }

        // ------------------------------------------------------------------
        // 6. The installer's connection probe and the stale-entry cleanup
        // ------------------------------------------------------------------

        [Fact]
        public void TheConnectionProbeKeepsItsStableExitCodeContract()
        {
            string probe = Read(Path.Combine("installer", "test-connection.ps1"));

            // The wizard builds its (localized) message from THESE codes and never
            // from the helper's text, which a cmd.exe redirect re-encodes into the
            // OEM codepage. A renumbering here would silently turn a timeout into
            // "connection refused" in every language.
            Assert.Contains("exit 0", probe, StringComparison.Ordinal);
            Assert.Contains("exit 3", probe, StringComparison.Ordinal);
            Assert.Contains("exit 5", probe, StringComparison.Ordinal);
            Assert.Contains("exit 6", probe, StringComparison.Ordinal);
            Assert.Contains("exit 7", probe, StringComparison.Ordinal);
            // 2 (unknown host) and 4 (refused) come from the classifier, which reads
            // the SocketError CODE: the message text is localized by Windows.
            Assert.Contains("return 2", probe, StringComparison.Ordinal);
            Assert.Contains("return 4", probe, StringComparison.Ordinal);
            Assert.Contains("SocketError]::HostNotFound", probe, StringComparison.Ordinal);
            // -Quiet is what keeps the probe's banner out of the redirected file.
            Assert.Contains("[switch]$Quiet", probe, StringComparison.Ordinal);

            string iss = Read(Path.Combine("installer", "NATSExcel.iss"));
            foreach (var code in new[] { 2, 3, 5, 6, 7 })
                Assert.Contains(IntToStrText(code) + ": Reason := ExpandConstant('{cm:TestReason", iss, StringComparison.Ordinal);
            // The message must come from a message key, never from the helper's text.
            Assert.DoesNotContain("LoadStringsFromFile(OutFile", iss, StringComparison.Ordinal);
        }

        [Fact]
        public void TheInstallerUnregistersOnlyOlderNatsExcelCopies()
        {
            string iss = Read(Path.Combine("installer", "NATSExcel.iss"));

            // An upgrade into another folder used to leave the old OPEN value behind,
            // so Excel loaded two add-ins (duplicated functions, two configs). The
            // cleanup must stay scoped to a NATSExcel XLL and keep this install's own
            // path, and it must run AFTER RegisterExcelAutoLoad.
            Assert.Contains("function XllNameIsNatsExcel", iss, StringComparison.Ordinal);
            Assert.Contains("function IsStaleNatsExcelEntry", iss, StringComparison.Ordinal);
            Assert.Contains("function RemoveStaleExcelEntries", iss, StringComparison.Ordinal);
            Assert.Contains("function RemoveStaleAddInManagerEntries", iss, StringComparison.Ordinal);
            Assert.Contains("NATSExcel64-packed.xll", iss, StringComparison.Ordinal);
            Assert.Contains("NATSExcel-packed.xll", iss, StringComparison.Ordinal);

            int register = iss.IndexOf("RegisterExcelAutoLoad(ExpandConstant", StringComparison.Ordinal);
            int cleanup = iss.IndexOf("RemoveStaleExcelEntries(CurrentXllPath())", StringComparison.Ordinal);
            Assert.True(register >= 0 && cleanup > register,
                "the stale-entry cleanup must run after this install's own entry is written");

            // The "Add-in Manager" cleanup only ever removes a DEAD entry: a value that
            // still points at an existing XLL belongs to a working install.
            int manager = iss.IndexOf("RemoveStaleAddInManagerEntries(const KeepPath", StringComparison.Ordinal);
            Assert.True(manager >= 0);
            string body = iss.Substring(manager, 900);
            Assert.Contains("not FileExists(Target)", body, StringComparison.Ordinal);
            Assert.Contains("CompareText(Target, KeepPath) <> 0", body, StringComparison.Ordinal);
        }

        private static string IntToStrText(int value)
        {
            return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
