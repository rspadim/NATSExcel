using System;
using System.IO;
using System.Linq;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Round-2 regression pins for the invariants that live in code the unit
    /// project cannot reference. <c>NatsRtd.cs</c> is excluded from every test
    /// project (it is the Excel/COM surface, covered by the Excel E2E), so the
    /// two RTD fixes from the round-2 audit are pinned here by a source scan:
    /// they are structural invariants of one file, and a scan is the only
    /// automated check available without Excel.
    ///
    /// A scan is a weak test (it pins shape, not behaviour) but it is strictly
    /// better than nothing: the alternative is that a future edit silently
    /// reverts the fix and no layer notices.
    /// </summary>
    public class RtdInvariantTests
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

        /// <summary>
        /// KVGET/MSG commit BYTES on success and TEXT on failure. Both markers
        /// being live at once made a recovery (the JS-off branch) or a repeated
        /// outage invisible: the cell kept a stale #ERROR/value forever. The fix
        /// is that each commit clears the OTHER marker, so exactly one is live.
        /// </summary>
        [Fact]
        public void CommitPolledBytesAndText_ClearTheOppositeMarker()
        {
            string src = Read("NatsRtd.cs");

            string bytes = ExtractMethod(src, "public void CommitPolledBytes(byte[] value)");
            string text = ExtractMethod(src, "public void CommitPolledText(string value)");

            Assert.Contains("_hasLastPolledBytes = true;", bytes);
            Assert.Contains("_hasLastPolledText = false;", bytes);
            Assert.Contains("_hasLastPolledText = true;", text);
            Assert.Contains("_hasLastPolledBytes = false;", text);
        }

        /// <summary>
        /// The RTD topic registry is capped, so a runaway sheet cannot register
        /// an unbounded number of topics (each with a subscription or a poll
        /// entry) and degrade Excel silently.
        /// </summary>
        [Fact]
        public void TheRtdTopicRegistryIsCapped()
        {
            string src = Read("NatsRtd.cs");
            Assert.Contains("MaxTopics", src);
            Assert.Contains("too many RTD topics", src);
        }

        /// <summary>The RTD polled KVKEYS list is bounded BEFORE it is sorted and
        /// serialised: a bucket with millions of subjects must not rebuild and
        /// JSON-encode the whole set every tick (the cell shows 255 chars).</summary>
        [Fact]
        public void RtdKvKeysIsCappedBeforeSorting()
        {
            string src = Read("NatsRtd.cs");
            Assert.Contains("MaxRtdKeys", src);

            int cap = src.IndexOf("keys.Count >= MaxRtdKeys", StringComparison.Ordinal);
            int sort = src.IndexOf("keys.Sort(StringComparer.Ordinal)", StringComparison.Ordinal);
            Assert.True(cap > 0, "the KVKEYS cap is gone");
            Assert.True(sort > 0, "the KVKEYS sort is gone");
            Assert.True(cap < sort, "the KVKEYS cap must run BEFORE the sort");
        }

        /// <summary>Every RTD cell goes through the 255-char bound, not the
        /// 32767 one: Excel-DNA's Topic.FixValue truncates at 255 with no marker,
        /// so the suffix must be applied here.</summary>
        [Fact]
        public void RtdCellTextUsesTheRtdBound()
        {
            string src = Read("NatsRtd.cs");
            Assert.Contains("NatsResultFormatter.BoundRtdCellText(value)", src);
            Assert.DoesNotContain("NatsResultFormatter.BoundCellText(value)", src);
        }

        /// <summary>
        /// A reload that snapshotted an RTD instance must not resurrect its
        /// timers after Teardown (a stranded timer roots the object and no later
        /// teardown can reach it).
        /// </summary>
        [Fact]
        public void ApplyConfigBailsOutAfterTeardown()
        {
            string src = Read("NatsRtd.cs");
            Assert.Contains("private volatile bool _terminated;", src);

            string teardown = ExtractMethod(src, "private void Teardown()");
            Assert.Contains("_terminated = true;", teardown);

            string apply = ExtractMethod(src, "private void ApplyConfig()");
            Assert.Contains("if (_terminated)", apply);
            // The guard must appear again before the timer swap: a teardown can
            // start between the entry check and the recreation.
            int guards = CountOccurrences(apply, "if (_terminated)");
            Assert.True(guards >= 2, "ApplyConfig needs the guard both at entry and before the timer swap");
        }

        // ------------------------------------------------------------------

        private static string ExtractMethod(string source, string signature)
        {
            int start = source.IndexOf(signature, StringComparison.Ordinal);
            Assert.True(start >= 0, "method not found: " + signature);
            // Walk braces from the first '{' after the signature.
            int open = source.IndexOf('{', start);
            Assert.True(open >= 0, "no body for " + signature);
            int depth = 0;
            for (int i = open; i < source.Length; i++)
            {
                if (source[i] == '{') depth++;
                else if (source[i] == '}')
                {
                    depth--;
                    if (depth == 0)
                        return source.Substring(open, i - open + 1);
                }
            }
            Assert.True(false, "unterminated body for " + signature);
            return null;
        }

        private static int CountOccurrences(string text, string needle)
        {
            int n = 0, i = 0;
            while ((i = text.IndexOf(needle, i, StringComparison.Ordinal)) >= 0) { n++; i += needle.Length; }
            return n;
        }
    }
}
