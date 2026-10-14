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
        // Round-4 pins. The round-4 audit found fixes from earlier rounds that no
        // layer would notice a regression in, plus two defects in this file (a
        // credential path that bypassed masking and a teardown/reload race that
        // could strand a live timer pair). The file still cannot be referenced by
        // any test project, so these are source scans - weak by construction, and
        // documented as such above.
        // ------------------------------------------------------------------

        /// <summary>
        /// The caps must be exact values, not just present identifiers: a scan for
        /// "MaxRtdKeys" also passes if the cap is raised to int.MaxValue, which
        /// would remove the protection while keeping the name.
        /// </summary>
        [Fact]
        public void TheRtdCapsAreBoundedToExactValues()
        {
            string src = Read("NatsRtd.cs");
            Assert.Contains("const int MaxTopics = 10000;", src);
            Assert.Contains("const int MaxRtdKeys = 5000;", src);
        }

        /// <summary>
        /// The polled cluster cell carries SERVER-SUPPLIED text in every branch: the
        /// cluster name, the advertised connect URLs (which embed a userinfo when
        /// the server was configured with one) and the raw $SYS ping reply. The UDF
        /// side masks the same values; this cell must not be the one surface that
        /// prints a credential.
        /// </summary>
        [Fact]
        public void PollClusterMasksTheServerSuppliedText()
        {
            string poll = ExtractMethod(Read("NatsRtd.cs"), "private void PollCluster(");

            Assert.Contains("AppConfig.MaskNatsEndpoint(", poll);

            // Before the dedup comparison, so the deduped value is what was delivered.
            int mask = poll.IndexOf("AppConfig.MaskNatsEndpoint(", StringComparison.Ordinal);
            int dedup = poll.IndexOf("HasChangedPolledText(", StringComparison.Ordinal);
            Assert.True(dedup > 0, "the polled cluster dedup is gone");
            Assert.True(mask < dedup, "the cluster text must be masked BEFORE the dedup compares it");
        }

        /// <summary>
        /// A poll error must clear the value markers, so a recovery to the payload
        /// the cell held before the blip is still pushed instead of being suppressed
        /// as a repeat. It has to stay AFTER the error's own dedup (or a repeated
        /// outage would re-push its error every tick) and BEFORE the error is pushed
        /// (or the next successful read would be deduped against the error text).
        /// </summary>
        [Fact]
        public void ThePollErrorPathClearsTheValueMarkers()
        {
            string poll = ExtractMethod(Read("NatsRtd.cs"), "private void PollEndpoint(");

            int dedup = poll.IndexOf("HasChangedPolledText(error)", StringComparison.Ordinal);
            int invalidate = poll.IndexOf("td.InvalidatePolledMarkers();", StringComparison.Ordinal);
            int push = poll.IndexOf("PushPolled(td, error", StringComparison.Ordinal);

            Assert.True(dedup > 0, "the poll ERROR dedup is gone");
            Assert.True(invalidate > 0, "the poll error path no longer clears the value markers");
            Assert.True(push > 0, "the poll error is no longer pushed");
            Assert.True(dedup < invalidate, "the error must be deduped BEFORE the markers are cleared");
            Assert.True(invalidate < push, "the markers must be cleared BEFORE the error is pushed");
        }

        /// <summary>
        /// The reload path still re-applies the RTD tuning and still re-checks the
        /// tombstone from BOTH sides of the timer swap. DisposeTimer nulls the field
        /// before disposing, so a Teardown starting after the pre-swap check would
        /// have disposed null and left the brand-new pair running on an instance no
        /// longer in Instances - a live timer rooting the whole server, unreachable
        /// by any later teardown.
        /// </summary>
        [Fact]
        public void ApplyConfigReChecksTheTombstoneAfterStartingTheNewTimers()
        {
            string src = Read("NatsRtd.cs");
            Assert.Contains("NatsRuntime.ConfigReloadHook = RefreshFromConfig;", src);

            string apply = ExtractMethod(src, "private void ApplyConfig()");
            Assert.True(CountOccurrences(apply, "if (_terminated)") >= 3,
                "ApplyConfig must re-check the tombstone after publishing and starting the new pair, not only before the swap");

            int started = apply.IndexOf("_excelTimer.Start();", StringComparison.Ordinal);
            Assert.True(started > 0, "ApplyConfig no longer starts the excel timer");
            int lastGuard = apply.LastIndexOf("if (_terminated)", StringComparison.Ordinal);
            Assert.True(lastGuard > started, "the final _terminated check must come AFTER the timers are started");

            // And it must unwind its OWN pair, not merely return.
            string tail = apply.Substring(lastGuard);
            Assert.Contains("DisposeTimer(ref _natsTimer);", tail);
            Assert.Contains("DisposeTimer(ref _excelTimer);", tail);
        }

        /// <summary>
        /// The RTDCLUSTER vocabulary error is user-visible text that lives in this
        /// file only (the parser tests cover the grammar, not the runtime reply), so
        /// nothing else can pin it.
        /// </summary>
        [Fact]
        public void TheClusterVocabularyErrorTextIsFrozen()
        {
            Assert.Contains(
                "RTDCLUSTER expects what to be one of [cluster, nodes, leader, lag]",
                Read("NatsRtd.cs"));
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
