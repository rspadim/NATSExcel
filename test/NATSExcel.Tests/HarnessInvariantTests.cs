using System;
using System.IO;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The test harness is code too, and its failure mode is destructive: a script
    /// that stops "whatever is listening on the port" or quits "an EXCEL.EXE" kills
    /// a developer's own server or an unsaved workbook. These are source scans
    /// (PowerShell cannot be referenced from here), and each one pins a safety
    /// contract that a past incident cost real time to diagnose:
    ///
    /// <list type="bullet">
    /// <item>a run against port 4222 answered by somebody else's long-lived
    /// nats-server, whose failover step then tried to kill that server and died on
    /// "Access denied";</item>
    /// <item>the crashed run's HIDDEN Excel keeping the copied XLL loaded, so every
    /// later run failed on "being used by another process" naming no culprit.</item>
    /// </list>
    ///
    /// A scan is a weak test, but the alternative is nothing, and the last two
    /// harness regressions were exactly of this shape.
    /// </summary>
    public class HarnessInvariantTests
    {
        private static string Read(string relative)
        {
            return File.ReadAllText(Path.Combine(RepoRoot(), relative));
        }

        internal static string RepoRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "NATSExcel.sln")))
                dir = dir.Parent;
            Assert.NotNull(dir);
            return dir.FullName;
        }

        /// <summary>
        /// The E2E's failover journey kills the node it finds on the port, so it must
        /// prove the node is one of OURS (the command line names a nats.conf under
        /// our temp work root) before touching it - and it must skip, not throw,
        /// when it cannot.
        /// </summary>
        [Fact]
        public void TheE2EProvesOwnershipBeforeItKillsANode()
        {
            string src = Read(Path.Combine("test", "Run-ExcelE2E.ps1"));

            Assert.Contains("function Test-OwnedTestServer(", src);
            Assert.Contains("-not (Test-OwnedTestServer $port)", src);
            Assert.Contains("not a disposable test node", src);

            // A denied kill (another user / a protected process) must be caught: the
            // journey is the LAST step, so an unhandled error there skips the summary
            // and leaves the hidden Excel alive holding the copied XLL.
            Assert.Contains("try { Stop-Process -Id $ownerPid -Force -ErrorAction Stop }", src);
            Assert.Contains("catch", src);
        }

        /// <summary>The stale hidden Excel from a crashed run is stopped by PROOF
        /// (invisible AND holding our own destination XLL); the user's visible Excel
        /// is only reported, never killed.</summary>
        [Fact]
        public void TheAddInFolderHelperOnlyStopsAStaleHiddenExcel()
        {
            string src = Read(Path.Combine("test", "TestSupport", "New-AddInFolder.ps1"));

            Assert.Contains("function Clear-StaleAddInHost(", src);
            Assert.Contains("Clear-StaleAddInHost $destXll", src);
            // The visible-instance refusal must come BEFORE the stop loop.
            int visibleGuard = src.IndexOf("if ($proc.MainWindowHandle -ne 0)", StringComparison.Ordinal);
            int stopStale = src.IndexOf("stopping the stale hidden Excel", StringComparison.Ordinal);
            Assert.True(visibleGuard > 0, "the visible-Excel guard is gone");
            Assert.True(stopStale > visibleGuard,
                "the visible-Excel guard must run before anything is stopped");

            // Only a process that really has OUR copy loaded is a candidate.
            Assert.Contains("$proc.Modules | Where-Object", src);
        }

        /// <summary>The cluster teardown never stops a server whose port/conf is not
        /// ours, and says so instead of guessing.</summary>
        [Fact]
        public void TheClusterTeardownRefusesStrangers()
        {
            string src = Read(Path.Combine("test", "TestSupport", "Stop-Cluster.ps1"));

            Assert.Contains("function Test-Ours(", src);
            Assert.Contains("leaving it alone", src);
        }

        /// <summary>
        /// The harness must give each node a monitoring port, or the smoke suite's
        /// wire-level refcount proof (which asks /connz on port+4000) can only ever
        /// SKIP. The unit test pins the conf line because a "helpful" cleanup there
        /// would silently downgrade a real proof to a skip again.
        /// </summary>
        [Fact]
        public void TheHarnessExposesAMonitoringPortForConnz()
        {
            string src = Read(Path.Combine("test", "TestSupport", "Start-Cluster.ps1"));

            Assert.Contains("http_port: $monitorPort", src);
            Assert.Contains("$monitorPort = $port + 4000", src);
        }

        /// <summary>
        /// The smoke suite parses /connz's <c>subscriptions_list</c> as an ARRAY of
        /// subject names (?subs=1). It was parsed as a JObject, so the probe always
        /// answered "no monitoring port" - the check could never pass even against a
        /// server that exposed /connz, and the wire-level proof of "one physical
        /// subscription for N listeners" never ran.
        /// </summary>
        [Fact]
        public void TheConnzProbeParsesTheSubscriptionListTheServerSends()
        {
            string src = Read(Path.Combine("test", "SmokeTests", "Program.cs"));

            Assert.Contains("connection[\"subscriptions_list\"] as JArray", src);
            Assert.DoesNotContain("connection[\"subscriptions_list\"] as JObject", src);
        }
    }
}
