using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NatsExcel.Tools;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Behavioural pin for the installer's connection probe: the probe must never
    /// write to, or publish into, a server it does not own. The address the user
    /// types in the wizard is very often somebody else's server, so the verdict
    /// "reached + authenticated" has to come from a connect and a PING alone.
    ///
    /// <para>
    /// This is the strong half of the pair: it drives the REAL validator code
    /// (<see cref="ConnectionTester.TestAsync"/>, the exact method
    /// <c>NatsValidate --test</c> calls, and the exact client the add-in uses)
    /// against a fake NATS server built here, and asserts on the bytes the client
    /// actually put on the wire. A source scan cannot prove that - it only proves
    /// what somebody typed. Between them, a "publish to prove the subscription
    /// works" step cannot come back unnoticed.
    /// </para>
    ///
    /// <para>
    /// The one subscription the client may legitimately create is its own
    /// <c>_INBOX.*</c> plumbing (the NATS client opens a shared inbox when it
    /// connects, for request/reply); it belongs to the probe's own namespace and
    /// mutates nothing on the server. Everything else - <c>PUB</c>, <c>HPUB</c>,
    /// a JetStream API call (<c>$JS.API.*</c> request) - is a failure.
    /// </para>
    ///
    /// <para>
    /// In its OWN collection so it never runs in parallel with the rest: it binds
    /// a loopback listener and drives a real NATS client through a full connect,
    /// and under xUnit's default parallelism that occasionally lost the race
    /// (1 failure in 13 suite runs, 0 in 10 isolated runs). Serialising it is the
    /// fix, not loosening the assertion.
    /// </para>
    /// </summary>
    [Collection("installer-probe-network")]
    public class InstallerProbeReadOnlyTests
    {
        /// <summary>
        /// Runs the real ConnectionTester against a loopback fake server and
        /// returns every line the client sent, plus the test outcome.
        /// </summary>
        [Fact]
        public async Task TheProbeConnectsAndPingsWithoutPublishingAnything()
        {
            using (var server = new FakeNatsServer())
            {
                server.Start();
                // A resolved endpoint aimed at the fake server, exactly as the
                // config loader would produce it for nats://127.0.0.1:<port>.
                var endpoint = new ResolvedEndpoint
                {
                    Urls = new List<string> { "nats://127.0.0.1:" + server.Port },
                    Auth = new AuthConfig(),
                    TimeoutMs = 3000,
                    Reconnect = new ReconnectConfig(),
                    ClusterLabel = "local",
                };

                TestResult result = await ConnectionTester.TestAsync(endpoint, 3000);

                // The probe must actually have PROVEN the connection: a test that
                // silently failed would make the assertions below vacuous.
                Assert.True(result.Ok, "the probe did not connect: " + result.Summary);
                Assert.Contains("PONG", string.Join("\n", result.Lines), StringComparison.Ordinal);
                Assert.True(server.SawConnect, "the client never sent CONNECT");
                Assert.True(server.SawPing, "the client never sent PING");

                string wire = server.Transcript();

                // The rule: no publish, in any form.
                Assert.DoesNotContain("PUB ", wire, StringComparison.Ordinal);
                Assert.DoesNotContain("HPUB ", wire, StringComparison.Ordinal);

                // No JetStream API traffic either: creating a stream to "check"
                // JetStream would be the same violation as a publish, and it is a
                // request (published to $JS.API.*) under the hood.
                Assert.DoesNotContain("$JS.API", wire, StringComparison.Ordinal);

                // And no other client-side write verb. (SUB is allowed only for the
                // client's own inbox, asserted below.)
                foreach (string verb in new[] { "UNSUB", "CONNECT {\"verbose\":true" })
                    Assert.DoesNotContain(verb, wire, StringComparison.Ordinal);

                foreach (string line in server.ClientLines())
                {
                    if (!line.StartsWith("SUB ", StringComparison.Ordinal))
                        continue;
                    // The only subject the probe may subscribe to is its own inbox.
                    Assert.Contains("_INBOX.", line, StringComparison.Ordinal);
                }
            }
        }

        /// <summary>
        /// A server that answers INFO and PONG and records everything the client
        /// sends. Deliberately minimal: the probe only needs INFO + PONG, and being
        /// minimal is what makes an unexpected verb obvious.
        /// </summary>
        private sealed class FakeNatsServer : IDisposable
        {
            private readonly TcpListener _listener;
            private readonly List<string> _lines = new List<string>();
            private readonly object _sync = new object();
            private readonly CancellationTokenSource _cts = new CancellationTokenSource();
            private Task _accept;

            public FakeNatsServer()
            {
                _listener = new TcpListener(IPAddress.Loopback, 0);
                _listener.Start();
                Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
            }

            public int Port { get; }

            public bool SawConnect { get; private set; }
            public bool SawPing { get; private set; }

            public void Start()
            {
                _accept = Task.Run(AcceptLoopAsync);
            }

            public string Transcript()
            {
                lock (_sync)
                    return string.Join("\n", _lines);
            }

            public List<string> ClientLines()
            {
                lock (_sync)
                    return _lines.ToList();
            }

            private async Task AcceptLoopAsync()
            {
                while (!_cts.IsCancellationRequested)
                {
                    TcpClient client;
                    try
                    {
                        client = await _listener.AcceptTcpClientAsync().ConfigureAwait(false);
                    }
                    catch
                    {
                        return; // listener stopped
                    }

                    _ = Task.Run(() => ServeAsync(client));
                }
            }

            private async Task ServeAsync(TcpClient client)
            {
                using (client)
                {
                    try
                    {
                        NetworkStream stream = client.GetStream();
                        // INFO: enough of the real banner for the client to accept it.
                        string info = "INFO {\"server_id\":\"FAKE\",\"server_name\":\"fake\",\"version\":\"2.10.0\","
                                    + "\"go\":\"go1.21\",\"host\":\"127.0.0.1\",\"port\":" + Port + ","
                                    + "\"headers\":true,\"max_payload\":1048576,"
                                    + "\"proto\":1,\"client_id\":1}\r\n";
                        byte[] infoBytes = Encoding.ASCII.GetBytes(info);
                        await stream.WriteAsync(infoBytes, 0, infoBytes.Length).ConfigureAwait(false);
                        await stream.FlushAsync().ConfigureAwait(false);

                        var buffer = new byte[8192];
                        var pending = new StringBuilder();
                        while (!_cts.IsCancellationRequested)
                        {
                            int n = await stream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                            if (n <= 0)
                                return;

                            pending.Append(Encoding.ASCII.GetString(buffer, 0, n));
                            // Split complete CRLF lines out of the pending buffer.
                            while (true)
                            {
                                string all = pending.ToString();
                                int crlf = all.IndexOf("\r\n", StringComparison.Ordinal);
                                if (crlf < 0)
                                    break;
                                string line = all.Substring(0, crlf);
                                pending.Remove(0, crlf + 2);

                                lock (_sync)
                                    _lines.Add(line);
                                if (line.StartsWith("CONNECT", StringComparison.Ordinal))
                                    SawConnect = true;
                                if (line == "PING")
                                {
                                    SawPing = true;
                                    byte[] pong = Encoding.ASCII.GetBytes("PONG\r\n");
                                    await stream.WriteAsync(pong, 0, pong.Length).ConfigureAwait(false);
                                    await stream.FlushAsync().ConfigureAwait(false);
                                }
                            }
                        }
                    }
                    catch
                    {
                        // A fake server must never fault the test process.
                    }
                }
            }

            public void Dispose()
            {
                _cts.Cancel();
                try { _listener.Stop(); } catch { }
                try { _accept?.Wait(2000); } catch { }
                _cts.Dispose();
            }
        }
    }
}
