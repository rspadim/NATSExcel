using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NATS.Client.Core;

namespace NatsExcel.Tools
{
    internal sealed class TestResult
    {
        public bool Ok { get; set; }
        public string Summary { get; set; }
        public List<string> Lines { get; } = new List<string>();
    }

    /// <summary>
    /// Connects with the same NATS.Net client the add-in uses, so every
    /// credential style is exercised for real (token, user/password, NKey seed,
    /// JWT + seed, creds file, mTLS). This is the half the PowerShell test
    /// cannot do: it only speaks the wire protocol.
    /// </summary>
    internal static class ConnectionTester
    {
        public static async Task<TestResult> Test(ResolvedEndpoint endpoint, int timeoutMs)
        {
            var result = new TestResult();
            if (endpoint.Urls.Count == 0)
            {
                result.Ok = false;
                result.Summary = "no URL to connect to.";
                return result;
            }

            int timeout = timeoutMs > 0 ? timeoutMs : endpoint.TimeoutMs;

            var opts = new NatsOpts
            {
                Url = string.Join(",", endpoint.Urls),
                Name = "NATSExcel-validate",
                ConnectTimeout = TimeSpan.FromMilliseconds(timeout),
                AuthOpts = BuildAuth(endpoint.Auth),
                TlsOpts = BuildTls(endpoint.Auth),
                MaxReconnectRetry = 0, // a validator must fail fast, not retry forever
            };

            using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(timeout + 2000));
            NatsConnection connection = null;
            try
            {
                connection = new NatsConnection(opts);
                await connection.ConnectAsync().ConfigureAwait(false);

                var info = connection.ServerInfo;
                if (info != null)
                {
                    result.Lines.Add("server    : " + info.Name + " (v" + info.Version + ")");
                    result.Lines.Add("cluster   : " + (string.IsNullOrEmpty(info.Cluster) ? "(none - standalone)" : info.Cluster));
                    result.Lines.Add("jetstream : " + (info.JetStreamAvailable ? "available" : "not enabled on this server"));
                    if (info.ClientConnectUrls != null && info.ClientConnectUrls.Length > 0)
                        result.Lines.Add("peers     : " + info.ClientConnectUrls.Length + " node(s) discovered");
                }

                // A real round trip proves the connection works end to end.
                // PingAsync returns the round-trip time, not a bool.
                TimeSpan rtt = await connection.PingAsync().ConfigureAwait(false);
                result.Lines.Add("ping      : PONG in " + rtt.TotalMilliseconds.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + " ms");

                // The configured cluster name (optional assertion) is a warning,
                // never a failure: the label is the user's, the announced name
                // is the server's.
                if (!string.IsNullOrEmpty(endpoint.ExpectCluster) && info != null)
                {
                    if (!string.Equals(info.Cluster, endpoint.ExpectCluster, StringComparison.OrdinalIgnoreCase))
                        result.Lines.Add("WARNING   : expectCluster='" + endpoint.ExpectCluster +
                                         "' but the server is in cluster '" + (info.Cluster ?? "(none)") + "'");
                    else
                        result.Lines.Add("cluster   : matches expectCluster '" + endpoint.ExpectCluster + "'");
                }

                result.Ok = true;
                result.Summary = "server " + info?.Name + " v" + info?.Version + " answered in " + rtt.TotalMilliseconds.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + " ms.";
            }
            catch (OperationCanceledException)
            {
                result.Ok = false;
                result.Summary = "timed out after " + timeout + " ms.";
            }
            catch (Exception ex)
            {
                result.Ok = false;
                result.Summary = Describe(ex);
            }
            finally
            {
                if (connection != null)
                {
                    try { await connection.DisposeAsync().ConfigureAwait(false); } catch { /* best effort */ }
                }
            }

            return result;
        }

        private static NatsAuthOpts BuildAuth(AuthConfig auth)
        {
            if (auth == null) return NatsAuthOpts.Default;

            // The client builds a native error for an invalid combination; we
            // mirror the precedence documented in DESIGN.md §8.
            if (!string.IsNullOrEmpty(auth.CredsFile))
                return new NatsAuthOpts { CredsFile = auth.CredsFile };
            if (!string.IsNullOrEmpty(auth.Jwt) && !string.IsNullOrEmpty(auth.NkeySeed))
                return new NatsAuthOpts { Jwt = auth.Jwt, Seed = ReadSeed(auth.NkeySeed) };
            if (!string.IsNullOrEmpty(auth.NkeySeed))
                return new NatsAuthOpts { Seed = ReadSeed(auth.NkeySeed) };
            if (!string.IsNullOrEmpty(auth.Token))
                return new NatsAuthOpts { Token = auth.Token };
            if (!string.IsNullOrEmpty(auth.User))
                return new NatsAuthOpts { Username = auth.User, Password = auth.Password };
            return NatsAuthOpts.Default;
        }

        /// <summary>
        /// The seed field accepts either the seed text itself (an "SU..." string)
        /// or a path to the .nk file that holds it. Passing a path where the
        /// client expects a seed fails with "Invalid base32 character", so a
        /// readable file is always read first.
        /// </summary>
        private static string ReadSeed(string seedOrPath)
        {
            if (string.IsNullOrWhiteSpace(seedOrPath))
                return seedOrPath;
            string value = seedOrPath.Trim();
            if (value.StartsWith("SU", StringComparison.Ordinal))
                return value;
            if (File.Exists(value))
                return File.ReadAllText(value).Trim();
            return value;
        }

        private static NatsTlsOpts BuildTls(AuthConfig auth)
        {
            if (auth == null || (!auth.Tls && string.IsNullOrEmpty(auth.TlsCert)))
                return NatsTlsOpts.Default;

            var tls = new NatsTlsOpts
            {
                Mode = TlsMode.Require,
                // A .pfx/.p12 bundle is one file + password; a .pem pair would
                // use CertFile/KeyFile, which the config does not model yet.
                CertBundleFile = string.IsNullOrEmpty(auth.TlsCert) ? null : auth.TlsCert,
                CertBundleFilePassword = string.IsNullOrEmpty(auth.TlsCertPassword) ? null : auth.TlsCertPassword,
            };
            return tls;
        }

        private static string Describe(Exception ex)
        {
            var messages = new List<string>();
            for (var e = ex; e != null; e = e.InnerException)
                messages.Add(e.Message);
            return string.Join(" | ", messages);
        }
    }
}
