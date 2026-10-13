using System;
using System.Collections.Generic;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// <see cref="AppConfig.ResolveEndpointCore"/> and <see cref="AppConfig.ResolveSettings"/>
    /// are the endpoint-resolution and settings-merge seam (DESIGN.md §8.1,
    /// AGENTS config contract). Both are exercised offline: the pure resolver is
    /// fed dictionaries directly, and the settings merge parses a small JSON
    /// document and installs it through the frozen <c>AppConfig.CurrentOverrideForTests</c>
    /// seam, so no machine-local <c>NATSExcel.json</c> is involved.
    /// </summary>
    [Collection("config-seam")]
    public class AppConfigResolutionTests
    {
        // ------------------------------------------------------------------
        // ResolveEndpointCore: the 4-step order.
        // ------------------------------------------------------------------

        [Fact]
        public void ResolveEndpointCore_ClusterName_ResolvesBeforeEverythingElse()
        {
            var servers = new Dictionary<string, string> { { "x", "nats://server" } };
            var clusters = new Dictionary<string, string> { { "x", "nats://cluster" } };

            Assert.Equal("nats://cluster",
                AppConfig.ResolveEndpointCore("x", "nats://default", servers, clusters));
        }

        [Fact]
        public void ResolveEndpointCore_ServerName_ResolvesWhenNoClusterMatches()
        {
            var servers = new Dictionary<string, string> { { "s", "nats://server" } };
            var clusters = new Dictionary<string, string>();

            Assert.Equal("nats://server",
                AppConfig.ResolveEndpointCore("s", "nats://default", servers, clusters));
        }

        [Fact]
        public void ResolveEndpointCore_RawUrl_IsUsedAsIs()
        {
            Assert.Equal("nats://raw:4222",
                AppConfig.ResolveEndpointCore("nats://raw:4222", "nats://default", null, null));

            // A comma-joined URL list also carries a scheme and passes through.
            Assert.Equal("nats://a:4222,nats://b:4222",
                AppConfig.ResolveEndpointCore("nats://a:4222,nats://b:4222", "nats://default", null, null));
        }

        [Fact]
        public void ResolveEndpointCore_BlankEndpoint_UsesTheDefaultEndpoint()
        {
            Assert.Equal("nats://default",
                AppConfig.ResolveEndpointCore(null, "nats://default", null, null));
            Assert.Equal("nats://default",
                AppConfig.ResolveEndpointCore("   ", "nats://default", null, null));
        }

        [Fact]
        public void ResolveEndpointCore_BlankEndpointAndBlankDefault_ReturnsNull()
        {
            Assert.Null(AppConfig.ResolveEndpointCore("   ", null, null, null));
            Assert.Null(AppConfig.ResolveEndpointCore("   ", "  ", null, null));
        }

        [Fact]
        public void ResolveEndpointCore_DanglingName_ReturnsNull_NeverTheDefault()
        {
            // The whole point of the null: a formula-named endpoint that resolves
            // to nothing must surface a #CONFIG error instead of silently
            // connecting to a hidden localhost default (DESIGN.md §8.1).
            Assert.Null(AppConfig.ResolveEndpointCore(
                "does-not-exist", "nats://localhost:4222",
                new Dictionary<string, string>(), new Dictionary<string, string>()));
        }

        [Fact]
        public void ResolveEndpointCore_ExactNameWinsOverACaseInsensitiveMatch()
        {
            var clusters = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Feed", "nats://upper" },
                { "FEED", "nats://exact" },
            };

            // Exact "FEED" is found directly, not via the fallback scan.
            Assert.Equal("nats://exact",
                AppConfig.ResolveEndpointCore("FEED", null, null, clusters));
            // An exact match for "Feed" is likewise preferred.
            Assert.Equal("nats://upper",
                AppConfig.ResolveEndpointCore("Feed", null, null, clusters));
        }

        [Fact]
        public void ResolveEndpointCore_DifferentlyCasedName_ResolvesCaseInsensitively()
        {
            var servers = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Production", "nats://prod:4222" },
            };

            Assert.Equal("nats://prod:4222",
                AppConfig.ResolveEndpointCore("production", null, servers, null));
        }

        [Fact]
        public void ResolveEndpointCore_BlankClusterUrl_FallsThroughToTheServerOfTheSameName()
        {
            var clusters = new Dictionary<string, string> { { "x", "" } };
            var servers = new Dictionary<string, string> { { "x", "nats://server" } };

            Assert.Equal("nats://server",
                AppConfig.ResolveEndpointCore("x", null, servers, clusters));
        }

        [Fact]
        public void ResolveEndpointCore_TrimsTheEndpointBeforeResolving()
        {
            var servers = new Dictionary<string, string> { { "s", "nats://server" } };
            Assert.Equal("nats://server",
                AppConfig.ResolveEndpointCore("  s  ", null, servers, null));
        }

        // ------------------------------------------------------------------
        // ResolveSettings: shared vs per-cluster.
        // ------------------------------------------------------------------

        private static AppConfig.Root Parse(string json)
        {
            return AppConfig.Sanitize(AppConfig.Parse(json));
        }

        private static void WithConfig(AppConfig.Root root, Action<AppConfig.Root> body)
        {
            var previous = AppConfig.CurrentOverrideForTests;
            try
            {
                AppConfig.CurrentOverrideForTests = root;
                body(root);
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = previous;
            }
        }

        [Fact]
        public void ResolveSettings_AuthWithNoClusterOverride_UsesTheSharedAuth()
        {
            var root = Parse(@"{
                ""Default"": ""local"",
                ""Auth"": { ""user"": ""shared-user"", ""password"": ""shared-pass"" },
                ""Clusters"": { ""local"": { ""urls"": [""nats://h:4222""] } }
            }");

            WithConfig(root, _ =>
            {
                var settings = AppConfig.ResolveSettings("local");
                Assert.Equal("shared-user", settings.Auth.User);
                Assert.Equal("shared-pass", settings.Auth.Password);
                Assert.Null(settings.Auth.Token);
            });
        }

        [Fact]
        public void ResolveSettings_ResolvedUrl_MapsBackToItsClusterAndItsAuth()
        {
            // The connection manager receives the RESOLVED URL, not the name the
            // user typed, so the cluster's auth must still be found by matching
            // that URL against the cluster's urls (DESIGN.md §8.1).
            var root = Parse(@"{
                ""Default"": ""local"",
                ""Auth"": { ""user"": ""shared-user"" },
                ""Clusters"": {
                    ""tokened"": { ""urls"": [""nats://tok:4222""], ""auth"": { ""token"": ""cluster-token"" } }
                }
            }");

            WithConfig(root, _ =>
            {
                var settings = AppConfig.ResolveSettings("nats://tok:4222");
                Assert.Equal("cluster-token", settings.Auth.Token);
            });
        }

        [Fact]
        public void ResolveSettings_ClusterAuth_OverridesTheFieldItDeclares()
        {
            var root = Parse(@"{
                ""Default"": ""local"",
                ""Auth"": { ""user"": ""shared-user"", ""token"": ""shared-token"" },
                ""Clusters"": {
                    ""tokened"": { ""urls"": [""nats://tok:4222""], ""auth"": { ""token"": ""cluster-token"" } }
                }
            }");

            WithConfig(root, _ =>
            {
                Assert.Equal("cluster-token", AppConfig.ResolveSettings("tokened").Auth.Token);
            });
        }

        [Fact]
        public void ResolveSettings_ClusterAuth_KeepsSharedFieldsItDidNotDeclare()
        {
            // DESIGN.md §8.1: "a cluster that declares only a token keeps the
            // shared user/password for the fields it did not set". The cluster
            // merge must be field-by-field over the shared Auth.
            var root = Parse(@"{
                ""Default"": ""local"",
                ""Auth"": { ""user"": ""shared-user"", ""password"": ""shared-pass"" },
                ""Clusters"": {
                    ""tokened"": { ""urls"": [""nats://tok:4222""], ""auth"": { ""token"": ""cluster-token"" } }
                }
            }");

            WithConfig(root, _ =>
            {
                var settings = AppConfig.ResolveSettings("tokened");
                Assert.Equal("cluster-token", settings.Auth.Token);      // declared: wins
                Assert.Equal("shared-user", settings.Auth.User);          // not declared: kept
                Assert.Equal("shared-pass", settings.Auth.Password);      // not declared: kept
            });
        }

        [Fact]
        public void ResolveSettings_ClusterAuth_KeepsTheSharedUserWhenOnlyThePasswordIsOverridden()
        {
            // Same documented field-by-field merge, and the case that breaks a real
            // connection: a cluster that overrides only the password must keep the
            // shared user, otherwise the merged pair is (null user, cluster pass)
            // and the server rejects it.
            var root = Parse(@"{
                ""Default"": ""local"",
                ""Auth"": { ""user"": ""shared-user"", ""password"": ""shared-pass"" },
                ""Clusters"": {
                    ""rotated"": { ""urls"": [""nats://rot:4222""], ""auth"": { ""password"": ""new-pass"" } }
                }
            }");

            WithConfig(root, _ =>
            {
                var settings = AppConfig.ResolveSettings("rotated");
                Assert.Equal("new-pass", settings.Auth.Password);        // declared: wins
                Assert.Equal("shared-user", settings.Auth.User);          // not declared: kept
            });
        }

        [Fact]
        public void ResolveSettings_ClusterTimeout_OverridesTheSharedTimeout()
        {
            var root = Parse(@"{
                ""Default"": ""local"",
                ""Auth"": {},
                ""Clusters"": {
                    ""local"":    { ""urls"": [""nats://h:4222""] },
                    ""slow"":     { ""urls"": [""nats://s:4222""], ""timeout"": 2500 }
                }
            }");

            WithConfig(root, _ =>
            {
                Assert.Equal(2500, AppConfig.ResolveSettings("slow").TimeoutMs);
                Assert.Equal(root.TimeoutMs, AppConfig.ResolveSettings("local").TimeoutMs);
            });
        }

        [Fact]
        public void ResolveSettings_ClusterReconnect_OverridesTheDeclaredField()
        {
            var root = Parse(@"{
                ""Default"": ""local"",
                ""Clusters"": {
                    ""slow"": { ""urls"": [""nats://s:4222""], ""reconnect"": { ""maxAttempts"": 5 } }
                }
            }");

            WithConfig(root, _ =>
            {
                Assert.Equal(5, AppConfig.ResolveSettings("slow").Reconnect.MaxAttempts);
            });
        }
    }
}
