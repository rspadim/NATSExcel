using System;
using System.Text;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Regression tests for round-1 audit fixes. Each pins a defect that had no
    /// coverage before: if one of these fails, the fix was reverted or a
    /// neighbouring change reintroduced the bug.
    /// </summary>
    [Collection("config-seam")]
    public class AuditRound1RegressionTests
    {
        // ------------------------------------------------------------------
        // MaskNatsEndpoint: the secret-path/query forms (D3).
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("credsFile=C:\\Users\\alice\\user.creds", "credsFile=****")]
        [InlineData("tlsCert=C:\\certs\\client.pfx", "tlsCert=****")]
        [InlineData("tlsCertPassword=hunter2", "tlsCertPassword=****")]
        [InlineData("?token=abc123", "?token=****")]
        [InlineData("credsFile: C:\\Program Files\\x.pfx", "credsFile=****")]
        public void MaskNatsEndpoint_MasksSecretFileAndQueryForms(string input, string expectedFragment)
        {
            string masked = AppConfig.MaskNatsEndpoint(input);
            Assert.Contains(expectedFragment, masked);
            // The secret value itself must be gone.
            Assert.DoesNotContain("alice", masked);
            Assert.DoesNotContain("hunter2", masked);
            Assert.DoesNotContain("abc123", masked);
        }

        [Fact]
        public void MaskNatsEndpoint_LeavesAPlainEndpointAlone()
        {
            Assert.Equal("nats://localhost:4222", AppConfig.MaskNatsEndpoint("nats://localhost:4222"));
        }

        // ------------------------------------------------------------------
        // Version: the bare InformationalVersion, never the SDK's +<sha> (D4).
        // ------------------------------------------------------------------

        [Fact]
        public void Version_ReadsTheInformationalVersion_NotTheAssemblyVersion()
        {
            // The fix: AppConfig.Version reads AssemblyInformationalVersion (the
            // tag the CI stamps), not Assembly.GetName().Version (always 1.0.0.0).
            // In this test assembly that means the linked test project's own
            // informational version, whatever it is - the point is that it is the
            // informational attribute, not the assembly version.
            var attr = (System.Reflection.AssemblyInformationalVersionAttribute)
                Attribute.GetCustomAttribute(
                    typeof(AppConfig).Assembly,
                    typeof(System.Reflection.AssemblyInformationalVersionAttribute));
            Assert.NotNull(attr);
            Assert.Equal(attr.InformationalVersion, AppConfig.Version);
        }

        // ------------------------------------------------------------------
        // BoundCellText / BoundRtdCellText: never split a surrogate pair (B5),
        // and the RTD bound is 255 (C4).
        // ------------------------------------------------------------------

        [Fact]
        public void BoundCellText_NeverSplitsASurrogatePair()
        {
            // 'a' * 32766 then an astral pair then more: the cut lands inside the pair.
            var text = new string('a', 32766) + char.ConvertFromUtf32(0x1F600) + new string('b', 50);
            string bounded = NatsResultFormatter.BoundCellText(text);
            Assert.True(bounded.Length <= NatsResultFormatter.MaxCellTextLength);
            // No lone surrogate anywhere in the result.
            for (int i = 0; i < bounded.Length; i++)
                Assert.False(char.IsHighSurrogate(bounded[i]) &&
                             (i + 1 >= bounded.Length || !char.IsLowSurrogate(bounded[i + 1])),
                             "lone high surrogate at index " + i);
        }

        [Fact]
        public void BoundRtdCellText_BoundsAt255()
        {
            string bounded = NatsResultFormatter.BoundRtdCellText(new string('x', 400));
            Assert.True(bounded.Length <= NatsResultFormatter.MaxRtdCellTextLength,
                "RTD bound must be 255 (Excel-DNA FixValue truncates there): got " + bounded.Length);
            Assert.Contains("truncated", bounded);
        }

        [Fact]
        public void BoundRtdCellText_ShortTextIsUnchanged()
        {
            Assert.Equal("hello", NatsResultFormatter.BoundRtdCellText("hello"));
        }

        // ------------------------------------------------------------------
        // The legacy nested Default: a wrongly-typed nested value is dropped on
        // its own, not fatal (D1).
        // ------------------------------------------------------------------

        [Fact]
        public void LegacyNestedDefault_WrongTypedTimeout_DoesNotDiscardTheFile()
        {
            var root = AppConfig.ReadConfig(
                @"{""Default"": {""cluster"": ""prod"", ""timeout"": ""abc""}, ""AllowAdminWrites"": true}",
                "test.json");
            // The file survived: the flag is kept, the bad timeout fell back.
            Assert.True(root.AllowAdminWrites);
            Assert.Equal("prod", root.Default);
        }

        // ------------------------------------------------------------------
        // mTLS: a tlsCert without tls implies TLS (security finding 8).
        // ------------------------------------------------------------------

        [Fact]
        public void TlsCertWithoutTlsFlag_ImpliesTls()
        {
            var root = AppConfig.Sanitize(AppConfig.Parse(
                @"{""Auth"": {""tlsCert"": ""C:\\certs\\client.pfx""}}"));
            Assert.True(root.Auth.Tls);
        }

        [Fact]
        public void ExplicitTlsFalse_IsNotOverriddenByAPresentCert()
        {
            var root = AppConfig.Sanitize(AppConfig.Parse(
                @"{""Auth"": {""tls"": false, ""tlsCert"": ""C:\\certs\\client.pfx""}}"));
            Assert.False(root.Auth.Tls);
        }

        [Fact]
        public void ClusterTlsCertOnly_KeepsTheImpliedTlsAndTheCertificate()
        {
            // A cluster that declares ONLY tlsCert implies TLS. The merge used to
            // drop that implied value (the presence flag tracked the literal `tls`
            // key alone), so the cluster's mTLS silently became an opportunistic
            // connection that presented no client certificate.
            var root = AppConfig.Sanitize(AppConfig.Parse(@"{
                ""Default"": ""local"",
                ""Clusters"": {
                    ""local"": {
                        ""urls"": [""nats://localhost:4222""],
                        ""auth"": { ""tlsCert"": ""C:\\certs\\client.pfx"" }
                    }
                }
            }"));

            var previous = AppConfig.CurrentOverrideForTests;
            try
            {
                AppConfig.CurrentOverrideForTests = root;
                var auth = AppConfig.ResolveSettings("local").Auth;
                Assert.True(auth.Tls);                                   // implied, kept
                Assert.Equal("C:\\certs\\client.pfx", auth.TlsCert);     // and the cert
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = previous;
            }
        }

        [Fact]
        public void ClusterExplicitTlsFalse_BeatsAClusterCert()
        {
            // An explicit `tls:false` in the cluster still wins over its own cert:
            // the flag is declared, so the merge honours it.
            var root = AppConfig.Sanitize(AppConfig.Parse(@"{
                ""Default"": ""local"",
                ""Auth"": { ""tls"": true },
                ""Clusters"": {
                    ""plain"": {
                        ""urls"": [""nats://localhost:4222""],
                        ""auth"": { ""tls"": false, ""tlsCert"": ""C:\\c.pfx"" }
                    }
                }
            }"));

            var previous = AppConfig.CurrentOverrideForTests;
            try
            {
                AppConfig.CurrentOverrideForTests = root;
                Assert.False(AppConfig.ResolveSettings("plain").Auth.Tls);
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = previous;
            }
        }

        // ------------------------------------------------------------------
        // JetStream domain routing (PLAN §9.1#8): the knob was parsed and surfaced
        // but never used, so a multi-domain deployment talked to the default
        // $JS.API subjects instead of the cluster's.
        // ------------------------------------------------------------------

        [Fact]
        public void ResolveJetStreamDomain_ReturnsTheDeclaredDomain()
        {
            var root = AppConfig.Sanitize(AppConfig.Parse(@"{
                ""Default"": ""west"",
                ""Clusters"": {
                    ""west"": { ""urls"": [""nats://west:4222""], ""jetstream"": { ""domain"": ""WEST"" } },
                    ""plain"": { ""urls"": [""nats://plain:4222""] }
                }
            }"));
            var previous = AppConfig.CurrentOverrideForTests;
            try
            {
                AppConfig.CurrentOverrideForTests = root;
                Assert.Equal("WEST", AppConfig.ResolveJetStreamDomain("west"));
                // Also by the resolved URL, not only by the label.
                Assert.Equal("WEST", AppConfig.ResolveJetStreamDomain("nats://west:4222"));
                // A cluster that declares none must stay on the default subjects.
                Assert.Null(AppConfig.ResolveJetStreamDomain("plain"));
                // A raw URL that matches no cluster has no domain.
                Assert.Null(AppConfig.ResolveJetStreamDomain("nats://nowhere:4222"));
            }
            finally
            {
                AppConfig.CurrentOverrideForTests = previous;
            }
        }

        // ------------------------------------------------------------------
        // Masking: the CLIENT's own credential-file phrasing (round-2 lens C).
        // ------------------------------------------------------------------

        [Theory]
        [InlineData("Failed to load credentials: Error loading creds file 'C:\\certs\\user.creds': Could not find file", "user.creds")]
        [InlineData("The process cannot access the file 'C:\\certs\\user.nk' because it is being used by another process.", "user.nk")]
        [InlineData("The file 'C:\\certs\\client.pfx' is not a valid PFX", "client.pfx")]
        public void MaskNatsEndpoint_MasksAQuotedCredentialFilePath(string input, string fileName)
        {
            string masked = AppConfig.MaskNatsEndpoint(input);
            // The path is gone; the file NAME may remain (it is not the secret).
            Assert.DoesNotContain("C:\\certs", masked);
            Assert.Contains("****", masked);
            // The extension still names the KIND of file, which is what makes the
            // message useful.
            Assert.Contains(System.IO.Path.GetExtension(fileName), masked);
        }

        [Fact]
        public void MaskNatsEndpoint_LeavesAnOrdinaryQuotedFilePathAlone()
        {
            // Only credential-store extensions are redacted: a normal path in an
            // unrelated message is not a secret and must stay readable.
            string text = "Cannot read 'C:\\logs\\report.csv'";
            Assert.Equal(text, AppConfig.MaskNatsEndpoint(text));
        }

        // ------------------------------------------------------------------
        // The control-line guard counts UTF-8 BYTES, not UTF-16 chars (B4).
        // ------------------------------------------------------------------

        [Fact]
        public void RequireSubject_RejectsAnOverLongUtf8Subject()
        {
            // 1024 astral chars = 4096 UTF-8 bytes, over the 2048-byte budget,
            // though only 2048 UTF-16 code units.
            var sb = new StringBuilder();
            for (int i = 0; i < 1024; i++) sb.Append(char.ConvertFromUtf32(0x1F600));
            string text = sb.ToString();
            Assert.NotNull(NatsUdfCommon.RequireSubject(text, allowWildcard: true));
        }

        [Fact]
        public void RequireSubject_AcceptsARealisticSubject()
        {
            Assert.Null(NatsUdfCommon.RequireSubject("orders.created.eu", allowWildcard: true));
        }

        // ------------------------------------------------------------------
        // ConfigSchema: the cross-section warnings (F1.5d/F1.5e).
        // ------------------------------------------------------------------

        [Fact]
        public void Validate_ReportsAnAmbiguousClusterAndServerName()
        {
            string path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "natsexcel-ambig-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                System.IO.File.WriteAllText(path,
                    @"{""Default"": ""x"", ""Servers"": {""x"": ""nats://h:4222""}, ""Clusters"": {""x"": {""urls"": [""nats://h:4222""]}}}");
                string report = ConfigSchema.Validate(path);
                Assert.Contains("WARNING", report);
                Assert.Contains("both Clusters and Servers", report);
            }
            finally
            {
                System.IO.File.Delete(path);
            }
        }
    }
}
