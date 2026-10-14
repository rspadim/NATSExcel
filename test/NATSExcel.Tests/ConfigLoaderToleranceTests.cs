using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Regression tests for the config loader's tolerance contract (PLAN F1.3):
    /// one wrongly-typed value is dropped ON ITS OWN and the rest of the file is
    /// kept. These drive the real load path (<see cref="AppConfig.ReadConfig"/>,
    /// the same code <c>Load</c>/<c>Reload</c> call) rather than the schema walk
    /// in isolation, because the bug they pin lived in the ORDER of the load
    /// steps: the strict parse ran first, so a single bad value threw and
    /// discarded the whole file (URLs, per-cluster auth, delivery tuning).
    /// </summary>
    [Collection("config-seam")]
    public class ConfigLoaderToleranceTests
    {
        private static AppConfig.Root Load(string json)
        {
            return AppConfig.ReadConfig(json, "test-NATSExcel.json");
        }

        private static void WithConfig(AppConfig.Root root, System.Action<AppConfig.Root> body)
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
        public void OneWrongTypedValue_DoesNotDiscardTheRestOfTheFile()
        {
            // A single unrelated type error ("ConflationMs" as a string) must not
            // repoint the cluster URL, drop the per-cluster token, or reset the
            // delivery tuning. Before the fix every one of these was lost.
            var root = Load(@"{
                ""Default"": ""local"",
                ""Servers"": { ""dev"": ""nats://localhost:4222"" },
                ""Clusters"": {
                    ""local"": { ""urls"": [""dev""] },
                    ""secure"": { ""urls"": [""nats://secure.example:4222""], ""auth"": { ""token"": ""SECRET"" } }
                },
                ""RTD"": { ""NatsUpdateRateMs"": 5555, ""NatsUpdateStyle"": ""Timer"" },
                ""UDF"": { ""cluster"": ""secure"" },
                ""JetStream"": { ""enabled"": false },
                ""ConflationMs"": ""200""
            }");

            // The cluster URL still resolves through the Servers alias.
            WithConfig(root, _ =>
            {
                Assert.Equal("nats://localhost:4222", AppConfig.ResolveEndpoint("local", "none"));
                // The per-cluster credential survived.
                Assert.Equal("SECRET", AppConfig.ResolveSettings("secure").Auth.Token);
            });
            // The delivery tuning survived.
            Assert.Equal(5555, root.NatsUpdateRateMs);
            Assert.Equal("Timer", root.NatsUpdateStyle);
            // The JetStream toggle survived.
            Assert.False(root.JetStreamEnabled);
            // The bad value itself was dropped (ConflationMs defaults to null -
            // unset - so the string became "not declared", not a broken number).
            Assert.Null(root.ConflationMs);
        }

        [Fact]
        public void NestedLegacyDefault_IsStillHonoured()
        {
            // The documented back-compat: an older config nested the settings
            // under Default as an object. The schema models Default as a plain
            // name, so without the legacy handling the whole tolerant path was
            // taken and the nested Default was lost.
            var root = Load(@"{ ""Default"": { ""cluster"": ""prod"", ""timeout"": 2000 } }");

            Assert.Equal("prod", root.Default);
            Assert.Equal(2000, root.TimeoutMs);
        }

        [Fact]
        public void AValidFile_KeepsEveryDeclaredValue()
        {
            var root = Load(@"{
                ""Default"": ""local"",
                ""Auth"": { ""user"": ""u"", ""password"": ""p"" },
                ""Clusters"": { ""local"": { ""urls"": [""nats://h:4222""] } },
                ""Reconnect"": { ""maxAttempts"": 9 },
                ""AllowAdminWrites"": true,
                ""ConflationMs"": 250
            }");

            Assert.Equal("local", root.Default);
            Assert.Equal("u", root.Auth.User);
            Assert.Equal(9, root.Reconnect.MaxAttempts);
            Assert.True(root.AllowAdminWrites);
            Assert.Equal(250, root.ConflationMs);
        }

        [Fact]
        public void AStructuralFailure_FallsBackToDefaults()
        {
            // A non-JSON document is the documented fallback: the parse throws and
            // the caller (Load) turns it into the built-in defaults.
            Assert.ThrowsAny<System.Exception>(
                () => AppConfig.ReadConfig("{ this is not json", "test.json"));
        }
    }
}
