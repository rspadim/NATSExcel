using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// <see cref="ConfigSchema"/> is the per-value schema walk behind the config
    /// contract (PLAN F1.3): one wrongly-typed value is dropped on its own and an
    /// issue is recorded, while the rest of the file survives. The schema is the
    /// embedded <c>NATSExcel.schema.json</c>, so these tests exercise the real
    /// schema, not a stand-in.
    /// </summary>
    public class ConfigSchemaTests
    {
        private static (JObject cleaned, List<ConfigIssue> issues) Clean(string json)
        {
            JObject root = ConfigSchema.ParseWithComments(json);
            var issues = new List<ConfigIssue>();
            var cleaned = ConfigSchema.Clean(root, issues) as JObject;
            return (cleaned, issues);
        }

        // ------------------------------------------------------------------
        // ParseWithComments.
        // ------------------------------------------------------------------

        [Fact]
        public void ParseWithComments_AcceptsLineAndBlockComments()
        {
            JObject root = ConfigSchema.ParseWithComments(
                "{\n// a line comment\n\"a\": 1, /* a block comment */ \"b\": 2\n}");
            Assert.NotNull(root);
            Assert.Equal(1, (int)root["a"]);
            Assert.Equal(2, (int)root["b"]);
        }

        [Fact]
        public void ParseWithComments_NonObjectRoot_ReturnsNull()
        {
            Assert.Null(ConfigSchema.ParseWithComments("[1, 2, 3]"));
            Assert.Null(ConfigSchema.ParseWithComments("\"just a string\""));
            Assert.Null(ConfigSchema.ParseWithComments("42"));
        }

        // ------------------------------------------------------------------
        // Clean: per-value drop, issues recorded.
        // ------------------------------------------------------------------

        [Fact]
        public void Clean_WronglyTypedValue_IsDroppedAndRecorded_WhileTheRestSurvives()
        {
            var (cleaned, issues) = Clean(
                "{ \"ConflationMs\": \"not-an-integer\", \"SyncWrite\": \"fireforget\" }");

            Assert.NotNull(cleaned);
            Assert.Null(cleaned["ConflationMs"]);                 // dropped
            Assert.Equal("fireforget", (string)cleaned["SyncWrite"]); // kept

            ConfigIssue issue = Assert.Single(issues);
            Assert.Equal("/ConflationMs", issue.Path);
            Assert.Contains("expected", issue.Message);
        }

        [Fact]
        public void Clean_AdditionalPropertiesFalse_DropsAnUnknownKey()
        {
            var (cleaned, issues) = Clean(
                "{ \"NotARealSetting\": 1, \"SyncWrite\": \"sync\" }");

            Assert.NotNull(cleaned);
            Assert.Null(cleaned["NotARealSetting"]);
            Assert.Equal("sync", (string)cleaned["SyncWrite"]);

            ConfigIssue issue = Assert.Single(issues);
            Assert.Equal("/NotARealSetting", issue.Path);
            Assert.Equal("unknown key", issue.Message);
        }

        [Fact]
        public void Clean_TypeArray_AcceptsBothAStringAndNull()
        {
            var (withString, stringIssues) = Clean(
                "{ \"Auth\": { \"user\": \"alice\" } }");
            Assert.Empty(stringIssues);
            Assert.Equal("alice", (string)withString["Auth"]["user"]);

            var (withNull, nullIssues) = Clean(
                "{ \"Auth\": { \"user\": null } }");
            Assert.Empty(nullIssues);
            Assert.Equal(JTokenType.Null, withNull["Auth"]["user"].Type);
        }

        [Fact]
        public void Clean_TypeArray_StillDropsAValueOfTheWrongType()
        {
            // ["string","null"] accepts neither an object nor a number.
            var (cleaned, issues) = Clean(
                "{ \"Auth\": { \"user\": 123 } }");

            Assert.Null(cleaned["Auth"]["user"]);
            ConfigIssue issue = Assert.Single(issues);
            Assert.Equal("/Auth/user", issue.Path);
        }

        [Fact]
        public void Clean_Pattern_IsMatchedCaseInsensitively()
        {
            // The SyncWrite pattern spells out ASCII classes; "SYNC" must pass.
            var (upper, upperIssues) = Clean("{ \"SyncWrite\": \"SYNC\" }");
            Assert.Empty(upperIssues);
            Assert.Equal("SYNC", (string)upper["SyncWrite"]);

            var (mixed, mixedIssues) = Clean("{ \"SyncWrite\": \"FireForget-All\" }");
            Assert.Empty(mixedIssues);
            Assert.Equal("FireForget-All", (string)mixed["SyncWrite"]);
        }

        [Fact]
        public void Clean_Pattern_RejectsAValueOutsideTheAllowedSet()
        {
            var (cleaned, issues) = Clean("{ \"SyncWrite\": \"sync-and-also-loud\" }");
            Assert.Null(cleaned["SyncWrite"]);
            ConfigIssue issue = Assert.Single(issues);
            Assert.Equal("/SyncWrite", issue.Path);
            Assert.Contains("allowed values", issue.Message);
        }

        [Fact]
        public void Clean_MinItemsOnClusterUrls_ReportsAnEmptyArray()
        {
            // Clusters.*.urls declares minItems: 1; an empty array passes every
            // item check but still describes a cluster with no URL.
            var (cleaned, issues) = Clean(
                "{ \"Clusters\": { \"c\": { \"urls\": [] } } }");

            ConfigIssue issue = Assert.Single(issues);
            Assert.Equal("/Clusters/c/urls", issue.Path);
            Assert.Contains("at least 1 item", issue.Message);
            Assert.Null(cleaned["Clusters"]["c"]["urls"]);
        }

        [Fact]
        public void Clean_WellFormedDocument_ProducesNoIssues()
        {
            var (cleaned, issues) = Clean(
                "{ \"SyncWrite\": \"sync\", \"Clusters\": { \"c\": { \"urls\": [\"nats://h:4222\"] } } }");
            Assert.Empty(issues);
            Assert.Equal("nats://h:4222", (string)cleaned["Clusters"]["c"]["urls"][0]);
        }

        [Fact]
        public void Clean_NullRoot_ReturnsNull()
        {
            Assert.Null(ConfigSchema.Clean(null, new List<ConfigIssue>()));
        }
    }
}
