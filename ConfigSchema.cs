using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace NATSExcel
{
    /// <summary>A single problem found while validating NATSExcel.json.</summary>
    internal sealed class ConfigIssue
    {
        /// <summary>JSON path of the offending value (e.g. "/RTD/NatsUpdateStyle").</summary>
        internal string Path { get; }

        /// <summary>Human-readable reason the value was dropped.</summary>
        internal string Message { get; }

        internal ConfigIssue(string path, string message)
        {
            Path = path ?? "";
            Message = message ?? "";
        }

        /// <summary>One readable line, e.g. "/Default/timeout: expected an integer".</summary>
        public override string ToString()
        {
            return (Path.Length == 0 ? "/" : Path) + ": " + Message;
        }
    }

    /// <summary>
    /// Validates NATSExcel.json against the embedded JSON Schema and produces a
    /// cleaned copy with every invalid value removed, so one wrong value can no
    /// longer discard the whole file: the previous tolerant parser failed at the
    /// JSON READER level for a bool/object/array in a numeric position (the Error
    /// callback cannot repair that), which collapsed the file to defaults.
    ///
    /// The schema lives in NATSExcel.schema.json. This class implements a small
    /// schema walk over Newtonsoft's JToken tree (no NJsonSchema dependency); it
    /// mirrors the rules of installer/validate-config.ps1: a `type` string OR
    /// array (e.g. ["string","null"]), `minimum`/`maximum`, a case-insensitive
    /// ASCII `pattern`, nested `properties`, `additionalProperties` (false or a
    /// schema object), `items` for arrays and `required`. Comments (// and
    /// /* */) are accepted on load.
    /// </summary>
    internal static class ConfigSchema
    {
        /// <summary>The configuration file this schema describes. Kept here so a
        /// caller that only has ConfigSchema can still name the file.</summary>
        internal const string ConfigFileName = "NATSExcel.json";

        /// <summary>The embedded schema resource name.</summary>
        private const string ResourceName = "NATSExcel.schema.json";

        private static readonly Lazy<string> _schemaText = new Lazy<string>(ReadSchemaText);

        private static readonly Lazy<JObject> _schema = new Lazy<JObject>(ParseSchema);

        /// <summary>The schema document, verbatim (for inspection through
        /// NatsUDFConfigSchema). Empty when neither the embedded resource nor a
        /// copy next to the assembly is available.</summary>
        internal static string SchemaText
        {
            get { return _schemaText.Value; }
        }

        /// <summary>Parses JSON accepting // and /* */ comments (Newtonsoft's
        /// reader skips comments when loading into a JToken). Returns null when the
        /// text is not a JSON object; the caller treats that as an unusable file.
        /// Throws on malformed JSON (the caller's catch handles it).</summary>
        internal static JObject ParseWithComments(string json)
        {
            if (!string.IsNullOrEmpty(json))
                json = json.TrimStart(' ', '\t', '\r', '\n', '\uFEFF');
            using (var reader = new JsonTextReader(new StringReader(json)))
            {
                reader.DateParseHandling = DateParseHandling.None;
                var token = JToken.ReadFrom(reader);
                return token as JObject;
            }
        }

        /// <summary>Returns a copy of <paramref name="root"/> with every schema
        /// violation removed, appending one <see cref="ConfigIssue"/> per removal.
        /// Deserializing the result cannot fail on a wrong-typed value because the
        /// reader never sees it.
        ///
        /// When no schema is available (the embedded resource was not added to the
        /// project yet and no copy sits next to the assembly) the document is
        /// returned UNCHANGED: validation degrades to today's behavior instead of
        /// dropping every key. Returns null when the root itself is null.
        ///
        /// The document and its schema are walked TOGETHER (not by error-path
        /// strings) so the exact JSON property names are preserved.</summary>
        internal static JToken Clean(JObject root, IList<ConfigIssue> issues)
        {
            if (root == null)
                return null;
            var schema = _schema.Value;
            if (schema == null)
                return root;

            // The schema declares its patterns case-insensitive; .NET folds case
            // using the CURRENT culture, so under tr-TR a capital I folds to the
            // dotless "ı" and a valid value would be dropped. Pin the invariant
            // culture for the walk (mirrors RedisExcel's ConfigSchema).
            var saved = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
                return Prune(root, schema, "", issues);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = saved;
            }
        }

        /// <summary>Validates a NATSExcel.json and returns a short report for a
        /// cell. <paramref name="path"/> = null/blank validates the configuration
        /// this process would load; otherwise the given file is validated.
        /// Comments (// and /* */) are accepted. Returns "OK (...)" when valid,
        /// otherwise "Error: ..." plus one line per problem. Never throws.</summary>
        internal static string Validate(string path = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path))
                    path = FirstExistingPath();
                else
                    path = path.Trim();

                if (path == null)
                    return "OK (no " + ConfigFileName + " found; using defaults)";
                if (!File.Exists(path))
                    return "Error: file not found: " + path;

                string json = File.ReadAllText(path);
                JObject root;
                try
                {
                    root = ParseWithComments(json);
                }
                catch (Exception ex)
                {
                    return "Error: " + ex.Message + " (" + path + ")";
                }
                if (root == null)
                    return "Error: the config root is not a JSON object (" + path + ")";

                var issues = new List<ConfigIssue>();
                Clean(root, issues);

                // Cross-section checks the schema cannot express. A name present
                // in BOTH `Clusters` and `Servers` resolves to the cluster (the
                // documented order), but silently - surface it as a warning so a
                // typo or a stale alias is visible (PLAN F1.5d).
                var warnings = new List<string>();
                CollectCrossSectionWarnings(root, warnings);

                if (issues.Count == 0 && warnings.Count == 0)
                    return "OK (" + path + ")";

                var lines = new List<string>();
                if (issues.Count > 0)
                    lines.Add("Error: " + issues.Count + " problem(s) in " + path);
                else
                    lines.Add("OK (" + path + ") with " + warnings.Count + " warning(s)");
                foreach (var issue in issues)
                    lines.Add(issue.ToString());
                foreach (var warning in warnings)
                    lines.Add("WARNING: " + warning);
                return string.Join("\n", lines);
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

        /// <summary>
        /// Cross-section checks the JSON Schema cannot express: a name that is both
        /// a <c>Servers</c> alias and a <c>Clusters</c> name resolves to the
        /// cluster silently (DESIGN §8.1 order), which is worth a warning.
        /// </summary>
        private static void CollectCrossSectionWarnings(JObject root, IList<string> warnings)
        {
            var servers = root["Servers"] as JObject;
            var clusters = root["Clusters"] as JObject;
            if (servers != null && clusters != null)
            {
                foreach (var cluster in clusters.Properties())
                {
                    foreach (var server in servers.Properties())
                    {
                        if (string.Equals(cluster.Name, server.Name, StringComparison.OrdinalIgnoreCase))
                            warnings.Add("'" + cluster.Name + "' is defined in both Clusters and Servers; a formula using it resolves to the cluster (rename one to avoid the ambiguity)");
                    }
                }
            }

            // A declared expectCluster is an assertion made at connect time: the
            // validator tool reports the match/mismatch against the live server,
            // and this cell reports that the assertion is configured (PLAN F1.5e).
            if (clusters != null)
            {
                foreach (var cluster in clusters.Properties())
                {
                    var expect = (cluster.Value as JObject)?["expectCluster"];
                    if (expect != null && expect.Type == JTokenType.String && !string.IsNullOrWhiteSpace((string)expect))
                        warnings.Add("cluster '" + cluster.Name + "' asserts expectCluster='" + (string)expect + "' (checked against the live server when a formula connects)");
                }
            }
        }

        // ---- schema loading ---------------------------------------------------

        private static string ReadSchemaText()
        {
            var assembly = typeof(ConfigSchema).Assembly;
            // Primary path: the resource embedded as NATSExcel.schema.json. The
            // csproj owner must add <EmbeddedResource Include="NATSExcel.schema.json" />;
            // MSBuild may prefix the name with the root namespace, so also scan.
            try
            {
                using (var stream = assembly.GetManifestResourceStream(ResourceName))
                {
                    if (stream != null)
                        using (var reader = new StreamReader(stream))
                            return reader.ReadToEnd();
                }
            }
            catch { /* fall through */ }

            try
            {
                foreach (var name in assembly.GetManifestResourceNames())
                {
                    if (!name.EndsWith(ResourceName, StringComparison.Ordinal))
                        continue;
                    using (var stream = assembly.GetManifestResourceStream(name))
                    {
                        if (stream != null)
                            using (var reader = new StreamReader(stream))
                                return reader.ReadToEnd();
                    }
                }
            }
            catch { /* fall through */ }

            // Fallback: the schema file next to the assembly.
            try
            {
                string beside = Path.Combine(AppContext.BaseDirectory, ResourceName);
                if (File.Exists(beside))
                    return File.ReadAllText(beside);
            }
            catch { /* fall through */ }

            return string.Empty;
        }

        private static JObject ParseSchema()
        {
            string text = _schemaText.Value;
            if (string.IsNullOrWhiteSpace(text))
                return null;
            try
            {
                return JObject.Parse(text);
            }
            catch (JsonException)
            {
                return null;
            }
        }

        // ---- schema walk ------------------------------------------------------

        private static JToken Prune(JToken value, JObject schema, string path, IList<ConfigIssue> issues)
        {
            var typeSpec = schema["type"];
            if (typeSpec != null && !TypeMatches(value, typeSpec))
            {
                issues?.Add(new ConfigIssue(path, "expected " + DescribeExpected(typeSpec)));
                return null;
            }

            if (value.Type == JTokenType.Object)
                return PruneObject((JObject)value, schema, path, issues);
            if (value.Type == JTokenType.Array)
                return PruneArray((JArray)value, schema, path, issues);
            return PruneScalar(value, schema, path, issues);
        }

        private static JToken PruneObject(JObject value, JObject schema, string path, IList<ConfigIssue> issues)
        {
            var properties = schema["properties"] as JObject;
            var additional = schema["additionalProperties"];
            var cleaned = new JObject();

            foreach (var prop in value.Properties())
            {
                string childPath = path + "/" + prop.Name;
                var declared = properties?[prop.Name] as JObject;
                if (declared != null)
                {
                    // The `Default` back-compat shape: an older config nested the
                    // settings under `Default` as an object. The schema models
                    // `Default` as a plain name, so the nested object would be
                    // "wrong typed" and dropped - losing a legacy file's cluster
                    // AND (because any drop drags the file through the tolerant
                    // re-parse) exposing the old bug where every nested section
                    // was discarded. Keep it verbatim and record no issue: the
                    // documented back-compat path in `AppConfig.Parse` reads it
                    // (cluster/timeout/auth) and an accepted legacy shape is not
                    // a schema violation.
                    if (path.Length == 0 && prop.Name == "Default" && prop.Value.Type == JTokenType.Object)
                        cleaned[prop.Name] = prop.Value;
                    else
                        Append(cleaned, prop.Name, Prune(prop.Value, declared, childPath, issues));
                }
                else if (additional is JObject additionalSchema)
                {
                    // Free-form keys validated against the additional schema
                    // (Servers: string; Clusters: a cluster object).
                    Append(cleaned, prop.Name, Prune(prop.Value, additionalSchema, childPath, issues));
                }
                else if (additional != null && additional.Type == JTokenType.Boolean && (bool)additional)
                {
                    cleaned[prop.Name] = prop.Value; // additionalProperties: true
                }
                else
                {
                    // additionalProperties is false (or omitted, as the schema
                    // always declares it): an unknown key is dropped and reported.
                    issues?.Add(new ConfigIssue(childPath, "unknown key"));
                }
            }

            var required = schema["required"] as JArray;
            if (required != null)
            {
                foreach (var entry in required)
                {
                    string name = (string)entry;
                    if (name != null && cleaned.Property(name) == null)
                        issues?.Add(new ConfigIssue(path + "/" + name, "required key is missing"));
                }
            }
            return cleaned;
        }

        private static JToken PruneArray(JArray value, JObject schema, string path, IList<ConfigIssue> issues)
        {
            var items = schema["items"] as JObject;
            if (items == null)
                return value; // no item schema: nothing to check

            var cleaned = new JArray();
            for (int i = 0; i < value.Count; i++)
            {
                var pruned = Prune(value[i], items, path + "[" + i + "]", issues);
                // One bad item voids the whole array: a partially valid list could
                // violate minItems and the array is only meaningful complete.
                if (pruned == null)
                    return null;
                cleaned.Add(pruned);
            }
            // minItems (e.g. Cluster.urls >= 1): an empty array passes every item
            // check but still leaves the cluster with no URL, so it is rejected
            // here (the array itself is dropped, like every other invalid value).
            var minItems = schema["minItems"];
            if (minItems != null && cleaned.Count < minItems.Value<int>())
            {
                issues?.Add(new ConfigIssue(path, "needs at least " + minItems + " item(s)"));
                return null;
            }
            return cleaned;
        }

        private static JToken PruneScalar(JToken value, JObject schema, string path, IList<ConfigIssue> issues)
        {
            if (value.Type == JTokenType.String)
            {
                var pattern = (string)schema["pattern"];
                if (!string.IsNullOrEmpty(pattern) && !MatchesPattern((string)value, pattern))
                {
                    issues?.Add(new ConfigIssue(path, "not one of the allowed values"));
                    return null;
                }
            }

            if (value.Type == JTokenType.Integer || value.Type == JTokenType.Float)
            {
                double number;
                try
                {
                    number = value.Value<double>();
                }
                catch
                {
                    issues?.Add(new ConfigIssue(path, "invalid number"));
                    return null;
                }
                var minimum = schema["minimum"];
                if (minimum != null && number < minimum.Value<double>())
                {
                    issues?.Add(new ConfigIssue(path, "value is below the minimum " + minimum));
                    return null;
                }
                var maximum = schema["maximum"];
                if (maximum != null && number > maximum.Value<double>())
                {
                    issues?.Add(new ConfigIssue(path, "value is above the maximum " + maximum));
                    return null;
                }
            }
            return value;
        }

        private static void Append(JObject target, string name, JToken pruned)
        {
            if (pruned != null)
                target[name] = pruned;
        }

        /// <summary>TRUE when the value matches the schema's `type` keyword, which
        /// may be a single name or an array of names (["string","null"]).</summary>
        private static bool TypeMatches(JToken value, JToken typeSpec)
        {
            if (typeSpec == null)
                return true;
            if (typeSpec.Type == JTokenType.Array)
            {
                foreach (var candidate in (JArray)typeSpec)
                    if (TypeMatches(value, candidate))
                        return true;
                return false;
            }

            string type = (string)typeSpec;
            if (type == null)
                return true;
            switch (type)
            {
                case "null": return value.Type == JTokenType.Null;
                case "string": return value.Type == JTokenType.String;
                case "integer": return value.Type == JTokenType.Integer;
                case "number": return value.Type == JTokenType.Integer || value.Type == JTokenType.Float;
                case "boolean": return value.Type == JTokenType.Boolean;
                case "object": return value.Type == JTokenType.Object;
                case "array": return value.Type == JTokenType.Array;
                default: return true; // unknown keyword: never drop a user value
            }
        }

        private static string DescribeExpected(JToken typeSpec)
        {
            if (typeSpec.Type == JTokenType.Array)
            {
                var names = new List<string>();
                foreach (var candidate in (JArray)typeSpec)
                    names.Add(TypeLabel((string)candidate));
                return string.Join(" or ", names);
            }
            return TypeLabel((string)typeSpec);
        }

        private static string TypeLabel(string type)
        {
            switch (type)
            {
                case "null": return "null";
                case "string": return "a string";
                case "integer": return "an integer";
                case "number": return "a number";
                case "boolean": return "true or false";
                case "object": return "an object";
                case "array": return "an array";
                default: return type ?? "a value";
            }
        }

        /// <summary>Matches a schema `pattern` case-insensitively (ASCII). The
        /// schema spells out ASCII classes, but a leading (?i) is honored too and
        /// the match is culture-invariant, mirroring validate-config.ps1.</summary>
        private static bool MatchesPattern(string value, string pattern)
        {
            string body = pattern;
            if (body.StartsWith("(?i)", StringComparison.Ordinal))
                body = body.Substring(4);
            try
            {
                return Regex.IsMatch(value, "^(?i:" + body + ")$", RegexOptions.CultureInvariant);
            }
            catch (ArgumentException)
            {
                return true; // a broken pattern must not drop a user value
            }
        }

        /// <summary>The first existing config candidate (the file the loader read,
        /// then the add-in folder, the user profile, the Excel folder and
        /// C:\Windows), or null when none exists. The add-in folder comes from the
        /// XLL path (<c>ExcelDnaUtil.XllPath</c>), not
        /// <c>AppContext.BaseDirectory</c>: inside a packed byte-loaded XLL the
        /// latter is Excel's directory, which the loader never searches, so the
        /// diagnostic would name a file the runtime did not consider.</summary>
        private static string FirstExistingPath()
        {
            foreach (var path in AppConfig.CandidatePaths())
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                    return path;
            }
            return null;
        }
    }
}
