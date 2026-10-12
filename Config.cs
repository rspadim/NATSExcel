using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;

namespace NATSExcel
{
    /// <summary>Write mode names accepted by <see cref="AppConfig.Root.SyncWrite"/>
    /// (DESIGN.md §8.1: sync = wait for the JetStream ack/flush; fireforget =
    /// result-agnostic publish without awaiting the ack).</summary>
    internal static class SyncWriteModes
    {
        public const string Sync = "sync";
        public const string FireForget = "fireforget";
        public const string FireForgetAll = "fireforget-all";
    }

    /// <summary>
    /// Loads NATSExcel.json once per process. Comments (// and /* */) are
    /// accepted; a missing file falls back to safe defaults. The full config
    /// model (clusters, auth styles, reconnect) is described in DESIGN.md §8.
    /// </summary>
    public static class AppConfig
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public const string ConfigFileName = "NATSExcel.json";
        public const string DefaultUrl = "nats://localhost:4222";

        private static readonly Lazy<Root> _current = new Lazy<Root>(Load);

        /// <summary>Test-only override (PLAN §1.4): when set, <see cref="Current"/>
        /// returns this instance instead of the machine's file, so the unit suite
        /// is hermetic. Assigned only by the linked test sources.</summary>
#pragma warning disable 0649 // assigned only by the linked unit test sources
        internal static Root CurrentOverrideForTests;
#pragma warning restore 0649

        public static Root Current => CurrentOverrideForTests ?? _current.Value;

        internal static string ConfigPath { get; private set; }

        public static string Version =>
            typeof(AppConfig).Assembly.GetName().Version?.ToString() ?? "dev";

        private static Root Load()
        {
            foreach (var path in CandidatePaths())
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    continue;
                try
                {
                    ConfigPath = path;
                    logger.Info("AppConfig: loaded configuration from " + path);
                    return Parse(File.ReadAllText(path));
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "AppConfig: error reading " + path + ", using defaults");
                    break; // first existing candidate stops the search
                }
            }
            logger.Info("AppConfig: no configuration file loaded, using defaults");
            return new Root();
        }

        private static IEnumerable<string> CandidatePaths()
        {
            string addInDir = null;
            try
            {
                var xll = ExcelDna.Integration.ExcelDnaUtil.XllPath;
                if (!string.IsNullOrWhiteSpace(xll)) addInDir = Path.GetDirectoryName(xll);
            }
            catch { /* no Excel context (tests) */ }

            if (addInDir != null) yield return Path.Combine(addInDir, ConfigFileName);
            yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ConfigFileName);
            string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrEmpty(windows)) yield return Path.Combine(windows, ConfigFileName);
        }

        /// <summary>Parse a config document (accepts // and /* */ comments).</summary>
        internal static Root Parse(string text)
        {
            var root = JObject.Parse(StripComments(text));
            var cfg = new Root();
            var def = root["Default"] as JObject;
            if (def != null)
            {
                cfg.Cluster = (string)def["cluster"];
                cfg.TimeoutMs = (int?)def["timeout"] ?? cfg.TimeoutMs;
                var auth = def["auth"] as JObject;
                if (auth != null)
                {
                    cfg.Auth.User = (string)auth["user"];
                    cfg.Auth.Password = (string)auth["password"];
                    cfg.Auth.Token = (string)auth["token"];
                    cfg.Auth.Jwt = (string)auth["jwt"];
                    cfg.Auth.NkeySeed = (string)auth["nkeySeed"];
                    cfg.Auth.CredsFile = (string)auth["credsFile"];
                    cfg.Auth.Tls = (bool?)auth["tls"] ?? false;
                }
            }
            var js = root["JetStream"] as JObject;
            if (js != null)
            {
                cfg.JetStreamEnabled = (bool?)js["enabled"] ?? cfg.JetStreamEnabled;
                cfg.JetStreamTimeoutMs = (int?)js["defaultTimeoutMs"] ?? cfg.JetStreamTimeoutMs;
            }
            // RTD/UDF keep only their own tuning plus an optional cluster override;
            // a null override falls back to Default.cluster at resolution time.
            var rtd = root["RTD"] as JObject;
            if (rtd != null)
                cfg.RtdCluster = (string)rtd["cluster"];
            var udf = root["UDF"] as JObject;
            if (udf != null)
                cfg.UdfCluster = (string)udf["cluster"];
            // Process-wide behavior toggles (shipped at the root of NATSExcel.json).
            // Each keeps its safe default when the key is absent.
            cfg.AllowAdminWrites = (bool?)root["AllowAdminWrites"] ?? cfg.AllowAdminWrites;
            cfg.SkipRepeatedMessages = (bool?)root["SkipRepeatedMessages"] ?? cfg.SkipRepeatedMessages;
            cfg.CoalesceRealtimeUpdates = (bool?)root["CoalesceRealtimeUpdates"] ?? cfg.CoalesceRealtimeUpdates;
            cfg.ConflationMs = (int?)root["ConflationMs"] ?? cfg.ConflationMs;
            cfg.PublishDedupCacheSize = (int?)root["PublishDedupCacheSize"] ?? cfg.PublishDedupCacheSize;
            cfg.SyncWrite = NormalizeSyncWrite((string)root["SyncWrite"] ?? cfg.SyncWrite);
            cfg.AsyncWrites = (bool?)root["AsyncWrites"] ?? cfg.AsyncWrites;
            cfg.UpdateCheck = (bool?)root["UpdateCheck"] ?? cfg.UpdateCheck;
            var servers = root["Servers"] as JObject;
            if (servers != null)
                foreach (var kv in servers)
                    if (kv.Value.Type == JTokenType.String)
                        cfg.Servers[kv.Key] = (string)kv.Value;

            var clusters = root["Clusters"] as JObject;
            if (clusters != null)
                foreach (var kv in clusters)
                {
                    var cj = kv.Value as JObject;
                    if (cj == null) continue;
                    var c = new Cluster { Label = kv.Key, ExpectCluster = (string)cj["expectCluster"] };
                    var urls = cj["urls"] as JArray;
                    if (urls != null)
                        foreach (var u in urls)
                            if (u.Type == JTokenType.String)
                                c.Urls.Add(ResolveUrl(cfg, (string)u));
                    cfg.Clusters[kv.Key] = c;
                }
            return cfg;
        }

        private static string ResolveUrl(Root cfg, string urlOrAlias)
        {
            if (urlOrAlias.Contains("://")) return urlOrAlias;
            if (cfg.Servers.TryGetValue(urlOrAlias, out var mapped)) return mapped;
            // Exact match keeps priority; a differently-cased alias still resolves
            // (DESIGN.md §8.1: exact-first, then case-insensitive).
            foreach (var kv in cfg.Servers)
                if (string.Equals(kv.Key, urlOrAlias, StringComparison.OrdinalIgnoreCase))
                    return kv.Value;
            return urlOrAlias;
        }

        /// <summary>Strip // and /* */ comments, string-aware (so nats:// survives).</summary>
        internal static string StripComments(string text)
        {
            var sb = new System.Text.StringBuilder(text.Length);
            bool inString = false, inLine = false, inBlock = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                char c2 = i + 1 < text.Length ? text[i + 1] : '\0';
                if (inLine) { if (c == '\n') { inLine = false; sb.Append(c); } continue; }
                if (inBlock) { if (c == '*' && c2 == '/') { inBlock = false; i++; } continue; }
                if (inString)
                {
                    sb.Append(c);
                    if (c == '\\' && i + 1 < text.Length) { sb.Append(c2); i++; continue; }
                    if (c == '"') inString = false;
                    continue;
                }
                if (c == '"') { inString = true; sb.Append(c); continue; }
                if (c == '/' && c2 == '/') { inLine = true; i++; continue; }
                if (c == '/' && c2 == '*') { inBlock = true; i++; continue; }
                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Redacts every credential <see cref="MaskNatsEndpoint"/> knows about in
        /// arbitrary text (a URL, a connection string, an exception message)
        /// before it reaches a cell or a log line. A string with no credential is
        /// returned unchanged. Patterns, in order: URL userinfo, key=value pairs,
        /// then the bare JWT and NKey-seed tokens (which can appear on their own).
        /// </summary>
        public static string MaskNatsEndpoint(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            string masked = UriUserInfo.Replace(text, "****@");
            masked = CredentialPair.Replace(masked, "$1$2=****");
            masked = JwtToken.Replace(masked, "****");
            masked = NkeySeedToken.Replace(masked, "****");
            return masked;
        }

        // nats://user:SECRET@host / nats://:SECRET@host / nats://user@host:
        // whatever sits between '//' and '@' is URL userinfo and must be redacted.
        private static readonly Regex UriUserInfo = new Regex(
            @"(?<=//)[^/@\s]+@", RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // A credential keyword (password/pass/user/username/token) followed by '='.
        // Preceded by the string start, whitespace, '&' (query string), ',' or ';'
        // (connection-string delimiter), so a user key like "key=user=abc" is NOT
        // rewritten; optional whitespace is allowed around '=', and a quoted value
        // ('...' or "...") is handled. The value runs to the next delimiter so
        // "password=ab,cd" is fully masked. Group 1 is the leading delimiter and
        // group 2 the keyword, so "$1$2=****" keeps them; example: "token=abc" ->
        // "token=****".
        private static readonly Regex CredentialPair = new Regex(
            @"(^|[\s&,;])(password|pass|user|username|token)\s*=\s*(?:""[^""]*""|'[^']*'|[^\s&,;]*)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // A JWT: "eyJ" (base64url of '{"') followed by one or more base64url
        // chars and dots (a JWT has header.payload.signature). A long segment is
        // assumed to be a token; short values that merely start with "eyJ" are
        // left alone. Replacement is "****".
        private static readonly Regex JwtToken = new Regex(
            @"eyJ[A-Za-z0-9_\-\.]{16,}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // An NKey seed: "SU" followed by base32 (A-Z, 2-7) for the signature
        // ("S" + a type char + the payload). A short "SU..." value is not a seed
        // and is left alone. Replacement is "****".
        private static readonly Regex NkeySeedToken = new Regex(
            @"SU[A-Z2-7]{20,}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>
        /// Resolves a UDF/RTD endpoint reference (DESIGN.md §8.1): a Clusters
        /// name → its URL list (comma-joined), then a Servers alias → its URL,
        /// then a raw URL/list, then the default. Matching is exact-name-first,
        /// then case-insensitive. A non-empty reference that resolves to nothing
        /// returns null, so the caller reports a #CONFIG error instead of silently
        /// connecting to localhost.
        /// </summary>
        public static string ResolveEndpoint(string endpoint, string defaultEndpoint)
        {
            var cfg = Current;
            return ResolveEndpointCore(endpoint, defaultEndpoint, cfg.Servers, ClusterUrlMap(cfg));
        }

        /// <summary>Resolves a cluster reference for RTD topics; a blank one falls
        /// back to RTD.cluster, then Default.cluster, then the built-in default.</summary>
        public static string ResolveRtdEndpoint(string endpoint)
        {
            var cfg = Current;
            return ResolveEndpointCore(endpoint, ResolveDefaultEndpoint(cfg, cfg.RtdCluster), cfg.Servers, ClusterUrlMap(cfg));
        }

        /// <summary>Resolves a cluster reference for UDF calls; a blank one falls
        /// back to UDF.cluster, then Default.cluster, then the built-in default.</summary>
        public static string ResolveUdfEndpoint(string endpoint)
        {
            var cfg = Current;
            return ResolveEndpointCore(endpoint, ResolveDefaultEndpoint(cfg, cfg.UdfCluster), cfg.Servers, ClusterUrlMap(cfg));
        }

        /// <summary>Projects a cluster name onto its comma-joined URL list, the
        /// flat form the pure resolver consumes. A cluster with no URLs is left
        /// out so a same-named <c>Servers</c> alias can still resolve.</summary>
        private static Dictionary<string, string> ClusterUrlMap(Root cfg)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in cfg.Clusters)
                if (kv.Value != null && kv.Value.Urls.Count > 0)
                    map[kv.Key] = string.Join(",", kv.Value.Urls);
            return map;
        }

        /// <summary>
        /// Pure endpoint resolution (kept internal for unit tests). See
        /// <see cref="ResolveEndpoint"/> for the 4-step order. <paramref name="clusters"/>
        /// maps a cluster name to its comma-joined URL list.
        /// </summary>
        internal static string ResolveEndpointCore(
            string endpoint, string defaultEndpoint,
            IDictionary<string, string> servers, IDictionary<string, string> clusters)
        {
            string name = string.IsNullOrWhiteSpace(endpoint) ? defaultEndpoint : endpoint.Trim();
            if (string.IsNullOrWhiteSpace(name))
                return null;
            if (clusters != null && TryGet(clusters, name, out var cluster) && !string.IsNullOrWhiteSpace(cluster))
                return cluster;
            if (servers != null && TryGet(servers, name, out var server) && !string.IsNullOrWhiteSpace(server))
                return server;
            // Step 3: a raw URL or URL list (always carries a scheme) is used
            // as-is. Anything else is a dangling named reference: return null so
            // the caller reports a #CONFIG error instead of connecting to a hidden
            // localhost default (DESIGN.md §8.1).
            if (name.Contains("://"))
                return name;
            return null;
        }

        /// <summary>
        /// The default endpoint reference for a blank formula argument: the
        /// section's cluster override, then Default.cluster, then the built-in
        /// URL. A configured default cluster is resolved through the same steps;
        /// when it does not resolve, the built-in default applies (a misconfigured
        /// <c>Default.cluster</c> must not break every unauthored call, unlike a
        /// dangling name the user typed in a formula).
        /// </summary>
        private static string ResolveDefaultEndpoint(Root cfg, string sectionCluster)
        {
            string reference = !string.IsNullOrWhiteSpace(sectionCluster) ? sectionCluster
                : !string.IsNullOrWhiteSpace(cfg.Cluster) ? cfg.Cluster
                : null;
            if (reference == null)
                return DefaultUrl;
            return ResolveEndpointCore(reference, DefaultUrl, cfg.Servers, ClusterUrlMap(cfg));
        }

        /// <summary>Exact-name lookup first, then a case-insensitive scan (the
        /// dictionary may itself be ordinal, in which case the scan is what
        /// implements DESIGN.md §8.1's fallback).</summary>
        private static bool TryGet<T>(IDictionary<string, T> map, string key, out T value)
        {
            if (map.TryGetValue(key, out value)) return true;
            foreach (var kv in map)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase))
                {
                    value = kv.Value;
                    return true;
                }
            }
            return false;
        }

        /// <summary>Normalizes a SyncWrite mode to lowercase; unknown or blank
        /// values fall back to "fireforget". Kept internal for tests.</summary>
        internal static string NormalizeSyncWrite(string mode)
        {
            switch (mode?.Trim().ToLowerInvariant())
            {
                case SyncWriteModes.Sync: return SyncWriteModes.Sync;
                case SyncWriteModes.FireForget: return SyncWriteModes.FireForget;
                case SyncWriteModes.FireForgetAll: return SyncWriteModes.FireForgetAll;
                default: return SyncWriteModes.FireForget;
            }
        }

        /// <summary>
        /// Resolves the nkeySeed field (AGENTS golden rule 11): a value starting
        /// with "SU" is the seed itself; otherwise a readable file path is read
        /// and trimmed (the .nk file case); anything else passes through. The
        /// NATS.Net client wants the seed, not the path, so a path passed verbatim
        /// fails with "Invalid base32 character".
        /// </summary>
        internal static string ResolveSeed(string seedOrPath)
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

        /// <summary>Configured write mode ("sync", "fireforget" or "fireforget-all").</summary>
        public static string SyncWrite => Current.SyncWrite;

        /// <summary>Whether writes are dispatched off the Excel thread (Observe).</summary>
        public static bool AsyncWrites => Current.AsyncWrites;

        /// <summary>Whether admin/destructive functions are enabled. Default false.</summary>
        public static bool AllowAdminWrites => Current.AllowAdminWrites;

        /// <summary>Whether identical consecutive payloads are skipped. Default true.</summary>
        public static bool SkipRepeatedMessages => Current.SkipRepeatedMessages;

        /// <summary>Whether real-time pushes are coalesced. Default true.</summary>
        public static bool CoalesceRealtimeUpdates => Current.CoalesceRealtimeUpdates;

        /// <summary>Explicit real-time conflation window in ms, or null when the
        /// boolean decides.</summary>
        public static int? ConflationMs => Current.ConflationMs;

        /// <summary>Capacity of the PublishIfChanged dedup cache. Default 10000.</summary>
        public static int PublishDedupCacheSize => Current.PublishDedupCacheSize;

        public class Root
        {
            public string Cluster { get; set; } = "local";
            public int TimeoutMs { get; set; } = 1000;
            public Auth Auth { get; set; } = new Auth();
            public bool JetStreamEnabled { get; set; } = true;
            public int JetStreamTimeoutMs { get; set; } = 5000;

            /// <summary>Optional RTD cluster override; null falls back to <see cref="Cluster"/>.</summary>
            public string RtdCluster { get; set; }

            /// <summary>Optional UDF cluster override; null falls back to <see cref="Cluster"/>.</summary>
            public string UdfCluster { get; set; }

            /// <summary>Admin/destructive functions (Stream*/Consumer*/Purge/Delete)
            /// require this to be true. Default off.</summary>
            public bool AllowAdminWrites { get; set; }

            /// <summary>Skip identical consecutive payloads (ordinal text) for
            /// literal subjects and polled reads. Default on.</summary>
            public bool SkipRepeatedMessages { get; set; } = true;

            /// <summary>Real-time: push the latest value at most once per window.
            /// Default on; superseded by an explicit <see cref="ConflationMs"/>.</summary>
            public bool CoalesceRealtimeUpdates { get; set; } = true;

            /// <summary>Explicit real-time conflation window in milliseconds
            /// (0 = off). Null leaves the <see cref="CoalesceRealtimeUpdates"/>
            /// boolean in charge.</summary>
            public int? ConflationMs { get; set; }

            /// <summary>Capacity of the PublishIfChanged dedup cache (LRU). Default 10000.</summary>
            public int PublishDedupCacheSize { get; set; } = 10000;

            /// <summary>Write mode: "sync", "fireforget" or "fireforget-all"
            /// (normalized to lowercase; unknown/blank falls back to "fireforget").</summary>
            public string SyncWrite { get; set; } = SyncWriteModes.FireForget;

            /// <summary>Dispatch writes off the Excel thread (Observe). Default on.</summary>
            public bool AsyncWrites { get; set; } = true;

            /// <summary>Background check for a newer release. Default on.</summary>
            public bool UpdateCheck { get; set; } = true;

            public Dictionary<string, string> Servers { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, Cluster> Clusters { get; } = new Dictionary<string, Cluster>(StringComparer.OrdinalIgnoreCase);

            /// <summary>The URL list of the default cluster, or the built-in default.</summary>
            public IReadOnlyList<string> DefaultUrls()
            {
                if (Clusters.TryGetValue(Cluster, out var c) && c.Urls.Count > 0)
                    return c.Urls;
                return new[] { DefaultUrl };
            }
        }

        public sealed class Cluster
        {
            public string Label { get; set; }
            public string ExpectCluster { get; set; }
            public List<string> Urls { get; } = new List<string>();
        }

        public sealed class Auth
        {
            public string User { get; set; }
            public string Password { get; set; }
            public string Token { get; set; }
            public string Jwt { get; set; }
            public string NkeySeed { get; set; }
            public string CredsFile { get; set; }
            public bool Tls { get; set; }
        }

        /// <summary>
        /// Fills a root with safe defaults and normalizes the values a malformed
        /// file could leave unusable (PLAN F1.4). Kept internal for tests.
        /// </summary>
        internal static Root Sanitize(Root config)
        {
            config = config ?? new Root();
            if (string.IsNullOrWhiteSpace(config.Cluster)) config.Cluster = "local";
            if (config.TimeoutMs <= 0) config.TimeoutMs = 1000;
            if (config.JetStreamTimeoutMs <= 0) config.JetStreamTimeoutMs = 5000;
            if (config.PublishDedupCacheSize <= 0) config.PublishDedupCacheSize = 10000;
            if (config.ConflationMs.HasValue)
                config.ConflationMs = Conflation.Sanitize(config.ConflationMs.Value);
            config.SyncWrite = NormalizeSyncWrite(config.SyncWrite);
            config.Auth = config.Auth ?? new Auth();
            return config;
        }
    }
}
