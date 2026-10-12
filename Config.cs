using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;

namespace NATSExcel
{
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

        public static Root Current => _current.Value;

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
            return cfg.Servers.TryGetValue(urlOrAlias, out var mapped) ? mapped : urlOrAlias;
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

        public sealed class Root
        {
            public string Cluster { get; set; } = "local";
            public int TimeoutMs { get; set; } = 1000;
            public Auth Auth { get; set; } = new Auth();
            public bool JetStreamEnabled { get; set; } = true;
            public int JetStreamTimeoutMs { get; set; } = 5000;
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
    }
}
