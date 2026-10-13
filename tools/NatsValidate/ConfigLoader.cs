using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NatsExcel.Tools
{
    /// <summary>
    /// Minimal stand-in for the add-in's config model, kept in the tool so it
    /// can be validated stand-alone (the add-in's own AppConfig is not built
    /// here). Mirrors the keys in NATSExcel.schema.json.
    /// </summary>
    internal sealed class ConfigRoot
    {
        public string DefaultCluster { get; set; }
        public int TimeoutMs { get; set; } = 1000;
        public AuthConfig DefaultAuth { get; set; } = new AuthConfig();
        public ReconnectConfig Reconnect { get; set; } = new ReconnectConfig();
        public bool JetStreamEnabled { get; set; } = true;
        public int JetStreamTimeoutMs { get; set; } = 5000;
        public Dictionary<string, string> Servers { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, ClusterConfig> Clusters { get; set; } = new Dictionary<string, ClusterConfig>(StringComparer.OrdinalIgnoreCase);
        public string FilePath { get; set; }
    }

    internal sealed class ClusterConfig
    {
        public List<string> Urls { get; set; } = new List<string>();
        public string ExpectCluster { get; set; }
        public AuthConfig Auth { get; set; }
        public int? TimeoutMs { get; set; }
        public ReconnectConfig Reconnect { get; set; }
    }

    internal sealed class AuthConfig
    {
        public string User { get; set; }
        public string Password { get; set; }
        public string Token { get; set; }
        public string Jwt { get; set; }
        public string NkeySeed { get; set; }
        public string CredsFile { get; set; }
        public bool Tls { get; set; }
        public string TlsCert { get; set; }
        public string TlsCertPassword { get; set; }
    }

    internal sealed class ReconnectConfig
    {
        public int MaxAttempts { get; set; } = 60;
        public int WaitMs { get; set; } = 2000;
        public int JitterMs { get; set; } = 100;
    }

    /// <summary>A resolved endpoint: the concrete URL list plus the effective options.</summary>
    internal sealed class ResolvedEndpoint
    {
        public List<string> Urls { get; set; } = new List<string>();
        public AuthConfig Auth { get; set; } = new AuthConfig();
        public int TimeoutMs { get; set; } = 1000;
        public ReconnectConfig Reconnect { get; set; } = new ReconnectConfig();
        public string ClusterLabel { get; set; }
        public string ExpectCluster { get; set; }
    }

    internal sealed class ConfigReferenceException : Exception
    {
        public ConfigReferenceException(string message) : base(message) { }
    }

    internal static class ConfigLoader
    {
        public static ConfigRoot Load(string path, IList<string> problems)
        {
            if (path == null)
                return new ConfigRoot();

            string raw = File.ReadAllText(path);
            JObject root;
            try
            {
                root = ParseWithComments(raw);
            }
            catch (JsonException ex)
            {
                throw new IOException("the file is not valid JSON: " + ex.Message, ex);
            }

            var config = new ConfigRoot { FilePath = path };

            // Default is a plain NAME in the shipped shape (a Clusters name, a
            // Servers name or a raw URL) - it holds no connection settings. The
            // add-in accepts the older nested object { "cluster": ..., "auth":
            // ... } as a back-compat fallback (Config.cs); mirror both here, or a
            // shipped-shape config resolves nothing and the cluster (and its
            // expectCluster assertion) is never reached.
            var defToken = root["Default"];
            if (defToken != null && defToken.Type == JTokenType.String)
            {
                config.DefaultCluster = ((string)defToken).Trim();
            }
            else if (defToken is JObject def)
            {
                config.DefaultCluster = (string)def["cluster"];
                config.TimeoutMs = ReadInt(def, "timeout", config.TimeoutMs, problems, "/Default/timeout");
                config.DefaultAuth = ReadAuth(def["auth"], problems, "/Default/auth") ?? config.DefaultAuth;
                config.Reconnect = ReadReconnect(def["reconnect"], problems, "/Default/reconnect") ?? config.Reconnect;
            }

            // The shared, process-wide blocks; a cluster overrides them field by
            // field (DESIGN.md section 8). The top-level Auth/Reconnect are the
            // shipped location and merge OVER the nested Default fallback above
            // field by field, exactly like the add-in (Config.cs).
            var sharedAuth = ReadAuth(root["Auth"], problems, "/Auth");
            if (sharedAuth != null) config.DefaultAuth = Merge(config.DefaultAuth, sharedAuth);
            var sharedReconnect = ReadReconnect(root["Reconnect"], problems, "/Reconnect");
            if (sharedReconnect != null) config.Reconnect = sharedReconnect;

            var js = root["JetStream"] as JObject;
            if (js != null)
            {
                config.JetStreamEnabled = ReadBool(js, "enabled", config.JetStreamEnabled);
                config.JetStreamTimeoutMs = ReadInt(js, "defaultTimeoutMs", config.JetStreamTimeoutMs, problems, "/JetStream/defaultTimeoutMs");
            }

            var servers = root["Servers"] as JObject;
            if (servers != null)
                foreach (var kv in servers)
                {
                    if (kv.Value.Type == JTokenType.String) config.Servers[kv.Key] = (string)kv.Value;
                    else problems.Add("/Servers/" + kv.Key + " : expected a string");
                }

            var clusters = root["Clusters"] as JObject;
            if (clusters != null)
                foreach (var kv in clusters)
                {
                    var cj = kv.Value as JObject;
                    if (cj == null) { problems.Add("/Clusters/" + kv.Key + " : expected an object"); continue; }
                    var c = new ClusterConfig();
                    var urls = cj["urls"] as JArray;
                    if (urls == null || urls.Count == 0)
                        problems.Add("/Clusters/" + kv.Key + "/urls : at least one URL is required");
                    else
                        foreach (var u in urls)
                            if (u.Type == JTokenType.String) c.Urls.Add((string)u);
                            else problems.Add("/Clusters/" + kv.Key + "/urls : expected strings");
                    c.ExpectCluster = (string)cj["expectCluster"];
                    c.Auth = ReadAuth(cj["auth"], problems, "/Clusters/" + kv.Key + "/auth");
                    c.Reconnect = ReadReconnect(cj["reconnect"], problems, "/Clusters/" + kv.Key + "/reconnect");
                    var jsc = cj["jetstream"] as JObject;
                    config.Clusters[kv.Key] = c;
                }

            // A cluster/alias referenced but not defined is the one failure that
            // must never pass silently.
            return config;
        }

        /// <summary>
        /// Resolve the cluster a formula would use (label = an argument, or the default).
        /// </summary>
        public static ResolvedEndpoint Resolve(ConfigRoot config, string reference)
        {
            string label = string.IsNullOrWhiteSpace(reference) ? config.DefaultCluster : reference.Trim();

            // 1) a Clusters name (the dictionary is ordinal-ignore-case, so the
            //    label resolves case-insensitively like the add-in's resolver).
            if (!string.IsNullOrEmpty(label) && config.Clusters.TryGetValue(label, out var cluster))
            {
                var ep = new ResolvedEndpoint
                {
                    ClusterLabel = label,
                    ExpectCluster = cluster.ExpectCluster,
                    TimeoutMs = cluster.TimeoutMs ?? config.TimeoutMs,
                    Reconnect = cluster.Reconnect ?? config.Reconnect,
                    Auth = Merge(config.DefaultAuth, cluster.Auth),
                };
                foreach (var u in cluster.Urls)
                    ep.Urls.Add(ResolveUrl(config, u));
                return ep;
            }

            // 2) a Servers alias
            if (!string.IsNullOrEmpty(label) && config.Servers.TryGetValue(label, out var serverUrl))
                return new ResolvedEndpoint { Urls = { serverUrl }, Auth = config.DefaultAuth, TimeoutMs = config.TimeoutMs, Reconnect = config.Reconnect };

            // 3) a raw URL (or list)
            if (!string.IsNullOrEmpty(label) && LooksLikeUrl(label))
                return FromRawUrls(config, SplitUrls(label).ToList());

            // 4) nothing configured -> the built-in default
            if (string.IsNullOrEmpty(label))
                return FromRawUrls(config, new List<string> { DefaultUrlOr(config) });

            throw new ConfigReferenceException(
                "'" + label + "' is not a known cluster or server; check Default.cluster, Clusters and Servers (a dangling name is not silently ignored).");
        }

        /// <summary>
        /// Builds an endpoint from concrete URLs. Because the add-in resolves a
        /// formula to a URL list before dispatch and maps that URL back to the
        /// cluster that owns it (AppConfig.ResolveSettings/FindCluster), a raw
        /// URL that happens to be one of a cluster's URLs must still surface that
        /// cluster's <c>expectCluster</c> - otherwise the assertion is dead for a
        /// config whose Default is a URL.
        /// </summary>
        private static ResolvedEndpoint FromRawUrls(ConfigRoot config, List<string> urls)
        {
            var ep = new ResolvedEndpoint
            {
                Urls = urls,
                Auth = config.DefaultAuth,
                TimeoutMs = config.TimeoutMs,
                Reconnect = config.Reconnect,
            };

            // Only an unambiguous URL match is used (two clusters sharing a URL
            // must not pick arbitrarily). The reference matches a cluster URL
            // ordinal-ignore-case, after a trim - the same compare the add-in uses.
            ClusterConfig unique = null;
            string uniqueLabel = null;
            foreach (var kv in config.Clusters)
            {
                if (kv.Value == null) continue;
                foreach (var raw in kv.Value.Urls)
                {
                    string resolved = LooksLikeUrl(raw) ? raw
                        : (config.Servers.TryGetValue(raw, out var mapped) ? mapped : raw);
                    if (urls.Any(u => string.Equals(u.Trim(), resolved.Trim(), StringComparison.OrdinalIgnoreCase)))
                    {
                        if (unique != null && !ReferenceEquals(unique, kv.Value)) return ep; // ambiguous: stay silent
                        unique = kv.Value;
                        uniqueLabel = kv.Key;
                    }
                }
            }

            if (unique != null)
            {
                ep.ClusterLabel = uniqueLabel;
                ep.ExpectCluster = unique.ExpectCluster;
                ep.TimeoutMs = unique.TimeoutMs ?? config.TimeoutMs;
                ep.Reconnect = unique.Reconnect ?? config.Reconnect;
                ep.Auth = Merge(config.DefaultAuth, unique.Auth);
            }
            return ep;
        }

        private static string DefaultUrlOr(ConfigRoot config)
        {
            if (config.Servers.TryGetValue("local", out var local)) return local;
            return "nats://localhost:4222";
        }

        private static string ResolveUrl(ConfigRoot config, string url)
        {
            if (LooksLikeUrl(url)) return url;
            if (config.Servers.TryGetValue(url, out var mapped)) return mapped;
            throw new ConfigReferenceException("cluster entry '" + url + "' is neither a URL nor a known Servers alias.");
        }

        private static bool LooksLikeUrl(string s)
        {
            return s != null && (s.Contains("://") || Regex.IsMatch(s, @"^[^,\s]+:\d+"));
        }

        private static IEnumerable<string> SplitUrls(string list)
        {
            foreach (var u in list.Split(','))
                if (!string.IsNullOrWhiteSpace(u))
                    yield return u.Trim();
        }

        private static AuthConfig Merge(AuthConfig baseAuth, AuthConfig over)
        {
            if (over == null) return baseAuth ?? new AuthConfig();
            var a = new AuthConfig
            {
                User = over.User ?? baseAuth?.User,
                Password = over.Password ?? baseAuth?.Password,
                Token = over.Token ?? baseAuth?.Token,
                Jwt = over.Jwt ?? baseAuth?.Jwt,
                NkeySeed = over.NkeySeed ?? baseAuth?.NkeySeed,
                CredsFile = over.CredsFile ?? baseAuth?.CredsFile,
                Tls = over.Tls || (baseAuth?.Tls ?? false),
                TlsCert = over.TlsCert ?? baseAuth?.TlsCert,
                TlsCertPassword = over.TlsCertPassword ?? baseAuth?.TlsCertPassword,
            };
            return a;
        }

        public static string MaskUrls(IEnumerable<string> urls)
        {
            return string.Join(",", urls.Select(MaskUrl));
        }

        /// <summary>Redact the userinfo of a URL so a credential never reaches the console.</summary>
        public static string MaskUrl(string url)
        {
            if (string.IsNullOrEmpty(url)) return url;
            return Regex.Replace(url, @"(?<=//)[^/@\s]+@", "****@");
        }

        private static AuthConfig ReadAuth(JToken token, IList<string> problems, string path)
        {
            var j = token as JObject;
            if (j == null) return null;
            return new AuthConfig
            {
                User = (string)j["user"],
                Password = (string)j["password"],
                Token = (string)j["token"],
                Jwt = (string)j["jwt"],
                NkeySeed = (string)j["nkeySeed"],
                CredsFile = (string)j["credsFile"],
                Tls = ReadBool(j, "tls", false),
                TlsCert = (string)j["tlsCert"],
                TlsCertPassword = (string)j["tlsCertPassword"],
            };
        }

        private static ReconnectConfig ReadReconnect(JToken token, IList<string> problems, string path)
        {
            var j = token as JObject;
            if (j == null) return null;
            return new ReconnectConfig
            {
                MaxAttempts = ReadInt(j, "maxAttempts", 60, problems, path + "/maxAttempts"),
                WaitMs = ReadInt(j, "waitMs", 2000, problems, path + "/waitMs"),
                JitterMs = ReadInt(j, "jitterMs", 100, problems, path + "/jitterMs"),
            };
        }

        private static int ReadInt(JObject o, string name, int fallback, IList<string> problems, string path)
        {
            var t = o[name];
            if (t == null) return fallback;
            if (t.Type == JTokenType.Integer) return (int)t;
            problems.Add(path + " : expected an integer, got " + t.Type);
            return fallback;
        }

        private static bool ReadBool(JObject o, string name, bool fallback)
        {
            var t = o[name];
            if (t == null) return fallback;
            return t.Type == JTokenType.Boolean ? (bool)t : fallback;
        }

        /// <summary>The add-in accepts // and /* */ comments; strip them, string-aware.</summary>
        public static JObject ParseWithComments(string text)
        {
            return JObject.Parse(StripComments(text));
        }

        private static string StripComments(string text)
        {
            var sb = new StringBuilder(text.Length);
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
    }
}
